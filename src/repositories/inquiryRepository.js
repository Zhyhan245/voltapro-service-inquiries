const { runInTransaction } = require('../db/connection');

/**
 * Репозиторій звернень, контактів клієнтів та історії змін.
 */
function createInquiryRepository(db) {
  function findByIdempotencyKey(key) {
    if (!key) return null;
    return db
      .prepare(
        `SELECT i.*, s.title AS service_title, s.category AS service_category,
                c.contact_name, c.company_name, c.phone, c.email,
                e.full_name AS assigned_employee_name
         FROM inquiries i
         JOIN services s ON s.id = i.service_id
         JOIN client_contacts c ON c.id = i.contact_id
         LEFT JOIN employees e ON e.id = i.assigned_employee_id
         WHERE i.idempotency_key = ?`
      )
      .get(key);
  }

  function createWithContactAndHistory({
    contact,
    inquiry,
    historyComment
  }) {
    return runInTransaction(db, () => {
      const contactRes = db
        .prepare(
          `INSERT INTO client_contacts (contact_name, company_name, phone, email, created_at)
           VALUES (?, ?, ?, ?, ?)`
        )
        .run(
          contact.contactName,
          contact.companyName,
          contact.phone,
          contact.email,
          inquiry.createdAt
        );
      const contactId = Number(contactRes.lastInsertRowid);

      const year = inquiry.createdAt.slice(0, 4);
      // Тимчасовий унікальний номер перед формуванням людиночитаного VP-YYYY-NNNN
      const tempTicket = `TMP-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

      const inqRes = db
        .prepare(
          `INSERT INTO inquiries (
            ticket_number, access_code, service_id, contact_id, assigned_employee_id,
            subject, message, status, rejection_reason, public_response, idempotency_key,
            sla_business_days, due_date, created_at, updated_at, closed_at
          ) VALUES (?, ?, ?, ?, NULL, ?, ?, 'NEW', NULL, NULL, ?, ?, ?, ?, ?, NULL)`
        )
        .run(
          tempTicket,
          inquiry.accessCode,
          inquiry.serviceId,
          contactId,
          inquiry.subject,
          inquiry.message,
          inquiry.idempotencyKey,
          inquiry.slaBusinessDays,
          inquiry.dueDate,
          inquiry.createdAt,
          inquiry.createdAt
        );
      const inquiryId = Number(inqRes.lastInsertRowid);
      const ticketNumber = `VP-${year}-${String(inquiryId).padStart(4, '0')}`;

      db.prepare('UPDATE inquiries SET ticket_number = ? WHERE id = ?').run(
        ticketNumber,
        inquiryId
      );

      db.prepare(
        `INSERT INTO inquiry_history (
          inquiry_id, author_employee_id, event_type, old_status, new_status,
          comment, is_internal, created_at
        ) VALUES (?, NULL, 'CREATED', NULL, 'NEW', ?, 0, ?)`
      ).run(inquiryId, historyComment, inquiry.createdAt);

      return findById(inquiryId);
    });
  }

  function findById(id) {
    return db
      .prepare(
        `SELECT i.*, s.title AS service_title, s.category AS service_category,
                c.contact_name, c.company_name, c.phone, c.email,
                e.full_name AS assigned_employee_name, e.department AS assigned_department
         FROM inquiries i
         JOIN services s ON s.id = i.service_id
         JOIN client_contacts c ON c.id = i.contact_id
         LEFT JOIN employees e ON e.id = i.assigned_employee_id
         WHERE i.id = ?`
      )
      .get(id);
  }

  function findByTicketNumber(ticketNumber) {
    return db
      .prepare(
        `SELECT i.*, s.title AS service_title, s.category AS service_category,
                c.contact_name, c.company_name, c.phone, c.email,
                e.full_name AS assigned_employee_name
         FROM inquiries i
         JOIN services s ON s.id = i.service_id
         JOIN client_contacts c ON c.id = i.contact_id
         LEFT JOIN employees e ON e.id = i.assigned_employee_id
         WHERE UPPER(i.ticket_number) = UPPER(?)`
      )
      .get(ticketNumber);
  }

  function listHistory(inquiryId, { includeInternal = false } = {}) {
    const internalFilter = includeInternal ? '' : 'AND h.is_internal = 0';
    return db
      .prepare(
        `SELECT h.id, h.inquiry_id, h.event_type, h.old_status, h.new_status,
                h.comment, h.is_internal, h.created_at,
                e.full_name AS author_name
         FROM inquiry_history h
         LEFT JOIN employees e ON e.id = h.author_employee_id
         WHERE h.inquiry_id = ? ${internalFilter}
         ORDER BY h.created_at ASC, h.id ASC`
      )
      .all(inquiryId);
  }

  function listForManager({
    employee,
    status = '',
    search = '',
    scope = 'all',
    overdueOnly = false,
    referenceDate = new Date().toISOString().slice(0, 10),
    page = 1,
    pageSize = 15
  }) {
    const conditions = [];
    const params = [];

    // Звичайний менеджер бачить лише свої або нерозподілені заявки
    if (employee.role !== 'SENIOR_MANAGER') {
      conditions.push('(i.assigned_employee_id IS NULL OR i.assigned_employee_id = ?)');
      params.push(employee.id);
    }

    if (scope === 'mine') {
      conditions.push('i.assigned_employee_id = ?');
      params.push(employee.id);
    } else if (scope === 'unassigned') {
      conditions.push('i.assigned_employee_id IS NULL');
    }

    if (status) {
      conditions.push('i.status = ?');
      params.push(status);
    }

    if (overdueOnly) {
      conditions.push("i.status IN ('NEW', 'IN_PROGRESS') AND i.due_date < ?");
      params.push(referenceDate);
    }

    if (search) {
      const term = `%${search.toLowerCase()}%`;
      conditions.push(
        '(uk_lower(i.ticket_number) LIKE ? OR uk_lower(i.subject) LIKE ? OR uk_lower(c.contact_name) LIKE ? OR uk_lower(c.email) LIKE ?)'
      );
      params.push(term, term, term, term);
    }

    const whereClause = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
    const safePage = Math.max(1, Number(page) || 1);
    const safePageSize = Math.min(50, Math.max(1, Number(pageSize) || 15));
    const offset = (safePage - 1) * safePageSize;

    const totalRow = db
      .prepare(
        `SELECT COUNT(*) AS total
         FROM inquiries i
         JOIN client_contacts c ON c.id = i.contact_id
         ${whereClause}`
      )
      .get(...params);

    const items = db
      .prepare(
        `SELECT i.id, i.ticket_number, i.service_id, i.assigned_employee_id,
                i.subject, i.status, i.sla_business_days, i.due_date,
                i.created_at, i.updated_at, i.closed_at,
                s.title AS service_title, s.category AS service_category,
                c.contact_name, c.company_name, c.email,
                e.full_name AS assigned_employee_name
         FROM inquiries i
         JOIN services s ON s.id = i.service_id
         JOIN client_contacts c ON c.id = i.contact_id
         LEFT JOIN employees e ON e.id = i.assigned_employee_id
         ${whereClause}
         ORDER BY
           CASE WHEN i.status IN ('NEW', 'IN_PROGRESS') AND i.due_date < ? THEN 0 ELSE 1 END ASC,
           i.due_date ASC,
           i.id DESC
         LIMIT ? OFFSET ?`
      )
      .all(...params, referenceDate, safePageSize, offset);

    return {
      items,
      total: Number(totalRow.total),
      page: safePage,
      pageSize: safePageSize
    };
  }

  function assignEmployee({ inquiryId, targetEmployeeId, targetEmployeeName, actorEmployeeId, nowIso }) {
    return runInTransaction(db, () => {
      db.prepare(
        'UPDATE inquiries SET assigned_employee_id = ?, updated_at = ? WHERE id = ?'
      ).run(targetEmployeeId, nowIso, inquiryId);

      const comment = targetEmployeeId
        ? `Призначено відповідального: ${targetEmployeeName}`
        : 'Знято призначення відповідального';

      db.prepare(
        `INSERT INTO inquiry_history (
          inquiry_id, author_employee_id, event_type, old_status, new_status,
          comment, is_internal, created_at
        ) VALUES (?, ?, 'ASSIGNMENT', NULL, NULL, ?, 1, ?)`
      ).run(inquiryId, actorEmployeeId, comment, nowIso);

      return findById(inquiryId);
    });
  }

  function changeStatus({
    inquiryId,
    oldStatus,
    newStatus,
    rejectionReason,
    publicResponse,
    comment,
    actorEmployeeId,
    nowIso
  }) {
    return runInTransaction(db, () => {
      const closedAt =
        newStatus === 'COMPLETED' || newStatus === 'REJECTED' ? nowIso : null;

      db.prepare(
        `UPDATE inquiries
         SET status = ?,
             rejection_reason = COALESCE(?, rejection_reason),
             public_response = COALESCE(?, public_response),
             updated_at = ?,
             closed_at = ?
         WHERE id = ?`
      ).run(
        newStatus,
        newStatus === 'IN_PROGRESS' ? null : rejectionReason,
        publicResponse || null,
        nowIso,
        closedAt,
        inquiryId
      );

      db.prepare(
        `INSERT INTO inquiry_history (
          inquiry_id, author_employee_id, event_type, old_status, new_status,
          comment, is_internal, created_at
        ) VALUES (?, ?, 'STATUS_CHANGE', ?, ?, ?, 0, ?)`
      ).run(inquiryId, actorEmployeeId, oldStatus, newStatus, comment, nowIso);

      return findById(inquiryId);
    });
  }

  function addComment({ inquiryId, authorEmployeeId, comment, isInternal, nowIso }) {
    return runInTransaction(db, () => {
      const eventType = isInternal ? 'INTERNAL_NOTE' : 'PUBLIC_COMMENT';
      db.prepare(
        `INSERT INTO inquiry_history (
          inquiry_id, author_employee_id, event_type, old_status, new_status,
          comment, is_internal, created_at
        ) VALUES (?, ?, ?, NULL, NULL, ?, ?, ?)`
      ).run(inquiryId, authorEmployeeId, eventType, comment, isInternal ? 1 : 0, nowIso);

      if (!isInternal) {
        db.prepare(
          'UPDATE inquiries SET public_response = ?, updated_at = ? WHERE id = ?'
        ).run(comment, nowIso, inquiryId);
      } else {
        db.prepare('UPDATE inquiries SET updated_at = ? WHERE id = ?').run(nowIso, inquiryId);
      }

      return findById(inquiryId);
    });
  }

  function listActiveForSlaReport(employee) {
    const conditions = ["i.status IN ('NEW', 'IN_PROGRESS')"];
    const params = [];
    if (employee.role !== 'SENIOR_MANAGER') {
      conditions.push('(i.assigned_employee_id IS NULL OR i.assigned_employee_id = ?)');
      params.push(employee.id);
    }

    return db
      .prepare(
        `SELECT i.id, i.ticket_number, i.subject, i.status, i.sla_business_days,
                i.due_date, i.created_at, i.assigned_employee_id,
                s.title AS service_title, s.category AS service_category,
                c.contact_name, c.company_name,
                e.full_name AS assigned_employee_name
         FROM inquiries i
         JOIN services s ON s.id = i.service_id
         JOIN client_contacts c ON c.id = i.contact_id
         LEFT JOIN employees e ON e.id = i.assigned_employee_id
         WHERE ${conditions.join(' AND ')}
         ORDER BY i.due_date ASC, i.id ASC`
      )
      .all(...params);
  }

  return {
    findByIdempotencyKey,
    createWithContactAndHistory,
    findById,
    findByTicketNumber,
    listHistory,
    listForManager,
    assignEmployee,
    changeStatus,
    addComment,
    listActiveForSlaReport
  };
}

module.exports = { createInquiryRepository };
