const { generateAccessCode } = require('../utils/security');
const { addBusinessDays, evaluateSlaState, toIsoDateOnly } = require('../domain/calendarRules');
const {
  DomainError,
  validateStatusTransition,
  assertEmployeeCanAccessInquiry
} = require('../domain/inquiryRules');
const {
  validateCreateInquiryInput,
  validatePublicLookupQuery,
  cleanText
} = require('../domain/validationRules');

/**
 * Прикладний сервіс керування зверненнями клієнтів та контролем SLA.
 */
function createInquiryService({
  inquiryRepo,
  serviceRepo,
  employeeRepo,
  calendarRepo
}) {
  function getReferenceDate(overrideDate) {
    return overrideDate ? toIsoDateOnly(overrideDate) : toIsoDateOnly();
  }

  function toPublicInquiryDto(row, historyRows, exceptionsMap, referenceDate) {
    const sla = evaluateSlaState(row, referenceDate, exceptionsMap);
    return {
      ticketNumber: row.ticket_number,
      serviceId: row.service_id,
      serviceTitle: row.service_title,
      serviceCategory: row.service_category,
      subject: row.subject,
      status: row.status,
      rejectionReason: row.status === 'REJECTED' ? row.rejection_reason : null,
      publicResponse: row.public_response,
      assignedEmployeeName: row.assigned_employee_name || 'Очікує призначення',
      slaBusinessDays: row.sla_business_days,
      dueDate: row.due_date,
      slaState: sla.state,
      overdueBusinessDays: sla.overdueBusinessDays,
      createdAt: row.created_at,
      updatedAt: row.updated_at,
      closedAt: row.closed_at,
      publicHistory: historyRows.map((h) => ({
        id: h.id,
        eventType: h.event_type,
        oldStatus: h.old_status,
        newStatus: h.new_status,
        comment: h.comment,
        authorName: h.author_name || 'Система',
        createdAt: h.created_at
      }))
    };
  }

  function toManagerListItemDto(row, exceptionsMap, referenceDate) {
    const sla = evaluateSlaState(row, referenceDate, exceptionsMap);
    return {
      id: row.id,
      ticketNumber: row.ticket_number,
      serviceId: row.service_id,
      serviceTitle: row.service_title,
      serviceCategory: row.service_category,
      contactName: row.contact_name,
      companyName: row.company_name,
      email: row.email,
      assignedEmployeeId: row.assigned_employee_id,
      assignedEmployeeName: row.assigned_employee_name,
      subject: row.subject,
      status: row.status,
      slaBusinessDays: row.sla_business_days,
      dueDate: row.due_date,
      slaState: sla.state,
      overdueBusinessDays: sla.overdueBusinessDays,
      remainingBusinessDays: sla.remainingBusinessDays,
      createdAt: row.created_at,
      updatedAt: row.updated_at,
      closedAt: row.closed_at
    };
  }

  function toManagerDetailDto(row, historyRows, exceptionsMap, referenceDate) {
    const sla = evaluateSlaState(row, referenceDate, exceptionsMap);
    return {
      id: row.id,
      ticketNumber: row.ticket_number,
      serviceId: row.service_id,
      serviceTitle: row.service_title,
      serviceCategory: row.service_category,
      contact: {
        name: row.contact_name,
        companyName: row.company_name,
        phone: row.phone,
        email: row.email
      },
      assignedEmployeeId: row.assigned_employee_id,
      assignedEmployeeName: row.assigned_employee_name,
      subject: row.subject,
      message: row.message,
      status: row.status,
      rejectionReason: row.rejection_reason,
      publicResponse: row.public_response,
      slaBusinessDays: row.sla_business_days,
      dueDate: row.due_date,
      slaState: sla.state,
      overdueBusinessDays: sla.overdueBusinessDays,
      remainingBusinessDays: sla.remainingBusinessDays,
      createdAt: row.created_at,
      updatedAt: row.updated_at,
      closedAt: row.closed_at,
      history: historyRows.map((h) => ({
        id: h.id,
        eventType: h.event_type,
        oldStatus: h.old_status,
        newStatus: h.new_status,
        comment: h.comment,
        isInternal: Boolean(h.is_internal),
        authorName: h.author_name || 'Система',
        createdAt: h.created_at
      }))
    };
  }

  /**
   * Створює нове звернення (із захистом від дублювання через idempotencyKey).
   */
  function createPublicInquiry(body) {
    const input = validateCreateInquiryInput(body);

    if (input.idempotencyKey) {
      const existing = inquiryRepo.findByIdempotencyKey(input.idempotencyKey);
      if (existing) {
        return {
          duplicate: true,
          ticketNumber: existing.ticket_number,
          accessCode: existing.access_code,
          status: existing.status,
          serviceTitle: existing.service_title,
          slaBusinessDays: existing.sla_business_days,
          dueDate: existing.due_date,
          createdAt: existing.created_at
        };
      }
    }

    const service = serviceRepo.findById(input.serviceId);
    if (!service || !service.is_active) {
      throw new DomainError(
        'SERVICE_NOT_FOUND',
        404,
        'Обрану послугу не знайдено або вона тимчасово недоступна.'
      );
    }

    const nowIso = new Date().toISOString();
    const exceptionsMap = calendarRepo.getExceptionsMap();
    const dueDate = addBusinessDays(nowIso, service.sla_business_days, exceptionsMap);
    const accessCode = generateAccessCode();

    const created = inquiryRepo.createWithContactAndHistory({
      contact: {
        contactName: input.contactName,
        companyName: input.companyName,
        phone: input.phone,
        email: input.email
      },
      inquiry: {
        accessCode,
        serviceId: service.id,
        subject: input.subject,
        message: input.message,
        idempotencyKey: input.idempotencyKey,
        slaBusinessDays: service.sla_business_days,
        dueDate,
        createdAt: nowIso
      },
      historyComment: `Звернення зареєстровано у системі. Очікуваний строк опрацювання: ${service.sla_business_days} роб. дн. (до ${dueDate}).`
    });

    return {
      duplicate: false,
      ticketNumber: created.ticket_number,
      accessCode: created.access_code,
      status: created.status,
      serviceTitle: created.service_title,
      slaBusinessDays: created.sla_business_days,
      dueDate: created.due_date,
      createdAt: created.created_at
    };
  }

  /**
   * Публічний пошук заявки за номером та кодом доступу (без внутрішніх приміток).
   */
  function lookupPublicInquiry(query, referenceDateOverride) {
    const { ticketNumber, accessCode } = validatePublicLookupQuery(query);
    const inquiry = inquiryRepo.findByTicketNumber(ticketNumber);
    if (!inquiry || inquiry.access_code !== accessCode) {
      throw new DomainError(
        'INQUIRY_NOT_FOUND',
        404,
        'Звернення з таким номером і кодом доступу не знайдено.'
      );
    }

    const history = inquiryRepo.listHistory(inquiry.id, { includeInternal: false });
    const exceptionsMap = calendarRepo.getExceptionsMap();
    return toPublicInquiryDto(
      inquiry,
      history,
      exceptionsMap,
      getReferenceDate(referenceDateOverride)
    );
  }

  function listManagerInquiries(employee, filters = {}) {
    const referenceDate = getReferenceDate(filters.referenceDate);
    const exceptionsMap = calendarRepo.getExceptionsMap();

    const result = inquiryRepo.listForManager({
      employee,
      status: cleanText(filters.status),
      search: cleanText(filters.search),
      scope: cleanText(filters.scope) || 'all',
      overdueOnly: filters.overdueOnly === 'true' || filters.overdueOnly === true,
      referenceDate,
      page: filters.page,
      pageSize: filters.pageSize
    });

    return {
      items: result.items.map((item) =>
        toManagerListItemDto(item, exceptionsMap, referenceDate)
      ),
      total: result.total,
      page: result.page,
      pageSize: result.pageSize
    };
  }

  function getManagerInquiry(employee, inquiryId, referenceDateOverride) {
    const inquiry = inquiryRepo.findById(Number(inquiryId));
    if (!inquiry) {
      throw new DomainError('INQUIRY_NOT_FOUND', 404, 'Звернення не знайдено.');
    }
    assertEmployeeCanAccessInquiry(employee, inquiry);

    const history = inquiryRepo.listHistory(inquiry.id, { includeInternal: true });
    const exceptionsMap = calendarRepo.getExceptionsMap();
    return toManagerDetailDto(
      inquiry,
      history,
      exceptionsMap,
      getReferenceDate(referenceDateOverride)
    );
  }

  function assignInquiry(employee, inquiryId, targetEmployeeIdRaw) {
    const inquiry = inquiryRepo.findById(Number(inquiryId));
    if (!inquiry) {
      throw new DomainError('INQUIRY_NOT_FOUND', 404, 'Звернення не знайдено.');
    }
    assertEmployeeCanAccessInquiry(employee, inquiry);

    let targetEmployeeId = null;
    let targetEmployeeName = null;

    if (targetEmployeeIdRaw !== null && targetEmployeeIdRaw !== undefined && targetEmployeeIdRaw !== '') {
      const parsedId = Number(targetEmployeeIdRaw);
      if (!Number.isInteger(parsedId) || parsedId <= 0) {
        throw new DomainError('VALIDATION_ERROR', 400, 'Некоректний ідентифікатор працівника.');
      }
      if (employee.role !== 'SENIOR_MANAGER' && parsedId !== employee.id) {
        throw new DomainError(
          'ASSIGN_FORBIDDEN',
          403,
          'Менеджер може призначити заявку лише на себе; перерозподіл на інших доступний старшому менеджеру.'
        );
      }
      const target = employeeRepo.findById(parsedId);
      if (!target || !target.is_active) {
        throw new DomainError('EMPLOYEE_NOT_FOUND', 404, 'Працівника не знайдено або деактивовано.');
      }
      targetEmployeeId = target.id;
      targetEmployeeName = target.full_name;
    } else if (employee.role !== 'SENIOR_MANAGER') {
      throw new DomainError(
        'UNASSIGN_FORBIDDEN',
        403,
        'Зняти відповідального може лише старший менеджер.'
      );
    }

    const updated = inquiryRepo.assignEmployee({
      inquiryId: inquiry.id,
      targetEmployeeId,
      targetEmployeeName,
      actorEmployeeId: employee.id,
      nowIso: new Date().toISOString()
    });

    const history = inquiryRepo.listHistory(updated.id, { includeInternal: true });
    return toManagerDetailDto(
      updated,
      history,
      calendarRepo.getExceptionsMap(),
      getReferenceDate()
    );
  }

  function updateInquiryStatus(employee, inquiryId, body) {
    const inquiry = inquiryRepo.findById(Number(inquiryId));
    if (!inquiry) {
      throw new DomainError('INQUIRY_NOT_FOUND', 404, 'Звернення не знайдено.');
    }
    assertEmployeeCanAccessInquiry(employee, inquiry);

    const nextStatus = cleanText(body?.status);
    const publicResponse = cleanText(body?.publicResponse);
    const normalizedReason = validateStatusTransition({
      currentStatus: inquiry.status,
      nextStatus,
      employeeRole: employee.role,
      rejectionReason: body?.rejectionReason
    });

    const statusLabels = {
      NEW: 'Нова',
      IN_PROGRESS: 'В роботі',
      COMPLETED: 'Завершена',
      REJECTED: 'Відхилена'
    };

    let comment = `Статус змінено: «${statusLabels[inquiry.status]}» → «${statusLabels[nextStatus]}».`;
    if (nextStatus === 'REJECTED') {
      comment += ` Причина: ${normalizedReason}`;
    } else if (publicResponse) {
      comment += ` Коментар: ${publicResponse}`;
    }

    const updated = inquiryRepo.changeStatus({
      inquiryId: inquiry.id,
      oldStatus: inquiry.status,
      newStatus: nextStatus,
      rejectionReason: normalizedReason,
      publicResponse: publicResponse || null,
      comment,
      actorEmployeeId: employee.id,
      nowIso: new Date().toISOString()
    });

    const history = inquiryRepo.listHistory(updated.id, { includeInternal: true });
    return toManagerDetailDto(
      updated,
      history,
      calendarRepo.getExceptionsMap(),
      getReferenceDate()
    );
  }

  function addInquiryComment(employee, inquiryId, body) {
    const inquiry = inquiryRepo.findById(Number(inquiryId));
    if (!inquiry) {
      throw new DomainError('INQUIRY_NOT_FOUND', 404, 'Звернення не знайдено.');
    }
    assertEmployeeCanAccessInquiry(employee, inquiry);

    const comment = cleanText(body?.comment);
    if (comment.length < 2 || comment.length > 2000) {
      throw new DomainError(
        'VALIDATION_ERROR',
        400,
        'Текст коментаря має містити від 2 до 2000 символів.',
        [{ field: 'comment', message: 'Введіть текст від 2 до 2000 символів.' }]
      );
    }

    const isInternal = Boolean(body?.isInternal);
    const updated = inquiryRepo.addComment({
      inquiryId: inquiry.id,
      authorEmployeeId: employee.id,
      comment,
      isInternal,
      nowIso: new Date().toISOString()
    });

    const history = inquiryRepo.listHistory(updated.id, { includeInternal: true });
    return toManagerDetailDto(
      updated,
      history,
      calendarRepo.getExceptionsMap(),
      getReferenceDate()
    );
  }

  /**
   * Формує звіт про прострочені заявки та заявки з дедлайном сьогодні.
   */
  function getOverdueReport(employee, referenceDateOverride) {
    const referenceDate = getReferenceDate(referenceDateOverride);
    const exceptionsMap = calendarRepo.getExceptionsMap();
    const activeRows = inquiryRepo.listActiveForSlaReport(employee);

    const overdue = [];
    const dueToday = [];
    let onTimeCount = 0;

    for (const row of activeRows) {
      const sla = evaluateSlaState(row, referenceDate, exceptionsMap);
      const item = {
        id: row.id,
        ticketNumber: row.ticket_number,
        subject: row.subject,
        serviceTitle: row.service_title,
        serviceCategory: row.service_category,
        contactName: row.contact_name,
        companyName: row.company_name,
        assignedEmployeeName: row.assigned_employee_name || 'Не призначено',
        status: row.status,
        slaBusinessDays: row.sla_business_days,
        dueDate: row.due_date,
        slaState: sla.state,
        overdueBusinessDays: sla.overdueBusinessDays
      };

      if (sla.state === 'OVERDUE') {
        overdue.push(item);
      } else if (sla.state === 'DUE_TODAY') {
        dueToday.push(item);
      } else {
        onTimeCount += 1;
      }
    }

    overdue.sort((a, b) => b.overdueBusinessDays - a.overdueBusinessDays);

    return {
      referenceDate,
      summary: {
        activeTotal: activeRows.length,
        overdueCount: overdue.length,
        dueTodayCount: dueToday.length,
        onTimeCount
      },
      overdue,
      dueToday,
      calendarExceptions: calendarRepo.listExceptions()
    };
  }

  return {
    createPublicInquiry,
    lookupPublicInquiry,
    listManagerInquiries,
    getManagerInquiry,
    assignInquiry,
    updateInquiryStatus,
    addInquiryComment,
    getOverdueReport
  };
}

module.exports = { createInquiryService };
