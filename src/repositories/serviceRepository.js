/**
 * Репозиторій послуг підприємства.
 */
function createServiceRepository(db) {
  function list({ search = '', category = '', onlyActive = true } = {}) {
    const conditions = [];
    const params = [];

    if (onlyActive) {
      conditions.push('is_active = 1');
    }
    if (category) {
      conditions.push('category = ?');
      params.push(category);
    }
    if (search) {
      conditions.push('(uk_lower(title) LIKE ? OR uk_lower(short_description) LIKE ? OR uk_lower(category) LIKE ?)');
      const term = `%${search.toLowerCase()}%`;
      params.push(term, term, term);
    }

    const whereClause = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
    return db
      .prepare(
        `SELECT id, slug, title, category, short_description, full_description,
                base_price_uah, sla_business_days, is_active, created_at
         FROM services
         ${whereClause}
         ORDER BY category ASC, id ASC`
      )
      .all(...params);
  }

  function findById(id) {
    return db
      .prepare(
        `SELECT id, slug, title, category, short_description, full_description,
                base_price_uah, sla_business_days, is_active, created_at
         FROM services
         WHERE id = ?`
      )
      .get(id);
  }

  return { list, findById };
}

module.exports = { createServiceRepository };
