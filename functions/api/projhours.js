/**
 * /api/projhours — Cloudflare Pages Function backing the Daily Truck Tracker.
 * Stores team-entered projected hours + notes (with a daily snapshot) in D1 so
 * they persist for months, are shared across the team, and can be compared
 * against Payworks actuals later. Protected by the same site sign-in gate.
 *
 * Requires a D1 binding named "DB" on the Pages project
 * (Settings -> Functions -> D1 database bindings -> DB = element-tracker).
 *
 *   GET  /api/projhours?date=YYYY-MM-DD        -> { rows: [...] } for that day
 *   GET  /api/projhours?from=YYYY-MM-DD&to=... -> rows in a date range (reporting)
 *   POST /api/projhours  {date, truck, proj_hrs, notes, driver, bu, revenue, stops, actual_hrs}
 */
const JSON_HEADERS = { 'content-type': 'application/json; charset=utf-8' };
const num = v => { const n = parseFloat(v); return isNaN(n) ? null : n; };
const int = v => { const n = parseInt(v, 10); return isNaN(n) ? null : n; };

export async function onRequest(context) {
  const { request, env } = context;
  const db = env.DB;
  if (!db) {
    return json({ error: 'D1 binding "DB" is not configured on this Pages project.' }, 500);
  }
  const url = new URL(request.url);

  try {
    if (request.method === 'GET') {
      const date = url.searchParams.get('date');
      const from = url.searchParams.get('from');
      const to = url.searchParams.get('to');
      let res;
      if (date) {
        res = await db.prepare('SELECT * FROM proj_hours WHERE date = ?').bind(date).all();
      } else if (from && to) {
        res = await db.prepare('SELECT * FROM proj_hours WHERE date BETWEEN ? AND ? ORDER BY date, truck')
                      .bind(from, to).all();
      } else {
        res = await db.prepare('SELECT * FROM proj_hours ORDER BY date DESC, truck LIMIT 5000').all();
      }
      return json({ rows: res.results || [] });
    }

    if (request.method === 'POST') {
      const b = await request.json();
      if (!b.date || b.truck == null || b.truck === '') {
        return json({ error: 'date and truck are required' }, 400);
      }
      const who = request.headers.get('Cf-Access-Authenticated-User-Email') || b.entered_by || 'team';
      const now = new Date().toISOString();
      await db.prepare(
        `INSERT INTO proj_hours
           (date, truck, proj_hrs, notes, driver, bu, revenue, rev_d, rev_p, stops, actual_hrs, entered_by, entered_at, updated_at)
         VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
         ON CONFLICT(date, truck) DO UPDATE SET
           proj_hrs   = excluded.proj_hrs,
           notes      = excluded.notes,
           driver     = excluded.driver,
           bu         = excluded.bu,
           revenue    = excluded.revenue,
           rev_d      = excluded.rev_d,
           rev_p      = excluded.rev_p,
           stops      = excluded.stops,
           actual_hrs = COALESCE(excluded.actual_hrs, proj_hours.actual_hrs),
           updated_at = excluded.updated_at`
      ).bind(
        b.date, String(b.truck), num(b.proj_hrs), b.notes || null, b.driver || null, b.bu || null,
        num(b.revenue), num(b.rev_d), num(b.rev_p), int(b.stops), num(b.actual_hrs), who, now, now
      ).run();
      return json({ ok: true, entered_by: who });
    }

    return json({ error: 'method not allowed' }, 405);
  } catch (e) {
    return json({ error: String(e && e.message || e) }, 500);
  }
}

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), { status, headers: JSON_HEADERS });
}
