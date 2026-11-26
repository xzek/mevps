/**
 * mevps - VPS 服务器管理面板
 * Cloudflare Workers + D1 全栈单文件（API + 前端）
 *
 * 部署方式（Cloudflare 控制台浏览器操作）：
 *   1. 创建 D1 数据库，名字固定：mevps
 *   2. 创建 Worker，把本文件全部内容粘贴进编辑器并部署
 *   3. 在 Worker 的 Settings -> Bindings 添加 D1 绑定：
 *      Variable name（变量名）固定为  db
 *      D1 database 选择 mevps
 *   4. 首次访问自动建表并初始化（默认账号 admin / 123456）
 */

const SESSION_COOKIE = 'mevps_session';
const SESSION_DAYS = 7;
let INITED = false;

/* ---------------- 基础工具 ---------------- */

function hex(buf) {
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

async function hashPassword(password, salt) {
  const data = new TextEncoder().encode('mevps::' + salt + '::' + password);
  const digest = await crypto.subtle.digest('SHA-256', data);
  return hex(digest);
}

function now() {
  return Date.now();
}

function json(data, status) {
  return new Response(JSON.stringify(data), {
    status: status || 200,
    headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' },
  });
}

function getCookie(request, name) {
  const raw = request.headers.get('cookie') || '';
  for (const part of raw.split(';')) {
    const i = part.indexOf('=');
    if (i < 0) continue;
    if (part.slice(0, i).trim() === name) return decodeURIComponent(part.slice(i + 1).trim());
  }
  return null;
}

function clientIp(request) {
  return request.headers.get('cf-connecting-ip') || '-';
}

/* ---------------- 图片验证码 ---------------- */

function randCode(n) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const buf = crypto.getRandomValues(new Uint8Array(n));
  let out = '';
  for (let i = 0; i < n; i++) out += chars[buf[i] % chars.length];
  return out;
}

function genCaptchaSvg(code) {
  const w = 120, h = 32;
  let deco = '';
  for (let i = 0; i < 5; i++) {
    const x1 = Math.random() * w, y1 = Math.random() * h;
    const x2 = Math.random() * w, y2 = Math.random() * h;
    deco += '<line x1="' + x1.toFixed(1) + '" y1="' + y1.toFixed(1) + '" x2="' + x2.toFixed(1) + '" y2="' + y2.toFixed(1) + '" stroke="#c7c9f5" stroke-width="1.2"/>';
  }
  for (let i = 0; i < 22; i++) {
    deco += '<circle cx="' + (Math.random() * w).toFixed(1) + '" cy="' + (Math.random() * h).toFixed(1) + '" r="' + (1 + Math.random() * 1.5).toFixed(1) + '" fill="#b9bcf0"/>';
  }
  let text = '';
  const step = w / (code.length + 1);
  for (let i = 0; i < code.length; i++) {
    const x = step * (i + 0.8);
    const y = 19 + Math.random() * 5;
    const rot = -18 + Math.random() * 36;
    const size = 17 + Math.random() * 4;
    text += '<text x="' + x.toFixed(1) + '" y="' + y.toFixed(1) + '" font-size="' + size.toFixed(1) + '" font-family="Arial, sans-serif" font-weight="bold" fill="#4b53c9" transform="rotate(' + rot.toFixed(1) + ' ' + x.toFixed(1) + ' ' + y.toFixed(1) + ')">' + code[i] + '</text>';
  }
  return '<svg xmlns="http://www.w3.org/2000/svg" width="' + w + '" height="' + h + '" viewBox="0 0 ' + w + ' ' + h + '"><rect width="' + w + '" height="' + h + '" rx="6" fill="#eef0ff"/>' + deco + text + '</svg>';
}

/* ---------------- 设置与邮件 ---------------- */

async function getSettingsMap(env) {
  const rows = (await env.db.prepare('SELECT key, value FROM settings').all()).results || [];
  const map = {};
  for (const r of rows) map[r.key] = r.value;
  return map;
}

async function sendOutlookMail(env, s, toEmail, subject, htmlBody) {
  const tenant = String(s.outlook_tenant_id || '').trim();
  const clientId = String(s.outlook_client_id || '').trim();
  const secret = String(s.outlook_client_secret || '').trim();
  const sender = String(s.outlook_sender || '').trim();
  if (!tenant || !clientId || !secret || !sender) {
    throw new Error('Outlook (Graph API) 未配置完整');
  }
  const tokenRes = await fetch('https://login.microsoftonline.com/' + tenant + '/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: secret,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials',
    }),
  });
  if (!tokenRes.ok) throw new Error('获取 Graph API 令牌失败 (' + tokenRes.status + ')');
  const tokenData = await tokenRes.json();
  const sendRes = await fetch('https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(sender) + '/sendMail', {
    method: 'POST',
    headers: {
      authorization: 'Bearer ' + tokenData.access_token,
      'content-type': 'application/json',
    },
    body: JSON.stringify({
      message: {
        subject,
        body: { contentType: 'HTML', content: htmlBody },
        toRecipients: [{ emailAddress: { address: toEmail } }],
      },
      saveToSentItems: true,
    }),
  });
  if (!sendRes.ok) throw new Error('发送邮件失败 (' + sendRes.status + ')');
}

/* ---------------- 数据库初始化（自动建表 + 默认数据） ---------------- */

async function ensureInit(env) {
  if (INITED) return;

  await env.db.batch([
    env.db.prepare(
      'CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL, salt TEXT NOT NULL, created_at INTEGER)'
    ),
    env.db.prepare(
      'CREATE TABLE IF NOT EXISTS sessions (token TEXT PRIMARY KEY, user_id INTEGER NOT NULL, created_at INTEGER, expires_at INTEGER)'
    ),
    env.db.prepare(
      'CREATE TABLE IF NOT EXISTS vps (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, ip TEXT DEFAULT "", provider TEXT DEFAULT "", location TEXT DEFAULT "", cpu TEXT DEFAULT "", memory TEXT DEFAULT "", disk TEXT DEFAULT "", os TEXT DEFAULT "", status TEXT DEFAULT "running", expire_date TEXT DEFAULT "", price REAL DEFAULT 0, note TEXT DEFAULT "", created_at INTEGER, updated_at INTEGER)'
    ),
    env.db.prepare(
      'CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER DEFAULT 0, username TEXT DEFAULT "", action TEXT DEFAULT "", target TEXT DEFAULT "", detail TEXT DEFAULT "", ip TEXT DEFAULT "", created_at INTEGER)'
    ),
    env.db.prepare('CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)'),
    env.db.prepare('CREATE INDEX IF NOT EXISTS idx_sessions_user ON sessions(user_id)'),
    env.db.prepare('CREATE INDEX IF NOT EXISTS idx_logs_created ON logs(created_at)'),
    env.db.prepare('CREATE INDEX IF NOT EXISTS idx_vps_status ON vps(status)'),
    env.db.prepare('CREATE TABLE IF NOT EXISTS captcha (id TEXT PRIMARY KEY, code TEXT NOT NULL, created_at INTEGER)'),
    env.db.prepare('CREATE TABLE IF NOT EXISTS pw_resets (token TEXT PRIMARY KEY, user_id INTEGER NOT NULL, created_at INTEGER, expires_at INTEGER, used INTEGER DEFAULT 0)'),
  ]);

  // 默认管理员：admin / 123456
  const admin = await env.db.prepare('SELECT id FROM users WHERE username = ?').bind('admin').first();
  if (!admin) {
    const salt = hex(crypto.getRandomValues(new Uint8Array(16)));
    const hash = await hashPassword('123456', salt);
    await env.db
      .prepare('INSERT INTO users (username, password_hash, salt, created_at) VALUES (?,?,?,?)')
      .bind('admin', hash, salt, now())
      .run();
  }

  // 默认系统变量（全部存 D1）
  const defaults = {
    site_name: 'mevps',
    admin_email: 'mevps',
    expire_warn_days: '7',
    captcha_enabled: '1',
    forgot_enabled: '1',
    outlook_tenant_id: '',
    outlook_client_id: '',
    outlook_client_secret: '',
    outlook_sender: '',
  };
  for (const key of Object.keys(defaults)) {
    await env.db
      .prepare('INSERT OR IGNORE INTO settings (key, value) VALUES (?,?)')
      .bind(key, defaults[key])
      .run();
  }

  INITED = true;
}

/* ---------------- 会话与日志 ---------------- */

async function getUser(request, env) {
  const token = getCookie(request, SESSION_COOKIE);
  if (!token) return null;
  const row = await env.db
    .prepare('SELECT s.expires_at AS expires_at, u.id AS id, u.username AS username FROM sessions s JOIN users u ON u.id = s.user_id WHERE s.token = ?')
    .bind(token)
    .first();
  if (!row) return null;
  if (row.expires_at < now()) {
    await env.db.prepare('DELETE FROM sessions WHERE token = ?').bind(token).run();
    return null;
  }
  return { id: row.id, username: row.username };
}

async function addLog(env, request, user, action, target, detail) {
  await env.db
    .prepare('INSERT INTO logs (user_id, username, action, target, detail, ip, created_at) VALUES (?,?,?,?,?,?,?)')
    .bind(user ? user.id : 0, user ? user.username : '-', action, target || '', detail || '', clientIp(request), now())
    .run();
}

/* ---------------- API 路由 ---------------- */

async function handleApi(request, env, url) {
  const path = url.pathname;
  const method = request.method;

  // ---- 登录（公开） ----
  if (path === '/api/login' && method === 'POST') {
    const body = await request.json().catch(() => ({}));
    const username = String(body.username || '').trim();
    const password = String(body.password || '');
    const smap = await getSettingsMap(env);
    if (smap.captcha_enabled === '1') {
      const cid = String(body.captcha_id || '');
      const ccode = String(body.captcha_code || '').trim().toUpperCase();
      const cap = await env.db.prepare('SELECT code FROM captcha WHERE id = ?').bind(cid).first();
      if (cid) await env.db.prepare('DELETE FROM captcha WHERE id = ?').bind(cid).run();
      if (!cap || !ccode || cap.code !== ccode) {
        await addLog(env, request, null, '登录失败', username, '验证码错误');
        return json({ error: '验证码错误' }, 400);
      }
    }
    const user = await env.db.prepare('SELECT * FROM users WHERE username = ?').bind(username).first();
    if (!user) {
      await addLog(env, request, null, '登录失败', username, '用户不存在');
      return json({ error: '用户名或密码错误' }, 401);
    }
    const hash = await hashPassword(password, user.salt);
    if (hash !== user.password_hash) {
      await addLog(env, request, null, '登录失败', username, '密码错误');
      return json({ error: '用户名或密码错误' }, 401);
    }
    const token = crypto.randomUUID() + '-' + hex(crypto.getRandomValues(new Uint8Array(16)));
    await env.db
      .prepare('INSERT INTO sessions (token, user_id, created_at, expires_at) VALUES (?,?,?,?)')
      .bind(token, user.id, now(), now() + SESSION_DAYS * 86400000)
      .run();
    await addLog(env, request, { id: user.id, username: user.username }, '登录系统', '', '');
    return new Response(JSON.stringify({ ok: true, username: user.username }), {
      status: 200,
      headers: {
        'content-type': 'application/json; charset=utf-8',
        'set-cookie': SESSION_COOKIE + '=' + token + '; Path=/; HttpOnly; SameSite=Lax; Secure; Max-Age=' + SESSION_DAYS * 86400,
      },
    });
  }

  // ---- 图片验证码（公开） ----
  if (path === '/api/captcha' && method === 'GET') {
    const smap = await getSettingsMap(env);
    const code = randCode(4);
    const id = crypto.randomUUID();
    await env.db.prepare('INSERT INTO captcha (id, code, created_at) VALUES (?,?,?)').bind(id, code, now()).run();
    await env.db.prepare('DELETE FROM captcha WHERE created_at < ?').bind(now() - 600000).run();
    return json({ id, svg: genCaptchaSvg(code), enabled: smap.captcha_enabled === '1' });
  }

  // ---- 忘记密码：邮箱找回（公开） ----
  if (path === '/api/forgot-password' && method === 'POST') {
    const body = await request.json().catch(() => ({}));
    const email = String(body.email || '').trim().toLowerCase();
    const smap = await getSettingsMap(env);
    const adminEmail = String(smap.admin_email || '').trim().toLowerCase();
    if (smap.forgot_enabled !== '1' || !email || !adminEmail || email !== adminEmail) {
      await addLog(env, request, null, '找回密码', email || '-', '邮箱未登记或功能未启用');
      return json({ ok: true, message: '如果该邮箱已登记，重置邮件已发送' });
    }
    try {
      const urow = await env.db.prepare('SELECT id FROM users ORDER BY id LIMIT 1').first();
      const token = crypto.randomUUID() + '-' + hex(crypto.getRandomValues(new Uint8Array(16)));
      await env.db
        .prepare('INSERT INTO pw_resets (token, user_id, created_at, expires_at, used) VALUES (?,?,?,?,0)')
        .bind(token, urow.id, now(), now() + 1800000)
        .run();
      const origin = new URL(request.url).origin;
      const link = origin + '/?reset=' + encodeURIComponent(token);
      const mailBody =
        '<div style="font-family:sans-serif;line-height:1.8"><h2>mevps 密码重置</h2>' +
        '<p>你正在找回 mevps 管理面板密码，点击下面的链接重置（30 分钟内有效）：</p>' +
        '<p><a href="' + link + '">' + link + '</a></p>' +
        '<p>如果不是你本人操作，请忽略本邮件。</p></div>';
      await sendOutlookMail(env, smap, email, 'mevps 密码重置', mailBody);
      await addLog(env, request, null, '找回密码', email, '重置邮件已发送');
    } catch (err) {
      await addLog(env, request, null, '找回密码', email, '发送失败: ' + String((err && err.message) || err));
      return json({ error: '邮件发送失败，请检查 Outlook (Graph API) 设置' }, 500);
    }
    return json({ ok: true, message: '如果该邮箱已登记，重置邮件已发送' });
  }

  // ---- 重置密码（公开，凭邮件链接令牌） ----
  if (path === '/api/reset-password' && method === 'POST') {
    const body = await request.json().catch(() => ({}));
    const token = String(body.token || '');
    const newPassword = String(body.new_password || '');
    if (newPassword.length < 6) return json({ error: '新密码至少 6 位' }, 400);
    const row = await env.db.prepare('SELECT * FROM pw_resets WHERE token = ?').bind(token).first();
    if (!row || row.used || row.expires_at < now()) return json({ error: '重置链接无效或已过期' }, 400);
    const salt = hex(crypto.getRandomValues(new Uint8Array(16)));
    const hash = await hashPassword(newPassword, salt);
    await env.db.prepare('UPDATE users SET password_hash = ?, salt = ? WHERE id = ?').bind(hash, salt, row.user_id).run();
    await env.db.prepare('UPDATE pw_resets SET used = 1 WHERE token = ?').bind(token).run();
    await addLog(env, request, null, '重置密码', '通过邮件链接', '');
    return json({ ok: true });
  }

  // ---- 其余接口需要登录 ----
  const user = await getUser(request, env);
  if (!user) return json({ error: '未登录' }, 401);

  if (path === '/api/me' && method === 'GET') {
    return json({ username: user.username });
  }

  if (path === '/api/logout' && method === 'POST') {
    const token = getCookie(request, SESSION_COOKIE);
    if (token) await env.db.prepare('DELETE FROM sessions WHERE token = ?').bind(token).run();
    await addLog(env, request, user, '退出登录', '', '');
    return new Response(JSON.stringify({ ok: true }), {
      status: 200,
      headers: {
        'content-type': 'application/json; charset=utf-8',
        'set-cookie': SESSION_COOKIE + '=; Path=/; HttpOnly; SameSite=Lax; Secure; Max-Age=0',
      },
    });
  }

  // ---- 仪表盘 ----
  if (path === '/api/dashboard' && method === 'GET') {
    const settings = await env.db.prepare('SELECT key, value FROM settings').all();
    const smap = {};
    for (const r of settings.results || []) smap[r.key] = r.value;
    const warnDays = parseInt(smap.expire_warn_days, 10) || 7;

    const total = (await env.db.prepare('SELECT COUNT(*) AS c FROM vps').first()).c;
    const running = (await env.db.prepare("SELECT COUNT(*) AS c FROM vps WHERE status = 'running'").first()).c;
    const stopped = (await env.db.prepare("SELECT COUNT(*) AS c FROM vps WHERE status = 'stopped'").first()).c;
    const maintenance = (await env.db.prepare("SELECT COUNT(*) AS c FROM vps WHERE status = 'maintenance'").first()).c;
    const expiring = (
      await env.db
        .prepare("SELECT COUNT(*) AS c FROM vps WHERE expire_date != '' AND julianday(expire_date) - julianday(date('now')) <= ? AND julianday(expire_date) - julianday(date('now')) >= 0")
        .bind(warnDays)
        .first()
    ).c;
    const expired = (
      await env.db
        .prepare("SELECT COUNT(*) AS c FROM vps WHERE expire_date != '' AND julianday(expire_date) < julianday(date('now'))")
        .first()
    ).c;
    const monthly = (await env.db.prepare('SELECT IFNULL(SUM(price),0) AS s FROM vps').first()).s;

    const providers = (
      await env.db.prepare("SELECT IFNULL(NULLIF(provider,''),'未填写') AS provider, COUNT(*) AS c FROM vps GROUP BY provider ORDER BY c DESC LIMIT 6").all()
    ).results || [];

    const recentLogs = (
      await env.db.prepare('SELECT username, action, target, detail, created_at FROM logs ORDER BY id DESC LIMIT 6').all()
    ).results || [];

    return json({
      total,
      running,
      stopped,
      maintenance,
      expiring,
      expired,
      monthly,
      warn_days: warnDays,
      providers,
      status_dist: [
        { label: '运行中', value: running },
        { label: '已停止', value: stopped },
        { label: '维护中', value: maintenance },
      ],
      recent_logs: recentLogs,
    });
  }

  // ---- VPS 管理 ----
  if (path === '/api/vps' && method === 'GET') {
    const q = (url.searchParams.get('q') || '').trim();
    const status = (url.searchParams.get('status') || '').trim();
    const page = Math.max(1, parseInt(url.searchParams.get('page'), 10) || 1);
    const size = 10;

    const where = [];
    const args = [];
    if (q) {
      where.push('(name LIKE ? OR ip LIKE ? OR provider LIKE ? OR location LIKE ? OR note LIKE ?)');
      const like = '%' + q + '%';
      args.push(like, like, like, like, like);
    }
    if (status) {
      where.push('status = ?');
      args.push(status);
    }
    const whereSql = where.length ? ' WHERE ' + where.join(' AND ') : '';

    const total = (await env.db.prepare('SELECT COUNT(*) AS c FROM vps' + whereSql).bind(...args).first()).c;
    const list = (
      await env.db
        .prepare('SELECT * FROM vps' + whereSql + ' ORDER BY id DESC LIMIT ? OFFSET ?')
        .bind(...args, size, (page - 1) * size)
        .all()
    ).results || [];
    return json({ list, total, page, size });
  }

  if (path === '/api/vps' && method === 'POST') {
    const b = await request.json().catch(() => ({}));
    const name = String(b.name || '').trim();
    if (!name) return json({ error: '名称不能为空' }, 400);
    const row = {
      name,
      ip: String(b.ip || ''),
      provider: String(b.provider || ''),
      location: String(b.location || ''),
      cpu: String(b.cpu || ''),
      memory: String(b.memory || ''),
      disk: String(b.disk || ''),
      os: String(b.os || ''),
      status: ['running', 'stopped', 'maintenance'].includes(b.status) ? b.status : 'running',
      expire_date: String(b.expire_date || ''),
      price: parseFloat(b.price) || 0,
      note: String(b.note || ''),
    };
    const res = await env.db
      .prepare('INSERT INTO vps (name, ip, provider, location, cpu, memory, disk, os, status, expire_date, price, note, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)')
      .bind(row.name, row.ip, row.provider, row.location, row.cpu, row.memory, row.disk, row.os, row.status, row.expire_date, row.price, row.note, now(), now())
      .run();
    await addLog(env, request, user, '新增VPS', row.name, row.ip + ' / ' + row.provider);
    return json({ ok: true, id: res.meta && res.meta.last_row_id });
  }

  const vpsMatch = path.match(/^\/api\/vps\/(\d+)$/);
  if (vpsMatch) {
    const id = parseInt(vpsMatch[1], 10);
    const exists = await env.db.prepare('SELECT * FROM vps WHERE id = ?').bind(id).first();
    if (!exists) return json({ error: '记录不存在' }, 404);

    if (method === 'PUT') {
      const b = await request.json().catch(() => ({}));
      const name = String(b.name || '').trim();
      if (!name) return json({ error: '名称不能为空' }, 400);
      await env.db
        .prepare('UPDATE vps SET name=?, ip=?, provider=?, location=?, cpu=?, memory=?, disk=?, os=?, status=?, expire_date=?, price=?, note=?, updated_at=? WHERE id=?')
        .bind(
          name,
          String(b.ip || ''),
          String(b.provider || ''),
          String(b.location || ''),
          String(b.cpu || ''),
          String(b.memory || ''),
          String(b.disk || ''),
          String(b.os || ''),
          ['running', 'stopped', 'maintenance'].includes(b.status) ? b.status : 'running',
          String(b.expire_date || ''),
          parseFloat(b.price) || 0,
          String(b.note || ''),
          now(),
          id
        )
        .run();
      await addLog(env, request, user, '编辑VPS', name, 'ID ' + id);
      return json({ ok: true });
    }

    if (method === 'DELETE') {
      await env.db.prepare('DELETE FROM vps WHERE id = ?').bind(id).run();
      await addLog(env, request, user, '删除VPS', exists.name, 'ID ' + id);
      return json({ ok: true });
    }
  }

  // ---- 操作日志 ----
  if (path === '/api/logs' && method === 'GET') {
    const action = (url.searchParams.get('action') || '').trim();
    const page = Math.max(1, parseInt(url.searchParams.get('page'), 10) || 1);
    const size = 15;

    const where = [];
    const args = [];
    if (action) {
      where.push('action LIKE ?');
      args.push('%' + action + '%');
    }
    const whereSql = where.length ? ' WHERE ' + where.join(' AND ') : '';

    const total = (await env.db.prepare('SELECT COUNT(*) AS c FROM logs' + whereSql).bind(...args).first()).c;
    const list = (
      await env.db
        .prepare('SELECT id, username, action, target, detail, ip, created_at FROM logs' + whereSql + ' ORDER BY id DESC LIMIT ? OFFSET ?')
        .bind(...args, size, (page - 1) * size)
        .all()
    ).results || [];
    return json({ list, total, page, size });
  }

  // ---- 系统设置（所有变量存 D1） ----
  if (path === '/api/settings' && method === 'GET') {
    const rows = (await env.db.prepare('SELECT key, value FROM settings').all()).results || [];
    const map = {};
    for (const r of rows) map[r.key] = r.value;
    return json(map);
  }

  if (path === '/api/settings' && method === 'POST') {
    const b = await request.json().catch(() => ({}));
    const allowed = ['site_name', 'admin_email', 'expire_warn_days', 'captcha_enabled', 'forgot_enabled', 'outlook_tenant_id', 'outlook_client_id', 'outlook_client_secret', 'outlook_sender'];
    const changed = [];
    for (const key of allowed) {
      if (b[key] !== undefined) {
        let value = String(b[key]).trim();
        if (key === 'captcha_enabled' || key === 'forgot_enabled') {
          if (value !== '0' && value !== '1') return json({ error: '开关值无效' }, 400);
        }
        if (key === 'expire_warn_days') {
          const n = parseInt(value, 10);
          if (isNaN(n) || n < 1 || n > 90) return json({ error: '到期提醒天数需为 1-90 的整数' }, 400);
          value = String(n);
        }
        await env.db.prepare('INSERT INTO settings (key, value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value = excluded.value').bind(key, value).run();
        changed.push(key);
      }
    }
    await addLog(env, request, user, '修改设置', '', changed.join(', '));
    return json({ ok: true });
  }

  // ---- 修改密码 ----
  if (path === '/api/password' && method === 'POST') {
    const b = await request.json().catch(() => ({}));
    const oldPassword = String(b.old_password || '');
    const newPassword = String(b.new_password || '');
    if (newPassword.length < 6) return json({ error: '新密码至少 6 位' }, 400);
    const u = await env.db.prepare('SELECT * FROM users WHERE id = ?').bind(user.id).first();
    const oldHash = await hashPassword(oldPassword, u.salt);
    if (oldHash !== u.password_hash) return json({ error: '当前密码不正确' }, 400);
    const salt = hex(crypto.getRandomValues(new Uint8Array(16)));
    const newHash = await hashPassword(newPassword, salt);
    await env.db.prepare('UPDATE users SET password_hash = ?, salt = ? WHERE id = ?').bind(newHash, salt, user.id).run();
    await addLog(env, request, user, '修改密码', user.username, '');
    return json({ ok: true });
  }

  return json({ error: 'Not Found' }, 404);
}

/* ---------------- 入口 ---------------- */

export default {
  async fetch(request, env) {
    try {
      await ensureInit(env);
      const url = new URL(request.url);
      if (url.pathname.startsWith('/api/')) {
        return await handleApi(request, env, url);
      }
      // 页面由 Static Assets（public/）在边缘直接提供，仅未命中静态资源的请求才会走到这里
      return new Response('Not Found', { status: 404 });
    } catch (err) {
      return json({ error: String((err && err.message) || err) }, 500);
    }
  },
};
