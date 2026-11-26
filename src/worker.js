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
  const w = 120, h = 44;
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
    const y = 28 + Math.random() * 8;
    const rot = -18 + Math.random() * 36;
    const size = 22 + Math.random() * 6;
    text += '<text x="' + x.toFixed(1) + '" y="' + y.toFixed(1) + '" font-size="' + size.toFixed(1) + '" font-family="Arial, sans-serif" font-weight="bold" fill="#4b53c9" transform="rotate(' + rot.toFixed(1) + ' ' + x.toFixed(1) + ' ' + y.toFixed(1) + ')">' + code[i] + '</text>';
  }
  return '<svg xmlns="http://www.w3.org/2000/svg" width="' + w + '" height="' + h + '" viewBox="0 0 ' + w + ' ' + h + '"><rect width="' + w + '" height="' + h + '" rx="8" fill="#eef0ff"/>' + deco + text + '</svg>';
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
      if (request.method === 'GET') {
        return new Response(HTML, {
          headers: { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store' },
        });
      }
      return new Response('Not Found', { status: 404 });
    } catch (err) {
      return json({ error: String((err && err.message) || err) }, 500);
    }
  },
};

/* ================= 前端页面（soybean-admin 风格） ================= */

const HTML = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>mevps · VPS 管理系统</title>
<style>
:root {
  --primary: #646cff;
  --primary-hover: #5259e8;
  --primary-weak: #eef0ff;
  --bg: #f0f2f5;
  --border: #e5e6eb;
  --text-1: #1d2129;
  --text-2: #4e5969;
  --text-3: #86909c;
}
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; height: 100%; }
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", "Microsoft YaHei", sans-serif;
  background: var(--bg); color: var(--text-1); font-size: 14px;
}
button { font-family: inherit; }
a { color: var(--primary); text-decoration: none; }

/* ===== 登录页 ===== */
.login-wrap {
  position: relative; display: flex; height: 100vh; align-items: center; justify-content: center;
  background: linear-gradient(135deg, #e0e7ff 0%, #c7d2fe 50%, #a5b4fc 100%); overflow: hidden;
}
.blob-1 {
  position: absolute; top: -220px; left: -160px; width: 680px; height: 680px; border-radius: 50%;
  background: linear-gradient(135deg, #818cf8, #646cff); opacity: .45;
}
.blob-2 {
  position: absolute; bottom: -260px; right: -140px; width: 720px; height: 560px;
  border-radius: 50% 50% 0 0; background: linear-gradient(180deg, #646cff, #818cf8); opacity: .55;
}
.login-card {
  position: relative; width: 448px; background: #fff; border-radius: 12px;
  padding: 36px 40px 28px; box-shadow: 0 12px 48px rgba(60, 70, 180, .22);
}
.login-brand { display: flex; align-items: center; gap: 12px; margin-bottom: 26px; }
.brand-logo {
  width: 40px; height: 40px; border-radius: 10px; background: linear-gradient(135deg, #8b93ff, #646cff);
  color: #fff; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 20px;
}
.brand-name { font-size: 18px; font-weight: 600; color: var(--text-1); }
.brand-sub { font-size: 12px; color: var(--text-3); margin-top: 2px; }
.login-tab { text-align: center; font-size: 15px; font-weight: 600; color: var(--primary); margin-bottom: 22px; }
.login-extra { display: flex; gap: 10px; margin-top: 16px; justify-content: center; }

/* ===== 表单 ===== */
.form-input {
  width: 100%; height: 42px; padding: 0 12px; border: 1px solid var(--border); border-radius: 8px;
  font-size: 14px; outline: none; background: #fff; transition: all .15s;
}
.form-input:focus { border-color: var(--primary); box-shadow: 0 0 0 3px rgba(100, 108, 255, .12); }
.form-input::placeholder { color: #94a3b8; }
.login-row { display: flex; align-items: center; justify-content: space-between; margin: -4px 0 18px; }
.remember { display: flex; align-items: center; gap: 6px; font-size: 12px; color: var(--text-2); cursor: pointer; }
.remember input { accent-color: var(--primary); }
.forgot { font-size: 12px; color: var(--primary); cursor: pointer; }
.forgot:hover { text-decoration: underline; }
.greet { display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; }
.greet-title { font-size: 18px; font-weight: 600; }
.greet-sub { font-size: 13px; color: var(--text-3); margin-top: 6px; }
.greet-meta { font-size: 12px; color: var(--text-3); }
.form-item { margin-bottom: 16px; }
.form-item label { display: block; font-size: 13px; color: var(--text-2); margin-bottom: 6px; }
.form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px 16px; }

/* ===== 按钮 ===== */
.btn {
  padding: 8px 18px; border-radius: 6px; border: 1px solid transparent; cursor: pointer;
  font-size: 14px; transition: all .15s;
}
.btn-primary { background: var(--primary); color: #fff; }
.btn-primary:hover { background: var(--primary-hover); }
.btn-pill {
  width: 100%; height: 44px; border-radius: 40px;
  background: linear-gradient(135deg, #646cff 0%, #818cf8 100%); color: #fff;
  font-size: 15px; font-weight: 500; border: none; cursor: pointer; transition: all .15s;
}
.btn-pill:hover { background: linear-gradient(135deg, #5259e8 0%, #646cff 100%); }
.chip {
  height: 34px; padding: 0 16px; border: none; border-radius: 6px;
  background: var(--primary); color: #fff; font-size: 12px; cursor: pointer;
}
.chip:hover { background: var(--primary-hover); }
.chip-ghost { background: var(--primary-weak); color: var(--primary); }
.btn-ghost { background: #fff; border-color: var(--border); color: var(--text-2); }
.btn-ghost:hover { border-color: var(--primary); color: var(--primary); }
.btn-danger { background: #fff; border-color: #ffccc7; color: #d03050; }
.btn-danger:hover { background: #fff1f0; }
.btn:disabled { opacity: .5; cursor: not-allowed; }
.btn-sm { padding: 4px 12px; font-size: 12px; }

/* ===== 布局 ===== */
.app { display: flex; height: 100vh; overflow: hidden; }
.sidebar {
  width: 240px; background: #fff; border-right: 1px solid var(--border);
  display: flex; flex-direction: column; flex-shrink: 0;
}
.logo {
  height: 64px; display: flex; align-items: center; gap: 12px; padding: 0 16px;
  border-bottom: 1px solid var(--border);
}
.logo-badge {
  width: 32px; height: 32px; border-radius: 8px; background: linear-gradient(135deg, #8b93ff, #646cff);
  color: #fff; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 16px;
}
.logo-name { font-size: 16px; font-weight: 600; letter-spacing: .5px; }
.logo-sub { font-size: 10px; color: var(--text-3); margin-top: 1px; }
.menu { padding: 8px 0; flex: 1; }
.menu-item {
  display: flex; align-items: center; gap: 10px; height: 40px; padding: 0 16px; margin: 2px 8px;
  border-radius: 4px; cursor: pointer; color: var(--text-2); font-size: 14px; user-select: none;
}
.menu-item:hover { background: #f5f5f5; color: var(--text-1); }
.menu-item.active { background: #e8eaff; color: var(--primary); font-weight: 500; }
.menu-item svg { width: 16px; height: 16px; flex-shrink: 0; }
.sidebar-foot { padding: 14px 16px; border-top: 1px solid var(--border); font-size: 11px; color: var(--text-3); }

.main { flex: 1; display: flex; flex-direction: column; overflow: hidden; }
.header {
  height: 56px; background: #fff; border-bottom: 1px solid var(--border);
  display: flex; align-items: center; justify-content: space-between; padding: 0 20px; flex-shrink: 0;
}
.header .title { font-size: 16px; font-weight: 600; }
.user-chip { display: flex; align-items: center; gap: 10px; font-size: 13px; color: var(--text-2); }
.avatar {
  width: 28px; height: 28px; border-radius: 50%; background: linear-gradient(135deg, #8b93ff, #646cff);
  color: #fff; display: flex; align-items: center; justify-content: center; font-weight: 600; font-size: 12px;
}
.content { flex: 1; overflow: auto; padding: 24px; }

/* ===== 卡片 ===== */
.card { background: #fff; border-radius: 8px; padding: 24px; box-shadow: 0 1px 2px rgba(0, 0, 0, .05); }
.cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; }
.stat-card { border-radius: 12px; padding: 20px 24px; color: #fff; overflow: hidden; }
.stat-card .t { font-size: 14px; font-weight: 500; opacity: .95; display: flex; align-items: center; gap: 8px; }
.stat-card .t svg { width: 16px; height: 16px; }
.stat-card .n { font-size: 32px; font-weight: 600; text-align: right; margin-top: 14px; }
.stat-card .s { font-size: 12px; opacity: .85; margin-top: 8px; }
.g1 { background: linear-gradient(90deg, #ec5e8f, #8b5cf6); }
.g2 { background: linear-gradient(90deg, #8b5cf6, #646cff); }
.g3 { background: linear-gradient(90deg, #38bdf8, #22d3ee); }
.g4 { background: linear-gradient(90deg, #f59e0b, #fb923c); }
.row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-top: 16px; }
.section-title { font-size: 15px; font-weight: 600; margin: 0 0 16px; color: var(--text-1); }
.card-block { margin-top: 16px; }

/* ===== 图表 ===== */
.bar-row { display: flex; align-items: center; gap: 10px; margin: 12px 0; }
.bar-label { width: 110px; font-size: 13px; color: var(--text-2); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.bar-track { flex: 1; height: 10px; background: #f2f3f5; border-radius: 5px; overflow: hidden; }
.bar-fill { height: 100%; background: linear-gradient(90deg, #a5b4fc, #646cff); border-radius: 5px; }
.bar-num { width: 30px; text-align: right; font-size: 13px; color: var(--text-2); }
.donut-wrap { display: flex; align-items: center; gap: 28px; justify-content: center; padding: 10px 0; }
.donut { width: 160px; height: 160px; border-radius: 50%; position: relative; flex-shrink: 0; }
.donut-hole {
  position: absolute; inset: 32px; background: #fff; border-radius: 50%;
  display: flex; flex-direction: column; align-items: center; justify-content: center;
}
.donut-hole .n { font-size: 24px; font-weight: 700; }
.donut-hole .t { font-size: 11px; color: var(--text-3); }
.legend-item { display: flex; align-items: center; gap: 8px; margin: 8px 0; font-size: 13px; color: var(--text-2); }
.legend-dot { width: 10px; height: 10px; border-radius: 3px; }
.log-line {
  display: flex; align-items: center; gap: 10px; padding: 10px 0;
  border-bottom: 1px solid #f2f3f5; font-size: 13px;
}
.log-line:last-child { border-bottom: none; }
.log-time { color: var(--text-3); width: 130px; flex-shrink: 0; }
.log-tag {
  background: var(--primary-weak); color: var(--primary); border-radius: 4px;
  padding: 1px 8px; font-size: 12px; flex-shrink: 0;
}

/* ===== 表格 ===== */
.toolbar { display: flex; gap: 10px; margin-bottom: 16px; align-items: center; flex-wrap: wrap; }
.toolbar .spacer { flex: 1; }
select.form-input { width: auto; padding: 8px 10px; height: 40px; }
table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; }
th {
  text-align: left; font-weight: 500; font-size: 13px; color: var(--text-2);
  background: #f7f8fa; padding: 12px; border-bottom: 1px solid var(--border); white-space: nowrap;
}
td { padding: 12px; border-bottom: 1px solid #f2f3f5; font-size: 13px; color: var(--text-1); }
tr:hover td { background: #fafbfc; }
.cell-main { font-weight: 500; }
.cell-sub { color: var(--text-3); font-size: 12px; margin-top: 2px; }
.badge {
  display: inline-flex; align-items: center; gap: 6px; padding: 2px 10px;
  border-radius: 10px; font-size: 12px; white-space: nowrap;
}
.badge .dot { width: 6px; height: 6px; border-radius: 50%; }
.badge-running { background: #e8fff4; color: #0e9f6e; }
.badge-running .dot { background: #0e9f6e; }
.badge-stopped { background: #f2f3f5; color: #86909c; }
.badge-stopped .dot { background: #86909c; }
.badge-maintenance { background: #fff7e6; color: #d97706; }
.badge-maintenance .dot { background: #d97706; }
.badge-warn { background: #fff7e6; color: #d97706; }
.badge-expired { background: #fff1f0; color: #d03050; }
.pager { display: flex; gap: 10px; align-items: center; justify-content: flex-end; margin-top: 16px; }
.pager-info { color: var(--text-3); font-size: 13px; }
.empty { text-align: center; color: var(--text-3); padding: 40px 0; }

/* ===== 弹窗 ===== */
.modal-mask {
  position: fixed; inset: 0; background: rgba(0, 0, 0, .35); display: none;
  align-items: center; justify-content: center; z-index: 50;
}
.modal {
  background: #fff; border-radius: 12px; width: 620px; max-width: 94vw;
  padding: 24px; max-height: 86vh; overflow: auto;
}
.modal-title { font-size: 16px; font-weight: 600; margin: 0 0 18px; }
.modal-foot { display: flex; justify-content: flex-end; gap: 10px; margin-top: 20px; }
.modal-sm { width: 400px; }

/* ===== 验证码 ===== */
.captcha-row { display: flex; gap: 10px; align-items: center; }
.captcha-box {
  width: 120px; height: 44px; border: 1px solid var(--border); border-radius: 8px;
  overflow: hidden; cursor: pointer; flex-shrink: 0; background: #eef0ff;
}
.captcha-box svg { display: block; }

/* ===== 开关（naive 风格） ===== */
.switch-row {
  display: flex; align-items: center; justify-content: space-between;
  padding: 12px 0; border-bottom: 1px solid #f2f3f5; gap: 16px;
}
.switch-row:last-child { border-bottom: none; }
.switch-label { font-size: 14px; font-weight: 500; }
.switch {
  width: 44px; height: 22px; border-radius: 11px; background: #c0c4cc;
  position: relative; cursor: pointer; transition: background .2s; flex-shrink: 0;
}
.switch .knob {
  position: absolute; top: 2px; left: 2px; width: 18px; height: 18px; border-radius: 50%;
  background: #fff; transition: left .2s; box-shadow: 0 1px 3px rgba(0, 0, 0, .15);
}
.switch.on { background: var(--primary); }
.switch.on .knob { left: 24px; }

/* ===== 设置页 ===== */
.settings-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; align-items: start; }
.hint { font-size: 12px; color: var(--text-3); margin-top: 14px; line-height: 1.6; text-align: center; }

/* ===== Toast ===== */
.toast {
  position: fixed; top: 20px; right: 20px; background: #1d2129; color: #fff;
  padding: 11px 18px; border-radius: 8px; font-size: 13px; opacity: 0;
  transform: translateY(-8px); transition: all .25s; z-index: 99; pointer-events: none;
}
.toast.show { opacity: 1; transform: translateY(0); }
.toast.err { background: #d03050; }

@media (max-width: 960px) {
  .cards { grid-template-columns: repeat(2, 1fr); }
  .row2 { grid-template-columns: 1fr; }
  .settings-grid { grid-template-columns: 1fr; }
  .login-card { width: 92vw; }
}
</style>
</head>
<body>

<!-- ================= 登录 ================= -->
<div class="login-wrap" id="view-login">
  <div class="blob-1"></div>
  <div class="blob-2"></div>
  <div class="login-card">
    <div class="login-brand">
      <div class="brand-logo">M</div>
      <div>
        <div class="brand-name">VPS 服务器管理系统</div>
        <div class="brand-sub">VPS 服务器管理面板 · Cloudflare Workers + D1</div>
      </div>
    </div>
    <form id="login-form">
      <div class="form-item">
        <input class="form-input" id="login-username" autocomplete="username" placeholder="请输入用户名">
      </div>
      <div class="form-item">
        <input class="form-input" id="login-password" type="password" autocomplete="current-password" placeholder="请输入密码">
      </div>
      <div class="form-item captcha-row" id="captcha-row">
        <input class="form-input" id="login-captcha" placeholder="验证码" maxlength="4" autocomplete="off">
        <div class="captcha-box" id="captcha-box" onclick="loadCaptcha()" title="点击刷新"></div>
      </div>
      <div class="login-row">
        <label class="remember"><input type="checkbox" id="login-remember" checked> 记住我</label>
        <a class="forgot" href="javascript:void(0)" onclick="showForgot()">忘记密码？</a>
      </div>
      <button class="btn-pill" type="submit">确 认</button>
    </form>
    <p class="hint">mevps · Powered by CF Workers + D1</p>
  </div>
</div>

<!-- ================= 主界面 ================= -->
<div class="app" id="view-app" style="display:none">
  <aside class="sidebar">
    <div class="logo">
      <div class="logo-badge">M</div>
      <div>
        <div class="logo-name" id="brand-name">mevps</div>
        <div class="logo-sub">VPS 管理系统</div>
      </div>
    </div>
    <nav class="menu">
      <div class="menu-item active" data-page="dashboard" onclick="switchPage('dashboard')">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="9" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="14" y="12" width="7" height="9" rx="1"/><rect x="3" y="16" width="7" height="5" rx="1"/></svg>
        数据仪表盘
      </div>
      <div class="menu-item" data-page="vps" onclick="switchPage('vps')">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="7" rx="2"/><rect x="3" y="13" width="18" height="7" rx="2"/><circle cx="7.5" cy="7.5" r="1" fill="currentColor"/><circle cx="7.5" cy="16.5" r="1" fill="currentColor"/></svg>
        VPS 管理
      </div>
      <div class="menu-item" data-page="logs" onclick="switchPage('logs')">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M8 6h13M8 12h13M8 18h13"/><circle cx="4" cy="6" r="1" fill="currentColor"/><circle cx="4" cy="12" r="1" fill="currentColor"/><circle cx="4" cy="18" r="1" fill="currentColor"/></svg>
        操作日志
      </div>
      <div class="menu-item" data-page="settings" onclick="switchPage('settings')">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 8h10M18 8h2M4 16h2M10 16h10"/><circle cx="16" cy="8" r="2"/><circle cx="8" cy="16" r="2"/></svg>
        系统设置
      </div>
    </nav>
    <div class="sidebar-foot">mevps · Powered by CF Workers + D1</div>
  </aside>

  <div class="main">
    <div class="header">
      <div class="title" id="page-title">数据仪表盘</div>
      <div class="user-chip">
        <div class="avatar" id="user-avatar">A</div>
        <span id="user-name">admin</span>
        <button class="btn btn-ghost btn-sm" onclick="logout()">退出</button>
      </div>
    </div>

    <div class="content">
      <!-- 仪表盘 -->
      <section id="page-dashboard">
        <div class="card greet">
          <div>
            <div class="greet-title" id="greet-text">欢迎回来</div>
            <div class="greet-sub">祝你今天工作顺利，机器稳定运行</div>
          </div>
          <div class="greet-meta">mevps · VPS 服务器管理面板</div>
        </div>
        <div class="cards">
          <div class="stat-card g1">
            <div class="t"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="7" rx="2"/><rect x="3" y="13" width="18" height="7" rx="2"/></svg>VPS 总数</div>
            <div class="n" id="d-total">0</div>
            <div class="s">全部纳管资产</div>
          </div>
          <div class="stat-card g2">
            <div class="t"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/></svg>运行中</div>
            <div class="n" id="d-running">0</div>
            <div class="s">状态正常的机器</div>
          </div>
          <div class="stat-card g4">
            <div class="t"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/></svg>即将到期</div>
            <div class="n" id="d-expiring">0</div>
            <div class="s">按设置的提醒天数统计</div>
          </div>
          <div class="stat-card g3">
            <div class="t"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 3v18M7 7h7a3 3 0 010 6H8a3 3 0 000 6h8"/></svg>月度支出</div>
            <div class="n" id="d-monthly">¥0</div>
            <div class="s">所有 VPS 月付合计</div>
          </div>
        </div>

        <div class="row2">
          <div class="card">
            <div class="section-title">服务商分布</div>
            <div id="d-providers"></div>
          </div>
          <div class="card">
            <div class="section-title">状态分布</div>
            <div class="donut-wrap">
              <div class="donut" id="d-donut">
                <div class="donut-hole"><div class="n" id="d-donut-n">0</div><div class="t">VPS 总数</div></div>
              </div>
              <div id="d-legend"></div>
            </div>
          </div>
        </div>

        <div class="card card-block">
          <div class="section-title">最近操作</div>
          <div id="d-logs"></div>
        </div>
      </section>

      <!-- VPS 管理 -->
      <section id="page-vps" style="display:none">
        <div class="toolbar">
          <input class="form-input" id="vps-q" placeholder="搜索名称 / IP / 服务商 / 机房" style="width:260px" onkeydown="if(event.key==='Enter'){searchVps()}">
          <select class="form-input" id="vps-status" onchange="searchVps()">
            <option value="">全部状态</option>
            <option value="running">运行中</option>
            <option value="stopped">已停止</option>
            <option value="maintenance">维护中</option>
          </select>
          <button class="btn btn-ghost" onclick="searchVps()">查询</button>
          <div class="spacer"></div>
          <button class="btn btn-primary" onclick="showVpsModal()">＋ 新增 VPS</button>
        </div>
        <table>
          <thead>
            <tr>
              <th>名称</th><th>IP 地址</th><th>服务商 / 机房</th><th>配置</th><th>系统</th>
              <th>状态</th><th>到期时间</th><th>价格/月</th><th>操作</th>
            </tr>
          </thead>
          <tbody id="vps-tbody"></tbody>
        </table>
        <div class="pager" id="vps-pager"></div>
      </section>

      <!-- 操作日志 -->
      <section id="page-logs" style="display:none">
        <div class="toolbar">
          <select class="form-input" id="log-action" onchange="searchLogs()">
            <option value="">全部操作</option>
            <option value="登录">登录相关</option>
            <option value="新增">新增</option>
            <option value="编辑">编辑</option>
            <option value="删除">删除</option>
            <option value="设置">修改设置</option>
            <option value="密码">修改密码</option>
          </select>
          <div class="spacer"></div>
          <button class="btn btn-ghost" onclick="searchLogs()">刷新</button>
        </div>
        <table>
          <thead>
            <tr><th>时间</th><th>操作人</th><th>操作</th><th>对象</th><th>详情</th><th>IP</th></tr>
          </thead>
          <tbody id="log-tbody"></tbody>
        </table>
        <div class="pager" id="log-pager"></div>
      </section>

      <!-- 系统设置 -->
      <section id="page-settings" style="display:none">
        <div class="settings-grid">
          <div class="card">
            <div class="section-title">系统变量</div>
            <p class="hint-l" style="margin:-8px 0 14px">所有变量统一存储于 D1 数据库，保存后立即生效。</p>
            <div class="form-item"><label>站点名称</label><input class="form-input" id="s-site_name"></div>
            <div class="form-item"><label>管理员邮箱</label><input class="form-input" id="s-admin_email"></div>
            <div class="form-item"><label>到期提醒天数</label><input class="form-input" id="s-expire_warn_days" type="number" min="1" max="90"></div>
            <button class="btn btn-primary" onclick="saveSettings()">保存设置</button>
          </div>
          <div class="card">
            <div class="section-title">安全与登录</div>
            <div class="switch-row">
              <div>
                <div class="switch-label">启用登录验证码</div>
                <div class="hint-l">登录时需要输入图片验证码</div>
              </div>
              <div class="switch on" id="sw-captcha_enabled" onclick="toggleSwitch('captcha_enabled')"><div class="knob"></div></div>
            </div>
            <div class="switch-row">
              <div>
                <div class="switch-label">启用邮箱找回密码</div>
                <div class="hint-l">忘记密码时通过管理员邮箱发送重置链接</div>
              </div>
              <div class="switch on" id="sw-forgot_enabled" onclick="toggleSwitch('forgot_enabled')"><div class="knob"></div></div>
            </div>
          </div>
          <div class="card">
            <div class="section-title">Outlook (Graph API) 设置</div>
            <p class="hint-l" style="margin:-8px 0 14px">用于「忘记密码」时通过 Outlook 发送重置邮件，需在 Azure 注册应用并授权 Mail.Send。</p>
            <div class="form-item"><label>租户 ID (Tenant ID)</label><input class="form-input" id="s-outlook_tenant_id" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"></div>
            <div class="form-item"><label>客户端 ID (Client ID)</label><input class="form-input" id="s-outlook_client_id" placeholder="Application (client) ID"></div>
            <div class="form-item"><label>客户端机密 (Client Secret)</label><input class="form-input" id="s-outlook_client_secret" type="password" placeholder="Client Secret Value"></div>
            <div class="form-item"><label>发件邮箱 (Send As)</label><input class="form-input" id="s-outlook_sender" placeholder="sender@yourdomain.com"></div>
            <button class="btn btn-primary" onclick="saveSettings()">保存设置</button>
          </div>
          <div class="card">
            <div class="section-title">修改密码</div>
            <div class="form-item"><label>当前密码</label><input class="form-input" id="p-old" type="password" autocomplete="current-password"></div>
            <div class="form-item"><label>新密码</label><input class="form-input" id="p-new" type="password" autocomplete="new-password" placeholder="至少 6 位"></div>
            <div class="form-item"><label>确认新密码</label><input class="form-input" id="p-new2" type="password" autocomplete="new-password"></div>
            <button class="btn btn-primary" onclick="changePassword()">修改密码</button>
          </div>
        </div>
      </section>
    </div>
  </div>
</div>

<!-- VPS 弹窗 -->
<div class="modal-mask" id="vps-mask">
  <div class="modal">
    <div class="modal-title" id="vps-modal-title">新增 VPS</div>
    <div class="form-grid">
      <div class="form-item"><label>名称 *</label><input class="form-input" id="f-name" placeholder="如：搬瓦工-香港-01"></div>
      <div class="form-item"><label>IP 地址</label><input class="form-input" id="f-ip" placeholder="如：1.2.3.4"></div>
      <div class="form-item"><label>服务商</label><input class="form-input" id="f-provider" placeholder="如：BandwagonHost"></div>
      <div class="form-item"><label>机房位置</label><input class="form-input" id="f-location" placeholder="如：香港 HK"></div>
      <div class="form-item"><label>CPU</label><input class="form-input" id="f-cpu" placeholder="如：2 核"></div>
      <div class="form-item"><label>内存</label><input class="form-input" id="f-memory" placeholder="如：4 GB"></div>
      <div class="form-item"><label>磁盘</label><input class="form-input" id="f-disk" placeholder="如：60 GB SSD"></div>
      <div class="form-item"><label>操作系统</label><input class="form-input" id="f-os" placeholder="如：Ubuntu 24.04"></div>
      <div class="form-item"><label>状态</label>
        <select class="form-input" id="f-status">
          <option value="running">运行中</option>
          <option value="stopped">已停止</option>
          <option value="maintenance">维护中</option>
        </select>
      </div>
      <div class="form-item"><label>到期时间</label><input class="form-input" id="f-expire_date" type="date"></div>
      <div class="form-item"><label>价格（元/月）</label><input class="form-input" id="f-price" type="number" min="0" step="0.01" placeholder="如：35"></div>
      <div class="form-item"><label>备注</label><input class="form-input" id="f-note" placeholder="选填"></div>
    </div>
    <div class="modal-foot">
      <button class="btn btn-ghost" onclick="hideVpsModal()">取消</button>
      <button class="btn btn-primary" onclick="saveVps()">保存</button>
    </div>
  </div>
</div>

<!-- 忘记密码弹窗 -->
<div class="modal-mask" id="forgot-mask">
  <div class="modal modal-sm">
    <div class="modal-title">找回密码</div>
    <p class="hint-l" style="margin:-8px 0 16px">输入管理员邮箱，我们会发送一封密码重置邮件（30 分钟内有效）。</p>
    <div class="form-item"><input class="form-input" id="forgot-email" type="email" placeholder="请输入邮箱"></div>
    <div class="modal-foot">
      <button class="btn btn-ghost" onclick="hideForgot()">取消</button>
      <button class="btn btn-primary" onclick="submitForgot()">发送邮件</button>
    </div>
  </div>
</div>

<!-- 重置密码弹窗 -->
<div class="modal-mask" id="reset-mask">
  <div class="modal modal-sm">
    <div class="modal-title">重置密码</div>
    <div class="form-item"><label>新密码</label><input class="form-input" id="reset-new" type="password" placeholder="至少 6 位"></div>
    <div class="form-item"><label>确认新密码</label><input class="form-input" id="reset-new2" type="password"></div>
    <div class="modal-foot">
      <button class="btn btn-primary" onclick="submitReset()">重置密码</button>
    </div>
  </div>
</div>

<div class="toast" id="toast"></div>

<script>
'use strict';

var state = {
  user: null,
  page: 'dashboard',
  vpsPage: 1,
  vpsList: [],
  logPage: 1,
  editingId: 0,
  captchaId: '',
  resetToken: ''
};

function $(id) { return document.getElementById(id); }

function esc(s) {
  return String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function toast(msg, isErr) {
  var el = $('toast');
  el.textContent = msg;
  el.className = 'toast show' + (isErr ? ' err' : '');
  setTimeout(function () { el.className = 'toast' + (isErr ? ' err' : ''); }, 2400);
}

function fmtTime(ms) {
  if (!ms) return '-';
  var d = new Date(ms);
  function p(n) { return (n < 10 ? '0' : '') + n; }
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
}

function daysLeft(dateStr) {
  if (!dateStr) return null;
  var t = new Date(dateStr + 'T23:59:59').getTime() - Date.now();
  return Math.ceil(t / 86400000);
}

function api(path, method, body) {
  var opt = { method: method || 'GET' };
  if (body) {
    opt.headers = { 'content-type': 'application/json' };
    opt.body = JSON.stringify(body);
  }
  return fetch(path, opt).then(function (res) {
    return res.json().catch(function () { return {}; }).then(function (data) {
      if (res.status === 401) {
        showLogin();
        throw new Error(data.error || '请先登录');
      }
      if (!res.ok) throw new Error(data.error || '请求失败');
      return data;
    });
  });
}

/* ---------- 视图切换 ---------- */

function showLogin() {
  $('view-login').style.display = 'flex';
  $('view-app').style.display = 'none';
}

function showApp() {
  $('view-login').style.display = 'none';
  $('view-app').style.display = 'flex';
  var name = state.user || 'admin';
  $('user-name').textContent = name;
  $('user-avatar').textContent = name.charAt(0).toUpperCase();
  var h = new Date().getHours();
  var greet = h < 6 ? '凌晨好' : (h < 9 ? '早安' : (h < 12 ? '上午好' : (h < 14 ? '中午好' : (h < 18 ? '下午好' : '晚上好'))));
  $('greet-text').textContent = greet + '，' + name + '，今天又是充满活力的一天';
  switchPage('dashboard');
}

var PAGE_TITLES = { dashboard: '数据仪表盘', vps: 'VPS 管理', logs: '操作日志', settings: '系统设置' };

function switchPage(page) {
  state.page = page;
  var items = document.querySelectorAll('.menu-item');
  for (var i = 0; i < items.length; i++) {
    items[i].className = 'menu-item' + (items[i].getAttribute('data-page') === page ? ' active' : '');
  }
  var secs = ['dashboard', 'vps', 'logs', 'settings'];
  for (var j = 0; j < secs.length; j++) {
    $('page-' + secs[j]).style.display = secs[j] === page ? '' : 'none';
  }
  $('page-title').textContent = PAGE_TITLES[page] || '';
  if (page === 'dashboard') renderDashboard();
  if (page === 'vps') renderVps();
  if (page === 'logs') renderLogs();
  if (page === 'settings') renderSettings();
}

/* ---------- 仪表盘 ---------- */

function renderDashboard() {
  api('/api/dashboard').then(function (d) {
    $('d-total').textContent = d.total;
    $('d-running').textContent = d.running;
    $('d-expiring').textContent = d.expiring;
    $('d-monthly').textContent = '¥' + (Math.round(d.monthly * 100) / 100);

    var maxP = 1;
    for (var i = 0; i < d.providers.length; i++) {
      if (d.providers[i].c > maxP) maxP = d.providers[i].c;
    }
    var ph = '';
    if (!d.providers.length) {
      ph = '<div class="empty">暂无数据</div>';
    } else {
      for (var j = 0; j < d.providers.length; j++) {
        var p = d.providers[j];
        ph += '<div class="bar-row"><div class="bar-label" title="' + esc(p.provider) + '">' + esc(p.provider) + '</div>' +
          '<div class="bar-track"><div class="bar-fill" style="width:' + Math.round(p.c / maxP * 100) + '%"></div></div>' +
          '<div class="bar-num">' + p.c + '</div></div>';
      }
    }
    $('d-providers').innerHTML = ph;

    var colors = ['#646cff', '#2dd4bf', '#f59e0b'];
    var total = 0;
    for (var k = 0; k < d.status_dist.length; k++) total += d.status_dist[k].value;
    $('d-donut-n').textContent = total;
    var acc = 0;
    var segs = [];
    var lh = '';
    for (var m = 0; m < d.status_dist.length; m++) {
      var item = d.status_dist[m];
      var deg = total > 0 ? item.value / total * 360 : 0;
      segs.push(colors[m] + ' ' + acc + 'deg ' + (acc + deg) + 'deg');
      acc += deg;
      lh += '<div class="legend-item"><span class="legend-dot" style="background:' + colors[m] + '"></span>' +
        esc(item.label) + ' · ' + item.value + ' 台</div>';
    }
    $('d-donut').style.background = 'conic-gradient(' + segs.join(',') + ')';
    $('d-legend').innerHTML = lh;

    var logsHtml = '';
    if (!d.recent_logs.length) {
      logsHtml = '<div class="empty">暂无操作记录</div>';
    } else {
      for (var n = 0; n < d.recent_logs.length; n++) {
        var lg = d.recent_logs[n];
        logsHtml += '<div class="log-line"><span class="log-time">' + fmtTime(lg.created_at) + '</span>' +
          '<span class="log-tag">' + esc(lg.action) + '</span>' +
          '<span>' + esc(lg.username) + ' ' + esc(lg.target) + (lg.detail ? ' · ' + esc(lg.detail) : '') + '</span></div>';
      }
    }
    $('d-logs').innerHTML = logsHtml;
  }).catch(function (e) {
    toast(e.message, true);
  });
}

/* ---------- VPS 管理 ---------- */

function searchVps() {
  state.vpsPage = 1;
  renderVps();
}

function gotoVpsPage(p) {
  state.vpsPage = p;
  renderVps();
}

function renderVps() {
  var q = encodeURIComponent($('vps-q').value.trim());
  var st = encodeURIComponent($('vps-status').value);
  api('/api/vps?q=' + q + '&status=' + st + '&page=' + state.vpsPage).then(function (d) {
    state.vpsList = d.list;
    var rows = '';
    if (!d.list.length) {
      rows = '<tr><td colspan="9"><div class="empty">暂无 VPS 记录，点击右上角「新增 VPS」添加</div></td></tr>';
    } else {
      for (var i = 0; i < d.list.length; i++) {
        var v = d.list[i];
        var badge = v.status === 'running' ? 'badge-running' : (v.status === 'stopped' ? 'badge-stopped' : 'badge-maintenance');
        var statusText = v.status === 'running' ? '运行中' : (v.status === 'stopped' ? '已停止' : '维护中');
        var dl = daysLeft(v.expire_date);
        var expireCell = '-';
        if (v.expire_date) {
          if (dl < 0) expireCell = esc(v.expire_date) + ' <span class="badge badge-expired">已过期</span>';
          else if (dl <= 7) expireCell = esc(v.expire_date) + ' <span class="badge badge-warn">' + dl + ' 天</span>';
          else expireCell = esc(v.expire_date);
        }
        rows += '<tr>' +
          '<td><div class="cell-main">' + esc(v.name) + '</div>' + (v.note ? '<div class="cell-sub">' + esc(v.note) + '</div>' : '') + '</td>' +
          '<td>' + esc(v.ip || '-') + '</td>' +
          '<td><div>' + esc(v.provider || '-') + '</div><div class="cell-sub">' + esc(v.location || '') + '</div></td>' +
          '<td>' + esc([v.cpu, v.memory, v.disk].filter(Boolean).join(' / ') || '-') + '</td>' +
          '<td>' + esc(v.os || '-') + '</td>' +
          '<td><span class="badge ' + badge + '"><span class="dot"></span>' + statusText + '</span></td>' +
          '<td>' + expireCell + '</td>' +
          '<td>¥' + (v.price || 0) + '</td>' +
          '<td><button class="btn btn-ghost btn-sm" onclick="editVps(' + v.id + ')">编辑</button> ' +
          '<button class="btn btn-danger btn-sm" onclick="deleteVps(' + v.id + ')">删除</button></td>' +
          '</tr>';
      }
    }
    $('vps-tbody').innerHTML = rows;

    var pages = Math.max(1, Math.ceil(d.total / d.size));
    $('vps-pager').innerHTML =
      '<span class="pager-info">共 ' + d.total + ' 条</span>' +
      '<button class="btn btn-ghost btn-sm" ' + (d.page <= 1 ? 'disabled' : '') + ' onclick="gotoVpsPage(' + (d.page - 1) + ')">上一页</button>' +
      '<span class="pager-info">' + d.page + ' / ' + pages + '</span>' +
      '<button class="btn btn-ghost btn-sm" ' + (d.page >= pages ? 'disabled' : '') + ' onclick="gotoVpsPage(' + (d.page + 1) + ')">下一页</button>';
  }).catch(function (e) {
    toast(e.message, true);
  });
}

function showVpsModal() {
  state.editingId = 0;
  $('vps-modal-title').textContent = '新增 VPS';
  var ids = ['name', 'ip', 'provider', 'location', 'cpu', 'memory', 'disk', 'os', 'expire_date', 'price', 'note'];
  for (var i = 0; i < ids.length; i++) $('f-' + ids[i]).value = '';
  $('f-status').value = 'running';
  $('vps-mask').style.display = 'flex';
}

function hideVpsModal() {
  $('vps-mask').style.display = 'none';
}

function editVps(id) {
  var v = null;
  for (var i = 0; i < state.vpsList.length; i++) {
    if (state.vpsList[i].id === id) v = state.vpsList[i];
  }
  if (!v) return;
  state.editingId = id;
  $('vps-modal-title').textContent = '编辑 VPS';
  $('f-name').value = v.name || '';
  $('f-ip').value = v.ip || '';
  $('f-provider').value = v.provider || '';
  $('f-location').value = v.location || '';
  $('f-cpu').value = v.cpu || '';
  $('f-memory').value = v.memory || '';
  $('f-disk').value = v.disk || '';
  $('f-os').value = v.os || '';
  $('f-status').value = v.status || 'running';
  $('f-expire_date').value = v.expire_date || '';
  $('f-price').value = v.price || '';
  $('f-note').value = v.note || '';
  $('vps-mask').style.display = 'flex';
}

function saveVps() {
  var body = {
    name: $('f-name').value,
    ip: $('f-ip').value,
    provider: $('f-provider').value,
    location: $('f-location').value,
    cpu: $('f-cpu').value,
    memory: $('f-memory').value,
    disk: $('f-disk').value,
    os: $('f-os').value,
    status: $('f-status').value,
    expire_date: $('f-expire_date').value,
    price: $('f-price').value,
    note: $('f-note').value
  };
  if (!body.name.trim()) { toast('请填写名称', true); return; }
  var req = state.editingId
    ? api('/api/vps/' + state.editingId, 'PUT', body)
    : api('/api/vps', 'POST', body);
  req.then(function () {
    hideVpsModal();
    toast('保存成功');
    renderVps();
  }).catch(function (e) {
    toast(e.message, true);
  });
}

function deleteVps(id) {
  if (!confirm('确定删除这条 VPS 记录吗？')) return;
  api('/api/vps/' + id, 'DELETE').then(function () {
    toast('已删除');
    renderVps();
  }).catch(function (e) {
    toast(e.message, true);
  });
}

/* ---------- 操作日志 ---------- */

function searchLogs() {
  state.logPage = 1;
  renderLogs();
}

function gotoLogPage(p) {
  state.logPage = p;
  renderLogs();
}

function renderLogs() {
  var act = encodeURIComponent($('log-action').value.trim());
  api('/api/logs?action=' + act + '&page=' + state.logPage).then(function (d) {
    var rows = '';
    if (!d.list.length) {
      rows = '<tr><td colspan="6"><div class="empty">暂无日志</div></td></tr>';
    } else {
      for (var i = 0; i < d.list.length; i++) {
        var lg = d.list[i];
        rows += '<tr>' +
          '<td>' + fmtTime(lg.created_at) + '</td>' +
          '<td>' + esc(lg.username) + '</td>' +
          '<td><span class="log-tag">' + esc(lg.action) + '</span></td>' +
          '<td>' + esc(lg.target || '-') + '</td>' +
          '<td>' + esc(lg.detail || '-') + '</td>' +
          '<td>' + esc(lg.ip) + '</td>' +
          '</tr>';
      }
    }
    $('log-tbody').innerHTML = rows;

    var pages = Math.max(1, Math.ceil(d.total / d.size));
    $('log-pager').innerHTML =
      '<span class="pager-info">共 ' + d.total + ' 条</span>' +
      '<button class="btn btn-ghost btn-sm" ' + (d.page <= 1 ? 'disabled' : '') + ' onclick="gotoLogPage(' + (d.page - 1) + ')">上一页</button>' +
      '<span class="pager-info">' + d.page + ' / ' + pages + '</span>' +
      '<button class="btn btn-ghost btn-sm" ' + (d.page >= pages ? 'disabled' : '') + ' onclick="gotoLogPage(' + (d.page + 1) + ')">下一页</button>';
  }).catch(function (e) {
    toast(e.message, true);
  });
}

/* ---------- 系统设置 ---------- */

function renderSettings() {
  api('/api/settings').then(function (s) {
    $('s-site_name').value = s.site_name || '';
    $('s-admin_email').value = s.admin_email || '';
    $('s-expire_warn_days').value = s.expire_warn_days || '7';
    $('s-outlook_tenant_id').value = s.outlook_tenant_id || '';
    $('s-outlook_client_id').value = s.outlook_client_id || '';
    $('s-outlook_client_secret').value = s.outlook_client_secret || '';
    $('s-outlook_sender').value = s.outlook_sender || '';
    $('sw-captcha_enabled').classList.toggle('on', s.captcha_enabled === '1');
    $('sw-forgot_enabled').classList.toggle('on', s.forgot_enabled === '1');
    if (s.site_name) {
      $('brand-name').textContent = s.site_name;
      document.title = s.site_name + ' · VPS 管理系统';
    }
  }).catch(function (e) {
    toast(e.message, true);
  });
}

function saveSettings() {
  api('/api/settings', 'POST', {
    site_name: $('s-site_name').value,
    admin_email: $('s-admin_email').value,
    expire_warn_days: $('s-expire_warn_days').value,
    captcha_enabled: $('sw-captcha_enabled').classList.contains('on') ? '1' : '0',
    forgot_enabled: $('sw-forgot_enabled').classList.contains('on') ? '1' : '0',
    outlook_tenant_id: $('s-outlook_tenant_id').value.trim(),
    outlook_client_id: $('s-outlook_client_id').value.trim(),
    outlook_client_secret: $('s-outlook_client_secret').value.trim(),
    outlook_sender: $('s-outlook_sender').value.trim()
  }).then(function () {
    toast('设置已保存');
    renderSettings();
  }).catch(function (e) {
    toast(e.message, true);
  });
}

function changePassword() {
  var oldP = $('p-old').value;
  var newP = $('p-new').value;
  var newP2 = $('p-new2').value;
  if (!oldP || !newP) { toast('请填写完整', true); return; }
  if (newP !== newP2) { toast('两次输入的新密码不一致', true); return; }
  api('/api/password', 'POST', { old_password: oldP, new_password: newP }).then(function () {
    toast('密码已修改');
    $('p-old').value = '';
    $('p-new').value = '';
    $('p-new2').value = '';
  }).catch(function (e) {
    toast(e.message, true);
  });
}

/* ---------- 登录 / 退出 ---------- */

function logout() {
  api('/api/logout', 'POST').then(function () {
    state.user = null;
    showLogin();
  }).catch(function () {
    showLogin();
  });
}

function loadCaptcha() {
  api('/api/captcha').then(function (d) {
    state.captchaId = d.id;
    $('captcha-box').innerHTML = d.svg;
    $('login-captcha').value = '';
    $('captcha-row').style.display = d.enabled ? 'flex' : 'none';
  }).catch(function () {});
}

function showForgot() {
  $('forgot-email').value = '';
  $('forgot-mask').style.display = 'flex';
}

function hideForgot() {
  $('forgot-mask').style.display = 'none';
}

function submitForgot() {
  var email = $('forgot-email').value.trim();
  if (!email) { toast('请输入邮箱', true); return; }
  api('/api/forgot-password', 'POST', { email: email }).then(function (d) {
    hideForgot();
    toast(d.message || '如果该邮箱已登记，重置邮件已发送');
  }).catch(function (e) {
    toast(e.message, true);
  });
}

function submitReset() {
  var p1 = $('reset-new').value;
  var p2 = $('reset-new2').value;
  if (p1.length < 6) { toast('新密码至少 6 位', true); return; }
  if (p1 !== p2) { toast('两次输入的新密码不一致', true); return; }
  api('/api/reset-password', 'POST', { token: state.resetToken, new_password: p1 }).then(function () {
    $('reset-mask').style.display = 'none';
    toast('密码已重置，请使用新密码登录');
    state.resetToken = '';
    window.history.replaceState({}, '', location.pathname);
  }).catch(function (e) {
    toast(e.message, true);
  });
}

function toggleSwitch(key) {
  $('sw-' + key).classList.toggle('on');
}

$('login-form').addEventListener('submit', function (e) {
  e.preventDefault();
  api('/api/login', 'POST', {
    username: $('login-username').value.trim(),
    password: $('login-password').value,
    captcha_id: state.captchaId,
    captcha_code: $('login-captcha').value.trim()
  }).then(function (d) {
    state.user = d.username;
    showApp();
  }).catch(function (err) {
    toast(err.message, true);
    loadCaptcha();
  });
});

/* ---------- 启动 ---------- */

loadCaptcha();

var resetToken = new URLSearchParams(location.search).get('reset');
if (resetToken) {
  state.resetToken = resetToken;
  $('reset-mask').style.display = 'flex';
}

api('/api/me').then(function (d) {
  state.user = d.username;
  showApp();
}).catch(function () {
  showLogin();
});
</script>
</body>
</html>`;
