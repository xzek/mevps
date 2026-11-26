/**
 * mevps - VPS 服务器管理面板
 * Cloudflare Workers + D1 全栈单文件（API + 前端）
 *
 * 部署方式（Cloudflare 控制台浏览器操作）：
 *   1. 创建 D1 数据库，名字固定：mymail
 *   2. 创建 Worker，把本文件全部内容粘贴进编辑器并部署
 *   3. 在 Worker 的 Settings -> Bindings 添加 D1 绑定：
 *      Variable name（变量名）固定为  db
 *      D1 database 选择 mymail
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
    admin_email: 'mymail',
    expire_warn_days: '7',
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
    const allowed = ['site_name', 'admin_email', 'expire_warn_days'];
    const changed = [];
    for (const key of allowed) {
      if (b[key] !== undefined) {
        let value = String(b[key]).trim();
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
  --primary: #18a058;
  --primary-hover: #36ad60;
  --primary-light: #e8f7f0;
  --bg: #f5f7f9;
  --border: #e5e6eb;
  --text: #262626;
  --text-2: #595959;
  --text-3: #8c8c8c;
}
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; height: 100%; }
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", "Microsoft YaHei", sans-serif;
  background: var(--bg); color: var(--text); font-size: 14px;
}
button { font-family: inherit; }
a { color: var(--primary); text-decoration: none; }

/* ===== 登录页 ===== */
.login-wrap { display: flex; height: 100vh; }
.login-brand {
  flex: 1; background: linear-gradient(135deg, #18a058 0%, #0c7a44 60%, #095c33 100%);
  color: #fff; padding: 64px; display: flex; flex-direction: column; justify-content: center;
}
.login-brand h1 { font-size: 42px; margin: 0 0 12px; letter-spacing: 1px; }
.login-brand p { font-size: 16px; opacity: .92; margin: 6px 0; }
.login-brand .feature { margin-top: 34px; }
.login-brand .feature div {
  display: flex; align-items: center; gap: 10px; margin: 12px 0; font-size: 15px; opacity: .95;
}
.login-brand .feature .tick {
  width: 22px; height: 22px; border-radius: 50%; background: rgba(255,255,255,.22);
  display: inline-flex; align-items: center; justify-content: center; font-size: 13px;
}
.login-box { width: 460px; background: #fff; display: flex; align-items: center; justify-content: center; }
.login-card { width: 320px; }
.login-card h2 { margin: 0 0 6px; font-size: 24px; }
.login-card .sub { color: var(--text-3); margin: 0 0 28px; font-size: 13px; }

/* ===== 表单 ===== */
.form-input {
  width: 100%; padding: 10px 12px; border: 1px solid var(--border); border-radius: 6px;
  font-size: 14px; outline: none; background: #fff; transition: border .15s;
}
.form-input:focus { border-color: var(--primary); box-shadow: 0 0 0 2px rgba(24,160,88,.12); }
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
.btn-ghost { background: #fff; border-color: var(--border); color: var(--text-2); }
.btn-ghost:hover { border-color: var(--primary); color: var(--primary); }
.btn-danger { background: #fff; border-color: #ffccc7; color: #d03050; }
.btn-danger:hover { background: #fff1f0; }
.btn-block { width: 100%; padding: 11px; }
.btn:disabled { opacity: .5; cursor: not-allowed; }
.btn-sm { padding: 4px 12px; font-size: 12px; }

/* ===== 布局 ===== */
.app { display: flex; height: 100vh; overflow: hidden; }
.sidebar {
  width: 220px; background: #fff; border-right: 1px solid var(--border);
  display: flex; flex-direction: column; flex-shrink: 0;
}
.logo {
  height: 56px; display: flex; align-items: center; gap: 10px; padding: 0 16px;
  border-bottom: 1px solid var(--border);
}
.logo-badge {
  width: 32px; height: 32px; border-radius: 8px; background: var(--primary); color: #fff;
  display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 16px;
}
.logo-name { font-size: 17px; font-weight: 700; letter-spacing: .5px; }
.logo-sub { font-size: 10px; color: var(--text-3); margin-top: 1px; }
.menu { padding: 12px 8px; flex: 1; }
.menu-item {
  display: flex; align-items: center; gap: 10px; padding: 10px 12px; border-radius: 6px;
  cursor: pointer; color: var(--text-2); margin-bottom: 4px; font-size: 14px; user-select: none;
}
.menu-item:hover { background: #f2f3f5; }
.menu-item.active { background: var(--primary-light); color: var(--primary); font-weight: 600; }
.menu-item svg { width: 18px; height: 18px; flex-shrink: 0; }
.sidebar-foot { padding: 14px 16px; border-top: 1px solid var(--border); font-size: 11px; color: var(--text-3); }

.main { flex: 1; display: flex; flex-direction: column; overflow: hidden; }
.header {
  height: 56px; background: #fff; border-bottom: 1px solid var(--border);
  display: flex; align-items: center; justify-content: space-between; padding: 0 20px; flex-shrink: 0;
}
.header .title { font-size: 16px; font-weight: 600; }
.user-chip {
  display: flex; align-items: center; gap: 10px; font-size: 13px; color: var(--text-2);
}
.avatar {
  width: 32px; height: 32px; border-radius: 50%; background: var(--primary-light); color: var(--primary);
  display: flex; align-items: center; justify-content: center; font-weight: 600;
}
.content { flex: 1; overflow: auto; padding: 20px; }

/* ===== 卡片 ===== */
.card { background: #fff; border-radius: 8px; padding: 18px 20px; box-shadow: 0 1px 2px rgba(0,0,0,.03); }
.cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; }
.stat-title { color: var(--text-3); font-size: 13px; display: flex; align-items: center; gap: 8px; }
.stat-value { font-size: 28px; font-weight: 700; margin-top: 8px; }
.stat-sub { font-size: 12px; color: var(--text-3); margin-top: 6px; }
.row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-top: 16px; }
.section-title { font-size: 15px; font-weight: 600; margin: 0 0 14px; }
.card + .card-block { margin-top: 16px; }
.card-block { margin-top: 16px; }

/* ===== 图表 ===== */
.bar-row { display: flex; align-items: center; gap: 10px; margin: 11px 0; }
.bar-label { width: 110px; font-size: 13px; color: var(--text-2); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.bar-track { flex: 1; height: 10px; background: #f2f3f5; border-radius: 5px; overflow: hidden; }
.bar-fill { height: 100%; background: linear-gradient(90deg, #63e2b7, #18a058); border-radius: 5px; }
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
  display: flex; align-items: center; gap: 10px; padding: 9px 0;
  border-bottom: 1px solid #f2f3f5; font-size: 13px;
}
.log-line:last-child { border-bottom: none; }
.log-time { color: var(--text-3); width: 130px; flex-shrink: 0; }
.log-tag {
  background: var(--primary-light); color: var(--primary); border-radius: 4px;
  padding: 1px 8px; font-size: 12px; flex-shrink: 0;
}

/* ===== 表格 ===== */
.toolbar { display: flex; gap: 10px; margin-bottom: 14px; align-items: center; flex-wrap: wrap; }
.toolbar .spacer { flex: 1; }
select.form-input { width: auto; padding: 8px 10px; }
table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; }
th {
  text-align: left; font-weight: 600; font-size: 13px; color: var(--text-2);
  background: #fafafa; padding: 11px 12px; border-bottom: 1px solid var(--border); white-space: nowrap;
}
td { padding: 11px 12px; border-bottom: 1px solid #f2f3f5; font-size: 13px; }
tr:hover td { background: #fafcfb; }
.cell-main { font-weight: 500; }
.cell-sub { color: var(--text-3); font-size: 12px; margin-top: 2px; }
.badge {
  display: inline-flex; align-items: center; gap: 6px; padding: 2px 10px;
  border-radius: 10px; font-size: 12px; white-space: nowrap;
}
.badge .dot { width: 6px; height: 6px; border-radius: 50%; }
.badge-running { background: #e8f7f0; color: #18a058; }
.badge-running .dot { background: #18a058; }
.badge-stopped { background: #f2f3f5; color: #86909c; }
.badge-stopped .dot { background: #86909c; }
.badge-maintenance { background: #fdf3e7; color: #d97706; }
.badge-maintenance .dot { background: #d97706; }
.badge-warn { background: #fff7e6; color: #d97706; }
.badge-expired { background: #fff1f0; color: #d03050; }
.pager { display: flex; gap: 10px; align-items: center; justify-content: flex-end; margin-top: 14px; }
.pager-info { color: var(--text-3); font-size: 13px; }
.empty { text-align: center; color: var(--text-3); padding: 40px 0; }

/* ===== 弹窗 ===== */
.modal-mask {
  position: fixed; inset: 0; background: rgba(0,0,0,.35); display: none;
  align-items: center; justify-content: center; z-index: 50;
}
.modal {
  background: #fff; border-radius: 10px; width: 620px; max-width: 94vw;
  padding: 22px 24px; max-height: 86vh; overflow: auto;
}
.modal-title { font-size: 16px; font-weight: 600; margin: 0 0 18px; }
.modal-foot { display: flex; justify-content: flex-end; gap: 10px; margin-top: 20px; }

/* ===== 设置页 ===== */
.settings-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; align-items: start; }
.hint { font-size: 12px; color: var(--text-3); margin-top: 6px; line-height: 1.6; }

/* ===== Toast ===== */
.toast {
  position: fixed; top: 20px; right: 20px; background: #262626; color: #fff;
  padding: 11px 18px; border-radius: 8px; font-size: 13px; opacity: 0;
  transform: translateY(-8px); transition: all .25s; z-index: 99; pointer-events: none;
}
.toast.show { opacity: 1; transform: translateY(0); }
.toast.err { background: #d03050; }

@media (max-width: 960px) {
  .cards { grid-template-columns: repeat(2, 1fr); }
  .row2 { grid-template-columns: 1fr; }
  .settings-grid { grid-template-columns: 1fr; }
  .login-brand { display: none; }
}
</style>
</head>
<body>

<!-- ================= 登录 ================= -->
<div class="login-wrap" id="view-login">
  <div class="login-brand">
    <h1>mevps</h1>
    <p>VPS 服务器管理面板</p>
    <p>Cloudflare Workers · D1 · 轻量部署</p>
    <div class="feature">
      <div><span class="tick">&#10003;</span> 数据仪表盘，资产状态一目了然</div>
      <div><span class="tick">&#10003;</span> VPS 全生命周期信息管理</div>
      <div><span class="tick">&#10003;</span> 完整操作日志审计</div>
      <div><span class="tick">&#10003;</span> 系统变量统一存储于 D1</div>
    </div>
  </div>
  <div class="login-box">
    <div class="login-card">
      <h2>欢迎回来</h2>
      <p class="sub">登录 mevps 管理系统</p>
      <form id="login-form">
        <div class="form-item">
          <label>用户名</label>
          <input class="form-input" id="login-username" autocomplete="username" placeholder="请输入用户名">
        </div>
        <div class="form-item">
          <label>密码</label>
          <input class="form-input" id="login-password" type="password" autocomplete="current-password" placeholder="请输入密码">
        </div>
        <button class="btn btn-primary btn-block" type="submit">登 录</button>
      </form>
      <p class="hint">默认账号：admin / 123456，登录后请在「系统设置」中及时修改密码。</p>
    </div>
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
        <div class="cards">
          <div class="card"><div class="stat-title">VPS 总数</div><div class="stat-value" id="d-total">0</div><div class="stat-sub">全部纳管资产</div></div>
          <div class="card"><div class="stat-title">运行中</div><div class="stat-value" id="d-running" style="color:#18a058">0</div><div class="stat-sub">状态正常的机器</div></div>
          <div class="card"><div class="stat-title">即将到期</div><div class="stat-value" id="d-expiring" style="color:#d97706">0</div><div class="stat-sub">按设置的提醒天数统计</div></div>
          <div class="card"><div class="stat-title">月度支出</div><div class="stat-value" id="d-monthly">¥0</div><div class="stat-sub">所有 VPS 月付合计</div></div>
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
            <p class="hint" style="margin-top:-6px">所有变量统一存储于 D1 数据库，保存后立即生效。</p>
            <div class="form-item"><label>站点名称</label><input class="form-input" id="s-site_name"></div>
            <div class="form-item"><label>管理员邮箱</label><input class="form-input" id="s-admin_email"></div>
            <div class="form-item"><label>到期提醒天数</label><input class="form-input" id="s-expire_warn_days" type="number" min="1" max="90"></div>
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

<div class="toast" id="toast"></div>

<script>
'use strict';

var state = {
  user: null,
  page: 'dashboard',
  vpsPage: 1,
  vpsList: [],
  logPage: 1,
  editingId: 0
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

    var colors = ['#18a058', '#63e2b7', '#f2c94c'];
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
    expire_warn_days: $('s-expire_warn_days').value
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

$('login-form').addEventListener('submit', function (e) {
  e.preventDefault();
  api('/api/login', 'POST', {
    username: $('login-username').value.trim(),
    password: $('login-password').value
  }).then(function (d) {
    state.user = d.username;
    showApp();
  }).catch(function (err) {
    toast(err.message, true);
  });
});

/* ---------- 启动 ---------- */

api('/api/me').then(function (d) {
  state.user = d.username;
  showApp();
}).catch(function () {
  showLogin();
});
</script>
</body>
</html>`;
