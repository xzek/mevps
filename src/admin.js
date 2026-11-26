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

/* ---------- 移动端侧边栏抽屉 ---------- */
function toggleSidebar() {
  var sb = document.querySelector('.sidebar');
  var mask = $('sidebar-mask');
  var open = !sb.classList.contains('open');
  sb.classList.toggle('open', open);
  mask.classList.toggle('show', open);
}
function closeSidebar() {
  document.querySelector('.sidebar').classList.remove('open');
  $('sidebar-mask').classList.remove('show');
}

function switchPage(page) {
  state.page = page;
  if (window.innerWidth <= 960) closeSidebar();
  var items = document.querySelectorAll('.menu-item');
  for (var i = 0; i < items.length; i++) {
    items[i].classList.toggle('active', items[i].getAttribute('data-page') === page);
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
