# mevps · VPS 服务器管理面板

基于 **Cloudflare Workers + D1 + Static Assets** 的轻量 VPS 服务器管理项目：`src/worker.js` 只写接口，管理界面由 Cloudflare 静态资源在边缘直接提供（更快、不消耗 Worker 请求额度）。

后台面板样式参考 [soybean-admin](https://github.com/soybeanjs/soybean-admin)：清爽绿色主题、侧边栏菜单 + 卡片式内容区。

## 功能

| 左侧菜单 | 功能 |
| --- | --- |
| 数据仪表盘 | VPS 总数 / 运行中 / 即将到期 / 月度支出统计卡片，服务商分布、状态分布图，最近操作记录 |
| VPS 管理 | VPS 资产增删改查：名称、IP、服务商、机房、配置（CPU/内存/磁盘）、系统、状态、到期时间、价格、备注；支持搜索、状态筛选、分页 |
| 操作日志 | 全部操作自动留痕（登录/登出/新增/编辑/删除/改设置/改密码），含操作人、对象、详情、IP，支持筛选与分页 |
| 系统设置 | 系统变量（站点名称、管理员邮箱、到期提醒天数）统一存储于 D1；登录安全开关（图片验证码 / 邮箱找回密码）；Outlook (Graph API) 邮件配置；修改登录密码 |

## 登录安全与密码找回

- **图片验证码**：登录页自带 SVG 图片验证码（点击刷新），可在「系统设置 → 安全与登录」一键开关，验证码开关状态存 D1
- **忘记密码**：点击登录页「忘记密码？」弹窗输入管理员邮箱，系统通过 **Outlook (Graph API)** 发送重置链接（30 分钟有效，一次性令牌）
- **Outlook (Graph API) 配置**：在「系统设置」中填入 Azure 应用的 Tenant ID、Client ID、Client Secret 和发件邮箱即可；需给应用授予 `Mail.Send` 应用权限并同意管理员授权

## 默认登录

- 用户名：`admin`
- 密码：`123456`

> ⚠️ 首次登录后请立即在「系统设置」中修改默认密码。

## 部署（Cloudflare 控制台，浏览器操作）

### 方式 A：连接 Git 仓库部署（推荐，全自动 ✅）

1. Cloudflare 控制台 → Workers 和 Pages → 创建 → **导入存储库 (Import a repository)**，选择 `xzek/mevps`。
2. 保持默认设置直接部署。

就这么简单：`wrangler.toml` 中已声明 D1 绑定（变量名 `db`，数据库名 `mevps`），部署时 Cloudflare 会**自动创建 D1 数据库并完成绑定**（官方 Automatic provisioning 功能）；构建命令会自动执行生成静态资源（`public/`），数据表和初始配置（admin/123456 等）由 Worker 在首次请求时**自动创建写入**。

### 方式 B：命令行部署（wrangler）

```bash
npm install                   # 安装构建依赖（esbuild）
npm run build                 # 生成静态资源 public/（wrangler deploy 时也会自动执行）
npx wrangler d1 create mevps  # 自动建库部署可跳过此步；手动建库则把 database_id 填入 wrangler.toml
npx wrangler deploy
```

> ⚠️ 管理界面已改为 **Static Assets** 静态资源（需构建产物 `public/`），不再支持「控制台编辑器粘贴代码」部署，请使用以上两种方式之一。

### 访问

打开 Worker 的 `*.workers.dev` 域名，首次访问会自动创建数据表并初始化，直接用 `admin / 123456` 登录。

## 项目结构

```
mevps/
├── src/
│   ├── worker.js     # 接口层：/api/* 全部逻辑
│   ├── index.html    # 管理界面 HTML 壳（模板，构建时注入哈希资源引用）
│   ├── admin.css     # 管理界面样式源码
│   ├── admin.js      # 管理界面脚本源码
│   └── _headers      # 静态资源缓存策略
├── build.mjs         # 构建脚本：esbuild 压缩 + 内容哈希文件名 → 输出 public/
├── public/           # 构建产物（静态资源，边缘直接提供，已提交入库）
├── package.json      # 构建依赖（esbuild）
├── wrangler.toml     # 部署配置（Static Assets + D1 绑定）
└── README.md
```

## 数据说明

- **所有系统变量都存在 D1 里**（`settings` 表，key-value 形式），修改后立即生效
- 数据表首次访问自动创建：`users`（账号）、`sessions`（会话）、`vps`（机器资产）、`logs`（操作日志）、`settings`（系统变量）
- 默认系统变量：`site_name=mevps`、`admin_email=mevps`、`expire_warn_days=7`

## 安全提示

- 上线后第一时间修改默认密码 `123456`
- 建议在 Cloudflare 上给 Worker 自定义域名，并开启 Access 保护（可选）
- 登录会话有效期 7 天，所有登录与操作均记录日志
