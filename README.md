# mevps · VPS 服务器管理面板

基于 **Cloudflare Workers + D1** 的轻量 VPS 服务器管理项目，前后端全部在一个 Worker 文件里，
可以通过 **Cloudflare 控制台（纯浏览器操作）** 完成部署，无需命令行、无需构建。

后台面板样式参考 [soybean-admin](https://github.com/soybeanjs/soybean-admin)：清爽绿色主题、侧边栏菜单 + 卡片式内容区。

## 功能

| 左侧菜单 | 功能 |
| --- | --- |
| 数据仪表盘 | VPS 总数 / 运行中 / 即将到期 / 月度支出统计卡片，服务商分布、状态分布图，最近操作记录 |
| VPS 管理 | VPS 资产增删改查：名称、IP、服务商、机房、配置（CPU/内存/磁盘）、系统、状态、到期时间、价格、备注；支持搜索、状态筛选、分页 |
| 操作日志 | 全部操作自动留痕（登录/登出/新增/编辑/删除/改设置/改密码），含操作人、对象、详情、IP，支持筛选与分页 |
| 系统设置 | 系统变量（站点名称、管理员邮箱、到期提醒天数）统一存储于 D1；修改登录密码 |

## 默认登录

- 用户名：`admin`
- 密码：`123456`

> ⚠️ 首次登录后请立即在「系统设置」中修改默认密码。

## 浏览器部署步骤（Cloudflare 控制台）

### 1. 创建 D1 数据库

1. 打开 [Cloudflare Dashboard](https://dash.cloudflare.com/) → **Workers & Pages** → **D1**（或「存储和数据库 → D1」）
2. 点击 **Create database**，数据库名称固定填写：**`mevps`**
3. 创建完成即可，无需手动建表（首次访问自动建表并写入默认数据）

### 2. 创建 Worker 并粘贴代码

1. **Workers & Pages** → **Create** → **Create Worker**，名字填 `mevps`
2. 点击 **Edit code**（编辑代码），把仓库中 `src/worker.js` 的**全部内容**粘贴进去（替换默认代码）
3. 点击 **Deploy** 保存

### 3. 绑定 D1（变量名固定为 db）

**方式一：控制台手动粘贴部署时**

1. 回到 Worker 的 **Settings** → **Bindings** → **Add binding**
2. 选择 **D1 database**：
   - **Variable name（变量名）** 固定填：**`db`**
   - **D1 database** 选择：**`mevps`**
3. 保存并再次 **Deploy**

**方式二：GitHub 连接自动部署（Workers Builds）**

1. **Workers & Pages** → **Create** → **Import a repository**，选择本仓库
2. Build command 填：`npx wrangler deploy`（其余留空）
3. Deploy 即可 —— `wrangler.toml` 里只声明了 `binding = "db"` 和 `database_name = "mevps"`，
   **不填 database_id，部署时会自动创建（或复用已存在的）`mevps` 数据库并完成绑定**，无需任何手动配置

### 4. 访问

打开 Worker 的 `*.workers.dev` 域名，首次访问会自动创建数据表并初始化，直接用 `admin / 123456` 登录。

## 项目结构

```
mevps/
├── src/
│   └── worker.js     # 全部代码（API + 前端页面），粘贴进 CF Worker 即可运行
├── wrangler.toml     # 可选：使用 Wrangler CLI 部署时的配置
└── README.md
```

## 数据说明

- **所有系统变量都存在 D1 里**（`settings` 表，key-value 形式），修改后立即生效
- 数据表首次访问自动创建：`users`（账号）、`sessions`（会话）、`vps`（机器资产）、`logs`（操作日志）、`settings`（系统变量）
- 默认系统变量：`site_name=mevps`、`admin_email=mevps`、`expire_warn_days=7`

## 可选：Wrangler CLI 部署

```bash
npm i -g wrangler
wrangler login
wrangler deploy               # 首次部署自动创建 mevps 数据库并绑定
```

## 安全提示

- 上线后第一时间修改默认密码 `123456`
- 建议在 Cloudflare 上给 Worker 自定义域名，并开启 Access 保护（可选）
- 登录会话有效期 7 天，所有登录与操作均记录日志
