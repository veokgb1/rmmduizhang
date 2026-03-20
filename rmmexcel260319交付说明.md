# 财务对账管理系统 - 代理部署方案 (二代架构)

## 架构说明
本系统采用前后端分离的无服务器 (Serverless) 架构：
- **前端页面**：纯 HTML/JS，使用 Base64 编码内嵌于 Cloudflare Worker 中，实现极速加载。
- **代理层**：Cloudflare Worker，负责域名绑定、跨域处理 (CORS) 以及请求转发。
- **后端数据**：Google Apps Script (GAS) + Google Sheets，处理核心业务逻辑与数据存储。

## 部署与操作明细

### 1. 前端 HTML 压缩 (Base64)
为了避免 HTML 代码中的反引号与 Worker 脚本冲突，需要将 `index.html` 转换为 Base64 编码。
**操作方法：**
1. 在浏览器打开空白页，按 F12 进入控制台 (Console)。
2. 粘贴专用的转换脚本（弹出框输入 HTML 源码，自动生成 Base64 并复制）。
3. 将生成的 Base64 字符串填入 `worker.js` 中的 `HTML_B64` 变量内。

### 2. Cloudflare 安全配置 (Secrets)
为保证后端入口安全，`GAS_ID` 严禁硬编码在代码中，必须使用 Cloudflare 的环境变量进行保护。
**操作方法：**
1. 登录 Cloudflare -> 选择对应的 Worker -> 设置 (Settings) -> 变量和机密 (Variables and Secrets)。
2. 添加机密 (Secrets)：
   - **名称**: `GAS_ID`
   - **值**: (填入 Google Apps Script 部署生成的 Deployment ID)
3. 代码中通过 `env.GAS_ID` 进行调用。

### 3. Google Apps Script 配置说明
- 核心代码存放在后端的 `.gs` 文件中，不在本公开仓库中展示。
- 后端脚本必须发布为 Web App，且访问权限需设置为 "Anyone" (允许 Cloudflare 转发的请求进入)。
- 真正的敏感机密（如 Gemini AI API Key）存放在 Google 后端的 Script Properties 中，前端绝对隔离，确保系统安全。

财务对账系统 - 核心部署备忘录 (2026版)
⚠️ 特别提醒： 本仓库代码已进行“脱敏处理”。GitHub 上的 worker.js 和 .gs 文件中不包含真实的 API Key 和 ID。真实的“钥匙”锁在云端后台。

1. 部署基本信息
部署日期：2026年3月19日

当前状态：二代安全架构已全面跑通

核心访问域名：https://rmm.veokgb1.top

2. 部署具体位置（钥匙在哪里？）
当以后需要更换 Key 或重新部署时，请去以下两个地方找“钥匙”：

位置 A：Cloudflare Worker 后台 (中转站)

路径：Worker -> 设置 (Settings) -> 变量和机密 (Variables and Secrets)

存了什么：GAS_ID (Google 脚本的部署 ID)

作用：让 Worker 知道去哪里找你的 Google 后厨。

位置 B：Google Apps Script 后台 (数据后厨)

路径：脚本编辑器 -> 项目设置 (齿轮图标) -> 脚本属性 (Script Properties)

存了什么：GEMINI_API_KEY 和 TOKEN_SECRET

作用：处理 AI 对账逻辑和登录加密。