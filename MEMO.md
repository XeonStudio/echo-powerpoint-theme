# PowerPoint Style Echo Plugin 开发与版本规范备忘录

## 1. 单向基准演进原则 (One-Way Baseline Evolution)
- **唯一母版基准**：此后所有代码修改、功能新增与缺陷修复，**默认且强制必须以现有最新版本（最新干净基准副本）为唯一母版**。
- **严禁交叉修改历史淘汰版本**：严禁回退修改数个版本更迭前的废弃代码、中间备份或过时分支文件，杜绝历史旧代码反向污染最新工程。
- **改动审计记录**：每一次版本迭代必须严格记录改动原因、变更代码范围并在 CHANGELOG.md 与 GitHub Release 中清晰阐述。

## 2. 语义化版本规范 (Semantic Versioning)
- **版本号递增**：每次发布严格遵循 Semantic Versioning 规范（例如 v1.0.0 -> v1.1.0 -> v1.2.0 -> v1.3.0）。
- **清单同步**：所有配置文件中的版本号必须保持完全一致：
  - `package.json`
  - `echo.workshop.project.json`
  - `content/echo.workshop.json`
  - `content/theme.json`
  - `CHANGELOG.md`
  - `MEMO.md`

## 3. 零 Emoji 强制门禁 (Zero-Emoji Policy)
- 工程内所有源文件（包括 `.js`, `.mjs`, `.cjs`, `.json`, `.css`, `.html`, `.md`, `.d.ts`）严禁包含任何 Emoji 图标字符。
- 界面图标统一采用 Office 365 / Microsoft Fluent Design 矢量 SVG。
- 每次提交前必须强制执行 `node scan-emoji.cjs` 检查，确保检测结果为 0 Emoji。

## 4. 工坊 SDK 质量门禁与多端同步 (SDK Quality Gate & Distribution)
- **SDK 检测标准**：每次构建必须通过 `echo-workshop-sdk.mjs check .` 全部 12 项检查（12 pass, 0 warning, 0 blocker）。
- **多端同步清单**：每次最新构建通过 `sync-packages.cjs` 自动同步至以下目标环境：
  1. Scratch 研发工程仓库 (`echo-powerpoint-theme`)
  2. 桌面工程母版 (`Desktop/echo.powerpoint-theme`)
  3. 桌面独立分发包 (`Desktop/echo.powerpoint-theme-package`)
  4. 桌面 ZIP 发布归档 (`Desktop/echo.powerpoint-theme-x.x.x.zip`)
  5. ECHO Steam 工坊安装运行时 (`AppData/Roaming/ECHO Steam/workshop/installed/...`)
  6. ECHO 宿主 Mods 扩展目录 (`ECHO/Mods/installed/echo.powerpoint-theme`)

## 5. GitHub 远端同步与 Release 整合规范 (GitHub Releases Integration)
- **远端仓库基准**：`https://github.com/XeonStudio/echo-powerpoint-theme.git`
- **提交与打标**：最新构建经过自动化验证通过后，默认提交 Git Commit，生成对应版本的 Git Tag（如 `v1.3.0`），推送到 GitHub 远端。
- **Release 整合发布**：在 GitHub 上整合详细更新日志，发布正式 Release 供用户下载与追溯。
