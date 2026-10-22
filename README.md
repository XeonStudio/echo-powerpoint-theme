# Microsoft PowerPoint · ECHO 创意工坊外观插件

[![ECHO Workshop](https://img.shields.io/badge/ECHO%20Workshop-Theme-c43e1c.svg)](https://store.steampowered.com/app/5105090/ECHO/)
[![Version](https://img.shields.io/badge/version-1.2.0-blue.svg)](CHANGELOG.md)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

经典 Microsoft PowerPoint 演示文稿主题界面皮肤（针对 ECHO 音乐播放器的创意工坊深度定制主题）。

---

## 核心特性

1. **经典 PowerPoint 顶部与功能区（Fluent Ribbon）**
   - **标牌栏**：呈现演示文稿文件名（`[歌曲名] - Microsoft PowerPoint`）、快速访问工具栏、播放列表全局即时搜索与系统控制。
   - **二级功能选项卡**：集成「开始」、「设计」、「编辑」、「多轨」、「剪辑」、「效果」、「缩放」、「播放」、「窗口」、「帮助」等 10 个选项卡面板。
   - **实时字体排版控制**：在功能区可直接实时调节歌词的字体家族、字号大小、加粗、倾斜、下划线、文字阴影与高亮着色。
   - **统一高精度排版**：所有 45 个功能区大按钮图标与文字基准线严格水平对齐，零错位参差。

2. **多模式幻灯片版式（Multi-Layout）**
   - **经典黑胶唱机模式 (Classic Vinyl)**：高精 3D 质感黑胶唱片匀速旋转，实时解析并展示音频硬件与流解码属性看板（FLAC/MP3/WAV 采样率、位深度、比特率与电平）。
   - **网易云双列模式 (NetEase Style)**：左侧大封面海报、右侧自适应双列动态流线歌词。
   - **Apple Music 流体模式 (Apple Music Fluid)**：根据专辑主色自适应动态流体背景，配合居中大字歌词排版。

3. **预设记忆与自动持久化恢复系统 (Preset Memory & Persistence)**
   - 突破工坊 iframe 沙箱环境隔离，通过宿主 Electron Mod Bridge 持久化通道无缝跨会话保存。
   - 自动记忆明暗主题、当前功能区选项卡、幻灯片排版、缩放比例、背景遮罩/模糊度、转速、双语翻译开关及字体排版，再次进入时原样还原。

4. **左侧幻灯片播放队列**
   - 以 PowerPoint 经典 16:9 幻灯片编号缩略图形式展示当前播放队列中的所有曲目。
   - 当前播放曲目呈现 Office 经典橙红选框高亮，支持一键点击任意幻灯片即时切歌。

5. **底部演讲者备注控制台**
   - 集成 PPT 备注提示栏、播放/暂停、快进/快退、曲目切换、音量滑块、幻灯片总数与缩放控制条（支持 50% ~ 200% 任意无级缩放）。

---

## 版本历史概览

| 版本号 | 发布日期 | 重点更新内容 |
| :--- | :--- | :--- |
| **v1.2.0** | 2026-10-22 | 修复功能区大按钮排版错落问题（去除括号快捷键说明）；新增完整预设记忆与自动恢复系统；运行鲁棒性加固。 |
| **v1.1.0** | 2026-10-21 | 新增黑胶/网易云/Apple Music 三大多元排版布局；新增背景高斯模糊与遮罩浓度调节；支持 F5 全屏放映。 |
| **v1.0.0** | 2026-10-20 | 初始版本发布，经典 Office PowerPoint Fluent Ribbon 界面复刻、幻灯片缩略图播放队列与实时歌词排版。 |

详细改动说明见 [CHANGELOG.md](CHANGELOG.md)。

---

## 项目结构

```text
echo-powerpoint-theme/
├── .echo-sdk/                 # ECHO 官方工坊 SDK 与契约定义
├── content/                   # 创意工坊分发内容目录
│   ├── ui/                    # 运行态前端资源
│   │   ├── app.css            # Fluent UI 样式表
│   │   ├── app.js             # 主题核心逻辑与预设持久化桥接
│   │   └── index.html         # PowerPoint 经典界面骨架
│   ├── echo.mod.json          # Mod 元数据
│   ├── echo.workshop.json     # 工坊清单配置 (SHA256 完整性校验)
│   ├── mod.js                 # 宿主 Electron 桥接服务
│   └── theme.json             # 调色板与外观契约配置
├── echo.mod.json              # 顶层 Mod 声明
├── echo.workshop.project.json # 工坊工程配置
├── mod.js                     # 顶层宿主脚本
├── package.json               # 项目包配置
├── preview.png                # 工坊封面预览图 (256x256)
├── CHANGELOG.md               # 详细版本改动说明
├── README.md                  # 项目使用与技术说明文档
├── scan-emoji.cjs             # 表情符静态审查脚本
└── sync-packages.cjs          # 多端自动同步与发布打包工具
```

---

## 安装与开发调试

### 环境要求
- ECHO 音乐播放器版本 >= 26.9.13
- Node.js >= 18.0.0

### 本地构建与验证
```bash
# 验证代码库是否包含非法 emoji 字符
node scan-emoji.cjs

# 运行 ECHO Workshop SDK 规范质量审查
node ./.echo-sdk/bin/echo-workshop-sdk.mjs check .

# 自动同步更新到本地 ECHO 工坊运行库、桌面免安装包与 Release ZIP
node sync-packages.cjs
```

---

## 开源协议

本项目采用 [MIT License](LICENSE) 开源。
