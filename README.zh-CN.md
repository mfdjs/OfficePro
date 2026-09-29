# OfficePro

**基于 Electron 的一体化跨平台办公套件** —— 文字处理、电子表格、演示文稿、PDF 工具、文档扫描和轻量数据库，全部集成在一个桌面应用中。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Electron](https://img.shields.io/badge/Electron-22-47848F.svg)](https://www.electronjs.org/)
[![Platform: Windows](https://img.shields.io/badge/Platform-Windows-0078D4.svg)](#下载)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](CONTRIBUTING.md)

[English](README.md) | 简体中文

> **本项目正在寻求社区贡献者。** OfficePro 已经是一个可正常使用的应用，但它是个人开发者的作品，界面集中在一个较大的单文件里，测试也比较有限。如果你擅长 Electron、前端架构、文件格式解析或测试，非常欢迎你的帮助。请见 [求助方向](#求助方向)，并直接发起 PR 或 Issue。

---

## 功能特性

- **文字**：富文本编辑、标题/样式、表格、图片、修订、批注、脚注/尾注、引文与目录、邮件合并，以及本地智能写作辅助。
- **表格**：公式与函数库、图表、数据验证、条件格式、排序/筛选、冻结窗格、分组、名称管理器、分类汇总、工作表保护。
- **演示**：幻灯片、形状、文本框、艺术字、动画与切换、主题与变体、母版、演示者视图，以及完整的放映。
- **PDF 工具**：查看、合并/拆分、压缩、加密/解密、提取页面、密文遮盖、与其他格式互转、OCR。
- **扫描王**：相机/导入扫描、批量处理、OCR、滤镜、旋转/翻转、亮度/对比度/锐化、老照片修复。
- **数据库**：轻量的表与查询。
- **多语言界面**：可在运行时切换 10 种语言。

## 支持的文件格式

OfficePro 会尽力打开你交给它的任何文档，包括：

- **文档**：doc、docx、docm、dot、dotx、odt、rtf、wps、pages、txt、md
- **表格**：xls、xlsx、xlsm、xlsb、csv、tsv、ods、numbers、dif、prn
- **演示**：ppt、pptx、pps、ppsx、pot、potx、odp、key
- **PDF 与电子书**：pdf、xps、oxps、epub、mobi、azw、azw3、fb2
- **图片**：jpg、png、bmp、gif、webp、tiff、svg、ico、avif、heic
- **文本/源代码**：json、xml、yaml、html、css、js、ts、py、java、c、cpp、cs、go、rs、php、sh、sql 等
- **压缩包与媒体**：zip、tar、gz、7z、rar、eml、msg、字体、mp3、wav、mp4、mkv 等

加密或损坏的文件会被识别并给出清晰提示，而不是显示乱码。

## 界面语言

界面支持：中文（默认）、英文、日文、韩文、西班牙文、法文、德文、俄文、葡萄牙文，以及阿拉伯文（完整的从右到左布局）。翻译文件位于 [`locales/`](locales)，由 [`i18n.js`](i18n.js) 应用。

## 开发环境快速开始

要求：**Node.js 16+** 与 **npm**。

```bash
# 1. 克隆
git clone https://github.com/<你的用户名>/OfficePro.git
cd OfficePro

# 2. 安装依赖
npm install

# 3. 开发运行
npm start
```

如需在固定端口开启调试协议：

```bash
npx electron . --no-sandbox --remote-debugging-port=9222
```

### OCR 语言数据

简体中文 OCR 训练数据未纳入仓库（约 19 MB）。Tesseract.js 可在首次使用时自动下载；如需手动放置：

```bash
mkdir -p tessdata
curl -L -o tessdata/chi_sim.traineddata.gz \
  https://tessdata.projectnaptha.com/4.0.0/chi_sim.traineddata.gz
```

## 构建打包

```bash
npm run build
```

会在 `dist/` 下生成 NSIS 安装包、绿色单文件 `.exe` 以及解压目录。

## 技术栈

- [Electron](https://www.electronjs.org/)：桌面外壳
- 原生 HTML/CSS/JavaScript：无前端框架
- [mammoth](https://github.com/mwilliamson/mammoth.js) / [word-extractor](https://github.com/ArtifexSoftware/word-extractor)：Word 解析
- [SheetJS (xlsx)](https://sheetjs.com/)：表格解析
- [JSZip](https://stuk.github.io/jszip/)：Office Open XML / zip 处理
- [pdf-parse](https://gitlab.com/autokent/pdf-parse)：PDF 文本
- [tesseract.js](https://github.com/naptha/tesseract.js)：OCR
- [docx](https://github.com/dolanmiu/docx)：文档生成

## 项目结构

```
OfficePro/
├── index.html           # 主界面（全部模块）
├── main.js              # Electron 主进程
├── preload.js           # 安全桥接
├── file-parser.js       # 多格式文件解析
├── officepro-enhance.js # 布局/快捷键/OCR 钩子
├── officepro-real.js    # 具体功能实现
├── i18n.js              # 国际化运行时
├── locales/             # 界面翻译（9 种语言）
├── tessdata/            # OCR 数据（未入库）
└── package.json
```

## 路线图

- [ ] 将较大的 `index.html` 拆分成为可维护的模块/组件
- [ ] 建立正规的单元测试与 CI
- [ ] 提升 Office Open XML 的解析/渲染保真度
- [ ] macOS 与 Linux 构建
- [ ] 大文件性能与流式处理
- [ ] 无障碍（a11y）改造
- [ ] 校对并扩充社区翻译

## 求助方向

项目尤其希望在以下方面得到帮助：

1. **架构**：把单文件界面重构为组件，分离关注点。
2. **测试与 CI**：引入单元测试和自动化构建/发布流程。
3. **文件格式保真度**：处理 Word/Excel/PPT/PDF 的更多边界情况。
4. **跨平台**：在 macOS/Linux 上验证并打包。
5. **翻译**：审校并改进自带语言包。
6. **代码现代化**：TypeScript、ESM 以及组件化的视图层。

发起 PR 前请阅读 [CONTRIBUTING.md](CONTRIBUTING.md)。

## 参与贡献

欢迎提交贡献、Bug 报告与功能建议。请 Fork 仓库、新建分支并发起 PR，详见 [CONTRIBUTING.md](CONTRIBUTING.md)。

## 开源协议

基于 [MIT 协议](LICENSE) 发布。
