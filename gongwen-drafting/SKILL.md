---
name: gongwen-drafting
description: Draft, revise, route, validate, and deliver Chinese official request and procedure Word documents from a self-contained skill package, especially 文物保护许可申请/文物核查请示、勘探验收请示、开工请示、勘探计划备案请示、勘探报告备案请示、申请开展考古勘探工作请示、考古勘探工作支持请示. Use when the user asks 帮我写/拟/生成/整理成请示/修改公文/调格式/出一版 Word/公文写作. Do not use for 文物调查报告、市级调查报告、自治区调查报告; route those to gongwen-survey-report.
---

# 公文撰写总流程 Skill

## 版本定位

- 当前版本：v1.0（2026-06-26）
- 版本含义：稳定版；请示、函、许可、备案、验收、开工等公文流程已完成拆分、模板路由、红头固定版式、Word 桌面交付和自检。

## 核心原则

本 skill 是自包含业务 skill，也就是默认运行入口和迁移单元：`references/` 内置规则、索引、文种知识、政策规范和历史参考件，`assets/` 内置可复用 Word 模板、固定版式范本、字体和模板图片，`scripts/` 内置本地公文工具链。使用时直接调用本 skill，并优先从本 skill 内取规则、模板、参考和脚本；只有用户另行指定外部材料时，才到外部路径读取。文物调查报告已拆分为独立 `gongwen-survey-report` skill，本 skill 只保留请示、函、许可、备案、验收、开工等公文流程。

本项目副本位于 `/Users/drevan01/Desktop/待出公文/skills/gongwen-drafting/`。在 `/Users/drevan01/Desktop/待出公文` 中使用时，先读取项目根目录 `AGENTS.md` 和 `references/project-local.md`，再按本 skill 的通用流程执行。

执行时把本 skill 当作公文写作工作库使用：先读本 skill 内规则，再判断文种和办理阶段，最后交付桌面上的格式化 Word。不要凭通用公文知识覆盖 `references/project-rules/`、文种手册、模板内化文件或脚本规则。

协作时先判断信息状态：

- 开放区：任务清楚时直接处理。
- 隐藏区：项目背景、偏好或材料不足但会影响方向时，集中问最关键问题。
- 盲区：用户反复觉得不对时，主动指出可能忽略的文种、流程或模板路径。
- 未知区：任务还在探索时，给出多个可比较方向，不急着锁死结论。

## 先读顺序

默认先读：

1. `SKILL.md`
2. 项目副本中读取 `references/project-local.md`
3. `references/routing.md`
4. `references/drafting-rules.md`
5. `references/word-output.md`
6. `references/resource-map.md`

根据任务再读本 skill 的参考文件：

- `/Users/drevan01/Desktop/待出公文` 项目内规则、规则沉淀和本地材料边界：`references/project-local.md`
- 文种判断、模板路由、资料需求：`references/routing.md`
- 写作底线、禁用词、长期记忆边界：`references/drafting-rules.md`
- Word 生成、格式化、校验和 OCR：`references/word-output.md`
- 内置资源位置和使用边界：`references/resource-map.md`
- 项目长期规则副本：`references/project-rules/`
- AI 接手说明副本：`references/project-handoff/`
- 资料索引、文种手册、模板内化和文件卡片：`references/knowledge-index/`

## 工作流

1. 识别用户意图：生成新公文、修改既有公文、调 Word 格式、整理规则、维护资料库或打包迁移。
2. 判断文种、办理阶段、发文主体和受文关系。能从材料判断就继续；判断会明显错向时再问用户。若识别为文物调查报告、市级调查报告或自治区调查报告，停止本 skill 流程并转用 `gongwen-survey-report`。
3. 按“地区 + 文种”优先检索模板索引、文种手册、模板内化文件和参考件，不要先套通用模板。
4. 请示类正文先写背景、事实依据或工作进展，再写请示事项。
5. 生成或修改时保持事实口径：项目名称、单位名称、面积、日期、数量、附件名称按材料原文或用户确认处理；可以写“项目面积”或“项目用地面积”，禁止写“调查面积”和“勘探面积”。
6. 输出时默认生成一份桌面格式化 `.docx`。最终调整版式时，必须以 `assets/fixed-format/固定版式——文物保护许可核查请示格式范本.docx` 作为固定版式 Word 样张，逐项核对红头、标题、正文、附件、落款的字体字号、段落缩进、附件悬挂缩进、制表位和落款对齐；带红头且暂无正式文号或签发人时，必须在红头信息行保留 `【文号】`、`【签发人】` 文字占位；除非地区模板有明确硬性差异，不得用通用公文记忆或手工空格替代样张版式。
7. 大幅生成或修改后，按项目规则补成稿复盘；资料库结构、入口、规则、脚本或索引说明变化必须写健康检查日志。
8. 用户明确要求完善、固定或更新 skill 时，优先修改本项目副本；确认稳定后再同步到全局 skill。

## 专项路由

遇到文物保护许可申请、文物核查请示、附件悬挂缩进、落款对齐或固定版式问题时，加载并执行：

优先读取本 skill 内的固定版式规则、模板内化和工具链；如当前环境另有 `gongwen-hecha-fixed-format` 专项 skill，也可加载它作为核查请示固定版式补充。

不要让本总流程 skill 覆盖专项 skill 的版式细节。总流程负责入口、路由、交付链路和资料库纪律；专项 skill 和本 skill 内置固定版式工具负责红头、标题、正文、附件、落款的硬校验。

## 工具入口

常用命令：

```bash
python3 scripts/gongwen-agent/gongwen_agent.py index --ocr --ocr-max-pages 3 --coverage-per-type 5
python3 scripts/gongwen-agent/gongwen_agent.py draft project.json
python3 scripts/gongwen-agent/gongwen_agent.py web --host 127.0.0.1 --port 8765
python3 scripts/check_gongwen_workspace.py
```

快速检查 skill 入口和关键规则：

```bash
python3 scripts/check_gongwen_workspace.py
```

生成前检查地区模板路由：

```bash
python3 scripts/gongwen-agent/check_template_route.py --region 库伦旗 --doc-type 文物核查请示
python3 scripts/gongwen-agent/check_template_route.py --region 库伦旗 --doc-type 文物核查请示 --planned-route generic
```

新 Mac 一键准备运行环境：

```bash
./scripts/install_environment.sh
python3 scripts/check_environment.py
```

OCR 需求默认使用本地 Paddle OCR，路径为 `/Users/drevan01/Desktop/OCR`；如当前项目脚本只支持其他 OCR 入口，先说明差异再执行。
