# Codex / Trae Solo 接手说明

你现在接手的是一个中文公文写作 skill。迁移目标是另一台 Mac，上面可能运行 Codex 或 Trae Solo。对外迁移包本身就是 `gongwen-drafting` skill；不要依赖原电脑的 `~/.codex/memories`，也不要凭通用记忆补规则。默认直接读取 skill 内的 `SKILL.md` 和 `references/` 执行；只有用户另行提供原始工作库时，才把外部资料库作为补充材料。

## 1. 当前项目是什么

本 skill 用于整理、检索、改写和生成中文公文材料，重点是文物手续相关公文。

常见文种包括：

- 文物保护许可申请 / 文物核查请示
- 勘探验收请示
- 发掘请示
- 开工请示
- 勘探计划备案请示
- 勘探报告备案请示
- 考古勘探工作支持请示
- 文物保护安全责任书
- 文物调查意见

文物调查报告、市级调查报告、自治区调查报告另由独立 `gongwen-survey-report` skill 处理。

## 2. 先读这些文件

如果你拿到的是 `gongwen-drafting.skill`，它应只包含一个自包含 `gongwen-drafting/` skill 文件夹。先安装或读取这个 skill：`references/` 里有规则、索引、文种手册和模板内化，`assets/` 里有政策规范、Word 模板和历史参考件，`scripts/` 里有本地工具链。用户另行提供的新项目材料应作为当前任务输入处理。

1. `00_开始使用_请先读.md`
   - 人和 AI 的最短入口。
   - 先确认资料库起读顺序、`T/P/K` 映射、常用命令、修改日志纪律和禁止事项。

2. `AGENTS.md`
   - 最高优先级项目记忆入口。
   - 包含写作风格、文种判断、资料需求、面积禁忌、固定文种规则、Word 版式规则、资料库规范化规则和跨设备迁移规则。

3. `0.资料索引/00-资料库规范化总览.md`
   - 当前资料库按《智能工具结构文档》形成的 `T/P/K` 映射说明。
   - 第一批规范化只建立入口、映射和日志，不移动原始政策、模板和参考公文目录。
   - 后续真实移动目录前，必须先检查 Obsidian 链接、脚本相对路径和迁移包清单。

4. `agent_handoff/PROJECT_MEMORY_EXPORT.md`
   - 从过往协作和本机 memories 显化出来的接手摘要。
   - 用于理解用户偏好、质量底线、常见错误和迁移边界。

5. `agent_handoff/skills/gongwen-drafting/SKILL.md`
   - 公文撰写总流程 skill。
   - 负责启动判断、文种路由、资料需求、写作底线、Word 交付链路、OCR 提示和资料库维护边界。
   - 如果当前 AI 支持本地 skill，可安装或引用整个 `agent_handoff/skills/gongwen-drafting/`；如果不支持，直接阅读该 `SKILL.md` 和其 `references/`。

6. `agent_handoff/skills/gongwen-hecha-fixed-format/SKILL.md`
   - 文物保护许可申请 / 文物核查请示的固定版式专项 skill。
   - 其中附件说明和落款参数也是跨文种固定版式规则，勘探验收、开工请示等 Word 公文同样要复用。
   - 如果当前 AI 支持本地 skill，可安装或引用整个 `agent_handoff/skills/gongwen-hecha-fixed-format/`。
   - 如果 Trae Solo 不支持 Codex skill 安装，直接阅读该 `SKILL.md` 并按流程执行。

7. `agent/README.md`
   - 本地公文撰写 Agent 的入口和快速命令。

8. `agent/使用说明.md`
   - Mac 运行环境、依赖、输入字段、建库、检索、生成和复核流程。

9. `0.资料索引/00-资料库总览.md`
   - Obsidian 资料库索引总览。
   - 需要找模板、参考件、文种索引时再继续读取其中链接。
   - 写作判断优先读取 `0.资料索引/00-AI工作台.md` 和 `0.资料索引/文种手册/00-文种手册总览.md`，再进入具体文种手册、模板内化文件和文件卡片。
   - 写文物调查报告、市级调查报告、自治区调查报告时，转用独立 `gongwen-survey-report` skill。
   - 写协助勘探、考古勘探工作支持类请示时，先读 `0.资料索引/文种手册/考古勘探工作支持请示-作战手册.md`，不要并入普通“申请开展考古勘探工作请示”。
   - 生成或大幅修改公文后，可按 `0.资料索引/生成记录/成稿后复盘模板.md` 写复盘卡；尚未固化的经验先进入 `0.资料索引/规则候选/`；资料库体检记录放入 `0.资料索引/健康检查/`。
   - 每日首次维护资料库或接手项目时，按 `0.资料索引/健康检查/参考公文每日更新检查清单.md` 检查 `3.公文参考/` 新增公文并整理归档，再把可复用规律回流到索引、手册、模板和版式检查中。

## 3. 最高优先级规则

- `AGENTS.md` 是长期规则主入口；不要另建或依赖 `公文写作规则.md`。
- 公文写作交付链路固定为：先按文种手册、地区模板和资料库固定版式脚本生成桌面 `.docx` 成稿；不再调用 `docformat-gui`、`agent/format_existing_docx.py` 或外部一键格式化引擎覆盖资料库既有版式。除非用户明确索要，不交付 Markdown 草稿或中间过程文件。
- 生成或大幅重建文物保护许可申请 / 核查请示时，优先遵守 `agent_handoff/skills/gongwen-hecha-fixed-format/SKILL.md` 和 `agent/` 固定版式脚本；最终交付前不得再走 `docformat-gui` 二次格式化。
- 附件说明和落款属于资料库专项生成工具中的固定版式区域；使用 `agent/gongwen_format.py`、`agent/make_hecha_docx.py` 等脚本生成或大幅重建公文时，应复用其中附件悬挂缩进、制表位和落款对齐参数。既有成稿需要调格式时，按当前文种和问题区域外科式处理，不得使用通用一键格式化工具重排整篇。
- 请示类正文必须先写背景、事实依据或工作进展，再写请示事项；不得开头直接写“现申请”“特此请示”等结论句。
- 可以写“项目面积”或“项目用地面积”；禁止写“调查面积”和“勘探面积”。
- 报送对象不得提前写死；能从材料识别就使用，识别不了必须问用户。
- 默认优先拟建设单位、项目单位、企业发出的文件；不主动拟文物部门内部转报请示。
- 模板型公文默认只替换应替换内容，不擅自扩写、加附件、加校验页或重排整篇。
- 新增参考公文应按现有文种、地区和文件角色目录归档；依据参考公文校正通用/地区模板和版式时，必须区分稳定规律与单份个案。
- 文物调查报告不是本 skill 的内置文种；遇到该类报告，转用独立 `gongwen-survey-report` skill。
- 正文段落结构服从公文规范，不为压缩页数合并应独立成段的“此函”“特此请示”“妥否，请批示”等收束语。
- 没有联系人和电话时不写，不留空占位；不得输出 `XX`、`20XX`、`待补充`。

## 4. 文件角色

- 根目录最短入口：`00_开始使用_请先读.md`
- 版本日志：`99_版本日志.md`
- 项目记忆：`AGENTS.md`
- 资料库规范化总览：`0.资料索引/00-资料库规范化总览.md`
- 历史记忆显化摘要：`agent_handoff/PROJECT_MEMORY_EXPORT.md`
- AI 可执行 skill：`agent_handoff/skills/gongwen-drafting/SKILL.md`、`agent_handoff/skills/gongwen-hecha-fixed-format/SKILL.md`
- 本地工具代码：`agent/`
- 固定版式工具：`agent/gongwen_format.py`、`agent/make_hecha_docx.py`、`agent/check_hecha_docx.py`、`agent/check_template_route.py`
- 工具说明：`agent/README.md`、`agent/使用说明.md`、`agent/固定版式说明.md`
- 资料索引与 AI 编译 Wiki：`0.资料索引/`、`.obsidian/`
- AI 写作入口：`0.资料索引/00-AI工作台.md`
- 文种作战手册：`0.资料索引/文种手册/`
- Word 模板内化：`0.资料索引/模板内化/`
- 成稿复盘：`0.资料索引/生成记录/`
- 规则候选：`0.资料索引/规则候选/`
- 健康检查：`0.资料索引/健康检查/`
- 规范化日志：`0.资料索引/健康检查/2026-06-15-资料库规范化日志.md`
- 资料本体：`1.政策法规与规范/`、`2.公文模板/`、`3.公文参考/`
- 运行缓存：`.gongwen_agent/`、`.venv/`、日志、pid 和本机配置文件，迁移包通常不带，到新 Mac 后重新生成或重新配置

## 5. 接手后验证命令

在 skill 目录运行：

```bash
python3 -m unittest agent/test_gongwen_agent.py
```

检查 Word 生成和固定版式工具链：

```bash
python3 -m py_compile agent/gongwen_agent.py agent/gongwen_format.py agent/make_hecha_docx.py agent/check_hecha_docx.py agent/check_template_route.py
python3 agent/check_template_route.py --region 库伦旗 --doc-type 文物核查请示
```

轻量重建索引：

```bash
python3 agent/gongwen_agent.py index --ocr --ocr-max-pages 3 --coverage-per-type 5
```

启动网页表单：

```bash
python3 agent/gongwen_agent.py web --host 127.0.0.1 --port 8765
```

核查请示固定版式生成与校验：

```bash
python3 agent/make_hecha_docx.py agent/sample_hecha_project.json --out /tmp/核查请示.docx --audit
python3 agent/check_hecha_docx.py /tmp/核查请示.docx
```

如果缺依赖，先读 `agent/使用说明.md`，在 Mac 上检查 `python3`、`python-docx`、`pdftotext`、`pdftoppm`、`tesseract` 和中文 OCR 语言包。

## 6. 用户发起任务时怎么判断

- 用户说“帮我写”“帮我拟”“帮我生成”“整理成请示”“出一版公文”等，并提供项目资料时，视为启动公文撰写流程。
- 先判断文种、办理阶段、发文主体和受文关系；能判断就直接按对应规则起草。
- 只有在无法判断文种、办理阶段、发文主体或收文关系，且继续生成会导致方向明显错误时，才集中追问关键问题。
- 能从资料中确认的信息直接使用；不能确认且影响事实、主送机关、附件一致性的，向用户说明缺口。
- 输出以桌面最终格式化 `.docx` 为主；默认保存到桌面，除非用户指定路径。

## 7. 明确禁止

- 不依赖原电脑 `~/.codex/memories`。
- 不凭通用公文知识覆盖本项目规则。
- 不把一次性修改要求自动写成长期规则；只有用户明确要求“固定”“记住”“写入项目记忆”“以后都按这个来”时才固化。
- 不把成稿复盘、规则候选或健康检查记录当作硬规则；硬规则仍以 `AGENTS.md`、正式 skill 和脚本为准。
- 不把 `.gongwen_agent/` 当作必须迁移内容；新 Mac 重新建索引。
- 不再迁移 `agent/format_existing_docx.py` 和 `agent/vendor/docformat_engine/`；迁移包必须保留资料库固定版式脚本和地区模板路由门禁。
- 不把复杂业务规则塞给人读；复杂规则由 `AGENTS.md`、`PROJECT_MEMORY_EXPORT.md`、skill 和脚本承载。

## 8. 给人的跨设备说明

人只需要做四件事：

1. 把 `gongwen-drafting.skill` 当作 skill 包导入或解压，里面应只有一个自包含 `gongwen-drafting/` skill 文件夹。
2. 打开 Codex 或 Trae Solo，让 AI 先读取 `gongwen-drafting/SKILL.md`。
3. 模板、历史参考件和本地格式化脚本默认已经在 skill 的 `assets/`、`references/` 和 `scripts/` 中；只有用户另行提供外部材料时，才读取外部路径。

不再需要“完整运行包”；默认新电脑已有 Codex、Trae Solo 等 AI 工具。给 AI 助手的迁移包不是整库备份，而是 skill。
