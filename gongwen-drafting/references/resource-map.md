# 内置资源地图

本 skill 是自包含迁移包。除用户另行提供的新项目材料外，优先使用本 skill 内资源。

## references/

- `routing.md`：文种路由、资料需求和常见文种判断。
- `drafting-rules.md`：写作底线、禁用词、长期记忆边界和维护边界。
- `word-output.md`：Word 输出、格式化、OCR 和验证命令。
- `policies-and-standards/`：政策规范、格式标准等阅读参考材料。
- `reference-documents/`：历史参考公文和案例材料，用于学习结构、措辞、附件和地区差异。
- `fixed-format/`：固定版式范本的校验说明、检查记录和解释性材料。
- `project-rules/`：从资料库根目录同步的长期入口与版本记录，包括 `AGENTS.md`、`00_开始使用_请先读.md`、`99_版本日志.md`。
- `project-handoff/`：AI 接手摘要、项目记忆导出和依赖说明。
- `knowledge-index/`：原 `0.资料索引/`，包括文种手册、模板内化、文件卡片、生成记录、规则候选和健康检查。

## assets/

- `templates/`：原 `2.公文模板/` 中可复用的 Word 模板和模板图片等素材。
- `fixed-format/`：固定版式 Word 范本等可复制、填充或作为版式依据的源文件。
- `fonts/`：公文标题和正文相关字体包。

`references/` 是需要阅读和理解的材料层；`assets/` 是可复制、填充、转换或作为输出依据的素材层。需要阅读时先优先使用 `references/knowledge-index/` 中的文件卡片、文种手册和模板内化；只有需要核对原模板、固定版式 Word 范本或字体素材时，再打开 `assets/` 中的原始文件。

文物调查报告、市级调查报告、自治区调查报告的模板、案例、手册和模板内化文件已拆分到独立 `gongwen-survey-report` skill；本 skill 不再内置 `10.文物调查报告` 参考件和调查报告模板。

## scripts/

- `check_gongwen_workspace.py`：检查 skill 包是否具备关键入口。
- `install_environment.sh`：macOS 一键安装基础运行环境，包括 Python 依赖和 Homebrew 工具包。
- `check_environment.py`：检查新电脑运行环境是否具备 Python 依赖、PDF 工具、OCR 工具和内置格式化脚本。
- `gongwen-agent/`：原 `agent/` 工具链核心副本，包含生成、格式化、固定版式、模板内化和模板路由门禁脚本；不包含外部消息桥接、本机配置或运行日志。

默认使用 skill 内置的 `scripts/gongwen-agent/`。只有用户另行指定外部项目或原始工作库时，才切换到外部路径。

## 不包含

- `.gongwen_agent/` 索引缓存。
- `.venv/` 虚拟环境。
- `.DS_Store`、`.WeDrive`、`__pycache__`。
- 本机配置、运行日志、pid 文件、桌面成稿和旧 zip 包。
- 文物调查报告专属模板、案例和手册；这些由 `gongwen-survey-report` 单独迁移。
- `assets/reference-documents/` 和 `assets/policies-and-standards/`；参考案例和政策规范必须放在 `references/` 下。
