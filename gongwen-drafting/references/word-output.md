# Word 输出和固定版式校验

## 交付原则

公文写作交付链路固定为：AI 按文种手册、地区模板、本 skill 内置固定版式 Word 样张和脚本直接生成桌面 `.docx` 成稿；不再调用 `docformat-gui`、旧版 `agent/format_existing_docx.py` 或外部一键格式化引擎覆盖既有版式。

除非用户明确索要，不额外交付 Markdown 草稿、未格式化 Word、中间过程文件或“问题摘要.md”。缺材料、缺正式报告、缺 KML 等风险默认只在聊天交付说明中简要提醒；只有用户明确要求“生成问题摘要”“列问题清单”“另存 Markdown”时，才额外交付对应 Markdown。

## 最终版式硬规则

最终调整版式时，必须以 `assets/fixed-format/固定版式——文物保护许可核查请示格式范本.docx` 作为固定版式 Word 样张。定稿前逐项核对红头、标题、正文、附件、落款，并重点检查字体字号、段落缩进、附件悬挂缩进、制表位和落款对齐：

- 红头：发文机关标志、文号/签发人行、红色分隔线的位置、颜色、线型和粗细；未取得正式文号或签发人时，文号/签发人行必须分别保留 `【文号】`、`【签发人】` 文字占位，不得写成裸 `文号` 或空的 `签发人：`。
- 标题：标题字体字号、居中、与红线间距、回行位置和整体形态。
- 正文：正文字体字号、行距、首行缩进、段前段后、页面版心和段落层次。
- 附件：`附件：1.` 连续起排，序号列对齐，附件名称首字对齐，长附件续行与附件名称首字对齐；使用悬挂缩进或制表位，不用连续空格。
- 落款：署名、日期、右空四字、署名相对日期居中和页面位置。

地区模板、文种手册和固定版式样张发生冲突时，优先保留地区模板的内容口径和明确硬性版式；其余通用版式区域按固定版式 Word 样张处理。不得只凭通用公文知识、自然语言描述或手工空格临时凑版式。

## 默认输出位置

生成或修改后的 `.docx` 默认保存到电脑桌面。不要放入资料库、skill 包或运行缓存目录，除非用户另行指定路径。

## 常用命令

新 Mac 先准备环境：

```bash
./scripts/install_environment.sh
python3 scripts/check_environment.py
```

建立或刷新轻量索引：

```bash
python3 scripts/gongwen-agent/gongwen_agent.py index --ocr --ocr-max-pages 3 --coverage-per-type 5
```

命令行生成：

```bash
python3 scripts/gongwen-agent/gongwen_agent.py draft project.json
```

启动网页表单：

```bash
python3 scripts/gongwen-agent/gongwen_agent.py web --host 127.0.0.1 --port 8765
```

核查请示专项生成与校验：

```bash
python3 scripts/gongwen-agent/make_hecha_docx.py scripts/gongwen-agent/sample_hecha_project.json --out /tmp/核查请示.docx --audit
python3 scripts/gongwen-agent/check_hecha_docx.py /tmp/核查请示.docx
```

## 验证

常规修改后至少运行：

```bash
python3 -m py_compile scripts/gongwen-agent/gongwen_agent.py scripts/gongwen-agent/gongwen_format.py scripts/gongwen-agent/make_hecha_docx.py scripts/gongwen-agent/check_hecha_docx.py scripts/gongwen-agent/check_template_route.py
python3 -m unittest discover -s scripts/gongwen-agent -p 'test_*.py'
python3 scripts/check_gongwen_workspace.py
```

如修改核查请示固定版式或附件/落款逻辑，额外运行专项生成和校验命令。

## OCR

需要 OCR 时默认使用本地 Paddle OCR，路径为 `/Users/drevan01/Desktop/OCR`。如果当前项目脚本使用 tesseract、macOS Vision 或其他 OCR 作为实际入口，先向用户说明本次使用的工具和原因。
