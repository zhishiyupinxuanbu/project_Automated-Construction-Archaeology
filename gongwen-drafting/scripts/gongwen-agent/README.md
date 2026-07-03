# 专属公文撰写 Agent

这是资料库内的本地 MVP 工具，目标是把现有历史公文和模板转成可检索知识库，并根据项目要素生成符合资料库固定版式规则的 Word 成稿。

完整运行环境、输入字段、建库命令和复核清单见：[使用说明](./使用说明.md)。

迁移到其他电脑使用时，见：[迁移打包工具](./迁移打包工具/)；迁移成品只保留资料库根目录下重新生成的 `公文撰写Agent给AI助手迁移包.zip`。后续迁移默认目标电脑已有 Codex、Trae Solo 等 AI 工具，不再制作“完整运行包”。

## 快速开始

使用 Codex bundled Python 或系统 Python 均可，当前环境已具备 `python-docx` 和 `pdftotext`。

```bash
python3 agent/gongwen_agent.py index
python3 agent/gongwen_agent.py web --port 8765
```

打开：

```text
http://127.0.0.1:8765
```

生成结果默认写入电脑桌面；只有在无桌面环境或用户另行指定路径时，才使用任务输出目录。公文写作链路直接按资料库文种手册、地区模板和固定版式脚本生成 `.docx`，不再调用外部一键格式化工具：

```text
~/Desktop/
```

## 命令行生成

先准备一个 JSON：

```json
{
  "business_type": "文物保护许可申请",
  "issuing_org": "XX公司",
  "recipient_org": "准格尔旗文物局",
  "project_name": "XX项目",
  "construction_unit": "XX公司",
  "location": "鄂尔多斯市准格尔旗XX镇",
  "scale": "新建线路长度约10公里，配套建设相关设施",
  "land_area": "0.12公顷",
  "approval_basis": "项目核准文件",
  "contact": "张三",
  "phone": "13800000000",
  "special_notes": ""
}
```

然后运行，桌面只会保留最终 Word：

```bash
python3 agent/gongwen_agent.py draft project.json
```

## 既有公文格式调整

给已经写好的 Word 公文调格式时，按当前文种手册、地区模板和固定版式脚本做外科式局部修正，不再使用通用一键格式化工具重排整篇。

## 固定版式核查请示与跨文种附件/落款

跨电脑稳定生成“文物保护许可申请/核查请示”时，优先使用固定版式脚本，不要只靠自然语言规则：

```bash
python3 agent/make_hecha_docx.py agent/sample_hecha_project.json --out /tmp/核查请示.docx --audit
python3 agent/check_hecha_docx.py /tmp/核查请示.docx
```

详细说明见：[固定版式说明](./固定版式说明.md)。

红头、附件说明和落款是资料库专项生成工具中的固定版式区域。生成勘探验收、开工请示、备案请示等 Word 公文时，可按对应模板或专项脚本处理；既有成稿需要调格式时，也应按当前文种和问题区域外科式处理。

## 已实现

- DOCX 模板直接解析入库。
- PDF 使用 `pdftotext` 尝试解析；扫描件可通过 OCR 入库。
- OCR 优先使用 `tesseract`，已安装中文语言包时自动使用 `chi_sim+chi_tra+eng`；缺少 tesseract 时会尝试 macOS Vision 备用方案。
- 按事项类型、旗区、项目关键词检索模板和历史参考件。
- 自动生成正文文本和 Word 成稿，使用资料库固定版式规则输出最终 `.docx`。
- 使用地区模板路由门禁和固定版式脚本降低模板错用、附件缩进和落款漂移风险。
- 对项目名称、建设单位、建设地点、用地面积、发文主体、主送机关做一致性校验。

## OCR 重建索引

全量 OCR：

```bash
python3 agent/gongwen_agent.py index --force --ocr
```

更快的可用模式，只识别每个 PDF 前 3 页：

```bash
python3 agent/gongwen_agent.py index --ocr --ocr-max-pages 3
```

轻量覆盖模式：每类公文最多入库 5 篇，优先模板和已解析文件，适合先把 Agent 跑顺：

```bash
python3 agent/gongwen_agent.py index --ocr --ocr-max-pages 3 --coverage-per-type 5
```

如果某个 PDF 转图片超时，工具会跳过该文件并保留为 `needs_ocr`，不会中断整个建库。

## 下一步建议

- 补充每类公文的字段模板和审批依据知识，提升正文准确度。
- 增加正式审核模式：输出修改痕迹、风险提示和引用来源。
