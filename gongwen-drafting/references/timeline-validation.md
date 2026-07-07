# 勘探请示时间校验硬闸门

## 适用文种

生成以下公文前必须执行时间校验：

- 申请开展考古勘探工作请示
- 勘探计划备案请示
- 其他实际属于勘探请示链路的请示

相关来源文件包括文物调查报告、文物复函/回函、勘探计划和勘探报告。

## 硬规则

1. 勘探请示必须在文物调查工作结束之后。
2. 勘探请示日期不得早于企业取得文物复函/回函日期；同一天可以，取得回函后 2-3 天再报也可以。
3. 勘探计划中所写勘探开始日期应当在打请示 2-3 天后。
4. 任一时间缺失、冲突或顺序不合规时，必须停止生成并报错给用户，不得自动调整日期或继续出 Word。

## 可填写字段

`project.json` 和本地 Web 表单可填写：

- `request_date`：请示日期
- `reply_date`：文物复函/回函日期
- `investigation_end_date`：文物调查结束日期
- `planned_start_date`：计划勘探开始日期
- `planned_end_date`：计划勘探结束日期
- `prospecting_time`：勘探时间范围

CLI 也可传入项目源资料目录，让脚本补充读取调查报告、回函/复函、勘探请示、勘探计划和勘探报告中的日期：

```bash
python3 scripts/gongwen-agent/gongwen_agent.py draft project.json --project-dir /path/to/project-source
```

## 工具入口

独立校验脚本：

```bash
python3 scripts/gongwen-agent/validate_project_timeline.py --project-json project.json --project-dir /path/to/project-source --doc-type 勘探请示
```

检查已有计划或报告 Word 内嵌图片水印日期时，传入 `--docx`；脚本会抽取 Word 内 `word/media/` 图片，并默认调用本地 Paddle OCR 尝试识别水印日期。识别到水印日期的图片参与范围校验，未识别到水印日期的图片会汇总提示用户人工核对：

```bash
python3 scripts/gongwen-agent/validate_project_timeline.py --project-json project.json --docx /path/to/计划或报告.docx --doc-type 勘探请示
```

该脚本返回非零时，必须中止生成。
