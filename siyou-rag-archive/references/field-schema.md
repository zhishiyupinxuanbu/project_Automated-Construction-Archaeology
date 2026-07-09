# 字段设计 v1.0

字段分成五张核心表，避免把文件、页面、图片、文书、候选卷别混成一张不可维护的大表。

## file_manifest

每个原始文件一行。

| 字段 | 说明 |
| --- | --- |
| file_id | 稳定 ID，可用哈希或序号 |
| original_path | 原始文件绝对路径 |
| original_name | 原始文件名 |
| extension | 扩展名 |
| file_size | 文件大小 |
| sha256 | 文件哈希 |
| source_year | 来源年度，如 2025/2026/未知 |
| source_folder | 来源文件夹 |
| material_kind | 文书/表格/图片/PDF/图纸/压缩包/其他 |
| process_status | 未处理/已抽文本/已OCR/待展开/失败/待核 |
| needs_manual_review | 是否需人工复核 |
| notes | 备注 |

## page_evidence

每页或每个可定位证据块一行。

| 字段 | 说明 |
| --- | --- |
| page_id | 页级 ID |
| file_id | 关联 file_manifest |
| page_no | 页号 |
| page_image_path | 页图路径 |
| ocr_text_path | OCR 文本路径 |
| ocr_text | OCR 全文或主要文本 |
| ocr_confidence | OCR 置信度 |
| is_blank | 是否空白页 |
| has_redhead | 是否疑似红头 |
| has_seal | 是否疑似印章 |
| has_table | 是否表格 |
| has_photo | 是否照片页 |
| evidence_excerpt | 支持判断的原文片段 |

## doc_extract

每个独立文书、表格、报告或材料单元一行。

| 字段 | 说明 |
| --- | --- |
| doc_id | 文档单元 ID |
| file_id | 来源文件 |
| candidate_title | 候选题名 |
| formal_title | 人工确认正式题名 |
| document_no | 文号，按原文括号形态转录 |
| document_date | 成文日期 |
| signing_date | 落款日期 |
| issuer | 发文单位 |
| receiver | 收文单位 |
| responsible_unit | 责任单位/建设单位/签署主体 |
| heritage_sites | 涉及文保点 |
| township | 涉及乡镇 |
| document_type | 通知/报告/函/请示/批复/责任书/合同/简报/统计表/调查表/规划/图纸/照片等 |
| keywords | 关键词 |
| page_count | 页数 |
| field_completeness | 字段完整度 |
| uncertain_fields | 待核字段 |

## classification_score

一个 doc 可有多个候选卷别。

| 字段 | 说明 |
| --- | --- |
| doc_id | 关联文档单元 |
| candidate_volume | 候选卷号 |
| candidate_volume_name | 候选卷名 |
| rule_hits | 命中的分类规则 |
| field_hits | 命中的关键字段 |
| keyword_hits | 命中的关键词 |
| path_signal | 来源路径辅助判断 |
| similar_examples | RAG 召回相似样例 |
| score_total | 总分 |
| confidence_level | 高/中/低/待核 |
| suggested_include | 是否建议归入 |
| evidence_page_ids | 支撑页 |
| manual_volume | 人工确认卷别 |
| manual_comment | 人工意见 |

## media_asset

图片、CAD、压缩包内媒体等特殊材料。

| 字段 | 说明 |
| --- | --- |
| asset_id | 媒体 ID |
| file_id | 来源文件 |
| media_type | 现场照片/文物本体照片/会议照片/资料检查照片/拓片/图纸/CAD/压缩包 |
| image_path | 图片或转换图路径 |
| thumbnail_path | 缩略图路径 |
| exif_time | EXIF 时间 |
| image_size | 尺寸 |
| contains_text | 是否含文字 |
| ocr_text | 图片 OCR 文本 |
| visual_summary | 图片内容简述 |
| related_site | 关联文保点 |
| related_activity | 关联活动 |
| candidate_volume | 候选卷别 |
| duplicate_group | 重复图片组 |
| quality_flag | 模糊/低质/正常 |
| manual_note | 人工说明 |
