# 字体资产说明

本目录用于记录文物影响评估 Word 成稿的字体环境要求，不随 skill 分发商业或系统字体文件。

## 固定要求

- 中文字体：宋体。
- 西文和数字字体：Times New Roman。
- 正式 Word 成稿应保留上述字体名，不因目标机器缺失字体而自动改用其他字体。

## 迁移规则

- 迁移 skill 到新机器后，先运行 `scripts/check_font_environment.py` 检查字体环境。
- 如果缺少宋体或 Times New Roman，应提示用户安装系统字体、Office 字体或由用户所在单位确认的正版字体包。
- 不把宋体、Times New Roman 等系统或商业字体文件直接放入 skill 包内迁移。
- 临时预览可以使用目标机器的替代字体显示，但正式提交前必须在 Word 中确认字体效果。

## 检查命令

```bash
python3 scripts/check_font_environment.py
```

如需把缺失字体作为迁移验收失败项：

```bash
python3 scripts/check_font_environment.py --严格
```
