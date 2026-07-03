# Word 字体与版式迁移要求 v0.1

本文件用于 skill 迁移到新机器后检查 Word 成稿环境，避免报告正文生成后因字体缺失、替换或目录处理方式不同导致版式偏差。

## 字体打包边界

- skill 不内置、不复制、不分发宋体、Times New Roman 等系统或商业字体文件。
- `assets/fonts/README.md` 只记录字体环境要求和迁移检查方式。
- 如果目标机器缺少所需字体，应由用户安装系统字体、Office 字体或单位确认的正版字体包。
- AI 不应自动下载、打包或替换商业字体。

## 成稿字体要求

- 中文文字使用宋体。
- 西文和数字使用 Times New Roman。
- 正文字号小四。
- 正文行距固定值 20 磅。
- 图片所在行为单倍行距。
- 标题、题注、附件版式继续按 `references/10-docx成稿规则.md` 执行。

## 迁移后检查流程

1. 在新机器上进入 skill 目录。
2. 运行：

```bash
python3 scripts/check_font_environment.py
```

3. 若需要把字体缺失作为迁移失败条件，运行：

```bash
python3 scripts/check_font_environment.py --严格
```

4. 如果脚本提示字体缺失，不直接修改报告字体；应在交付说明中提醒用户补装字体并在 Word 中复核。
5. 生成 Word 后，仍需人工检查封面、目录预留页、正文、题注、图片和附件版式。

## 缺失处理原则

- 缺少宋体：提示用户补装宋体、SimSun 或单位认可的同名字体环境；正式成稿仍要求宋体。
- 缺少 Times New Roman：提示用户补装 Times New Roman 或 Office 字体环境；正式成稿仍要求 Times New Roman。
- 不能因本机缺字而在 skill 模板中永久改成其他字体。
- 若项目临时只需预览稿，可在交付说明中标明“当前机器字体环境待复核”。
