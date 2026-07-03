# project_Automated-Construction-Archaeology
有标准、能自动化、可复用、可管理的智能化执行体系。

## Skills

仓库根目录仅保留两个智能勘探技能，便于直接复制与按版本更新：

| 中文显示名 | Skill name | 版本 | 用途 |
| --- | --- | --- | --- |
| 智能勘探报告 | `smart-prospecting-report` | `v1.0.1` | 生成考古调查勘探报告 DOCX 和检查成果 |
| 智能勘探工作计划 | `smart-prospecting-plan` | `v1.0.1` | 生成考古调查勘探工作计划 DOCX 和检查成果 |

## 安装

两个智能勘探 skill 位于仓库根目录。把需要使用的 skill 文件夹复制到本机 Codex skills 目录：

```bash
mkdir -p ~/.codex/skills
cp -R smart-prospecting-report ~/.codex/skills/
cp -R smart-prospecting-plan ~/.codex/skills/
```

复制后重启 Codex，使 skill description 和中文显示名被重新读取。
