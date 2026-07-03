# project_Automated-Construction-Archaeology
有标准、能自动化、可复用、可管理的智能化执行体系。

## Skills

所有 skill 文件夹均放在仓库根目录；旧的 `skills/` 汇总目录不再使用。

| 中文显示名 | Skill name | 版本 | 用途 |
| --- | --- | --- | --- |
| 智能勘探报告 | `smart-prospecting-report` | `v1.0.1` | 生成考古调查勘探报告 DOCX 和检查成果 |
| 智能勘探工作计划 | `smart-prospecting-plan` | `v1.0.1` | 生成考古调查勘探工作计划 DOCX 和检查成果 |
| 文物影响评估 | `heritage-impact-assessment` | `v0.3.1` | 文物影响评估材料处理与报告生成 |

## 安装

把需要使用的 skill 文件夹复制到本机 Codex skills 目录：

```bash
mkdir -p ~/.codex/skills
cp -R smart-prospecting-report ~/.codex/skills/
cp -R smart-prospecting-plan ~/.codex/skills/
```

复制后重启 Codex，使 skill description 和中文显示名被重新读取。
