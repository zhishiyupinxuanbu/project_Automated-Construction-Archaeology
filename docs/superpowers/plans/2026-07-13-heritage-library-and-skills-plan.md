# 文物影响评估资料库与 Skill 分离 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将混合工作目录重建为纯专业资料库，并把两个 Skill 规范化到常用 Git 仓库后同步安装到 Codex。

**Architecture:** Git 仓库是 Skill 唯一维护源，`~/.codex/skills` 是安装镜像，桌面资料库只保存专业资料。迁移采用清单先行、保留资料移动、明确缓存删除、最终哈希与测试校验的顺序。

**Tech Stack:** zsh、Python 3、pytest、Git、SHA-256 文件清单。

## Global Constraints

- 不修改常用仓库中本次范围以外的 Skill。
- 不删除最终报告、评审意见、项目审批材料和完整项目证据链。
- 原始报告、论文、法规 PDF 和项目案例不得留在 Skill 包内。
- 本机 Paddle OCR 固定路径 `/Users/drevan01/Desktop/OCR` 保持不变。
- 不推送远程仓库。

---

### Task 1: 建立迁移基线

**Files:**
- Create: `/Users/drevan01/Desktop/影响评估与保护方案skill/.整理清单/pre-migration-files.tsv`
- Create: `/Users/drevan01/Desktop/影响评估与保护方案skill/.整理清单/pre-migration-summary.txt`

- [ ] 记录所有文件的相对路径、大小和修改时间。
- [ ] 记录顶层目录体量与 Git 初始状态。
- [ ] 运行影响评估 Skill 现有测试，确认基线。

### Task 2: 用测试锁定 Skill 标准结构

**Files:**
- Create: `/Users/drevan01/Desktop/project_Automated-Construction-Archaeology-gongwen-ready/heritage-impact-assessment/tests/test_skill_package_layout.py`
- Create: `/Users/drevan01/Desktop/project_Automated-Construction-Archaeology-gongwen-ready/heritage-protection-plan/tests/test_skill_package_layout.py`

- [ ] 先写包结构测试，要求核心目录齐全且禁止缓存、原始成稿库和旧资料库路径。
- [ ] 运行测试并确认保护方案目录缺失、影响评估仍含旧引用时测试失败。

### Task 3: 规范 Git 中两个 Skill

**Files:**
- Modify: `/Users/drevan01/Desktop/project_Automated-Construction-Archaeology-gongwen-ready/heritage-impact-assessment/**`
- Create: `/Users/drevan01/Desktop/project_Automated-Construction-Archaeology-gongwen-ready/heritage-protection-plan/**`
- Modify: `/Users/drevan01/Desktop/project_Automated-Construction-Archaeology-gongwen-ready/README.md`
- Create: `/Users/drevan01/Desktop/project_Automated-Construction-Archaeology-gongwen-ready/.gitignore`

- [ ] 将本机安装版的最新影响评估内容同步到 Git，排除 `source-originals` 和缓存。
- [ ] 将保护方案迁入 Git，排除 `source-originals` 和缓存。
- [ ] 把原始样本路由改为外部资料库，更新旧绝对路径。
- [ ] 统一 `CHANGELOG.md`、测试目录和忽略规则。
- [ ] 运行包结构测试与两个 Skill 的全部测试。

### Task 4: 重建专业资料库

**Files:**
- Create: `/Users/drevan01/Desktop/文物影响评估与保护方案资料库/**`
- Remove after migration: `/Users/drevan01/Desktop/影响评估与保护方案skill/**`

- [ ] 创建六类专业资料目录和说明索引。
- [ ] 移动法规标准、方法论文献、文物对象资料、完整项目案例、独立报告样本和结构化知识。
- [ ] 保留完整项目内部现有证据链。
- [ ] 删除 Skill 历史副本、迁移包、工具环境、缓存、临时输出、空目录和一次性脚本。
- [ ] 生成迁移后的文件清单，并核对保留对象。

### Task 5: 同步安装并完成验证

**Files:**
- Replace: `/Users/drevan01/.codex/skills/heritage-impact-assessment/**`
- Replace: `/Users/drevan01/.codex/skills/heritage-protection-plan/**`

- [ ] 从 Git 维护源同步两个 Skill 到本机安装目录。
- [ ] 再次运行包结构检查和全部测试。
- [ ] 检查资料库根目录不存在 Skill、脚本、缓存和临时输出。
- [ ] 保存 Git 状态、资料库体量和最终校验摘要。
