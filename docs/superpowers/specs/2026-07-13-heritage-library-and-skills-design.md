# 文物影响评估资料库与 Skill 分离设计

## 目标

将 `/Users/drevan01/Desktop/影响评估与保护方案skill` 重建为纯专业资料库；影响评估与保护方案两个 Skill 只在常用 Git 仓库维护，并同步安装到本机 Codex。

## 权威位置

- Git 维护源：`/Users/drevan01/Desktop/project_Automated-Construction-Archaeology-gongwen-ready`
- 本机安装副本：`/Users/drevan01/.codex/skills`
- 专业资料库：`/Users/drevan01/Desktop/文物影响评估与保护方案资料库`

Git 仓库是 Skill 的唯一维护源。本机安装目录只保存可运行副本，不反向作为长期维护源。

## Skill 标准结构

两个 Skill 均位于 Git 仓库根目录，采用同一结构：

```text
skill-name/
├── SKILL.md
├── agents/openai.yaml
├── assets/
├── references/
├── scripts/
├── tests/
├── CHANGELOG.md
└── .gitignore
```

- `assets/`：会被复制、填写或用于生成成果的模板和静态资源。
- `references/`：Codex 执行任务时按条件读取的规则、边界和说明。
- `scripts/`：OCR、抽取、生成和校验等确定性工具。
- `tests/`：脚本回归和包结构校验。
- 原始报告、论文、法规 PDF、项目案例不得放入 Skill。
- 不保留 `references/source-originals/`、`.DS_Store`、`__pycache__`、`.pytest_cache` 和运行输出。

## 资料库结构

```text
文物影响评估与保护方案资料库/
├── 00_资料库说明与索引/
├── 01_法规政策与标准/
│   ├── 法律法规与政策/
│   ├── 法规条文库/
│   └── 环境与施工标准资料库/
├── 02_方法论与专题研究/
│   ├── 文物影响评估方法论/
│   ├── 参考论文/
│   └── 相关书籍/
├── 03_文物对象与保护规划资料/
├── 04_项目案例库/
│   ├── 觉海寺改扩建项目/
│   ├── 正蓝旗中央厨房项目/
│   └── 萨拉乌苏公路项目/
├── 05_报告与方案样本/
│   ├── 文物影响评估报告/
│   ├── 文物保护方案/
│   └── 评审意见与修改对照/
└── 06_结构化知识与专题索引/
```

完整项目案例保留项目内部证据链，不拆散项目材料。独立样本报告按成果类型归档。结构化条文库、知识笔记和检索索引保留，但工具运行环境和生成缓存删除。

## 删除边界

直接删除：`.DS_Store`、`__pycache__`、`.pytest_cache`、`tmp/`、`.codex_tmp_*`、临时渲染页、空演练目录、一次性生成脚本、可重新安装的旧 PDF/OCR 工具环境、历史 Skill 副本和迁移包。

不得按过程文件删除：最终报告、评审意见、修改对照、项目审批材料、调查勘探报告、保护规划、法规标准和具有复用价值的案例版本。

## 验证

- 迁移前后对保留文件生成数量、大小和 SHA-256 清单。
- 两个 Skill 通过包结构检查和现有测试。
- Git 仓库与本机安装副本内容一致（缓存除外）。
- 资料库根目录不再出现 Skill、代码工具、缓存或临时目录。
