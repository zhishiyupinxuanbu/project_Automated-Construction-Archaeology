#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import subprocess
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path


PROJECT_RE = re.compile(r"^\d{8}-.+")
RED_HEAD_HINTS = ("请示", "函", "意见", "报告", "验收", "开工", "发掘")
DOC_EXTS = {".doc", ".docx", ".pdf", ".ofd"}


REQUIRED_TOP = ["1.项目资料", "2.商务资料", "3.执行资料", "4.成果资料"]
EXEC_DIRS = ["1.外业成果", "2.CAD工程", "3.内业成果", "4.报告"]
FIELD_PHOTO_DIRS = [
    "1.项目地块现状照",
    "2.走访调查照",
    "3.实地踏查照",
    "4.信息采集照",
    "5.勘探单元布置照",
    "6.布设探孔照",
    "7.普探工作照",
    "8.取样记录照",
    "9.勘探后局部照",
    "10.勘探后航拍照",
    "11.文献资料收集与整理工作照",
    "12.标准孔照",
    "13.遗迹",
    "14.资料整理工作照",
]
DEFAULT_REPORT_DIR = Path("/Users/hero/Desktop/NAS归档检查报告")


@dataclass
class Finding:
    severity: str
    path: str
    message: str


@dataclass
class ProjectAudit:
    path: Path
    findings: list[Finding] = field(default_factory=list)
    file_count: int = 0
    dir_count: int = 0

    def add(self, severity: str, path: Path, message: str):
        self.findings.append(Finding(severity, str(path), message))


def run(cmd: list[str]) -> str:
    try:
        return subprocess.check_output(cmd, text=True, stderr=subprocess.DEVNULL)
    except Exception:
        return ""


def mounted_smb_shares() -> list[Path]:
    shares: list[Path] = []
    for line in run(["mount"]).splitlines():
        if "smbfs" not in line or " on /Volumes/" not in line:
            continue
        mount_point = line.split(" on ", 1)[1].split(" (", 1)[0]
        p = Path(mount_point)
        if p.exists() and p.is_dir():
            shares.append(p)
    return sorted(set(shares), key=lambda p: str(p))


def list_mounts(_args: argparse.Namespace) -> int:
    shares = mounted_smb_shares()
    if not shares:
        print("No SMB shares mounted under /Volumes.")
        return 1
    for share in shares:
        print(share)
    return 0


def mkdirs(paths: list[Path]):
    for p in paths:
        p.mkdir(parents=True, exist_ok=True)


def create_project(args: argparse.Namespace) -> int:
    root = Path(args.root).expanduser()
    project = root / args.name
    if not root.exists():
        raise SystemExit(f"Root does not exist: {root}")
    dirs = [
        project / "1.项目资料",
        project / "2.商务资料" / "合同扫描",
        project / "3.执行资料" / "1.外业成果",
        project / "3.执行资料" / "2.CAD工程",
        project / "3.执行资料" / "3.内业成果" / "图纸" / "单个标准孔",
        project / "3.执行资料" / "3.内业成果" / "图纸" / "单个遗迹",
        project / "3.执行资料" / "3.内业成果" / "图纸" / "遗迹分布示意图",
        project / "3.执行资料" / "3.内业成果" / "表格",
        project / "3.执行资料" / "4.报告",
        project / "4.成果资料" / "1.验收请示",
        project / "4.成果资料" / "2.三级联调",
        project / "4.成果资料" / "3.用地申请" / "附件",
        project / "4.成果资料" / "4.发掘申请" / "附件",
    ]
    field_root = project / "3.执行资料" / "1.外业成果"
    for name in FIELD_PHOTO_DIRS:
        dirs.append(field_root / name)
    if args.with_relic_subdirs:
        dirs.extend([field_root / "13.遗迹" / "土样照", field_root / "13.遗迹" / "现场照"])
    mkdirs(dirs)
    print(project)
    return 0


def find_projects(root: Path) -> list[Path]:
    if not root.exists() or not root.is_dir():
        return []
    projects = []
    for child in root.iterdir():
        if child.is_dir() and (PROJECT_RE.match(child.name) or child.name.endswith("项目")):
            projects.append(child)
    return sorted(projects, key=lambda p: p.name)


def has_any_file(path: Path) -> bool:
    return path.exists() and any(p.is_file() for p in path.rglob("*"))


def audit_project(project: Path) -> ProjectAudit:
    audit = ProjectAudit(path=project)
    all_dirs = [p for p in project.rglob("*") if p.is_dir()]
    all_files = [p for p in project.rglob("*") if p.is_file()]
    audit.dir_count = len(all_dirs)
    audit.file_count = len(all_files)

    for name in REQUIRED_TOP:
        p = project / name
        if not p.is_dir():
            audit.add("ERROR", p, "缺少项目一级目录")

    commerce = project / "2.商务资料" / "合同扫描"
    if not commerce.is_dir():
        audit.add("WARN", commerce, "缺少合同扫描目录")

    exec_root = project / "3.执行资料"
    for name in EXEC_DIRS:
        p = exec_root / name
        if not p.is_dir():
            audit.add("ERROR", p, "执行资料缺少真实模板目录")

    field_root = exec_root / "1.外业成果"
    if field_root.is_dir():
        zone_dirs = [p for p in field_root.iterdir() if p.is_dir() and p.name.startswith("勘探区域")]
        photo_roots = zone_dirs or [field_root]
        for base in photo_roots:
            for name in FIELD_PHOTO_DIRS:
                p = base / name
                if not p.is_dir():
                    audit.add("WARN", p, "外业成果照片分类不完整")
            relic = base / "13.遗迹"
            if relic.is_dir() and has_any_file(relic) and not ((relic / "土样照").is_dir() and (relic / "现场照").is_dir()):
                audit.add("WARN", relic, "遗迹目录已有文件，但缺少土样照/现场照分组")

    indoor = exec_root / "3.内业成果"
    for p, msg in [
        (indoor / "图纸", "内业成果缺少图纸目录"),
        (indoor / "图纸" / "单个标准孔", "图纸缺少单个标准孔目录"),
        (indoor / "表格", "内业成果缺少表格目录"),
    ]:
        if not p.is_dir():
            audit.add("WARN", p, msg)

    tables = indoor / "表格"
    if tables.is_dir():
        expected_table_hints = ["标准孔坐标", "勘探单元", "四至范围坐标"]
        names = [p.name for p in tables.iterdir() if p.is_file()]
        for hint in expected_table_hints:
            if not any(hint in name for name in names):
                audit.add("INFO", tables, f"未发现常见表格：{hint}")

    report_dir = exec_root / "4.报告"
    if report_dir.is_dir() and not has_any_file(report_dir):
        audit.add("INFO", report_dir, "报告目录为空，确认项目阶段是否尚未形成报告")

    project_materials = project / "1.项目资料"
    if project_materials.is_dir():
        for file in project_materials.rglob("*"):
            if not file.is_file() or file.suffix.lower() not in DOC_EXTS:
                continue
            if any(hint in file.name for hint in RED_HEAD_HINTS):
                audit.add("WARN", file, "疑似红头/报送文件位于 1.项目资料；若为公司生成文件，应放 4.成果资料")

    return audit


def render_report(audits: list[ProjectAudit], roots: list[Path]) -> str:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        "# NAS 项目资料归档检查报告",
        "",
        f"- 检查时间：{now}",
        f"- 检查根目录：{', '.join(str(r) for r in roots)}",
        f"- 项目数量：{len(audits)}",
        "",
    ]
    total = sum(len(a.findings) for a in audits)
    by_sev: dict[str, int] = {}
    for audit in audits:
        for finding in audit.findings:
            by_sev[finding.severity] = by_sev.get(finding.severity, 0) + 1
    lines.extend([
        "## 汇总",
        "",
        f"- 发现项总数：{total}",
        f"- ERROR：{by_sev.get('ERROR', 0)}",
        f"- WARN：{by_sev.get('WARN', 0)}",
        f"- INFO：{by_sev.get('INFO', 0)}",
        "",
    ])
    for audit in audits:
        lines.extend([
            f"## {audit.path}",
            "",
            f"- 文件夹数：{audit.dir_count}",
            f"- 文件数：{audit.file_count}",
            "",
        ])
        if not audit.findings:
            lines.extend(["未发现明显归档问题。", ""])
            continue
        for finding in audit.findings:
            lines.append(f"- `{finding.severity}` {finding.message}：`{finding.path}`")
        lines.append("")
    return "\n".join(lines)


def default_report_path() -> Path:
    stamp = datetime.now().strftime("%Y%m%d")
    return DEFAULT_REPORT_DIR / f"nas-audit-{stamp}.md"


def write_report(text: str, output: str | None) -> None:
    if not output:
        output = str(default_report_path())
    if output:
        out = Path(output).expanduser()
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text, encoding="utf-8")
        print(out)


def audit_roots(roots: list[Path], output: str | None) -> int:
    audits: list[ProjectAudit] = []
    for root in roots:
        for project in find_projects(root):
            audits.append(audit_project(project))
    text = render_report(audits, roots)
    write_report(text, output)
    return 0


def audit(args: argparse.Namespace) -> int:
    return audit_roots([Path(r).expanduser() for r in args.roots], args.output)


def audit_mounted(args: argparse.Namespace) -> int:
    roots = mounted_smb_shares()
    return audit_roots(roots, args.output)


def main() -> int:
    parser = argparse.ArgumentParser(description="Create and audit NAS project archives.")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("list-mounts")
    p.set_defaults(func=list_mounts)

    p = sub.add_parser("create-project")
    p.add_argument("--root", required=True)
    p.add_argument("--name", required=True)
    p.add_argument("--with-relic-subdirs", action="store_true")
    p.set_defaults(func=create_project)

    p = sub.add_parser("audit")
    p.add_argument("--roots", nargs="+", required=True)
    p.add_argument("--output")
    p.set_defaults(func=audit)

    p = sub.add_parser("audit-mounted")
    p.add_argument("--output")
    p.set_defaults(func=audit_mounted)

    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
