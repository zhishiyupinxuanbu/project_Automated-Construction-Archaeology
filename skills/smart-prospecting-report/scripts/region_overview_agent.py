#!/usr/bin/env python3
"""Generate county-level region overview text through configured LLM APIs.

The public entry points in this module are intentionally small so they can be
wrapped by an MCP tool later:

- ``generate_region_overview(fields)`` for Python callers.
- ``python region_overview_agent.py --input-json request.json`` for JSON callers.
- stdin JSON when ``--input-json -`` is used.

Only the county name and generic writing rules are sent to the external API.
Project names, coordinates, project locations, construction details, photos,
and user-filled reference text must stay local. The generated text must
describe the county only, not the specific project.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any


DOUBAO_ARK_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3"
DOUBAO_ARK_MODEL = "doubao-seed-2-0-lite-260215"
DEEPSEEK_BASE_URL = "https://api.deepseek.com"
DEEPSEEK_MODEL = "deepseek-chat"
OPENAI_RESPONSES_MODEL = "gpt-5.2"
REGION_OVERVIEW_KEYS = [
    "项目所在地旗县地理位置概况",
    "项目所在地旗县行政区划与社会经济概况",
    "项目所在地旗县气候条件",
    "项目所在地旗县历史沿革",
]
REGION_FIELD_SPECS = {
    "项目所在地旗县地理位置概况": (
        "地理位置",
        "写成 2-3 个自然段，550-650 个汉字，覆盖行政隶属、地理方位、邻接关系、地貌位置、面积、名称由来、交通区位、自然资源、历史文化资源或生态资源。",
        450,
    ),
    "项目所在地旗县行政区划与社会经济概况": (
        "行政区划与社会经济概况",
        "写成 1-2 个自然段，330-400 个汉字，覆盖镇级行政区划、主导产业、产业结构、资源开发、城镇基础、交通条件和居民生产生活概况。",
        250,
    ),
    "项目所在地旗县气候条件": (
        "气候条件",
        "写成 1-2 个自然段，330-420 个汉字，覆盖地形地貌、气候类型、气温、降水、无霜期、日照、风沙、蒸发等指标，不写作业建议。",
        250,
    ),
    "项目所在地旗县历史沿革": (
        "历史沿革",
        "写成 8-11 个自然段，1250-1400 个汉字，按秦汉、魏晋南北朝、隋唐、辽金元、明清、民国、中华人民共和国成立后等阶段组织，重点写行政隶属、建置沿革、族群活动和区域交通/军政地位。",
        1000,
    ),
}


def clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\ufeff", "")
    text = re.sub(r"[\u200b\u200c\u200d\u2060]", "", text)
    return text.strip()


def script_root() -> Path:
    return Path(__file__).resolve().parents[2]


def load_env_file(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        key, value = stripped.split("=", 1)
        key = key.strip()
        if key and key not in os.environ:
            os.environ[key] = value.strip().strip('"').strip("'")


def load_project_env() -> None:
    root = script_root()
    load_env_file(root / ".env.local")
    load_env_file(root / "智能生成报告技能资料" / ".env.local")


def infer_county(fields: dict[str, Any]) -> str:
    direct = clean(fields.get("项目所在地旗县"))
    if direct:
        direct_matches = re.findall(r"(?:^|省|自治区|市|盟|州|地区)([\u4e00-\u9fff]{2,8}(?:旗|县|市|区))", direct)
        return direct_matches[-1] if direct_matches else direct
    text = clean(fields.get("项目位置"))
    matches = re.findall(r"(?:^|省|自治区|市|盟|州|地区)([\u4e00-\u9fff]{2,8}(?:旗|县|市|区))", text)
    for match in reversed(matches):
        if match not in {"自治区", "开发区", "工业区", "园区"}:
            return match
    return ""


REGION_RESEARCH_KEYS = [
    "地理位置与邻接",
    "地形地貌与自然资源",
    "行政区划与交通",
    "社会经济与产业",
    "气候资料",
    "历史沿革资料",
    "需剔除或慎用的信息",
]

REGION_KNOWN_FACTS = {
    "额济纳旗": {
        "行政区划与交通": "下辖3个镇、6个苏木：达来呼布镇、东风镇、哈日布日格德音乌拉镇、苏泊淖尔苏木、赛汉陶来苏木、马鬃山苏木、巴彦陶来苏木、温图高勒苏木、巴音陶海苏木。旗人民政府驻达来呼布镇。",
    }
}


def apply_known_region_facts(county: str, research: dict[str, str]) -> dict[str, str]:
    known = REGION_KNOWN_FACTS.get(county)
    if not known:
        return research
    merged = dict(research)
    for key, value in known.items():
        existing = clean(merged.get(key))
        merged[key] = value if not existing else f"{value}\n{existing}"
    return merged


def enforce_known_region_facts(county: str, fields: dict[str, str]) -> dict[str, str]:
    if county != "额济纳旗":
        return fields
    corrected = dict(fields)
    key = "项目所在地旗县行政区划与社会经济概况"
    value = clean(corrected.get(key))
    if not value:
        return corrected
    correct_sentence = (
        "额济纳旗位于内蒙古自治区阿拉善盟最西端，总面积约11.46万平方公里，"
        "下辖3个镇、6个苏木，即达来呼布镇、东风镇、哈日布日格德音乌拉镇、"
        "苏泊淖尔苏木、赛汉陶来苏木、马鬃山苏木、巴彦陶来苏木、温图高勒苏木、"
        "巴音陶海苏木，旗人民政府驻达来呼布镇。"
    )
    value = re.sub(
        r"额济纳旗位于.*?旗人民政府驻达来呼布镇。",
        correct_sentence,
        value,
        count=1,
    )
    if "下辖3个镇、6个苏木" not in value:
        value = f"{correct_sentence}{value}"
    corrected[key] = value
    return corrected


def build_region_research_prompt(county: str) -> str:
    return f"""
你是考古调查勘探报告的资料检索与资料整理助手。请只围绕“{county}”这个旗县级行政区，整理可用于“项目区域概况”的资料卡。

资料整理要求：
1. 尽量保留可用于报告正文的具体事实，包括行政隶属、地理方位、邻接关系、面积、地貌、河流/湖泊/沙漠/戈壁、矿产或农牧资源、生态/历史文化资源、镇级行政区划、交通、主导产业、气候指标、历史建置沿革等。
2. 只剔除明显错误、相互冲突且无法判断、宣传口号、荣誉名单、旅游广告、网页链接、来源痕迹、与旗县无关的项目资料。
3. 对不确定数字用“约、一般、多在、通常”等稳健口径表达，不要编造不可核验的文号、批复号、统计排名或精确年份。
4. 历史沿革资料要尽量按时间顺序保留，不要只写摘要；秦汉、魏晋南北朝、隋唐、辽金元、明清、民国、中华人民共和国成立后等阶段能写则尽量保留。
5. 出于资料安全要求，本次请求只提供旗县名称，不提供项目名称、项目坐标、项目位置、建设内容、人工填写表参考内容或其他项目资料。

只输出 JSON 对象，键名必须严格为：
{json.dumps(REGION_RESEARCH_KEYS, ensure_ascii=False)}
"""


def build_region_overview_prompt(county: str, research: dict[str, str] | None = None) -> str:
    research_text = ""
    if research:
        research_text = f"""
已整理的资料卡如下。除“需剔除或慎用的信息”外，请尽量吸收和保留资料卡中的可用事实，但必须按报告逻辑重组，不要机械堆砌：
{json.dumps(research, ensure_ascii=False, indent=2)}
"""
    return f"""
你是考古调查勘探报告的“项目区域概况”撰写 agent。请只围绕“{county}”这个旗县级行政区生成正式报告正文，不要结合任何具体建设项目、矿山项目或红线范围情况。
{research_text}

写作任务：
1. “地理位置”按旧报告范文的厚度写成 2-3 个自然段，合计约 550-650 个汉字：先写 {county} 的行政隶属、地理方位、邻接关系、地貌位置和面积；再写名称由来、交通区位、自然资源、历史文化资源或生态资源。不要写具体项目位置。
2. “行政区划与社会经济概况”写 330-400 个汉字，覆盖 {county} 的镇级行政区划、主导产业、产业结构、资源开发、城镇基础、交通条件和居民生产生活概况。不要写荣誉名单、宣传口号或旅游广告。
3. “气候特征”只写 {county} 的地形地貌和气候条件，约 330-420 个汉字，采用旧报告中“气候类型、降水、气温、无霜期、日照、风沙、蒸发”等指标型写法；数据无法稳妥核实时使用“约、一般、呈现”等稳健表述，不写“对考古调查勘探工作影响”“需注意防暑防寒”等作业建议。
4. “历史沿革”按时间顺序写清楚 {county} 的历史沿革情况，约 1250-1400 个汉字，分 8-11 段，优先按秦汉、魏晋南北朝、隋唐、辽金元、明清、民国、中华人民共和国成立后等阶段组织，重点写行政隶属、建置沿革、族群活动和区域交通/军政地位。
5. 四个字段合计控制在 2000-2400 个汉字左右，目标 2200-2350 个汉字。
6. 优先使用资料卡中的公开可靠事实，表述稳健客观；不要写宣传口号、网页链接、来源说明、脚注或 Markdown。
7. 不要编造不可核验的具体文号、批复号、统计数值、荣誉称号或考古结论。涉及 GDP、人口、面积、降水等统计数据时，只使用常见公开口径；不确定时用概括性表述。
8. 以上限制只针对 agent 生成的旗县概况字段，不影响模板中固定的“项目名称坐落于项目位置……”开头句；agent 生成字段本身不得出现“本项目”“该项目”“项目区”“项目所在地”“项目用地”“项目红线”“建设内容”“矿区范围”“坐落于”等把文字指向具体项目的表述。
9. 输出正文应像报告正文，不要像百科词条堆砌；自然资源和社会经济内容要融入段落，不要列清单。
10. 出于资料安全要求，本次请求只提供旗县名称，不提供项目名称、项目坐标、项目位置、建设内容、人工填写表参考内容或其他项目资料。请仅依据旗县级公开常识生成。

只输出 JSON 对象，键名必须严格为：
{json.dumps(REGION_OVERVIEW_KEYS, ensure_ascii=False)}
"""


def build_region_overview_repair_prompt(
    county: str,
    previous: dict[str, str],
    warnings: list[str],
    research: dict[str, str] | None = None,
) -> str:
    research_note = ""
    if research:
        research_note = f"""
可用资料卡如下。除“需剔除或慎用的信息”外，请尽量保留资料卡里的事实，并按报告逻辑重组：
{json.dumps(research, ensure_ascii=False, indent=2)}
"""
    return f"""
你上一次生成的“{county}”项目区域概况未达到字数硬性要求，请根据下面的问题重写完整 JSON。
{research_note}

必须修正的问题：
{chr(10).join(f"- {warning}" for warning in warnings)}

硬性字数要求：
1. 地理位置：550-650 个汉字，目标 600 个汉字，2-3 个自然段。
2. 行政区划与社会经济概况：330-400 个汉字，目标 360 个汉字。
3. 气候条件：330-420 个汉字，目标 380 个汉字，只写地形地貌和气候事实，不写作业建议。
4. 历史沿革：1250-1400 个汉字，目标 1300 个汉字，8-11 段，按时间顺序写。
5. 四个字段合计 2000-2400 个汉字，目标 2200-2350 个汉字。

上一次输出如下，仅供你判断哪里过短，不要照抄短文本：
{json.dumps(previous, ensure_ascii=False, indent=2)}

请只输出 JSON 对象，键名必须严格为：
{json.dumps(REGION_OVERVIEW_KEYS, ensure_ascii=False)}
"""


def build_single_region_field_prompt(
    county: str,
    key: str,
    previous: str = "",
    warning: str = "",
    research: dict[str, str] | None = None,
) -> str:
    title, requirement, _min_chars = REGION_FIELD_SPECS[key]
    research_note = ""
    if research:
        research_note = f"""
可用资料卡如下。除“需剔除或慎用的信息”外，请尽量保留与“{title}”相关的事实，并按报告逻辑重组：
{json.dumps(research, ensure_ascii=False, indent=2)}
"""
    previous_note = ""
    if previous or warning:
        target_note = ""
        if key == "项目所在地旗县历史沿革":
            target_note = "本次必须写成 8-11 个自然段，至少 1300 个汉字，重点补足各历史阶段的行政隶属、军政建置、交通地位和族群活动。"
        elif key == "项目所在地旗县地理位置概况":
            target_note = "本次必须写成 2-3 个自然段，至少 650 个汉字，补充地貌格局、资源类型、交通区位和历史文化资源的叙述性文字。"
        elif key == "项目所在地旗县气候条件":
            target_note = "本次必须写成 1-2 个自然段，至少 420 个汉字，补充地形地貌对气候的影响、降水季节分布、风沙和蒸发等指标性叙述，不写作业建议。"
        previous_note = f"""
上一次该字段未达标：
- 问题：{warning}
- 上一次文本：{previous}

请在事实稳健的前提下重写并明显扩充，不要压缩。{target_note}
"""
    return f"""
你是考古调查勘探报告的“项目区域概况”撰写 agent。请只围绕“{county}”这个旗县级行政区撰写“{title}”字段，不要结合任何具体建设项目、矿山项目或红线范围情况。
{research_note}

写作要求：
{requirement}
不要写宣传口号、网页链接、来源说明、脚注或 Markdown；不要编造不可核验的具体文号、批复号、统计数值、荣誉称号或考古结论。
不得出现“本项目”“该项目”“项目区”“项目所在地”“项目用地”“项目红线”“建设内容”“矿区范围”“坐落于”等把文字指向具体项目的表述。
出于资料安全要求，本次请求只提供旗县名称，不提供项目名称、项目坐标、项目位置、建设内容、人工填写表参考内容或其他项目资料。请仅依据旗县级公开常识生成。
{previous_note}
只输出 JSON 对象，键名必须严格为：{json.dumps([key], ensure_ascii=False)}
"""


def extract_json_object(text: str, required_keys: list[str]) -> dict[str, str]:
    match = re.search(r"\{.*\}", text, flags=re.S)
    if not match:
        raise ValueError("agent 响应中未找到 JSON 对象。")
    raw = match.group(0)
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        data = loose_extract_json_object(raw, required_keys)
    missing = [key for key in required_keys if not clean(data.get(key))]
    if missing:
        raise ValueError(f"agent 响应缺少字段：{', '.join(missing)}")
    return {key: clean(data.get(key)) for key in required_keys}


def loose_extract_json_object(text: str, required_keys: list[str]) -> dict[str, str]:
    data: dict[str, str] = {}
    positions: list[tuple[int, str, int]] = []
    for key in required_keys:
        match = re.search(rf'["“”]?{re.escape(key)}["“”]?', text)
        if match:
            positions.append((match.start(), key, match.end()))
    positions.sort()
    for index, (_, key, key_end) in enumerate(positions):
        next_start = positions[index + 1][0] if index + 1 < len(positions) else len(text)
        chunk = text[key_end:next_start]
        chunk = re.sub(r"^[\s:：,，\"“”]+", "", chunk)
        chunk = re.sub(r"[\s,，}\"“”]+$", "", chunk)
        chunk = chunk.replace("\\n", "\n").replace('\\"', '"')
        data[key] = clean(chunk)
    return data


def extract_region_json(text: str, required_keys: list[str] | None = None) -> dict[str, str]:
    return extract_json_object(text, required_keys or REGION_OVERVIEW_KEYS)


def extract_research_json(text: str) -> dict[str, str]:
    return extract_json_object(text, REGION_RESEARCH_KEYS)


def count_chinese_chars(values: dict[str, str]) -> int:
    return sum(len(re.findall(r"[\u4e00-\u9fff]", value)) for value in values.values())


def count_chinese_text(value: str) -> int:
    return len(re.findall(r"[\u4e00-\u9fff]", value))


def validate_region_overview(values: dict[str, str], *, strict: bool = False) -> list[str]:
    warnings: list[str] = []
    total = count_chinese_chars(values)
    geo_chars = count_chinese_text(values.get("项目所在地旗县地理位置概况", ""))
    admin_chars = count_chinese_text(values.get("项目所在地旗县行政区划与社会经济概况", ""))
    climate_chars = count_chinese_text(values.get("项目所在地旗县气候条件", ""))
    history_chars = count_chinese_text(values.get("项目所在地旗县历史沿革", ""))
    if total < 2000 or total > 2400:
        warnings.append(f"四部分合计约 {total} 个汉字，应控制在 2000-2400 个汉字。")
    if geo_chars < 450:
        warnings.append(f"地理位置约 {geo_chars} 个汉字，应不少于 450 个汉字。")
    if admin_chars < 250:
        warnings.append(f"行政区划与社会经济概况约 {admin_chars} 个汉字，应不少于 250 个汉字。")
    if climate_chars < 250:
        warnings.append(f"气候条件约 {climate_chars} 个汉字，应不少于 250 个汉字。")
    if history_chars < 1000:
        warnings.append(f"历史沿革约 {history_chars} 个汉字，应不少于 1000 个汉字。")
    if strict and warnings:
        raise ValueError("；".join(warnings))
    return warnings


def region_candidate_score(values: dict[str, str], warnings: list[str]) -> int:
    score = sum(count_chinese_text(values.get(key, "")) for key in REGION_OVERVIEW_KEYS)
    return score - len(warnings) * 300


def ark_chat_completion(prompt: str) -> str:
    api_key = os.environ.get("ARK_API_KEY") or os.environ.get("DOUBAO_ARK_API_KEY")
    if not api_key:
        raise RuntimeError("未配置 ARK_API_KEY 或 DOUBAO_ARK_API_KEY。")
    model = os.environ.get("DOUBAO_ARK_MODEL") or os.environ.get("REGION_OVERVIEW_MODEL") or DOUBAO_ARK_MODEL
    base_url = os.environ.get("DOUBAO_ARK_BASE_URL", DOUBAO_ARK_BASE_URL).rstrip("/")
    if base_url.endswith("/chat/completions"):
        base_url = base_url[: -len("/chat/completions")]
    endpoint = base_url if base_url.endswith("/responses") else f"{base_url}/responses"
    payload = {
        "model": model,
        "input": [
            {
                "role": "system",
                "content": "你是严谨的中文考古调查勘探报告写作助手，只输出可解析 JSON。",
            },
            {"role": "user", "content": prompt},
        ],
        "text": {"format": {"type": "json_object"}},
    }
    request = urllib.request.Request(
        endpoint,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=180) as response:
            result = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"豆包 Ark 请求失败：HTTP {exc.code} {details[:500]}") from exc
    texts: list[str] = []
    for item in result.get("output", []):
        for content in item.get("content", []):
            if content.get("type") in {"output_text", "text"} and content.get("text"):
                texts.append(content["text"])
    if not texts and result.get("output_text"):
        texts.append(result["output_text"])
    if not texts:
        raise RuntimeError("豆包 Ark Responses 响应中没有正文 text。")
    return "\n".join(texts)


def deepseek_chat_completion(prompt: str) -> str:
    api_key = os.environ.get("DEEPSEEK_API_KEY")
    if not api_key:
        raise RuntimeError("未配置 DEEPSEEK_API_KEY。")
    model = os.environ.get("DEEPSEEK_MODEL") or os.environ.get("REGION_OVERVIEW_MODEL") or DEEPSEEK_MODEL
    base_url = os.environ.get("DEEPSEEK_BASE_URL", DEEPSEEK_BASE_URL).rstrip("/")
    if base_url.endswith("/chat/completions"):
        endpoint = base_url
    else:
        endpoint = f"{base_url}/chat/completions"
    payload = {
        "model": model,
        "messages": [
            {
                "role": "system",
                "content": "你是严谨的中文考古调查勘探报告写作助手，只输出可解析 JSON。",
            },
            {"role": "user", "content": prompt},
        ],
        "response_format": {"type": "json_object"},
        "temperature": 0.2,
        "max_tokens": 8192,
    }
    request = urllib.request.Request(
        endpoint,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=180) as response:
            result = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"DeepSeek 请求失败：HTTP {exc.code} {details[:500]}") from exc
    choices = result.get("choices") or []
    texts = [
        clean((choice.get("message") or {}).get("content"))
        for choice in choices
        if clean((choice.get("message") or {}).get("content"))
    ]
    if not texts:
        raise RuntimeError("DeepSeek 响应中没有 message.content。")
    return "\n".join(texts)


def generate_region_overview_fieldwise(
    county: str,
    completion: Any,
    research: dict[str, str] | None = None,
) -> dict[str, str]:
    generated: dict[str, str] = {}
    for key in REGION_OVERVIEW_KEYS:
        previous = ""
        warning = ""
        _, _, min_chars = REGION_FIELD_SPECS[key]
        for _ in range(4):
            data = extract_region_json(completion(build_single_region_field_prompt(county, key, previous, warning, research)), [key])
            value = clean(data.get(key))
            chars = count_chinese_text(value)
            if chars >= min_chars:
                generated[key] = value
                break
            previous = value
            warning = f"{key} 约 {chars} 个汉字，应不少于 {min_chars} 个汉字。"
        if key not in generated:
            generated[key] = previous
    return generated


def repair_short_region_fields(
    county: str,
    completion: Any,
    generated: dict[str, str],
    warnings: list[str],
    research: dict[str, str] | None = None,
) -> tuple[dict[str, str], list[str]]:
    current = dict(generated)
    current_warnings = list(warnings)
    for _ in range(4):
        changed = False
        for key in REGION_OVERVIEW_KEYS:
            field_warnings = [warning for warning in current_warnings if key.replace("项目所在地旗县", "").replace("概况", "")[:4] in warning or REGION_FIELD_SPECS[key][0] in warning]
            if not field_warnings:
                continue
            previous = current.get(key, "")
            data = extract_region_json(
                completion(build_single_region_field_prompt(county, key, previous, "；".join(field_warnings), research)),
                [key],
            )
            value = clean(data.get(key))
            if count_chinese_text(value) > count_chinese_text(previous):
                current[key] = value
                changed = True
        current_warnings = validate_region_overview(current, strict=False)
        if not current_warnings or not changed:
            break
    return current, current_warnings


def openai_responses_completion(prompt: str) -> str:
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("未配置 OPENAI_API_KEY。")
    model = os.environ.get("REGION_OVERVIEW_OPENAI_MODEL") or os.environ.get("OPENAI_MODEL") or OPENAI_RESPONSES_MODEL
    payload = {
        "model": model,
        "input": [
            {
                "role": "system",
                "content": "你是严谨的中文考古调查勘探报告写作助手，只输出可解析 JSON。",
            },
            {"role": "user", "content": prompt},
        ],
        "text": {"format": {"type": "json_object"}},
    }
    request = urllib.request.Request(
        "https://api.openai.com/v1/responses",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=180) as response:
            result = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"OpenAI Responses 请求失败：HTTP {exc.code} {details[:500]}") from exc
    texts: list[str] = []
    for item in result.get("output", []):
        for content in item.get("content", []):
            if content.get("type") in {"output_text", "text"} and content.get("text"):
                texts.append(content["text"])
    if not texts and result.get("output_text"):
        texts.append(result["output_text"])
    if not texts:
        raise RuntimeError("OpenAI Responses 响应中没有正文 text。")
    return "\n".join(texts)


def generate_region_overview(
    fields: dict[str, Any],
    *,
    dry_run_prompt: bool = False,
    strict: bool = False,
) -> dict[str, Any]:
    load_project_env()
    county = infer_county(fields)
    if not county:
        raise ValueError("缺少项目所在地旗县，且无法从项目位置中识别旗县。")
    prompt = build_region_overview_prompt(county)
    if dry_run_prompt:
        return {
            "ok": True,
            "provider": "deepseek" if os.environ.get("DEEPSEEK_API_KEY") else "doubao-ark",
            "model": os.environ.get("DEEPSEEK_MODEL")
            or os.environ.get("DOUBAO_ARK_MODEL")
            or os.environ.get("REGION_OVERVIEW_MODEL")
            or DEEPSEEK_MODEL,
            "county": county,
            "prompt": prompt,
            "research_prompt": build_region_research_prompt(county),
        }
    attempts: list[tuple[str, str, Any]] = []
    if os.environ.get("DEEPSEEK_API_KEY"):
        attempts.append(("deepseek", os.environ.get("DEEPSEEK_MODEL") or os.environ.get("REGION_OVERVIEW_MODEL") or DEEPSEEK_MODEL, deepseek_chat_completion))
    attempts.append(("doubao-ark", os.environ.get("DOUBAO_ARK_MODEL") or os.environ.get("REGION_OVERVIEW_MODEL") or DOUBAO_ARK_MODEL, ark_chat_completion))
    attempts.append(("openai-responses", os.environ.get("REGION_OVERVIEW_OPENAI_MODEL") or os.environ.get("OPENAI_MODEL") or OPENAI_RESPONSES_MODEL, openai_responses_completion))

    errors: list[str] = []
    generated: dict[str, str] | None = None
    research: dict[str, str] = {}
    provider = ""
    model = ""
    warnings: list[str] = []
    best_generated: dict[str, str] | None = None
    best_research: dict[str, str] = {}
    best_provider = ""
    best_model = ""
    best_warnings: list[str] = []
    best_score: int | None = None
    for candidate_provider, candidate_model, completion in attempts:
        try:
            research = extract_research_json(completion(build_region_research_prompt(county)))
            research = apply_known_region_facts(county, research)
        except Exception as research_exc:
            errors.append(f"{candidate_provider} 资料卡生成失败：{research_exc}")
            research = apply_known_region_facts(county, {})
        candidate_prompt = build_region_overview_prompt(county, research)
        try:
            for _ in range(4):
                generated = extract_region_json(completion(candidate_prompt))
                generated = enforce_known_region_facts(county, generated)
                warnings = validate_region_overview(generated, strict=False)
                if warnings:
                    generated, warnings = repair_short_region_fields(county, completion, generated, warnings, research)
                    generated = enforce_known_region_facts(county, generated)
                    warnings = validate_region_overview(generated, strict=False)
                score = region_candidate_score(generated, warnings)
                if best_score is None or score > best_score:
                    best_generated = dict(generated)
                    best_research = dict(research)
                    best_provider = candidate_provider
                    best_model = candidate_model
                    best_warnings = list(warnings)
                    best_score = score
                if not warnings:
                    provider = candidate_provider
                    model = candidate_model
                    break
                candidate_prompt = build_region_overview_repair_prompt(county, generated, warnings, research)
            if not provider:
                generated = generate_region_overview_fieldwise(county, completion, research)
                generated = enforce_known_region_facts(county, generated)
                warnings = validate_region_overview(generated, strict=False)
                if warnings:
                    generated, warnings = repair_short_region_fields(county, completion, generated, warnings, research)
                    generated = enforce_known_region_facts(county, generated)
                    warnings = validate_region_overview(generated, strict=False)
                score = region_candidate_score(generated, warnings)
                if best_score is None or score > best_score:
                    best_generated = dict(generated)
                    best_research = dict(research)
                    best_provider = candidate_provider
                    best_model = candidate_model
                    best_warnings = list(warnings)
                    best_score = score
                if not warnings:
                    provider = candidate_provider
                    model = candidate_model
            if provider:
                break
            raise ValueError("；".join(warnings))
        except Exception as exc:
            errors.append(f"{candidate_provider} 失败：{exc}")
            generated = None
            warnings = []
    if generated is None or not provider:
        if best_generated and not strict:
            generated = best_generated
            research = best_research
            provider = best_provider
            model = best_model
            warnings = best_warnings
        else:
            raise RuntimeError("；".join(errors))
    if generated is None or not provider:
        raise RuntimeError("；".join(errors))
    warnings = validate_region_overview(generated, strict=strict)
    return {
        "ok": True,
        "provider": provider,
        "model": model,
        "county": county,
        "fields": generated,
        "research": research,
        "warnings": warnings,
    }


def mcp_tool_schema() -> dict[str, Any]:
    return {
        "name": "generate_region_overview",
        "description": "生成考古调查勘探报告/计划中项目区域概况的地理位置、行政区划与社会经济概况、气候条件和历史沿革。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "项目所在地旗县": {"type": "string"},
                "项目位置": {"type": "string"},
                "项目所在地旗县地理位置概况": {"type": "string"},
                "项目所在地旗县行政区划与社会经济概况": {"type": "string"},
                "项目所在地旗县气候条件": {"type": "string"},
                "项目所在地旗县历史沿革": {"type": "string"},
                "strict": {"type": "boolean", "default": False},
            },
            "required": [],
        },
        "outputSchema": {
            "type": "object",
            "properties": {
                "ok": {"type": "boolean"},
                "provider": {"type": "string"},
                "model": {"type": "string"},
                "county": {"type": "string"},
                "fields": {
                    "type": "object",
                    "properties": {key: {"type": "string"} for key in REGION_OVERVIEW_KEYS},
                    "required": REGION_OVERVIEW_KEYS,
                },
                "research": {
                    "type": "object",
                    "properties": {key: {"type": "string"} for key in REGION_RESEARCH_KEYS},
                },
                "warnings": {"type": "array", "items": {"type": "string"}},
            },
            "required": ["ok", "provider", "model", "county", "fields"],
        },
    }


def read_json_arg(path: str) -> dict[str, Any]:
    if path == "-":
        return json.loads(sys.stdin.read())
    return json.loads(Path(path).read_text(encoding="utf-8"))


def main() -> int:
    parser = argparse.ArgumentParser(description="豆包 Ark 项目区域概况 agent")
    parser.add_argument("--input-json", default="-", help="JSON 输入文件；使用 - 从 stdin 读取")
    parser.add_argument("--output-json", default="-", help="JSON 输出文件；使用 - 输出到 stdout")
    parser.add_argument("--print-schema", action="store_true", help="打印预留 MCP 工具 schema")
    parser.add_argument("--dry-run-prompt", action="store_true", help="只生成提示词，不调用外部 API")
    parser.add_argument("--strict", action="store_true", help="字数不满足要求时返回错误")
    args = parser.parse_args()

    if args.print_schema:
        result = mcp_tool_schema()
    else:
        request = read_json_arg(args.input_json)
        fields = request.get("fields", request)
        strict = bool(request.get("strict", args.strict))
        result = generate_region_overview(fields, dry_run_prompt=args.dry_run_prompt, strict=strict)

    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output_json == "-":
        print(text)
    else:
        Path(args.output_json).write_text(text + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
