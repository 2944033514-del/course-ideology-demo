import html
import io
import math
import re
from difflib import SequenceMatcher
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd
import streamlit as st


APP_TITLE = "来华留学生课程思政数据要素自动识别演示系统"
EXPORT_NAME = "识别结果导出.xlsx"
APP_DIR = Path(__file__).parent
MANUAL_RECORD_NAME = "manual_records.xlsx"
CODING_FRAMEWORK_NAME = "coding_framework.xlsx"

RESULT_COLUMNS = [
    "文本编号",
    "课次",
    "原文句段",
    "识别来源",
    "是否含课程思政要素",
    "机器/人工主标签",
    "机器/人工辅标签",
    "证据词",
    "置信度",
    "标注理由",
    "可用于教学切入点",
]

CATEGORY_INFO = {
    "C1": {"label": "在华学习的中国社会与发展理解", "color": "#2F80ED"},
    "C2": {"label": "文化理解与跨文化沟通", "color": "#27AE60"},
    "C3": {"label": "法治规范与学术伦理", "color": "#F2994A"},
    "C4": {"label": "专业责任与职业素养", "color": "#9B51E0"},
    "C5": {"label": "全球胜任与友好交流", "color": "#00A6B4"},
    "C6": {"label": "学习适应与发展支持", "color": "#607D9A"},
}

FALLBACK_LEXICON = {
    "C1": {
        "strong": [
            "当代中国",
            "中国社会",
            "社会现象",
            "社会发展",
            "发展实际",
            "中国发展",
            "现实国情",
            "城市化进程",
            "地区差异",
            "交通",
            "就业",
            "生活方式",
            "中国国情",
            "中国式现代化",
            "改革开放",
            "国家治理",
            "发展实践",
            "中国案例",
            "中国方案",
        ],
        "weak": ["变化", "发展", "不同地区", "人口", "城市", "社会", "经济", "生活", "中国经验", "制度环境", "基层治理", "中国概况"],
        "required": ["中国", "当代中国", "留学生", "认识", "了解", "发展实际", "社会现象", "课程", "案例", "学习", "分析", "理解", "比较", "实践", "教学"],
        "exclude": ["中国学生", "中文名"],
    },
    "C2": {
        "strong": ["中华文化", "传统文化", "文明互鉴", "跨文化交流", "文化自信", "国际传播"],
        "weak": ["文化传播", "跨文化阐释", "中国故事", "文化比较", "文化差异", "文化理解"],
        "required": ["表达", "理解", "沟通", "比较", "阐释", "交流", "语言", "受众"],
        "exclude": ["经济增长", "制度治理"],
    },
    "C3": {
        "strong": ["遵守中国法律", "学术诚信", "科研伦理", "知识产权", "数据安全", "考试纪律"],
        "weak": ["引用规范", "查重", "实验伦理", "隐私保护", "校纪校规", "法律法规"],
        "required": ["规范", "遵守", "纪律", "诚信", "伦理", "法规", "保护", "安全"],
        "exclude": ["请勿迟到", "禁止饮食"],
    },
    "C4": {
        "strong": ["职业素养", "专业责任", "行业规范", "工程伦理", "医德医风", "职业伦理"],
        "weak": ["社会责任", "岗位责任", "职业规范", "质量意识", "服务意识", "实践能力"],
        "required": ["专业", "职业", "岗位", "行业", "实践", "责任", "伦理", "服务"],
        "exclude": ["普通课堂要求"],
    },
    "C5": {
        "strong": ["全球胜任力", "国际合作", "友好交流", "全球治理", "人类命运共同体", "可持续发展"],
        "weak": ["国际理解", "世界各国", "共同应对", "国际视野", "民心相通", "交流互鉴"],
        "required": ["全球", "国际", "合作", "交流", "共同", "世界", "友好", "跨国"],
        "exclude": ["仅介绍外国地名"],
    },
    "C6": {
        "strong": ["学习适应", "发展支持", "心理支持", "学业支持", "生活适应", "辅导服务"],
        "weak": ["适应能力", "不太习惯", "学习困难", "语言支持", "校园服务", "帮助学生"],
        "required": ["适应", "支持", "帮助", "辅导", "学习", "生活", "心理", "发展"],
        "exclude": ["旅游偏好"],
    },
}


def split_terms(value):
    if pd.isna(value):
        return []
    parts = re.split(r"[;；,，、\n/|]+", str(value))
    return [p.strip() for p in parts if p and p.strip() and p.strip().lower() != "nan"]


def find_col(columns, candidates):
    normalized = {str(c).replace(" ", "").lower(): c for c in columns}
    for key in candidates:
        key_norm = key.replace(" ", "").lower()
        for norm, original in normalized.items():
            if key_norm in norm or norm in key_norm:
                return original
    return None


def locate_excel():
    path = APP_DIR / CODING_FRAMEWORK_NAME
    return path if path.exists() else None


def locate_manual_record_excel():
    path = APP_DIR / MANUAL_RECORD_NAME
    return path if path.exists() else None


def normalize_text(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", "", str(value).strip())


def clean_cell(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def extract_category_id(label):
    match = re.search(r"\b(C[1-6])\b", str(label))
    return match.group(1) if match else ""


@st.cache_data(show_spinner=False)
def load_category_names(excel_path):
    names = {cid: info["label"] for cid, info in CATEGORY_INFO.items()}
    if not excel_path or not Path(excel_path).exists():
        return names

    try:
        df = pd.read_excel(excel_path, sheet_name="类目体系")
    except Exception:
        return names

    id_col = find_col(df.columns, ["类目ID", "一级类目ID", "类别ID", "编码"])
    name_col = find_col(df.columns, ["一级类目", "类目名称", "类别名称", "标签名称"])
    if id_col and name_col:
        for _, row in df.iterrows():
            cid = str(row.get(id_col, "")).strip()
            label = str(row.get(name_col, "")).strip()
            if cid in names and label and label.lower() != "nan":
                names[cid] = label
    return names


@st.cache_data(show_spinner=False)
def load_manual_records(record_path):
    if not record_path or not Path(record_path).exists():
        return pd.DataFrame()
    try:
        df = pd.read_excel(record_path, sheet_name="教材语料备案表")
    except Exception:
        return pd.DataFrame()

    required_cols = [
        "文本编号",
        "课次",
        "课文主题",
        "段落编号",
        "句段序号",
        "原文句段",
        "是否含课程思政要素",
        "人工主标签",
        "人工辅标签",
        "证据词",
        "标注理由",
        "可用于教学切入点",
    ]
    for col in required_cols:
        if col not in df.columns:
            df[col] = ""
    df = df[required_cols].copy()
    df["原文句段_norm"] = df["原文句段"].apply(normalize_text)
    return df


def match_manual_records(text, manual_df, threshold=0.82):
    input_norm = normalize_text(text)
    if manual_df.empty or not input_norm:
        return pd.DataFrame(columns=RESULT_COLUMNS), []

    matched = []
    for _, row in manual_df.iterrows():
        segment_norm = row["原文句段_norm"]
        if len(segment_norm) < 6:
            continue
        exact_or_contains = (
            input_norm == segment_norm
            or segment_norm in input_norm
            or input_norm in segment_norm
        )
        similarity = SequenceMatcher(None, input_norm, segment_norm).ratio()
        if exact_or_contains or similarity >= threshold:
            matched.append((row, similarity))

    if not matched:
        return pd.DataFrame(columns=RESULT_COLUMNS), []

    # 包含匹配可能一次命中多个备案句段；纯相似匹配则保留最接近的一条，避免误报过多。
    contains_hits = [
        item
        for item in matched
        if item[0]["原文句段_norm"] in input_norm or input_norm in item[0]["原文句段_norm"]
    ]
    selected = contains_hits if contains_hits else [max(matched, key=lambda item: item[1])]

    rows = []
    spans = []
    for row, similarity in selected:
        has_element = clean_cell(row["是否含课程思政要素"])
        is_positive = has_element == "是"
        main_label = clean_cell(row["人工主标签"]) if is_positive else "无明显课程思政要素"
        aux_label = clean_cell(row["人工辅标签"]) if is_positive else ""
        reason = clean_cell(row["标注理由"])
        if not is_positive:
            reason = "该句段已备案，但人工标注为无明显课程思政要素。"

        rows.append(
            {
                "文本编号": clean_cell(row["文本编号"]),
                "课次": clean_cell(row["课次"]),
                "原文句段": clean_cell(row["原文句段"]),
                "识别来源": "教材备案表匹配",
                "是否含课程思政要素": has_element or "否",
                "机器/人工主标签": main_label,
                "机器/人工辅标签": aux_label,
                "证据词": clean_cell(row["证据词"]) if is_positive else "",
                "置信度": 0.95,
                "标注理由": reason,
                "可用于教学切入点": clean_cell(row["可用于教学切入点"]) if is_positive else "",
            }
        )

        cid = extract_category_id(main_label)
        if is_positive and cid:
            spans.append(
                {
                    "text": clean_cell(row["原文句段"]),
                    "cid": cid,
                    "words": split_terms(row["证据词"]),
                    "source": "教材备案表匹配",
                }
            )

    return pd.DataFrame(rows, columns=RESULT_COLUMNS), spans


@st.cache_data(show_spinner=False)
def load_lexicon(excel_path):
    lexicon = {
        cid: {
            "strong": list(data["strong"]),
            "weak": list(data["weak"]),
            "required": list(data["required"]),
            "exclude": list(data["exclude"]),
        }
        for cid, data in FALLBACK_LEXICON.items()
    }

    if not excel_path or not Path(excel_path).exists():
        return lexicon

    try:
        df = pd.read_excel(excel_path, sheet_name="辅助词库")
    except Exception:
        return lexicon

    id_col = find_col(df.columns, ["类目ID", "一级类目ID", "适用类目", "类别ID", "编码"])
    strong_col = find_col(df.columns, ["强触发词", "机器识别辅助词库", "辅助词库", "触发词", "关键词"])
    weak_col = find_col(df.columns, ["弱触发词", "同义词", "扩展表达", "近义词", "同义词/扩展表达"])
    required_col = find_col(df.columns, ["必要语境", "共现语境", "课程语境", "使用条件", "触发条件"])
    exclude_col = find_col(df.columns, ["排除语境", "排除条件", "反例语境", "使用限制"])

    if not id_col:
        return lexicon

    for _, row in df.iterrows():
        cid = str(row.get(id_col, "")).strip()
        if cid not in lexicon:
            continue
        if strong_col:
            lexicon[cid]["strong"].extend(split_terms(row.get(strong_col)))
        if weak_col:
            lexicon[cid]["weak"].extend(split_terms(row.get(weak_col)))
        if required_col:
            lexicon[cid]["required"].extend(split_terms(row.get(required_col)))
        if exclude_col and "使用限制" not in str(exclude_col):
            lexicon[cid]["exclude"].extend(split_terms(row.get(exclude_col)))

    for cid in lexicon:
        for field in lexicon[cid]:
            lexicon[cid][field] = sorted(set(lexicon[cid][field]), key=len, reverse=True)
    return lexicon


@st.cache_data(show_spinner=False)
def load_manual_lexicon(record_path):
    addition = {cid: {"strong": [], "weak": [], "required": [], "exclude": []} for cid in CATEGORY_INFO}
    if not record_path or not Path(record_path).exists():
        return addition
    try:
        df = pd.read_excel(record_path, sheet_name="可补充辅助词库")
    except Exception:
        return addition

    id_col = find_col(df.columns, ["类目ID", "一级类目ID", "类别ID", "编码"])
    strong_col = find_col(df.columns, ["强触发词", "强触发", "核心词"])
    weak_col = find_col(df.columns, ["弱触发词", "弱触发", "扩展词"])
    required_col = find_col(df.columns, ["必要语境", "共现语境", "课程语境"])
    exclude_col = find_col(df.columns, ["排除语境", "排除条件"])

    if not id_col:
        return addition

    for _, row in df.iterrows():
        cid = clean_cell(row.get(id_col))
        if cid not in addition:
            continue
        if strong_col:
            addition[cid]["strong"].extend(split_terms(row.get(strong_col)))
        if weak_col:
            addition[cid]["weak"].extend(split_terms(row.get(weak_col)))
        if required_col:
            addition[cid]["required"].extend(split_terms(row.get(required_col)))
        if exclude_col:
            addition[cid]["exclude"].extend(split_terms(row.get(exclude_col)))
    return addition


def merge_lexicon(base, addition):
    merged = {
        cid: {field: list(values) for field, values in fields.items()}
        for cid, fields in base.items()
    }
    for cid, fields in addition.items():
        if cid not in merged:
            continue
        for field, values in fields.items():
            merged[cid][field].extend(values)
    for cid in merged:
        for field in merged[cid]:
            merged[cid][field] = sorted(set(merged[cid][field]), key=len, reverse=True)
    return merged


def split_segments(text):
    text = re.sub(r"\s+", " ", text.strip())
    if not text:
        return []
    pieces = re.split(r"(?<=[。！？；;!?])\s*", text)
    return [p.strip() for p in pieces if p.strip()]


def score_segment(segment, lexicon):
    scores = defaultdict(float)
    evidence = defaultdict(list)
    debug = {}
    lowered = segment.lower()

    for cid, vocab in lexicon.items():
        strong_hits = [w for w in vocab["strong"] if w and w.lower() in lowered]
        weak_hits = [w for w in vocab["weak"] if w and w.lower() in lowered]
        required_hits = [w for w in vocab["required"] if w and w.lower() in lowered]
        exclude_hits = [w for w in vocab["exclude"] if w and w.lower() in lowered]

        if strong_hits:
            scores[cid] += 0.55 + min(0.25, 0.08 * len(strong_hits))
            evidence[cid].extend(strong_hits)

        if weak_hits and required_hits:
            scores[cid] += 0.35 + min(0.2, 0.06 * len(weak_hits))
            evidence[cid].extend(weak_hits + required_hits[:2])

        if exclude_hits:
            scores[cid] -= 0.35

        debug[cid] = {
            "命中的强触发词": "；".join(strong_hits),
            "命中的弱触发词": "；".join(weak_hits),
            "命中的必要语境": "；".join(required_hits),
            "得分": round(max(0.0, scores[cid]), 3),
        }

    return scores, evidence, debug


def confidence(score, best_score, cid):
    if score <= 0:
        return 0.0
    raw = 1 / (1 + math.exp(-2.8 * (score - 0.45)))
    if best_score and score < best_score:
        raw *= 0.86
    if cid == "C1" and score >= 0.55:
        raw = max(0.78, min(0.86, raw))
    return round(max(0.0, min(0.98, raw)), 2)


def recognize(text, lexicon, category_names):
    rows = []
    spans = []
    debug_rows = []

    for segment in split_segments(text):
        scores, evidence, debug = score_segment(segment, lexicon)
        for cid in CATEGORY_INFO:
            debug_rows.append(
                {
                    "原文片段": segment,
                    "类目": f"{cid} {category_names.get(cid, CATEGORY_INFO[cid]['label'])}",
                    "命中的强触发词": debug[cid]["命中的强触发词"],
                    "命中的弱触发词": debug[cid]["命中的弱触发词"],
                    "命中的必要语境": debug[cid]["命中的必要语境"],
                    "得分": debug[cid]["得分"],
                }
            )
        positive = [(cid, score) for cid, score in scores.items() if score >= 0.28 and evidence[cid]]
        if not positive:
            continue

        positive.sort(key=lambda item: item[1], reverse=True)
        main_id, best_score = positive[0]
        aux_id = positive[1][0] if len(positive) > 1 and positive[1][1] >= 0.28 else ""
        main_label = f"{main_id} {category_names.get(main_id, CATEGORY_INFO[main_id]['label'])}"
        aux_label = f"{aux_id} {category_names.get(aux_id, CATEGORY_INFO[aux_id]['label'])}" if aux_id else ""
        hit_words = sorted(set(evidence[main_id] + (evidence[aux_id] if aux_id else [])), key=len, reverse=True)

        rows.append(
            {
                "原文片段": segment,
                "命中证据词": "；".join(hit_words),
                "机器主标签": main_label,
                "机器辅标签": aux_label,
                "置信度": confidence(best_score, best_score, main_id),
            }
        )
        spans.append({"text": segment, "cid": main_id, "words": hit_words})

    return (
        pd.DataFrame(rows, columns=["原文片段", "命中证据词", "机器主标签", "机器辅标签", "置信度"]),
        spans,
        pd.DataFrame(debug_rows, columns=["原文片段", "类目", "命中的强触发词", "命中的弱触发词", "命中的必要语境", "得分"]),
    )


def convert_rule_results(rule_df):
    rows = []
    for _, row in rule_df.iterrows():
        evidence = clean_cell(row.get("命中证据词"))
        rows.append(
            {
                "文本编号": "",
                "课次": "",
                "原文句段": clean_cell(row.get("原文片段")),
                "识别来源": "规则词库识别",
                "是否含课程思政要素": "是",
                "机器/人工主标签": clean_cell(row.get("机器主标签")),
                "机器/人工辅标签": clean_cell(row.get("机器辅标签")),
                "证据词": evidence,
                "置信度": row.get("置信度", 0),
                "标注理由": f"规则词库根据“{evidence}”等证据词自动识别。" if evidence else "规则词库自动识别。",
                "可用于教学切入点": "",
            }
        )
    return pd.DataFrame(rows, columns=RESULT_COLUMNS)


def make_unrecognized_result(text):
    return pd.DataFrame(
        [
            {
                "文本编号": "",
                "课次": "",
                "原文句段": text.strip(),
                "识别来源": "未识别",
                "是否含课程思政要素": "否",
                "机器/人工主标签": "无明显课程思政要素",
                "机器/人工辅标签": "",
                "证据词": "",
                "置信度": 0.0,
                "标注理由": "未匹配到教材备案表句段，规则词库也未识别到明确课程思政要素。",
                "可用于教学切入点": "",
            }
        ],
        columns=RESULT_COLUMNS,
    )


def recognize_with_manual_priority(text, manual_df, lexicon, category_names):
    manual_result_df, manual_spans = match_manual_records(text, manual_df)
    if not manual_result_df.empty:
        debug_df = pd.DataFrame(columns=["原文片段", "类目", "命中的强触发词", "命中的弱触发词", "命中的必要语境", "得分"])
        return manual_result_df, manual_spans, debug_df

    rule_df, spans, debug_df = recognize(text, lexicon, category_names)
    if not rule_df.empty:
        return convert_rule_results(rule_df), spans, debug_df

    return make_unrecognized_result(text), [], debug_df


def render_highlight(text, spans):
    escaped = html.escape(text)
    ordered_spans = sorted(spans, key=lambda s: len(s["text"]), reverse=True)

    for span in ordered_spans:
        cid = span["cid"]
        color = CATEGORY_INFO[cid]["color"]
        label = cid
        escaped_segment = html.escape(span["text"])
        replacement = (
            f'<mark class="tag tag-{cid}" title="{label}">'
            f'<span class="tag-label">{label}</span>{escaped_segment}</mark>'
        )
        escaped = escaped.replace(escaped_segment, replacement, 1)
    return escaped.replace("\n", "<br>")


def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="识别结果")
    return output.getvalue()


def demo_text():
    return (
        "在论文写作课中，教师要求留学生学习引用规范、查重要求和学术诚信原则，避免侵犯知识产权。"
        "课堂还结合中国改革开放以来的城市发展案例，引导学生理解中国社会与国家治理实践。"
        "在跨文化交流任务中，学生需要用中文介绍中华文化，并比较不同国家的时间观念。"
        "学校为初到中国的学生提供学业支持、心理支持和生活适应辅导，帮助他们更好完成专业学习。"
    )


st.set_page_config(page_title=APP_TITLE, page_icon="📘", layout="wide")

st.markdown(
    """
    <style>
    .main .block-container {padding-top: 1.6rem; max-width: 1280px;}
    h1 {font-size: 2rem;}
    .highlight-box {
        min-height: 360px;
        padding: 18px;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
        background: #FBFCFE;
        line-height: 2.1;
        font-size: 16px;
    }
    mark.tag {
        color: #111827;
        padding: 4px 6px;
        border-radius: 6px;
        margin: 0 2px;
    }
    .tag-label {
        color: white;
        border-radius: 4px;
        padding: 1px 5px;
        margin-right: 5px;
        font-size: 12px;
        font-weight: 700;
    }
    .tag-C1 {background: rgba(47,128,237,.16); border-bottom: 2px solid #2F80ED;}
    .tag-C2 {background: rgba(39,174,96,.16); border-bottom: 2px solid #27AE60;}
    .tag-C3 {background: rgba(242,153,74,.20); border-bottom: 2px solid #F2994A;}
    .tag-C4 {background: rgba(155,81,224,.16); border-bottom: 2px solid #9B51E0;}
    .tag-C5 {background: rgba(0,166,180,.16); border-bottom: 2px solid #00A6B4;}
    .tag-C6 {background: rgba(96,125,154,.18); border-bottom: 2px solid #607D9A;}
    .tag-C1 .tag-label {background:#2F80ED;}
    .tag-C2 .tag-label {background:#27AE60;}
    .tag-C3 .tag-label {background:#F2994A;}
    .tag-C4 .tag-label {background:#9B51E0;}
    .tag-C5 .tag-label {background:#00A6B4;}
    .tag-C6 .tag-label {background:#607D9A;}
    .legend {display:flex; flex-wrap:wrap; gap:8px 14px; margin:.5rem 0 1rem;}
    .legend-item {display:flex; align-items:center; gap:6px; font-size:13px;}
    .dot {width:12px; height:12px; border-radius:50%;}
    </style>
    """,
    unsafe_allow_html=True,
)

excel_path = locate_excel()
manual_record_path = locate_manual_record_excel()
category_names = load_category_names(str(excel_path) if excel_path else "")
manual_df = load_manual_records(str(manual_record_path) if manual_record_path else "")
lexicon = merge_lexicon(
    load_lexicon(str(excel_path) if excel_path else ""),
    load_manual_lexicon(str(manual_record_path) if manual_record_path else ""),
)

st.title(APP_TITLE)
if manual_record_path:
    st.caption(f"已优先读取教材语料备案表：{manual_record_path}")
else:
    st.warning(f"未找到 {MANUAL_RECORD_NAME}，系统将直接使用规则词库识别。")

if excel_path:
    st.caption(f"已读取辅助词库：{excel_path}")
else:
    st.warning("未找到标注体系模板 Excel，系统已启用内置演示词库。")

legend_html = '<div class="legend">'
for cid, info in CATEGORY_INFO.items():
    legend_html += (
        f'<div class="legend-item"><span class="dot" style="background:{info["color"]}"></span>'
        f'{cid} {html.escape(category_names.get(cid, info["label"]))}</div>'
    )
legend_html += "</div>"
st.markdown(legend_html, unsafe_allow_html=True)

left, right = st.columns([0.95, 1.05], gap="large")

with left:
    text = st.text_area(
        "输入教材课文或论文语料文本",
        value=demo_text(),
        height=360,
        placeholder="请粘贴来华留学生教材课文、课程大纲或论文语料文本……",
    )
    col_a, col_b = st.columns([1, 1])
    with col_a:
        run = st.button("开始识别", type="primary", use_container_width=True)
    with col_b:
        clear = st.button("清空结果", use_container_width=True)

if clear:
    st.session_state.pop("result_df", None)
    st.session_state.pop("spans", None)
    st.session_state.pop("debug_df", None)
    st.session_state.pop("has_run", None)

if run:
    result_df, spans, debug_df = recognize_with_manual_priority(text, manual_df, lexicon, category_names)
    st.session_state["result_df"] = result_df
    st.session_state["spans"] = spans
    st.session_state["debug_df"] = debug_df
    st.session_state["has_run"] = True

result_df = st.session_state.get("result_df", pd.DataFrame(columns=RESULT_COLUMNS))
spans = st.session_state.get("spans", [])
debug_df = st.session_state.get("debug_df", pd.DataFrame(columns=["原文片段", "类目", "命中的强触发词", "命中的弱触发词", "命中的必要语境", "得分"]))
has_run = st.session_state.get("has_run", False)
has_manual_no_element = (
    has_run
    and not result_df.empty
    and (result_df["识别来源"] == "教材备案表匹配").any()
    and (result_df["是否含课程思政要素"] == "否").all()
)
has_unrecognized = has_run and not result_df.empty and (result_df["识别来源"] == "未识别").all()

with right:
    st.subheader("识别高亮")
    if spans:
        st.markdown(f'<div class="highlight-box">{render_highlight(text, spans)}</div>', unsafe_allow_html=True)
    elif has_manual_no_element:
        st.markdown('<div class="highlight-box">已备案，但无明显课程思政要素</div>', unsafe_allow_html=True)
    elif has_unrecognized:
        st.markdown('<div class="highlight-box">未识别到匹配结果</div>', unsafe_allow_html=True)
    elif has_run:
        st.markdown('<div class="highlight-box">未识别到匹配结果</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="highlight-box">点击“开始识别”后，这里将显示带颜色标注的课程思政元素。</div>', unsafe_allow_html=True)

st.subheader("识别结果表")
if not has_run:
    st.info("暂无识别结果。")
else:
    if has_manual_no_element:
        st.warning("该句段已备案，但人工标注为无明显课程思政要素。")
    elif has_unrecognized:
        st.warning("未识别到匹配结果。")
    st.dataframe(result_df, use_container_width=True, hide_index=True)
    excel_bytes = to_excel_bytes(result_df)
    (APP_DIR / EXPORT_NAME).write_bytes(excel_bytes)
    st.download_button(
        "导出 Excel",
        data=excel_bytes,
        file_name=EXPORT_NAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.subheader("调试区")
if not has_run:
    st.info("点击“开始识别”后显示命中词和各类目得分。")
elif not result_df.empty and (result_df["识别来源"] == "教材备案表匹配").any():
    st.info("本次结果来自教材备案表匹配，已优先采用人工标注基准。")
elif debug_df.empty:
    st.info("暂无规则词库调试信息。")
else:
    st.dataframe(debug_df, use_container_width=True, hide_index=True)

st.subheader("各类目识别数量")
if not has_run or result_df.empty:
    chart_df = pd.DataFrame({"类目": list(CATEGORY_INFO), "识别数量": [0] * len(CATEGORY_INFO)})
else:
    positive = result_df[result_df["是否含课程思政要素"] == "是"]
    counts = Counter(positive["机器/人工主标签"].str.extract(r"^(C\d)", expand=False).dropna())
    chart_df = pd.DataFrame({"类目": list(CATEGORY_INFO), "识别数量": [counts.get(cid, 0) for cid in CATEGORY_INFO]})

st.bar_chart(chart_df.set_index("类目"), height=260)
