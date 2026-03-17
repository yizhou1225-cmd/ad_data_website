import streamlit as st
import pandas as pd
import numpy as np
import io
import json
import xlsxwriter
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
import docx.opc.constants
import time

# ==========================================
# PART 1: 配置区域 (修复了字段映射)
# ==========================================

COMMON_METRICS = {
    "spend": ["花费金额(USD)", "花费金额 （USD）", "花费金额 (USD)", "花费金额", "Amount Spent"],
    "roas": ["广告花费回报 (ROAS) - 购物", "广告花费回报（ROAS）-购物", "ROAS", "Purchase ROAS"],
    "purchases": ["购买次数", "成效数量", "成效", "Purchases"],
    "cpa": ["单次购买费用", "单次购物成本", "单次成效成本", "单次成效费用", "Cost per Purchase"],
    "ctr": ["链接点击率", "链接点击率（%)", "链接点击率（%）", "CTR"],
    "cpm": ["千次展示费用", "CPM"],
    "clicks": ["点击", "链接点击", "Clicks"],
    "impressions": ["曝光", "展示次数", "Impressions"],
    "purchase_value": ["购买价值", "购物价值", "Purchase Value"],
    "aov": ["单次购买价值", "单次购物价值"]
}

# 框定「每一个 Sheet」需要抽取哪些指标
SHEET_MAPPINGS = {
    "整体数据": {
        **COMMON_METRICS,
        "date_range": ["时间范围"],
        "clicks_all": ["点击"],
        "landing_page_views": ["落地页浏览量"],
        "add_to_cart": ["加入购物车"],
        "initiate_checkout": ["结账发起次数"],
        "rate_click_to_lp": ["点击-落地页浏览转化率"],
        "rate_lp_to_atc": ["落地页浏览-加购转化率"],
        "rate_atc_to_ic": ["加购-结账转化率"],
        "rate_ic_to_pur": ["结账-购买转化率"]
    },
    "分时段数据": {
        **COMMON_METRICS,
        "date_range": ["时间范围"],
        "landing_page_views": ["落地页浏览量"],
        "add_to_cart": ["加入购物车"],
        "initiate_checkout": ["结账发起次数"],
        "rate_click_to_lp": ["点击-落地页浏览转化率"],
        "rate_lp_to_atc": ["落地页浏览-加购转化率"],
        "rate_atc_to_ic": ["加购-结账转化率"],
        "rate_ic_to_pur": ["结账-购买转化率"]
    },
    "异常指标": {
        "anomaly_metric_name": ["异常指标"],
        "mom_change": ["环比"]
    },
    "广告架构": {**COMMON_METRICS, "dimension_item": ["广告类型"]},
    "受众组": {
        **COMMON_METRICS,
        "dimension_item": ["广告组", "广告组Id", "Ad Set Name"],
        "custom_audience_settings": ["设置的自定义受众", "Custom Audiences"],
        "converting_keywords": ["产生成效的关键词", "Interests", "Keywords"],
        "converting_countries": ["产生成效的国家", "国家", "地区", "Country", "Region", "Location"],
        "converting_genders": ["产生成效的性别", "性别", "Gender"],
        "converting_ages": ["产生成效的年龄", "年龄", "Age", "Age Group"]
    },
    "受众类型": {**COMMON_METRICS, "dimension_item": ["受众类型"]},
    "国家": {**COMMON_METRICS, "dimension_item": ["国家/地区", "国家"]},
    "年龄": {**COMMON_METRICS, "dimension_item": ["年龄"]},
    "性别": {**COMMON_METRICS, "dimension_item": ["性别"]},
    "平台&版位": {**COMMON_METRICS, "dimension_item": ["平台&版位"]},
    "素材": {
        **COMMON_METRICS,
        "content_item": ["素材"],
        "cvr_lp_to_pur": ["落地页浏览-购买转化率"]
    },
    "落地页": {
        **COMMON_METRICS,
        "content_item": ["落地页url", "落地页"],
        "ctr_all": ["曝光-点击转化率"],
        "rate_lp_to_atc": ["落地页浏览-加购转化率", "落地页浏览-购物转化率"]
    }
}

GROUP_CONFIG = {
    "Master_Overview": ["整体数据", "分时段数据", "异常指标"],
    "Master_Breakdown": ["广告架构", "受众组", "受众类型", "国家", "年龄", "性别", "平台&版位"],
    "Master_Creative": ["素材", "落地页"]
}

REPORT_MAPPING = {
    "spend": "花费 ($)", "roas": "ROAS", "purchases": "购买次数", "purchase_value": "购买总价值",
    "cpa": "CPA ($)", "ctr": "CTR (%)", "cpm": "CPM ($)", "cpc": "CPC ($)", "aov": "客单价",
    "impressions": "展现量", "clicks_all": "点击量 (All)", "clicks": "点击量 (All)", "ctr_all": "点击率 (All)",
    "landing_page_views": "落地页访问量", "add_to_cart": "加购次数", "initiate_checkout": "结账发起数 (IC)",
    "rate_click_to_lp": "点击 → 落地页访问转化率", "rate_lp_to_atc": "落地页 → 加购转化率",
    "rate_atc_to_ic": "加购 → 购买转化率", "rate_ic_to_pur": "购买转化率",
    "cvr_purchase": "点击 → 购买转化率", "cvr_lp_to_pur": "CVR (全站转化率)",
    "date_range": "日期/时段", "campaign_type": "投放模式", "adset_name": "广告组ID", "adset_id": "广告组ID",
    "custom_audience_settings": "自定义受众源", "converting_keywords": "高潜兴趣词", "audience_type": "受众策略",
    "country": "国家", "age_group": "年龄", "gender": "性别", "creative_name": "素材名称", "placement": "版位",
    "landing_page_url": "页面 URL", "mom_change": "环比波动", "anomaly_metric_name": "异常项",
    "converting_countries": "产生成效的国家", "converting_genders": "产生成效的性别", "converting_ages": "产生成效的年龄"
}

# ✅ 增强了模糊匹配别名 (修复核心：增加了add_to_cart等字段的映射)
FIELD_ALIASES = {
    "adset_id": ["adset_id", "ad set id", "adset id", "广告组编号", "广告组id", "adset_name", "ad set name"],
    "converting_countries": ["converting_countries", "country", "region", "国家", "地区", "location"],
    "converting_genders": ["converting_genders", "gender", "性别"],
    "converting_ages": ["converting_ages", "age", "年龄", "age_group"],
    "converting_keywords": ["converting_keywords", "keywords", "interests", "兴趣", "关键词"],
    "spend": ["spend", "amount spent", "cost", "花费", "消耗"],
    "purchases": ["purchases", "results", "result", "成效", "购买"],
    "roas": ["roas", "return on ad spend", "purchase roas"],
    "purchase_value": ["purchase_value", "conversion value", "value", "总价值", "gmv", "购买总价值"],
    "clicks": ["clicks", "clicks (all)", "点击量", "clicks_all"],
    "impressions": ["impressions", "展示", "展现"],
    "ctr_all": ["ctr_all", "ctr (all)", "点击率 (all)"],
    # ✅ 修复位置：新增以下三行映射，确保计算函数能找到中文列名
    "add_to_cart": ["add_to_cart", "加入购物车", "加购", "cart"],
    "initiate_checkout": ["initiate_checkout", "结账发起次数", "结账", "checkout"],
    "landing_page_views": ["landing_page_views", "落地页浏览量", "落地页", "landing"]
}


# ==========================================
# PART 2: 核心工具函数 (已修复百分比识别问题)
# ==========================================

def parse_float(value):
    """辅助函数：清理数据并将字符串/数字安全转换为浮点数"""
    if value is None:
        return 0.0
    try:
        if isinstance(value, (int, float)):
            return float(value)
        return clean_numeric_strict(value)
    except (ValueError, TypeError):
        return 0.0

def safe_div(numerator, denominator, multiplier=1.0):
    n = parse_float(numerator)
    d = parse_float(denominator)
    if d > 0:
        return (n / d) * multiplier
    else:
        return 0.0

def clean_numeric(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).strip().replace('$', '').replace('¥', '').replace(',', '')
    if '%' in val_str: 
        val_str = val_str.replace('%', '')
        try: return float(val_str) / 100.0 
        except: return 0.0
    try: return float(val_str)
    except: return val

def clean_numeric_strict(val): 
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).strip().replace('$', '').replace('¥', '').replace(',', '')
    if '%' in val_str: 
        val_str = val_str.replace('%', '')
        try: return float(val_str) / 100.0
        except: return 0.0
    try: return float(val_str)
    except: return 0.0

def find_column_fuzzy(df, keywords):
    for kw in keywords:
        if kw in df.columns: return kw
    df_cols_norm = {c.lower().replace(' ', '').replace('_', ''): c for c in df.columns}
    for kw in keywords:
        kw_norm = kw.lower().replace(' ', '').replace('_', '')
        if kw_norm in df_cols_norm: return df_cols_norm[kw_norm]
    for col in df.columns:
        col_lower = col.lower()
        for kw in keywords:
            if kw.lower() in col_lower: return col
    return None

def calc_metrics_dict(df_chunk):
    res = {}
    if df_chunk.empty: return res
    sums = {}
    targets = ['spend', 'clicks', 'impressions', 'purchases', 'purchase_value',
               'landing_page_views', 'add_to_cart', 'initiate_checkout']
    
    for t in targets:
        aliases = FIELD_ALIASES.get(t, [t])
        if t == 'purchase_value' and 'value' not in aliases: aliases.append('value')
        col = find_column_fuzzy(df_chunk, aliases)
        if col:
             sums[t] = df_chunk[col].apply(clean_numeric_strict).sum()
        else:
             sums[t] = 0.0

    res['spend'] = parse_float(sums.get('spend', 0))
    res['impressions'] = parse_float(sums.get('impressions', 0))
    res['clicks'] = parse_float(sums.get('clicks', 0))
    res['purchases'] = parse_float(sums.get('purchases', 0))
    res['purchase_value'] = parse_float(sums.get('purchase_value', 0))
    res['add_to_cart'] = parse_float(sums.get('add_to_cart', 0)) # ✅ 确保写入结果
    res['roas'] = safe_div(sums.get('purchase_value'), sums.get('spend'))
    res['cpm'] = safe_div(sums.get('spend'), sums.get('impressions'), multiplier=1000)
    res['cpc'] = safe_div(sums.get('spend'), sums.get('clicks'))
    res['ctr'] = safe_div(sums.get('clicks'), sums.get('impressions'))
    res['cpa'] = safe_div(sums.get('spend'), sums.get('purchases'))
    res['cvr_purchase'] = safe_div(sums.get('purchases'), sums.get('clicks'))
    res['rate_click_to_lp'] = safe_div(sums.get('landing_page_views'), sums.get('clicks'))
    res['rate_lp_to_atc']   = safe_div(sums.get('add_to_cart'), sums.get('landing_page_views'))
    res['rate_atc_to_ic']   = safe_div(sums.get('initiate_checkout'), sums.get('add_to_cart'))
    res['rate_ic_to_pur']   = safe_div(sums.get('purchases'), sums.get('initiate_checkout'))
    res['aov'] = safe_div(sums.get('purchase_value'), sums.get('purchases'))

    date_col = find_column_fuzzy(df_chunk, ['date', 'time', 'range'])
    if date_col:
        try:
            dates = pd.to_datetime(df_chunk[date_col], errors='coerce').dropna()
            if not dates.empty: res['date_range'] = f"{dates.min():%Y-%m-%d} ~ {dates.max():%Y-%m-%d}"
            else: res['date_range'] = "-"
        except: res['date_range'] = "-"
    else: res['date_range'] = "-"
    return res 

def format_cell(key, val, is_mom=False):
    if isinstance(val, str): return val
    if is_mom:
        if key == 'date_range': return val
        return f"{val:+.2%}"
    k = str(key).lower()
    if 'roas' in k: return f"{val:.2f}"
    if any(x in k for x in ['rate', 'ctr', 'cvr', '点击率', '转化率', '着陆率', '意向率', '成功率']): 
        return f"{val:.2%}" 
    if any(x in k for x in ['spend', 'cpm', 'cpc', 'value', 'aov', 'cpa', '花费', '金额', '客单价', 'gmv', '价值']): return f"{val:,.2f}"
    if any(x in k for x in ['purchases', 'cart', 'click', '次数', '单量', '点击', '展现', '访问量', '发起数']): return f"{val:,.0f}"
    return f"{val}"

def extract_benchmark_values(df_bench):
    targets = {'roas': (['roas'], True), 'cpm': (['cpm'], False), 'ctr': (['ctr'], True), 'cpc': (['cpc'], False), 'cpa': (['cpa_purchase', 'cpa'], False)}
    extracted = {}
    for metric, (aliases, higher_better) in targets.items():
        found_col = None
        for alias in aliases:
            found_col = find_column_fuzzy(df_bench, [alias])
            if found_col: break
        if found_col:
            try:
                s = df_bench[found_col].apply(clean_numeric_strict)
                v = s[s>0].mean()
                if metric in ['ctr'] and v > 1.0:
                    v = v / 100.0
                if not pd.isna(v): extracted[metric] = [v, higher_better]
            except: pass
    return extracted

def add_hyperlink(paragraph, url, text, color="0000FF", underline=True):
    try:
        part = paragraph.part
        r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), r_id)
        new_run = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        if color:
            c = OxmlElement('w:color')
            c.set(qn('w:val'), color)
            rPr.append(c)
        if underline:
            u = OxmlElement('w:u')
            u.set(qn('w:val'), 'single')
            rPr.append(u)
        new_run.append(rPr)
        new_run.text = text
        hyperlink.append(new_run)
        paragraph._p.append(hyperlink)
        return hyperlink
    except: return None

def apply_report_labels(df, custom_mapping=None):
    if df.empty: return df
    mapping = REPORT_MAPPING.copy()
    if custom_mapping: mapping.update(custom_mapping)
    return df.rename(columns=mapping)

def add_df_to_word(doc, df, title, level=1):
    if df.empty: return
    doc.add_heading(title, level=level)
    t = doc.add_table(rows=df.shape[0]+1, cols=df.shape[1])
    t.style = 'Table Grid'
    is_creative = "素材" in title
    is_landing = "落地页" in title
    link_col_idx = -1
    for j, col in enumerate(df.columns):
        cell = t.cell(0, j)
        cell.text = str(col)
        if any(x in str(col).lower() for x in ["url", "link", "素材", "内容", "content"]): link_col_idx = j
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.bold = True
                r.font.size = Pt(8)
    for i in range(df.shape[0]):
        label_prefix = "素材" if is_creative else ("落地页" if is_landing else "")
        label_char = chr(65 + (i % 26))
        if i >= 26: label_char += str(i // 26)
        label_text = f"{label_prefix}{label_char}"
        for j in range(df.shape[1]):
            val = df.iat[i, j]
            cell = t.cell(i+1, j)
            if (is_creative or is_landing) and j == link_col_idx:
                try:
                    p = cell.paragraphs[0]
                    url = str(val).strip()
                    if len(url) > 5: add_hyperlink(p, url, label_text)
                    else: cell.text = label_text
                except: cell.text = label_text
            else:
                cell.text = str(val)
                if "结论" in str(df.columns[j]):
                    if "✅" in str(val): cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 128, 0)
                    if "⚠️" in str(val): cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 0, 0)
            for p in cell.paragraphs:
                for r in p.runs: r.font.size = Pt(8)
    doc.add_paragraph("\n")

import re

def _to_number_maybe(x):
    """把 '4,226.45' / '2.45%' / '+1.29%' 这类字符串转成 number。
    规则：
    - 百分号：转成 ratio（2.45% -> 0.0245）
    - 纯数字（含逗号/货币符号）：转 float
    - 其他：原样返回
    """
    if x is None:
        return None
    if isinstance(x, (int, float)) and not (np.isnan(x) or np.isinf(x)):
        return float(x)

    if not isinstance(x, str):
        return x

    s = x.strip()
    if s == "" or s.lower() == "nan":
        return None

    # 去掉货币符号与千分位
    s2 = s.replace(",", "").replace("$", "").replace("¥", "")

    # 百分比：+1.29% / -44.40% / 2.45%
    if s2.endswith("%"):
        try:
            return float(s2[:-1]) / 100.0
        except:
            return x

    # 普通数字
    if re.fullmatch(r"[+-]?\d+(\.\d+)?", s2):
        try:
            return float(s2)
        except:
            return x

    return x


def normalize_v2_payload(v2):
    """统一：
    1) 表格 rows 数值化（尤其 t1_data_overview）
    2) cpc 字段改名为 'CPC ($)'，并同步 columns 与 rows keys
    3) benchmark 的 '指标'：把 'CPC' 改成 'CPC ($)'
    """
    tables = v2.get("tables", {})
    for table_id, t in tables.items():
        cols = t.get("columns", [])
        rows = t.get("rows", [])

        # 2) 统一 cpc 列名：cpc -> CPC ($)
        if "cpc" in cols:
            cols = ["CPC ($)" if c == "cpc" else c for c in cols]
            t["columns"] = cols

        # rows：逐 cell 数值化 + key 改名
        new_rows = []
        for r in rows:
            if not isinstance(r, dict):
                new_rows.append(r)
                continue

            rr = {}
            for k, val in r.items():
                kk = "CPC ($)" if k == "cpc" else k
                rr[kk] = _to_number_maybe(val)
            new_rows.append(rr)

        t["rows"] = new_rows

        # 额外：benchmark 的指标值对齐
        if table_id == "t2_industry_benchmark":
            for rr in t["rows"]:
                if isinstance(rr, dict) and rr.get("指标") == "CPC":
                    rr["指标"] = "CPC ($)"

    return v2

def infer_column_types(columns):
    """
    基于列名推断类型：rate/money/count/text/mom/ratio
    你可以按你自己的列名继续补充关键词。
    """
    types = {}
    for c in columns:
        name = str(c)

        # 1) rate（百分比显示，JSON 存 0~1）
        if any(k in name for k in ["CTR", "转化率", "CVR", "点击 →", "落地页 →", "购买转化率", "%"]):
            types[name] = "rate"
            continue
            
        # 2) 文本列
        if any(k in name for k in ["日期", "时段", "国家", "性别", "年龄", "受众", "版位", "素材", "落地页", "URL", "页面", "指标", "对比结论"]):
            types[name] = "text"
            continue

        # 3) 环比 / MoM
        if any(k in name for k in ["环比", "MoM", "mom"]):
            types[name] = "mom"
            continue

        # 4) ROAS / ratio（不带%）
        if "ROAS" in name:
            types[name] = "ratio"
            continue

        # 5) money
        if any(k in name for k in ["花费", "金额", "CPM", "CPC", "CPA", "客单价", "购买总价值", "($)"]):
            types[name] = "money"
            continue

        # 6) count
        if any(k in name for k in ["次数", "量", "数", "展现", "点击", "访问", "加购", "购买"]):
            types[name] = "count"
            continue

        # 默认：text
        types[name] = "text"

    return types

def json_safe(obj):
    """
    递归清理 JSON 数据：
    - np.nan / inf / -inf → None
    - 确保 json.dumps 输出严格 JSON
    """
    if isinstance(obj, float):
        if np.isnan(obj) or np.isinf(obj):
            return None
        return obj

    if isinstance(obj, dict):
        return {k: json_safe(v) for k, v in obj.items()}

    if isinstance(obj, list):
        return [json_safe(v) for v in obj]

    return obj

def build_tables_and_plan(old_json: dict) -> dict:
    """
    把 processor.final_json（你现在的结构：章节key -> records）
    转成 v2：
      - report_meta
      - tables: table_id -> {title, columns, rows}
      - report_plan: sections/blocks（text 先留空，后续Day3让LLM填）
    """
    report_meta = {
        "report_title": old_json.get("report_title", "广告投放深度分析报告"),
        "generated_at": old_json.get("generated_at", pd.Timestamp.now().strftime("%Y-%m-%d"))
    }

    tables = {}
    sections = []

    def add_table(table_id: str, title: str, rows):
        if rows is None:
            return
        if isinstance(rows, list) and len(rows) > 0 and isinstance(rows[0], dict):
            columns = list(rows[0].keys())
        else:
            columns = []
        col_types = infer_column_types(columns)

        tables[table_id] = {
            "title": title,
            "columns": columns,
            "column_types": col_types,   # ✅ 新增这一行
            "rows": rows
        }

        if table_id == "t2_industry_benchmark":
            # benchmark 是“指标驱动表”，单靠列名推断不准，直接固定类型
            tables[table_id]["column_types"] = {
                "指标": "text",
                "当前账户": "number",
                "行业基准": "number",
                "对比结论": "text"
            }

    # 1. 数据大盘总览
    if "1_data_overview" in old_json:
        add_table("t1_data_overview", "1. 数据大盘总览", old_json["1_data_overview"])
        sections.append({
            "id": "overview",
            "title": "1. 数据大盘总览",
            "blocks": [
                {"type": "text", "text": ""},  # Day3 接 LLM 后填
                {"type": "table_ref", "table_id": "t1_data_overview"}
            ]
        })

    # 2. Benchmark
    if "2_industry_benchmark" in old_json:
        add_table("t2_industry_benchmark", "2. 行业 Benchmark 对比", old_json["2_industry_benchmark"])
        sections.append({
            "id": "benchmark",
            "title": "2. 行业 Benchmark 对比",
            "blocks": [
                {"type": "table_ref", "table_id": "t2_industry_benchmark"},
                {"type": "text", "text": ""}
            ]
        })

    # 3. 受众分析（old_json['3_audience_analysis'] 是 dict）
    aud = old_json.get("3_audience_analysis", {})
    if isinstance(aud, dict) and aud:
        mapping = {
            "3.1 国家分析": "t3_country",
            "3.2 性别分析": "t4_gender",
            "3.3 年龄分析": "t5_age",
            "3.4 受众组分析表": "t6_adset",
            "3.5 受众类型分析": "t7_audience_type"
        }
        blocks = [{"type": "text", "text": ""}]
        for sub_title, rows in aud.items():
            table_id = mapping.get(sub_title, f"t3_{sub_title.replace(' ', '').replace('.', '_')}")
            add_table(table_id, sub_title, rows)
            blocks.append({"type": "table_ref", "table_id": table_id})
            blocks.append({"type": "text", "text": ""})

        sections.append({
            "id": "audience",
            "title": "3. 受众分析",
            "blocks": blocks
        })

    # 4. 素材
    if "4_creative_analysis" in old_json:
        add_table("t8_creatives", "4. 素材分析", old_json["4_creative_analysis"])
        sections.append({
            "id": "creative",
            "title": "4. 素材分析",
            "blocks": [
                {"type": "text", "text": ""},
                {"type": "table_ref", "table_id": "t8_creatives"}
            ]
        })

    # 5. 版位分析
    placement = old_json.get("5_placement_analysis", {})
    if isinstance(placement, dict) and placement:
        if "top_spend" in placement:
            add_table("t9_placement_top_spend", "5.1 版位花费 TOP 5", placement["top_spend"])
        if "high_potential" in placement:
            add_table("t10_placement_high_potential", "5.2 版位高潜力", placement["high_potential"])
        sections.append({
            "id": "placement",
            "title": "5. 版位分析",
            "blocks": [
                {"type": "text", "text": ""},
                {"type": "table_ref", "table_id": "t9_placement_top_spend"},
                {"type": "table_ref", "table_id": "t10_placement_high_potential"},
                {"type": "text", "text": ""}
            ]
        })

    # 6. 落地页
    if "6_landing_page_analysis" in old_json:
        add_table("t11_landing_pages", "6. 落地页分析", old_json["6_landing_page_analysis"])
        sections.append({
            "id": "landing",
            "title": "6. 落地页分析",
            "blocks": [
                {"type": "text", "text": ""},
                {"type": "table_ref", "table_id": "t11_landing_pages"}
            ]
        })


    return {
        "report_meta": report_meta,
        "tables": tables,
        "report_plan": {"sections": sections}
    }


# ==========================================
# PART 3: 主逻辑类
# ==========================================

class AdReportProcessor:
    def __init__(self, raw_file, bench_file=None):
        self.raw_file = raw_file
        self.bench_file = bench_file
        self.processed_dfs = {}
        self.merged_dfs = {}
        self.final_json = {}
        self.doc = Document()

    def process_etl(self):
        xls = pd.ExcelFile(self.raw_file)
        for sheet_name, mapping in SHEET_MAPPINGS.items():
            if sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                final_cols = {}
                for std_col, raw_col_options in mapping.items():
                    matched_col = None
                    for option in raw_col_options:
                        if option in df.columns: matched_col = option; break
                        if not matched_col:
                            for df_col in df.columns:
                                if option.replace(" ", "") == df_col.replace(" ", ""): matched_col = df_col; break
                        if matched_col: break
                    if matched_col: final_cols[std_col] = matched_col

                if final_cols:
                    df_clean = df[list(final_cols.values())].rename(columns={v: k for k, v in final_cols.items()})
                    text_cols = ['date_range', 'anomaly_metric_name', 
                                 'converting_keywords', 'converting_countries', 'converting_genders', 'converting_ages', 
                                 'custom_audience_settings', 'dimension_item', 'content_item']
                    
                    for col in df_clean.columns:
                        if col not in text_cols:
                            df_clean[col] = df_clean[col].apply(clean_numeric)

                    if sheet_name in ["素材", "落地页", "受众组"]:
                        if "spend" in df_clean.columns:
                            df_clean = df_clean.sort_values("spend", ascending=False).head(10)

                    df_clean["Source_Sheet"] = sheet_name
                    self.processed_dfs[sheet_name] = df_clean

        for master_name, source_sheets in GROUP_CONFIG.items():
            dfs_to_merge = [self.processed_dfs[src] for src in source_sheets if src in self.processed_dfs]
            if dfs_to_merge:
                merged_df = pd.concat(dfs_to_merge, ignore_index=True)
                cols = list(merged_df.columns)
                priority_cols = ['Source_Sheet', 'date_range', 'dimension_item', 'content_item',
                                 'spend', 'roas', 'purchases', 'cpa']
                new_order = [c for c in priority_cols if c in cols] + [c for c in cols if c not in priority_cols]
                self.merged_dfs[master_name] = merged_df[new_order]

    def generate_report(self):
        benchmark_targets = {'roas': [2.0, True], 'cpm': [20.0, False], 'ctr': [0.015, True], 'cpc': [1.5, False], 'cpa': [30.0, False]}
        if self.bench_file:
            try:
                df_b = pd.read_excel(self.bench_file)
                benchmark_targets = extract_benchmark_values(df_b)
            except: pass

        self.doc.add_heading('广告投放深度分析报告', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.final_json = {"report_title": "广告投放深度分析报告", "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d")}

        # 1. 大盘总览
        df_ov = pd.DataFrame()
        if "Master_Overview" in self.merged_dfs:
            df_src = self.merged_dfs["Master_Overview"]
            mask = df_src['Source_Sheet'].astype(str).apply(lambda x: any(k in x for k in ["分时", "Time"]))
            df_ov = df_src[mask].copy() if not df_src[mask].empty else df_src.copy()

        if not df_ov.empty:
            date_col = find_column_fuzzy(df_ov, ['date', 'time', '时间'])
            if date_col:
                try:
                    df_ov['temp_date'] = pd.to_datetime(df_ov[date_col], errors='coerce')
                    df_clean = df_ov.dropna(subset=['temp_date']).sort_values('temp_date')
                    dates = df_clean['temp_date'].unique()
                    raw_overall = calc_metrics_dict(df_clean)
                    if len(dates) >= 2:
                        mid_date = dates[len(dates)//2]
                        raw_prev = calc_metrics_dict(df_clean[df_clean['temp_date'] < mid_date])
                        raw_curr = calc_metrics_dict(df_clean[df_clean['temp_date'] >= mid_date])
                        raw_mom = {}
                        for k, v_curr in raw_curr.items():
                            if k == 'date_range': raw_mom[k] = "-"
                            else:
                                v_prev = raw_prev.get(k, 0)
                                raw_mom[k] = (v_curr - v_prev) / v_prev if v_prev > 0 else 0.0
                    else:
                        raw_prev = {k: "-" for k in raw_overall}; raw_curr = raw_overall; raw_mom = {k: "-" for k in raw_overall}

                    col_order = ["date_range", "spend", "roas", "cpa", "cpm", "cpc", "ctr", "cvr_purchase",
                                 "rate_click_to_lp", "rate_lp_to_atc", "rate_ic_to_pur", "aov", "add_to_cart", "purchases", "purchase_value"]
                    final_data_raw = []
                    for label, r in zip(["整体数据", "前半周期", "后半周期", "环比"], [raw_overall, raw_prev, raw_curr, raw_mom]):
                        row = {"Label": label}
                        is_m = (label == "环比")
                        for c in col_order:
                            row[c] = r.get(c, 0.0)
                        row["date_range"] = label
                        final_data_raw.append(row)

                    df_raw = pd.DataFrame(final_data_raw, columns=col_order)

                    # Word 展示版本
                    df_disp = df_raw.copy()
                    for c in col_order:
                        if c == "date_range":
                            continue
                        is_mom_row = (df_disp["date_range"] == "环比")
                        df_disp.loc[~is_mom_row, c] = df_disp.loc[~is_mom_row, c].apply(
                            lambda v: format_cell(c, v, is_mom=False)
                        )
                        df_disp.loc[is_mom_row, c] = df_disp.loc[is_mom_row, c].apply(
                            lambda v: format_cell(c, v, is_mom=True)
                        )

                    df_disp = apply_report_labels(df_disp)
                    add_df_to_word(self.doc, df_disp, "1. 数据大盘总览", level=1)

                    self.final_json["1_data_overview"] = apply_report_labels(df_raw).to_dict(orient="records")


                    # 2. Benchmark
                    raw_current = calc_metrics_dict(df_clean)
                    bench_data = []
                    for metric_key in ['roas', 'cpm', 'ctr', 'cpc', 'cpa']:
                        curr_val = raw_current.get(metric_key, 0)
                        bench_val, higher_is_better = benchmark_targets.get(metric_key, [0, True])
                        conclusion = "-"
                        if curr_val != 0:
                            diff = curr_val - bench_val
                            if higher_is_better: conclusion = "✅ 优于大盘" if diff > 0 else ("⚠️ 低于大盘" if diff < 0 else "持平")
                            else: conclusion = "✅ 优于大盘" if diff < 0 else ("⚠️ 高于大盘" if diff > 0 else "持平")
                        bench_data.append({
                            "指标": REPORT_MAPPING.get(metric_key, metric_key.upper()),
                            "当前账户": format_cell(metric_key, curr_val),
                            "行业基准": format_cell(metric_key, bench_val),
                            "对比结论": conclusion
                        })
                    df_b = pd.DataFrame(bench_data)
                    add_df_to_word(self.doc, df_b, "2. 行业 Benchmark 对比", level=1)
                    self.final_json['2_industry_benchmark'] = df_b.to_dict(orient='records')
                except Exception as e: st.warning(f"大盘计算警告: {e}")

              # 3. 受众分析（拆分：受众组 + 受众类型）
        self.doc.add_heading("3. 受众分析", level=1)
        self.final_json['3_audience_analysis'] = {}

        def build_audience_table(df_curr, title, dim_label, top10=False, include_converting=False, level=2):
            if df_curr.empty:
                return

            # 补算常用指标（如果缺失）
            if 'clicks' in df_curr.columns and 'spend' in df_curr.columns and not find_column_fuzzy(df_curr, ['cpc']):
                df_curr['cpc'] = df_curr['spend'] / df_curr['clicks'].replace(0, np.nan)
            if 'impressions' in df_curr.columns and 'spend' in df_curr.columns and not find_column_fuzzy(df_curr, ['cpm']):
                df_curr['cpm'] = (df_curr['spend'] / df_curr['impressions'].replace(0, np.nan)) * 1000
                # 如果没有 CTR 列，先计算 CTR（比例）
                    # 如果没有 CTR 列，先计算 CTR（比例）
                   # 如果没有 CTR 列，先计算 CTR（比例）
            if not find_column_fuzzy(df_curr, ['ctr']):
                if 'impressions' in df_curr.columns and 'clicks' in df_curr.columns:
                    df_curr['ctr'] = df_curr['clicks'] / df_curr['impressions'].replace(0, np.nan)
                else:
                    df_curr['ctr'] = np.nan

            # 如果同时有 cpc 和 cpm，但 ctr 为空或为 0，用反推公式补 CTR
            if 'cpc' in df_curr.columns and 'cpm' in df_curr.columns:
                mask_fix = (df_curr['ctr'].isna() | (df_curr['ctr'] == 0)) & (df_curr['cpc'] > 0)
                if mask_fix.any():
                    df_curr.loc[mask_fix, 'ctr'] = (
                        df_curr.loc[mask_fix, 'cpm'] /
                        (df_curr.loc[mask_fix, 'cpc'] * 1000)
                    )

            # ✅ 保持为比例 0~1
            df_curr['ctr'] = df_curr['ctr'].fillna(0)



            if 'purchases' in df_curr.columns and 'spend' in df_curr.columns and not find_column_fuzzy(df_curr, ['cpa']):
                df_curr['cpa'] = df_curr['spend'] / df_curr['purchases'].replace(0, np.nan)

            # 需要输出的列
            req_cols = ["dimension_item", "spend", "ctr", "cpc", "cpm", "cpa", "roas"]
            if include_converting:
                req_cols += ["custom_audience_settings", "converting_countries", "converting_keywords", "converting_genders", "converting_ages"]

            rename_map = {}
            valid_cols = []
            for req in req_cols:
                aliases = FIELD_ALIASES.get(req, [req])
                found = find_column_fuzzy(df_curr, aliases)
                if found:
                    valid_cols.append(found)
                    rename_map[found] = req
                else:
                    default_val = "-" if ("converting" in req or req == "custom_audience_settings") else 0.0
                    df_curr[req] = default_val
                    valid_cols.append(req)

            df_final = df_curr[valid_cols].rename(columns=rename_map)

            # 文本列清理
            for t_col in ["custom_audience_settings", "converting_countries", "converting_keywords", "converting_genders", "converting_ages"]:
                if t_col in df_final.columns:
                    df_final[t_col] = df_final[t_col].fillna("-").astype(str).replace("nan", "-")

            # 过滤 unknow
            if "dimension_item" in df_final.columns:
                df_final = df_final[~df_final['dimension_item'].astype(str).str.lower().str.contains('unknow', na=False)]

            # Top10
            if top10 and 'spend' in df_final.columns:
                df_final = df_final.sort_values('spend', ascending=False).head(10)

            df_clean = df_final.round(2)
            df_display = apply_report_labels(df_clean, custom_mapping={'dimension_item': dim_label})

            add_df_to_word(self.doc, df_display, title, level=level)
            self.final_json['3_audience_analysis'][title] = df_display.to_dict(orient='records')

        if "Master_Breakdown" in self.merged_dfs:
            df_bd = self.merged_dfs["Master_Breakdown"]

            # 3.1 国家
            df_country = df_bd[df_bd["Source_Sheet"].astype(str) == "国家"].copy()
            build_audience_table(df_country, "3.1 国家分析", "国家", top10=True, include_converting=False, level=2)

            # 3.2 性别
            df_gender = df_bd[df_bd["Source_Sheet"].astype(str) == "性别"].copy()
            build_audience_table(df_gender, "3.2 性别分析", "性别", top10=False, include_converting=False, level=2)

            # 3.3 年龄
            df_age = df_bd[df_bd["Source_Sheet"].astype(str) == "年龄"].copy()
            build_audience_table(df_age, "3.3 年龄分析", "年龄段", top10=False, include_converting=False, level=2)

            # 3.4 受众组（单独表）
            df_adset = df_bd[df_bd["Source_Sheet"].astype(str) == "受众组"].copy()
            build_audience_table(df_adset, "3.4 受众组分析表", "受众组名称", top10=True, include_converting=True, level=2)

            # 3.5 受众类型（单独表）
            df_audtype = df_bd[df_bd["Source_Sheet"].astype(str) == "受众类型"].copy()
            build_audience_table(df_audtype, "3.5 受众类型分析", "受众类型", top10=False, include_converting=False, level=2)



        # 4. 素材与落地页
        if "Master_Creative" in self.merged_dfs:
            df_cr = self.merged_dfs["Master_Creative"]
            for title, keywords, label, json_key in [("4. 素材分析", ["素材", "Creative"], "素材名称", "4_creative_analysis"), ("6. 落地页分析", ["落地页", "Landing"], "落地页 URL", "6_landing_page_analysis")]:
                mask = df_cr['Source_Sheet'].astype(str).apply(lambda x: any(k in x for k in keywords))
                df_curr = df_cr[mask].copy()
                if not df_curr.empty:
                    if not find_column_fuzzy(df_curr, ['cpc']): df_curr['cpc'] = df_curr['spend'] / df_curr['clicks'].replace(0, np.nan) if 'clicks' in df_curr else 0
                    if not find_column_fuzzy(df_curr, ['cpa']): df_curr['cpa'] = df_curr['spend'] / df_curr['purchases'].replace(0, np.nan) if 'purchases' in df_curr else 0
                    if not find_column_fuzzy(df_curr, ['ctr']):
                         if 'impressions' in df_curr and 'clicks' in df_curr: df_curr['ctr'] = df_curr['clicks'] / df_curr['impressions'].replace(0, np.nan)
                         else: df_curr['ctr'] = np.nan
                    if 'cpc' in df_curr.columns and 'cpm' in df_curr.columns:
                        mask_fix = (df_curr['ctr'].isna() | (df_curr['ctr'] == 0)) & (df_curr['cpc'] > 0)
                        if mask_fix.any(): df_curr.loc[mask_fix, 'ctr'] = df_curr.loc[mask_fix, 'cpm'] / (df_curr.loc[mask_fix, 'cpc'] * 1000)
                    df_curr['ctr'] = df_curr['ctr'].fillna(0)

                    req_cols = ["content_item", "spend", "ctr", "cpc", "cpm", "roas", "cpa"]
                    rename_map = {}; valid_cols = []
                    for req in req_cols:
                        aliases = FIELD_ALIASES.get(req, [req])
                        found = find_column_fuzzy(df_curr, aliases)
                        if found: valid_cols.append(found); rename_map[found] = req
                        else: df_curr[req] = 0.0; valid_cols.append(req)
                    df_final = df_curr[valid_cols].rename(columns=rename_map)
                    if 'spend' in df_final.columns: df_final = df_final.sort_values('spend', ascending=False).head(10)
                    df_clean = df_final.round(2) 
                    
                    df_display = apply_report_labels(df_clean, custom_mapping={'content_item': label})
                    add_df_to_word(self.doc, df_display, title, level=1)
                    self.final_json[json_key] = df_display.to_dict(orient='records')
                    
        # 5. 版位
        if "Master_Breakdown" in self.merged_dfs:
             self.doc.add_heading("5. 版位分析", level=1)
             df_bd = self.merged_dfs["Master_Breakdown"]
             mask = df_bd['Source_Sheet'].astype(str).apply(lambda x: any(k in x for k in ["版位", "Placement"]))
             df_curr = df_bd[mask].copy()
             if not df_curr.empty:
                 if not find_column_fuzzy(df_curr, ['cpc']): df_curr['cpc'] = df_curr['spend'] / df_curr['clicks'].replace(0, np.nan) if 'clicks' in df_curr else 0
                 if not find_column_fuzzy(df_curr, ['cpa']): df_curr['cpa'] = df_curr['spend'] / df_curr['purchases'].replace(0, np.nan) if 'purchases' in df_curr else 0
                 if not find_column_fuzzy(df_curr, ['ctr']): df_curr['ctr'] = df_curr['clicks'] / df_curr['impressions'].replace(0, np.nan) if 'impressions' in df_curr else 0
                 if not find_column_fuzzy(df_curr, ['cpm']): df_curr['cpm'] = (df_curr['spend'] / df_curr['impressions'].replace(0, np.nan)) * 1000 if 'impressions' in df_curr else 0
                 if 'ctr' in df_curr.columns:
                     df_curr['ctr'] = df_curr['ctr'].fillna(0)
                 req_cols = ['dimension_item', 'spend', 'ctr', 'cpc', 'cpm', 'roas', 'cpa']
                 rename_map = {}; valid_cols = []
                 for c in req_cols:
                     aliases = FIELD_ALIASES.get(c, [c])
                     f = find_column_fuzzy(df_curr, aliases)
                     if f: valid_cols.append(f); rename_map[f] = c
                     else: df_curr[c] = 0.0; valid_cols.append(c)
                 df_clean = df_curr[valid_cols].rename(columns=rename_map).round(2)
                 
                 df_top5 = df_clean.sort_values('spend', ascending=False).head(5)
                 add_df_to_word(self.doc, apply_report_labels(df_top5, {'dimension_item': '版位'}), "5.1 版位花费 TOP 5", level=2)
                 
                 mean_ctr = df_clean['ctr'].mean(); mean_cpm = df_clean['cpm'].mean()
                 mask_pot = (df_clean['ctr'] > mean_ctr) & (df_clean['cpm'] < mean_cpm)
                 df_pot = df_clean[mask_pot].sort_values('ctr', ascending=False).head(5)
                 if df_pot.empty: df_pot = df_clean.sort_values('ctr', ascending=False).head(5)
                 add_df_to_word(self.doc, apply_report_labels(df_pot, {'dimension_item': '版位'}), "5.2 版位高潜力", level=2)
                 
                 self.final_json['5_placement_analysis'] = {
                     "top_spend": apply_report_labels(df_top5, {'dimension_item': '版位'}).to_dict('records'),
                     "high_potential": apply_report_labels(df_pot, {'dimension_item': '版位'}).to_dict('records')
                 }


# ==========================================
# PART 4: Streamlit UI
# ==========================================

def main():
    st.set_page_config(page_title="Auto-ad-data", layout="wide")

    st.markdown("""
        <style>
        .stApp {
            background: radial-gradient(circle at 50% 20%, rgba(255, 240, 200, 0.6) 0%, rgba(240, 200, 255, 0.4) 30%, rgba(255, 255, 255, 1) 70%);
            background-attachment: fixed;
            background-size: cover;
        }
        .main-title {
            text-align: center;
            font-size: 3.5rem;
            font-weight: 800;
            margin-bottom: 0.5rem;
            font-family: -apple-system, BlinkMacSystemFont, sans-serif;
            background: linear-gradient(135deg, #662D8C 0%, #ED1E79 100%);
            background: -webkit-linear-gradient(135deg, #662D8C 0%, #ED1E79 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        .sub-title {
            text-align: center;
            font-size: 1.1rem;
            color: #666;
            margin-bottom: 3rem;
        }
        [data-testid="stVerticalBlockBorderWrapper"] {
            background: rgba(255, 255, 255, 0.45) !important;
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
            border: 1px solid rgba(255, 255, 255, 0.8) !important;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.04);
            border-radius: 24px !important;
            padding: 1rem;
            transition: all 0.3s ease;
        }
        [data-testid="stVerticalBlockBorderWrapper"]:hover {
            transform: translateY(-3px);
            box-shadow: 0 15px 35px rgba(102, 45, 140, 0.1);
            border-color: rgba(237, 30, 121, 0.2) !important;
            background: rgba(255, 255, 255, 0.65) !important;
        }
        .card-header {
            text-align: center;
            font-weight: 600;
            color: #4A4A4A;
            margin-bottom: 0.5rem;
            letter-spacing: 0.5px;
        }
        .icon-container {
            text-align: center;
            font-size: 2.5rem;
            margin-bottom: 5px;
            filter: drop-shadow(0 2px 4px rgba(0,0,0,0.1));
        }
        [data-testid='stFileUploader'] section {
            background-color: rgba(255, 255, 255, 0.25);
            backdrop-filter: blur(15px);
            -webkit-backdrop-filter: blur(15px);
            border: 1.5px dashed rgba(102, 45, 140, 0.3);
            box-shadow: inset 0 0 20px rgba(255, 255, 255, 0.5);
            border-radius: 16px;
            padding: 1.5rem;
            transition: all 0.3s ease;
        }
        [data-testid='stFileUploader'] section:hover {
            background-color: rgba(255, 255, 255, 0.5);
            border-color: #ED1E79;
            box-shadow: 0 8px 20px rgba(102, 45, 140, 0.15);
        }
        [data-testid='stFileUploader'] button {
            border-radius: 20px;
            border-color: rgba(102, 45, 140, 0.2);
            color: #662D8C;
            background-color: rgba(255, 255, 255, 0.8);
        }
        .glass-info-box {
            padding: 1rem 1.5rem;
            margin-bottom: 1.5rem;
            display: flex;
            align-items: center;
            border-radius: 16px;
            background: rgba(255, 255, 255, 0.3);
            backdrop-filter: blur(15px);
            -webkit-backdrop-filter: blur(15px);
            border: 1px solid rgba(255, 255, 255, 0.7);
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
        }
        div.stButton > button {
            display: block;
            margin: 0 auto;
            width: 100%;
            background-image: linear-gradient(90deg, #B721FF 0%, #21D4FD 100%);
            color: white !important;
            border-radius: 30px;
            padding: 0.7rem 1.5rem;
            font-size: 1.1rem;
            font-weight: 600;
            border: none;
            box-shadow: 0 6px 20px rgba(183, 33, 255, 0.4); 
            transition: all 0.3s ease;
        }
        div.stButton > button:hover {
            opacity: 0.95;
            transform: translateY(-3px) scale(1.02);
            box-shadow: 0 10px 30px rgba(33, 212, 253, 0.5);
        }
        div.stDownloadButton > button {
            display: block;
            margin: 0 auto;
            width: 100%;
            background: linear-gradient(145deg, rgba(255, 255, 255, 1) 0%, rgba(245, 235, 255, 1) 100%);
            color: #662D8C !important; 
            border-radius: 30px;
            padding: 0.7rem 1.5rem;
            font-size: 1rem;
            font-weight: 700;
            border: 1px solid rgba(255, 255, 255, 0.8);
            box-shadow: 
                0 4px 10px rgba(102, 45, 140, 0.08),
                inset 0 1px 0 rgba(255, 255, 255, 0.9),
                inset 0 -2px 0 rgba(0, 0, 0, 0.03);
            transition: all 0.2s ease;
        }
        div.stDownloadButton > button:hover {
            transform: translateY(-2px);
            background: linear-gradient(145deg, #ffffff 0%, #f0e6ff 100%);
            border-color: #ED1E79;
            box-shadow: 0 8px 15px rgba(102, 45, 140, 0.15);
            color: #ED1E79 !important;
        }
        div[data-baseweb="notification"] {
            background-color: rgba(102, 45, 140, 0.05);
            border-left-color: #662D8C;
            border-radius: 12px;
        }
        </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="main-title">What can I help with?</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-title">请分别上传您的【周期性复盘报告】、【行业benchmark】数据文件，我将为您生成专业准确的广告优化【数据终表】。</div>', 
        unsafe_allow_html=True
    )

    col1, col_gap, col2 = st.columns([1, 0.1, 1])

    with col1:
        with st.container(border=True):
            st.markdown('<div class="icon-container">📊</div>', unsafe_allow_html=True)
            st.markdown('<div class="card-header">1.上传【周期性复盘报告】</div>', unsafe_allow_html=True)
            raw_file = st.file_uploader("", type=["xlsx", "xls"], key="raw_uploader", label_visibility="collapsed")

    with col2:
        with st.container(border=True):
            st.markdown('<div class="icon-container">🎯</div>', unsafe_allow_html=True)
            st.markdown('<div class="card-header">2.上传【行业 Benchmark]】</div>', unsafe_allow_html=True)
            bench_file = st.file_uploader("", type=["xlsx", "xls"], key="bench_uploader", label_visibility="collapsed")

    st.write("")
    st.write("")

    b_c1, b_c2, b_c3 = st.columns([1, 1, 1])
    with b_c2:
        start_btn = st.button("开始生成数据表 ✦", use_container_width=True)

    if start_btn:
        if not raw_file:
            st.error("⚠️ 请至少上传 [数据报表] 才能继续！")
            return

        processor = AdReportProcessor(raw_file, bench_file)

        try:
            with st.spinner("阶段 1/2: 数据清洗、Top10截断、降维合并..."):
                processor.process_etl()
                st.toast("✅ 阶段 1 完成：Master Tables 已生成", icon="✅")

            with st.expander("📄 点击查看处理后的数据预览 (Master Tables)", expanded=False):
                tabs = st.tabs(list(processor.merged_dfs.keys()))
                for i, (k, v) in enumerate(processor.merged_dfs.items()):
                    with tabs[i]: 
                        st.dataframe(v.head(20), use_container_width=True)

            with st.spinner("阶段 2/2: 生成架构诊断、Word报告 & JSON..."):
                processor.generate_report()
                st.toast("✅ 阶段 2 完成：所有报告已准备就绪", icon="🎉")
            
            st.balloons() 
            
            # 构建 v2 JSON
            v2_payload = build_tables_and_plan(processor.final_json)
            v2_payload = json_safe(v2_payload)
            
            # ✅ 统一规范化：数值化 + 命名对齐
            v2_payload = normalize_v2_payload(v2_payload)

            # 👇 这里加预览
            with st.expander("🔍 预览 v2 JSON（tables + report_plan）", expanded=False):
                st.json(v2_payload)
            
            st.markdown("### 📥 下载结果文件")
            
            with st.container(border=True):
                st.markdown("""
                    <div class="glass-info-box">
                        <span style="font-size: 1.2rem; margin-right: 0.8rem;">💡</span>
                        <span style="
                            font-weight: 600;
                            background: linear-gradient(135deg, #662D8C 0%, #ED1E79 100%);
                            -webkit-background-clip: text;
                            -webkit-text-fill-color: transparent;
                            font-family: -apple-system, BlinkMacSystemFont, sans-serif;
                        ">
                            建议：您可只选择下载 JSON 格式文件用于大模型分析，如有必要再下载其他格式文件。
                        </span>
                    </div>
                """, unsafe_allow_html=True)

                res_c1, res_c2, res_c3 = st.columns(3)

                json_str = json.dumps(v2_payload, indent=4, ensure_ascii=False)

                res_c1.download_button(
                    "📥 JSON (大模型分析)", 
                    json_str, 
                    "Ad_Report_Data.json", 
                    "application/json",
                    use_container_width=True
                )

                output_xls = io.BytesIO()
                with pd.ExcelWriter(output_xls, engine='xlsxwriter') as writer:
                    for name, df in processor.merged_dfs.items(): 
                        df.to_excel(writer, sheet_name=name, index=False)
                res_c2.download_button(
                    "📥 Excel (数据透视)", 
                    output_xls.getvalue(), 
                    "Merged_Ad_Report_Final.xlsx", 
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

                output_doc = io.BytesIO()
                processor.doc.save(output_doc)
                res_c3.download_button(
                    "📥 Word (数据审查)", 
                    output_doc.getvalue(), 
                    "Ad_Report_Final_V20_10.docx", 
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"❌ 处理过程中发生错误: {str(e)}")
            st.exception(e)

if __name__ == "__main__":
    main()
