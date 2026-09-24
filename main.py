import os
import io
import re
import sys
import math
import json
import time
import uuid
import secrets
import bcrypt
import certifi
import jwt
import difflib
from datetime import datetime, timedelta
from typing import List, Optional, Dict, Any, Tuple, Set
from fastapi import FastAPI, File, UploadFile, Form, HTTPException, Response, Depends
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.requests import Request
import pandas as pd
import numpy as np
import pymongo
from bson import ObjectId
from dotenv import load_dotenv
from app_logging import configure_logging, RequestLoggingMiddleware

# Load configuration and secrets from a local .env file, which is gitignored.
# override=False means real environment variables always win, so a hosted
# deployment (systemd, Docker, EC2) can set them directly with no .env present.
load_dotenv(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"),
    override=False,
)

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
logger = configure_logging(WORKSPACE_DIR)

app = FastAPI(
    title="Mutual Fund & Dynamic Excel Analytics Platform",
    description="High-performance mutual fund analysis with Sharpe Ratio, Rolling Returns, PyMongo & MongoDB Compass integration.",
    version="2.4.0"
)

app.add_middleware(RequestLoggingMiddleware)

TEMPLATES_DIR = os.path.join(WORKSPACE_DIR, "templates")
STATIC_DIR = os.path.join(WORKSPACE_DIR, "static")
STATIC_CSS_DIR = os.path.join(STATIC_DIR, "css")
STATIC_JS_DIR = os.path.join(STATIC_DIR, "js")
STATIC_IMAGES_DIR = os.path.join(STATIC_DIR, "images")

# Cache-busting token appended to css/js URLs. Bump it after a frontend change
# so browsers pick up the new file instead of a stale cached copy.
ASSET_VERSION = os.getenv("ASSET_VERSION", "8.0.2")

# Ensure static & templates exist
for _d in (STATIC_DIR, STATIC_CSS_DIR, STATIC_JS_DIR, STATIC_IMAGES_DIR, TEMPLATES_DIR):
    os.makedirs(_d, exist_ok=True)

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
templates = Jinja2Templates(directory=TEMPLATES_DIR)


# ==========================================
# HELPER FUNCTIONS: DYNAMIC EXCEL PARSING
# ==========================================

def clean_val(v):
    if pd.isna(v) or v is None:
        return None
    s = str(v).strip()
    if s in ["N/A*", "N/A", "NA", "-", "--", "null", "None", ""]:
        return None
    try:
        # Check if numeric
        s_clean = s.replace(",", "").replace("%", "")
        f = float(s_clean)
        if math.isnan(f) or math.isinf(f):
            return None
        return f
    except ValueError:
        return s


def normalize_excel_return_percentages(df, workbook, sheet_name, header_row):
    """Convert explicitly percent-formatted return cells to percentage points.

    A plain 0.12 stays 0.12; a numeric 0.12 displayed as 12% becomes 12.
    Ratio, AUM and other columns are left untouched.
    """
    mapping = map_fund_columns(df)
    columns = {df.columns.get_loc(mapping[key]) for key in
               ("rolling_1y", "rolling_2y", "rolling_3y", "rolling_5y")
               if mapping.get(key) is not None}
    if not columns:
        return

    def apply(row_index, column_index, number_format):
        # Quoted/escaped percent symbols are literals, not Excel scaling rules.
        fmt = re.sub(r'"[^"]*"|\\.|_.|\*.', '', number_format or '')
        if "%" not in fmt:
            return
        value = df.iat[row_index, column_index]
        if pd.notna(value) and isinstance(value, (int, float, np.number)):
            df.iat[row_index, column_index] = float(value) * 100

    if hasattr(workbook, "sheet_by_name"):  # xlrd, with formatting_info=True
        sheet = workbook.sheet_by_name(sheet_name)
        for row in range(len(df)):
            for col in columns:
                cell = sheet.cell(header_row + 1 + row, col)
                xf = workbook.xf_list[cell.xf_index]
                apply(row, col, workbook.format_map[xf.format_key].format_str)
    else:  # openpyxl: enumerate rows once (important for read-only workbooks).
        sheet = workbook[sheet_name]
        for row, cells in enumerate(sheet.iter_rows(
                min_row=header_row + 2, max_row=header_row + 1 + len(df),
                min_col=1, max_col=max(columns) + 1)):
            for col in columns:
                apply(row, col, cells[col].number_format)


def smart_read_sheet(file_bytes: bytes, filename: str, sheet_name=None) -> Dict[str, Any]:
    """
    Reads an Excel or CSV file dynamically, discovers sheets,
    finds the genuine header row by scanning first 15 rows,
    and returns a clean DataFrame and sheet metadata.
    """
    ext = os.path.splitext(filename)[1].lower()
    
    if ext == ".csv":
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), header=None)
        except Exception:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding="latin1", header=None)
        sheets = ["Default"]
        active_sheet = "Default"
    else:
        # Excel file (.xls or .xlsx)
        xl = pd.ExcelFile(io.BytesIO(file_bytes),
                          engine_kwargs={"formatting_info": True} if ext == ".xls" else {})
        sheets = xl.sheet_names
        active_sheet = sheet_name if sheet_name in sheets else sheets[0]
        df = xl.parse(active_sheet, header=None)

    if df.empty or len(df) < 2:
        if ext != ".csv":
            xl.close()
        return {"df": pd.DataFrame(), "sheets": sheets, "active_sheet": active_sheet, "header_row": 0}

    # Find the most likely header row by scoring text columns and non-null count
    best_header_row = 0
    max_score = -1
    
    header_keywords = [
        "scheme", "fund", "name", "category", "aum", "sharpe", "sharp",
        "ratio", "return", "std", "deviation", "inception", "rolling",
        "treynor", "information", "1y", "2y", "3y", "5y"
    ]

    for r_idx in range(min(15, len(df))):
        row_vals = [str(x).strip().lower() for x in df.iloc[r_idx].dropna().tolist()]
        if not row_vals:
            continue
        
        # Keyword matches
        matches = sum(1 for v in row_vals if any(kw in v for kw in header_keywords))
        unique_texts = len(set(row_vals))
        
        score = (matches * 10) + unique_texts + len(row_vals)
        if score > max_score and matches > 0:
            max_score = score
            best_header_row = r_idx

    # If no keyword matched, find first row with mostly unique non-numeric strings
    if max_score <= 0:
        for r_idx in range(min(10, len(df))):
            row_vals = [str(x).strip() for x in df.iloc[r_idx].dropna().tolist()]
            if len(row_vals) >= 2 and all(not x.replace('.', '', 1).isdigit() for x in row_vals):
                best_header_row = r_idx
                break

    # Extract headers and data
    raw_headers = [str(df.iloc[best_header_row, c]).strip() if pd.notna(df.iloc[best_header_row, c]) else f"Column_{c+1}" 
                   for c in range(df.shape[1])]
    
    # Handle duplicate header names
    seen = {}
    clean_headers = []
    for h in raw_headers:
        if not h or h.lower().startswith("unnamed:"):
            h = f"Column_{len(clean_headers)+1}"
        if h in seen:
            seen[h] += 1
            clean_headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            clean_headers.append(h)

    data_df = df.iloc[best_header_row + 1:].copy()
    data_df.columns = clean_headers
    if ext != ".csv":
        try:
            normalize_excel_return_percentages(data_df, xl.book, active_sheet, best_header_row)
        finally:
            xl.close()
    
    # Drop rows that are completely empty or look like trailing disclaimer text
    data_df = data_df.dropna(how="all")
    
    # Filter out footer disclaimers where only 1 column is populated with long text
    valid_rows = []
    for idx, row in data_df.iterrows():
        non_null_count = row.dropna().count()
        first_col_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ""
        if non_null_count <= 1 and (len(first_col_val) > 60 or "disclaimer" in first_col_val or "phone" in first_col_val or "email" in first_col_val or "arn-" in first_col_val):
            continue
        if non_null_count > 0:
            valid_rows.append(idx)
            
    if valid_rows:
        data_df = data_df.loc[valid_rows]

    return {
        "df": data_df.reset_index(drop=True),
        "sheets": sheets,
        "active_sheet": active_sheet,
        "header_row": best_header_row,
        "columns": clean_headers
    }


def map_fund_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """
    Fuzzy matches dataframe column names to standard mutual fund attributes.
    """
    mapping = {
        "name": None,
        "category": None,
        "aum": None,
        "sharpe": None,
        "beta": None,
        "alpha": None,
        "sortino": None,
        "info_ratio": None,
        "treynor": None,
        "std_dev": None,
        "inception_date": None,
        "rolling_1y": None,
        "rolling_2y": None,
        "rolling_3y": None,
        "rolling_5y": None
    }
    
    cols = df.columns.tolist()

    # Step 1: Detect Fund / Scheme Name column with priority
    # Exclude metadata columns that contain the word "fund" or "scheme" but are NOT the fund name
    exclude_name_words = [
        "house", "manager", "family", "size", "rating", "rank", "code",
        "class", "benchmark", "type", "category", "sub category", "corpus",
        "aum", "asset", "url", "link", "idcw", "nav", "date"
    ]
    
    # Priority 1: Exact column matches for scheme name
    for c in cols:
        cl = c.lower().replace("_", " ").replace("-", " ").strip()
        if cl in ["scheme name", "fund name", "scheme", "fund", "security name", "security", "scheme / plan name", "instrument name"]:
            mapping["name"] = c
            break

    # Priority 2: Contains "scheme name" or "fund name"
    if not mapping["name"]:
        for c in cols:
            cl = c.lower().replace("_", " ").replace("-", " ").strip()
            if any(k in cl for k in ["scheme name", "fund name", "scheme / plan", "security name"]):
                if not any(ex in cl for ex in ["house", "manager", "rating", "rank", "size"]):
                    mapping["name"] = c
                    break

    # Priority 3: Contains "scheme", "fund", or "security" without exclude words
    if not mapping["name"]:
        for c in cols:
            cl = c.lower().replace("_", " ").replace("-", " ").strip()
            if any(k in cl for k in ["scheme", "fund", "security"]):
                if not any(ex in cl for ex in exclude_name_words):
                    mapping["name"] = c
                    break
    
    for c in cols:
        cl = c.lower().replace("_", " ").replace("-", " ").strip()
        
        # Category
        if not mapping["category"] and any(k in cl for k in ["category", "asset class", "type", "sub category"]):
            mapping["category"] = c
        # AUM
        elif not mapping["aum"] and any(k in cl for k in ["aum", "asset", "corpus", "size"]):
            mapping["aum"] = c
        # Sharpe Ratio (handles Sharp / Sharpe)
        elif not mapping["sharpe"] and any(k in cl for k in ["sharpe", "sharp"]):
            mapping["sharpe"] = c
        # Beta
        elif not mapping["beta"] and any(k in cl for k in ["beta", "fund beta", "market beta", "systemic risk"]):
            mapping["beta"] = c
        # Alpha
        elif not mapping["alpha"] and any(k in cl for k in ["alpha", "jensen alpha", "jensens alpha", "excess alpha"]):
            mapping["alpha"] = c
        # Sortino Ratio
        elif not mapping["sortino"] and any(k in cl for k in ["sortino", "sortino ratio"]):
            mapping["sortino"] = c
        # Information Ratio
        elif not mapping["info_ratio"] and (any(k in cl for k in ["information ratio", "info ratio"]) or re.search(r"\bir\b", cl)):
            mapping["info_ratio"] = c
        # Treynor Ratio
        elif not mapping["treynor"] and any(k in cl for k in ["treynor", "treynor ratio"]):
            mapping["treynor"] = c
        # Standard Deviation
        elif not mapping["std_dev"] and any(k in cl for k in ["standard deviation", "std dev", "stddev", "volatility", "sd"]):
            mapping["std_dev"] = c
        # Inception Date
        elif not mapping["inception_date"] and any(k in cl for k in ["inception", "launch", "start date"]):
            mapping["inception_date"] = c
        # Rolling Returns
        elif not mapping["rolling_1y"] and any(k in cl for k in ["1y rolling", "1 yr rolling", "1y ret", "1y", "1 year"]):
            mapping["rolling_1y"] = c
        elif not mapping["rolling_2y"] and any(k in cl for k in ["2y rolling", "2 yr rolling", "2y ret", "2y", "2 year"]):
            mapping["rolling_2y"] = c
        elif not mapping["rolling_3y"] and any(k in cl for k in ["3y rolling", "3 yr rolling", "3y ret", "3y", "3 year"]):
            mapping["rolling_3y"] = c
        elif not mapping["rolling_5y"] and any(k in cl for k in ["5y rolling", "5 yr rolling", "5y ret", "5y", "5 year"]):
            mapping["rolling_5y"] = c

    # Fallback for name if still none
    if not mapping["name"] and len(cols) > 0:
        mapping["name"] = cols[0]

    return mapping


def analyze_fund_dataset(df: pd.DataFrame, col_map: Dict[str, Optional[str]], custom_benchmark: Optional[float] = None) -> Dict[str, Any]:
    """
    Computes all risk, return, benchmark averages, highlights, category aggregates, and distribution stats.
    """
    records = []
    name_col = col_map.get("name")
    cat_col = col_map.get("category")
    aum_col = col_map.get("aum")
    sharpe_col = col_map.get("sharpe")
    beta_col = col_map.get("beta")
    alpha_col = col_map.get("alpha")
    sortino_col = col_map.get("sortino")
    info_col = col_map.get("info_ratio")
    treynor_col = col_map.get("treynor")
    std_col = col_map.get("std_dev")
    inc_col = col_map.get("inception_date")
    r1_col = col_map.get("rolling_1y")
    r2_col = col_map.get("rolling_2y")
    r3_col = col_map.get("rolling_3y")
    r5_col = col_map.get("rolling_5y")

    sharpe_vals = []
    
    for _, row in df.iterrows():
        # A row with no value in the scheme-name column isn't a fund at all — it's almost
        # always leftover footer/disclaimer text that spilled into other columns (e.g. a
        # multi-line disclaimer paragraph that landed in the "Category" cell). Previously
        # this fell back to a placeholder name "Unknown Fund" and was kept as a fake record;
        # skip it instead so it can never surface as a "fund" anywhere downstream.
        if not name_col or pd.isna(row[name_col]):
            continue
        name = str(row[name_col]).strip()
        if not name or name.lower() in ["nan", "none", ""]:
            continue
            
        category = str(row[cat_col]).strip() if cat_col and pd.notna(row[cat_col]) else "General"
        aum = clean_val(row[aum_col]) if aum_col else None
        sharpe = clean_val(row[sharpe_col]) if sharpe_col else None
        beta = clean_val(row[beta_col]) if beta_col else None
        alpha = clean_val(row[alpha_col]) if alpha_col else None
        sortino = clean_val(row[sortino_col]) if sortino_col else None
        info_ratio = clean_val(row[info_col]) if info_col else None
        treynor = clean_val(row[treynor_col]) if treynor_col else None
        std_dev = clean_val(row[std_col]) if std_col else None
        inception = str(row[inc_col]).strip() if inc_col and pd.notna(row[inc_col]) else ""
        
        r1 = clean_val(row[r1_col]) if r1_col else None
        r2 = clean_val(row[r2_col]) if r2_col else None
        r3 = clean_val(row[r3_col]) if r3_col else None
        r5 = clean_val(row[r5_col]) if r5_col else None
        
        # Calculate rolling avg if available
        rollings = [x for x in [r1, r2, r3, r5] if isinstance(x, (int, float))]
        rolling_avg = round(float(np.mean(rollings)), 2) if rollings else None

        is_sharpe_valid = isinstance(sharpe, (int, float)) and not math.isnan(sharpe)
        if is_sharpe_valid:
            sharpe_vals.append(float(sharpe))

        records.append({
            "name": name,
            "category": category,
            "aum": aum if isinstance(aum, (int, float)) else None,
            "sharpe": sharpe if is_sharpe_valid else None,
            "sharpe_valid": is_sharpe_valid,
            "beta": beta if isinstance(beta, (int, float)) else None,
            "alpha": alpha if isinstance(alpha, (int, float)) else None,
            "sortino": sortino if isinstance(sortino, (int, float)) else None,
            "info_ratio": info_ratio if isinstance(info_ratio, (int, float)) else None,
            "treynor": treynor if isinstance(treynor, (int, float)) else None,
            "std_dev": std_dev if isinstance(std_dev, (int, float)) else None,
            "inception_date": inception,
            "rolling_1y": r1 if isinstance(r1, (int, float)) else None,
            "rolling_2y": r2 if isinstance(r2, (int, float)) else None,
            "rolling_3y": r3 if isinstance(r3, (int, float)) else None,
            "rolling_5y": r5 if isinstance(r5, (int, float)) else None,
            "rolling_avg": rolling_avg
        })

    return compute_records_analysis(records, custom_benchmark=custom_benchmark, col_map=col_map)


def compute_records_analysis(records: List[Dict[str, Any]], custom_benchmark: Optional[float] = None, col_map: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    Given a list of fund record dicts, computes Sharpe benchmarks, ranks,
    category breakdowns, outperformance deltas, and distribution statistics.
    """
    sharpe_vals = [r["sharpe"] for r in records if r.get("sharpe_valid") and r.get("sharpe") is not None]
    total_funds = len(records)
    valid_sharpes_count = len(sharpe_vals)
    
    if valid_sharpes_count > 0:
        avg_sharpe = float(np.mean(sharpe_vals))
        median_sharpe = float(np.median(sharpe_vals))
        max_sharpe = float(np.max(sharpe_vals))
        min_sharpe = float(np.min(sharpe_vals))
        std_sharpe = float(np.std(sharpe_vals))
    else:
        avg_sharpe = 0.0
        median_sharpe = 0.0
        max_sharpe = 0.0
        min_sharpe = 0.0
        std_sharpe = 0.0

    benchmark = custom_benchmark if custom_benchmark is not None else avg_sharpe
    benchmark_rounded = round(benchmark, 4)

    # Rank and flag above/below average
    sorted_indices = sorted(
        [i for i, r in enumerate(records) if r.get("sharpe_valid") and r.get("sharpe") is not None],
        key=lambda i: records[i]["sharpe"],
        reverse=True
    )
    
    # Clear previous ranks first
    for r in records:
        r["rank"] = None

    for rank, idx in enumerate(sorted_indices, start=1):
        records[idx]["rank"] = rank

    above_avg_count = 0
    below_avg_count = 0
    na_count = 0
    total_aum = 0.0

    for r in records:
        if r.get("aum") is not None and isinstance(r["aum"], (int, float)):
            total_aum += r["aum"]
            
        if r.get("sharpe_valid") and r.get("sharpe") is not None:
            is_above = r["sharpe"] >= benchmark_rounded
            r["above_avg"] = is_above
            r["sharpe_diff"] = round(r["sharpe"] - benchmark_rounded, 2)
            if is_above:
                above_avg_count += 1
            else:
                below_avg_count += 1
        else:
            r["above_avg"] = False
            r["sharpe_diff"] = None
            na_count += 1

    # Overall ratio averages
    info_vals = [r["info_ratio"] for r in records if r.get("info_ratio") is not None and isinstance(r["info_ratio"], (int, float))]
    avg_info_ratio = round(float(np.mean(info_vals)), 2) if info_vals else None
    treynor_vals = [r["treynor"] for r in records if r.get("treynor") is not None and isinstance(r["treynor"], (int, float))]
    avg_treynor = round(float(np.mean(treynor_vals)), 2) if treynor_vals else None

    # Category summaries
    categories = sorted(list(set(r["category"] for r in records if r.get("category"))))
    category_stats = {}
    for cat in categories:
        cat_records = [r for r in records if r.get("category") == cat]
        cat_sharpes = [r["sharpe"] for r in cat_records if r.get("sharpe_valid") and r.get("sharpe") is not None]
        cat_info = [r["info_ratio"] for r in cat_records if r.get("info_ratio") is not None and isinstance(r["info_ratio"], (int, float))]
        cat_treynor = [r["treynor"] for r in cat_records if r.get("treynor") is not None and isinstance(r["treynor"], (int, float))]
        cat_aum = sum(r["aum"] for r in cat_records if r.get("aum") is not None and isinstance(r["aum"], (int, float)))
        cat_above = sum(1 for r in cat_records if r.get("above_avg", False))
        
        category_stats[cat] = {
            "count": len(cat_records),
            "valid_sharpe_count": len(cat_sharpes),
            "avg_sharpe": round(float(np.mean(cat_sharpes)), 2) if cat_sharpes else None,
            "median_sharpe": round(float(np.median(cat_sharpes)), 2) if cat_sharpes else None,
            "avg_info_ratio": round(float(np.mean(cat_info)), 2) if cat_info else None,
            "avg_treynor": round(float(np.mean(cat_treynor)), 2) if cat_treynor else None,
            "total_aum": round(cat_aum, 2),
            "above_avg_count": cat_above,
            "outperformance_rate": round((cat_above / len(cat_records)) * 100, 1) if cat_records else 0
        }

    # Top performer
    top_performer = records[sorted_indices[0]] if sorted_indices else None

    # Distribution histogram
    hist_bins = [
        {"label": "< 0", "min": -999, "max": 0},
        {"label": "0.0 - 0.2", "min": 0, "max": 0.2},
        {"label": "0.2 - 0.4", "min": 0.2, "max": 0.4},
        {"label": "0.4 - 0.6", "min": 0.4, "max": 0.6},
        {"label": "0.6 - 0.8", "min": 0.6, "max": 0.8},
        {"label": "0.8 - 1.0", "min": 0.8, "max": 1.0},
        {"label": "1.0 - 1.2", "min": 1.0, "max": 1.2},
        {"label": "> 1.2", "min": 1.2, "max": 999}
    ]
    for b in hist_bins:
        b["count"] = sum(1 for s in sharpe_vals if b["min"] <= s < b["max"])

    return {
        "funds": records,
        "summary": {
            "total_funds": total_funds,
            "valid_sharpe_count": valid_sharpes_count,
            "avg_sharpe": round(avg_sharpe, 2),
            "avg_sharpe_exact": avg_sharpe,
            "median_sharpe": round(median_sharpe, 2),
            "max_sharpe": round(max_sharpe, 2),
            "min_sharpe": round(min_sharpe, 2),
            "std_sharpe": round(std_sharpe, 2),
            "avg_info_ratio": avg_info_ratio,
            "avg_treynor": avg_treynor,
            "benchmark": round(benchmark, 2),
            "above_avg_count": above_avg_count,
            "below_avg_count": below_avg_count,
            "na_count": na_count,
            "outperformance_rate": round((above_avg_count / valid_sharpes_count) * 100, 1) if valid_sharpes_count > 0 else 0.0,
            "total_aum": round(total_aum, 2),
            "category_count": len(categories),
            "top_performer": top_performer
        },
        "categories": categories,
        "category_stats": category_stats,
        "distribution": hist_bins,
        "column_mapping": col_map or {}
    }


# ==========================================
# SCHEME-NAME MATCHING
# ==========================================
_PLAN_MAP = {
    "reg": "regular", "regular": "regular", "rp": "regular",
    "dir": "direct", "direct": "direct", "dp": "direct",
    "gr": "growth", "growth": "growth", "g": "growth",
    "idcw": "idcw", "div": "idcw", "divd": "idcw", "dividend": "idcw",
    "payout": "payout", "reinv": "reinvest", "reinvest": "reinvest",
    "bonus": "bonus"
}

# Noise tokens that carry no fund identity
_NOISE_TOKENS = {
    "fund", "funds", "plan", "plans", "scheme", "schemes",
    "the", "of", "and", "mutual", "portfolio", "trust",
    "series", "option", "options", "class"
}

_COMPOUND_REPLACEMENTS = [
    (r'\b(small|mid|large|flexi|multi|micro|mega)cap\b', r'\1 cap'),
    (r'\b(tax)saver\b', r'\1 saver'),
    (r'\b(index)fund\b', r'\1 fund'),
    (r'\b(blue)chip\b', r'\1 chip'),
]

_CATEGORY_ABBREVIATIONS = [
    (r'\bretrmnt\b', 'retirement'),
    (r'\bchildrens\b', 'children'),
    (r'\bchildren\'s\b', 'children'),
    (r'\bcons\b', 'conservative'),
    (r'\bag\b', 'aggressive'),
    (r'\baggr\b', 'aggressive'),
    (r'\bdyn\b', 'dynamic'),
    (r'\bmod\b', 'moderate'),
    (r'\bsch\b', 'scheme'),
    (r'\bsavngs\b', 'savings'),
    (r'\bsavings\b', 'savings'),
    (r'\byojna\b', 'yojana'),
    (r'\bopp\b', 'opportunities'),
    (r'\bopps\b', 'opportunities'),
    (r'\bopportunity\b', 'opportunities'),
    (r'\beq\b', 'equity'),
    (r'\bequities\b', 'equity'),
    (r'\bbal\b', 'balanced'),
    (r'\badv\b', 'advantage'),
]

_AMC_ALIASES = [
    (r'\bicici\s+pru\b', 'icici prudential'),
    (r'\bppfas\b', 'parag parikh'),
    (r'\babsl\b', 'aditya birla sun life'),
    (r'\bbirla\s+sun\s+life\b', 'aditya birla sun life'),
    (r'\baditya\s+birla\b', 'aditya birla sun life'),
    (r'\bfranklin\s+templeton\b', 'franklin india'),
    (r'\bkotak\s+mahindra\b', 'kotak'),
    (r'\bnippon\s+india\b', 'nippon'),
    (r'\bmirae\s+asset\b', 'mirae'),
    (r'\bquant\s+mutual\b', 'quant'),
    (r'\btata\s+mutual\b', 'tata'),
    (r'\buti\s+mutual\b', 'uti'),
    (r'\bidfc\b', 'bandhan'),
]


def clean_scheme_name(name: Any) -> str:
    """Cleans punctuation, prefixes, AMC acronyms, and compound words."""
    if not name:
        return ""
    text = str(name).lower().strip()
    text = text.replace("&", " and ")
    
    # Strip metadata in parentheses like (formerly ...), (benchmark ...), (tier 1), etc.
    text = re.sub(r'\((formerly|erstwhile|benchmark|tier|star|isin|bse|nse)[^)]*\)', ' ', text)
    
    # Strip common category prefixes like "Equity : Small Cap - " or "Large Cap: "
    text = re.sub(r'^(equity|debt|hybrid|other)\s*:\s*([a-z\s]+-\s*)?', ' ', text)
    text = re.sub(r'^(large|mid|small|flexi|multi)\s*cap\s*:\s*', ' ', text)
    
    # Standardize compound words like "smallcap" -> "small cap"
    for pat, repl in _COMPOUND_REPLACEMENTS:
        text = re.sub(pat, repl, text)

    # Standardize category and scheme abbreviations (e.g. retrmnt -> retirement)
    for pat, repl in _CATEGORY_ABBREVIATIONS:
        text = re.sub(pat, repl, text)
        
    # Standardize AMC aliases
    for pat, repl in _AMC_ALIASES:
        text = re.sub(pat, repl, text)
        
    # Replace non-alphanumeric with spaces
    text = re.sub(r'[^a-z0-9]+', ' ', text)
    return text.strip()


def normalized_scheme_key(name: Any) -> str:
    """Strict key: clean alphanumeric."""
    cleaned = clean_scheme_name(name)
    return re.sub(r'[^a-z0-9]', '', cleaned)


def extract_scheme_tokens(name: Any) -> Tuple[Set[str], Set[str]]:
    """Splits fund name into core distinguishing tokens and plan/share-class tokens."""
    cleaned = clean_scheme_name(name)
    tokens = re.findall(r'[a-z0-9]+', cleaned)
    core = set()
    plan = set()
    for tok in tokens:
        if tok in _NOISE_TOKENS:
            continue
        if tok in _PLAN_MAP:
            plan.add(_PLAN_MAP[tok])
        else:
            core.add(tok)
    return core, plan


def are_plans_compatible(plan_a: Set[str], plan_b: Set[str]) -> bool:
    """
    Prevents cross-contamination between Direct and Regular share classes or Growth vs IDCW.
    If neither or only one specifies a plan, returns True (asymmetric tolerance).
    """
    if "direct" in plan_a and "regular" in plan_b:
        return False
    if "regular" in plan_a and "direct" in plan_b:
        return False
    if "growth" in plan_a and "idcw" in plan_b:
        return False
    if "idcw" in plan_a and "growth" in plan_b:
        return False
    return True


def scheme_signature(name: Any) -> frozenset:
    """Returns the set of core tokens for backward compatibility."""
    core, _ = extract_scheme_tokens(name)
    return frozenset(core)


def build_rolling_matcher(rolling_funds: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Indexes rolling-returns records for multi-tier matching:
    1. Exact normalized key
    2. Full signature (core + plan)
    3. Core signature (tolerant of omitted Regular/Growth)
    4. Token subset / superset
    5. Jaccard token similarity
    6. SequenceMatcher fuzzy similarity
    """
    records = []
    by_key: Dict[str, List[Dict[str, Any]]] = {}
    by_full_sig: Dict[frozenset, List[Dict[str, Any]]] = {}
    by_core_sig: Dict[frozenset, List[Dict[str, Any]]] = {}

    for rf in rolling_funds:
        raw_name = rf.get("name")
        if not raw_name:
            continue
        core, plan = extract_scheme_tokens(raw_name)
        key = normalized_scheme_key(raw_name)
        full_sig = frozenset(core | plan)
        core_sig = frozenset(core)
        
        entry = {
            "fund": rf,
            "raw_name": raw_name,
            "key": key,
            "core": core,
            "plan": plan,
            "full_sig": full_sig,
            "core_sig": core_sig,
            "cleaned": clean_scheme_name(raw_name)
        }
        records.append(entry)
        by_key.setdefault(key, []).append(entry)
        by_full_sig.setdefault(full_sig, []).append(entry)
        by_core_sig.setdefault(core_sig, []).append(entry)

    return {
        "records": records,
        "by_key": by_key,
        "by_full_sig": by_full_sig,
        "by_core_sig": by_core_sig,
    }


def match_rolling_record(matcher: Dict[str, Any], name: Any, consumed_ids: Optional[Set[int]] = None) -> Dict[str, Any]:
    """
    Matches a fund name against rolling records using multi-tiered resolution:
    - Tier 1: Exact alphanumeric key
    - Tier 2: Full signature (core + plan)
    - Tier 3: Core signature (plan compatible or omitted)
    - Tier 4: Token subset / superset
    - Tier 5: Jaccard token similarity
    - Tier 6: Fuzzy SequenceMatcher fallback
    """
    if not name:
        return {"record": None, "how": None, "ambiguous": False}
        
    consumed = consumed_ids or set()
    raw_query = str(name)
    q_cleaned = clean_scheme_name(raw_query)
    q_key = normalized_scheme_key(raw_query)
    q_core, q_plan = extract_scheme_tokens(raw_query)
    q_full_sig = frozenset(q_core | q_plan)
    q_core_sig = frozenset(q_core)

    def is_available(entry: Dict[str, Any]) -> bool:
        # Similar spelling cannot override an explicit size, index or series identity.
        identity_tokens = {"large", "mid", "small", "micro", "mega", "flexi", "multi", "index", "etf", "fof"}
        query_identity = (q_core & identity_tokens) | {t for t in q_core if t.isdigit()}
        entry_core = entry["core"]
        entry_identity = (entry_core & identity_tokens) | {t for t in entry_core if t.isdigit()}
        return id(entry["fund"]) not in consumed and query_identity == entry_identity

    # Tier 1: Exact normalized key match
    key_candidates = [e for e in matcher["by_key"].get(q_key, []) if is_available(e)]
    if len(key_candidates) == 1:
        return {"record": key_candidates[0]["fund"], "how": "exact", "ambiguous": False}
    elif len(key_candidates) > 1:
        compat = [e for e in key_candidates if are_plans_compatible(q_plan, e["plan"])]
        if len(compat) == 1:
            return {"record": compat[0]["fund"], "how": "exact", "ambiguous": False}

    # Tier 2: Full signature match (core + plan tokens exact match)
    full_sig_cands = [e for e in matcher["by_full_sig"].get(q_full_sig, []) if is_available(e)]
    if len(full_sig_cands) == 1:
        return {"record": full_sig_cands[0]["fund"], "how": "signature", "ambiguous": False}
    elif len(full_sig_cands) > 1:
        return {"record": None, "how": None, "ambiguous": True}

    # Tier 3: Core signature match (core matches, plan compatible or omitted in one sheet)
    core_sig_cands = [e for e in matcher["by_core_sig"].get(q_core_sig, []) if is_available(e)]
    compat_core = [e for e in core_sig_cands if are_plans_compatible(q_plan, e["plan"])]
    if len(compat_core) == 1:
        return {"record": compat_core[0]["fund"], "how": "signature", "ambiguous": False}
    elif len(compat_core) > 1:
        exact_plan = [e for e in compat_core if q_plan == e["plan"]]
        if len(exact_plan) == 1:
            return {"record": exact_plan[0]["fund"], "how": "signature", "ambiguous": False}
        return {"record": None, "how": None, "ambiguous": True}

    # Tier 4: Token subset / superset core match
    if len(q_core) >= 2:
        subset_cands = []
        for e in matcher["records"]:
            if not is_available(e):
                continue
            if not are_plans_compatible(q_plan, e["plan"]):
                continue
            e_core = e["core"]
            if len(e_core) < 2:
                continue
            overlap = q_core & e_core
            max_len = max(len(q_core), len(e_core))
            if (q_core.issubset(e_core) or e_core.issubset(q_core)) and (len(overlap) / max_len >= 0.70):
                subset_cands.append((len(overlap) / max_len, e))
                
        if subset_cands:
            subset_cands.sort(key=lambda x: x[0], reverse=True)
            best_score, best_cand = subset_cands[0]
            if len(subset_cands) == 1 or best_score > subset_cands[1][0] + 0.10:
                return {"record": best_cand["fund"], "how": "signature", "ambiguous": False}

    # Tier 5: Jaccard token similarity match
    if len(q_core) >= 2:
        jaccard_cands = []
        for e in matcher["records"]:
            if not is_available(e):
                continue
            if not are_plans_compatible(q_plan, e["plan"]):
                continue
            e_core = e["core"]
            if not e_core:
                continue
            intersection = len(q_core & e_core)
            union = len(q_core | e_core)
            sim = intersection / union if union > 0 else 0
            if sim >= 0.65:
                jaccard_cands.append((sim, e))
                
        if jaccard_cands:
            jaccard_cands.sort(key=lambda x: x[0], reverse=True)
            best_sim, best_cand = jaccard_cands[0]
            if len(jaccard_cands) == 1 or (best_sim - jaccard_cands[1][0] >= 0.12):
                return {"record": best_cand["fund"], "how": "signature", "ambiguous": False}

    # Tier 6: SequenceMatcher fuzzy match on cleaned strings
    if len(q_cleaned) >= 5:
        fuzzy_cands = []
        for e in matcher["records"]:
            if not is_available(e):
                continue
            if not are_plans_compatible(q_plan, e["plan"]):
                continue
            ratio = difflib.SequenceMatcher(None, q_cleaned, e["cleaned"]).ratio()
            if ratio >= 0.82:
                fuzzy_cands.append((ratio, e))
        if fuzzy_cands:
            fuzzy_cands.sort(key=lambda x: x[0], reverse=True)
            best_ratio, best_cand = fuzzy_cands[0]
            if len(fuzzy_cands) == 1 or (best_ratio - fuzzy_cands[1][0] >= 0.08):
                return {"record": best_cand["fund"], "how": "signature", "ambiguous": False}

    return {"record": None, "how": None, "ambiguous": False}


def merge_risk_and_rolling(risk_res: Dict[str, Any], rolling_res: Dict[str, Any], custom_benchmark: Optional[float] = None) -> Dict[str, Any]:
    """
    Merges risk ratio dataset with rolling returns dataset by exact fund-name matching.
    Funds that exist in the Rolling-Returns sheet but have no matching Risk-Ratios entry
    (e.g. a different scheme, or an entire category like Hybrid/Solution Oriented that the
    Risk-Ratios sheet doesn't cover) are NOT dropped — they're kept as their own rows using
    whatever data the Rolling-Returns sheet itself computed for them, and their category is
    folded into the combined universe so it shows up in the category filters.
    """
    risk_funds = risk_res["funds"]
    rolling_funds = rolling_res["funds"]

    matcher = build_rolling_matcher(rolling_funds)

    merged_funds = []
    matched_count = 0
    signature_matched_count = 0
    ambiguous_names = []
    consumed = set()

    for f in risk_funds:
        item = dict(f)

        # Multi-tiered match tolerant of Plan/Option absence, acronyms, compound words, and formatting
        result = match_rolling_record(matcher, f.get("name"), consumed_ids=consumed)
        matched_rolling = result["record"]
        if result["ambiguous"]:
            ambiguous_names.append(f.get("name"))

        if matched_rolling:
            matched_count += 1
            if result.get("how") != "exact":
                signature_matched_count += 1
            consumed.add(id(matched_rolling))
            item["std_dev"] = matched_rolling.get("std_dev") or item.get("std_dev")
            item["beta"] = matched_rolling.get("beta") or item.get("beta")
            item["alpha"] = matched_rolling.get("alpha") or item.get("alpha")
            item["sortino"] = matched_rolling.get("sortino") or item.get("sortino")
            item["inception_date"] = matched_rolling.get("inception_date") or item.get("inception_date")
            item["rolling_1y"] = matched_rolling.get("rolling_1y")
            item["rolling_2y"] = matched_rolling.get("rolling_2y")
            item["rolling_3y"] = matched_rolling.get("rolling_3y")
            item["rolling_5y"] = matched_rolling.get("rolling_5y")
            item["rolling_avg"] = matched_rolling.get("rolling_avg")

        merged_funds.append(item)

    # Keep Rolling-Returns funds that had no Risk-Ratios counterpart (different scheme,
    # or a category the Risk-Ratios sheet doesn't have at all) instead of silently
    # discarding that data.
    rolling_only_funds = [rf for rf in rolling_funds if id(rf) not in consumed]
    merged_funds.extend(dict(rf) for rf in rolling_only_funds)

    # Recompute universe-wide stats (ranks, benchmark comparison, category list,
    # category_stats, distribution) over the full combined fund list, so categories that
    # only exist in the Rolling-Returns sheet are included everywhere (category filters,
    # Best Funds category filter, etc.) rather than reflecting only the Risk-Ratios universe.
    recalculated = compute_records_analysis(
        merged_funds,
        custom_benchmark=custom_benchmark,
        col_map=risk_res.get("column_mapping")
    )
    recalculated["rolling_funds"] = rolling_funds
    recalculated["summary"]["rolling_matched_count"] = matched_count
    recalculated["summary"]["rolling_only_count"] = len(rolling_only_funds)
    recalculated["match_report"] = {
        "rolling_rows": len(rolling_funds),
        "matched": matched_count,
        "matched_by_signature": signature_matched_count,
        "unmatched": len(rolling_only_funds),
        "unmatched_names": [rf.get("name") for rf in rolling_only_funds][:50],
        "ambiguous_names": ambiguous_names[:50],
    }
    return recalculated



# ==========================================
# PERSISTENCE & DATABASE LAYER (PYMONGO / MONGODB COMPASS)
# ==========================================

MONGO_URI = os.getenv("MONGO_URI", "mongodb://127.0.0.1:27017/")
MONGO_DB_NAME = os.getenv("MONGO_DB_NAME", "mutual_funds_db")

MONGO_IS_REMOTE = (
    os.getenv("MONGO_IS_REMOTE") == "1"
    or MONGO_URI.startswith("mongodb+srv://")
    or not any(host in MONGO_URI for host in ("127.0.0.1", "localhost", "mongo"))
)
# Cloud round-trips need a longer handshake budget than a loopback connection.
MONGO_TIMEOUT_MS = int(os.getenv("MONGO_TIMEOUT_MS", "8000" if MONGO_IS_REMOTE else "2000"))

_mongo_client = None
_mongo_db = None
_mongo_last_failed_at = 0.0
_mongo_retry_cooldown_sec = 10


def require_mongo_db():
    """
    Returns the Mongo database, raising 503 when MongoDB is unreachable.
    """
    mdb = get_mongo_db()
    if mdb is None:
        raise HTTPException(
            status_code=503,
            detail="The MongoDB database is unreachable. Data is not being saved — "
                   "check the connection/network and try again."
        )
    return mdb


def get_mongo_db():
    """
    Connects to MongoDB using PyMongo.
    """
    global _mongo_client, _mongo_db, _mongo_last_failed_at
    if _mongo_db is not None:
        return _mongo_db

    if _mongo_last_failed_at and (time.monotonic() - _mongo_last_failed_at) < _mongo_retry_cooldown_sec:
        return None

    try:
        client_kwargs = {"serverSelectionTimeoutMS": MONGO_TIMEOUT_MS}
        if MONGO_IS_REMOTE:
            # Atlas and other hosted clusters require TLS with a trusted CA bundle.
            client_kwargs["tls"] = True
            client_kwargs["tlsCAFile"] = certifi.where()

        _mongo_client = pymongo.MongoClient(MONGO_URI, **client_kwargs)
        _mongo_client.server_info()  # Verify connection
        _mongo_db = _mongo_client[MONGO_DB_NAME]

        # Indexes: per-user lookups are the hot path now that data is account-scoped.
        _mongo_db["users"].create_index([("email", pymongo.ASCENDING)], unique=True)
        _mongo_db["saved_records"].create_index([("created_at_dt", pymongo.DESCENDING)])
        _mongo_db["saved_records"].create_index([("user_id", pymongo.ASCENDING),
                                                 ("created_at_dt", pymongo.DESCENDING)])
        _mongo_db["saved_records"].create_index([("uid", pymongo.ASCENDING)])
        _mongo_db["sessions"].create_index([("user_id", pymongo.ASCENDING),
                                            ("updated_at_dt", pymongo.DESCENDING)])
        _mongo_db["mutual_funds"].create_index([("name", pymongo.ASCENDING)])
        _mongo_db["mutual_funds"].create_index([("sharpe", pymongo.DESCENDING)])
        _mongo_db["mutual_funds"].create_index([("category", pymongo.ASCENDING)])
        _mongo_db["mutual_funds"].create_index([("user_id", pymongo.ASCENDING)])

        logger.info("Connected to server MongoDB database %s", MONGO_DB_NAME)
        return _mongo_db
    except Exception as e:
        _mongo_client = None
        _mongo_db = None
        _mongo_last_failed_at = time.monotonic()
        logger.exception("MongoDB connection failed")
        return None


# ==========================================
# AUTHENTICATION
# Accounts so each user's data is separate and reachable from any device.
# ==========================================

SECRET_PATH = os.path.join(WORKSPACE_DIR, ".app_secret")
AUTH_COOKIE = "mfa_session"
TOKEN_TTL_DAYS = int(os.getenv("AUTH_TOKEN_TTL_DAYS", "30"))
# Sign-in only by default: accounts are created by an admin via `python main.py create-user`.
ALLOW_SIGNUP = os.getenv("ALLOW_SIGNUP", "0") == "1"


def _load_or_create_secret() -> str:
    """
    Signing key for auth tokens. Taken from APP_SECRET when set (the right way to do it
    in a hosted deployment); otherwise generated once and kept in a local file so tokens
    survive restarts instead of logging everyone out on each reload.
    """
    env_secret = os.getenv("APP_SECRET")
    if env_secret:
        return env_secret
    try:
        if os.path.exists(SECRET_PATH):
            with open(SECRET_PATH, "r", encoding="utf-8") as f:
                existing = f.read().strip()
                if existing:
                    return existing
        generated = secrets.token_urlsafe(48)
        with open(SECRET_PATH, "w", encoding="utf-8") as f:
            f.write(generated)
        return generated
    except Exception:
        # Last resort: ephemeral key (tokens won't survive a restart).
        return secrets.token_urlsafe(48)


APP_SECRET = _load_or_create_secret()


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def verify_password(password: str, password_hash: str) -> bool:
    try:
        return bcrypt.checkpw(password.encode("utf-8"), password_hash.encode("utf-8"))
    except Exception:
        return False


def create_token(user_id: str, email: str) -> str:
    payload = {
        "sub": user_id,
        "email": email,
        "exp": datetime.utcnow() + timedelta(days=TOKEN_TTL_DAYS),
        "iat": datetime.utcnow(),
    }
    return jwt.encode(payload, APP_SECRET, algorithm="HS256")


def decode_token(token: str) -> Optional[Dict[str, Any]]:
    try:
        return jwt.decode(token, APP_SECRET, algorithms=["HS256"])
    except Exception:
        return None


def find_user_by_email(email: str) -> Optional[Dict[str, Any]]:
    email = (email or "").strip().lower()
    if not email:
        return None

    mdb = require_mongo_db()
    doc = mdb["users"].find_one({"email": email})
    if doc:
        return {"uid": str(doc["_id"]), "email": doc["email"],
                "password_hash": doc["password_hash"], "name": doc.get("name", "")}
    return None


def find_user_by_id(user_id: str) -> Optional[Dict[str, Any]]:
    if not user_id:
        return None

    mdb = require_mongo_db()
    doc = mdb["users"].find_one({"_id": user_id})
    if doc:
        return {"uid": str(doc["_id"]), "email": doc["email"],
                "password_hash": doc["password_hash"], "name": doc.get("name", "")}
    return None


def create_user(email: str, password: str, name: str) -> Dict[str, Any]:
    email = email.strip().lower()
    user_id = str(uuid.uuid4())
    now_dt = datetime.now()
    created_at = now_dt.strftime("%d %b %Y, %I:%M %p")
    password_hash = hash_password(password)

    mdb = require_mongo_db()
    try:
        mdb["users"].insert_one({
            "_id": user_id, "email": email, "password_hash": password_hash,
            "name": name, "created_at": created_at, "created_at_dt": now_dt,
        })
    except pymongo.errors.DuplicateKeyError:
        raise HTTPException(status_code=409, detail="An account with that email already exists.")

    return {"uid": user_id, "email": email, "name": name, "created_at": created_at}


def count_users() -> int:
    """Used by the sign-in screen to tell 'wrong password' apart from 'no accounts yet'."""
    try:
        mdb = get_mongo_db()
        if mdb is not None:
            return mdb["users"].count_documents({})
        return 0
    except Exception:
        return -1  # unknown (database unreachable)


async def get_current_user(request: Request) -> Dict[str, Any]:
    """FastAPI dependency: every data endpoint goes through this."""
    token = request.cookies.get(AUTH_COOKIE)
    if not token:
        raise HTTPException(status_code=401, detail="Not signed in.")
    payload = decode_token(token)
    if not payload or not payload.get("sub"):
        raise HTTPException(status_code=401, detail="Your session has expired. Please sign in again.")
    user = find_user_by_id(payload["sub"])
    if not user:
        raise HTTPException(status_code=401, detail="Account no longer exists.")
    return user



# ==========================================
# ==========================================
# FASTAPI ENDPOINTS
# ==========================================

ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".csv"}

@app.get("/", response_class=HTMLResponse)
async def get_index(request: Request):
    """
    Serves the Mutual Fund & Dynamic Excel Analytics dashboard.

    Rendered through Jinja2 (not read as a flat file) so every asset and API
    URL is built from the request's root_path. That is what lets one build run
    unchanged at a domain root, behind a sub-path reverse proxy, or on any
    cloud host without rebaking URLs into the HTML.
    """
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            # "" at a domain root; "/mfa" when a proxy mounts us at a sub-path.
            "root_path": request.scope.get("root_path", "").rstrip("/"),
            "asset_version": ASSET_VERSION,
        },
        headers={
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "Expires": "0",
        },
    )


@app.get("/favicon.ico", include_in_schema=False)
async def get_favicon():
    logo_path = os.path.join(STATIC_IMAGES_DIR, "arp-logo.png")
    if os.path.exists(logo_path):
        return FileResponse(logo_path, media_type="image/png")
    return Response(status_code=204)


@app.get("/api/health")
async def health():
    return {"status": "ok", "app": "Mutual Fund Analytics Platform", "version": "2.4.0"}


@app.post("/api/upload")
async def upload_files(
    files: Optional[List[UploadFile]] = File(None),
    risk_file: Optional[UploadFile] = File(None),
    rolling_file: Optional[UploadFile] = File(None),
    benchmark: Optional[float] = Form(None),
    user: Dict[str, Any] = Depends(get_current_user)
):
    """
    Strictly verifies and parses uploaded Excel/CSV spreadsheets only.
    Guarantees 100% of the returned data originates directly from the uploaded file(s).
    """
    all_uploads = []
    if files:
        all_uploads.extend(files)
    if risk_file:
        all_uploads.append(risk_file)
    if rolling_file:
        all_uploads.append(rolling_file)
        
    if not all_uploads:
        raise HTTPException(
            status_code=400, 
            detail="No file uploaded. Please upload a valid Excel spreadsheet (.xlsx, .xls, .xlsm) or .csv file."
        )
        
    try:
        parsed_files = []
        
        for file in all_uploads:
            filename = file.filename or "uploaded_spreadsheet.xlsx"
            ext = os.path.splitext(filename)[1].lower()
            
            # 1. Strict File Type Check
            if ext not in ALLOWED_EXTENSIONS:
                raise HTTPException(
                    status_code=400,
                    detail=f"Invalid file format for '{filename}'. Only Excel spreadsheets (.xlsx, .xls, .xlsm) or .csv files are supported."
                )
                
            content = await file.read()
            
            # 2. Strict File Content & Size Check
            if not content or len(content) == 0:
                raise HTTPException(
                    status_code=400,
                    detail=f"The uploaded file '{filename}' is empty (0 bytes). Please upload a valid Excel spreadsheet."
                )
                
            try:
                parsed = smart_read_sheet(content, filename)
            except Exception as read_err:
                raise HTTPException(
                    status_code=400,
                    detail=f"Failed to parse Excel file '{filename}'. Ensure it is a valid spreadsheet. Error: {str(read_err)}"
                )
                
            if parsed["df"].empty or len(parsed["df"]) == 0:
                raise HTTPException(
                    status_code=400,
                    detail=f"No data rows could be found in Excel file '{filename}'. Please verify the sheet structure."
                )
                
            col_map = map_fund_columns(parsed["df"])
            analysis = analyze_fund_dataset(parsed["df"], col_map, custom_benchmark=benchmark)
            
            has_rolling = any(col_map.get(k) is not None for k in ["rolling_1y", "rolling_2y", "rolling_3y", "rolling_5y"])
            has_sharpe = col_map.get("sharpe") is not None
            
            file_meta = {
                "filename": filename,
                "filesize_bytes": len(content),
                "filesize_kb": round(len(content) / 1024, 1),
                "sheets": parsed["sheets"],
                "active_sheet": parsed["active_sheet"],
                "header_row_index": parsed["header_row"] + 1,
                "columns": parsed["columns"],
                "mapping": col_map,
                "has_rolling": has_rolling,
                "has_sharpe": has_sharpe,
                "analysis": analysis,
                "row_count": len(parsed["df"])
            }
            parsed_files.append(file_meta)

        if not parsed_files:
            raise HTTPException(status_code=400, detail="Uploaded file(s) could not be processed.")

        # Smart merge if multiple files uploaded (e.g. Risk-Ratios + Rolling-Returns)
        if len(parsed_files) == 1:
            main_result = parsed_files[0]["analysis"]
            if parsed_files[0]["has_rolling"]:
                main_result["rolling_funds"] = parsed_files[0]["analysis"]["funds"]
            return JSONResponse(content={
                "status": "success",
                "source": "uploaded_excel_direct",
                "verified_from_excel": True,
                "file_count": 1,
                "files_meta": [
                    {
                        "filename": f["filename"],
                        "filesize_kb": f["filesize_kb"],
                        "active_sheet": f["active_sheet"],
                        "sheets": f["sheets"],
                        "row_count": f["row_count"],
                        "columns": f["columns"]
                    } for f in parsed_files
                ],
                "data": main_result
            })
        else:
            risk_item = next((f for f in parsed_files if f["has_sharpe"] and not f["has_rolling"]), parsed_files[0])
            rolling_item = next((f for f in parsed_files if f["has_rolling"]), parsed_files[1] if len(parsed_files) > 1 else None)
            
            if rolling_item and risk_item:
                combined = merge_risk_and_rolling(risk_item["analysis"], rolling_item["analysis"], custom_benchmark=benchmark)
                return JSONResponse(content={
                    "status": "success",
                    "source": "uploaded_excel_direct",
                    "verified_from_excel": True,
                    "file_count": len(parsed_files),
                    "files_meta": [
                        {
                            "filename": f["filename"],
                            "filesize_kb": f["filesize_kb"],
                            "active_sheet": f["active_sheet"],
                            "sheets": f["sheets"],
                            "row_count": f["row_count"],
                            "columns": f["columns"]
                        } for f in parsed_files
                    ],
                    "data": combined
                })
            else:
                return JSONResponse(content={
                    "status": "success",
                    "source": "uploaded_excel_direct",
                    "verified_from_excel": True,
                    "file_count": len(parsed_files),
                    "files_meta": [
                        {
                            "filename": f["filename"],
                            "filesize_kb": f["filesize_kb"],
                            "active_sheet": f["active_sheet"],
                            "sheets": f["sheets"],
                            "row_count": f["row_count"],
                            "columns": f["columns"]
                        } for f in parsed_files
                    ],
                    "data": parsed_files[0]["analysis"]
                })
                
    except HTTPException:
        raise
    except Exception as e:
        logger.exception("Error reading Excel spreadsheet")
        raise HTTPException(status_code=500, detail="Error reading Excel spreadsheet. Check server logs for details.")


@app.post("/api/upload-rolling-category")
async def upload_rolling_category(
    rolling_file: UploadFile = File(...),
    current_funds_json: str = Form(...),
    benchmark: Optional[float] = Form(None),
    user: Dict[str, Any] = Depends(get_current_user)
):
    """
    Sequentially merges category-wise rolling returns spreadsheets into the active
    mutual fund dataset without overwriting prior rolling metrics or other categories.
    """
    try:
        current_funds = json.loads(current_funds_json)
        if not isinstance(current_funds, list):
            raise HTTPException(status_code=400, detail="Invalid current funds list format.")
            
        filename = rolling_file.filename or "rolling_returns.xlsx"
        ext = os.path.splitext(filename)[1].lower()
        if ext not in ALLOWED_EXTENSIONS:
            raise HTTPException(
                status_code=400,
                detail=f"Invalid file format for '{filename}'. Only Excel spreadsheets (.xlsx, .xls, .xlsm) or .csv are supported."
            )
            
        content = await rolling_file.read()
        if not content or len(content) == 0:
            raise HTTPException(status_code=400, detail=f"The uploaded file '{filename}' is empty.")
            
        try:
            parsed = smart_read_sheet(content, filename)
        except Exception as read_err:
            raise HTTPException(status_code=400, detail=f"Failed to read sheet '{filename}': {str(read_err)}")
            
        if parsed["df"].empty or len(parsed["df"]) == 0:
            raise HTTPException(status_code=400, detail=f"No data rows found in '{filename}'.")
            
        col_map = map_fund_columns(parsed["df"])
        rolling_analysis = analyze_fund_dataset(parsed["df"], col_map, custom_benchmark=benchmark)
        incoming_funds = rolling_analysis["funds"]
        
        # Build lookup index for incoming rolling returns (exact key + signature)
        matcher = build_rolling_matcher(incoming_funds)

        merged_funds = []
        matched_count = 0
        signature_matched_count = 0
        matched_categories = set()
        ambiguous_names = []
        consumed = set()

        # Merge into existing fund universe. Exact normalized-name match wins, then a
        # signature match (tolerant of "Reg"/"Regular", "Gr"/"Growth", a missing "Fund",
        # punctuation, token order) used only when exactly one incoming record carries
        # that signature — see scheme_signature() for why that guard matters.
        for f in current_funds:
            item = dict(f)

            result = match_rolling_record(matcher, f.get("name"), consumed_ids=consumed)
            matched = result["record"]
            if result["ambiguous"]:
                ambiguous_names.append(f.get("name"))

            if matched:
                matched_count += 1
                if result.get("how") != "exact":
                    signature_matched_count += 1
                consumed.add(id(matched))
                if matched.get("rolling_1y") is not None: item["rolling_1y"] = matched["rolling_1y"]
                if matched.get("rolling_2y") is not None: item["rolling_2y"] = matched["rolling_2y"]
                if matched.get("rolling_3y") is not None: item["rolling_3y"] = matched["rolling_3y"]
                if matched.get("rolling_5y") is not None: item["rolling_5y"] = matched["rolling_5y"]
                # Derive the mean from every retained horizon, including earlier uploads.
                horizons = [item.get(key) for key in
                            ("rolling_1y", "rolling_2y", "rolling_3y", "rolling_5y")]
                values = [v for v in horizons if isinstance(v, (int, float))]
                item["rolling_avg"] = round(sum(values) / len(values), 2) if values else None

                # Update supplemental risk ratios if not already present
                if matched.get("std_dev") is not None and item.get("std_dev") is None:
                    item["std_dev"] = matched["std_dev"]
                if matched.get("beta") is not None and item.get("beta") is None:
                    item["beta"] = matched["beta"]
                if matched.get("alpha") is not None and item.get("alpha") is None:
                    item["alpha"] = matched["alpha"]
                if matched.get("sortino") is not None and item.get("sortino") is None:
                    item["sortino"] = matched["sortino"]
                if matched.get("inception_date") and not item.get("inception_date"):
                    item["inception_date"] = matched["inception_date"]

                cat = item.get("category") or matched.get("category")
                if cat:
                    matched_categories.add(cat)

            merged_funds.append(item)

        # Keep funds that exist only in this rolling-returns upload — e.g. a category (such
        # as Hybrid or Solution Oriented) that isn't present in the current universe at all —
        # instead of silently discarding that data.
        rolling_only_funds = [rf for rf in incoming_funds if id(rf) not in consumed]
        for rf in rolling_only_funds:
            merged_funds.append(dict(rf))
            cat = rf.get("category")
            if cat:
                matched_categories.add(cat)

        # Re-run full universe calculation on merged_funds
        recalculated = compute_records_analysis(merged_funds, custom_benchmark=benchmark, col_map=col_map)
        
        # Calculate category rolling coverage
        categories_rolling_stats = {}
        for cat in recalculated["categories"]:
            cat_funds = [f for f in recalculated["funds"] if f.get("category") == cat]
            with_rolling = sum(1 for f in cat_funds if f.get("rolling_1y") is not None or f.get("rolling_avg") is not None)
            categories_rolling_stats[cat] = {
                "total": len(cat_funds),
                "with_rolling": with_rolling,
                "coverage_pct": round((with_rolling / len(cat_funds)) * 100, 1) if cat_funds else 0
            }
            
        total_with_rolling = sum(1 for f in recalculated["funds"] if f.get("rolling_1y") is not None or f.get("rolling_avg") is not None)
        recalculated["summary"]["rolling_matched_count"] = total_with_rolling
        recalculated["summary"]["rolling_only_count"] = len(rolling_only_funds)

        message = f"Successfully merged rolling returns for {matched_count} funds across {len(matched_categories)} categories."
        if rolling_only_funds:
            message += f" Added {len(rolling_only_funds)} fund(s) found only in this sheet (not present in the existing universe)."

        return JSONResponse(content={
            "status": "success",
            "message": message,
            "filename": filename,
            "row_count": len(incoming_funds),
            "matched_count": matched_count,
            "rolling_only_count": len(rolling_only_funds),
            "total_with_rolling": total_with_rolling,
            "matched_categories": sorted(list(matched_categories)),
            "categories_rolling_stats": categories_rolling_stats,
            "rolling_funds": incoming_funds,
            "match_report": {
                "rolling_rows": len(incoming_funds),
                "matched": matched_count,
                "matched_by_signature": signature_matched_count,
                "unmatched": len(rolling_only_funds),
                "unmatched_names": [rf.get("name") for rf in rolling_only_funds][:50],
                "ambiguous_names": ambiguous_names[:50],
            },
            "data": recalculated
        })
    except HTTPException:
        raise
    except Exception as e:
        logger.exception("Failed to merge category rolling returns")
        raise HTTPException(status_code=500, detail="Failed to merge category rolling returns. Check server logs for details.")


@app.post("/api/recalculate")
async def recalculate_dataset(request: Request):
    """
    Recalculates universe statistics, ranks, and benchmark outperformance
    when new records are added or benchmarks are updated.
    """
    try:
        body = await request.json()
        funds = body.get("funds", [])
        benchmark = body.get("benchmark")
        if not funds:
            raise HTTPException(status_code=400, detail="No funds provided to recalculate.")
            
        recalculated = compute_records_analysis(funds, custom_benchmark=benchmark)
        return JSONResponse(content={
            "status": "success",
            "data": recalculated
        })
    except HTTPException:
        raise
    except Exception as e:
        logger.exception("Recalculation failed")
        raise HTTPException(status_code=500, detail="Recalculation failed. Check server logs for details.")


@app.post("/api/analyze-custom")
async def analyze_custom_sheet(
    file: UploadFile = File(...),
    sheet_name: Optional[str] = Form(None),
    name_col: Optional[str] = Form(None),
    cat_col: Optional[str] = Form(None),
    aum_col: Optional[str] = Form(None),
    sharpe_col: Optional[str] = Form(None),
    benchmark: Optional[float] = Form(None)
):
    """
    Allows user to customize column mappings and target metric thresholds for ANY custom dynamic sheet.
    """
    try:
        content = await file.read()
        parsed = smart_read_sheet(content, file.filename, sheet_name=sheet_name)
        
        # Override column mapping if provided
        col_map = map_fund_columns(parsed["df"])
        if name_col: col_map["name"] = name_col
        if cat_col: col_map["category"] = cat_col
        if aum_col: col_map["aum"] = aum_col
        if sharpe_col: col_map["sharpe"] = sharpe_col
        
        analysis = analyze_fund_dataset(parsed["df"], col_map, custom_benchmark=benchmark)
        
        return JSONResponse(content={
            "status": "success",
            "filename": file.filename,
            "active_sheet": parsed["active_sheet"],
            "sheets": parsed["sheets"],
            "columns": parsed["columns"],
            "mapping": col_map,
            "data": analysis
        })
    except Exception as e:
        logger.exception("Custom analysis failed")
        raise HTTPException(status_code=500, detail="Custom analysis failed. Check server logs for details.")


@app.post("/api/export")
async def export_data(data: Dict[str, Any], user: Dict[str, Any] = Depends(get_current_user)):
    """
    Exports enriched mutual fund dataset to Excel (.xlsx) with calculated rank,
    Sharpe comparison, delta, and formatting.
    """
    try:
        funds = data.get("funds", [])
        if not funds:
            raise HTTPException(status_code=400, detail="No fund data to export.")
            
        export_rows = []
        for f in funds:
            export_rows.append({
                "Rank": f.get("rank") or "-",
                "Scheme Name": f.get("name"),
                "Category": f.get("category"),
                "AUM (₹ Crore)": f.get("aum"),
                "Sharpe Ratio": f.get("sharpe"),
                "Performance Status": "≥ Average Sharpe" if f.get("above_avg") else ("< Average Sharpe" if f.get("sharpe_valid") else "N/A"),
                "Sharpe Delta vs Avg": f.get("sharpe_diff"),
                "Information Ratio": f.get("info_ratio"),
                "Treynor Ratio": f.get("treynor"),
                "Standard Deviation (%)": f.get("std_dev"),
                "Inception Date": f.get("inception_date"),
                "1Y Rolling (%)": f.get("rolling_1y"),
                "2Y Rolling (%)": f.get("rolling_2y"),
                "3Y Rolling (%)": f.get("rolling_3y"),
                "5Y Rolling (%)": f.get("rolling_5y"),
                "Average Rolling (%)": f.get("rolling_avg")
            })
            
        df_export = pd.DataFrame(export_rows)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_export.to_excel(writer, sheet_name="Mutual_Funds_Analysis", index=False)
            
            # Add summary sheet
            summary = data.get("summary", {})
            summary_rows = [
                {"Metric": "Total Funds Analyzed", "Value": summary.get("total_funds")},
                {"Metric": "Funds with Valid Sharpe Ratio", "Value": summary.get("valid_sharpe_count")},
                {"Metric": "Total Average Sharpe Ratio", "Value": summary.get("avg_sharpe")},
                {"Metric": "Funds ≥ Average Sharpe", "Value": summary.get("above_avg_count")},
                {"Metric": "Funds < Average Sharpe", "Value": summary.get("below_avg_count")},
                {"Metric": "Total Portfolio AUM (₹ Cr)", "Value": summary.get("total_aum")},
                {"Metric": "Top Outperformer", "Value": summary.get("top_performer", {}).get("name", "N/A") if summary.get("top_performer") else "N/A"}
            ]
            pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Executive_Summary", index=False)
            
        output.seek(0)
        
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=mutual_fund_analytics_report.xlsx"}
        )
    except Exception as e:
        logger.exception("Export generation failed")
        raise HTTPException(status_code=500, detail="Export generation failed. Check server logs for details.")


@app.post("/api/export-basket")
async def export_selected_basket(request: Request, user: Dict[str, Any] = Depends(get_current_user)):
    """
    Exports only the user-screened / selected fund basket to an executive Excel spreadsheet.
    """
    try:
        body = await request.json()
        funds = body.get("funds", [])
        basket_name = body.get("basket_name", "Curated_Mutual_Fund_Basket")
        filter_summary = body.get("filter_summary", "Custom Multi-Ratio Filter")
        benchmark = body.get("benchmark", 0.62)

        if not funds:
            raise HTTPException(status_code=400, detail="No funds provided in the selected basket to export.")

        export_rows = []
        total_aum = 0.0
        valid_sharpes = []

        for idx, f in enumerate(funds, start=1):
            sharpe_val = f.get("sharpe")
            if isinstance(sharpe_val, (int, float)):
                valid_sharpes.append(sharpe_val)
                status_str = "Outperformer (≥ Avg)" if sharpe_val >= benchmark else "Below Benchmark (< Avg)"
            else:
                status_str = "Track Record < 3Y (N/A)"

            aum_val = f.get("aum")
            if isinstance(aum_val, (int, float)):
                total_aum += aum_val

            export_rows.append({
                "Basket S.No.": idx,
                "Fund Rank": f.get("rank") or "N/A",
                "Scheme Name": f.get("name", "Unknown"),
                "Category": f.get("category", "General"),
                "AUM (₹ Cr)": round(aum_val, 2) if isinstance(aum_val, (int, float)) else "N/A",
                "Sharpe Ratio": round(sharpe_val, 2) if isinstance(sharpe_val, (int, float)) else "N/A",
                "Sharpe vs Benchmark": status_str,
                "Std Deviation (Volatility %)": round(f["std_dev"], 2) if isinstance(f.get("std_dev"), (int, float)) else "N/A",
                "Treynor Ratio": round(f["treynor"], 2) if isinstance(f.get("treynor"), (int, float)) else "N/A",
                "Information Ratio": round(f["info_ratio"], 2) if isinstance(f.get("info_ratio"), (int, float)) else "N/A",
                "1Y Rolling (%)": round(f["rolling_1y"], 2) if isinstance(f.get("rolling_1y"), (int, float)) else "N/A",
                "2Y Rolling (%)": round(f["rolling_2y"], 2) if isinstance(f.get("rolling_2y"), (int, float)) else "N/A",
                "3Y Rolling (%)": round(f["rolling_3y"], 2) if isinstance(f.get("rolling_3y"), (int, float)) else "N/A",
                "5Y Rolling (%)": round(f["rolling_5y"], 2) if isinstance(f.get("rolling_5y"), (int, float)) else "N/A",
                "Rolling Avg CAGR (%)": round(f["rolling_avg"], 2) if isinstance(f.get("rolling_avg"), (int, float)) else "N/A",
                "Inception Date": f.get("inception_date") or "N/A"
            })

        df_basket = pd.DataFrame(export_rows)

        # Summary Sheet
        avg_basket_sharpe = round(float(np.mean(valid_sharpes)), 2) if valid_sharpes else "N/A"
        above_count = sum(1 for s in valid_sharpes if s >= benchmark)

        summary_rows = [
            {"Metric": "Basket Title", "Value": basket_name},
            {"Metric": "Total Funds in Basket", "Value": len(funds)},
            {"Metric": "Funds with Valid Sharpe Ratio", "Value": len(valid_sharpes)},
            {"Metric": "Basket Average Sharpe Ratio", "Value": avg_basket_sharpe},
            {"Metric": "Active Benchmark Sharpe Threshold", "Value": round(benchmark, 2)},
            {"Metric": "Outperforming Funds in Basket", "Value": f"{above_count} ({round(above_count / len(valid_sharpes) * 100, 1)}%)" if valid_sharpes else "0%"},
            {"Metric": "Combined Basket AUM (₹ Cr)", "Value": round(total_aum, 2)},
            {"Metric": "Filter Criteria Applied", "Value": filter_summary},
            {"Metric": "Generated On", "Value": datetime.now().strftime("%d %b %Y, %I:%M %p")}
        ]
        df_summary = pd.DataFrame(summary_rows)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_basket.to_excel(writer, sheet_name="Curated_Fund_Basket", index=False)
            df_summary.to_excel(writer, sheet_name="Basket_Executive_Summary", index=False)

        output.seek(0)
        safe_filename = re.sub(r'[^a-zA-Z0-9_-]', '_', basket_name.lower()) + ".xlsx"

        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={safe_filename}"}
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.exception("Basket export generation failed")
        raise HTTPException(status_code=500, detail="Basket export generation failed. Check server logs for details.")

def init_db():
    # Initialize MongoDB connection and indexes
    get_mongo_db()

init_db()


# ==========================================
# AUTH ENDPOINTS
# ==========================================

def _set_auth_cookie(response: JSONResponse, token: str) -> JSONResponse:
    # secure=True requires HTTPS; enable it via COOKIE_SECURE=1 once served over TLS
    # (mandatory for internet deployment — passwords/tokens must not cross plain HTTP).
    response.set_cookie(
        key=AUTH_COOKIE,
        value=token,
        httponly=True,                 # not readable by JavaScript
        samesite="lax",
        secure=os.getenv("COOKIE_SECURE", "0") == "1",
        max_age=TOKEN_TTL_DAYS * 24 * 3600,
        path="/",
    )
    return response


@app.post("/api/auth/register")
async def register(request: Request):
    """
    Account creation. Public self-signup is OFF by default — the app shows a sign-in
    screen only, so accounts are provisioned by an admin:
        python main.py create-user <email> <password> ["Full Name"]
    Set ALLOW_SIGNUP=1 to re-open public registration.
    """
    if not ALLOW_SIGNUP:
        raise HTTPException(
            status_code=403,
            detail="Self sign-up is disabled. Ask an administrator to create your account."
        )

    body = await request.json()
    email = (body.get("email") or "").strip().lower()
    password = body.get("password") or ""
    name = (body.get("name") or "").strip() or email.split("@")[0]

    if "@" not in email or "." not in email.split("@")[-1]:
        raise HTTPException(status_code=400, detail="Enter a valid email address.")
    if len(password) < 8:
        raise HTTPException(status_code=400, detail="Password must be at least 8 characters.")

    if find_user_by_email(email):
        raise HTTPException(status_code=409, detail="An account with that email already exists.")

    user = create_user(email, password, name)
    token = create_token(user["uid"], user["email"])
    resp = JSONResponse(content={
        "status": "success",
        "user": {"uid": user["uid"], "email": user["email"], "name": user["name"]},
    })
    return _set_auth_cookie(resp, token)


@app.post("/api/auth/login")
async def login(request: Request):
    body = await request.json()
    email = (body.get("email") or "").strip().lower()
    password = body.get("password") or ""

    user = find_user_by_email(email)
    # Same message for unknown email and wrong password — don't reveal which accounts exist.
    if not user or not verify_password(password, user["password_hash"]):
        raise HTTPException(status_code=401, detail="Incorrect email or password.")

    token = create_token(user["uid"], user["email"])
    resp = JSONResponse(content={
        "status": "success",
        "user": {"uid": user["uid"], "email": user["email"], "name": user.get("name", "")},
    })
    return _set_auth_cookie(resp, token)


@app.post("/api/auth/logout")
async def logout():
    resp = JSONResponse(content={"status": "success"})
    resp.delete_cookie(AUTH_COOKIE, path="/")
    return resp


@app.get("/api/auth/me")
async def whoami(request: Request):
    """Used by the UI on load to decide between the login screen and the dashboard."""
    token = request.cookies.get(AUTH_COOKIE)
    payload = decode_token(token) if token else None
    if not payload or not payload.get("sub"):
        return {"status": "anonymous", "allow_signup": ALLOW_SIGNUP, "user_count": count_users()}
    try:
        user = find_user_by_id(payload["sub"])
    except HTTPException:
        # Database unreachable — report it rather than bouncing the user to the login screen
        return {"status": "error", "detail": "Database unreachable."}
    if not user:
        return {"status": "anonymous", "allow_signup": ALLOW_SIGNUP, "user_count": count_users()}
    return {
        "status": "success",
        "user": {"uid": user["uid"], "email": user["email"], "name": user.get("name", "")},
    }


# ==========================================
# LIVE SESSION PERSISTENCE (server-side, UUID-keyed, per user)
# ==========================================

@app.put("/api/session")
async def save_session(request: Request, user: Dict[str, Any] = Depends(get_current_user)):
    """
    Upserts the signed-in user's live working session (active dataset + UI state).
    Stored in the database against their account — never in browser storage — so signing
    in on any device brings the same working state back.
    """
    try:
        body = await request.json()
        uid = (body.get("uid") or "").strip() or str(uuid.uuid4())
        state = body.get("state")
        if state is None:
            raise HTTPException(status_code=400, detail="No session state provided.")

        now_dt = datetime.now()
        updated_at = now_dt.strftime("%d %b %Y, %I:%M %p")
        state_json = json.dumps(state)
        summary = (state.get("data") or {}).get("summary") or {}
        total_funds = int(summary.get("total_funds", 0) or 0)
        files_summary = state.get("files_summary", "") or ""

        mdb = require_mongo_db()
        mdb["sessions"].update_one(
            {"_id": uid, "user_id": user["uid"]},
            {"$set": {
                "user_id": user["uid"],
                "updated_at": updated_at,
                "updated_at_dt": now_dt,
                "total_funds": total_funds,
                "files_summary": files_summary,
                "state": state,
            },
             "$setOnInsert": {"created_at": updated_at}},
            upsert=True
        )

        return {"status": "success", "uid": uid, "updated_at": updated_at, "total_funds": total_funds}
    except HTTPException:
        raise
    except Exception as e:
        logger.exception("Failed to save session")
        raise HTTPException(status_code=500, detail="Failed to save session. Check server logs for details.")


@app.get("/api/session/latest")
async def get_latest_session(user: Dict[str, Any] = Depends(get_current_user)):
    """Most recently updated session for the signed-in user, or 'empty' when none."""
    mdb = require_mongo_db()
    doc = mdb["sessions"].find_one({"user_id": user["uid"]},
                                   sort=[("updated_at_dt", pymongo.DESCENDING)])
    if not doc:
        return {"status": "empty"}
    return {
        "status": "success",
        "uid": str(doc["_id"]),
        "updated_at": doc.get("updated_at", ""),
        "total_funds": doc.get("total_funds", 0),
        "state": doc.get("state"),
    }


@app.get("/api/session/{uid}")
async def get_session(uid: str, user: Dict[str, Any] = Depends(get_current_user)):
    """Returns one of the signed-in user's sessions by its UUID."""
    mdb = require_mongo_db()
    doc = mdb["sessions"].find_one({"_id": uid, "user_id": user["uid"]})
    if not doc:
        raise HTTPException(status_code=404, detail="Session not found.")
    return {
        "status": "success",
        "uid": str(doc["_id"]),
        "updated_at": doc.get("updated_at", ""),
        "total_funds": doc.get("total_funds", 0),
        "state": doc.get("state"),
    }


@app.delete("/api/session/{uid}")
async def delete_session(uid: str, user: Dict[str, Any] = Depends(get_current_user)):
    """Clears one of the signed-in user's sessions (used by 'start fresh')."""
    mdb = require_mongo_db()
    removed = mdb["sessions"].delete_one({"_id": uid, "user_id": user["uid"]}).deleted_count
    return {"status": "success", "removed": removed}


@app.get("/api/db/status")
async def get_db_status():
    """Report availability without exposing credentials, connection strings, or user counts."""
    mdb = get_mongo_db()
    if mdb is not None:
        try:
            mdb.command("ping")
            return {"status": "connected", "engine": "MongoDB (PyMongo)", "storage": "server"}
        except Exception:
            logger.exception("MongoDB status check failed")
    return {
        "status": "disconnected",
        "engine": "MongoDB (PyMongo)",
        "storage": "server",
        "message": "Server database is unreachable. Changes cannot be saved until it reconnects.",
    }


@app.get("/api/records")
async def list_saved_records(user: Dict[str, Any] = Depends(get_current_user)):
    """
    Returns the signed-in user's saved records from MongoDB.
    """
    mdb = require_mongo_db()
    docs = list(mdb["saved_records"].find(
        {"user_id": user["uid"]},
        {"data": 0, "funds": 0}
    ).sort("created_at_dt", pymongo.DESCENDING))

    records = []
    for d in docs:
        records.append({
            "id": d.get("uid") or str(d["_id"]),
            "name": d.get("name", "Untitled Analysis"),
            "created_at": d.get("created_at", ""),
            "total_funds": d.get("total_funds", 0),
            "valid_sharpe_count": d.get("valid_sharpe_count", 0),
            "avg_sharpe": d.get("avg_sharpe", 0.0),
            "above_avg_count": d.get("above_avg_count", 0),
            "below_avg_count": d.get("below_avg_count", 0),
            "outperformance_rate": d.get("outperformance_rate", 0.0),
            "total_aum": d.get("total_aum", 0.0),
            "top_performer_name": d.get("top_performer_name", ""),
            "top_performer_sharpe": d.get("top_performer_sharpe", 0.0),
            "files_summary": d.get("files_summary", ""),
            "storage_engine": "MongoDB"
        })
    return {"status": "success", "engine": "MongoDB", "count": len(records), "records": records}


@app.post("/api/records/save")
async def save_current_record(request: Request, user: Dict[str, Any] = Depends(get_current_user)):
    """
    Saves an analysed dataset against the signed-in user's account in MongoDB.
    """
    try:
        body = await request.json()
        name = body.get("name") or f"Analysis_{datetime.now().strftime('%Y-%m-%d_%H%M')}"
        data = body.get("data")
        files_summary = body.get("files_summary", "")

        if not data or not data.get("funds"):
            raise HTTPException(status_code=400, detail="No fund analysis data provided to save.")

        summary = data.get("summary", {})
        top_perf = summary.get("top_performer") or {}
        now_dt = datetime.now()
        created_at_str = now_dt.strftime("%d %b %Y, %I:%M %p")
        record_uid = str(uuid.uuid4())
        mdb = require_mongo_db()

        record_doc = {
            "uid": record_uid,
            "user_id": user["uid"],
            "name": name,
            "created_at": created_at_str,
            "created_at_dt": now_dt,
            "total_funds": int(summary.get("total_funds", 0)),
            "valid_sharpe_count": int(summary.get("valid_sharpe_count", 0)),
            "avg_sharpe": float(summary.get("avg_sharpe", 0.0)),
            "above_avg_count": int(summary.get("above_avg_count", 0)),
            "below_avg_count": int(summary.get("below_avg_count", 0)),
            "outperformance_rate": float(summary.get("outperformance_rate", 0.0)),
            "total_aum": float(summary.get("total_aum", 0.0)),
            "top_performer_name": str(top_perf.get("name", "")),
            "top_performer_sharpe": float(top_perf.get("sharpe") or 0.0),
            "files_summary": files_summary,
            "data": data
        }
        mdb["saved_records"].insert_one(record_doc)

        # Also insert individual funds into `mutual_funds` collection for querying in Compass
        funds_to_insert = []
        for f in data.get("funds", []):
            fund_doc = dict(f)
            fund_doc["snapshot_id"] = record_uid
            fund_doc["user_id"] = user["uid"]
            fund_doc["snapshot_name"] = name
            fund_doc["snapshot_date"] = created_at_str
            funds_to_insert.append(fund_doc)

        if funds_to_insert:
            mdb["mutual_funds"].insert_many(funds_to_insert)

        return {
            "status": "success",
            "message": f"Record '{name}' saved successfully.",
            "record_id": record_uid,
            "uid": record_uid,
            "name": name,
            "created_at": created_at_str,
            "engine": "MongoDB"
        }
    except HTTPException:
        raise
    except Exception as e:
        logger.exception("Failed to save record")
        raise HTTPException(status_code=500, detail="Failed to save record. Check server logs for details.")


@app.get("/api/records/{record_id}")
async def get_saved_record(record_id: str, user: Dict[str, Any] = Depends(get_current_user)):
    """
    Retrieves the full payload of one of the signed-in user's saved records from MongoDB.
    """
    mdb = require_mongo_db()
    doc = mdb["saved_records"].find_one({"uid": record_id, "user_id": user["uid"]})
    if not doc and ObjectId.is_valid(record_id):
        doc = mdb["saved_records"].find_one({"_id": ObjectId(record_id), "user_id": user["uid"]})
    if not doc:
        raise HTTPException(status_code=404, detail="Saved record not found.")

    meta = dict(doc)
    meta["id"] = meta.get("uid") or str(meta["_id"])
    data_payload = meta.pop("data", None)
    meta.pop("_id", None)
    return {
        "status": "success",
        "engine": "MongoDB",
        "record_meta": meta,
        "data": data_payload
    }


@app.delete("/api/records/{record_id}")
async def delete_saved_record(record_id: str, user: Dict[str, Any] = Depends(get_current_user)):
    """
    Deletes one of the signed-in user's saved records from MongoDB.
    """
    if not record_id or not record_id.strip():
        raise HTTPException(status_code=400, detail="A record id is required.")

    mdb = require_mongo_db()
    res = mdb["saved_records"].delete_one({"uid": record_id, "user_id": user["uid"]})
    if res.deleted_count == 0 and ObjectId.is_valid(record_id):
        res = mdb["saved_records"].delete_one({"_id": ObjectId(record_id), "user_id": user["uid"]})

    if res.deleted_count > 0:
        mdb["mutual_funds"].delete_many({"snapshot_id": record_id, "user_id": user["uid"]})
        return {"status": "success", "message": "Record deleted successfully."}

    raise HTTPException(status_code=404, detail="Record not found to delete.")


def find_available_port(start_port: int = 8000, max_attempts: int = 50) -> int:
    import socket
    for port in range(start_port, start_port + max_attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    return start_port


def _cli_create_user(argv: List[str]) -> int:
    """
    `python main.py create-user <email> <password> ["Full Name"]`

    The app shows a sign-in screen only (no public signup), so this is how accounts get
    created. Password can be omitted to be prompted for it without it landing in shell
    history.
    """
    import getpass

    if not argv:
        print("Usage: python main.py create-user <email> [password] [\"Full Name\"]")
        return 2

    email = argv[0].strip().lower()
    password = argv[1] if len(argv) > 1 else ""
    name = argv[2] if len(argv) > 2 else ""

    if "@" not in email or "." not in email.split("@")[-1]:
        print(f"✗ '{email}' is not a valid email address.")
        return 1

    if not password:
        password = getpass.getpass("Password (min 8 chars): ")
        confirm = getpass.getpass("Confirm password: ")
        if password != confirm:
            print("✗ Passwords do not match.")
            return 1

    if len(password) < 8:
        print("✗ Password must be at least 8 characters.")
        return 1

    try:
        if find_user_by_email(email):
            print(f"✗ An account for {email} already exists.")
            return 1
        user = create_user(email, password, name or email.split("@")[0])
    except HTTPException as e:
        print(f"✗ {e.detail}")
        return 1
    except Exception as e:
        print(f"✗ Could not create the account: {e}")
        return 1

    print(f"✓ Created account for {user['email']} (id {user['uid']}) in MongoDB.")
    print("  They can now sign in from any device.")
    return 0


if __name__ == "__main__":
    import uvicorn
    import socket

    # Admin sub-commands run and exit; anything else starts the server.
    if len(sys.argv) > 1 and sys.argv[1] == "create-user":
        raise SystemExit(_cli_create_user(sys.argv[2:]))

    target_port = find_available_port(int(os.getenv("PORT", "8005")))

    # HOST defaults to 0.0.0.0 so phones/laptops on the same network can reach the app.
    # Set HOST=127.0.0.1 to restrict it back to this machine only.
    host = os.getenv("HOST", "0.0.0.0")

    logger.info("Starting Mutual Fund Analytics on %s:%s", host, target_port)
    logger.info("Persistent MongoDB database: %s", MONGO_DB_NAME)
    n_users = count_users()
    if n_users == 0:
        logger.warning("No accounts exist. Create the first account with: python main.py create-user <email>")
    elif n_users > 0:
        logger.info("Accounts: %s; public signup: %s", n_users, ALLOW_SIGNUP)

    # One server process owns weekly rotation. Uvicorn lifecycle and errors use
    # our logging handlers; HTTP access is logged by RequestLoggingMiddleware.
    uvicorn.run(app, host=host, port=target_port, log_config=None, access_log=False)
