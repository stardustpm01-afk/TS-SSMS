"""
スキルシート変換ツール
- Excel/PDFのスキルシートをアップロード
- Gemini APIで情報抽出
- 自社レイアウト（SAP職務経歴書）形式で表示・PDF出力
"""

import streamlit as st
import google.generativeai as genai
import openpyxl
import pdfplumber
import fitz  # PyMuPDF
import json
import os
import re
from io import BytesIO
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT, TA_CENTER

# ─────────────────────────────────────────
# 設定
# ─────────────────────────────────────────
# APIキー：ローカル(.env) / Streamlit Cloud(secrets)両対応
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
except Exception:
    GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

# 日本語フォント設定（優先順位付きで検索）
FONT_NAME = "JPFont"
FONT_PATH = None

# 候補フォントパスリスト（Linux/Mac/Windows対応）
_font_candidates = [
    # Streamlit Cloud (Ubuntu) - IPAフォント
    "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf",
    "/usr/share/fonts/truetype/ipafont-gothic/ipagp.ttf",
    "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
    "/usr/share/fonts/truetype/ipafont-gothic/ipag.ttf",
    "/usr/share/fonts/opentype/ipafont/ipagp.ttf",
    "/usr/share/fonts/truetype/ipafont/ipagp.ttf",
    # Streamlit Cloud - Noto CJK
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
    # macOS
    "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
    "/Library/Fonts/Arial Unicode MS.ttf",
    # Windows
    "C:/Windows/Fonts/msgothic.ttc",
    "C:/Windows/Fonts/YuGothM.ttc",
    "C:/Windows/Fonts/meiryo.ttc",
]

for _p in _font_candidates:
    if os.path.exists(_p):
        try:
            pdfmetrics.registerFont(TTFont(FONT_NAME, _p))
            FONT_PATH = _p
            break
        except Exception:
            continue

# fc-list でも検索（上記で見つからない場合）
if not FONT_PATH:
    try:
        import subprocess
        result = subprocess.run(
            ["fc-list", ":lang=ja", "--format=%{file}\n"],
            capture_output=True, text=True, timeout=5
        )
        for _p in result.stdout.splitlines():
            if _p.endswith((".ttf", ".TTF", ".otf")):
                try:
                    pdfmetrics.registerFont(TTFont(FONT_NAME, _p))
                    FONT_PATH = _p
                    break
                except Exception:
                    continue
    except Exception:
        pass

# ─────────────────────────────────────────
# ページ設定
# ─────────────────────────────────────────
st.set_page_config(
    page_title="スキルシート変換ツール",
    page_icon="📄",
    layout="wide"
)

st.markdown("""
<style>
    .main-title {
        font-size: 2rem;
        font-weight: bold;
        color: #1a3a6b;
        margin-bottom: 0.5rem;
    }
    .section-header {
        background: #1a3a6b;
        color: white;
        padding: 6px 12px;
        border-radius: 4px;
        font-weight: bold;
        margin: 1rem 0 0.5rem 0;
    }
    .project-card {
        background: #f8f9fa;
        border-left: 4px solid #1a3a6b;
        padding: 10px 15px;
        margin: 8px 0;
        border-radius: 0 4px 4px 0;
    }
    .badge {
        background: #1a3a6b;
        color: white;
        padding: 2px 8px;
        border-radius: 10px;
        font-size: 0.8rem;
        margin-right: 4px;
    }
    .skill-tag {
        background: #e8f0fe;
        color: #1a3a6b;
        padding: 2px 8px;
        border-radius: 10px;
        font-size: 0.85rem;
        margin: 2px;
        display: inline-block;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# ユーティリティ関数
# ─────────────────────────────────────────

def extract_text_from_excel(file_bytes: bytes) -> str:
    """ExcelファイルからテキストをGrid形式で抽出"""
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    result = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        result.append(f"=== シート: {sheet_name} ===")
        for row in ws.iter_rows():
            row_texts = []
            for cell in row:
                val = cell.value
                if val is not None:
                    if isinstance(val, datetime):
                        val = val.strftime("%Y/%m")
                    row_texts.append(f"[{cell.coordinate}]{str(val).strip()}")
            if row_texts:
                result.append("  |  ".join(row_texts))
    return "\n".join(result)


def extract_text_from_pdf(file_bytes: bytes) -> str:
    """PDFからテキストを抽出"""
    text_parts = []
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                text_parts.append(f"=== ページ {i+1} ===\n{text}")
    return "\n".join(text_parts)


def extract_skills_with_gemini(file_text: str, api_key: str) -> dict:
    """Gemini APIでスキルシート情報をJSON抽出"""
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.5-flash")

    prompt = f"""
以下はスキルシート（職務経歴書）のテキストデータです。
このデータから必要な情報を抽出して、必ず以下のJSON形式のみで返してください。
説明文や```json```の囲みは不要です。JSONオブジェクトだけを返してください。

抽出するJSON形式:
{{
  "基本情報": {{
    "氏名": "",
    "フリガナ": "",
    "性別": "",
    "生年月日": "",
    "年齢": "",
    "未既婚": "",
    "国籍": "",
    "日本滞在年数": "",
    "住所": "",
    "最寄駅路線": "",
    "最寄駅名": ""
  }},
  "SAP情報": {{
    "モジュール": "",
    "ポジション": "",
    "SAP経験年数": ""
  }},
  "言語能力": {{
    "日本語": {{
      "日常会話": "",
      "業務会話": "",
      "読み": "",
      "書き": "",
      "仕様書読解": "",
      "仕様書作成": ""
    }},
    "英語": {{
      "日常会話": "",
      "業務会話": "",
      "読み": "",
      "書き": "",
      "仕様書読解": "",
      "仕様書作成": ""
    }}
  }},
  "取得資格": "",
  "得意分野": "",
  "自己PR": "",
  "職務経歴": [
    {{
      "No": 1,
      "開始年月": "",
      "終了年月": "",
      "業種": "",
      "プロジェクト概要": "",
      "担当業務": "",
      "OS_DB": "",
      "作業環境": "",
      "開発言語": "",
      "役割": "",
      "フェーズ": {{
        "分析調査": false,
        "提案管理レビュー": false,
        "要件定義": false,
        "基本設計": false,
        "詳細設計": false,
        "製造": false,
        "単体試験": false,
        "結合試験": false,
        "総合試験": false,
        "運用保守": false
      }}
    }}
  ]
}}

スキルシートデータ:
{file_text[:8000]}
"""

    response = model.generate_content(prompt)
    raw = response.text.strip()
    # JSONの抽出（```json...``` ブロックがあれば除去）
    raw = re.sub(r"^```json\s*", "", raw)
    raw = re.sub(r"^```\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw)



# ─────────────────────────────────────────
# ユーティリティ：None安全文字列変換
# ─────────────────────────────────────────
def safe_str(val, default="-"):
    """NoneやNaN等を安全に文字列変換"""
    if val is None:
        return default
    s = str(val).strip()
    return s if s else default


# ─────────────────────────────────────────
# Excel出力
# ─────────────────────────────────────────
def generate_excel(data: dict) -> bytes:
    """
    抽出データを整形されたExcel形式で出力（A4横向き印刷対応）

    列構成（全17列 A～Q）:
    A=No, B=開始年月, C=終了年月, D=業種,
    E=システム概要・担当業務（統合）, F=OS/言語/DB/ツール, G=役割,
    H-Q=フェーズ10列
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.page import PageMargins
    from io import BytesIO as _BytesIO

    wb = Workbook()
    ws = wb.active
    ws.title = "職務経歴書"

    # ─── スタイル定義 ─────────────────────────
    HDR_FILL   = PatternFill("solid", fgColor="1A3A6B")
    HDR_FONT   = Font(name="Meiryo UI", color="FFFFFF", bold=True, size=9)
    VAL_FONT   = Font(name="Meiryo UI", size=9)
    TTL_FONT   = Font(name="Meiryo UI", bold=True, size=14, color="1A3A6B")
    SEC_FONT   = Font(name="Meiryo UI", bold=True, size=11, color="1A3A6B")
    STRIPE     = PatternFill("solid", fgColor="EEF2FF")
    THIN       = Side(style="thin")
    BORDER     = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    C_CENTER   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    C_LEFT     = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    C_LEFT_TOP = Alignment(horizontal="left",   vertical="top",    wrap_text=True)

    def _hdr(r, c, v):
        cell = ws.cell(row=r, column=c, value=v)
        cell.fill, cell.font, cell.alignment, cell.border = HDR_FILL, HDR_FONT, C_CENTER, BORDER
        return cell

    def _val(r, c, v, align=None):
        cell = ws.cell(row=r, column=c, value=safe_str(v))
        cell.font, cell.alignment, cell.border = VAL_FONT, (align or C_LEFT), BORDER
        return cell

    def _mval(r, c1, c2, v, align=None):
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        return _val(r, c1, v, align or C_LEFT)

    def _row_height(r, text, min_h=18, cpl=40):
        n = max(1, len(str(text or "")) // cpl + str(text or "").count("\n") + 1)
        ws.row_dimensions[r].height = max(min_h, n * 13)

    # ─── 列幅設定 ─────────────────────────────
    col_w = {
        "A": 5,  "B": 10, "C": 10, "D": 13,
        "E": 48, "F": 18, "G": 8,
        "H": 4.5,"I": 4.5,"J": 4.5,"K": 4.5,"L": 4.5,
        "M": 4.5,"N": 4.5,"O": 4.5,"P": 4.5,"Q": 4.5,
    }
    for col_letter, w in col_w.items():
        ws.column_dimensions[col_letter].width = w

    # ─── ページ設定（A4横向き） ───────────────
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = 9
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.75, bottom=0.75)

    bi = data.get("基本情報", {}) or {}
    si = data.get("SAP情報",  {}) or {}
    LAST_COL = 17  # Q列

    # ══ 行1: タイトル ════════════════════════════
    r = 1
    ws.merge_cells(f"A{r}:Q{r}")
    t = ws.cell(row=r, column=1, value="職　務　経　歴　書")
    t.font, t.alignment = TTL_FONT, C_CENTER
    ws.row_dimensions[r].height = 28
    r += 1

    # ══ 基本情報テーブル（行2-5） ════════════════
    # 行2: フリガナ(A) | 値(B-D) | 性別(E) | 値(F) | 生年月日(G) | 値(H-K) | 未・既婚(L) | 値(M-Q)
    _hdr(r, 1, "フリガナ");   _mval(r, 2, 4,  bi.get("フリガナ"))
    _hdr(r, 5, "性別");       _val(r,  6,      bi.get("性別"), C_CENTER)
    _hdr(r, 7, "生年月日");   _mval(r, 8, 11,  bi.get("生年月日"))
    _hdr(r, 12, "未・既婚");  _mval(r, 13, 17, bi.get("未既婚"), C_CENTER)
    ws.row_dimensions[r].height = 18; r += 1

    # 行3: 氏名(A) | 値(B-D) | 国籍(E) | 値(F) | 日本滞在(G) | 値(H-K) | SAP経験(L) | 値(M-Q)
    _hdr(r, 1, "氏　名");     _mval(r, 2, 4,  bi.get("氏名"))
    _hdr(r, 5, "国籍");       _val(r,  6,      bi.get("国籍"), C_CENTER)
    _hdr(r, 7, "日本滞在");   _mval(r, 8, 11,  bi.get("日本滞在年数"))
    _hdr(r, 12, "SAP経験");   _mval(r, 13, 17, si.get("SAP経験年数"), C_CENTER)
    ws.row_dimensions[r].height = 18; r += 1

    # 行4: モジュール(A) | 値(B-D) | ポジション(E) | 値(F-I) | 住所(J) | 値(K-Q)
    _hdr(r, 1, "モジュール"); _mval(r, 2, 4,  si.get("モジュール"))
    _hdr(r, 5, "ポジション"); _mval(r, 6, 9,  si.get("ポジション"))
    _hdr(r, 10, "住　所");    _mval(r, 11, 17, bi.get("住所"))
    ws.row_dimensions[r].height = 18; r += 1

    # 行5: 最寄駅(A) | 値(B-Q)
    _hdr(r, 1, "最寄駅")
    station = f"{safe_str(bi.get('最寄駅路線',''), '')} {safe_str(bi.get('最寄駅名',''), '')}駅".strip()
    _mval(r, 2, 17, station if station != "駅" else "-")
    ws.row_dimensions[r].height = 18; r += 1

    # ══ 資格・得意分野・自己PR ════════════════════
    for label, key, min_h, cpl in [
        ("取得資格", "取得資格", 18, 70),
        ("得意分野", "得意分野", 18, 70),
        ("自己ＰＲ", "自己PR",   36, 55),
    ]:
        _hdr(r, 1, label)
        _mval(r, 2, 17, data.get(key), C_LEFT_TOP)
        _row_height(r, data.get(key), min_h=min_h, cpl=cpl)
        r += 1

    # 空行
    ws.row_dimensions[r].height = 6; r += 1

    # ══ 業務経歴タイトル ══════════════════════════
    ws.merge_cells(f"A{r}:Q{r}")
    sec = ws.cell(row=r, column=1, value="業　務　経　歴")
    sec.font, sec.alignment = SEC_FONT, C_CENTER
    ws.row_dimensions[r].height = 22; r += 1

    # ══ 業務経歴ヘッダー ══════════════════════════
    proj_header_row = r
    h_vals = [
        "No","開始年月","終了年月","業種",
        "システム概要・担当業務",
        "OS/言語/DB/ツール","役割",
        "分析\n調査","提案\n管理","要件\n定義","基本\n設計",
        "詳細\n設計","製造","単体\n試験","結合\n試験","総合\n試験","運用\n保守"
    ]
    for ci, hv in enumerate(h_vals, 1):
        _hdr(r, ci, hv)
    ws.row_dimensions[r].height = 30; r += 1

    # ══ 業務経歴データ行 ══════════════════════════
    phase_keys = [
        "分析調査","提案管理レビュー","要件定義","基本設計",
        "詳細設計","製造","単体試験","結合試験","総合試験","運用保守"
    ]

    for pi, proj in enumerate(data.get("職務経歴", []) or []):
        phases = proj.get("フェーズ", {}) or {}
        bg     = STRIPE if pi % 2 == 0 else None

        overview = safe_str(proj.get("プロジェクト概要"))
        tasks    = safe_str(proj.get("担当業務"))
        combined = f"{overview}\n【担当】{tasks}" if tasks not in ("-", overview) else overview

        os_db = safe_str(proj.get("OS_DB"))
        env   = safe_str(proj.get("作業環境"))
        lang  = safe_str(proj.get("開発言語"))
        env_text = "\n".join(x for x in [os_db, env, lang] if x != "-")

        row_vals = [
            safe_str(proj.get("No", pi+1)),
            safe_str(proj.get("開始年月")),
            safe_str(proj.get("終了年月")),
            safe_str(proj.get("業種")),
            combined,
            env_text,
            safe_str(proj.get("役割")),
        ] + ["●" if phases.get(k) else "" for k in phase_keys]

        for ci, v in enumerate(row_vals, 1):
            align = C_CENTER if ci >= 8 else                     (C_LEFT_TOP if ci in (5, 6) else                     (C_CENTER if ci <= 3 else C_LEFT))
            cell = ws.cell(row=r, column=ci, value=v)
            cell.font, cell.alignment, cell.border = VAL_FONT, align, BORDER
            if bg:
                cell.fill = bg

        max_len = max(len(combined), len(env_text))
        _row_height(r, max_len, min_h=22, cpl=35)
        r += 1

    # フリーズペイン（ヘッダー固定）
    ws.freeze_panes = ws.cell(row=proj_header_row + 1, column=1)
    ws.print_area   = f"A1:Q{r-1}"

    buf = _BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────
# PDF生成
# ─────────────────────────────────────────

def generate_pdf(data: dict) -> bytes:
    """抽出データから自社レイアウトPDFを生成"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=10*mm,
        leftMargin=10*mm,
        topMargin=10*mm,
        bottomMargin=10*mm
    )

    fn = FONT_NAME if FONT_PATH else "Helvetica"
    styles = getSampleStyleSheet()
    normal = ParagraphStyle("normal", fontName=fn, fontSize=8, leading=11)
    small  = ParagraphStyle("small",  fontName=fn, fontSize=7, leading=9)
    title_style = ParagraphStyle("title", fontName=fn, fontSize=16,
                                  alignment=TA_CENTER, textColor=colors.HexColor("#1a3a6b"))
    header_style = ParagraphStyle("header", fontName=fn, fontSize=8,
                                   textColor=colors.white)

    def hdr(text):
        return Paragraph(f"<b>{text}</b>", header_style)

    def cell(text, style=normal):
        return Paragraph(safe_str(text), style)

    story = []
    bi = data.get("基本情報", {})
    si = data.get("SAP情報", {})

    # タイトル
    story.append(Paragraph("職　務　経　歴　書", title_style))
    story.append(Spacer(1, 3*mm))

    # 更新日
    today = datetime.now().strftime("%Y/%m/%d")
    story.append(Paragraph(f"最終更新日：{today}", ParagraphStyle("right", fontName=fn, fontSize=8)))
    story.append(Spacer(1, 2*mm))

    # 基本情報テーブル
    header_color = colors.HexColor("#1a3a6b")
    basic_data = [
        [hdr("フリガナ"), cell(bi.get("フリガナ","")), hdr("性別"), cell(bi.get("性別","")),
         hdr("生年月日"), cell(bi.get("生年月日","")), hdr("未・既婚"), cell(bi.get("未既婚",""))],
        [hdr("氏　名"), cell(bi.get("氏名","")), hdr("国籍"), cell(bi.get("国籍","")),
         hdr("日本滞在"), cell(bi.get("日本滞在年数","")), hdr("SAP経験"), cell(si.get("SAP経験年数",""))],
        [hdr("モジュール"), cell(si.get("モジュール","")), hdr("ポジション"), cell(si.get("ポジション","")),
         hdr("住所"), Paragraph(bi.get("住所",""), normal), "", ""],
        [hdr("最寄駅"), Paragraph(f"{bi.get('最寄駅路線','')} {bi.get('最寄駅名','')}駅", normal),
         "", "", "", "", "", ""],
    ]
    col_w = [18*mm, 30*mm, 15*mm, 20*mm, 18*mm, 25*mm, 15*mm, 20*mm]
    t = Table(basic_data, colWidths=col_w, repeatRows=0)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (0,-1), header_color),
        ("BACKGROUND", (2,0), (2,-1), header_color),
        ("BACKGROUND", (4,0), (4,-1), header_color),
        ("BACKGROUND", (6,0), (6,-1), header_color),
        ("TEXTCOLOR", (0,0), (0,-1), colors.white),
        ("TEXTCOLOR", (2,0), (2,-1), colors.white),
        ("TEXTCOLOR", (4,0), (4,-1), colors.white),
        ("TEXTCOLOR", (6,0), (6,-1), colors.white),
        ("FONTNAME", (0,0), (-1,-1), fn),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("SPAN", (5,2), (7,2)),
        ("SPAN", (1,3), (7,3)),
    ]))
    story.append(t)
    story.append(Spacer(1, 3*mm))

    # 取得資格
    story.append(Table(
        [[hdr("取得資格"), cell(data.get("取得資格",""))]],
        colWidths=[18*mm, 162*mm],
        style=TableStyle([
            ("BACKGROUND", (0,0), (0,0), header_color),
            ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
            ("FONTNAME", (0,0), (-1,-1), fn),
            ("FONTSIZE", (0,0), (-1,-1), 8),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ])
    ))
    story.append(Spacer(1, 2*mm))

    # 得意分野
    story.append(Table(
        [[hdr("得意分野"), cell(data.get("得意分野",""))]],
        colWidths=[18*mm, 162*mm],
        style=TableStyle([
            ("BACKGROUND", (0,0), (0,0), header_color),
            ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
            ("FONTNAME", (0,0), (-1,-1), fn),
            ("FONTSIZE", (0,0), (-1,-1), 8),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ])
    ))
    story.append(Spacer(1, 2*mm))

    # 自己PR
    story.append(Table(
        [[hdr("自己PR"), cell(data.get("自己PR",""))]],
        colWidths=[18*mm, 162*mm],
        style=TableStyle([
            ("BACKGROUND", (0,0), (0,0), header_color),
            ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
            ("FONTNAME", (0,0), (-1,-1), fn),
            ("FONTSIZE", (0,0), (-1,-1), 8),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("MINROWHEIGHT", (0,0), (-1,-1), 20*mm),
        ])
    ))
    story.append(Spacer(1, 3*mm))

    # 職務経歴ヘッダー
    story.append(Paragraph("業　務　経　歴", title_style))
    story.append(Spacer(1, 2*mm))

    phase_labels = ["分析\n調査", "提案\n管理", "要件\n定義", "基本\n設計",
                    "詳細\n設計", "製造", "単体\n試験", "結合\n試験", "総合\n試験", "運用\n保守"]
    phase_keys   = ["分析調査","提案管理レビュー","要件定義","基本設計",
                    "詳細設計","製造","単体試験","結合試験","総合試験","運用保守"]

    proj_header = [
        [hdr("No"), hdr("作業期間"), hdr("業種"),
         hdr("システム概要・担当業務"), hdr("OS/言語/DB/ツール"),
         hdr("役割")] + [hdr(p) for p in phase_labels]
    ]
    col_w2 = [8*mm, 22*mm, 15*mm, 48*mm, 30*mm, 15*mm] + [7*mm]*10
    proj_table_data = proj_header

    for proj in data.get("職務経歴", []):
        phases = proj.get("フェーズ", {})
        period = f"{proj.get('開始年月','')}～\n{proj.get('終了年月','')}"
        env_text = f"{proj.get('OS_DB','')}\n{proj.get('作業環境','')}\n{proj.get('開発言語','')}"
        content = f"{proj.get('プロジェクト概要','')}\n【担当】{proj.get('担当業務','')}"
        row = [
            cell(str(proj.get("No",""))),
            cell(period, small),
            cell(proj.get("業種","")),
            cell(content, small),
            cell(env_text, small),
            cell(proj.get("役割",""))
        ] + [cell("●" if phases.get(k) else "-") for k in phase_keys]
        proj_table_data.append(row)

    pt = Table(proj_table_data, colWidths=col_w2, repeatRows=1)
    pt.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), header_color),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,-1), fn),
        ("FONTSIZE", (0,0), (-1,-1), 7),
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f0f4ff")]),
        ("ALIGN", (0,0), (0,-1), "CENTER"),
        ("ALIGN", (6,0), (-1,-1), "CENTER"),
    ]))
    story.append(pt)

    doc.build(story)
    return buffer.getvalue()


# ─────────────────────────────────────────
# メインUI
# ─────────────────────────────────────────

st.markdown('<div class="main-title">📄 スキルシート変換ツール</div>', unsafe_allow_html=True)
st.caption("各社バラバラのExcel/PDFスキルシートを自社フォーマットに自動変換します")

st.divider()

# サイドバー：API設定
with st.sidebar:
    st.header("⚙️ 設定")
    api_key = st.text_input(
        "Gemini APIキー",
        value=GEMINI_API_KEY,
        type="password",
        help="Google AI StudioのAPIキーを入力"
    )
    st.caption("🔒 キーはこのセッション内のみ使用されます")
    st.divider()
    st.markdown("**対応ファイル形式**")
    st.markdown("- 📊 Excel (.xlsx, .xls)\n- 📄 PDF (.pdf)")
    st.divider()
    st.markdown("**処理フロー**")
    st.markdown("1. ファイルアップロード\n2. テキスト抽出\n3. Gemini AI解析\n4. 自社フォーマット表示\n5. PDF出力")

# メインエリア
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown('<div class="section-header">📁 ファイルアップロード</div>', unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "スキルシートをアップロード（複数可）",
        type=["xlsx", "xls", "pdf"],
        accept_multiple_files=True,
        help="ExcelまたはPDF形式のスキルシートをドラッグ＆ドロップ"
    )

    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)}件のファイルが選択されました")
        for f in uploaded_files:
            size_kb = len(f.getvalue()) / 1024
            st.caption(f"• {f.name}（{size_kb:.1f} KB）")

with col2:
    st.markdown('<div class="section-header">🚀 変換実行</div>', unsafe_allow_html=True)
    if not api_key:
        st.warning("⚠️ サイドバーにGemini APIキーを入力してください")
    elif not uploaded_files:
        st.info("👆 左側からファイルをアップロードしてください")
    else:
        if st.button("🤖 AIで自動変換する", type="primary", use_container_width=True):
            st.session_state["results"] = []
            for uploaded_file in uploaded_files:
                with st.spinner(f"⏳ {uploaded_file.name} を処理中..."):
                    try:
                        file_bytes = uploaded_file.getvalue()
                        ext = uploaded_file.name.lower().split(".")[-1]

                        # テキスト抽出
                        if ext in ["xlsx", "xls"]:
                            text = extract_text_from_excel(file_bytes)
                        else:
                            text = extract_text_from_pdf(file_bytes)

                        # Gemini API呼び出し
                        extracted = extract_skills_with_gemini(text, api_key)
                        extracted["_filename"] = uploaded_file.name
                        st.session_state["results"].append(extracted)
                        st.success(f"✅ {uploaded_file.name} 完了！")
                    except json.JSONDecodeError as e:
                        st.error(f"❌ JSON解析エラー: {uploaded_file.name}\n{e}")
                    except Exception as e:
                        st.error(f"❌ エラー: {uploaded_file.name}\n{e}")

# ─────────────────────────────────────────
# 結果表示エリア
# ─────────────────────────────────────────
if "results" in st.session_state and st.session_state["results"]:
    st.divider()
    st.markdown("## 📊 抽出結果・プレビュー")

    for idx, data in enumerate(st.session_state["results"]):
        bi = data.get("基本情報", {})
        si = data.get("SAP情報", {})
        name = bi.get("氏名", f"ファイル{idx+1}")

        with st.expander(f"👤 {name}（{data.get('_filename','')}）", expanded=True):

            # 基本情報
            st.markdown('<div class="section-header">👤 基本情報</div>', unsafe_allow_html=True)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("氏名", bi.get("氏名", "-"))
            c2.metric("生年月日", bi.get("生年月日", "-"))
            c3.metric("性別", bi.get("性別", "-"))
            c4.metric("未・既婚", bi.get("未既婚", "-"))

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("モジュール", si.get("モジュール", "-"))
            c2.metric("ポジション", si.get("ポジション", "-"))
            c3.metric("SAP経験年数", si.get("SAP経験年数", "-"))
            c4.metric("最寄駅", f"{bi.get('最寄駅路線','')} {bi.get('最寄駅名','')}駅")

            # 資格・PR
            st.markdown('<div class="section-header">📜 資格・スキル</div>', unsafe_allow_html=True)
            st.info(f"**取得資格：** {data.get('取得資格', '-')}")

            with st.container():
                st.markdown("**得意分野**")
                st.text_area("得意分野", data.get("得意分野", ""), height=80,
                              label_visibility="collapsed", key=f"tokui_{idx}")
                st.markdown("**自己PR**")
                st.text_area("自己PR", data.get("自己PR", ""), height=80,
                              label_visibility="collapsed", key=f"pr_{idx}")

            # 職務経歴
            st.markdown('<div class="section-header">💼 職務経歴</div>', unsafe_allow_html=True)
            projects = data.get("職務経歴", [])
            if projects:
                for proj in projects:
                    phases = proj.get("フェーズ", {})
                    active = [k for k, v in phases.items() if v]
                    st.markdown(f"""
<div class="project-card">
<b>No.{proj.get('No','')} | {proj.get('開始年月','')} ～ {proj.get('終了年月','')}</b>
　<span class="badge">{proj.get('業種','')}</span>
　<span class="badge">{proj.get('役割','')}</span>
<br><br>
{proj.get('プロジェクト概要','')}
<br><br>
<b>担当：</b>{proj.get('担当業務','')}
<br>
<b>環境：</b>{proj.get('OS_DB','')} / {proj.get('作業環境','')} / {proj.get('開発言語','')}
<br>
<b>フェーズ：</b>{"　".join([f'<span class="skill-tag">{p}</span>' for p in active])}
</div>
""", unsafe_allow_html=True)
            else:
                st.warning("職務経歴が抽出できませんでした")

            # 出力ボタン
            st.markdown("---")
            col_pdf, col_xlsx = st.columns(2)

            with col_pdf:
                try:
                    pdf_bytes = generate_pdf(data)
                    safe_name = safe_str(bi.get("氏名", f"skillsheet_{idx}"), f"skillsheet_{idx}").replace(" ", "_")
                    st.download_button(
                        label="⬇️ PDFダウンロード（自社フォーマット）",
                        data=pdf_bytes,
                        file_name=f"{safe_name}_職務経歴書.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        type="primary",
                        key=f"dl_{idx}"
                    )
                except Exception as e:
                    st.error(f"PDF生成エラー: {e}")
                    import traceback
                    st.code(traceback.format_exc(), language="text")

            with col_xlsx:
                try:
                    excel_bytes = generate_excel(data)
                    safe_name = safe_str(bi.get("氏名", f"skillsheet_{idx}"), f"skillsheet_{idx}").replace(" ", "_")
                    st.download_button(
                        label="📊 Excelダウンロード",
                        data=excel_bytes,
                        file_name=f"{safe_name}_職務経歴書.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key=f"xl_{idx}"
                    )
                except Exception as e:
                    st.error(f"Excel生成エラー: {e}")

    # 一括PDF出力
    if len(st.session_state["results"]) > 1:
        st.divider()
        st.markdown("### 📦 一括ダウンロード")
        st.info("複数ファイルは個別にダウンロードしてください（上記ボタン）")
