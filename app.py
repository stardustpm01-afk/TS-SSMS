"""
スキルシート変換ツール
- Excel/PDFのスキルシートをアップロード
- Gemini APIで情報抽出
- 自社レイアウト（SAP職務経歴書）形式で表示・Excel出力
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
# PDF出力は不使用

# ─────────────────────────────────────────
# 設定
# ─────────────────────────────────────────
# APIキー：ローカル(.env) / Streamlit Cloud(secrets)両対応
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
except Exception:
    GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")


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
    抽出データを整形されたExcel形式で出力（罫線完全修正版）

    【修正ポイント】
    openpyxl はマージ後の内側セルをXMLに記録しない。
    そのためマージ前に全セルの外枠罫線を設定する。

    【列構成 全17列 A-Q】
    A=No  B=開始年月  C=終了年月  D=業種
    E=システム概要・担当業務  F=OS/言語/DB/ツール  G=役割
    H-Q=フェーズ10列

    【基本情報（縦線統一）】
    Row2: A B-D E F-G H I-L M N-Q （4グループ）
    Row3: 同上
    Row4: A B-D E F-G H I-Q      （住所を H列右から）
    Row5: A B-Q                  （最寄駅）
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.page import PageMargins
    from io import BytesIO as _BytesIO

    wb = Workbook()
    ws = wb.active
    ws.title = "職務経歴書"

    # ── スタイル定義 ────────────────────────────────
    HDR_FILL   = PatternFill("solid", fgColor="1A3A6B")
    HDR_FONT   = Font(name="Meiryo UI", color="FFFFFF", bold=True, size=9)
    VAL_FONT   = Font(name="Meiryo UI", size=9)
    TTL_FONT   = Font(name="Meiryo UI", bold=True, size=14, color="1A3A6B")
    SEC_FONT   = Font(name="Meiryo UI", bold=True, size=11, color="1A3A6B")
    STRIPE     = PatternFill("solid", fgColor="EEF2FF")
    THIN       = Side(style="thin")
    NONE_S     = Side(style=None)
    BORDER     = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    C_CENTER   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    C_LEFT     = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    C_LEFT_TOP = Alignment(horizontal="left",   vertical="top",    wrap_text=True)

    # ────────────────────────────────────────────────────────────────────
    # 【核心修正】マージ前に全セルの外枠罫線を設定するヘルパー
    # ────────────────────────────────────────────────────────────────────
    def _pre_border(r1, c1, r2, c2):
        """
        マージ前に範囲 (r1,c1)-(r2,c2) の全セルに外枠罫線を設定。
        左端列 → 左罫線のみ / 右端列 → 右罫線のみ / 上下端行 → 上下罫線
        ※ マージ前に呼ぶこと！
        """
        for row in range(r1, r2 + 1):
            for col in range(c1, c2 + 1):
                L = THIN if col == c1 else NONE_S
                R = THIN if col == c2 else NONE_S
                T = THIN if row == r1 else NONE_S
                B = THIN if row == r2 else NONE_S
                ws.cell(row=row, column=col).border = Border(
                    left=L, right=R, top=T, bottom=B)

    # ── ヘルパー関数 ────────────────────────────────
    def _hdr(r, c, v):
        """ヘッダーセル（単一セル・紺背景・白文字）"""
        cell = ws.cell(row=r, column=c, value=v)
        cell.fill      = HDR_FILL
        cell.font      = HDR_FONT
        cell.alignment = C_CENTER
        cell.border    = BORDER          # 単一セルは直接設定
        return cell

    def _mhdr(r, c1, c2, v):
        """マージヘッダー（紺背景）- マージ前に罫線設定"""
        _pre_border(r, c1, r, c2)
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        cell = ws.cell(row=r, column=c1, value=v)
        cell.fill      = HDR_FILL
        cell.font      = HDR_FONT
        cell.alignment = C_CENTER
        return cell

    def _mval(r, c1, c2, v, align=None):
        """マージ値セル - マージ前に罫線設定（核心修正）"""
        _pre_border(r, c1, r, c2)                          # ← マージ前に罫線！
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        cell = ws.cell(row=r, column=c1, value=safe_str(v))
        cell.font      = VAL_FONT
        cell.alignment = align or C_LEFT
        return cell

    def _title_row(r, c1, c2, v, font=None, align=None):
        """タイトル・セクション行（マージ）"""
        _pre_border(r, c1, r, c2)
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        cell = ws.cell(row=r, column=c1, value=v)
        cell.font      = font  or TTL_FONT
        cell.alignment = align or C_CENTER
        return cell

    def _row_h(r, *texts, min_h=18, cpl=40):
        """テキスト量から行高さを動的設定"""
        max_lines = 1
        for t in texts:
            s = str(t or "")
            lines = s.count("\n") + max(1, len(s) // cpl)
            max_lines = max(max_lines, lines)
        ws.row_dimensions[r].height = max(min_h, max_lines * 13)

    # ── 列幅設定（全17列 A-Q）────────────────────────
    col_w = {
        "A": 5,   "B": 10,  "C": 10,  "D": 13,
        "E": 48,  "F": 20,  "G": 8,
        "H": 5,   "I": 5,   "J": 5,   "K": 5,   "L": 5,
        "M": 5,   "N": 5,   "O": 5,   "P": 5,   "Q": 5,
    }
    for cl, w in col_w.items():
        ws.column_dimensions[cl].width = w

    # ── ページ設定（A4横向き） ─────────────────────────
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = 9
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.75, bottom=0.75)

    bi = data.get("基本情報", {}) or {}
    si = data.get("SAP情報",  {}) or {}

    # ══ Row1: タイトル ════════════════════════════════
    _title_row(1, 1, 17, "職　務　経　歴　書", TTL_FONT, C_CENTER)
    ws.row_dimensions[1].height = 28
    r = 2

    # ══ Row2: フリガナ / 性別 / 生年月日 / 未・既婚 ══════
    # 縦区切り: 1|2, 4|5, 5|6, 7|8, 8|9, 12|13, 13|14, 17
    _hdr(r, 1,  "フリガナ");   _mval(r, 2,  4,  bi.get("フリガナ"))
    _hdr(r, 5,  "性別");       _mval(r, 6,  7,  bi.get("性別"), C_CENTER)
    _hdr(r, 8,  "生年月日");   _mval(r, 9,  12, bi.get("生年月日"))
    _hdr(r, 13, "未・既婚");   _mval(r, 14, 17, bi.get("未既婚"), C_CENTER)
    ws.row_dimensions[r].height = 18;  r += 1

    # ══ Row3: 氏名 / 国籍 / 日本滞在 / SAP経験 ══════════
    # 縦区切り: 同上（統一）
    _hdr(r, 1,  "氏　名");     _mval(r, 2,  4,  bi.get("氏名"))
    _hdr(r, 5,  "国籍");       _mval(r, 6,  7,  bi.get("国籍"), C_CENTER)
    _hdr(r, 8,  "日本滞在");   _mval(r, 9,  12, bi.get("日本滞在年数"))
    _hdr(r, 13, "SAP経験");    _mval(r, 14, 17, si.get("SAP経験年数"), C_CENTER)
    ws.row_dimensions[r].height = 18;  r += 1

    # ══ Row4: モジュール / ポジション / 住所 ═════════════
    # 縦区切り: 1|2, 4|5, 5|6, 7|8, 8|9 ← Row2/3と最大限一致
    _hdr(r, 1, "モジュール");  _mval(r, 2, 4,  si.get("モジュール"))
    _hdr(r, 5, "ポジション");  _mval(r, 6, 7,  si.get("ポジション"))
    _hdr(r, 8, "住　所");      _mval(r, 9, 17, bi.get("住所"))
    ws.row_dimensions[r].height = 18;  r += 1

    # ══ Row5: 最寄駅 ══════════════════════════════════
    _hdr(r, 1, "最寄駅")
    route   = safe_str(bi.get("最寄駅路線", ""), "")
    name    = safe_str(bi.get("最寄駅名",   ""), "")
    station = " ".join(filter(None, [
        route,
        (name + "駅") if name else ""
    ])) or safe_str(bi.get("最寄駅", ""))
    _mval(r, 2, 17, station)
    ws.row_dimensions[r].height = 18;  r += 1

    # ══ Row6-8: 資格・得意分野・自己PR ═══════════════════
    for label, key, min_h, cpl in [
        ("取得資格", "取得資格", 18, 80),
        ("得意分野", "得意分野", 18, 80),
        ("自己ＰＲ", "自己PR",  36, 50),
    ]:
        _hdr(r, 1, label)
        _mval(r, 2, 17, data.get(key), C_LEFT_TOP)
        _row_h(r, data.get(key), min_h=min_h, cpl=cpl)
        r += 1

    # 区切り空行
    ws.row_dimensions[r].height = 5;  r += 1

    # ══ 業務経歴タイトル ══════════════════════════════
    _title_row(r, 1, 17, "業　務　経　歴", SEC_FONT, C_CENTER)
    ws.row_dimensions[r].height = 22;  r += 1

    # ══ 業務経歴ヘッダー行 ═══════════════════════════
    proj_header_row = r
    h_vals = [
        "No", "開始\n年月", "終了\n年月", "業種",
        "システム概要・担当業務",
        "OS/言語/DB/ツール", "役割",
        "分析\n調査","提案\n管理","要件\n定義","基本\n設計",
        "詳細\n設計","製造","単体\n試験","結合\n試験","総合\n試験","運用\n保守"
    ]
    for ci, hv in enumerate(h_vals, 1):
        _hdr(r, ci, hv)
    ws.row_dimensions[r].height = 30;  r += 1

    # ══ 業務経歴データ行 ══════════════════════════════
    phase_keys = [
        "分析調査","提案管理レビュー","要件定義","基本設計",
        "詳細設計","製造","単体試験","結合試験","総合試験","運用保守"
    ]

    for pi, proj in enumerate(data.get("職務経歴", []) or []):
        phases = proj.get("フェーズ", {}) or {}
        bg     = STRIPE if pi % 2 == 0 else None

        overview = safe_str(proj.get("プロジェクト概要"))
        tasks    = safe_str(proj.get("担当業務"))
        combined = (f"【概要】{overview}\n【担当】{tasks}"
                    if tasks not in ("-", overview, "") else overview)

        os_db    = safe_str(proj.get("OS_DB"))
        env      = safe_str(proj.get("作業環境"))
        lang     = safe_str(proj.get("開発言語"))
        env_text = "\n".join(x for x in [os_db, env, lang] if x and x != "-")

        row_vals = [
            safe_str(proj.get("No", pi + 1)),
            safe_str(proj.get("開始年月")),
            safe_str(proj.get("終了年月")),
            safe_str(proj.get("業種")),
            combined,
            env_text or "-",
            safe_str(proj.get("役割")),
        ] + ["●" if phases.get(k) else "" for k in phase_keys]

        for ci, v in enumerate(row_vals, 1):
            if ci <= 3:
                align = C_CENTER
            elif ci in (5, 6):
                align = C_LEFT_TOP
            else:
                align = C_CENTER if ci >= 8 else C_LEFT
            cell = ws.cell(row=r, column=ci, value=v)
            cell.font      = VAL_FONT
            cell.alignment = align
            cell.border    = BORDER
            if bg:
                cell.fill = bg

        _row_h(r, combined, env_text, min_h=25, cpl=38)
        r += 1

    # フリーズペイン・印刷範囲
    ws.freeze_panes = ws.cell(row=proj_header_row + 1, column=1)
    ws.print_area   = f"A1:Q{r - 1}"

    buf = _BytesIO()
    wb.save(buf)
    return buf.getvalue()




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

            # Excelダウンロードボタン
            st.markdown("---")
            try:
                excel_bytes = generate_excel(data)
                safe_name = safe_str(bi.get("氏名", f"skillsheet_{idx}"), f"skillsheet_{idx}").replace(" ", "_")
                st.download_button(
                    label="📊 Excelダウンロード（自社フォーマット）",
                    data=excel_bytes,
                    file_name=f"{safe_name}_職務経歴書.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary",
                    key=f"xl_{idx}"
                )
            except Exception as e:
                st.error(f"Excel生成エラー: {e}")
                import traceback
                st.code(traceback.format_exc(), language="text")

    # 一括PDF出力
    if len(st.session_state["results"]) > 1:
        st.divider()
        st.markdown("### 📦 一括ダウンロード")
        st.info("複数ファイルは個別にExcelダウンロードしてください（上記ボタン）")
