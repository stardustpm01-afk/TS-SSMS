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

# 日本語フォント設定（IPAゴシック or Noto Sans）
FONT_NAME = "IPAGothic"
try:
    import subprocess
    result = subprocess.run(
        ["fc-list", ":lang=ja", "--format=%{file}\n"],
        capture_output=True, text=True
    )
    font_paths = [l for l in result.stdout.splitlines() if l.endswith(".ttf") or l.endswith(".TTF")]
    FONT_PATH = font_paths[0] if font_paths else None
    if FONT_PATH:
        pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
except Exception:
    FONT_PATH = None

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
    model = genai.GenerativeModel("gemini-2.0-flash")

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
        return Paragraph(str(text) if text else "-", style)

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

            # PDF出力ボタン
            st.markdown("---")
            try:
                pdf_bytes = generate_pdf(data)
                safe_name = bi.get("氏名", f"skillsheet_{idx}").replace(" ", "_")
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

    # 一括PDF出力
    if len(st.session_state["results"]) > 1:
        st.divider()
        st.markdown("### 📦 一括ダウンロード")
        st.info("複数ファイルは個別にダウンロードしてください（上記ボタン）")
