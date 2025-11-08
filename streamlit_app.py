# -*- coding: utf-8 -*-
# 3分無料診断（Phase 1+2+3 完成版 / JST対応 / QR右寄せ / ロゴ堅牢化）
# - 設問10問 → スコア化 → 6タイプ判定 → 信号色表示
# - PDF出力（日本語TTF埋め込み、棒グラフ、ロゴ/ブランド色、URL横にQR）
# - ログ保存（Google Sheets / CSV）
# - OpenAIで“約300字”のAIコメント自動生成（Secrets/環境変数両対応）
# - セッション保持で再実行しても結果画面を維持
# - ロゴはローカル優先（/content/CImark.png 等）→ 失敗時はURL取得

import os
import io
import json
import time
import tempfile
from datetime import datetime, timedelta, timezone
import urllib.request

import streamlit as st
import pandas as pd
import altair as alt
import matplotlib.pyplot as plt

# ReportLab（PDF）
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily

# Matplotlib日本語フォント
from matplotlib import font_manager

# 画像・QR
from PIL import Image as PILImage
import qrcode

# Google Sheets（任意）
import gspread
from google.oauth2.service_account import Credentials

# ネットワークフォールバック
import requests

# ========= ブランド & 定数 =========
BRAND_BG = "#f0f7f7"
LOGO_LOCAL = "/content/CImark.png"  # Colabにアップしたら最優先で使用
LOGO_URL   = "https://victorconsulting.jp/wp-content/uploads/2025/10/CImark.png"
CTA_URL    = "https://victorconsulting.jp/spot-diagnosis/"
OPENAI_MODEL = "gpt-4o-mini"

# 日本時間
JST = timezone(timedelta(hours=9))

st.set_page_config(
    page_title="3分無料診断｜Victor Consulting",
    page_icon="✅",
    layout="centered",
    initial_sidebar_state="expanded"
)

# ---- session init ----
for k, v in {
    "result_ready": False, "df": None, "overall_avg": None, "signal": None,
    "main_type": None, "company": "", "email": "", "ai_comment": None
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ========= Secrets 安全読み取り =========
def read_secret(key: str, default=None):
    try:
        return st.secrets[key]
    except Exception:
        return os.environ.get(key, default)

# ========= 日本語TTF 登録（ReportLab & Matplotlib）=========
def setup_japanese_font():
    candidates = [
        "/content/NotoSansJP-Regular.ttf",
        "/mnt/data/NotoSansJP-Regular.ttf",
        "./NotoSansJP-Regular.ttf",
    ]
    font_path = next((p for p in candidates if os.path.exists(p)), None)
    if not font_path:
        return None
    try:
        pdfmetrics.registerFont(TTFont("JP", font_path))
        registerFontFamily("JP", normal="JP", bold="JP", italic="JP", boldItalic="JP")
    except Exception as e:
        print("ReportLab font register error:", e)
    try:
        font_manager.fontManager.addfont(font_path)
        fp = font_manager.FontProperties(fname=font_path)
        import matplotlib as mpl
        mpl.rcParams["font.family"] = fp.get_name()
        mpl.rcParams["axes.unicode_minus"] = False
    except Exception as e:
        print("Matplotlib font register error:", e)
    return font_path

FONT_PATH_IN_USE = setup_japanese_font()

# ========= スタイル =========
st.markdown(
    f"""
<style>
.stApp {{ background: {BRAND_BG}; }}
.block-container {{ padding-top: 2.8rem; }}   /* タイトル頭が切れないよう余白拡大 */
h1 {{ margin-top: .6rem; }}
.result-card {{
  background: white; border-radius: 14px; padding: 1.2rem 1.1rem;
  box-shadow: 0 6px 20px rgba(0,0,0,.06); border: 1px solid rgba(0,0,0,.06);
}}
.badge {{ display:inline-block; padding:.25rem .6rem; border-radius:999px; font-size:.9rem;
  font-weight:700; letter-spacing:.02em; margin-left:.5rem; }}
.badge-blue  {{ background:#e6f0ff; color:#0b5fff; border:1px solid #cfe3ff; }}
.badge-yellow{{ background:#fff6d8; color:#8a6d00; border:1px solid #ffecb3; }}
.badge-red   {{ background:#ffe6e6; color:#a80000; border:1px solid #ffc7c7; }}
.small-note {{ color:#666; font-size:.9rem; }}
hr {{ border:none; border-top:1px dotted #c9d7d7; margin:1.1rem 0; }}
</style>
""",
    unsafe_allow_html=True
)

# ========= ロゴ取得（ローカル優先 → URLフォールバック） =========
def path_or_download_logo() -> str | None:
    if os.path.exists(LOGO_LOCAL):
        return LOGO_LOCAL
    try:
        for _ in range(2):
            r = requests.get(LOGO_URL, timeout=8)
            if r.ok:
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                tmp.write(r.content); tmp.flush()
                return tmp.name
            time.sleep(1.2)
    except Exception:
        pass
    return None

# ========= サイドバー =========
with st.sidebar:
    logo_path = path_or_download_logo()
    if logo_path:
        st.image(logo_path, width=160)
    st.markdown("### 3分無料診断")
    st.markdown("- 入力は Yes/部分的/No と 5段階のみ\n- 機密数値は不要\n- 結果は 6タイプ＋赤/黄/青")
    st.caption("© Victor Consulting")

st.title("製造現場の“隠れたムダ”をあぶり出す｜3分無料診断")
st.write("**10問**に回答するだけで、貴社のリスク“構造”を可視化します。")

# ========= 設問 UI =========
YN3 = ["Yes", "部分的に", "No"]
FIVE = ["5（非常にある）", "4", "3", "2", "1（まったくない）"]

with st.form("diagnose_form"):
    st.subheader("① 在庫・運搬（資金の滞留）")
    q1 = st.radio("Q1. 完成品・仕掛品の在庫基準を数値で管理していますか？", YN3, index=1)
    q2 = st.radio("Q2. 在庫削減の責任部署（またはKPI）が明確ですか？", YN3, index=1)

    st.subheader("② 人材・技能承継（属人化リスク）")
    q3 = st.radio("Q3. 熟練者しか対応できない作業が3割以上ありますか？（Yesはリスク高）", YN3, index=2)
    q4 = st.radio("Q4. 作業標準書・マニュアルを継続更新できる体制がありますか？", YN3, index=1)

    st.subheader("③ 原価意識・改善文化（損失体質）")
    q5 = st.radio("Q5. 改善提案や原価削減の目標を数値で追っていますか？", YN3, index=1)
    q6 = st.radio("Q6. 現場リーダーがコスト感覚を持って行動していますか？", FIVE, index=2)

    st.subheader("④ 生産計画・変動対応（流れの乱れ）")
    q7 = st.radio("Q7. 受注変動や突発対応の標準ルールがありますか？", YN3, index=1)
    q8 = st.radio("Q8. リードタイム短縮の取組を定期的に見直していますか？", YN3, index=1)

    st.subheader("⑤ DX・情報共有（見える化不足）")
    q9  = st.radio("Q9. 現場の進捗や生産実績をリアルタイムで把握できますか？", YN3, index=2)
    q10 = st.radio("Q10. データをもとに経営会議や現場ミーティングを行っていますか？", YN3, index=1)

    st.markdown("---")
    company = st.text_input("会社名（任意）", value=st.session_state["company"])
    email   = st.text_input("メールアドレス（任意｜Phase 4で利用）", value=st.session_state["email"])
    submitted = st.form_submit_button("診断する")

# ========= スコア関数 =========
def to_score_yn3(ans: str, invert=False) -> int:
    base = {"Yes": 5, "部分的に": 3, "No": 1}
    val = base.get(ans, 3)
    return {5: 1, 3: 3, 1: 5}[val] if invert else val

def to_score_5scale(ans: str) -> int:
    return int(ans[0])

# ========= 型・コメント（静的デフォルト） =========
TYPE_TEXT = {
    "在庫滞留型": "過剰在庫やWIP滞留で資金が眠っている可能性が高い状態です。生産量ではなく“流れ”の設計に軸足を移しましょう。",
    "熟練依存型": "属人化により技能がブラックボックス化。ベテラン離職に伴う急落リスクが高い状態です。技能棚卸と多能工化の設計が急務です。",
    "原価ブラックボックス型": "コスト意識・原価の見える化が弱く、利益が目減りする体質です。現場まで“見える原価管理”を展開しましょう。",
    "変動脆弱型": "受注変動・突発に弱く、納期トラブルや残業増に直結しています。変動を“なくす”のではなく“流す”バッファ設計が肝要です。",
    "データ断絶型": "進捗・実績が見えず、意思決定が遅れがちです。まずは“見える化”から。現場と経営のデータ接続を整備しましょう。",
    "バランス良好型": "リスク分散と仕組み成熟が進んでいます。次の一手は“利益を生むデータ活用”と継続的なリードタイム短縮です。"
}

# ========= 保存系（Sheets / CSV）=========
def try_append_to_google_sheets(row_dict: dict, spreadsheet_id: str, service_json_str: str):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    info = json.loads(service_json_str)
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.sheet1
    if not ws.get_all_values():
        ws.append_row(list(row_dict.keys()))
    ws.append_row([row_dict[k] for k in row_dict.keys()])

def fallback_append_to_csv(row_dict: dict, csv_path="responses.csv"):
    df = pd.DataFrame([row_dict])
    if os.path.exists(csv_path):
        df.to_csv(csv_path, mode="a", header=False, index=False, encoding="utf-8")
    else:
        df.to_csv(csv_path, index=False, encoding="utf-8")

# ========= PDFユーティリティ =========
def build_bar_png(df: pd.DataFrame) -> bytes:
    fig, ax = plt.subplots(figsize=(5.2, 2.6), dpi=220)
    df_sorted = df.sort_values("平均スコア", ascending=True)
    ax.barh(df_sorted["カテゴリ"], df_sorted["平均スコア"])
    ax.set_xlim(0, 5)
    ax.set_xlabel("平均スコア（0-5）")
    ax.grid(axis="x", linestyle="--", alpha=0.3)
    if FONT_PATH_IN_USE:
        from matplotlib import font_manager as fm
        fp = fm.FontProperties(fname=FONT_PATH_IN_USE)
        ax.set_xlabel("平均スコア（0-5）", fontproperties=fp)
        for label in ax.get_yticklabels(): label.set_fontproperties(fp)
        for label in ax.get_xticklabels(): label.set_fontproperties(fp)
    buf = io.BytesIO()
    fig.tight_layout()
    fig.savefig(buf, format="png")
    plt.close(fig); buf.seek(0)
    return buf.read()

def image_with_max_width(path: str, max_w: int):
    with PILImage.open(path) as im:
        w, h = im.size
    if w <= max_w:
        return Image(path, width=w, height=h)
    new_h = h * (max_w / w)
    return Image(path, width=max_w, height=new_h)

def build_qr_png(data_url: str) -> bytes:
    img = qrcode.make(data_url)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.read()

def make_pdf_bytes(result: dict, df_scores: pd.DataFrame, brand_hex=BRAND_BG) -> bytes:
    # ロゴ解決（ローカル優先）
    logo_path = path_or_download_logo()
    bar_png = build_bar_png(df_scores)
    qr_png  = build_qr_png(CTA_URL)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
    )

    styles = getSampleStyleSheet()
    title = styles["Title"]; normal = styles["BodyText"]; h3 = styles["Heading3"]
    if FONT_PATH_IN_USE:
        title.fontName = normal.fontName = h3.fontName = "JP"

    elems = []
    # ロゴ（縦横比維持）
    if logo_path:
        elems.append(image_with_max_width(logo_path, max_w=140))
        elems.append(Spacer(1, 8))

    elems.append(Paragraph("3分無料診断レポート", title))
    elems.append(Spacer(1, 6))
    meta = (
        f"会社名：{result['company'] or '（未入力）'}　/　"
        f"実施日時：{result['dt']}　/　"
        f"信号：{result['signal']}　/　"
        f"タイプ：{result['main_type']}"
    )
    elems.append(Paragraph(meta, normal))
    elems.append(Spacer(1, 8))

    elems.append(Paragraph("診断コメント", h3))
    elems.append(Paragraph(result["comment"], normal))
    elems.append(Spacer(1, 8))

    # 表
    table_data = [["カテゴリ", "平均スコア（0-5）"]] + [
        [r["カテゴリ"], f"{r['平均スコア']:.2f}"] for _, r in df_scores.iterrows()
    ]
    tbl = Table(table_data, colWidths=[220, 150])
    style_list = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(brand_hex)),
        ("TEXTCOLOR",  (0, 0), (-1, 0), colors.black),
        ("GRID",       (0, 0), (-1, -1), 0.3, colors.grey),
        ("ALIGN",      (1, 1), (-1, -1), "CENTER"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
    ]
    if FONT_PATH_IN_USE:
        style_list.append(("FONTNAME", (0, 0), (-1, -1), "JP"))
    tbl.setStyle(TableStyle(style_list))
    elems.append(tbl)
    elems.append(Spacer(1, 8))

    # 棒グラフ
    bar_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    bar_tmp.write(bar_png); bar_tmp.flush()
    elems.append(Paragraph("カテゴリ別スコア（棒グラフ）", h3))
    elems.append(Image(bar_tmp.name, width=420, height=210))
    elems.append(Spacer(1, 8))

    # 「次の一手」：左に文言、右にQR を横並びにするためTableを使用
    elems.append(Paragraph("次の一手（90分スポット診断のご案内）", h3))

    # 左セル（URL文言）
    url_par = Paragraph(f"詳細・お申込み：<u>{CTA_URL}</u>", normal)

    # 右セル（QR画像：やや小さめで1ページに収める）
    qr_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    qr_tmp.write(qr_png); qr_tmp.flush()
    qr_img = Image(qr_tmp.name, width=60, height=60)

    next_table = Table(
        [[url_par, qr_img]],
        colWidths=[430, 80]  # 左を広め、右にQR
    )
    nt_style = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",  (1, 0), (1, 0), "RIGHT"),
    ]
    if FONT_PATH_IN_USE:
        nt_style.append(("FONTNAME", (0, 0), (-1, -1), "JP"))
    next_table.setStyle(TableStyle(nt_style))
    elems.append(next_table)

    doc.build(elems)
    buf.seek(0)
    return buf.read()

# ========= OpenAI: AIコメント生成 =========
def _openai_client(api_key: str):
    try:
        from openai import OpenAI  # 新SDK
        return "new", OpenAI(api_key=api_key)
    except Exception:
        import openai  # 旧SDK
        openai.api_key = api_key
        return "old", openai

def generate_ai_comment(company: str, main_type: str, df_scores: pd.DataFrame, overall_avg: float):
    api_key = read_secret("OPENAI_API_KEY", None)
    if not api_key:
        return None, "OpenAIのAPIキーが未設定です（Settings→Secrets または環境変数に OPENAI_API_KEY を設定）。"
    worst2 = df_scores.sort_values("平均スコア", ascending=True).head(2)["カテゴリ"].tolist()
    user_prompt = f"""
あなたは元製造部長の経営コンサルタントです。以下の診断結果を受け、経営者向けに約300字（260〜340字）の具体的コメントを日本語で書いてください。箇条書きは使わず、1段落で、余計な前置きや免責は不要。最後は「90分スポット診断」での次アクションを自然に促す一文で締めます。

[会社名] {company or "（未入力）"}
[全体平均] {overall_avg:.2f} / 5
[信号] {"青" if overall_avg>=4.0 else ("黄" if overall_avg>=2.6 else "赤")}
[タイプ] {main_type}
[弱点カテゴリTOP2] {", ".join(worst2)}
[5カテゴリ] {", ".join(df_scores["カテゴリ"].tolist())}
""".strip()

    mode, client = _openai_client(api_key)
    try:
        if mode == "new":
            resp = client.chat.completions.create(
                model=OPENAI_MODEL,
                messages=[
                    {"role": "system", "content": "専門的かつ簡潔。日本語。実務に直結する助言を。"},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=0.4,
                max_tokens=420,
            )
            text = resp.choices[0].message.content.strip()
        else:
            resp = client.ChatCompletion.create(
                model=OPENAI_MODEL,
                messages=[
                    {"role": "system", "content": "専門的かつ簡潔。日本語。実務に直結する助言を。"},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=0.4,
                max_tokens=420,
            )
            text = resp.choices[0].message["content"].strip()
        return text, None
    except Exception as e:
        return None, f"AIコメント生成でエラー: {e}"

# ========= 計算＆セッション保存 =========
if submitted:
    inv_scores    = [to_score_yn3(q1), to_score_yn3(q2)]
    skills_scores = [to_score_yn3(q3, invert=True), to_score_yn3(q4)]
    cost_scores   = [to_score_yn3(q5), to_score_5scale(q6)]
    plan_scores   = [to_score_yn3(q7), to_score_yn3(q8)]
    dx_scores     = [to_score_yn3(q9), to_score_yn3(q10)]

    df = pd.DataFrame({
        "カテゴリ": ["在庫・運搬","人材・技能承継","原価意識・改善文化","生産計画・変動対応","DX・情報共有"],
        "平均スコア": [
            sum(inv_scores)/2,
            sum(skills_scores)/2,
            sum(cost_scores)/2,
            sum(plan_scores)/2,
            sum(dx_scores)/2
        ]
    })
    overall_avg = df["平均スコア"].mean()

    if overall_avg >= 4.0:
        signal = ("青信号", "badge-blue")
    elif overall_avg >= 2.6:
        signal = ("黄信号", "badge-yellow")
    else:
        signal = ("赤信号", "badge-red")

    if (df["平均スコア"] >= 4.0).all():
        main_type = "バランス良好型"
    else:
        worst_row = df.sort_values("平均スコア").iloc[0]
        cat = worst_row["カテゴリ"]
        main_type = {
            "在庫・運搬": "在庫滞留型",
            "人材・技能承継": "熟練依存型",
            "原価意識・改善文化": "原価ブラックボックス型",
            "生産計画・変動対応": "変動脆弱型",
            "DX・情報共有": "データ断絶型"
        }[cat]

    # セッションへ保存
    st.session_state.update({
        "df": df, "overall_avg": overall_avg, "signal": signal,
        "main_type": main_type, "company": company, "email": email,
        "result_ready": True
    })

# ========= 結果画面（セッションから表示） =========
if st.session_state.get("result_ready"):
    df = st.session_state["df"]
    overall_avg = st.session_state["overall_avg"]
    signal = st.session_state["signal"]
    main_type = st.session_state["main_type"]
    company = st.session_state["company"]
    email = st.session_state["email"]

    current_time = datetime.now(JST).strftime("%Y-%m-%d %H:%M")

    st.markdown("### 診断結果")
    st.markdown(
        f"""
        <div class="result-card">
            <h3 style="margin:0 0 .3rem 0;">
              タイプ判定：{main_type} <span class="badge {signal[1]}">{signal[0]}</span>
            </h3>
            <div class="small-note">
              会社名：{company or "（未入力）"} ／ 実施日時：{current_time}
            </div>
            <hr/>
            <p style="margin:.2rem 0 0 0;">{TYPE_TEXT[main_type]}</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    chart = (
        alt.Chart(df)
        .mark_bar()
        .encode(
            x=alt.X("平均スコア:Q", scale=alt.Scale(domain=[0, 5])),
            y=alt.Y("カテゴリ:N", sort="-x"),
            tooltip=["カテゴリ", "平均スコア"]
        ).properties(height=220)
    )
    st.altair_chart(chart, use_container_width=True)
    st.dataframe(df.style.format({"平均スコア": "{:.2f}"}), use_container_width=True)

    # ===== AIコメント生成 =====
    with st.expander("AIコメント（約300字）を自動生成する", expanded=False):
        colA, colB = st.columns([1,1])
        if colA.button("AIコメントを生成", use_container_width=True):
            text, err = generate_ai_comment(company, main_type, df, overall_avg)
            if err:
                st.error(err)
            else:
                st.session_state["ai_comment"] = text
                st.success("AIコメントを生成しました。下に表示しています。")

        if colB.button("AIコメントをクリア", use_container_width=True):
            st.session_state["ai_comment"] = None

        if st.session_state["ai_comment"]:
            st.write(st.session_state["ai_comment"])
        else:
            st.caption("（未生成）ボタンを押すと、診断内容に沿った約300字のコメントを生成します。")

    st.success("PDF出力・ログ保存が使えます（下のボタン群）。")

    # PDF: AIコメントがあれば優先して使う
    comment_for_pdf = st.session_state["ai_comment"] or TYPE_TEXT[main_type]
    result_payload = {
        "company": company,
        "email": email,
        "dt": current_time,  # JST
        "signal": signal[0],
        "main_type": main_type,
        "comment": comment_for_pdf
    }
    pdf_bytes = make_pdf_bytes(result_payload, df, brand_hex=BRAND_BG)
    fname = f"VC_診断_{company or '匿名'}_{datetime.now(JST).strftime('%Y%m%d_%H%M')}.pdf"
    st.download_button("📄 PDFをダウンロード", data=pdf_bytes, file_name=fname, mime="application/pdf")

    # ログ保存（Sheets or CSV）
    with st.expander("管理者向け：ログ保存（Google Sheets / CSV）"):
        st.write("※ Google Sheets のサービスアカウントJSONとスプレッドシートIDがあれば、直接保存できます。無い場合はCSVに追記します。")
        secret_json     = read_secret("GOOGLE_SERVICE_JSON", None)
        secret_sheet_id = read_secret("SPREADSHEET_ID", None)

        sheet_id = st.text_input("スプレッドシートID（1A2B... の長いID）", value=secret_sheet_id or "")
        json_text = st.text_area("サービスアカウントJSON（貼り付け）", value=secret_json or "", height=140)

        row = {
            "timestamp": datetime.now(JST).isoformat(timespec="seconds"),
            "company": company, "email": email,
            "signal": signal[0], "main_type": main_type,
            "overall_avg": f"{overall_avg:.2f}",
            "inv_avg": f"{df.loc[df['カテゴリ']=='在庫・運搬','平均スコア'].values[0]:.2f}",
            "skills_avg": f"{df.loc[df['カテゴリ']=='人材・技能承継','平均スコア'].values[0]:.2f}",
            "cost_avg": f"{df.loc[df['カテゴリ']=='原価意識・改善文化','平均スコア'].values[0]:.2f}",
            "plan_avg": f"{df.loc[df['カテゴリ']=='生産計画・変動対応','平均スコア'].values[0]:.2f}",
            "dx_avg": f"{df.loc[df['カテゴリ']=='DX・情報共有','平均スコア'].values[0]:.2f}",
            "ai_comment": st.session_state["ai_comment"] or ""
        }

        col1, col2 = st.columns(2)
        if col1.button("Google Sheetsに保存"):
            try:
                if sheet_id and json_text:
                    try_append_to_google_sheets(row, sheet_id, json_text)
                    st.success("Google Sheetsに保存しました。")
                else:
                    st.warning("スプレッドシートID と サービスアカウントJSON を入力してください。")
            except Exception as e:
                st.error(f"Sheets保存でエラー：{e}")

        if col2.button("CSVに保存（responses.csv）"):
            try:
                fallback_append_to_csv(row)
                st.success("CSVに追記しました（アプリ直下の responses.csv）。")
            except Exception as e:
                st.error(f"CSV保存でエラー：{e}")

else:
    st.caption("フォームに回答し、「診断する」を押してください。")







