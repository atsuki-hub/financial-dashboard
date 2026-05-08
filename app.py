# -*- coding: utf-8 -*-
import os
import sys
import json
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from io import StringIO
from dotenv import load_dotenv
import anthropic
import fitz

load_dotenv()

try:
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    if "ANTHROPIC_API_KEY" in st.secrets:
        os.environ["ANTHROPIC_API_KEY"] = st.secrets["ANTHROPIC_API_KEY"]
except Exception:
    pass

st.set_page_config(
    page_title="CFO Cockpit | 中小企業向け財務分析",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==============================
# カスタムCSS（プロフェッショナルなトーン）
# ==============================
st.markdown("""
<style>
    .main .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1400px; }
    h1 { font-weight: 600; letter-spacing: -0.02em; color: #F3F4F6; }
    h2, h3 { font-weight: 500; color: #E5E7EB; }
    .stTabs [data-baseweb="tab-list"] {
        gap: 0; border-bottom: 1px solid #2A2F3A; background: transparent;
    }
    .stTabs [data-baseweb="tab"] {
        height: 48px; padding: 0 24px; background: transparent;
        color: #9CA3AF; font-weight: 500; border-radius: 0;
    }
    .stTabs [aria-selected="true"] {
        color: #5EAFFF; border-bottom: 2px solid #5EAFFF;
    }
    [data-testid="stMetric"] {
        background: #1A1F2B; padding: 16px 20px; border-radius: 4px;
        border-left: 3px solid #5EAFFF;
    }
    [data-testid="stMetricLabel"] { color: #9CA3AF; font-size: 0.85rem; }
    [data-testid="stMetricValue"] { color: #F3F4F6; font-weight: 600; }
    .lp-hero {
        background: linear-gradient(135deg, #1A1F2B 0%, #141821 100%);
        padding: 32px 40px; border-radius: 6px; margin-bottom: 24px;
        border: 1px solid #2A2F3A;
    }
    .lp-tag {
        display: inline-block; padding: 4px 12px; background: #1F3A5F;
        color: #5EAFFF; font-size: 0.75rem; border-radius: 3px;
        letter-spacing: 0.05em; margin-bottom: 12px;
    }
    .stButton > button {
        background: #5EAFFF; color: #0F1419; border: none; font-weight: 600;
    }
    .stButton > button:hover { background: #4A95E0; color: #0F1419; }
    .stDownloadButton > button { background: #2A2F3A; color: #E5E7EB; }
    section[data-testid="stSidebar"] { background: #141821; border-right: 1px solid #2A2F3A; }
    .stAlert { border-radius: 4px; }
</style>
""", unsafe_allow_html=True)

# ==============================
# ヒーロー
# ==============================
st.markdown("""
<div class="lp-hero">
  <span class="lp-tag">CFO COCKPIT</span>
  <h1 style="margin:0; font-size:2rem;">中小企業オーナーのための財務分析ダッシュボード</h1>
  <p style="color:#9CA3AF; margin-top:8px; font-size:1rem;">
    年商1〜5億円の事業に最適化。試算表PDFをアップロードするだけで、
    KPI診断・トレンド分析・経営シミュレーションを一気通貫で行います。
  </p>
</div>
""", unsafe_allow_html=True)

# ==============================
# サンプルデータ
# ==============================
SAMPLE_ACTUAL = """勘定科目,金額
売上高,180000000
売上原価,108000000
販売費及び一般管理費,48000000
営業外収益,800000
営業外費用,500000
特別利益,0
特別損失,0
流動資産,72000000
固定資産,45000000
流動負債,38000000
固定負債,22000000
純資産,57000000
売上債権,28000000
棚卸資産,15000000
仕入債務,18000000
"""

SAMPLE_MONTHLY = {
    1:  {"売上高": 13500000, "売上原価":  8100000, "販売費及び一般管理費": 3700000},
    2:  {"売上高": 14200000, "売上原価":  8500000, "販売費及び一般管理費": 3800000},
    3:  {"売上高": 16800000, "売上原価": 10000000, "販売費及び一般管理費": 4100000},
    4:  {"売上高": 15100000, "売上原価":  9100000, "販売費及び一般管理費": 3900000},
    5:  {"売上高": 15900000, "売上原価":  9500000, "販売費及び一般管理費": 4000000},
    6:  {"売上高": 16500000, "売上原価":  9900000, "販売費及び一般管理費": 4050000},
    7:  {"売上高": 17200000, "売上原価": 10300000, "販売費及び一般管理費": 4200000},
    8:  {"売上高": 14800000, "売上原価":  8900000, "販売費及び一般管理費": 3950000},
    9:  {"売上高": 15700000, "売上原価":  9400000, "販売費及び一般管理費": 4050000},
    10: {"売上高": 16100000, "売上原価":  9700000, "販売費及び一般管理費": 4100000},
    11: {"売上高": 17500000, "売上原価": 10500000, "販売費及び一般管理費": 4250000},
    12: {"売上高": 18900000, "売上原価": 11300000, "販売費及び一般管理費": 4400000},
}

INDUSTRY_BENCHMARKS = {
    "IT・ソフトウェア・SaaS": {
        "売上総利益率": (50, 70), "営業利益率": (10, 20), "流動比率": 150, "自己資本比率": 40,
        "特徴": "在庫を持たないため粗利率が高い。人件費が主なコスト。",
    },
    "コンサルティング・士業": {
        "売上総利益率": (60, 80), "営業利益率": (15, 25), "流動比率": 150, "自己資本比率": 50,
        "特徴": "原価が人件費中心で高利益率を実現可能。稼働率が要。",
    },
    "卸売・商社": {
        "売上総利益率": (15, 25), "営業利益率": (2, 5), "流動比率": 130, "自己資本比率": 25,
        "特徴": "薄利多売。在庫回転と運転資本管理が生命線。",
    },
    "製造業": {
        "売上総利益率": (20, 35), "営業利益率": (5, 10), "流動比率": 150, "自己資本比率": 35,
        "特徴": "設備投資と運転資本のバランスが課題。",
    },
    "建設・工事": {
        "売上総利益率": (15, 25), "営業利益率": (3, 8), "流動比率": 130, "自己資本比率": 25,
        "特徴": "受注産業のため期ズレが大きい。工事ごとの採算管理が重要。",
    },
    "サービス業（一般）": {
        "売上総利益率": (40, 60), "営業利益率": (8, 15), "流動比率": 120, "自己資本比率": 30,
        "特徴": "人件費率が高い。生産性指標を併せて管理する。",
    },
}

# ==============================
# ヘルパー関数
# ==============================
def extract_financials_from_pdf(pdf_file) -> pd.Series:
    pdf_bytes = pdf_file.read()
    table_text = ""
    fallback_text = ""

    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page_num, page in enumerate(doc):
            try:
                tabs = page.find_tables()
                if tabs.tables:
                    for table in tabs.tables:
                        rows = table.extract()
                        if not rows:
                            continue
                        table_text += f"\n--- ページ{page_num + 1} 表 ---\n"
                        for row in rows:
                            cells = [str(c).strip() if c is not None else "" for c in row]
                            table_text += " | ".join(cells) + "\n"
            except Exception:
                pass
            try:
                page_text = page.get_text("text") or ""
                fallback_text += page_text + "\n"
            except Exception:
                pass

    raw_input = table_text if table_text.strip() else fallback_text
    text = raw_input.encode("utf-8", errors="replace").decode("utf-8")

    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    prompt = (
        "以下は財務書類のPDFから抽出したテーブルデータです。\n"
        "損益計算書と貸借対照表の数値を読み取り、必ず下記JSONのみを返してください。\n"
        "最初の文字を { にしてください。説明文・コードブロックは不要です。\n\n"
        "【ルール】\n"
        "- 金額は円単位の整数（千円→×1000、百万円→×1000000）\n"
        "- 不明・該当なしは0\n"
        "- 売掛金=売上債権、買掛金=仕入債務、販管費=販売費及び一般管理費\n"
        "- 売上原価未記載: 売上高 - 売上総利益\n"
        "- 販管費未記載: 売上総利益 - 営業利益\n"
        "- 複数期間ある場合は最新期\n\n"
        '{"売上高":0,"売上原価":0,"販売費及び一般管理費":0,'
        '"営業外収益":0,"営業外費用":0,"特別利益":0,"特別損失":0,'
        '"流動資産":0,"固定資産":0,"流動負債":0,"固定負債":0,"純資産":0,'
        '"売上債権":0,"棚卸資産":0,"仕入債務":0}\n\n'
        "--- データ ---\n" + text[:20000]
    )
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = message.content[0].text.strip()

    if "```" in raw:
        for part in raw.split("```"):
            part = part.strip().lstrip("json").strip()
            if part.startswith("{"):
                raw = part
                break

    if not raw.startswith("{"):
        start = raw.find("{")
        end = raw.rfind("}") + 1
        if start != -1 and end > start:
            raw = raw[start:end]
        else:
            raise ValueError(f"JSON抽出失敗: {message.content[0].text[:200]}")

    extracted = json.loads(raw)
    defaults = {
        "売上高":0,"売上原価":0,"販売費及び一般管理費":0,
        "営業外収益":0,"営業外費用":0,"特別利益":0,"特別損失":0,
        "流動資産":0,"固定資産":0,"流動負債":0,"固定負債":0,"純資産":0,
        "売上債権":0,"棚卸資産":0,"仕入債務":0,
    }
    defaults.update({k: float(v) for k, v in extracted.items()})
    return pd.Series(defaults)


def load_data(source):
    df = pd.read_csv(source)
    df.columns = ["勘定科目", "金額"]
    df["金額"] = pd.to_numeric(df["金額"], errors="coerce").fillna(0)
    return df.set_index("勘定科目")["金額"]


def get(data, key, default=0):
    return data.get(key, default)


def calc_kpis(d):
    s = get(d, "売上高")
    cogs = get(d, "売上原価")
    sga = get(d, "販売費及び一般管理費")
    oi = get(d, "営業外収益")
    oe = get(d, "営業外費用")
    ca = get(d, "流動資産"); fa = get(d, "固定資産")
    cl = get(d, "流動負債"); fl = get(d, "固定負債")
    eq = get(d, "純資産")
    ar = get(d, "売上債権"); inv = get(d, "棚卸資産"); ap = get(d, "仕入債務")

    gp = s - cogs
    op = gp - sga
    ord_p = op + oi - oe
    ta = ca + fa
    pct = lambda a, b: round(a / b * 100, 1) if b else None
    days = lambda a, b: round(a / b * 365, 1) if b else None

    return {
        "売上高": s, "売上総利益": gp, "営業利益": op, "経常利益": ord_p,
        "売上総利益率": pct(gp, s), "営業利益率": pct(op, s), "経常利益率": pct(ord_p, s),
        "流動比率": pct(ca, cl), "自己資本比率": pct(eq, ta),
        "売上債権回転日数": days(ar, s), "棚卸資産回転日数": days(inv, s),
        "仕入債務回転日数": days(ap, cogs),
        "総資産": ta, "純資産": eq,
    }


def fmt_yen(v):
    if v is None: return "—"
    if abs(v) >= 1_0000_0000: return f"{v/1_0000_0000:.2f}億円"
    if abs(v) >= 10000: return f"{v/10000:,.0f}万円"
    return f"{v:,.0f}円"


def evaluate_kpi(value, benchmark_range):
    """ベンチマークとの差で評価（0-100）"""
    if value is None: return 0
    lo, hi = benchmark_range
    if value >= hi: return 100
    if value >= lo: return 70
    if value >= lo * 0.7: return 50
    return 25


def health_score(kpis, bm):
    gp_score = evaluate_kpi(kpis["売上総利益率"], bm["売上総利益率"])
    op_score = evaluate_kpi(kpis["営業利益率"], bm["営業利益率"])
    cr = kpis["流動比率"] or 0
    cr_score = 100 if cr >= bm["流動比率"] else max(0, int(cr / bm["流動比率"] * 100))
    er = kpis["自己資本比率"] or 0
    er_score = 100 if er >= bm["自己資本比率"] else max(0, int(er / bm["自己資本比率"] * 100))

    profitability = int((gp_score + op_score) / 2)
    safety = int((cr_score + er_score) / 2)
    total = int(profitability * 0.6 + safety * 0.4)
    return {"総合": total, "収益性": profitability, "安全性": safety}


def generate_ai_comment(kpis, industry, bm, memo=""):
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    kpi_text = "\n".join([
        f"売上高: {fmt_yen(kpis['売上高'])}",
        f"営業利益: {fmt_yen(kpis['営業利益'])} (率: {kpis['営業利益率']}%, 業種目安: {bm['営業利益率'][0]}〜{bm['営業利益率'][1]}%)",
        f"売上総利益率: {kpis['売上総利益率']}% (業種目安: {bm['売上総利益率'][0]}〜{bm['売上総利益率'][1]}%)",
        f"流動比率: {kpis['流動比率']}% (目安: {bm['流動比率']}%以上)",
        f"自己資本比率: {kpis['自己資本比率']}% (目安: {bm['自己資本比率']}%以上)",
        f"売上債権回転日数: {kpis['売上債権回転日数']}日",
    ])
    prompt = (
        f"あなたは中小企業オーナー（年商1〜5億円規模）専属のCFOアドバイザーです。\n"
        f"業種: {industry}（{bm['特徴']}）\n\n"
        "以下のKPIから、3〜4分で読める経営者向けレポートを書いてください。\n\n"
        "## エグゼクティブサマリー\n（2文。一番重要な所見と次のアクションを端的に）\n\n"
        "## 強み\n（箇条書き2点）\n\n"
        "## 課題と打ち手\n（箇条書き2〜3点。具体的な数値目標を含めて）\n\n"
        + (f"【経営者からの補足】\n{memo}\n\n" if memo.strip() else "")
        + "【KPI】\n" + kpi_text
    )
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1500,
        messages=[{"role": "user", "content": prompt}],
    )
    return msg.content[0].text


# ==============================
# タブ
# ==============================
tab1, tab2, tab3 = st.tabs(["KPI診断", "月次推移分析", "What-if シミュレーション"])

# ==============================
# TAB 1: KPI診断
# ==============================
with tab1:
    with st.sidebar:
        st.markdown("### データ入力")
        input_method = st.radio(
            "入力方法", ["CSV アップロード", "PDF（AI読み取り）", "サンプルデータ"],
            label_visibility="collapsed",
        )
        uploaded_file = None
        pdf_file = None
        if input_method == "CSV アップロード":
            with st.expander("CSVフォーマット"):
                st.code("勘定科目,金額\n売上高,180000000\n...", language="text")
            uploaded_file = st.file_uploader("CSV", type=["csv"], key="csv_upload",
                                             label_visibility="collapsed")
        elif input_method == "PDF（AI読み取り）":
            st.caption("freee / マネーフォワードの試算表PDFに対応")
            pdf_file = st.file_uploader("PDF", type=["pdf"], key="pdf_upload",
                                        label_visibility="collapsed")
            if pdf_file and st.button("AIで読み取る", use_container_width=True):
                with st.spinner("AI解析中..."):
                    try:
                        st.session_state["pdf_extracted"] = extract_financials_from_pdf(pdf_file)
                        st.success("読み取り完了")
                    except Exception as e:
                        st.error(f"エラー: {e}")
        else:
            st.caption("年商1.8億円の卸売・サービス業を想定したサンプル")

        st.divider()
        st.markdown("### 業種")
        industry = st.selectbox(
            "業種",
            list(INDUSTRY_BENCHMARKS.keys()),
            label_visibility="collapsed",
        )
        bm = INDUSTRY_BENCHMARKS[industry]
        st.caption(bm["特徴"])

    data = None
    if input_method == "PDF（AI読み取り）" and st.session_state.get("pdf_extracted") is not None:
        data = st.session_state["pdf_extracted"]
    elif input_method == "CSV アップロード" and uploaded_file:
        data = load_data(uploaded_file)
    elif input_method == "サンプルデータ":
        data = load_data(StringIO(SAMPLE_ACTUAL))

    if data is None:
        st.info("サイドバーからデータを入力してください。試しに「サンプルデータ」を選ぶと、すぐに動作を確認できます。")
    else:
        kpis = calc_kpis(data)
        scores = health_score(kpis, bm)

        # ====== 健全度スコア ======
        st.markdown("#### 財務健全度スコア")
        sc1, sc2, sc3 = st.columns([1.2, 1, 1])
        with sc1:
            color = "#10B981" if scores["総合"] >= 75 else "#F59E0B" if scores["総合"] >= 50 else "#EF4444"
            fig_g = go.Figure(go.Indicator(
                mode="gauge+number", value=scores["総合"],
                title={"text": "総合スコア", "font": {"size": 14, "color": "#9CA3AF"}},
                number={"font": {"size": 48, "color": color}},
                gauge={
                    "axis": {"range": [0, 100], "tickcolor": "#4B5563"},
                    "bar": {"color": color, "thickness": 0.3},
                    "bgcolor": "#1A1F2B", "borderwidth": 0,
                    "steps": [
                        {"range": [0, 50], "color": "#2D1B1F"},
                        {"range": [50, 75], "color": "#2D2519"},
                        {"range": [75, 100], "color": "#1A2D24"},
                    ],
                },
            ))
            fig_g.update_layout(height=220, margin=dict(t=30, b=10, l=10, r=10),
                                paper_bgcolor="rgba(0,0,0,0)", font={"color": "#E5E7EB"})
            st.plotly_chart(fig_g, use_container_width=True)
        with sc2:
            st.metric("収益性", f"{scores['収益性']}", help="粗利率・営業利益率の業種比較")
            st.caption("売上に対する利益の効率")
        with sc3:
            st.metric("安全性", f"{scores['安全性']}", help="流動比率・自己資本比率")
            st.caption("資金繰りと財務の堅牢性")

        st.divider()

        # ====== 損益サマリー ======
        st.markdown("#### 損益サマリー")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("売上高", fmt_yen(kpis["売上高"]))
        c2.metric("売上総利益", fmt_yen(kpis["売上総利益"]), f"{kpis['売上総利益率']}%")
        c3.metric("営業利益", fmt_yen(kpis["営業利益"]), f"{kpis['営業利益率']}%")
        c4.metric("経常利益", fmt_yen(kpis["経常利益"]), f"{kpis['経常利益率']}%")

        # ====== KPIテーブル＋ベンチマーク ======
        st.markdown("#### KPI vs 業種ベンチマーク")
        kpi_rows = [
            ("売上総利益率", f"{kpis['売上総利益率']}%", f"{bm['売上総利益率'][0]}〜{bm['売上総利益率'][1]}%"),
            ("営業利益率",   f"{kpis['営業利益率']}%",   f"{bm['営業利益率'][0]}〜{bm['営業利益率'][1]}%"),
            ("流動比率",     f"{kpis['流動比率']}%",     f"{bm['流動比率']}%以上"),
            ("自己資本比率", f"{kpis['自己資本比率']}%", f"{bm['自己資本比率']}%以上"),
            ("売上債権回転日数", f"{kpis['売上債権回転日数']}日", "短いほど良"),
            ("仕入債務回転日数", f"{kpis['仕入債務回転日数']}日", "長いほど良"),
        ]
        df_kpi = pd.DataFrame(kpi_rows, columns=["指標", "実績", "業種目安"])
        st.dataframe(df_kpi, use_container_width=True, hide_index=True)

        st.divider()

        # ====== ビジュアル ======
        col_l, col_r = st.columns(2)
        with col_l:
            st.markdown("#### 利益構造（ウォーターフォール）")
            fig_w = go.Figure(go.Waterfall(
                orientation="v",
                measure=["absolute", "relative", "total", "relative", "total"],
                x=["売上高", "売上原価", "売上総利益", "販管費", "営業利益"],
                y=[kpis["売上高"], -get(data, "売上原価"), kpis["売上総利益"],
                   -get(data, "販売費及び一般管理費"), kpis["営業利益"]],
                connector={"line": {"color": "#4B5563"}},
                increasing={"marker": {"color": "#5EAFFF"}},
                decreasing={"marker": {"color": "#EF6B6B"}},
                totals={"marker": {"color": "#10B981"}},
            ))
            fig_w.update_layout(
                showlegend=False, height=350,
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#1A1F2B",
                font={"color": "#E5E7EB"},
                xaxis={"gridcolor": "#2A2F3A"}, yaxis={"gridcolor": "#2A2F3A"},
                margin=dict(t=20, b=20, l=10, r=10),
            )
            st.plotly_chart(fig_w, use_container_width=True)

        with col_r:
            st.markdown("#### B/S 構成")
            fig_bs = go.Figure()
            fig_bs.add_trace(go.Bar(name="流動資産", x=["資産"], y=[get(data, "流動資産")],
                                     marker_color="#5EAFFF"))
            fig_bs.add_trace(go.Bar(name="固定資産", x=["資産"], y=[get(data, "固定資産")],
                                     marker_color="#3B7DD8"))
            fig_bs.add_trace(go.Bar(name="流動負債", x=["負債・資本"], y=[get(data, "流動負債")],
                                     marker_color="#EF6B6B"))
            fig_bs.add_trace(go.Bar(name="固定負債", x=["負債・資本"], y=[get(data, "固定負債")],
                                     marker_color="#B8484A"))
            fig_bs.add_trace(go.Bar(name="純資産",   x=["負債・資本"], y=[get(data, "純資産")],
                                     marker_color="#10B981"))
            fig_bs.update_layout(
                barmode="stack", height=350,
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#1A1F2B",
                font={"color": "#E5E7EB"},
                xaxis={"gridcolor": "#2A2F3A"}, yaxis={"gridcolor": "#2A2F3A"},
                margin=dict(t=20, b=20, l=10, r=10),
                legend={"bgcolor": "rgba(0,0,0,0)"},
            )
            st.plotly_chart(fig_bs, use_container_width=True)

        st.divider()

        # ====== AIコメント ======
        st.markdown("#### CFOコメント（AI生成）")
        memo = st.text_area(
            "経営状況の補足（任意）",
            placeholder="例: 主力顧客が1社撤退した／新規プロダクトの先行投資中",
            height=70, key="kpi_memo",
        )
        if st.button("AIにレポートを生成させる", key="gen_kpi"):
            with st.spinner("CFOが分析中..."):
                try:
                    st.session_state["kpi_comment"] = generate_ai_comment(
                        kpis, industry, bm, memo
                    )
                except Exception as e:
                    st.error(f"生成エラー: {e}")
        if st.session_state.get("kpi_comment"):
            st.markdown(st.session_state["kpi_comment"])


# ==============================
# TAB 2: 月次推移分析
# ==============================
with tab2:
    st.markdown("#### 月次推移分析")
    st.caption("12ヶ月分の月次データから、トレンド・季節性・異常値を可視化します。")

    use_sample_monthly = st.toggle("サンプル月次データを使用", value=True, key="toggle_monthly")

    if use_sample_monthly:
        monthly_data = SAMPLE_MONTHLY
    else:
        st.markdown("**月次データを入力**")
        rows = []
        cols = st.columns(4)
        for i in range(1, 13):
            col = cols[(i - 1) % 4]
            with col:
                rev = st.number_input(f"{i}月 売上高", value=int(SAMPLE_MONTHLY[i]["売上高"]),
                                       step=100000, key=f"rev_{i}")
                cogs = st.number_input(f"{i}月 売上原価", value=int(SAMPLE_MONTHLY[i]["売上原価"]),
                                        step=100000, key=f"cogs_{i}")
                sga = st.number_input(f"{i}月 販管費", value=int(SAMPLE_MONTHLY[i]["販売費及び一般管理費"]),
                                       step=100000, key=f"sga_{i}")
                rows.append((i, rev, cogs, sga))
        monthly_data = {i: {"売上高": r, "売上原価": c, "販売費及び一般管理費": s}
                        for i, r, c, s in rows}

    months = sorted(monthly_data.keys())
    revenue = [monthly_data[m]["売上高"] for m in months]
    cogs_arr = [monthly_data[m]["売上原価"] for m in months]
    sga_arr = [monthly_data[m]["販売費及び一般管理費"] for m in months]
    op_profit = [r - c - s for r, c, s in zip(revenue, cogs_arr, sga_arr)]

    df_m = pd.DataFrame({
        "月": [f"{m}月" for m in months],
        "売上高": revenue, "売上原価": cogs_arr, "販管費": sga_arr,
        "営業利益": op_profit,
    })

    # ====== サマリー ======
    total_rev = sum(revenue)
    total_op = sum(op_profit)
    avg_rev = total_rev / len(months)
    op_margin = (total_op / total_rev * 100) if total_rev else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("年間売上", fmt_yen(total_rev))
    c2.metric("年間営業利益", fmt_yen(total_op), f"{op_margin:.1f}%")
    c3.metric("月平均売上", fmt_yen(avg_rev))
    c4.metric("最高/最低月", f"{months[revenue.index(max(revenue))]}月 / {months[revenue.index(min(revenue))]}月")

    st.divider()

    # ====== 推移グラフ ======
    st.markdown("#### 売上・利益の推移")

    # 移動平均
    ma3 = pd.Series(revenue).rolling(3, min_periods=1).mean()

    fig_t = go.Figure()
    fig_t.add_trace(go.Bar(x=df_m["月"], y=revenue, name="売上高",
                            marker_color="#5EAFFF", opacity=0.8))
    fig_t.add_trace(go.Scatter(x=df_m["月"], y=ma3, name="3ヶ月移動平均",
                                line={"color": "#F59E0B", "width": 2, "dash": "dot"},
                                mode="lines+markers"))
    fig_t.add_trace(go.Scatter(x=df_m["月"], y=op_profit, name="営業利益",
                                line={"color": "#10B981", "width": 2.5},
                                mode="lines+markers", yaxis="y2"))
    fig_t.update_layout(
        height=420,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#1A1F2B",
        font={"color": "#E5E7EB"},
        xaxis={"gridcolor": "#2A2F3A"},
        yaxis={"title": "売上高（円）", "gridcolor": "#2A2F3A"},
        yaxis2={"title": "営業利益（円）", "overlaying": "y", "side": "right",
                "gridcolor": "rgba(0,0,0,0)"},
        legend={"bgcolor": "rgba(0,0,0,0)", "orientation": "h", "y": 1.1},
        margin=dict(t=40, b=20, l=10, r=10),
    )
    st.plotly_chart(fig_t, use_container_width=True)

    # ====== 異常値検知 ======
    avg = sum(revenue) / len(revenue)
    std = (sum((r - avg) ** 2 for r in revenue) / len(revenue)) ** 0.5
    anomalies = [(months[i], revenue[i]) for i in range(len(revenue))
                  if abs(revenue[i] - avg) > 1.5 * std]

    if anomalies:
        st.markdown("#### 異常値の検知")
        for m, v in anomalies:
            direction = "↑ 高水準" if v > avg else "↓ 低水準"
            st.warning(f"**{m}月**: {fmt_yen(v)} ({direction} / 平均比 {(v/avg-1)*100:+.1f}%)")

    # ====== 着地予測 ======
    st.divider()
    st.markdown("#### 着地予測")
    st.caption("直近3ヶ月の平均から、年間着地を試算します。")

    if len(months) >= 3:
        recent_avg_rev = sum(revenue[-3:]) / 3
        recent_avg_op = sum(op_profit[-3:]) / 3
        forecast_rev = sum(revenue) + recent_avg_rev * (12 - len(months)) if len(months) < 12 else sum(revenue)
        forecast_op = sum(op_profit) + recent_avg_op * (12 - len(months)) if len(months) < 12 else sum(op_profit)

        f1, f2, f3 = st.columns(3)
        f1.metric("年間売上見込み", fmt_yen(forecast_rev))
        f2.metric("年間営業利益見込み", fmt_yen(forecast_op))
        f3.metric("予想営業利益率", f"{(forecast_op/forecast_rev*100) if forecast_rev else 0:.1f}%")


# ==============================
# TAB 3: What-if シミュレーション
# ==============================
with tab3:
    st.markdown("#### What-if シミュレーション")
    st.caption("売上・コストが変動したとき、利益がどう変わるかをリアルタイムで可視化します。経営判断の意思決定に。")

    base_data = load_data(StringIO(SAMPLE_ACTUAL))
    base_kpis = calc_kpis(base_data)

    st.markdown("##### 変動シナリオ")
    sim_c1, sim_c2, sim_c3 = st.columns(3)
    with sim_c1:
        rev_change = st.slider("売上変化率（%）", -30, 50, 0, step=1, key="sim_rev")
    with sim_c2:
        cogs_change = st.slider("変動費率の変化（pt）", -10, 10, 0, step=1, key="sim_cogs",
                                help="売上原価率の変化幅。マイナスは原価改善。")
    with sim_c3:
        sga_change = st.slider("固定費変化率（%）", -20, 30, 0, step=1, key="sim_sga")

    # 計算
    new_rev = base_kpis["売上高"] * (1 + rev_change / 100)
    base_cogs_rate = get(base_data, "売上原価") / base_kpis["売上高"]
    new_cogs_rate = base_cogs_rate + cogs_change / 100
    new_cogs = new_rev * new_cogs_rate
    new_sga = get(base_data, "販売費及び一般管理費") * (1 + sga_change / 100)
    new_gp = new_rev - new_cogs
    new_op = new_gp - new_sga

    # ====== 結果サマリー ======
    st.divider()
    st.markdown("#### シミュレーション結果")

    res_c1, res_c2, res_c3 = st.columns(3)
    diff_rev = new_rev - base_kpis["売上高"]
    diff_op = new_op - base_kpis["営業利益"]
    new_op_margin = (new_op / new_rev * 100) if new_rev else 0

    res_c1.metric("シミュ後 売上", fmt_yen(new_rev),
                  f"{diff_rev/10000:+,.0f}万円 vs 現状")
    res_c2.metric("シミュ後 営業利益", fmt_yen(new_op),
                  f"{diff_op/10000:+,.0f}万円 vs 現状")
    res_c3.metric("営業利益率", f"{new_op_margin:.1f}%",
                  f"{new_op_margin - base_kpis['営業利益率']:+.1f}pt")

    # ====== 比較グラフ ======
    st.markdown("#### 現状 vs シミュレーション")
    fig_c = go.Figure()
    items = ["売上高", "売上総利益", "営業利益"]
    base_vals = [base_kpis["売上高"], base_kpis["売上総利益"], base_kpis["営業利益"]]
    sim_vals = [new_rev, new_gp, new_op]

    fig_c.add_trace(go.Bar(name="現状", x=items, y=base_vals, marker_color="#4B5563",
                            text=[fmt_yen(v) for v in base_vals], textposition="outside"))
    fig_c.add_trace(go.Bar(name="シミュレーション", x=items, y=sim_vals, marker_color="#5EAFFF",
                            text=[fmt_yen(v) for v in sim_vals], textposition="outside"))
    fig_c.update_layout(
        barmode="group", height=400,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#1A1F2B",
        font={"color": "#E5E7EB"},
        xaxis={"gridcolor": "#2A2F3A"}, yaxis={"gridcolor": "#2A2F3A"},
        legend={"bgcolor": "rgba(0,0,0,0)"},
        margin=dict(t=40, b=20, l=10, r=10),
    )
    st.plotly_chart(fig_c, use_container_width=True)

    # ====== シナリオ感度分析 ======
    st.divider()
    st.markdown("#### 感度分析：売上変化が利益に与える影響")
    sens_x = list(range(-30, 51, 5))
    sens_y = []
    for rc in sens_x:
        r = base_kpis["売上高"] * (1 + rc / 100)
        c = r * new_cogs_rate
        s = get(base_data, "販売費及び一般管理費") * (1 + sga_change / 100)
        sens_y.append(r - c - s)

    fig_s = go.Figure()
    fig_s.add_trace(go.Scatter(
        x=sens_x, y=sens_y, mode="lines+markers",
        line={"color": "#5EAFFF", "width": 2.5},
        marker={"size": 6},
    ))
    fig_s.add_hline(y=0, line_dash="dash", line_color="#EF6B6B",
                     annotation_text="損益分岐点", annotation_position="bottom right")
    fig_s.add_vline(x=rev_change, line_dash="dot", line_color="#F59E0B",
                     annotation_text="現在のシナリオ")
    fig_s.update_layout(
        height=380,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#1A1F2B",
        font={"color": "#E5E7EB"},
        xaxis={"title": "売上変化率（%）", "gridcolor": "#2A2F3A"},
        yaxis={"title": "営業利益（円）", "gridcolor": "#2A2F3A"},
        margin=dict(t=30, b=20, l=10, r=10),
    )
    st.plotly_chart(fig_s, use_container_width=True)

# ==============================
# フッター
# ==============================
st.markdown("""
<div style="margin-top:60px; padding-top:20px; border-top:1px solid #2A2F3A;
            color:#6B7280; font-size:0.85rem; text-align:center;">
  CFO Cockpit · Built with Streamlit + Claude API · MIT License
</div>
""", unsafe_allow_html=True)
