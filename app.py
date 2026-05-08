# -*- coding: utf-8 -*-
import os
import sys
import json
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from io import StringIO, BytesIO
from datetime import datetime, date, timedelta
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
  <h1 style="margin:0; font-size:2rem;">経営者の「眠れない夜」を解消する資金繰りダッシュボード</h1>
  <p style="color:#9CA3AF; margin-top:8px; font-size:1rem;">
    年商1〜5億円の中小企業オーナー向け。3ヶ月先までの資金ショートを先回りで検知し、
    原因と打ち手まで提示。会計ソフトでは届かない「経営判断のラストワンマイル」を埋めます。
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
# 資金繰り：サンプルデータ
# ==============================
def _today():
    return datetime.now().date()

def _make_sample_cashflow():
    t = _today()
    receivables = pd.DataFrame([
        {"取引先": "A社", "金額": 800000,  "予定日": t + timedelta(days=12)},
        {"取引先": "B社", "金額": 1200000, "予定日": t + timedelta(days=18)},
        {"取引先": "C社", "金額": 500000,  "予定日": t + timedelta(days=25)},
        {"取引先": "D社", "金額": 1500000, "予定日": t + timedelta(days=35)},
        {"取引先": "E社", "金額": 900000,  "予定日": t + timedelta(days=50)},
        {"取引先": "F社", "金額": 1100000, "予定日": t + timedelta(days=65)},
    ])
    payables = pd.DataFrame([
        {"科目": "外注費A",      "金額": 600000,  "予定日": t + timedelta(days=8)},
        {"科目": "機材購入",     "金額": 1500000, "予定日": t + timedelta(days=20)},
        {"科目": "法人税中間",   "金額": 800000,  "予定日": t + timedelta(days=40)},
        {"科目": "外注費B",      "金額": 400000,  "予定日": t + timedelta(days=55)},
    ])
    recurring = pd.DataFrame([
        {"科目": "家賃",         "金額": 380000,  "支払日": 25},
        {"科目": "人件費",       "金額": 4500000, "支払日": 25},
        {"科目": "通信費",       "金額": 50000,   "支払日": 27},
        {"科目": "水道光熱費",   "金額": 80000,   "支払日": 28},
    ])
    return 3500000, receivables, payables, recurring


# ==============================
# 資金繰り：日次残高シミュレーター
# ==============================
def build_daily_balance(start_balance, receivables, payables, recurring,
                         days=90, delay_days=0, extra_amount=0, extra_day=0):
    """
    指定期間の日次残高を計算する。
    delay_days: 全入金予定をN日後ろにずらす（シナリオ分析）
    extra_amount, extra_day: 指定日（今日からN日後）に追加支出を発生させる
    """
    today = _today()
    rows = []
    balance = float(start_balance)

    for i in range(days + 1):
        d = today + timedelta(days=i)
        income = 0.0
        expense = 0.0
        income_detail = []
        expense_detail = []

        if not receivables.empty:
            for _, r in receivables.iterrows():
                scheduled = pd.to_datetime(r["予定日"]).date()
                actual = scheduled + timedelta(days=int(delay_days))
                if actual == d:
                    income += float(r["金額"])
                    income_detail.append(f"{r['取引先']} {int(r['金額']):,}円")

        if not payables.empty:
            for _, p in payables.iterrows():
                if pd.to_datetime(p["予定日"]).date() == d:
                    expense += float(p["金額"])
                    expense_detail.append(f"{p['科目']} {int(p['金額']):,}円")

        if not recurring.empty:
            for _, rec in recurring.iterrows():
                if d.day == int(rec["支払日"]) and i > 0:  # 今日の固定支出は出ない
                    expense += float(rec["金額"])
                    expense_detail.append(f"{rec['科目']} {int(rec['金額']):,}円")

        if extra_amount > 0 and i == int(extra_day):
            expense += float(extra_amount)
            expense_detail.append(f"想定外支出 {int(extra_amount):,}円")

        balance += income - expense
        rows.append({
            "日付": d, "残高": balance,
            "入金": income, "支出": expense,
            "入金内訳": " / ".join(income_detail) if income_detail else "",
            "支出内訳": " / ".join(expense_detail) if expense_detail else "",
        })
    return pd.DataFrame(rows)


def detect_shortage(df_balance):
    """残高がマイナスに沈む最初の日を返す"""
    neg = df_balance[df_balance["残高"] < 0]
    if neg.empty:
        return None
    first = neg.iloc[0]
    return {"日付": first["日付"], "残高": first["残高"], "支出内訳": first["支出内訳"]}


def calc_risk_rank(df_balance, monthly_baseline):
    """ランクA〜E + 評価コメント"""
    shortage = detect_shortage(df_balance)
    today = _today()

    if shortage:
        days_to = (shortage["日付"] - today).days
        if days_to <= 30:
            return "E", "🚨 30日以内に資金ショートの恐れ。即時対応が必要です。"
        elif days_to <= 60:
            return "D", "⚠ 60日以内にショート見込み。早期の手当を推奨します。"
        else:
            return "D", "⚠ 90日以内にショート見込み。対応策の検討を始めてください。"

    min_balance = float(df_balance["残高"].min())
    if monthly_baseline > 0:
        ratio = min_balance / monthly_baseline
        if ratio >= 1.0:
            return "A", "✓ 資金繰り良好。月商1ヶ月分以上の余裕があります。"
        elif ratio >= 0.5:
            return "B", "✓ 概ね健全。月商0.5〜1ヶ月分の余裕があります。"
        else:
            return "C", "△ 余裕が薄い。月商0.5ヶ月分未満の最低残高で要注意。"
    return "B", "✓ ショートなし。"


def rule_based_advice(df_balance, receivables, payables, recurring, balance):
    """ルールベースのCFO的助言（最大3件）"""
    advice = []
    shortage = detect_shortage(df_balance)
    today = _today()

    # 1. ショート対応
    if shortage:
        d = shortage["日付"]
        days_to = (d - today).days
        # その日までの最大支出を特定
        target_day = df_balance[df_balance["日付"] == d].iloc[0]
        if target_day["支出内訳"]:
            advice.append(
                f"**【最優先】{d.strftime('%m月%d日')}（あと{days_to}日）にショート見込み**。"
                f"主要因: {target_day['支出内訳']}。"
                f"この支払いの分割交渉、または同日までの早期入金確保（請求前倒し・前金受領）を検討してください。"
            )
        # 必要資金額
        need = abs(int(shortage["残高"])) + 500000  # バッファ50万
        advice.append(
            f"**【手当の選択肢】** ショート回避には約 **{need:,}円** の追加資金が必要。"
            f"短期借入（当座借越・手形貸付）/ ファクタリング / 経営者借入 から選択を。"
        )

    # 2. 売掛金の集中リスク
    if not receivables.empty:
        total_recv = receivables["金額"].sum()
        max_recv = receivables["金額"].max()
        if total_recv > 0 and max_recv / total_recv >= 0.4:
            top = receivables.loc[receivables["金額"].idxmax()]
            advice.append(
                f"**【取引先集中リスク】** {top['取引先']}の入金（{int(top['金額']):,}円）が"
                f"全入金の{max_recv/total_recv*100:.0f}%を占めています。"
                f"この1社の遅延・倒産が致命傷になる構造。取引先分散の検討を。"
            )

    # 3. 固定費負担の警告
    if not recurring.empty and balance > 0:
        monthly_fixed = recurring["金額"].sum()
        if monthly_fixed > balance * 0.4:
            advice.append(
                f"**【固定費過重】** 月次固定費 **{int(monthly_fixed):,}円** が"
                f"現在残高の{monthly_fixed/balance*100:.0f}%。"
                f"売上が1ヶ月止まると {int(balance/monthly_fixed*30):,}日で資金が枯渇します。"
            )

    # 4. ポジティブな所見（健全な場合）
    if not advice:
        min_balance = float(df_balance["残高"].min())
        advice.append(
            f"**【現状評価】** 90日先までショートなし。最低残高 **{int(min_balance):,}円**。"
            f"この余裕を活かして、設備投資・人材投資・営業強化への戦略的支出を検討する好機です。"
        )

    return advice[:3]


def export_cashflow_csv(balance, receivables, payables, recurring):
    rows = [{"type": "balance", "name": "", "amount": int(balance), "date": "", "day": ""}]
    for _, r in receivables.iterrows():
        rows.append({"type": "receivable", "name": r["取引先"], "amount": int(r["金額"]),
                     "date": pd.to_datetime(r["予定日"]).strftime("%Y-%m-%d"), "day": ""})
    for _, p in payables.iterrows():
        rows.append({"type": "payable", "name": p["科目"], "amount": int(p["金額"]),
                     "date": pd.to_datetime(p["予定日"]).strftime("%Y-%m-%d"), "day": ""})
    for _, rec in recurring.iterrows():
        rows.append({"type": "recurring", "name": rec["科目"], "amount": int(rec["金額"]),
                     "date": "", "day": int(rec["支払日"])})
    return pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig")


def import_cashflow_csv(file):
    df = pd.read_csv(file)
    balance = 0.0
    recv, pay, rec = [], [], []
    for _, row in df.iterrows():
        t = str(row.get("type", "")).strip()
        amt = float(row.get("amount", 0) or 0)
        name = str(row.get("name", "") or "")
        if t == "balance":
            balance = amt
        elif t == "receivable":
            recv.append({"取引先": name, "金額": amt,
                         "予定日": pd.to_datetime(row["date"]).date()})
        elif t == "payable":
            pay.append({"科目": name, "金額": amt,
                        "予定日": pd.to_datetime(row["date"]).date()})
        elif t == "recurring":
            rec.append({"科目": name, "金額": amt, "支払日": int(row["day"])})
    return (balance,
            pd.DataFrame(recv) if recv else pd.DataFrame(columns=["取引先", "金額", "予定日"]),
            pd.DataFrame(pay) if pay else pd.DataFrame(columns=["科目", "金額", "予定日"]),
            pd.DataFrame(rec) if rec else pd.DataFrame(columns=["科目", "金額", "支払日"]))


# ==============================
# タブ
# ==============================
tab1, tab2, tab3 = st.tabs(["💰 資金繰り3ヶ月先見", "📊 KPI診断", "🔬 What-if"])

# ==============================
# TAB 2: KPI診断
# ==============================
with tab2:
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
# TAB 1: 資金繰り3ヶ月先見（KILLER FEATURE）
# ==============================
with tab1:
    st.markdown("#### 資金繰り3ヶ月先見")
    st.caption("売掛金の入金予定・支払予定・固定費から、先90日の現金残高をシミュレーション。"
               "ショートポイントの自動検出と打ち手の提示まで行います。")

    # 初期化
    if "cf_balance" not in st.session_state:
        b, r, p, rec = _make_sample_cashflow()
        st.session_state["cf_balance"] = b
        st.session_state["cf_recv"] = r
        st.session_state["cf_pay"] = p
        st.session_state["cf_rec"] = rec

    # === ヘッダーアクション ===
    act_c1, act_c2, act_c3, act_c4 = st.columns([1, 1, 1, 2])
    with act_c1:
        if st.button("サンプル読込", use_container_width=True, key="cf_sample"):
            b, r, p, rec = _make_sample_cashflow()
            st.session_state["cf_balance"] = b
            st.session_state["cf_recv"] = r
            st.session_state["cf_pay"] = p
            st.session_state["cf_rec"] = rec
            st.rerun()
    with act_c2:
        csv_bytes = export_cashflow_csv(
            st.session_state["cf_balance"], st.session_state["cf_recv"],
            st.session_state["cf_pay"], st.session_state["cf_rec"],
        )
        st.download_button(
            "CSV エクスポート", data=csv_bytes,
            file_name=f"cashflow_{_today().strftime('%Y%m%d')}.csv",
            mime="text/csv", use_container_width=True, key="cf_export",
        )
    with act_c3:
        upl = st.file_uploader("CSVインポート", type=["csv"], key="cf_import",
                                label_visibility="collapsed")
        if upl is not None:
            try:
                b, r, p, rec = import_cashflow_csv(upl)
                st.session_state["cf_balance"] = b
                st.session_state["cf_recv"] = r
                st.session_state["cf_pay"] = p
                st.session_state["cf_rec"] = rec
                st.success("CSVを取り込みました")
            except Exception as e:
                st.error(f"取り込みエラー: {e}")

    # === データ入力（折りたたみ） ===
    with st.expander("📝 データを入力・編集する", expanded=False):
        in_c1, in_c2 = st.columns([1, 3])
        with in_c1:
            new_bal = st.number_input(
                "現在の現金残高（円）",
                value=int(st.session_state["cf_balance"]),
                step=100000, format="%d", key="cf_balance_input",
            )
            st.session_state["cf_balance"] = new_bal

        st.markdown("**入金予定（売掛金）**")
        edited_recv = st.data_editor(
            st.session_state["cf_recv"], num_rows="dynamic", use_container_width=True,
            column_config={
                "金額": st.column_config.NumberColumn("金額（円）", format="%d", step=100000),
                "予定日": st.column_config.DateColumn("予定日"),
            },
            key="cf_recv_editor",
        )
        st.session_state["cf_recv"] = edited_recv

        st.markdown("**支払予定（一回限り）**")
        edited_pay = st.data_editor(
            st.session_state["cf_pay"], num_rows="dynamic", use_container_width=True,
            column_config={
                "金額": st.column_config.NumberColumn("金額（円）", format="%d", step=100000),
                "予定日": st.column_config.DateColumn("予定日"),
            },
            key="cf_pay_editor",
        )
        st.session_state["cf_pay"] = edited_pay

        st.markdown("**定期支出（毎月）**")
        edited_rec = st.data_editor(
            st.session_state["cf_rec"], num_rows="dynamic", use_container_width=True,
            column_config={
                "金額": st.column_config.NumberColumn("金額（円）", format="%d", step=10000),
                "支払日": st.column_config.NumberColumn("支払日（毎月）", min_value=1, max_value=31, step=1),
            },
            key="cf_rec_editor",
        )
        st.session_state["cf_rec"] = edited_rec

    # === 計算 ===
    df_balance = build_daily_balance(
        st.session_state["cf_balance"],
        st.session_state["cf_recv"], st.session_state["cf_pay"], st.session_state["cf_rec"],
        days=90,
    )
    shortage = detect_shortage(df_balance)

    # 月商ベースライン（3ヶ月入金合計÷3）
    monthly_baseline = float(st.session_state["cf_recv"]["金額"].sum()) / 3 if not st.session_state["cf_recv"].empty else 0
    rank, rank_msg = calc_risk_rank(df_balance, monthly_baseline)

    # === ショートアラート（最上部に固定表示） ===
    if shortage:
        days_to = (shortage["日付"] - _today()).days
        st.error(
            f"## 🚨 ショート警告：{shortage['日付'].strftime('%Y年%m月%d日')}（あと{days_to}日）\n\n"
            f"残高見込み: **{int(shortage['残高']):,}円**　／　主要因: {shortage['支出内訳']}"
        )

    # === サマリーカード ===
    sm_c1, sm_c2, sm_c3, sm_c4 = st.columns(4)
    min_balance = float(df_balance["残高"].min())
    min_day = df_balance.loc[df_balance["残高"].idxmin(), "日付"]
    end_balance = float(df_balance["残高"].iloc[-1])

    rank_color = {"A": "#10B981", "B": "#5EAFFF", "C": "#F59E0B", "D": "#F97316", "E": "#EF4444"}[rank]

    sm_c1.metric("現在残高", fmt_yen(st.session_state["cf_balance"]))
    sm_c2.metric("最低残高（90日内）", fmt_yen(min_balance),
                 f"{min_day.strftime('%m/%d')}",
                 delta_color="inverse" if min_balance < 0 else "off")
    sm_c3.metric("90日後残高", fmt_yen(end_balance),
                 f"{(end_balance - st.session_state['cf_balance'])/10000:+,.0f}万円")
    sm_c4.markdown(
        f"<div style='background:{rank_color}; padding:14px; border-radius:6px; "
        f"text-align:center; margin-top:8px;'>"
        f"<div style='color:#0F1419; font-size:0.75rem; font-weight:600; letter-spacing:0.1em;'>RISK RANK</div>"
        f"<div style='color:#0F1419; font-size:2.5rem; font-weight:700; line-height:1;'>{rank}</div>"
        f"</div>",
        unsafe_allow_html=True,
    )

    st.caption(rank_msg)
    st.divider()

    # === 残高推移グラフ ===
    st.markdown("#### 日次残高シミュレーション（先90日）")

    fig_cf = go.Figure()
    fig_cf.add_trace(go.Scatter(
        x=df_balance["日付"], y=df_balance["残高"],
        mode="lines", name="残高",
        line={"color": "#5EAFFF", "width": 2.5},
        fill="tozeroy", fillcolor="rgba(94, 175, 255, 0.12)",
        hovertemplate="<b>%{x|%Y-%m-%d}</b><br>残高: %{y:,.0f}円<extra></extra>",
    ))
    # ゼロライン
    fig_cf.add_hline(y=0, line_dash="dash", line_color="#EF4444", line_width=1.5)

    # 大型支出マーカー
    big_expenses = df_balance[df_balance["支出"] >= 500000]
    if not big_expenses.empty:
        fig_cf.add_trace(go.Scatter(
            x=big_expenses["日付"], y=big_expenses["残高"],
            mode="markers", name="主な支出日",
            marker={"size": 10, "color": "#F59E0B", "symbol": "triangle-down"},
            customdata=big_expenses[["支出", "支出内訳"]].values,
            hovertemplate="<b>%{x|%Y-%m-%d}</b><br>支出: %{customdata[0]:,.0f}円<br>%{customdata[1]}<extra></extra>",
        ))

    # ショートポイント
    if shortage:
        fig_cf.add_vline(
            x=shortage["日付"], line_dash="dot", line_color="#EF4444", line_width=2,
            annotation_text=f"⚠ ショート",
            annotation_position="top",
            annotation_font={"color": "#EF4444"},
        )

    fig_cf.update_layout(
        height=420,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#1A1F2B",
        font={"color": "#E5E7EB"},
        xaxis={"gridcolor": "#2A2F3A", "title": ""},
        yaxis={"gridcolor": "#2A2F3A", "title": "残高（円）", "tickformat": ",.0f"},
        legend={"bgcolor": "rgba(0,0,0,0)", "orientation": "h", "y": 1.08},
        margin=dict(t=40, b=20, l=10, r=10),
        hovermode="x unified",
    )
    st.plotly_chart(fig_cf, use_container_width=True)

    st.divider()

    # === シナリオ分析 ===
    st.markdown("#### シナリオ分析")
    st.caption("「もし入金が遅れたら」「予期せぬ支出が出たら」を試算できます。")

    sc_c1, sc_c2 = st.columns(2)
    with sc_c1:
        st.markdown("**入金遅延シナリオ**")
        delay_days = st.slider(
            "全入金が ◯日遅れたら", 0, 60, 0, step=1, key="cf_delay",
            help="主要取引先の支払い遅延を想定したシミュレーション",
        )
    with sc_c2:
        st.markdown("**予期せぬ支出シナリオ**")
        ex_amt_man = st.slider(
            "想定外支出（万円）", 0, 1000, 0, step=10, key="cf_ex_amt",
        )
        ex_day = st.slider(
            "発生する日（今日からN日後）", 1, 90, 30, step=1, key="cf_ex_day",
            disabled=(ex_amt_man == 0),
        )

    df_scenario = build_daily_balance(
        st.session_state["cf_balance"],
        st.session_state["cf_recv"], st.session_state["cf_pay"], st.session_state["cf_rec"],
        days=90, delay_days=delay_days,
        extra_amount=ex_amt_man * 10000, extra_day=ex_day,
    )
    scenario_shortage = detect_shortage(df_scenario)

    # 比較グラフ
    fig_sc = go.Figure()
    fig_sc.add_trace(go.Scatter(
        x=df_balance["日付"], y=df_balance["残高"],
        mode="lines", name="現状",
        line={"color": "#6B7280", "width": 2, "dash": "dot"},
    ))
    fig_sc.add_trace(go.Scatter(
        x=df_scenario["日付"], y=df_scenario["残高"],
        mode="lines", name="シナリオ後",
        line={"color": "#F59E0B", "width": 2.5},
        fill="tozeroy", fillcolor="rgba(245, 158, 11, 0.12)",
    ))
    fig_sc.add_hline(y=0, line_dash="dash", line_color="#EF4444")
    fig_sc.update_layout(
        height=350,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#1A1F2B",
        font={"color": "#E5E7EB"},
        xaxis={"gridcolor": "#2A2F3A"},
        yaxis={"gridcolor": "#2A2F3A", "title": "残高（円）", "tickformat": ",.0f"},
        legend={"bgcolor": "rgba(0,0,0,0)", "orientation": "h", "y": 1.1},
        margin=dict(t=40, b=20, l=10, r=10),
    )
    st.plotly_chart(fig_sc, use_container_width=True)

    sc_msg_c1, sc_msg_c2 = st.columns(2)
    sc_msg_c1.metric(
        "シナリオ後 最低残高",
        fmt_yen(float(df_scenario["残高"].min())),
        f"{(float(df_scenario['残高'].min()) - min_balance)/10000:+,.0f}万円 vs 現状",
        delta_color="inverse",
    )
    if scenario_shortage and not shortage:
        sc_msg_c2.error(f"⚠ このシナリオでは {scenario_shortage['日付'].strftime('%m/%d')} にショート")
    elif scenario_shortage and shortage:
        diff_days = (scenario_shortage["日付"] - shortage["日付"]).days
        sc_msg_c2.warning(f"ショート日が {abs(diff_days)}日 {'後ろ倒し' if diff_days > 0 else '前倒し'}")
    else:
        sc_msg_c2.success("✓ このシナリオでもショートなし")

    st.divider()

    # === ルールベース CFOアドバイス ===
    st.markdown("#### CFOアドバイス（ルールベース）")
    st.caption("財務状況から自動生成された経営アドバイス。")

    advice_list = rule_based_advice(
        df_balance, st.session_state["cf_recv"], st.session_state["cf_pay"],
        st.session_state["cf_rec"], st.session_state["cf_balance"],
    )
    for i, ad in enumerate(advice_list, 1):
        st.markdown(f"**{i}.** {ad}")


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
