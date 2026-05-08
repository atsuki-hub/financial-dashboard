# -*- coding: utf-8 -*-
import os
import sys
import json
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from io import StringIO
from dotenv import load_dotenv
import anthropic
import fitz  # PyMuPDF

load_dotenv()

try:
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# Streamlit CloudのSecretsからAPIキーを取得（ローカルは.envを使用）
try:
    if "ANTHROPIC_API_KEY" in st.secrets:
        os.environ["ANTHROPIC_API_KEY"] = st.secrets["ANTHROPIC_API_KEY"]
except Exception:
    pass

st.set_page_config(page_title="財務KPI分析ダッシュボード", page_icon="📊", layout="wide")
st.title("📊 財務KPI分析ダッシュボード")
st.caption("個人事業主・フリーランス向け 財務分析ツール")

with st.expander("📖 はじめての方へ — 使い方ガイド", expanded=False):
    st.markdown("""
#### このアプリでできること
試算表CSV・PDFをアップロードするだけで、財務KPIの分析・資金繰り予測・税金試算まで自動化できます。

#### おすすめの使い方フロー

| ステップ | タブ | 内容 |
|---------|------|------|
| **① まずここから** | 📈 KPI分析 | CSVまたはPDFをアップロードして財務状況を把握 |
| **② 予算と比較** | 🎯 予実分析 | 予算CSVと実績CSVで達成率・差異を確認 |
| **③ 月次トレンド** | 📅 月次推移 | 複数月のデータを登録してトレンドを把握 |
| **④ 年度末予測** | 🔮 着地見込み | 月次推移から年度末の着地を予測 |
| **⑤ 資金管理** | 💰 資金繰り | 売掛金・支払予定から3ヶ月先の資金繰りを予測 |
| **⑥ 黒字化分析** | 📉 損益分岐点 | あといくら売れば黒字になるかを計算 |
| **⑦ 節税試算** | 💼 フリーランス試算 | 年収から手取り・税金・節税効果を計算 |
| **⑧ 成長確認** | 📊 前期比較 | 前期と今期を並べて成長率を自動計算 |
| **⑨ 総合診断** | 🏆 健全度スコア | 財務の健全性を0〜100点でスコアリング |
| **⑩ 意思決定** | 🔬 What-if | 売上・コストを変化させたときの影響をシミュレーション |
| **⑪ 法人化検討** | 🏢 法人化判断 | 個人事業主 vs 法人の手取りを比較 |

#### CSVのフォーマット
```
勘定科目,金額
売上高,5000000
売上原価,3000000
販売費及び一般管理費,1000000
...
```
""")

# ---- サンプルデータ ----
SAMPLE_ACTUAL = """勘定科目,金額
売上高,50000000
売上原価,30000000
販売費及び一般管理費,10000000
営業外収益,500000
営業外費用,300000
特別利益,0
特別損失,0
流動資産,25000000
固定資産,15000000
流動負債,10000000
固定負債,8000000
純資産,22000000
売上債権,8000000
棚卸資産,5000000
仕入債務,4000000
"""

SAMPLE_BUDGET = """勘定科目,金額
売上高,55000000
売上原価,32000000
販売費及び一般管理費,9500000
営業外収益,400000
営業外費用,300000
特別利益,0
特別損失,0
流動資産,27000000
固定資産,15000000
流動負債,10000000
固定負債,8000000
純資産,24000000
売上債権,9000000
棚卸資産,5500000
仕入債務,4200000
"""

# 月ごとにばらつきを持たせたサンプル月次データ
SAMPLE_MONTHLY = {
    1:  {"売上高": 38000000, "売上原価": 22000000, "販売費及び一般管理費": 8000000, "営業外収益": 300000, "営業外費用": 200000},
    2:  {"売上高": 41000000, "売上原価": 24000000, "販売費及び一般管理費": 8200000, "営業外収益": 320000, "営業外費用": 210000},
    3:  {"売上高": 52000000, "売上原価": 31000000, "販売費及び一般管理費": 9500000, "営業外収益": 400000, "営業外費用": 250000},
    4:  {"売上高": 45000000, "売上原価": 27000000, "販売費及び一般管理費": 8800000, "営業外収益": 350000, "営業外費用": 230000},
    5:  {"売上高": 48000000, "売上原価": 28000000, "販売費及び一般管理費": 9000000, "営業外収益": 360000, "営業外費用": 240000},
    6:  {"売上高": 50000000, "売上原価": 30000000, "販売費及び一般管理費": 9200000, "営業外収益": 380000, "営業外費用": 250000},
}

ANNUAL_BUDGET = {
    "売上高": 600000000,
    "売上原価": 360000000,
    "販売費及び一般管理費": 108000000,
    "営業外収益": 4800000,
    "営業外費用": 3000000,
}

INCOME_ITEMS = ["売上高", "売上原価", "販売費及び一般管理費", "営業外収益", "営業外費用"]
REVENUE_ITEMS = {"売上高", "営業外収益"}

SAMPLE_AR = """取引先,金額,入金予定日
A社,300000,2026-05-20
B社,500000,2026-05-31
C社,200000,2026-06-15
D社,800000,2026-06-30
E社,400000,2026-07-20
F社,600000,2026-07-31
"""

SAMPLE_AP = """支払先,金額,支払予定日
家賃,150000,2026-05-25
水道光熱費,30000,2026-05-28
家賃,150000,2026-06-25
外注費,200000,2026-06-10
水道光熱費,30000,2026-06-28
家賃,150000,2026-07-25
ソフトウェア,50000,2026-07-15
水道光熱費,30000,2026-07-28
"""

# ---- ヘルパー関数 ----
def extract_financials_from_pdf(pdf_file) -> pd.Series:
    pdf_bytes = pdf_file.read()
    table_text = ""
    fallback_text = ""

    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page_num, page in enumerate(doc):
            # ① テーブル抽出を優先
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

            # ② テーブルが取れなかったページはテキストで補完
            try:
                page_text = page.get_text("text") or ""
                fallback_text += page_text + "\n"
            except Exception:
                pass

    # テーブルが取れていればテーブル優先、なければテキスト
    raw_input = table_text if table_text.strip() else fallback_text

    # 文字コードを安全なUTF-8に統一
    text = raw_input.encode("utf-8", errors="replace").decode("utf-8")

    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    prompt = (
        "以下は財務書類のPDFから抽出したテーブルデータです（「科目名 | 金額」形式）。\n"
        "損益計算書と貸借対照表の数値を読み取り、必ず下記JSONのみを返してください。\n"
        "最初の文字を { にしてください。説明文・前置き・コードブロックは不要です。\n\n"
        "【変換ルール】\n"
        "- 金額は円単位の整数（千円単位→×1000、百万円単位→×1000000）\n"
        "- 不明・該当なしは0\n"
        "- 売掛金=売上債権、買掛金=仕入債務、販管費=販売費及び一般管理費\n"
        "- 合計欄（〇〇合計）を使う\n"
        "- 売上原価の記載なし → 売上高 - 売上総利益 で計算\n"
        "- 販管費の記載なし → 売上総利益 - 営業利益 で計算\n"
        "- 複数期間ある場合は最新期を使う\n\n"
        '{"売上高":0,"売上原価":0,"販売費及び一般管理費":0,'
        '"営業外収益":0,"営業外費用":0,"特別利益":0,"特別損失":0,'
        '"流動資産":0,"固定資産":0,"流動負債":0,"固定負債":0,"純資産":0,'
        '"売上債権":0,"棚卸資産":0,"仕入債務":0}\n\n'
        "--- テーブルデータ ---\n"
        + text[:20000]
    )
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = message.content[0].text.strip()

    # コードブロックを除去
    if "```" in raw:
        for part in raw.split("```"):
            part = part.strip().lstrip("json").strip()
            if part.startswith("{"):
                raw = part
                break

    # JSONが見つからない場合は空のデフォルト値を返す
    if not raw.startswith("{"):
        start = raw.find("{")
        end = raw.rfind("}") + 1
        if start != -1 and end > start:
            raw = raw[start:end]
        else:
            raise ValueError(
                f"ClaudeがJSONを返しませんでした。\n返答内容: {message.content[0].text[:300]}"
            )

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
    売上高 = get(d, "売上高")
    売上原価 = get(d, "売上原価")
    販管費 = get(d, "販売費及び一般管理費")
    営業外収益 = get(d, "営業外収益")
    営業外費用 = get(d, "営業外費用")
    特別利益 = get(d, "特別利益")
    特別損失 = get(d, "特別損失")
    流動資産 = get(d, "流動資産")
    固定資産 = get(d, "固定資産")
    流動負債 = get(d, "流動負債")
    固定負債 = get(d, "固定負債")
    純資産 = get(d, "純資産")
    売上債権 = get(d, "売上債権")
    棚卸資産 = get(d, "棚卸資産")
    仕入債務 = get(d, "仕入債務")

    売上総利益 = 売上高 - 売上原価
    営業利益 = 売上総利益 - 販管費
    経常利益 = 営業利益 + 営業外収益 - 営業外費用
    税引前純利益 = 経常利益 + 特別利益 - 特別損失
    総資産 = 流動資産 + 固定資産
    総負債 = 流動負債 + 固定負債

    def pct(a, b):
        return round(a / b * 100, 1) if b else None

    days = lambda a, b: round(a / b * 365, 1) if b else None

    return {
        "売上総利益": 売上総利益,
        "営業利益": 営業利益,
        "経常利益": 経常利益,
        "税引前純利益": 税引前純利益,
        "売上総利益率": pct(売上総利益, 売上高),
        "営業利益率": pct(営業利益, 売上高),
        "経常利益率": pct(経常利益, 売上高),
        "流動比率": pct(流動資産, 流動負債),
        "自己資本比率": pct(純資産, 総資産),
        "売上債権回転日数": days(売上債権, 売上高),
        "棚卸資産回転日数": days(棚卸資産, 売上高),
        "仕入債務回転日数": days(仕入債務, 売上原価),
        "総資産": 総資産,
        "総負債": 総負債,
        "純資産": 純資産,
        "売上高": 売上高,
    }


def fmt_yen(v):
    if abs(v) >= 1_0000_0000:
        return f"{v/1_0000_0000:.1f}億円"
    elif abs(v) >= 10000:
        return f"{v/10000:.0f}万円"
    return f"{v:,.0f}円"


INDUSTRY_BENCHMARKS = {
    "IT・Web・SaaS": {
        "売上総利益率": "60〜80%", "営業利益率": "10〜20%", "流動比率": "150%以上", "自己資本比率": "40%以上",
        "特徴": "在庫を持たないため売上総利益率が高い。人件費が主なコストで販管費率が高くなりやすい。",
    },
    "デザイン・クリエイティブ": {
        "売上総利益率": "50〜70%", "営業利益率": "10〜20%", "流動比率": "150%以上", "自己資本比率": "40%以上",
        "特徴": "外注費が変動し売上総利益率がブレやすい。売掛金の回収管理が重要。",
    },
    "コンサルティング": {
        "売上総利益率": "60〜80%", "営業利益率": "15〜25%", "流動比率": "150%以上", "自己資本比率": "50%以上",
        "特徴": "原価が人件費中心のため高利益率が可能。稼働率と単価管理がKPIの核心。",
    },
    "飲食業": {
        "売上総利益率": "60〜70%", "営業利益率": "5〜10%", "流動比率": "100%以上", "自己資本比率": "20%以上",
        "特徴": "食材原価率30〜35%、人件費率30〜35%が目安。回転率と客単価が重要。",
    },
    "小売業": {
        "売上総利益率": "25〜40%", "営業利益率": "3〜8%", "流動比率": "120%以上", "自己資本比率": "30%以上",
        "特徴": "粗利率が低いため在庫回転と販管費管理が重要。棚卸資産回転日数に注目。",
    },
    "製造業": {
        "売上総利益率": "20〜35%", "営業利益率": "5〜10%", "流動比率": "150%以上", "自己資本比率": "35%以上",
        "特徴": "原材料・労務費・製造経費が原価を構成。設備投資と運転資本のバランスが課題。",
    },
    "建設・工事": {
        "売上総利益率": "15〜25%", "営業利益率": "3〜8%", "流動比率": "130%以上", "自己資本比率": "25%以上",
        "特徴": "受注産業のため完成工事高の変動が大きい。工事ごとの採算管理が重要。",
    },
    "医療・介護": {
        "売上総利益率": "40〜60%", "営業利益率": "5〜10%", "流動比率": "150%以上", "自己資本比率": "30%以上",
        "特徴": "診療報酬・介護報酬の制度変更に影響を受ける。人件費率が高い。",
    },
    "教育・スクール": {
        "売上総利益率": "50〜70%", "営業利益率": "10〜15%", "流動比率": "150%以上", "自己資本比率": "40%以上",
        "特徴": "前払い収入（前受金）が多く資金繰りが安定しやすい。生徒数が重要KPI。",
    },
    "不動産": {
        "売上総利益率": "30〜50%", "営業利益率": "10〜20%", "流動比率": "100%以上", "自己資本比率": "20%以上",
        "特徴": "借入依存度が高く自己資本比率が低めになりやすい。家賃収入の安定性が重要。",
    },
    "その他・未選択": {
        "売上総利益率": "業種により異なる", "営業利益率": "業種により異なる", "流動比率": "100%以上", "自己資本比率": "30%以上",
        "特徴": "業種ベンチマークなし。一般的な中小企業水準で評価します。",
    },
}


def generate_kpi_comment(kpis: dict, industry: str, memo: str) -> str:
    bm = INDUSTRY_BENCHMARKS.get(industry, INDUSTRY_BENCHMARKS["その他・未選択"])
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    kpi_text = "\n".join([
        f"売上高: {fmt_yen(kpis['売上高'])}",
        f"売上総利益率: {kpis['売上総利益率']}%（業種目安: {bm['売上総利益率']}）",
        f"営業利益率: {kpis['営業利益率']}%（業種目安: {bm['営業利益率']}）",
        f"経常利益率: {kpis['経常利益率']}%",
        f"流動比率: {kpis['流動比率']}%（業種目安: {bm['流動比率']}）",
        f"自己資本比率: {kpis['自己資本比率']}%（業種目安: {bm['自己資本比率']}）",
        f"売上債権回転日数: {kpis['売上債権回転日数']}日",
        f"棚卸資産回転日数: {kpis['棚卸資産回転日数']}日",
        f"仕入債務回転日数: {kpis['仕入債務回転日数']}日",
    ])
    prompt = (
        f"あなたはCFOとして個人事業主・中小企業の財務分析を行うアドバイザーです。\n"
        f"対象事業者の業種は「{industry}」です。\n"
        f"業種の特徴: {bm['特徴']}\n\n"
        "以下のKPIと業種ベンチマークをもとに、経営者向けの財務コメントを日本語で作成してください。\n\n"
        "【出力形式】\n"
        "## 収益性\n（2〜3文。業種ベンチマークと比較して利益率を評価し、改善点を指摘）\n\n"
        "## 安全性\n（2〜3文。流動比率・自己資本比率から資金繰りと財務健全性を評価）\n\n"
        "## 効率性\n（2〜3文。回転日数から運転資本の効率と資金回収の課題を指摘）\n\n"
        "## 総合評価とアドバイス\n（3〜4文。強み・弱みをまとめ、この業種に即した具体的なアクションを提案）\n\n"
        + (f"【経営者からの補足情報】\n{memo}\n\n" if memo.strip() else "")
        + "【注意】専門用語は避け、経営者が直感的に理解できる言葉で書いてください。\n\n"
        "【KPIデータ】\n" + kpi_text
    )
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1200,
        messages=[{"role": "user", "content": prompt}],
    )
    return msg.content[0].text


def generate_yoji_comment(df_result: pd.DataFrame, industry: str, memo: str) -> str:
    bm = INDUSTRY_BENCHMARKS.get(industry, INDUSTRY_BENCHMARKS["その他・未選択"])
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    rows_text = "\n".join([
        f"{row['項目']}: 予算{fmt_yen(row['予算'])} / 実績{fmt_yen(row['実績'])} / 差異{fmt_yen(row['差異'])} / 達成率{row['達成率(%)']}%"
        for _, row in df_result.iterrows()
    ])
    prompt = (
        f"あなたはCFOとして個人事業主・中小企業の予実管理を支援するアドバイザーです。\n"
        f"対象事業者の業種は「{industry}」です。業種の特徴: {bm['特徴']}\n\n"
        "以下の予実データをもとに、経営者向けの月次コメントを日本語で作成してください。\n\n"
        "【出力形式】\n"
        "## 予実サマリー\n（2〜3文。全体の達成状況を端的に）\n\n"
        "## 未達項目の要因分析\n（2〜3文。この業種の特性を踏まえて要因を指摘）\n\n"
        "## 次月へのアクション\n（箇条書き3点。この業種に即した具体的な改善アクション）\n\n"
        + (f"【経営者からの補足情報】\n{memo}\n\n" if memo.strip() else "")
        + "【注意】数字を引用しながら、経営者が次の行動を取りやすい文章にしてください。\n\n"
        "【予実データ】\n" + rows_text
    )
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1200,
        messages=[{"role": "user", "content": prompt}],
    )
    return msg.content[0].text


def calc_health_score(kpis):
    def score(val, thresholds):
        if val is None: return 0
        for limit, pts in thresholds:
            if val <= limit: return pts
        return thresholds[-1][1]

    gp  = score(kpis.get("売上総利益率") or 0, [(0,0),(10,15),(20,35),(30,55),(40,75),(100,100)])
    op  = score(kpis.get("営業利益率") or 0,  [(0,0),(3,15),(7,40),(12,65),(20,85),(100,100)])
    cr  = score(kpis.get("流動比率") or 0,    [(50,5),(80,20),(100,45),(130,65),(150,80),(200,100)])
    er  = score(kpis.get("自己資本比率") or 0,[(0,0),(20,25),(30,50),(40,70),(50,85),(100,100)])
    ar  = score(-(kpis.get("売上債権回転日数") or 60), [(-90,10),(-60,40),(-45,60),(-30,80),(-15,95),(0,100)])

    収益性 = int(gp * 0.5 + op * 0.5)
    安全性 = int(cr * 0.5 + er * 0.5)
    効率性 = ar
    総合   = int(収益性 * 0.40 + 安全性 * 0.40 + 効率性 * 0.20)
    return {"総合": 総合, "収益性": 収益性, "安全性": 安全性, "効率性": 効率性}


def calc_income_tax_jp(taxable):
    if taxable <= 0: return 0
    brackets = [(1950000,.05,0),(3300000,.10,97500),(6950000,.20,427500),
                (9000000,.23,636000),(18000000,.33,1536000),(40000000,.40,2796000),(1e9,.45,4796000)]
    for limit, rate, ded in brackets:
        if taxable <= limit:
            return max(0, int((taxable * rate - ded) * 1.021))
    return 0


def calc_incorporation(annual_revenue, annual_expense, owner_salary):
    # 個人事業主
    aokiro = 650000
    jigyou = max(0, annual_revenue - annual_expense - aokiro)
    nenkin_ind = 203760
    kokuho_ind = max(0, int((jigyou - 430000) * 0.10))
    sha_ind = nenkin_ind + kokuho_ind
    taxable_ind = max(0, jigyou - sha_ind - 480000)
    tax_ind = calc_income_tax_jp(taxable_ind) + max(0, int(taxable_ind * 0.10))
    takehome_ind = annual_revenue - annual_expense - tax_ind - sha_ind

    # 法人
    corp_profit = max(0, annual_revenue - annual_expense - owner_salary)
    corp_tax_rate = 0.234
    corp_tax = int(corp_profit * corp_tax_rate)
    sha_corp = int(owner_salary * 0.155)  # 健保+厚生年金（本人負担）
    kyuyo_kojo = min(owner_salary, max(550000, int(owner_salary * 0.40)))
    taxable_corp_owner = max(0, owner_salary - kyuyo_kojo - sha_corp - 480000)
    tax_corp_owner = calc_income_tax_jp(taxable_corp_owner) + max(0, int(taxable_corp_owner * 0.10))
    takehome_corp = owner_salary - tax_corp_owner - sha_corp
    retained = corp_profit - corp_tax

    return {
        "個人_手取り": takehome_ind, "個人_税社保": tax_ind + sha_ind,
        "法人_手取り": takehome_corp, "法人_税社保": tax_corp_owner + sha_corp + corp_tax,
        "法人_内部留保": retained, "法人_法人税": corp_tax,
        "差額": takehome_corp - takehome_ind,
    }


def calc_profit(monthly_dict):
    """月次dictから利益項目を計算して返す"""
    s = 売上高 = monthly_dict.get("売上高", 0)
    c = 売上原価 = monthly_dict.get("売上原価", 0)
    sg = 販管費 = monthly_dict.get("販売費及び一般管理費", 0)
    oi = monthly_dict.get("営業外収益", 0)
    oe = monthly_dict.get("営業外費用", 0)
    return {
        "売上高": s,
        "売上総利益": s - c,
        "営業利益": s - c - sg,
        "経常利益": s - c - sg + oi - oe,
    }


# ==============================
# タブ
# ==============================
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10, tab11 = st.tabs([
    "📈 KPI分析", "🎯 予実分析", "📅 月次推移", "🔮 着地見込み", "💰 資金繰り",
    "📉 損益分岐点", "💼 フリーランス試算", "📊 前期比較",
    "🏆 健全度スコア", "🔬 What-if", "🏢 法人化判断"
])


# ==============================
# TAB 1: KPI分析
# ==============================
with tab1:
    with st.sidebar:
        st.header("データ入力")
        input_method = st.radio("入力方法", ["CSV", "PDF（AI読み取り）"], horizontal=True)

        uploaded_file = None
        if input_method == "CSV":
            st.markdown("**CSVフォーマット**")
            st.code("勘定科目,金額\n売上高,50000000\n...")
            if st.button("サンプルデータで試す", key="sample_kpi"):
                st.session_state["use_sample"] = True
                st.session_state.pop("pdf_extracted", None)
            uploaded_file = st.file_uploader("実績CSVをアップロード", type=["csv"], key="kpi_upload")
        else:
            st.markdown("**対応PDFの例**")
            st.markdown("- freee / MFクラウドの試算表\n- 決算短信・有価証券報告書")
            pdf_file = st.file_uploader("PDFをアップロード", type=["pdf"], key="pdf_upload")
            if pdf_file and st.button("AIで読み取る", key="pdf_read"):
                with st.spinner("AIが財務データを読み取り中..."):
                    try:
                        st.session_state["pdf_extracted"] = extract_financials_from_pdf(pdf_file)
                        st.session_state.pop("use_sample", None)
                        st.success("読み取り完了！")
                    except Exception as e:
                        import traceback
                        tb = traceback.format_exc()
                        st.error("読み取りエラーが発生しました。")
                        st.code(tb, language="text")
            if st.session_state.get("pdf_extracted") is not None:
                with st.expander("抽出されたデータを確認・補完する"):
                    st.caption("0になっている項目は手入力で補完できます")
                    edited = {}
                    items_order = [
                        "売上高","売上原価","販売費及び一般管理費",
                        "営業外収益","営業外費用","特別利益","特別損失",
                        "流動資産","固定資産","流動負債","固定負債","純資産",
                        "売上債権","棚卸資産","仕入債務",
                    ]
                    for key in items_order:
                        val = int(st.session_state["pdf_extracted"].get(key, 0))
                        edited[key] = st.number_input(
                            key, value=val, step=10000, key=f"edit_{key}"
                        )
                    if st.button("この値で分析する", key="apply_edit"):
                        st.session_state["pdf_extracted"] = pd.Series(
                            {k: float(v) for k, v in edited.items()}
                        )
                        st.success("反映しました")

        st.divider()
        st.markdown("**必須勘定科目**")
        st.markdown("- 売上高 / 売上原価\n- 販売費及び一般管理費\n- 流動資産 / 流動負債\n- 固定資産 / 固定負債 / 純資産\n- 売上債権 / 棚卸資産 / 仕入債務")

    data = None
    if input_method == "PDF（AI読み取り）" and st.session_state.get("pdf_extracted") is not None:
        data = st.session_state["pdf_extracted"]
    elif uploaded_file:
        data = load_data(uploaded_file)
    elif st.session_state.get("use_sample"):
        data = load_data(StringIO(SAMPLE_ACTUAL))

    if data is None:
        st.info("サイドバーからCSVをアップロードするか、サンプルデータを試してください。")
    else:
        kpis = calc_kpis(data)

        st.subheader("損益サマリー")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("売上高", fmt_yen(kpis["売上高"]))
        c2.metric("売上総利益", fmt_yen(kpis["売上総利益"]))
        c3.metric("営業利益", fmt_yen(kpis["営業利益"]))
        c4.metric("経常利益", fmt_yen(kpis["経常利益"]))

        st.subheader("収益性指標")
        c1, c2, c3 = st.columns(3)
        for col, label in zip([c1, c2, c3], ["売上総利益率", "営業利益率", "経常利益率"]):
            v = kpis[label]
            col.metric(label, f"{v}%" if v is not None else "N/A")

        st.subheader("安全性指標")
        c1, c2 = st.columns(2)
        c1.metric("流動比率", f"{kpis['流動比率']}%" if kpis['流動比率'] else "N/A", help="100%以上が目安")
        c2.metric("自己資本比率", f"{kpis['自己資本比率']}%" if kpis['自己資本比率'] else "N/A", help="30%以上が目安")

        st.subheader("効率性指標（回転日数）")
        c1, c2, c3 = st.columns(3)
        c1.metric("売上債権回転日数", f"{kpis['売上債権回転日数']}日" if kpis['売上債権回転日数'] else "N/A")
        c2.metric("棚卸資産回転日数", f"{kpis['棚卸資産回転日数']}日" if kpis['棚卸資産回転日数'] else "N/A")
        c3.metric("仕入債務回転日数", f"{kpis['仕入債務回転日数']}日" if kpis['仕入債務回転日数'] else "N/A")

        st.divider()
        col_left, col_right = st.columns(2)
        with col_left:
            st.subheader("利益ウォーターフォール")
            fig = go.Figure(go.Waterfall(
                orientation="v",
                measure=["absolute", "relative", "total", "relative", "total"],
                x=["売上高", "売上原価", "売上総利益", "販管費", "営業利益"],
                y=[kpis["売上高"], -get(data, "売上原価"), kpis["売上総利益"],
                   -get(data, "販売費及び一般管理費"), kpis["営業利益"]],
                connector={"line": {"color": "rgb(63,63,63)"}},
                increasing={"marker": {"color": "#2196F3"}},
                decreasing={"marker": {"color": "#EF5350"}},
                totals={"marker": {"color": "#4CAF50"}},
            ))
            fig.update_layout(showlegend=False, height=350)
            st.plotly_chart(fig, use_container_width=True)

        with col_right:
            st.subheader("収益性KPI比較")
            kpi_labels = ["売上総利益率", "営業利益率", "経常利益率"]
            kpi_values = [kpis[k] for k in kpi_labels]
            fig2 = px.bar(x=kpi_labels, y=kpi_values, labels={"x": "", "y": "%"},
                          color=kpi_values, color_continuous_scale=["#EF5350", "#FFC107", "#4CAF50"],
                          range_color=[0, 30])
            fig2.update_layout(showlegend=False, height=350, coloraxis_showscale=False)
            fig2.update_traces(text=[f"{v}%" for v in kpi_values], textposition="outside")
            st.plotly_chart(fig2, use_container_width=True)

        st.subheader("貸借対照表サマリー")
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(name="流動資産", x=["資産"], y=[get(data, "流動資産")], marker_color="#42A5F5"))
        fig3.add_trace(go.Bar(name="固定資産", x=["資産"], y=[get(data, "固定資産")], marker_color="#1565C0"))
        fig3.add_trace(go.Bar(name="流動負債", x=["負債・純資産"], y=[get(data, "流動負債")], marker_color="#EF9A9A"))
        fig3.add_trace(go.Bar(name="固定負債", x=["負債・純資産"], y=[get(data, "固定負債")], marker_color="#C62828"))
        fig3.add_trace(go.Bar(name="純資産", x=["負債・純資産"], y=[get(data, "純資産")], marker_color="#66BB6A"))
        fig3.update_layout(barmode="stack", height=350, yaxis_title="金額（円）")
        st.plotly_chart(fig3, use_container_width=True)

        # ---- AI コメント ----
        st.divider()
        st.subheader("🤖 AIコメント")
        industry = st.selectbox(
            "業種を選択", list(INDUSTRY_BENCHMARKS.keys()), key="industry_kpi"
        )
        memo = st.text_area(
            "補足情報（任意）", placeholder="例: 今月は新規顧客獲得に注力した／季節要因で売上が下がった",
            key="memo_kpi", height=80
        )
        if st.button("AIに財務コメントを生成させる", key="gen_kpi_comment"):
            with st.spinner("AIが分析中..."):
                try:
                    comment = generate_kpi_comment(kpis, industry, memo)
                    st.session_state["kpi_comment"] = comment
                except Exception as e:
                    st.error(f"コメント生成エラー: {e}")
        if st.session_state.get("kpi_comment"):
            st.markdown(st.session_state["kpi_comment"])


# ==============================
# TAB 2: 予実分析
# ==============================
with tab2:
    st.subheader("🎯 予実分析")
    st.caption("予算と実績のCSVをアップロードして、差異・達成率を自動分析します")

    col_a, col_b = st.columns(2)
    with col_a:
        budget_file = st.file_uploader("予算CSVをアップロード", type=["csv"], key="budget_upload")
    with col_b:
        actual_file = st.file_uploader("実績CSVをアップロード", type=["csv"], key="actual_upload")

    if st.button("サンプルデータで予実分析を試す", key="sample_yoji"):
        st.session_state["use_sample_yoji"] = True

    budget_data = None
    actual_data = None
    if budget_file and actual_file:
        budget_data = load_data(budget_file)
        actual_data = load_data(actual_file)
    elif st.session_state.get("use_sample_yoji"):
        budget_data = load_data(StringIO(SAMPLE_BUDGET))
        actual_data = load_data(StringIO(SAMPLE_ACTUAL))

    if budget_data is None or actual_data is None:
        st.info("予算・実績の両方のCSVをアップロードするか、サンプルデータをお試しください。")
    else:
        budget_kpis = calc_kpis(budget_data)
        actual_kpis = calc_kpis(actual_data)

        profit_items = {"売上高": True, "売上総利益": True, "営業利益": True, "経常利益": True}
        rows = []
        for label, higher_is_better in profit_items.items():
            b = budget_kpis[label]
            a = actual_kpis[label]
            diff = a - b
            rate = round(a / b * 100, 1) if b else None
            rows.append({"項目": label, "予算": b, "実績": a, "差異": diff, "達成率(%)": rate,
                         "higher_is_better": higher_is_better})
        df_result = pd.DataFrame(rows)

        st.subheader("利益サマリー")
        cols = st.columns(4)
        for i, row in df_result.iterrows():
            delta_color = "normal" if (row["差異"] >= 0) == row["higher_is_better"] else "inverse"
            cols[i].metric(row["項目"], fmt_yen(row["実績"]),
                           delta=f"{row['差異']/10000:.0f}万円", delta_color=delta_color)

        st.divider()
        st.subheader("予算 vs 実績 比較")
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(name="予算", x=df_result["項目"], y=df_result["予算"],
                                 marker_color="#90CAF9", text=[fmt_yen(v) for v in df_result["予算"]],
                                 textposition="outside"))
        fig_bar.add_trace(go.Bar(name="実績", x=df_result["項目"], y=df_result["実績"],
                                 marker_color="#1565C0", text=[fmt_yen(v) for v in df_result["実績"]],
                                 textposition="outside"))
        fig_bar.update_layout(barmode="group", height=400, yaxis_title="金額（円）")
        st.plotly_chart(fig_bar, use_container_width=True)

        st.subheader("達成率")
        gauge_cols = st.columns(4)
        for i, row in df_result.iterrows():
            rate = row["達成率(%)"] or 0
            color = "#4CAF50" if rate >= 100 else "#FFC107" if rate >= 90 else "#EF5350"
            fig_g = go.Figure(go.Indicator(
                mode="gauge+number", value=rate, number={"suffix": "%"},
                title={"text": row["項目"]},
                gauge={"axis": {"range": [0, 130]}, "bar": {"color": color},
                       "steps": [{"range": [0, 90], "color": "#FFEBEE"},
                                 {"range": [90, 100], "color": "#FFF9C4"},
                                 {"range": [100, 130], "color": "#E8F5E9"}],
                       "threshold": {"line": {"color": "black", "width": 2}, "value": 100}},
            ))
            fig_g.update_layout(height=250, margin=dict(t=40, b=10, l=10, r=10))
            gauge_cols[i].plotly_chart(fig_g, use_container_width=True)

        st.subheader("勘定科目別 差異明細")
        all_items = sorted(set(budget_data.index) | set(actual_data.index))
        detail_rows = [{"勘定科目": item, "予算": budget_data.get(item, 0),
                        "実績": actual_data.get(item, 0),
                        "差異": actual_data.get(item, 0) - budget_data.get(item, 0),
                        "達成率(%)": round(actual_data.get(item, 0) / budget_data.get(item, 0) * 100, 1)
                        if budget_data.get(item, 0) else None}
                       for item in all_items]
        df_detail = pd.DataFrame(detail_rows)

        def highlight_diff(val):
            if isinstance(val, (int, float)):
                return "color: #2E7D32" if val > 0 else "color: #C62828" if val < 0 else ""
            return ""

        styled = df_detail.style.format({
            "予算": fmt_yen, "実績": fmt_yen, "差異": fmt_yen,
            "達成率(%)": lambda v: f"{v}%" if v is not None else "N/A",
        }).applymap(highlight_diff, subset=["差異"])
        st.dataframe(styled, use_container_width=True, hide_index=True)

        # ---- AI コメント ----
        st.divider()
        st.subheader("🤖 AIコメント")
        industry_y = st.selectbox(
            "業種を選択", list(INDUSTRY_BENCHMARKS.keys()), key="industry_yoji"
        )
        memo_y = st.text_area(
            "補足情報（任意）", placeholder="例: 今月は大型案件が遅延した／広告費を前倒し投下した",
            key="memo_yoji", height=80
        )
        if st.button("AIに予実コメントを生成させる", key="gen_yoji_comment"):
            with st.spinner("AIが分析中..."):
                try:
                    comment = generate_yoji_comment(df_result, industry_y, memo_y)
                    st.session_state["yoji_comment"] = comment
                except Exception as e:
                    st.error(f"コメント生成エラー: {e}")
        if st.session_state.get("yoji_comment"):
            st.markdown(st.session_state["yoji_comment"])


# ==============================
# TAB 3: 月次推移
# ==============================
with tab3:
    st.subheader("📅 月次推移")
    st.caption("月ごとの試算表CSVを登録して、KPIのトレンドを確認します")

    if "monthly_data" not in st.session_state:
        st.session_state["monthly_data"] = {}

    col_inp, col_status = st.columns([2, 1])
    with col_inp:
        selected_month = st.selectbox("月を選択", options=list(range(1, 13)),
                                      format_func=lambda m: f"{m}月")
        monthly_file = st.file_uploader(f"{selected_month}月のCSVをアップロード",
                                        type=["csv"], key=f"monthly_{selected_month}")
        if monthly_file:
            st.session_state["monthly_data"][selected_month] = load_data(monthly_file)
            st.success(f"{selected_month}月のデータを登録しました")

        if st.button("サンプルデータ（1〜6月）を読み込む", key="sample_monthly"):
            for m, d in SAMPLE_MONTHLY.items():
                st.session_state["monthly_data"][m] = pd.Series(d)
            st.success("1〜6月のサンプルデータを読み込みました")

    with col_status:
        st.markdown("**登録済み月**")
        registered = sorted(st.session_state["monthly_data"].keys())
        if registered:
            for m in registered:
                st.markdown(f"✅ {m}月")
            if st.button("リセット", key="reset_monthly"):
                st.session_state["monthly_data"] = {}
                st.rerun()
        else:
            st.markdown("*まだありません*")

    monthly = st.session_state["monthly_data"]
    if len(monthly) < 2:
        st.info("2ヶ月以上のデータを登録するとグラフが表示されます。")
    else:
        months_sorted = sorted(monthly.keys())
        month_labels = [f"{m}月" for m in months_sorted]

        profits = {m: calc_profit(monthly[m]) for m in months_sorted}

        st.divider()

        # ---- 売上・利益 推移 ----
        st.subheader("売上・利益 月次推移")
        items = ["売上高", "売上総利益", "営業利益", "経常利益"]
        colors = ["#1565C0", "#42A5F5", "#4CAF50", "#FFC107"]
        fig_trend = go.Figure()
        for item, color in zip(items, colors):
            values = [profits[m][item] for m in months_sorted]
            fig_trend.add_trace(go.Scatter(
                x=month_labels, y=values, name=item, mode="lines+markers",
                line=dict(color=color, width=2),
                marker=dict(size=8),
                text=[fmt_yen(v) for v in values],
                hovertemplate="%{x}: %{text}<extra>" + item + "</extra>",
            ))
        fig_trend.update_layout(height=400, yaxis_title="金額（円）", hovermode="x unified")
        st.plotly_chart(fig_trend, use_container_width=True)

        # ---- 利益率 推移 ----
        st.subheader("利益率 月次推移")
        fig_rate = go.Figure()
        rate_items = [
            ("売上総利益率", "#42A5F5"),
            ("営業利益率", "#4CAF50"),
            ("経常利益率", "#FFC107"),
        ]
        for label, color in rate_items:
            rates = []
            for m in months_sorted:
                p = profits[m]
                s = monthly[m].get("売上高", 0)
                if label == "売上総利益率":
                    rates.append(round(p["売上総利益"] / s * 100, 1) if s else 0)
                elif label == "営業利益率":
                    rates.append(round(p["営業利益"] / s * 100, 1) if s else 0)
                else:
                    rates.append(round(p["経常利益"] / s * 100, 1) if s else 0)
            fig_rate.add_trace(go.Scatter(
                x=month_labels, y=rates, name=label, mode="lines+markers",
                line=dict(color=color, width=2), marker=dict(size=8),
                text=[f"{r}%" for r in rates],
                hovertemplate="%{x}: %{text}<extra>" + label + "</extra>",
            ))
        fig_rate.update_layout(height=350, yaxis_title="%", hovermode="x unified")
        st.plotly_chart(fig_rate, use_container_width=True)

        # ---- 移動平均 + 異常値検知 ----
        st.subheader("売上高 移動平均・異常値検知")
        sales_series = pd.Series([profits[m]["売上高"] for m in months_sorted], index=months_sorted)
        ma3 = sales_series.rolling(3, min_periods=2).mean()
        mean_val = sales_series.mean()
        std_val  = sales_series.std()
        anomaly  = (sales_series - mean_val).abs() > 1.5 * std_val if std_val > 0 else pd.Series([False]*len(sales_series), index=months_sorted)

        fig_ma = go.Figure()
        fig_ma.add_trace(go.Bar(
            x=month_labels, y=sales_series.values, name="売上高",
            marker_color=["#EF5350" if anomaly[m] else "#90CAF9" for m in months_sorted],
        ))
        fig_ma.add_trace(go.Scatter(
            x=month_labels, y=ma3.values, name="3ヶ月移動平均",
            mode="lines+markers", line=dict(color="#1565C0", width=2, dash="dot"),
        ))
        fig_ma.add_hline(y=mean_val, line_dash="dash", line_color="#FFC107",
                         annotation_text="平均", annotation_position="right")
        fig_ma.update_layout(height=380, yaxis_title="金額（円）")
        st.plotly_chart(fig_ma, use_container_width=True)
        if anomaly.any():
            months_anomaly = [f"{m}月" for m in months_sorted if anomaly[m]]
            st.warning(f"⚠️ 異常値を検知: {', '.join(months_anomaly)}（平均から大きく乖離）")

        # ---- 前月比テーブル ----
        st.subheader("前月比テーブル")
        table_rows = []
        for i, m in enumerate(months_sorted):
            row = {"月": f"{m}月"}
            for item in ["売上高", "売上総利益", "営業利益", "経常利益"]:
                v = profits[m][item]
                row[item] = fmt_yen(v)
                if i > 0:
                    prev = profits[months_sorted[i - 1]][item]
                    diff_pct = round((v - prev) / abs(prev) * 100, 1) if prev else 0
                    sign = "▲" if diff_pct >= 0 else "▼"
                    row[f"{item}_前月比"] = f"{sign}{abs(diff_pct)}%"
                else:
                    row[f"{item}_前月比"] = "-"
            table_rows.append(row)

        df_monthly = pd.DataFrame(table_rows)
        st.dataframe(df_monthly, use_container_width=True, hide_index=True)


# ==============================
# TAB 4: 着地見込み
# ==============================
with tab4:
    st.subheader("🔮 着地見込み（年度末予測）")
    st.caption("登録済みの月次実績をもとに、1月〜12月の年度末着地を予測します")

    monthly = st.session_state.get("monthly_data", {})

    if len(monthly) < 1:
        st.info("先に「月次推移」タブで月次データを登録してください。")
    else:
        months_sorted = sorted(monthly.keys())
        n = len(months_sorted)
        remaining = 12 - max(months_sorted)

        profits = {m: calc_profit(monthly[m]) for m in months_sorted}

        # ---- 月次平均で残月を補完 ----
        avg = {}
        ytd = {}
        for item in ["売上高", "売上総利益", "営業利益", "経常利益"]:
            vals = [profits[m][item] for m in months_sorted]
            ytd[item] = sum(vals)
            avg[item] = sum(vals) / n

        landing = {item: ytd[item] + avg[item] * remaining for item in avg}

        # ---- サマリーメトリクス ----
        st.subheader(f"着地見込み（{max(months_sorted)}月実績 + 残{remaining}ヶ月予測）")
        cols = st.columns(4)
        items = ["売上高", "売上総利益", "営業利益", "経常利益"]
        for i, item in enumerate(items):
            ytd_val = ytd[item]
            land_val = landing[item]
            budget_val = ANNUAL_BUDGET.get(item)
            delta_str = None
            delta_color = "off"
            if budget_val:
                diff = land_val - budget_val
                delta_str = f"予算比 {diff/10000:+.0f}万円"
                delta_color = "normal" if diff >= 0 else "inverse"
            cols[i].metric(
                label=item,
                value=fmt_yen(land_val),
                delta=delta_str,
                delta_color=delta_color,
                help=f"YTD実績: {fmt_yen(ytd_val)}",
            )

        st.divider()

        # ---- 積み上げ棒グラフ：実績 vs 予測 ----
        st.subheader("実績 + 予測の内訳")
        fig_land = go.Figure()
        for item in items:
            actual_vals = [profits[m][item] for m in months_sorted]
            forecast_vals = [avg[item]] * remaining
            x_actual = [f"{m}月" for m in months_sorted]
            x_forecast = [f"{m}月(予)" for m in range(max(months_sorted) + 1, 13)]

        fig_land = go.Figure()
        actual_totals = [ytd[item] for item in items]
        forecast_totals = [avg[item] * remaining for item in items]
        budget_totals = [ANNUAL_BUDGET.get(item, 0) for item in items]

        fig_land.add_trace(go.Bar(name="YTD実績", x=items, y=actual_totals,
                                  marker_color="#1565C0",
                                  text=[fmt_yen(v) for v in actual_totals],
                                  textposition="inside"))
        fig_land.add_trace(go.Bar(name="残月予測", x=items, y=forecast_totals,
                                  marker_color="#90CAF9",
                                  text=[fmt_yen(v) for v in forecast_totals],
                                  textposition="inside"))
        fig_land.add_trace(go.Scatter(name="年間予算", x=items, y=budget_totals,
                                      mode="markers", marker=dict(symbol="line-ew", size=20,
                                      color="red", line=dict(width=3, color="red"))))
        fig_land.update_layout(barmode="stack", height=450, yaxis_title="金額（円）")
        st.plotly_chart(fig_land, use_container_width=True)

        # ---- 月次推移 + 予測ライン ----
        st.subheader("月次推移と予測ライン（売上高）")
        actual_months = [f"{m}月" for m in months_sorted]
        actual_sales = [profits[m]["売上高"] for m in months_sorted]
        forecast_months = [f"{m}月" for m in range(max(months_sorted) + 1, 13)]
        forecast_sales = [avg["売上高"]] * remaining

        fig_fc = go.Figure()
        fig_fc.add_trace(go.Scatter(
            x=actual_months, y=actual_sales, name="実績",
            mode="lines+markers", line=dict(color="#1565C0", width=2), marker=dict(size=8),
        ))
        if forecast_months:
            fig_fc.add_trace(go.Scatter(
                x=[actual_months[-1]] + forecast_months,
                y=[actual_sales[-1]] + forecast_sales,
                name="予測", mode="lines+markers",
                line=dict(color="#90CAF9", width=2, dash="dash"), marker=dict(size=8),
            ))
        fig_fc.add_hline(
            y=ANNUAL_BUDGET["売上高"] / 12, line_dash="dot", line_color="red",
            annotation_text="月次予算ペース",
        )
        fig_fc.update_layout(height=380, yaxis_title="金額（円）", hovermode="x unified")
        st.plotly_chart(fig_fc, use_container_width=True)

        # ---- 達成率見込みゲージ ----
        st.subheader("年間予算 達成見込み")
        gauge_cols = st.columns(4)
        for i, item in enumerate(items):
            budget_val = ANNUAL_BUDGET.get(item, 0)
            rate = round(landing[item] / budget_val * 100, 1) if budget_val else 0
            color = "#4CAF50" if rate >= 100 else "#FFC107" if rate >= 90 else "#EF5350"
            fig_g = go.Figure(go.Indicator(
                mode="gauge+number", value=rate, number={"suffix": "%"},
                title={"text": item},
                gauge={"axis": {"range": [0, 130]}, "bar": {"color": color},
                       "steps": [{"range": [0, 90], "color": "#FFEBEE"},
                                 {"range": [90, 100], "color": "#FFF9C4"},
                                 {"range": [100, 130], "color": "#E8F5E9"}],
                       "threshold": {"line": {"color": "black", "width": 2}, "value": 100}},
            ))
            fig_g.update_layout(height=250, margin=dict(t=40, b=10, l=10, r=10))
            gauge_cols[i].plotly_chart(fig_g, use_container_width=True)

        st.caption(f"※ 予測は{max(months_sorted)}月までの月次平均を残{remaining}ヶ月に適用した単純予測です。")


# ==============================
# TAB 5: 資金繰り
# ==============================
with tab5:
    import datetime

    st.subheader("💰 資金繰り予測（3ヶ月）")
    st.caption("現在の現金残高・売掛金・支払予定から、3ヶ月先までの資金繰りを予測します")

    # ---- 設定入力 ----
    col_set1, col_set2 = st.columns(2)
    with col_set1:
        current_cash = st.number_input(
            "現在の現金残高（円）", min_value=0, value=1_500_000, step=100_000,
            help="今日時点の手元現金・預金の合計"
        )
    with col_set2:
        min_threshold = st.number_input(
            "最低確保したい残高（円）", min_value=0, value=500_000, step=100_000,
            help="この金額を下回るとアラートが出ます"
        )

    # ---- CSV アップロード ----
    col_ar, col_ap = st.columns(2)
    with col_ar:
        st.markdown("**売掛金CSV**")
        st.caption("フォーマット: 取引先,金額,入金予定日")
        ar_file = st.file_uploader("売掛金CSVをアップロード", type=["csv"], key="ar_upload")
    with col_ap:
        st.markdown("**支払予定CSV**")
        st.caption("フォーマット: 支払先,金額,支払予定日")
        ap_file = st.file_uploader("支払予定CSVをアップロード", type=["csv"], key="ap_upload")

    if st.button("サンプルデータで試す", key="sample_cf"):
        st.session_state["use_sample_cf"] = True

    with st.expander("支払いサイト設定（詳細オプション）"):
        ar_delay = st.slider("売掛金の平均回収遅延（日）", 0, 60, 0,
                             help="例：翌月末払いなら15〜45日程度の遅延")

    # ---- データ読み込み ----
    ar_df = ap_df = None
    if ar_file and ap_file:
        ar_df = pd.read_csv(ar_file)
        ap_df = pd.read_csv(ap_file)
        ar_df.columns = ["取引先", "金額", "入金予定日"]
        ap_df.columns = ["支払先", "金額", "支払予定日"]
    elif st.session_state.get("use_sample_cf"):
        ar_df = pd.read_csv(StringIO(SAMPLE_AR))
        ap_df = pd.read_csv(StringIO(SAMPLE_AP))
        ar_df.columns = ["取引先", "金額", "入金予定日"]
        ap_df.columns = ["支払先", "金額", "支払予定日"]

    if ar_df is None or ap_df is None:
        st.info("売掛金・支払予定の両方のCSVをアップロードするか、サンプルデータをお試しください。")
    else:
        ar_df["金額"] = pd.to_numeric(ar_df["金額"], errors="coerce").fillna(0)
        ap_df["金額"] = pd.to_numeric(ap_df["金額"], errors="coerce").fillna(0)
        ar_df["日付"] = pd.to_datetime(ar_df["入金予定日"])
        ap_df["日付"] = pd.to_datetime(ap_df["支払予定日"])

        today = datetime.date.today()
        end_date = today + datetime.timedelta(days=92)  # 約3ヶ月

        # 3ヶ月以内のデータに絞る
        ar_3m = ar_df[ar_df["日付"].dt.date <= end_date].copy()
        ap_3m = ap_df[ap_df["日付"].dt.date <= end_date].copy()

        # 月ごとに集計
        ar_3m["月"] = ar_3m["日付"].dt.to_period("M")
        ap_3m["月"] = ap_3m["日付"].dt.to_period("M")

        months = pd.period_range(start=pd.Timestamp(today), periods=3, freq="M")
        monthly_rows = []
        balance = current_cash
        for month in months:
            inflow = ar_3m[ar_3m["月"] == month]["金額"].sum()
            outflow = ap_3m[ap_3m["月"] == month]["金額"].sum()
            net = inflow - outflow
            balance += net
            monthly_rows.append({
                "月": str(month),
                "入金": inflow,
                "支出": outflow,
                "収支": net,
                "残高": balance,
            })
        df_cf = pd.DataFrame(monthly_rows)

        # ---- サマリーメトリクス ----
        st.divider()
        min_balance = df_cf["残高"].min()
        min_month = df_cf.loc[df_cf["残高"].idxmin(), "月"]
        end_balance = df_cf["残高"].iloc[-1]
        total_inflow = df_cf["入金"].sum()
        total_outflow = df_cf["支出"].sum()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("現在の残高", fmt_yen(current_cash))
        c2.metric("3ヶ月後の残高", fmt_yen(end_balance),
                  delta=fmt_yen(end_balance - current_cash),
                  delta_color="normal" if end_balance >= current_cash else "inverse")
        c3.metric("最低残高", fmt_yen(min_balance),
                  delta=f"{min_month}",
                  delta_color="normal" if min_balance >= min_threshold else "inverse")
        c4.metric("3ヶ月合計 入金", fmt_yen(total_inflow))

        # アラート
        if min_balance < min_threshold:
            st.error(
                f"⚠️ {min_month} に残高が {fmt_yen(min_balance)} まで低下します。"
                f"最低確保額 {fmt_yen(min_threshold)} を下回る見込みです。"
            )
        else:
            st.success(f"✅ 3ヶ月間、残高は最低確保額（{fmt_yen(min_threshold)}）を上回る見込みです。")

        # 黒字倒産リスクスコア
        total_ar = ar_3m["金額"].sum()
        ar_to_cash = total_ar / current_cash if current_cash > 0 else 99
        risk_score = 0
        if ar_to_cash > 3: risk_score += 40
        elif ar_to_cash > 1.5: risk_score += 20
        if min_balance < min_threshold: risk_score += 40
        elif min_balance < min_threshold * 1.5: risk_score += 20
        if ar_delay > 30: risk_score += 20
        elif ar_delay > 15: risk_score += 10
        risk_score = min(100, risk_score)
        risk_color = "#EF5350" if risk_score >= 60 else "#FFC107" if risk_score >= 30 else "#4CAF50"
        risk_label = "高リスク" if risk_score >= 60 else "要注意" if risk_score >= 30 else "低リスク"

        st.subheader("⚠️ 黒字倒産リスクスコア")
        rc1, rc2, rc3 = st.columns(3)
        rc1.metric("リスクスコア", f"{risk_score}/100", delta=risk_label, delta_color="off")
        rc2.metric("売掛金/現金 比率", f"{ar_to_cash:.1f}倍",
                   help="売掛金が現金の何倍あるか。高いほどキャッシュ化前に資金ショートしやすい")
        rc3.metric("3ヶ月最低残高", fmt_yen(min_balance))
        if risk_score >= 60:
            st.error("売掛金の回収促進・早期入金交渉・つなぎ融資の検討をお勧めします。")
        elif risk_score >= 30:
            st.warning("資金繰りに注意が必要です。入金サイトの短縮を検討してください。")

        st.divider()

        # ---- 残高推移グラフ ----
        st.subheader("現金残高 推移")
        balance_points = [current_cash] + df_cf["残高"].tolist()
        month_labels = ["現在"] + df_cf["月"].tolist()

        fig_balance = go.Figure()
        colors_line = ["#EF5350" if b < min_threshold else "#1565C0" for b in balance_points]
        fig_balance.add_trace(go.Scatter(
            x=month_labels, y=balance_points,
            mode="lines+markers+text",
            line=dict(color="#1565C0", width=3),
            marker=dict(size=12,
                        color=["#EF5350" if b < min_threshold else "#1565C0" for b in balance_points]),
            text=[fmt_yen(b) for b in balance_points],
            textposition="top center",
            name="残高",
        ))
        fig_balance.add_hline(
            y=min_threshold, line_dash="dash", line_color="#EF5350",
            annotation_text=f"最低確保額 {fmt_yen(min_threshold)}",
            annotation_position="bottom right",
        )
        fig_balance.update_layout(height=380, yaxis_title="金額（円）", showlegend=False)
        st.plotly_chart(fig_balance, use_container_width=True)

        # ---- 月次入出金ウォーターフォール ----
        st.subheader("月次 入出金 内訳")
        wf_x = []
        wf_y = []
        wf_measure = []
        wf_x.append("期首残高")
        wf_y.append(current_cash)
        wf_measure.append("absolute")
        for _, row in df_cf.iterrows():
            wf_x.append(f"{row['月']} 入金")
            wf_y.append(row["入金"])
            wf_measure.append("relative")
            wf_x.append(f"{row['月']} 支出")
            wf_y.append(-row["支出"])
            wf_measure.append("relative")
        wf_x.append("3ヶ月後残高")
        wf_y.append(end_balance)
        wf_measure.append("total")

        fig_wf = go.Figure(go.Waterfall(
            orientation="v", measure=wf_measure, x=wf_x, y=wf_y,
            connector={"line": {"color": "rgb(63,63,63)"}},
            increasing={"marker": {"color": "#4CAF50"}},
            decreasing={"marker": {"color": "#EF5350"}},
            totals={"marker": {"color": "#1565C0"}},
        ))
        fig_wf.update_layout(height=420, showlegend=False, yaxis_title="金額（円）")
        st.plotly_chart(fig_wf, use_container_width=True)

        # ---- 明細テーブル ----
        st.subheader("入出金 明細（日付順）")

        ar_detail = ar_3m[["日付", "取引先", "金額"]].copy()
        ar_detail["区分"] = "入金"
        ar_detail = ar_detail.rename(columns={"取引先": "相手先"})

        ap_detail = ap_3m[["日付", "支払先", "金額"]].copy()
        ap_detail["区分"] = "支出"
        ap_detail = ap_detail.rename(columns={"支払先": "相手先"})

        detail_all = pd.concat([ar_detail, ap_detail]).sort_values("日付").reset_index(drop=True)
        detail_all["日付"] = detail_all["日付"].dt.strftime("%Y-%m-%d")
        detail_all["金額"] = detail_all["金額"].apply(fmt_yen)
        detail_all = detail_all[["日付", "区分", "相手先", "金額"]]

        def highlight_kubun(row):
            color = "#E8F5E9" if row["区分"] == "入金" else "#FFEBEE"
            return [f"background-color: {color}"] * len(row)

        st.dataframe(
            detail_all.style.apply(highlight_kubun, axis=1),
            use_container_width=True, hide_index=True
        )


# ==============================
# TAB 6: 損益分岐点分析
# ==============================
with tab6:
    st.subheader("📉 損益分岐点分析（CVP分析）")
    st.caption("固定費・変動費を入力して、黒字転換に必要な売上高を計算します")

    col_in1, col_in2 = st.columns(2)
    with col_in1:
        st.markdown("**費用の入力**")
        fixed_cost = st.number_input("固定費（月額・円）", value=500000, step=10000,
                                      help="家賃・人件費・リース料など売上に関係なくかかる費用")
        variable_rate = st.slider("変動費率（%）", min_value=0, max_value=95, value=40,
                                   help="売上高に比例してかかる費用の割合（原材料・外注費など）")
        actual_sales = st.number_input("現在の売上高（月額・円）", value=1000000, step=10000)
    with col_in2:
        st.markdown("**目標設定**")
        target_profit = st.number_input("目標営業利益（月額・円）", value=100000, step=10000,
                                         help="この利益を達成するために必要な売上高を逆算します")

    # 計算
    contribution_rate = (100 - variable_rate) / 100  # 限界利益率
    bep = fixed_cost / contribution_rate if contribution_rate > 0 else 0
    target_sales = (fixed_cost + target_profit) / contribution_rate if contribution_rate > 0 else 0
    actual_profit = actual_sales * contribution_rate - fixed_cost
    margin_of_safety = (actual_sales - bep) / actual_sales * 100 if actual_sales > 0 else 0

    st.divider()

    # メトリクス
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("損益分岐点売上高", fmt_yen(bep), help="この売上を超えると黒字")
    c2.metric("目標達成に必要な売上高", fmt_yen(target_sales))
    c3.metric("現在の営業利益", fmt_yen(actual_profit),
              delta_color="normal" if actual_profit >= 0 else "inverse",
              delta="黒字" if actual_profit >= 0 else "赤字")
    c4.metric("安全余裕率", f"{margin_of_safety:.1f}%",
              help="売上が何%下がっても黒字を維持できるか。高いほど安全")

    if actual_profit < 0:
        st.error(f"現在赤字です。黒字化にはあと {fmt_yen(bep - actual_sales)} の売上増加が必要です。")
    elif margin_of_safety < 20:
        st.warning(f"安全余裕率が低めです。売上が {margin_of_safety:.1f}% 下がると赤字転落します。")
    else:
        st.success(f"黒字経営を維持しています。安全余裕率 {margin_of_safety:.1f}%")

    # CVPグラフ
    st.subheader("売上・費用・利益グラフ")
    max_sales = max(actual_sales, bep, target_sales) * 1.5
    x_range = list(range(0, int(max_sales) + 1, int(max_sales / 20)))

    fig_cvp = go.Figure()
    fig_cvp.add_trace(go.Scatter(
        x=x_range, y=x_range, name="売上高", mode="lines",
        line=dict(color="#1565C0", width=2)
    ))
    fig_cvp.add_trace(go.Scatter(
        x=x_range, y=[fixed_cost + x * (variable_rate / 100) for x in x_range],
        name="総費用", mode="lines", line=dict(color="#EF5350", width=2)
    ))
    fig_cvp.add_vline(x=bep, line_dash="dash", line_color="#FF9800",
                      annotation_text=f"損益分岐点 {fmt_yen(bep)}", annotation_position="top right")
    fig_cvp.add_vline(x=actual_sales, line_dash="dot", line_color="#1565C0",
                      annotation_text=f"現在 {fmt_yen(actual_sales)}", annotation_position="top left")
    if target_sales != bep:
        fig_cvp.add_vline(x=target_sales, line_dash="dot", line_color="#4CAF50",
                          annotation_text=f"目標 {fmt_yen(target_sales)}", annotation_position="bottom right")
    fig_cvp.update_layout(height=420, xaxis_title="売上高（円）", yaxis_title="金額（円）",
                          hovermode="x unified")
    st.plotly_chart(fig_cvp, use_container_width=True)

    # 感度分析テーブル
    st.subheader("感度分析：売上高シナリオ別 営業利益")
    scenarios = [0.7, 0.8, 0.9, 1.0, 1.1, 1.2, 1.3]
    sens_rows = []
    for s in scenarios:
        s_sales = actual_sales * s
        s_profit = s_sales * contribution_rate - fixed_cost
        sens_rows.append({
            "シナリオ": f"売上 {int(s*100)}%",
            "売上高": fmt_yen(s_sales),
            "営業利益": fmt_yen(s_profit),
            "黒字/赤字": "🟢 黒字" if s_profit >= 0 else "🔴 赤字",
        })
    st.dataframe(pd.DataFrame(sens_rows), use_container_width=True, hide_index=True)


# ==============================
# TAB 7: フリーランス試算
# ==============================
with tab7:
    st.subheader("💼 フリーランス・個人事業主 手取り試算")
    st.caption("年収・経費から所得税・住民税・社会保険料を概算し、手取り額を試算します（簡易計算）")

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**収入**")
        calc_type = st.radio("入力方法", ["年収で入力", "稼働日数×単価で入力"], horizontal=True)
        if calc_type == "年収で入力":
            annual_revenue = st.number_input("年間売上高（円）", value=6000000, step=100000)
        else:
            unit_price = st.number_input("日単価（円）", value=50000, step=5000)
            working_days = st.number_input("年間稼働日数", value=200, step=5)
            annual_revenue = unit_price * working_days
            st.info(f"年間売上高: {fmt_yen(annual_revenue)}")

    with col_b:
        st.markdown("**経費・控除**")
        annual_expense = st.number_input("年間経費（円）", value=1000000, step=100000,
                                          help="交通費・通信費・消耗品・外注費など")
        aokiro = st.selectbox("青色申告特別控除",
                               ["65万円（e-Tax）", "55万円（紙申告）", "10万円（簡易帳簿）", "なし（白色申告）"])
        aokiro_map = {"65万円（e-Tax）": 650000, "55万円（紙申告）": 550000,
                      "10万円（簡易帳簿）": 100000, "なし（白色申告）": 0}
        aokiro_deduct = aokiro_map[aokiro]

    # 所得税計算
    def calc_income_tax(taxable_income):
        brackets = [
            (1950000, 0.05, 0),
            (3300000, 0.10, 97500),
            (6950000, 0.20, 427500),
            (9000000, 0.23, 636000),
            (18000000, 0.33, 1536000),
            (40000000, 0.40, 2796000),
            (float("inf"), 0.45, 4796000),
        ]
        if taxable_income <= 0:
            return 0
        for limit, rate, deduction in brackets:
            if taxable_income <= limit:
                tax = taxable_income * rate - deduction
                return max(0, int(tax * 1.021))  # 復興特別所得税込み
        return 0

    jigyou_shotoku = max(0, annual_revenue - annual_expense - aokiro_deduct)
    nenkin = 203760  # 国民年金 2024年
    kokuho = max(0, int((jigyou_shotoku - 430000) * 0.10))  # 国民健康保険（簡易）
    shakai_hoken = nenkin + kokuho
    kiso_kojo = 480000  # 基礎控除
    taxable = max(0, jigyou_shotoku - shakai_hoken - kiso_kojo)
    shotoku_zei = calc_income_tax(taxable)
    jumin_zei = max(0, int(taxable * 0.10))
    total_tax = shotoku_zei + jumin_zei + shakai_hoken
    take_home = annual_revenue - annual_expense - total_tax
    effective_rate = total_tax / annual_revenue * 100 if annual_revenue > 0 else 0

    st.divider()

    # メトリクス
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("年間売上高", fmt_yen(annual_revenue))
    c2.metric("事業所得", fmt_yen(jigyou_shotoku))
    c3.metric("税・社保 合計", fmt_yen(total_tax))
    c4.metric("手取り（概算）", fmt_yen(take_home),
              delta=f"実効負担率 {effective_rate:.1f}%",
              delta_color="off")

    # 内訳グラフ
    st.subheader("手取りの内訳")
    labels = ["手取り", "所得税", "住民税", "国民健康保険", "国民年金", "経費"]
    values = [take_home, shotoku_zei, jumin_zei, kokuho, nenkin, annual_expense]
    colors = ["#4CAF50", "#EF5350", "#FF7043", "#FFA726", "#FFCA28", "#90CAF9"]
    fig_pie = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.45,
        marker=dict(colors=colors),
        textinfo="label+percent",
    ))
    fig_pie.update_layout(height=380, showlegend=True)
    st.plotly_chart(fig_pie, use_container_width=True)

    # 詳細内訳テーブル
    st.subheader("詳細内訳")
    breakdown = [
        {"項目": "年間売上高", "金額": fmt_yen(annual_revenue), "備考": ""},
        {"項目": "　経費", "金額": f"▲ {fmt_yen(annual_expense)}", "備考": ""},
        {"項目": "　青色申告特別控除", "金額": f"▲ {fmt_yen(aokiro_deduct)}", "備考": aokiro},
        {"項目": "事業所得", "金額": fmt_yen(jigyou_shotoku), "備考": ""},
        {"項目": "　国民年金", "金額": f"▲ {fmt_yen(nenkin)}", "備考": "月16,980円×12ヶ月"},
        {"項目": "　国民健康保険", "金額": f"▲ {fmt_yen(kokuho)}", "備考": "自治体により異なる（概算）"},
        {"項目": "　基礎控除", "金額": f"▲ {fmt_yen(kiso_kojo)}", "備考": ""},
        {"項目": "課税所得", "金額": fmt_yen(taxable), "備考": ""},
        {"項目": "　所得税（復興税込）", "金額": f"▲ {fmt_yen(shotoku_zei)}", "備考": ""},
        {"項目": "　住民税", "金額": f"▲ {fmt_yen(jumin_zei)}", "備考": "一律10%"},
        {"項目": "手取り（概算）", "金額": fmt_yen(take_home), "備考": ""},
    ]
    st.dataframe(pd.DataFrame(breakdown), use_container_width=True, hide_index=True)
    st.caption("※ 本試算は概算です。実際の税額は税理士にご確認ください。")

    # 節税シミュレーション
    st.divider()
    st.subheader("💡 節税シミュレーション")
    ideco_monthly = st.slider("iDeCo 掛け金（月額・円）", 0, 68000, 20000, step=1000,
                               help="自営業者の上限は月68,000円")
    kyosai_monthly = st.slider("小規模企業共済 掛け金（月額・円）", 0, 70000, 30000, step=1000,
                                help="掛け金全額が所得控除になる")
    ideco_annual   = ideco_monthly * 12
    kyosai_annual  = kyosai_monthly * 12
    new_taxable    = max(0, taxable - ideco_annual - kyosai_annual)
    new_tax        = calc_income_tax_jp(new_taxable) + max(0, int(new_taxable * 0.10))
    tax_saving     = (shotoku_zei + jumin_zei) - new_tax
    new_takehome   = take_home + tax_saving - ideco_annual - kyosai_annual
    actual_saving  = new_takehome - take_home

    sc1, sc2, sc3 = st.columns(3)
    sc1.metric("節税額（税金減少）", fmt_yen(tax_saving))
    sc2.metric("掛け金合計（年）", fmt_yen(ideco_annual + kyosai_annual))
    sc3.metric("実質手取り増減", fmt_yen(actual_saving),
               delta_color="normal" if actual_saving >= 0 else "inverse")
    st.caption("※ iDeCo・共済の掛け金は将来の受取時に課税されます。手取りへの即時効果と老後資産形成のバランスで検討ください。")


# ==============================
# TAB 8: 前期比較
# ==============================
with tab8:
    st.subheader("📊 前期比較")
    st.caption("今期と前期の試算表CSVを読み込んで、KPIの変化・成長率を比較します")

    col_c, col_d = st.columns(2)
    with col_c:
        st.markdown("**今期**")
        current_file = st.file_uploader("今期CSVをアップロード", type=["csv"], key="current_upload")
    with col_d:
        st.markdown("**前期**")
        prev_file = st.file_uploader("前期CSVをアップロード", type=["csv"], key="prev_upload")

    if st.button("サンプルデータで試す", key="sample_yoy"):
        st.session_state["use_sample_yoy"] = True

    cur_data = prev_data = None
    if current_file and prev_file:
        cur_data = load_data(current_file)
        prev_data = load_data(prev_file)
    elif st.session_state.get("use_sample_yoy"):
        cur_data = load_data(StringIO(SAMPLE_ACTUAL))
        prev_str = """勘定科目,金額
売上高,44000000
売上原価,27000000
販売費及び一般管理費,9000000
営業外収益,400000
営業外費用,280000
特別利益,0
特別損失,0
流動資産,22000000
固定資産,14000000
流動負債,9000000
固定負債,8000000
純資産,19000000
売上債権,7000000
棚卸資産,4500000
仕入債務,3500000
"""
        prev_data = load_data(StringIO(prev_str))

    if cur_data is None or prev_data is None:
        st.info("今期・前期の両方のCSVをアップロードするか、サンプルデータをお試しください。")
    else:
        cur_kpis = calc_kpis(cur_data)
        prev_kpis = calc_kpis(prev_data)

        # 比較テーブル作成
        compare_items = [
            ("売上高", False),
            ("売上総利益", False),
            ("営業利益", False),
            ("経常利益", False),
            ("売上総利益率", True),
            ("営業利益率", True),
            ("経常利益率", True),
            ("流動比率", True),
            ("自己資本比率", True),
        ]

        rows = []
        for key, is_pct in compare_items:
            cur_val = cur_kpis.get(key)
            prev_val = prev_kpis.get(key)
            if cur_val is None or prev_val is None:
                continue
            change = cur_val - prev_val
            change_rate = round(change / abs(prev_val) * 100, 1) if prev_val else None
            unit = "%" if is_pct else "円"
            rows.append({
                "項目": key,
                "前期": f"{prev_val}{unit}" if is_pct else fmt_yen(prev_val),
                "今期": f"{cur_val}{unit}" if is_pct else fmt_yen(cur_val),
                "増減": f"{'+' if change >= 0 else ''}{change:.1f}{unit}" if is_pct
                        else fmt_yen(change),
                "増減率": f"{'+' if (change_rate or 0) >= 0 else ''}{change_rate}%" if change_rate else "N/A",
                "_change": change,
            })

        df_compare = pd.DataFrame(rows)

        # サマリーメトリクス
        st.subheader("主要KPI 前期比")
        m_cols = st.columns(4)
        for i, key in enumerate(["売上高", "売上総利益", "営業利益", "経常利益"]):
            cur_v = cur_kpis[key]
            prev_v = prev_kpis[key]
            diff = cur_v - prev_v
            rate = round(diff / abs(prev_v) * 100, 1) if prev_v else 0
            m_cols[i].metric(key, fmt_yen(cur_v),
                              delta=f"{'+' if rate >= 0 else ''}{rate}%",
                              delta_color="normal" if diff >= 0 else "inverse")

        st.divider()

        # 前期vs今期 グループバーチャート
        st.subheader("売上・利益 前期比較")
        bar_items = ["売上高", "売上総利益", "営業利益", "経常利益"]
        fig_yoy = go.Figure()
        fig_yoy.add_trace(go.Bar(
            name="前期", x=bar_items,
            y=[prev_kpis[k] for k in bar_items],
            marker_color="#90CAF9",
            text=[fmt_yen(prev_kpis[k]) for k in bar_items],
            textposition="outside",
        ))
        fig_yoy.add_trace(go.Bar(
            name="今期", x=bar_items,
            y=[cur_kpis[k] for k in bar_items],
            marker_color="#1565C0",
            text=[fmt_yen(cur_kpis[k]) for k in bar_items],
            textposition="outside",
        ))
        fig_yoy.update_layout(barmode="group", height=420, yaxis_title="金額（円）")
        st.plotly_chart(fig_yoy, use_container_width=True)

        # 利益率 レーダーチャート
        st.subheader("KPI レーダーチャート（前期vs今期）")
        radar_items = ["売上総利益率", "営業利益率", "経常利益率", "流動比率", "自己資本比率"]
        def safe_val(kpis, key):
            v = kpis.get(key)
            return v if v is not None else 0

        fig_radar = go.Figure()
        fig_radar.add_trace(go.Scatterpolar(
            r=[safe_val(prev_kpis, k) for k in radar_items] + [safe_val(prev_kpis, radar_items[0])],
            theta=radar_items + [radar_items[0]],
            fill="toself", name="前期", line_color="#90CAF9",
        ))
        fig_radar.add_trace(go.Scatterpolar(
            r=[safe_val(cur_kpis, k) for k in radar_items] + [safe_val(cur_kpis, radar_items[0])],
            theta=radar_items + [radar_items[0]],
            fill="toself", name="今期", line_color="#1565C0",
        ))
        fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True)), height=420)
        st.plotly_chart(fig_radar, use_container_width=True)

        # 詳細比較テーブル
        st.subheader("全KPI 前期比較テーブル")
        def color_change(val):
            if isinstance(val, str) and val.startswith("+"):
                return "color: #2E7D32"
            elif isinstance(val, str) and val.startswith("-") and val != "-":
                return "color: #C62828"
            return ""
        display_df = df_compare.drop(columns=["_change"])
        st.dataframe(
            display_df.style.applymap(color_change, subset=["増減", "増減率"]),
            use_container_width=True, hide_index=True
        )


# ==============================
# TAB 9: 財務健全度スコア
# ==============================
with tab9:
    st.subheader("🏆 財務健全度スコア")
    st.caption("KPI分析タブでデータを読み込んだ後に使用してください")

    kpi_src = None
    if st.session_state.get("pdf_extracted") is not None:
        kpi_src = st.session_state["pdf_extracted"]
    elif st.session_state.get("use_sample"):
        kpi_src = load_data(StringIO(SAMPLE_ACTUAL))

    if kpi_src is None:
        st.info("先に「KPI分析」タブでデータを読み込んでください。")
    else:
        kpis_s = calc_kpis(kpi_src)
        scores = calc_health_score(kpis_s)

        # 総合スコアゲージ
        total = scores["総合"]
        color = "#4CAF50" if total >= 70 else "#FFC107" if total >= 40 else "#EF5350"
        label = "優良" if total >= 70 else "普通" if total >= 40 else "要改善"

        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=total,
            title={"text": f"総合スコア　{label}"},
            delta={"reference": 70, "valueformat": ".0f"},
            gauge={
                "axis": {"range": [0, 100]},
                "bar": {"color": color},
                "steps": [
                    {"range": [0, 40],  "color": "#FFEBEE"},
                    {"range": [40, 70], "color": "#FFF9C4"},
                    {"range": [70, 100],"color": "#E8F5E9"},
                ],
                "threshold": {"line": {"color": "black", "width": 3}, "value": 70},
            },
        ))
        fig_gauge.update_layout(height=300)
        st.plotly_chart(fig_gauge, use_container_width=True)

        # カテゴリ別スコア
        st.subheader("カテゴリ別スコア")
        cats = ["収益性", "安全性", "効率性"]
        cat_scores = [scores[c] for c in cats]
        cat_colors = ["#4CAF50" if s >= 70 else "#FFC107" if s >= 40 else "#EF5350" for s in cat_scores]

        c1, c2, c3 = st.columns(3)
        for col, cat, score, clr in zip([c1, c2, c3], cats, cat_scores, cat_colors):
            fig_mini = go.Figure(go.Indicator(
                mode="gauge+number",
                value=score,
                title={"text": cat},
                gauge={
                    "axis": {"range": [0, 100]},
                    "bar": {"color": clr},
                    "steps": [{"range": [0, 40], "color": "#FFEBEE"},
                               {"range": [40, 70], "color": "#FFF9C4"},
                               {"range": [70, 100], "color": "#E8F5E9"}],
                },
            ))
            fig_mini.update_layout(height=220, margin=dict(t=40, b=10, l=10, r=10))
            col.plotly_chart(fig_mini, use_container_width=True)

        # レーダーチャート
        st.subheader("スコアレーダー")
        fig_r = go.Figure(go.Scatterpolar(
            r=cat_scores + [cat_scores[0]],
            theta=cats + [cats[0]],
            fill="toself",
            line_color=color,
            fillcolor=color.replace(")", ",0.2)").replace("rgb", "rgba") if "rgb" in color else color,
        ))
        fig_r.update_layout(polar=dict(radialaxis=dict(range=[0, 100])), height=380)
        st.plotly_chart(fig_r, use_container_width=True)

        # 改善アドバイス
        st.subheader("改善ポイント")
        advices = []
        if scores["収益性"] < 70:
            if (kpis_s.get("売上総利益率") or 0) < 30:
                advices.append("**売上総利益率が低め** — 原価削減・単価引き上げを検討してください")
            if (kpis_s.get("営業利益率") or 0) < 5:
                advices.append("**営業利益率が低め** — 販管費の見直し・固定費削減が効果的です")
        if scores["安全性"] < 70:
            if (kpis_s.get("流動比率") or 0) < 120:
                advices.append("**流動比率が低め** — 短期借入の返済計画見直しや売掛金の早期回収を")
            if (kpis_s.get("自己資本比率") or 0) < 30:
                advices.append("**自己資本比率が低め** — 利益の内部留保を積み上げて財務基盤を強化しましょう")
        if scores["効率性"] < 70:
            advices.append("**売上債権回転日数が長め** — 入金サイトの短縮・早期入金割引の活用を検討してください")

        if advices:
            for a in advices:
                st.markdown(f"- {a}")
        else:
            st.success("全カテゴリで良好なスコアです。現状の財務規律を維持してください。")


# ==============================
# TAB 10: What-if シミュレーター
# ==============================
with tab10:
    st.subheader("🔬 What-if シミュレーター")
    st.caption("売上・コストを変化させたときのKPIへの影響をリアルタイムで確認できます")

    kpi_base = None
    if st.session_state.get("pdf_extracted") is not None:
        kpi_base = st.session_state["pdf_extracted"]
    elif st.session_state.get("use_sample"):
        kpi_base = load_data(StringIO(SAMPLE_ACTUAL))

    if kpi_base is None:
        st.info("先に「KPI分析」タブでデータを読み込んでください。")
    else:
        base_kpis = calc_kpis(kpi_base)
        st.markdown("#### ベースライン")
        b1, b2, b3 = st.columns(3)
        b1.metric("売上高", fmt_yen(base_kpis["売上高"]))
        b2.metric("営業利益", fmt_yen(base_kpis["営業利益"]))
        b3.metric("営業利益率", f"{base_kpis['営業利益率']}%")

        st.divider()
        st.markdown("#### 変化を設定してください")
        wc1, wc2 = st.columns(2)
        with wc1:
            sales_chg   = st.slider("売上高の変化率 (%)", -50, 100, 0)
            cogs_chg    = st.slider("売上原価の変化率 (%)", -50, 50, 0)
        with wc2:
            sga_chg     = st.slider("販管費の変化率 (%)", -50, 50, 0)
            fixed_cut   = st.number_input("固定費の削減額（円）", value=0, step=100000)

        # What-if計算
        w = kpi_base.copy()
        w["売上高"] = w.get("売上高", 0) * (1 + sales_chg / 100)
        w["売上原価"] = w.get("売上原価", 0) * (1 + cogs_chg / 100)
        w["販売費及び一般管理費"] = max(0, w.get("販売費及び一般管理費", 0) * (1 + sga_chg / 100) - fixed_cut)
        new_kpis = calc_kpis(w)

        st.divider()
        st.markdown("#### シミュレーション結果")
        items_wi = [
            ("売上高", False), ("売上総利益", False), ("営業利益", False),
            ("売上総利益率", True), ("営業利益率", True), ("経常利益率", True),
        ]
        cols_wi = st.columns(3)
        for i, (key, is_pct) in enumerate(items_wi):
            base_v = base_kpis.get(key) or 0
            new_v  = new_kpis.get(key) or 0
            diff   = new_v - base_v
            unit   = "%" if is_pct else ""
            fmt    = (lambda v: f"{v:.1f}%") if is_pct else fmt_yen
            delta_str = f"{'+' if diff >= 0 else ''}{diff:.1f}%" if is_pct else fmt_yen(diff)
            cols_wi[i % 3].metric(key, fmt(new_v), delta=delta_str,
                                   delta_color="normal" if diff >= 0 else "inverse")

        # ビフォーアフターグラフ
        st.subheader("ビフォーアフター比較")
        bar_keys = ["売上高", "売上総利益", "営業利益"]
        fig_wi = go.Figure()
        fig_wi.add_trace(go.Bar(name="現状", x=bar_keys,
                                 y=[base_kpis.get(k, 0) for k in bar_keys],
                                 marker_color="#90CAF9",
                                 text=[fmt_yen(base_kpis.get(k, 0)) for k in bar_keys],
                                 textposition="outside"))
        fig_wi.add_trace(go.Bar(name="What-if", x=bar_keys,
                                 y=[new_kpis.get(k, 0) for k in bar_keys],
                                 marker_color="#1565C0",
                                 text=[fmt_yen(new_kpis.get(k, 0)) for k in bar_keys],
                                 textposition="outside"))
        fig_wi.update_layout(barmode="group", height=400, yaxis_title="金額（円）")
        st.plotly_chart(fig_wi, use_container_width=True)

        # 利益率の変化
        st.subheader("利益率の変化")
        rate_keys = ["売上総利益率", "営業利益率", "経常利益率"]
        fig_rate_wi = go.Figure()
        fig_rate_wi.add_trace(go.Bar(name="現状", x=rate_keys,
                                      y=[base_kpis.get(k, 0) for k in rate_keys],
                                      marker_color="#90CAF9"))
        fig_rate_wi.add_trace(go.Bar(name="What-if", x=rate_keys,
                                      y=[new_kpis.get(k, 0) for k in rate_keys],
                                      marker_color="#1565C0"))
        fig_rate_wi.update_layout(barmode="group", height=350, yaxis_title="%")
        st.plotly_chart(fig_rate_wi, use_container_width=True)


# ==============================
# TAB 11: 法人化判断シミュレーター
# ==============================
with tab11:
    st.subheader("🏢 法人化判断シミュレーター")
    st.caption("個人事業主と法人化した場合の手取り・税負担を比較します（簡易計算）")

    ic1, ic2 = st.columns(2)
    with ic1:
        corp_revenue  = st.number_input("年間売上高（円）", value=10000000, step=500000, key="corp_rev")
        corp_expense  = st.number_input("年間経費（役員報酬以外）", value=2000000, step=100000, key="corp_exp",
                                         help="仕入・外注費・家賃など")
    with ic2:
        owner_salary  = st.number_input("法人化した場合の役員報酬（年額）", value=6000000, step=100000,
                                         help="法人から自分に支払う給与。節税の主要な調整弁です")
        st.caption("役員報酬を変えると法人税と個人の手取りが変わります")

    result = calc_incorporation(corp_revenue, corp_expense, owner_salary)

    st.divider()

    # 比較メトリクス
    col_ind, col_corp = st.columns(2)
    with col_ind:
        st.markdown("### 個人事業主")
        st.metric("手取り（年）", fmt_yen(result["個人_手取り"]))
        st.metric("税・社保 合計", fmt_yen(result["個人_税社保"]))

    with col_corp:
        st.markdown("### 法人化")
        st.metric("手取り（年）", fmt_yen(result["法人_手取り"]),
                   delta=fmt_yen(result["差額"]),
                   delta_color="normal" if result["差額"] >= 0 else "inverse")
        st.metric("税・社保 合計（法人税含む）", fmt_yen(result["法人_税社保"]))
        st.metric("法人内部留保", fmt_yen(result["法人_内部留保"]),
                   help="法人に残るお金。将来の設備投資・退職金の原資になります")

    if result["差額"] > 0:
        st.success(f"法人化すると手取りが年 {fmt_yen(result['差額'])} 増える見込みです。")
    elif result["差額"] < 0:
        st.warning(f"現在の条件では個人事業主のほうが手取りが {fmt_yen(abs(result['差額']))} 多い見込みです。役員報酬を調整してみてください。")

    st.divider()

    # 比較グラフ
    st.subheader("手取り vs 税・社保 比較")
    fig_corp = go.Figure()
    fig_corp.add_trace(go.Bar(
        name="手取り", x=["個人事業主", "法人化"],
        y=[result["個人_手取り"], result["法人_手取り"]],
        marker_color="#4CAF50",
        text=[fmt_yen(result["個人_手取り"]), fmt_yen(result["法人_手取り"])],
        textposition="outside",
    ))
    fig_corp.add_trace(go.Bar(
        name="税・社保", x=["個人事業主", "法人化"],
        y=[result["個人_税社保"], result["法人_税社保"]],
        marker_color="#EF5350",
        text=[fmt_yen(result["個人_税社保"]), fmt_yen(result["法人_税社保"])],
        textposition="outside",
    ))
    fig_corp.update_layout(barmode="group", height=420, yaxis_title="金額（円）")
    st.plotly_chart(fig_corp, use_container_width=True)

    # 役員報酬別シミュレーション
    st.subheader("役員報酬を変えたときの手取り推移")
    salary_range = range(2000000, min(corp_revenue - corp_expense, 15000000), 500000)
    sim_rows = [calc_incorporation(corp_revenue, corp_expense, s) for s in salary_range]
    fig_sim = go.Figure()
    fig_sim.add_trace(go.Scatter(
        x=list(salary_range), y=[r["法人_手取り"] for r in sim_rows],
        name="法人化 手取り", mode="lines", line=dict(color="#1565C0", width=2),
    ))
    fig_sim.add_hline(y=result["個人_手取り"], line_dash="dash", line_color="#EF5350",
                      annotation_text=f"個人事業主 手取り {fmt_yen(result['個人_手取り'])}",
                      annotation_position="right")
    fig_sim.add_vline(x=owner_salary, line_dash="dot", line_color="#FFC107",
                      annotation_text="現在の設定", annotation_position="top")
    fig_sim.update_layout(height=380, xaxis_title="役員報酬（円）", yaxis_title="手取り（円）")
    st.plotly_chart(fig_sim, use_container_width=True)
    st.caption("※ 本試算は概算です。実際の法人化判断は税理士・司法書士にご相談ください。")
