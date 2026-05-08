# CFO Cockpit

> 年商1〜5億円規模の中小企業オーナーのための財務分析ダッシュボード

試算表PDFをアップロードするだけで、KPI診断・トレンド分析・経営シミュレーションを一気通貫で提供します。
Anthropic Claude を活用し、CFO目線の所見と打ち手を自動生成。

## デモ

[https://financial-dashboard-2zpfnq7jneyn8pyb5xya55.streamlit.app/](https://financial-dashboard-2zpfnq7jneyn8pyb5xya55.streamlit.app/)

## 主な機能

| 機能 | 内容 |
|------|------|
| **KPI診断** | PDF/CSV試算表をアップロード → 業種別ベンチマークと比較した健全度スコアを生成。AIがCFO目線でレポートを書き出します。 |
| **月次推移分析** | 12ヶ月の売上・利益を可視化。3ヶ月移動平均・異常値検知・年間着地予測を含みます。 |
| **What-if シミュレーション** | 売上・原価率・固定費の変動が営業利益にどう影響するかをスライダーでリアルタイム計算。感度分析グラフ付き。 |

## 設計思想

- **業種別ベンチマーク**: IT・コンサル・卸売・製造・建設・サービス業の6業種に対応。粗利率・営業利益率・流動比率・自己資本比率の業種目安と自社実績を比較します。
- **AI生成レポート**: 単なる数字の羅列ではなく、Claude APIによる「経営者向けの所見と打ち手」を生成。エグゼクティブサマリー・強み・課題と打ち手の3パート構成。
- **PDF AI読み取り**: freee / マネーフォワード等の試算表PDFをClaudeが解析し、勘定科目を自動抽出。

## 技術スタック

- **フレームワーク**: [Streamlit](https://streamlit.io/)
- **データ処理**: pandas
- **可視化**: Plotly
- **PDF解析**: PyMuPDF
- **AI**: [Anthropic Claude API](https://www.anthropic.com/)（claude-sonnet-4-6）

## ローカル起動

```bash
git clone https://github.com/<your-username>/financial-dashboard.git
cd financial-dashboard

pip install -r requirements.txt

# .env を作成し ANTHROPIC_API_KEY を設定
echo "ANTHROPIC_API_KEY=sk-ant-xxxx" > .env

streamlit run app.py
```

## CSVフォーマット

```csv
勘定科目,金額
売上高,180000000
売上原価,108000000
販売費及び一般管理費,48000000
営業外収益,800000
営業外費用,500000
流動資産,72000000
固定資産,45000000
流動負債,38000000
固定負債,22000000
純資産,57000000
売上債権,28000000
棚卸資産,15000000
仕入債務,18000000
```

## ライセンス

MIT License
