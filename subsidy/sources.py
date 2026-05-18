"""海外輸出関連の補助金を発信する既定ソース定義

各ソースは Google Custom Search の `site:` 修飾子で範囲を絞って検索する。
カテゴリは「行政 / 金融機関 / 自治体 / 公的機関 / 営利法人・団体」のいずれか。
"""

from __future__ import annotations

# --- 既定ソース ---
# Google Custom Search の site: で絞り込むためのドメインと、ヒトに見せる機関名。
PREDEFINED_SOURCES: list[dict] = [
    # === 行政（中央省庁） ===
    {"name": "経済産業省", "category": "行政", "domain": "meti.go.jp"},
    {"name": "中小企業庁", "category": "行政", "domain": "chusho.meti.go.jp"},
    {"name": "農林水産省", "category": "行政", "domain": "maff.go.jp"},
    {"name": "外務省", "category": "行政", "domain": "mofa.go.jp"},
    {"name": "総務省", "category": "行政", "domain": "soumu.go.jp"},

    # === 公的機関（独立行政法人・公益法人） ===
    {"name": "JETRO（日本貿易振興機構）", "category": "公的機関", "domain": "jetro.go.jp"},
    {"name": "NEDO（新エネルギー・産業技術総合開発機構）", "category": "公的機関", "domain": "nedo.go.jp"},
    {"name": "中小企業基盤整備機構", "category": "公的機関", "domain": "smrj.go.jp"},
    {"name": "J-Net21（中小機構運営）", "category": "公的機関", "domain": "j-net21.smrj.go.jp"},
    {"name": "ミラサポplus", "category": "公的機関", "domain": "mirasapo-plus.go.jp"},
    {"name": "JFOODO（日本食品海外プロモーションセンター）", "category": "公的機関", "domain": "jfoodo.jetro.go.jp"},
    {"name": "JICA（国際協力機構）", "category": "公的機関", "domain": "jica.go.jp"},

    # === 金融機関 ===
    {"name": "JBIC（国際協力銀行）", "category": "金融機関", "domain": "jbic.go.jp"},
    {"name": "日本政策金融公庫", "category": "金融機関", "domain": "jfc.go.jp"},
    {"name": "商工組合中央金庫", "category": "金融機関", "domain": "shokochukin.co.jp"},
    {"name": "NEXI（日本貿易保険）", "category": "金融機関", "domain": "nexi.go.jp"},
    {"name": "DBJ（日本政策投資銀行）", "category": "金融機関", "domain": "dbj.jp"},

    # === 自治体（代表的な海外展開支援機関） ===
    {"name": "東京都産業労働局", "category": "自治体", "domain": "sangyo-rodo.metro.tokyo.lg.jp"},
    {"name": "東京都中小企業振興公社", "category": "自治体", "domain": "tokyo-kosha.or.jp"},
    {"name": "大阪産業局", "category": "自治体", "domain": "obda.or.jp"},
    {"name": "大阪府", "category": "自治体", "domain": "pref.osaka.lg.jp"},
    {"name": "愛知県", "category": "自治体", "domain": "pref.aichi.jp"},
    {"name": "神奈川県産業振興センター", "category": "自治体", "domain": "kipc.or.jp"},
    {"name": "横浜企業経営支援財団", "category": "自治体", "domain": "idec.or.jp"},
    {"name": "京都府", "category": "自治体", "domain": "pref.kyoto.jp"},
    {"name": "兵庫県", "category": "自治体", "domain": "web.pref.hyogo.lg.jp"},
    {"name": "福岡県", "category": "自治体", "domain": "pref.fukuoka.lg.jp"},
    {"name": "北海道", "category": "自治体", "domain": "pref.hokkaido.lg.jp"},

    # === 営利法人・団体（民間でも補助金/助成金情報を発信） ===
    {"name": "ジェグテック（中小機構 民間連携）", "category": "営利団体", "domain": "jgoodtech.smrj.go.jp"},
    {"name": "日本商工会議所", "category": "営利団体", "domain": "jcci.or.jp"},
    {"name": "東京商工会議所", "category": "営利団体", "domain": "tokyo-cci.or.jp"},
    {"name": "経団連（日本経済団体連合会）", "category": "営利団体", "domain": "keidanren.or.jp"},
]


# --- キーワード（海外輸出に関係する語） ---
# 検索クエリは「キーワード」と「補助金」「助成金」「補助事業」などを組み合わせる。
EXPORT_KEYWORDS: list[str] = [
    "海外展開",
    "海外輸出",
    "輸出促進",
    "越境EC",
    "国際化",
    "海外進出",
    "海外販路開拓",
    "海外見本市",
    "海外プロモーション",
]

SUBSIDY_TERMS: list[str] = [
    "補助金",
    "助成金",
    "補助事業",
    "支援事業",
]
