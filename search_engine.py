"""Google Custom Search APIを使った企業検索エンジン + 電話番号・住所スクレイピング"""

from __future__ import annotations

import re
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from urllib.parse import urlparse
from datetime import datetime
from typing import Optional

import requests
from bs4 import BeautifulSoup


PREFECTURES = [
    "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
    "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
    "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県",
    "岐阜県", "静岡県", "愛知県", "三重県",
    "滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県",
    "鳥取県", "島根県", "岡山県", "広島県", "山口県",
    "徳島県", "香川県", "愛媛県", "高知県",
    "福岡県", "佐賀県", "長崎県", "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県",
]

INDUSTRIES = [
    "IT/SaaS",
    "製造業",
    "小売・流通",
    "金融・保険",
    "不動産",
    "医療・ヘルスケア",
    "建設・土木",
    "教育",
    "飲食・フードサービス",
    "広告・メディア",
    "物流・運輸",
    "エネルギー",
    "コンサルティング",
    "人材サービス",
]

PRODUCT_GENRES = [
    "ソフトウェア・アプリ",
    "ハードウェア・機器",
    "クラウドサービス",
    "食品・飲料",
    "化粧品・日用品",
    "アパレル・ファッション",
    "建材・資材",
    "医薬品・医療機器",
    "自動車・部品",
    "電子部品・半導体",
    "化学製品",
    "金融商品・保険",
    "教育・研修サービス",
    "人材・採用サービス",
    "広告・マーケティング",
    "物流・倉庫サービス",
    "コンサルティング",
    "不動産・施設管理",
]

MANUFACTURING_CATEGORIES = [
    "食品・飲料",
    "菓子・スイーツ",
    "化粧品・スキンケア",
    "日用品・トイレタリー",
    "アパレル・衣料品",
    "家具・インテリア",
    "家電・生活家電",
    "キッチン用品・調理器具",
    "文具・事務用品",
    "玩具・ホビー",
    "スポーツ用品",
    "ペット用品",
    "ベビー・キッズ用品",
    "健康食品・サプリメント",
    "その他消費財",
]

# 電話番号の正規表現パターン
_PHONE_PATTERNS = [
    re.compile(r"(?:TEL|tel|Tel|電話|Phone|phone|☎|℡)[：:\s]*(\d{2,4}[-ー]\d{1,4}[-ー]\d{3,4})"),
    re.compile(r"(?:TEL|tel|Tel|電話|Phone|phone|☎|℡)[：:\s]*(0\d{9,10})"),
    re.compile(r"(0\d{1,4}[-ー]\d{1,4}[-ー]\d{3,4})"),
]

# 住所の正規表現パターン
_POSTAL_PATTERN = re.compile(r"〒?\s*(\d{3}[-ー]\d{4})")
_ADDRESS_PATTERN = re.compile(
    r"((?:北海道|青森県|岩手県|宮城県|秋田県|山形県|福島県|茨城県|栃木県|群馬県|埼玉県|千葉県|東京都|神奈川県|"
    r"新潟県|富山県|石川県|福井県|山梨県|長野県|岐阜県|静岡県|愛知県|三重県|"
    r"滋賀県|京都府|大阪府|兵庫県|奈良県|和歌山県|"
    r"鳥取県|島根県|岡山県|広島県|山口県|徳島県|香川県|愛媛県|高知県|"
    r"福岡県|佐賀県|長崎県|熊本県|大分県|宮崎県|鹿児島県|沖縄県)"
    r"[^\n\r<]{2,40})"
)


def _extract_phone(text: str) -> str:
    """テキストから電話番号を抽出する"""
    for pattern in _PHONE_PATTERNS:
        m = pattern.search(text)
        if m:
            phone = m.group(1).replace("ー", "-")
            # フリーダイヤル（0120/0800）は除外しない
            if len(phone.replace("-", "")) >= 10:
                return phone
    return ""


def _extract_address(text: str) -> str:
    """テキストから住所を抽出する"""
    m = _ADDRESS_PATTERN.search(text)
    if m:
        addr = m.group(1).strip()
        # HTMLタグや不要な文字を除去
        addr = re.sub(r"<[^>]+>", "", addr)
        return addr
    return ""


def scrape_company_details(url: str) -> dict:
    """企業サイトにアクセスして電話番号・住所を抽出する

    Returns:
        {"phone": str, "address": str}
    """
    result = {"phone": "", "address": ""}

    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                          "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        resp = requests.get(url, headers=headers, timeout=5, allow_redirects=True)
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding or "utf-8"
        html = resp.text
    except Exception:
        return result

    soup = BeautifulSoup(html, "html.parser")

    # scriptタグとstyleタグを除去
    for tag in soup(["script", "style"]):
        tag.decompose()

    text = soup.get_text(separator="\n")

    # 電話番号を抽出
    result["phone"] = _extract_phone(text)

    # 住所を抽出
    result["address"] = _extract_address(text)

    # 会社概要ページや問い合わせページも試行
    if not result["phone"] or not result["address"]:
        subpages = _find_company_info_links(soup, url)
        for sub_url in subpages[:3]:
            try:
                resp2 = requests.get(sub_url, headers=headers, timeout=5, allow_redirects=True)
                resp2.raise_for_status()
                resp2.encoding = resp2.apparent_encoding or "utf-8"
                soup2 = BeautifulSoup(resp2.text, "html.parser")
                for tag in soup2(["script", "style"]):
                    tag.decompose()
                text2 = soup2.get_text(separator="\n")

                if not result["phone"]:
                    result["phone"] = _extract_phone(text2)
                if not result["address"]:
                    result["address"] = _extract_address(text2)

                if result["phone"] and result["address"]:
                    break
            except Exception:
                continue

    return result


def _find_company_info_links(soup: BeautifulSoup, base_url: str) -> list[str]:
    """会社概要・お問い合わせなどのリンクを探す"""
    keywords = ["会社概要", "企業情報", "company", "about", "corporate", "お問い合わせ", "contact", "access", "アクセス"]
    links = []
    parsed = urlparse(base_url)
    base = f"{parsed.scheme}://{parsed.netloc}"

    for a in soup.find_all("a", href=True):
        href = a.get("href", "")
        text = a.get_text(strip=True).lower()
        href_lower = href.lower()

        if any(kw in text or kw in href_lower for kw in keywords):
            if href.startswith("http"):
                # 同一ドメインのみ
                if urlparse(href).netloc == parsed.netloc:
                    links.append(href)
            elif href.startswith("/"):
                links.append(base + href)

    return list(dict.fromkeys(links))  # 重複除去、順序保持


def build_search_query(keyword: str, industries: list[str], prefecture: str | None = None) -> str:
    """検索クエリを業種・地域条件から構築する"""
    parts = [keyword]

    if industries:
        industry_str = " OR ".join(industries)
        parts.append(f"({industry_str})")

    parts.append("会社 企業")

    if prefecture:
        parts.append(prefecture)

    return " ".join(parts)


def search_companies(
    api_key: str,
    search_engine_id: str,
    keyword: str,
    industries: list[str] | None = None,
    prefecture: str | None = None,
    num_results: int = 10,
    scrape_details: bool = True,
    progress_callback=None,
) -> list[dict]:
    """Google Custom Search APIで企業を検索する

    Args:
        progress_callback: (current, total) を受け取るコールバック関数（進捗表示用）

    Returns:
        list[dict]: 検索結果のリスト
    """
    service = build("customsearch", "v1", developerKey=api_key)

    query = build_search_query(keyword, industries or [], prefecture)

    results = []
    seen_domains = set()

    # Google Custom Search APIは1回のリクエストで最大10件
    fetched = 0
    start_index = 1

    while fetched < num_results:
        batch_size = min(10, num_results - fetched)

        try:
            response = (
                service.cse()
                .list(
                    q=query,
                    cx=search_engine_id,
                    num=batch_size,
                    start=start_index,
                    lr="lang_ja",
                    gl="jp",
                )
                .execute()
            )
        except HttpError as e:
            if e.resp.status == 429:
                raise RuntimeError("API利用制限に達しました。しばらく待ってから再試行してください。")
            elif e.resp.status == 403:
                raise RuntimeError("APIキーが無効か、権限がありません。設定を確認してください。")
            else:
                raise RuntimeError(f"API呼び出しエラー: {e.resp.status} - {e.reason}")

        items = response.get("items", [])
        if not items:
            break

        for item in items:
            url = item.get("link", "")
            domain = urlparse(url).netloc

            # 同一ドメインの重複を除外
            if domain in seen_domains:
                continue
            seen_domains.add(domain)

            company_name = _extract_company_name(item.get("title", ""))
            snippet = item.get("snippet", "").replace("\n", " ").strip()

            # スニペットから電話番号を抽出（第1段階）
            phone = _extract_phone(snippet)
            address = _extract_address(snippet)

            results.append(
                {
                    "company_name": company_name,
                    "url": url,
                    "domain": domain,
                    "snippet": snippet,
                    "phone": phone,
                    "address": address,
                    "industry": ", ".join(industries) if industries else "",
                    "prefecture": prefecture or "",
                    "searched_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                }
            )
            fetched += 1

            if fetched >= num_results:
                break

        start_index += batch_size

        # 検索結果の総数を超えたら終了
        total_results = int(response.get("searchInformation", {}).get("totalResults", "0"))
        if start_index > total_results:
            break

    # 第2段階: サイトスクレイピングで電話番号・住所を補完
    if scrape_details:
        for i, result in enumerate(results):
            if progress_callback:
                progress_callback(i + 1, len(results))

            if result["phone"] and result["address"]:
                continue

            details = scrape_company_details(result["url"])
            if not result["phone"] and details["phone"]:
                result["phone"] = details["phone"]
            if not result["address"] and details["address"]:
                result["address"] = details["address"]

    return results


def _extract_company_name(title: str) -> str:
    """検索結果のタイトルから企業名を推定抽出する"""
    # よくある区切り文字で分割して最初の部分を取得
    separators = [" | ", " - ", " – ", "｜", "／", " :: "]
    for sep in separators:
        if sep in title:
            title = title.split(sep)[0]
            break

    # 末尾の一般的なサフィックスを除去
    suffixes = ["公式サイト", "公式ホームページ", "ホームページ", "コーポレートサイト", "TOP", "トップ"]
    for suffix in suffixes:
        if title.endswith(suffix):
            title = title[: -len(suffix)].strip()

    return title.strip()
