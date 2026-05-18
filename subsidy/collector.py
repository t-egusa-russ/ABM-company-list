"""補助金情報の収集ロジック

Google Custom Search API を使い、既定ソース（site: 絞り込み）と
汎用キーワード検索の両方を実行し、重複を除去して返す。
"""

from __future__ import annotations

import re
import time
from dataclasses import dataclass, asdict
from datetime import datetime
from typing import Optional
from urllib.parse import urlparse

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from .sources import PREDEFINED_SOURCES, EXPORT_KEYWORDS, SUBSIDY_TERMS


@dataclass
class SubsidyItem:
    title: str
    summary: str
    source_name: str
    source_category: str  # 行政 / 金融機関 / 自治体 / 公的機関 / 営利団体
    url: str
    domain: str
    matched_keyword: str
    collected_at: str

    def to_dict(self) -> dict:
        return asdict(self)


# 海外輸出と関係しない結果を弾く除外語
_EXCLUDE_KEYWORDS = (
    "求人", "採用情報", "リクルート",
)


def _is_relevant(title: str, snippet: str) -> bool:
    """タイトル＋スニペットが補助金として妥当か簡易判定"""
    text = f"{title} {snippet}"
    if any(ng in text for ng in _EXCLUDE_KEYWORDS):
        return False
    has_subsidy = any(term in text for term in SUBSIDY_TERMS)
    has_export = any(kw in text for kw in EXPORT_KEYWORDS) or "海外" in text or "輸出" in text
    return has_subsidy and has_export


def _normalize_title(title: str) -> str:
    """タイトルから区切り以降のサイト名等を除く"""
    for sep in (" | ", " - ", " – ", "｜", "／", " :: "):
        if sep in title:
            title = title.split(sep)[0]
            break
    return title.strip()


def _execute_search(
    service,
    query: str,
    search_engine_id: str,
    num: int = 10,
) -> list[dict]:
    """1回のCustom Search呼び出し（最大10件）。エラーは握りつぶさず例外で上げる。"""
    try:
        resp = (
            service.cse()
            .list(
                q=query,
                cx=search_engine_id,
                num=num,
                lr="lang_ja",
                gl="jp",
            )
            .execute()
        )
    except HttpError as e:
        status = e.resp.status if hasattr(e, "resp") else "unknown"
        if status == 429:
            raise RuntimeError("Custom Search API 利用上限に達しました")
        if status == 403:
            raise RuntimeError("Custom Search API のキーまたは権限が無効です")
        raise RuntimeError(f"Custom Search API エラー ({status}): {e}")
    return resp.get("items", [])


def collect_subsidies(
    api_key: str,
    search_engine_id: str,
    results_per_query: int = 5,
    sleep_seconds: float = 0.3,
    logger=None,
) -> list[SubsidyItem]:
    """海外輸出関連の補助金情報を収集する

    Args:
        api_key: Google Custom Search API キー
        search_engine_id: 検索エンジンID（Programmable Search Engine）
        results_per_query: 1検索クエリ当たりの取得件数（最大10）
        sleep_seconds: API呼び出し間のスリープ秒数（レート制限対策）
        logger: callable(str) でログを受け取る（省略可）

    Returns:
        list[SubsidyItem]
    """
    if not api_key or not search_engine_id:
        raise RuntimeError("GOOGLE_API_KEY と GOOGLE_SEARCH_ENGINE_ID を設定してください")

    def log(msg: str):
        if logger:
            logger(msg)

    service = build("customsearch", "v1", developerKey=api_key)
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    items: dict[str, SubsidyItem] = {}  # URL をキーに重複排除

    # --- 1) 既定ソースごとに site: 絞り込み検索 ---
    for source in PREDEFINED_SOURCES:
        # 「海外展開 補助金 site:domain」のような形に集約（クエリ数を抑制）
        query = (
            f"(海外展開 OR 海外輸出 OR 輸出 OR 越境EC OR 海外販路) "
            f"(補助金 OR 助成金 OR 支援事業) "
            f"site:{source['domain']}"
        )
        log(f"[既定ソース] {source['name']} ({source['domain']}) を検索中...")
        try:
            results = _execute_search(service, query, search_engine_id, num=results_per_query)
        except RuntimeError as e:
            log(f"  ! 失敗: {e}")
            continue
        except Exception as e:  # noqa: BLE001 — 個別ソースの失敗で全体を止めない
            log(f"  ! 予期せぬエラー: {e}")
            continue

        for item in results:
            url = item.get("link", "")
            if not url or url in items:
                continue
            title = _normalize_title(item.get("title", ""))
            snippet = item.get("snippet", "").replace("\n", " ").strip()
            if not _is_relevant(title, snippet):
                continue
            items[url] = SubsidyItem(
                title=title,
                summary=snippet,
                source_name=source["name"],
                source_category=source["category"],
                url=url,
                domain=urlparse(url).netloc,
                matched_keyword=f"site:{source['domain']}",
                collected_at=now,
            )

        time.sleep(sleep_seconds)

    # --- 2) 汎用キーワード検索（ソース横断） ---
    for kw in EXPORT_KEYWORDS:
        for term in SUBSIDY_TERMS:
            query = f"{kw} {term}"
            log(f"[キーワード] '{query}' を検索中...")
            try:
                results = _execute_search(service, query, search_engine_id, num=results_per_query)
            except RuntimeError as e:
                log(f"  ! 失敗: {e}")
                continue
            except Exception as e:  # noqa: BLE001
                log(f"  ! 予期せぬエラー: {e}")
                continue

            for item in results:
                url = item.get("link", "")
                if not url or url in items:
                    continue
                title = _normalize_title(item.get("title", ""))
                snippet = item.get("snippet", "").replace("\n", " ").strip()
                if not _is_relevant(title, snippet):
                    continue

                domain = urlparse(url).netloc
                source_name, source_category = _guess_source(domain)
                items[url] = SubsidyItem(
                    title=title,
                    summary=snippet,
                    source_name=source_name,
                    source_category=source_category,
                    url=url,
                    domain=domain,
                    matched_keyword=f"{kw} {term}",
                    collected_at=now,
                )

            time.sleep(sleep_seconds)

    log(f"収集完了: {len(items)} 件（重複排除後）")
    return list(items.values())


def _guess_source(domain: str) -> tuple[str, str]:
    """ドメインから既定ソース定義に該当する機関名・カテゴリを推定する"""
    for s in PREDEFINED_SOURCES:
        if domain.endswith(s["domain"]):
            return s["name"], s["category"]
    # 既定リスト外でも、TLDからおおまかな分類を当てる
    if domain.endswith(".go.jp"):
        return domain, "行政"
    if domain.endswith(".lg.jp"):
        return domain, "自治体"
    if domain.endswith(".or.jp"):
        return domain, "公的機関"
    if domain.endswith(".ac.jp"):
        return domain, "公的機関"
    return domain, "営利団体"
