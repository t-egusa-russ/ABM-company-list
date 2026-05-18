"""週次補助金レポート 実行エントリポイント

GitHub Actions（毎週月曜 12:00 JST）から呼び出される想定。
ローカル実行も可能。

環境変数:
    GOOGLE_API_KEY              Custom Search API キー
    GOOGLE_SEARCH_ENGINE_ID     Programmable Search Engine の ID
    SMTP_HOST / SMTP_PORT       メールサーバ（例: smtp.gmail.com / 587）
    SMTP_USER / SMTP_PASSWORD   SMTP認証情報（Gmailはアプリパスワード）
    SMTP_USE_SSL                "true" で SMTPS（465）を使用
    MAIL_FROM                   送信元アドレス
    MAIL_TO                     送信先（カンマ区切りで複数指定可）

任意:
    DRY_RUN=1                   メール送信をスキップしてログのみ
    RESULTS_PER_QUERY=5         1クエリあたりの取得件数（1〜10）
"""

from __future__ import annotations

import os
import sys
from datetime import datetime
from pathlib import Path

# プロジェクトルートを sys.path に追加（GitHub Actions からも使えるように）
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from dotenv import load_dotenv  # noqa: E402

from subsidy.collector import collect_subsidies  # noqa: E402
from subsidy.db import (  # noqa: E402
    init_db,
    upsert_subsidies,
    get_subsidies_since,
    get_all_subsidies,
)
from subsidy.report import build_excel_report, build_html_report  # noqa: E402
from subsidy.mailer import send_report_email  # noqa: E402


def _log(msg: str) -> None:
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


def main() -> int:
    load_dotenv()

    api_key = os.getenv("GOOGLE_API_KEY", "")
    engine_id = os.getenv("GOOGLE_SEARCH_ENGINE_ID", "")
    if not api_key or not engine_id:
        _log("ERROR: GOOGLE_API_KEY / GOOGLE_SEARCH_ENGINE_ID が未設定です")
        return 2

    results_per_query = int(os.getenv("RESULTS_PER_QUERY", "5"))
    dry_run = os.getenv("DRY_RUN", "").lower() in ("1", "true", "yes")

    _log("補助金情報の収集を開始")
    init_db()

    items = collect_subsidies(
        api_key=api_key,
        search_engine_id=engine_id,
        results_per_query=results_per_query,
        logger=_log,
    )
    _log(f"収集件数: {len(items)}")

    new_count, update_count = upsert_subsidies([i.to_dict() for i in items])
    _log(f"DB更新: 新規 {new_count} 件 / 既存更新 {update_count} 件")

    new_df = get_subsidies_since(days=7)
    all_df = get_all_subsidies()
    _log(f"今週分: {len(new_df)} 件 / 累計: {len(all_df)} 件")

    today = datetime.now()
    excel_bytes = build_excel_report(new_df if not new_df.empty else all_df)
    html_body = build_html_report(new_df=new_df, all_df=all_df, report_date=today)
    subject = f"週次 海外輸出 補助金レポート ({today:%Y-%m-%d}) - 新規 {len(new_df)} 件"
    filename = f"weekly_subsidy_report_{today:%Y%m%d}.xlsx"

    # Artifact用にファイル出力（GitHub Actions で参照しやすいように）
    out_dir = ROOT / "out"
    out_dir.mkdir(exist_ok=True)
    (out_dir / filename).write_bytes(excel_bytes)
    (out_dir / f"weekly_subsidy_report_{today:%Y%m%d}.html").write_text(
        html_body, encoding="utf-8"
    )
    _log(f"レポートを {out_dir} に書き出しました")

    if dry_run:
        _log("DRY_RUN=1 のためメール送信をスキップしました")
        return 0

    try:
        send_report_email(
            subject=subject,
            html_body=html_body,
            attachment_bytes=excel_bytes,
            attachment_filename=filename,
        )
        _log("メール送信に成功しました")
    except Exception as e:
        _log(f"ERROR: メール送信に失敗しました: {e}")
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
