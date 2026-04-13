"""ABMアプローチ 企業リストアップアプリ - メインUI"""

import os
import streamlit as st
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()


def _get_env(key: str, default: str = "") -> str:
    """環境変数 → st.secrets → default の優先順で値を取得"""
    val = os.getenv(key, "")
    if val:
        return val
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


from search_engine import (
    search_companies,
    PREFECTURES,
    INDUSTRIES,
    PRODUCT_GENRES,
    MANUFACTURING_CATEGORIES,
)
from data_manager import (
    add_company,
    add_companies_bulk,
    get_all_companies,
    get_companies_filtered,
    update_company,
    delete_company,
    delete_companies,
    export_to_excel,
    EMPLOYEE_RANGES,
    OVERSEAS_OPTIONS,
)

# --- ページ設定 ---
st.set_page_config(
    page_title="テレアポ用 企業リストアップ",
    page_icon="📞",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- カスタムCSS ---
st.markdown(
    """
    <style>
    .main-header {
        font-size: 1.8rem;
        font-weight: 700;
        color: #2B5797;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 0.95rem;
        color: #666;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background: #f8f9fa;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
        border: 1px solid #e9ecef;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 2rem;
    }
    .phone-highlight {
        color: #d63384;
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ===== サイドバー =====
with st.sidebar:
    st.header("設定")

    # API設定（.envファイルから自動読み込み）
    with st.expander("API設定", expanded=False):
        api_key = st.text_input(
            "Google API Key",
            value=_get_env("GOOGLE_API_KEY"),
            type="password",
            key="api_key",
            help="Google Cloud ConsoleでCustom Search APIのキーを取得してください",
        )
        search_engine_id = st.text_input(
            "Search Engine ID",
            value=_get_env("GOOGLE_SEARCH_ENGINE_ID"),
            key="search_engine_id",
            help="Programmable Search Engineで作成したSearch Engine IDを入力",
        )

        if api_key and search_engine_id:
            st.success("API設定済み")
        else:
            st.info("APIキーを設定すると企業検索が利用できます")

    st.divider()

    # フィルター設定
    st.subheader("検索フィルター")

    industry_mode = st.radio(
        "業種カテゴリ",
        options=["製造業（消費財）", "全業種"],
        key="industry_mode",
        horizontal=True,
    )

    if industry_mode == "製造業（消費財）":
        selected_industries = st.multiselect(
            "製造業カテゴリを選択",
            options=MANUFACTURING_CATEGORIES,
            default=[],
            key="filter_mfg_categories",
        )
    else:
        selected_industries = st.multiselect(
            "業種を選択",
            options=INDUSTRIES,
            default=[],
            key="filter_industries",
        )

    custom_industry = st.text_input(
        "カスタム業種（任意）",
        placeholder="例: 農業, 宇宙産業",
        key="custom_industry",
    )
    if custom_industry:
        extra = [s.strip() for s in custom_industry.split(",") if s.strip()]
        selected_industries = selected_industries + extra

    selected_prefecture = st.selectbox(
        "地域",
        options=["指定なし"] + PREFECTURES,
        key="filter_prefecture",
    )
    if selected_prefecture == "指定なし":
        selected_prefecture = None

    st.divider()

    # 登録済みデータのフィルター
    st.subheader("リスト表示フィルター")
    list_keyword = st.text_input("キーワード検索", placeholder="会社名・概要で絞り込み", key="list_keyword")
    list_industry = st.selectbox(
        "業種で絞り込み",
        options=["すべて"] + INDUSTRIES + MANUFACTURING_CATEGORIES,
        key="list_industry_filter",
    )
    list_product_genre = st.selectbox(
        "商品ジャンルで絞り込み",
        options=["すべて"] + PRODUCT_GENRES,
        key="list_product_genre_filter",
    )
    list_employee_range = st.selectbox(
        "従業員数規模で絞り込み",
        options=["すべて"] + EMPLOYEE_RANGES,
        key="list_employee_range_filter",
    )
    list_has_overseas = st.selectbox(
        "海外支店の有無で絞り込み",
        options=["すべて"] + OVERSEAS_OPTIONS,
        key="list_overseas_filter",
    )


# ===== メインエリア =====
st.markdown('<p class="main-header">📞 テレアポ用 企業リストアップ</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="sub-header">Google検索APIで企業情報を自動収集し、電話番号付きのExcelリストを出力します</p>',
    unsafe_allow_html=True,
)

# サマリー指標
df_all = get_all_companies()
col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric("登録企業数", len(df_all))
with col2:
    phone_count = 0
    if not df_all.empty and "phone" in df_all.columns:
        phone_count = len(df_all[df_all["phone"].fillna("").str.strip() != ""])
    st.metric("電話番号あり", phone_count)
with col3:
    n_industries = df_all["industry"].nunique() if not df_all.empty else 0
    st.metric("業種数", n_industries)
with col4:
    n_prefectures = df_all["prefecture"].nunique() if not df_all.empty else 0
    st.metric("地域数", n_prefectures)
with col5:
    today_count = 0
    if not df_all.empty and "created_at" in df_all.columns:
        today = datetime.now().strftime("%Y-%m-%d")
        today_count = len(df_all[df_all["created_at"].str.startswith(today, na=False)])
    st.metric("本日追加", today_count)

st.divider()

# --- タブ構成 ---
tab_search, tab_list, tab_add = st.tabs(["企業検索", "企業リスト", "手動追加"])

# ===== タブ1: 企業検索 =====
with tab_search:
    st.subheader("企業を検索")

    if not api_key or not search_engine_id:
        st.warning("サイドバーの「API設定」からGoogle API KeyとSearch Engine IDを入力してください。")
    else:
        search_col1, search_col2 = st.columns([3, 1])
        with search_col1:
            search_keyword = st.text_input(
                "検索キーワード",
                placeholder="例: 化粧品 製造 OEM",
                key="search_keyword",
            )
        with search_col2:
            num_results = st.selectbox("取得件数", options=[10, 20, 30, 50], key="num_results")

        scrape_details = st.checkbox(
            "企業サイトから電話番号・住所を自動取得する（時間がかかります）",
            value=True,
            key="scrape_details",
        )

        # 現在の検索条件を表示
        conditions = []
        if selected_industries:
            conditions.append(f"業種: {', '.join(selected_industries)}")
        if selected_prefecture:
            conditions.append(f"地域: {selected_prefecture}")
        if conditions:
            st.caption(f"検索条件: {' / '.join(conditions)}")

        if st.button("検索実行", type="primary", use_container_width=True):
            if not search_keyword:
                st.error("検索キーワードを入力してください")
            else:
                progress_bar = st.progress(0, text="企業を検索中...")
                status_text = st.empty()

                def update_progress(current, total):
                    progress_bar.progress(
                        current / total,
                        text=f"企業サイトをスクレイピング中... ({current}/{total})",
                    )

                try:
                    results = search_companies(
                        api_key=api_key,
                        search_engine_id=search_engine_id,
                        keyword=search_keyword,
                        industries=selected_industries if selected_industries else None,
                        prefecture=selected_prefecture,
                        num_results=num_results,
                        scrape_details=scrape_details,
                        progress_callback=update_progress if scrape_details else None,
                    )
                    st.session_state["search_results"] = results
                    progress_bar.empty()
                    status_text.empty()
                except RuntimeError as e:
                    progress_bar.empty()
                    st.error(str(e))
                except Exception as e:
                    progress_bar.empty()
                    st.error(f"検索エラー: {e}")

        # 検索結果の表示
        if "search_results" in st.session_state and st.session_state["search_results"]:
            results = st.session_state["search_results"]
            phone_found = sum(1 for r in results if r.get("phone"))
            st.success(f"{len(results)} 件の企業が見つかりました（電話番号取得: {phone_found} 件）")

            # 一括保存ボタン
            if st.button("検索結果をすべて保存", type="secondary"):
                added, skipped, messages = add_companies_bulk(results)
                if added > 0:
                    st.success(f"{added} 件を保存しました")
                if skipped > 0:
                    st.warning(f"{skipped} 件はスキップ（重複）")
                    for msg in messages[:5]:
                        st.caption(msg)

            # 各検索結果を表示
            for i, result in enumerate(results):
                phone_label = f" | 📞 {result['phone']}" if result.get("phone") else " | 📞 未取得"
                with st.expander(f"{result['company_name']}{phone_label}", expanded=False):
                    r_col1, r_col2 = st.columns(2)
                    with r_col1:
                        st.markdown(f"**会社名:** {result['company_name']}")
                        if result.get("phone"):
                            st.markdown(f"**電話番号:** :red[{result['phone']}]")
                        else:
                            st.markdown("**電話番号:** 未取得")
                        if result.get("address"):
                            st.markdown(f"**住所:** {result['address']}")
                        st.markdown(f"**URL:** {result['url']}")
                    with r_col2:
                        st.markdown(f"**概要:** {result['snippet']}")
                        if result.get("industry"):
                            st.markdown(f"**業種:** {result['industry']}")
                        if result.get("prefecture"):
                            st.markdown(f"**地域:** {result['prefecture']}")

                    if st.button("この企業を保存", key=f"save_{i}"):
                        success, msg = add_company(result)
                        if success:
                            st.success(msg)
                        else:
                            st.warning(msg)


# ===== タブ2: 企業リスト =====
with tab_list:
    st.subheader("登録済み企業リスト")

    # フィルター適用
    filter_industry = None if list_industry == "すべて" else list_industry
    filter_keyword = list_keyword if list_keyword else None
    filter_product_genre = None if list_product_genre == "すべて" else list_product_genre
    filter_employee_range = None if list_employee_range == "すべて" else list_employee_range
    filter_has_overseas = None if list_has_overseas == "すべて" else list_has_overseas

    df = get_companies_filtered(
        industry=filter_industry,
        keyword=filter_keyword,
        product_genre=filter_product_genre,
        employee_range=filter_employee_range,
        has_overseas=filter_has_overseas,
    )

    if df.empty:
        st.info("企業が登録されていません。「企業検索」または「手動追加」から企業を追加してください。")
    else:
        # 操作ボタン行
        action_col1, action_col2, action_col3 = st.columns([1, 1, 3])
        with action_col1:
            excel_data = export_to_excel(df)
            filename = f"テレアポリスト_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            st.download_button(
                "Excelダウンロード",
                data=excel_data,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        with action_col2:
            st.caption(f"{len(df)} 件表示中")

        # テーブル表示
        display_cols = ["company_name", "phone", "address", "industry", "product_genre",
                        "prefecture", "url", "snippet", "employee_count",
                        "employee_range", "revenue", "has_overseas", "notes"]
        display_cols = [c for c in display_cols if c in df.columns]
        display_df = df[display_cols].copy()
        col_labels = {
            "company_name": "会社名", "phone": "電話番号", "address": "住所",
            "industry": "業種", "product_genre": "商品ジャンル",
            "prefecture": "所在地", "url": "URL", "snippet": "概要",
            "employee_count": "従業員数", "employee_range": "従業員数規模",
            "revenue": "売上規模", "has_overseas": "海外支店", "notes": "備考",
        }
        display_df.columns = [col_labels.get(c, c) for c in display_cols]

        st.dataframe(
            display_df,
            use_container_width=True,
            height=400,
            column_config={
                "URL": st.column_config.LinkColumn("URL", display_text="リンク"),
                "概要": st.column_config.TextColumn("概要", width="large"),
                "電話番号": st.column_config.TextColumn("電話番号", width="medium"),
                "住所": st.column_config.TextColumn("住所", width="large"),
            },
        )

        # 個別企業の操作
        st.divider()
        st.subheader("企業の編集・削除")

        company_options = {
            f"{row['company_name']} (ID: {row['id']})": row["id"]
            for _, row in df.iterrows()
        }

        if company_options:
            selected_company_label = st.selectbox(
                "企業を選択",
                options=list(company_options.keys()),
                key="edit_company_select",
            )
            selected_id = company_options[selected_company_label]
            selected_row = df[df["id"] == selected_id].iloc[0]

            with st.expander("編集", expanded=False):
                edit_col1, edit_col2 = st.columns(2)
                with edit_col1:
                    edit_name = st.text_input("会社名", value=selected_row.get("company_name", ""), key="edit_name")
                    edit_phone = st.text_input("電話番号", value=selected_row.get("phone", ""), key="edit_phone")
                    edit_address = st.text_input("住所", value=selected_row.get("address", ""), key="edit_address")
                    edit_industry = st.text_input("業種", value=selected_row.get("industry", ""), key="edit_industry")
                    edit_product_genre = st.selectbox(
                        "商品ジャンル",
                        options=[""] + PRODUCT_GENRES,
                        index=(PRODUCT_GENRES.index(selected_row.get("product_genre", "")) + 1)
                            if selected_row.get("product_genre", "") in PRODUCT_GENRES else 0,
                        key="edit_product_genre",
                    )
                    edit_prefecture = st.text_input("所在地", value=selected_row.get("prefecture", ""), key="edit_pref")
                    edit_url = st.text_input("URL", value=selected_row.get("url", ""), key="edit_url")
                with edit_col2:
                    edit_employee = st.text_input("従業員数", value=selected_row.get("employee_count", ""), key="edit_emp")
                    edit_employee_range = st.selectbox(
                        "従業員数規模",
                        options=[""] + EMPLOYEE_RANGES,
                        index=(EMPLOYEE_RANGES.index(selected_row.get("employee_range", "")) + 1)
                            if selected_row.get("employee_range", "") in EMPLOYEE_RANGES else 0,
                        key="edit_emp_range",
                    )
                    edit_revenue = st.text_input("売上規模", value=selected_row.get("revenue", ""), key="edit_rev")
                    edit_has_overseas = st.selectbox(
                        "海外支店の有無",
                        options=[""] + OVERSEAS_OPTIONS,
                        index=(OVERSEAS_OPTIONS.index(selected_row.get("has_overseas", "")) + 1)
                            if selected_row.get("has_overseas", "") in OVERSEAS_OPTIONS else 0,
                        key="edit_overseas",
                    )
                    edit_notes = st.text_area("備考", value=selected_row.get("notes", ""), key="edit_notes")

                if st.button("更新", key="update_btn"):
                    update_company(selected_id, {
                        "company_name": edit_name,
                        "phone": edit_phone,
                        "address": edit_address,
                        "industry": edit_industry,
                        "product_genre": edit_product_genre,
                        "prefecture": edit_prefecture,
                        "url": edit_url,
                        "employee_count": edit_employee,
                        "employee_range": edit_employee_range,
                        "revenue": edit_revenue,
                        "has_overseas": edit_has_overseas,
                        "notes": edit_notes,
                    })
                    st.success("更新しました")
                    st.rerun()

            if st.button("この企業を削除", type="secondary", key="delete_btn"):
                delete_company(selected_id)
                st.success("削除しました")
                st.rerun()


# ===== タブ3: 手動追加 =====
with tab_add:
    st.subheader("企業を手動で追加")

    with st.form("manual_add_form"):
        add_col1, add_col2 = st.columns(2)
        with add_col1:
            m_name = st.text_input("会社名 *", key="m_name")
            m_phone = st.text_input("電話番号", key="m_phone", placeholder="例: 03-1234-5678")
            m_address = st.text_input("住所", key="m_address", placeholder="例: 東京都渋谷区...")
            m_industry = st.selectbox(
                "業種",
                options=[""] + INDUSTRIES + MANUFACTURING_CATEGORIES,
                key="m_industry",
            )
            m_product_genre = st.selectbox("商品ジャンル", options=[""] + PRODUCT_GENRES, key="m_product_genre")
            m_prefecture = st.selectbox("所在地", options=[""] + PREFECTURES, key="m_prefecture")
            m_url = st.text_input("URL", key="m_url", placeholder="https://example.com")
        with add_col2:
            m_employee = st.text_input("従業員数", key="m_employee", placeholder="例: 100名")
            m_employee_range = st.selectbox("従業員数規模", options=[""] + EMPLOYEE_RANGES, key="m_employee_range")
            m_revenue = st.text_input("売上規模", key="m_revenue", placeholder="例: 10億円")
            m_has_overseas = st.selectbox("海外支店の有無", options=[""] + OVERSEAS_OPTIONS, key="m_overseas")
            m_snippet = st.text_area("概要", key="m_snippet", placeholder="企業の概要を入力")
            m_notes = st.text_area("備考", key="m_notes", placeholder="営業メモなど")

        submitted = st.form_submit_button("追加", type="primary", use_container_width=True)

        if submitted:
            if not m_name:
                st.error("会社名は必須です")
            else:
                success, msg = add_company({
                    "company_name": m_name,
                    "phone": m_phone,
                    "address": m_address,
                    "url": m_url,
                    "domain": "",
                    "snippet": m_snippet,
                    "industry": m_industry,
                    "product_genre": m_product_genre,
                    "prefecture": m_prefecture,
                    "employee_count": m_employee,
                    "employee_range": m_employee_range,
                    "revenue": m_revenue,
                    "has_overseas": m_has_overseas,
                    "notes": m_notes,
                    "searched_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                })
                if success:
                    st.success(f"{m_name} を登録しました")
                else:
                    st.warning(msg)
