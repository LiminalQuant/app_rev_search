from __future__ import annotations

import hashlib
from io import BytesIO
from typing import Any

import numpy as np
import pandas as pd
import plotly.express as px
import requests
import streamlit as st
from google_play_scraper import Sort, reviews_all


# =========================
# CONFIG
# =========================

DEFAULT_GOOGLE_APP_ID = "ru.mts.smartmed"
DEFAULT_APPLE_APP_ID = "6736902949"

STANDARD_COLUMNS = [
    "source",
    "store_country",
    "review_id",
    "user_name",
    "review_title",
    "review_description",
    "rating",
    "review_date",
    "developer_response",
    "developer_response_date",
    "thumbs_up",
    "app_version",
]


# =========================
# UTILS
# =========================

def empty_reviews_df() -> pd.DataFrame:
    return pd.DataFrame(columns=STANDARD_COLUMNS)


def stable_review_id(*parts: Any) -> str:
    """
    Детерминированный ID вместо uuid4.
    Если перезапустить парсер, ID останется тем же.
    Это важно для дедупликации.
    """
    raw = "|".join("" if part is None else str(part) for part in parts)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


def get_nested(data: dict, *keys: str) -> Any:
    current = data

    for key in keys:
        if not isinstance(current, dict):
            return None

        current = current.get(key)

    return current


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    for column in STANDARD_COLUMNS:
        if column not in df.columns:
            df[column] = None

    return df[STANDARD_COLUMNS]


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="reviews")

    return output.getvalue()


# =========================
# GOOGLE PLAY
# =========================

def fetch_google_play_reviews(
    app_id: str,
    lang: str = "ru",
    country: str = "ru",
) -> tuple[pd.DataFrame, str | None]:
    try:
        raw_reviews = reviews_all(
            app_id,
            sleep_milliseconds=100,
            lang=lang,
            country=country,
            sort=Sort.NEWEST,
        )

    except Exception as exc:
        return empty_reviews_df(), f"Google Play не отдал отзывы: {exc}"

    if not raw_reviews:
        return empty_reviews_df(), "Google Play: отзывы не найдены."

    df = pd.json_normalize(raw_reviews)

    df = df.rename(
        columns={
            "reviewId": "review_id",
            "userName": "user_name",
            "content": "review_description",
            "score": "rating",
            "at": "review_date",
            "replyContent": "developer_response",
            "repliedAt": "developer_response_date",
            "thumbsUpCount": "thumbs_up",
            "reviewCreatedVersion": "app_version",
        }
    )

    df["source"] = "Google Play"
    df["store_country"] = country.upper()
    df["review_title"] = None

    return normalize_columns(df), None


# =========================
# APP STORE RSS
# =========================

def parse_apple_review_entry(entry: dict, country: str) -> dict | None:
    """
    В RSS первый entry иногда может быть мета-информацией о приложении.
    Отзыв отличаем по наличию im:rating.
    """
    rating = get_nested(entry, "im:rating", "label")

    if rating is None:
        return None

    review_description = get_nested(entry, "content", "label")
    review_date = get_nested(entry, "updated", "label")
    user_name = get_nested(entry, "author", "name", "label")
    review_title = get_nested(entry, "title", "label")

    review_id = (
        get_nested(entry, "id", "label")
        or stable_review_id("app_store", country, user_name, review_date, review_title, review_description)
    )

    return {
        "source": "App Store",
        "store_country": country.upper(),
        "review_id": review_id,
        "user_name": user_name,
        "review_title": review_title,
        "review_description": review_description,
        "rating": rating,
        "review_date": review_date,
        "developer_response": None,
        "developer_response_date": None,
        "thumbs_up": None,
        "app_version": None,
    }


def fetch_app_store_rss_reviews(
    app_id: str,
    countries: tuple[str, ...] = ("ru", "us"),
    max_pages: int = 10,
    timeout: int = 15,
) -> tuple[pd.DataFrame, str | None]:
    rows: list[dict] = []
    warnings: list[str] = []

    headers = {
        "User-Agent": "Mozilla/5.0 reviews-dashboard/1.0"
    }

    for country in countries:
        before_country_count = len(rows)

        for page in range(1, max_pages + 1):
            url = (
                f"https://itunes.apple.com/{country}/rss/customerreviews/"
                f"page={page}/id={app_id}/sortby=mostrecent/json"
            )

            try:
                response = requests.get(url, headers=headers, timeout=timeout)
                response.raise_for_status()
                payload = response.json()

            except Exception as exc:
                warnings.append(f"App Store {country.upper()}, page {page}: {exc}")
                break

            entries = payload.get("feed", {}).get("entry", [])

            if isinstance(entries, dict):
                entries = [entries]

            page_rows = []

            for entry in entries:
                parsed = parse_apple_review_entry(entry, country)

                if parsed:
                    page_rows.append(parsed)

            if not page_rows:
                break

            rows.extend(page_rows)

        if len(rows) == before_country_count:
            warnings.append(f"App Store {country.upper()}: отзывы не получены.")

    if not rows:
        message = "App Store: отзывы не получены."
        if warnings:
            message += " " + " | ".join(warnings)

        return empty_reviews_df(), message

    df = pd.DataFrame(rows)
    df = df.drop_duplicates(subset=["review_id"])

    warning_message = " | ".join(warnings) if warnings else None

    return normalize_columns(df), warning_message


# =========================
# DATA PREP
# =========================

def prepare_reviews(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    df = df.copy()

    df["rating"] = pd.to_numeric(df["rating"], errors="coerce")

    df["review_date"] = (
        pd.to_datetime(df["review_date"], errors="coerce", utc=True)
        .dt.tz_convert(None)
    )

    df = df.dropna(subset=["review_date", "rating"])

    df["date"] = df["review_date"].dt.floor("D")
    df["year"] = df["date"].dt.year
    df["month"] = df["date"].dt.month
    df["year_month"] = df["date"].dt.to_period("M").astype(str)

    df["rating_group"] = np.where(
        df["rating"] < 4,
        "Отрицательный",
        "Положительный",
    )

    df["value"] = 1

    return df.sort_values("review_date", ascending=False)


@st.cache_data(ttl=60 * 60, show_spinner=False)
def load_reviews(
    google_app_id: str,
    google_lang: str,
    google_country: str,
    apple_app_id: str,
    apple_countries: tuple[str, ...],
) -> tuple[pd.DataFrame, list[str]]:
    google_df, google_warning = fetch_google_play_reviews(
        app_id=google_app_id,
        lang=google_lang,
        country=google_country,
    )

    apple_df, apple_warning = fetch_app_store_rss_reviews(
        app_id=apple_app_id,
        countries=apple_countries,
    )

    df = pd.concat([google_df, apple_df], ignore_index=True)
    df = prepare_reviews(df)

    warnings = [
        warning
        for warning in [google_warning, apple_warning]
        if warning
    ]

    return df, warnings


# =========================
# STREAMLIT APP
# =========================

def main() -> None:
    st.set_page_config(
        page_title="Отзывы SmartMed",
        layout="wide",
    )

    st.title("Отзывы SmartMed / Storemed")
    st.markdown("Дашборд по отзывам из Google Play и App Store.")

    with st.sidebar:
        st.header("Источники")

        google_app_id = st.text_input(
            "Google Play app_id",
            value=DEFAULT_GOOGLE_APP_ID,
        )

        google_lang = st.text_input(
            "Google Play lang",
            value="ru",
        ).strip().lower()

        google_country = st.text_input(
            "Google Play country",
            value="ru",
        ).strip().lower()

        apple_app_id = st.text_input(
            "App Store app_id",
            value=DEFAULT_APPLE_APP_ID,
        )

        apple_countries = st.multiselect(
            "App Store countries",
            options=["ru", "us", "de", "gb"],
            default=["ru", "us"],
        )

        if st.button("Обновить данные"):
            load_reviews.clear()

    with st.spinner("Загружаю отзывы..."):
        df, warnings = load_reviews(
            google_app_id=google_app_id,
            google_lang=google_lang,
            google_country=google_country,
            apple_app_id=apple_app_id,
            apple_countries=tuple(apple_countries),
        )

    for warning in warnings:
        st.warning(warning)

    if df.empty:
        st.error("Данных нет. Проверь app_id, страну магазина или доступность источника.")
        st.stop()

    with st.sidebar:
        st.header("Фильтры")

        years = sorted(df["year"].dropna().astype(int).unique(), reverse=True)
        months = sorted(df["month"].dropna().astype(int).unique())
        sources = sorted(df["source"].dropna().unique())
        countries = sorted(df["store_country"].dropna().unique())
        rating_groups = sorted(df["rating_group"].dropna().unique())

        default_year = [max(years)] if years else []

        selected_years = st.multiselect(
            "Год",
            options=years,
            default=default_year,
        )

        selected_months = st.multiselect(
            "Месяц",
            options=months,
            default=months,
        )

        selected_sources = st.multiselect(
            "Источник",
            options=sources,
            default=sources,
        )

        selected_countries = st.multiselect(
            "Страна магазина",
            options=countries,
            default=countries,
        )

        selected_rating_groups = st.multiselect(
            "Группа рейтинга",
            options=rating_groups,
            default=rating_groups,
        )

    mask = (
        df["year"].isin(selected_years)
        & df["month"].isin(selected_months)
        & df["source"].isin(selected_sources)
        & df["store_country"].isin(selected_countries)
        & df["rating_group"].isin(selected_rating_groups)
    )

    df_selection = df.loc[mask].copy()

    if df_selection.empty:
        st.warning("По выбранным фильтрам данных нет.")
        st.stop()

    # =========================
    # KPI
    # =========================

    average_rating = round(df_selection["rating"].mean(), 2)
    total_reviews = len(df_selection)
    negative_share = round((df_selection["rating"] < 4).mean() * 100, 1)
    positive_share = round((df_selection["rating"] >= 4).mean() * 100, 1)

    star_rating = "⭐" * int(round(average_rating, 0))

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Средний рейтинг", f"{average_rating} {star_rating}")
    col2.metric("Всего отзывов", f"{total_reviews}")
    col3.metric("Доля отрицательных", f"{negative_share}%")
    col4.metric("Доля положительных", f"{positive_share}%")

    st.markdown("---")

    # =========================
    # CHARTS
    # =========================

    source_rating = (
        df_selection
        .groupby(["source", "rating_group"], as_index=False)
        .agg(value=("value", "sum"))
    )

    fig_source = px.bar(
        source_rating,
        x="value",
        y="source",
        color="rating_group",
        orientation="h",
        barmode="stack",
        title="Оценки по источникам",
        color_discrete_map={
            "Отрицательный": "red",
            "Положительный": "green",
        },
    )

    st.plotly_chart(fig_source, use_container_width=True)

    monthly = (
        df_selection
        .groupby(["year_month", "rating_group"], as_index=False)
        .agg(value=("value", "sum"))
        .sort_values("year_month")
    )

    fig_month = px.bar(
        monthly,
        x="year_month",
        y="value",
        color="rating_group",
        barmode="stack",
        title="Отзывы по месяцам",
        color_discrete_map={
            "Отрицательный": "red",
            "Положительный": "green",
        },
    )

    st.plotly_chart(fig_month, use_container_width=True)

    daily = (
        df_selection
        .groupby(["date", "rating_group"], as_index=False)
        .agg(value=("value", "sum"))
        .sort_values("date")
    )

    fig_date = px.bar(
        daily,
        x="date",
        y="value",
        color="rating_group",
        barmode="stack",
        title="Отзывы по датам",
        color_discrete_map={
            "Отрицательный": "red",
            "Положительный": "green",
        },
    )

    st.plotly_chart(fig_date, use_container_width=True)

    st.markdown("---")

    # =========================
    # TABLE + DOWNLOAD
    # =========================

    visible_columns = [
        "source",
        "store_country",
        "review_date",
        "rating",
        "rating_group",
        "user_name",
        "review_title",
        "review_description",
        "developer_response",
        "thumbs_up",
        "app_version",
    ]

    st.subheader("Таблица отзывов")
    st.dataframe(
        df_selection[visible_columns],
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        label="Скачать Excel",
        data=to_excel_bytes(df_selection[visible_columns]),
        file_name="smartmed_reviews.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
