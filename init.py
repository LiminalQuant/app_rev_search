from __future__ import annotations

import hashlib
import re
from io import BytesIO
from typing import Any

import numpy as np
import pandas as pd
import plotly.express as px
import requests
import streamlit as st
from google_play_scraper import Sort, reviews_all


# ============================================================
# CONFIG
# ============================================================

DEFAULT_GOOGLE_APP_ID = "ru.mts.smartmed"
DEFAULT_APPLE_APP_ID = "6736902949"

PROBLEM_RATING_MAX = 3

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

PROBLEM_PATTERNS = {
    "Авторизация / вход": [
        "войти",
        "вход",
        "авторизац",
        "логин",
        "пароль",
        "код",
        "смс",
        "sms",
        "не приходит код",
        "не могу войти",
        "не заходит",
    ],
    "Запись к врачу": [
        "запис",
        "запись",
        "прием",
        "приём",
        "слот",
        "расписан",
        "нет времени",
        "нет записи",
        "врач не отображ",
        "не могу записаться",
    ],
    "Отмена / перенос": [
        "отмен",
        "перенос",
        "перезапис",
        "сдвинули",
        "пропала запись",
        "исчезла запись",
    ],
    "Оплата": [
        "оплат",
        "платеж",
        "платёж",
        "карта",
        "деньги",
        "списали",
        "чек",
        "не проходит оплата",
        "ошибка оплаты",
    ],
    "Документы / результаты": [
        "анализ",
        "результат",
        "заключен",
        "заключение",
        "документ",
        "pdf",
        "справк",
        "не отображаются результаты",
        "не загрузились",
    ],
    "ДМС / страховая": [
        "дмс",
        "страхов",
        "полис",
        "согласован",
        "гарантийное письмо",
        "страховая",
    ],
    "Стабильность приложения": [
        "не работает",
        "сломалось",
        "ошибка",
        "баг",
        "вылет",
        "вылетает",
        "завис",
        "тормоз",
        "тормозит",
        "белый экран",
        "краш",
        "не открывается",
    ],
    "Уведомления": [
        "уведомлен",
        "push",
        "пуш",
        "напоминан",
        "не пришло уведомление",
        "не пришла смс",
    ],
    "Поддержка": [
        "поддержка",
        "оператор",
        "чат",
        "не отвеч",
        "обратная связь",
        "никто не отвечает",
        "дозвониться",
    ],
    "Личный кабинет / профиль": [
        "личный кабинет",
        "профиль",
        "данные",
        "пациент",
        "ребен",
        "ребён",
        "карта пациента",
    ],
    "UX / навигация": [
        "неудоб",
        "непонят",
        "найти",
        "интерфейс",
        "меню",
        "поиск",
        "куда нажать",
    ],
    "Клиника / сервис": [
        "администратор",
        "регистратура",
        "клиника",
        "врач",
        "медсестра",
        "очередь",
        "ожидание",
    ],
    "Обновление приложения": [
        "после обновления",
        "обновление",
        "обновили",
        "новая версия",
        "раньше работало",
    ],
}


# ============================================================
# COMMON UTILS
# ============================================================

def empty_reviews_df() -> pd.DataFrame:
    return pd.DataFrame(columns=STANDARD_COLUMNS)


def stable_review_id(*parts: Any) -> str:
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


def safe_text(value: Any) -> str:
    if pd.isna(value):
        return ""

    return str(value)


def to_excel_report(
    filtered_df: pd.DataFrame,
    problem_df: pd.DataFrame,
    topic_summary: pd.DataFrame,
    weekly_summary: pd.DataFrame,
    anomaly_days: pd.DataFrame,
    source_summary: pd.DataFrame,
) -> bytes:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        filtered_df.to_excel(writer, index=False, sheet_name="reviews_filtered")
        problem_df.to_excel(writer, index=False, sheet_name="problem_reviews")
        topic_summary.to_excel(writer, index=False, sheet_name="problem_topics")
        weekly_summary.to_excel(writer, index=False, sheet_name="weekly_summary")
        anomaly_days.to_excel(writer, index=False, sheet_name="anomaly_days")
        source_summary.to_excel(writer, index=False, sheet_name="source_summary")

    return output.getvalue()


# ============================================================
# GOOGLE PLAY
# ============================================================

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

    df["source"] = "Android / Google Play"
    df["store_country"] = country.upper()
    df["review_title"] = None

    return normalize_columns(df), None


# ============================================================
# APP STORE RSS
# ============================================================

def parse_apple_review_entry(entry: dict, country: str) -> dict | None:
    rating = get_nested(entry, "im:rating", "label")

    if rating is None:
        return None

    review_description = get_nested(entry, "content", "label")
    review_date = get_nested(entry, "updated", "label")
    user_name = get_nested(entry, "author", "name", "label")
    review_title = get_nested(entry, "title", "label")

    review_id = (
        get_nested(entry, "id", "label")
        or stable_review_id(
            "app_store",
            country,
            user_name,
            review_date,
            review_title,
            review_description,
        )
    )

    return {
        "source": "iOS / App Store",
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
        "User-Agent": "Mozilla/5.0 app-reviews-monitor/1.0"
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


# ============================================================
# ENRICHMENT / PATTERN ANALYTICS
# ============================================================

def detect_problem_topics(text: str, rating: float) -> list[str]:
    if pd.isna(rating) or rating > PROBLEM_RATING_MAX:
        return []

    text = safe_text(text).lower()
    matched_topics = []

    for topic, patterns in PROBLEM_PATTERNS.items():
        for pattern in patterns:
            if pattern.lower() in text:
                matched_topics.append(topic)
                break

    if not matched_topics:
        matched_topics.append("Без классификации")

    return matched_topics


def detect_matched_patterns(text: str, rating: float) -> str:
    if pd.isna(rating) or rating > PROBLEM_RATING_MAX:
        return ""

    text = safe_text(text).lower()
    matched = []

    for topic, patterns in PROBLEM_PATTERNS.items():
        for pattern in patterns:
            if pattern.lower() in text:
                matched.append(pattern)

    return ", ".join(sorted(set(matched)))


def prepare_reviews(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    df = df.copy()

    df["rating"] = pd.to_numeric(df["rating"], errors="coerce")

    df["review_date"] = (
        pd.to_datetime(df["review_date"], errors="coerce", utc=True)
        .dt.tz_convert(None)
    )

    df["developer_response_date"] = (
        pd.to_datetime(df["developer_response_date"], errors="coerce", utc=True)
        .dt.tz_convert(None)
    )

    df = df.dropna(subset=["review_date", "rating"])

    df["date"] = df["review_date"].dt.floor("D")
    df["year"] = df["date"].dt.year
    df["month"] = df["date"].dt.month
    df["year_month"] = df["date"].dt.to_period("M").astype(str)
    df["week_start"] = df["date"].dt.to_period("W").apply(lambda period: period.start_time)

    df["review_text"] = (
        df["review_title"].fillna("").astype(str)
        + " "
        + df["review_description"].fillna("").astype(str)
    ).str.strip()

    df["text_len"] = df["review_text"].str.len()
    df["has_text"] = df["text_len"] >= 20

    df["is_problem"] = df["rating"] <= PROBLEM_RATING_MAX
    df["is_hard_problem"] = df["rating"] <= 2
    df["is_positive"] = df["rating"] >= 4
    df["is_empty_positive"] = (df["rating"] >= 4) & (df["text_len"] < 20)
    df["is_rich_problem"] = (df["rating"] <= 2) & (df["text_len"] >= 100)

    df["rating_group"] = np.select(
        [
            df["rating"] <= 2,
            df["rating"] == 3,
            df["rating"] >= 4,
        ],
        [
            "1–2★ жёсткий негатив",
            "3★ умеренная проблема",
            "4–5★ позитив / нейтрально",
        ],
        default="Без оценки",
    )

    df["problem_topics"] = df.apply(
        lambda row: detect_problem_topics(row["review_text"], row["rating"]),
        axis=1,
    )

    df["problem_topics_str"] = df["problem_topics"].apply(
        lambda items: ", ".join(items) if items else ""
    )

    df["matched_patterns"] = df.apply(
        lambda row: detect_matched_patterns(row["review_text"], row["rating"]),
        axis=1,
    )

    df["has_developer_response"] = df["developer_response"].notna()

    df["response_days"] = (
        df["developer_response_date"] - df["review_date"]
    ).dt.days

    now = df["date"].max()

    df["days_from_last_review"] = (now - df["date"]).dt.days

    df["recency_weight"] = np.select(
        [
            df["days_from_last_review"] <= 30,
            df["days_from_last_review"] <= 90,
        ],
        [
            1.30,
            1.10,
        ],
        default=1.00,
    )

    df["problem_weight"] = np.select(
        [
            df["rating"] == 1,
            df["rating"] == 2,
            df["rating"] == 3,
        ],
        [
            3.0,
            2.0,
            1.0,
        ],
        default=0.0,
    )

    df["problem_score"] = (
        df["problem_weight"]
        * df["recency_weight"]
        + df["is_rich_problem"].astype(int) * 0.5
    )

    df["value"] = 1

    return df.sort_values("review_date", ascending=False)


def build_topic_summary(df: pd.DataFrame) -> pd.DataFrame:
    problem_df = df[df["is_problem"]].copy()

    if problem_df.empty:
        return pd.DataFrame()

    exploded = problem_df.explode("problem_topics")
    exploded = exploded.rename(columns={"problem_topics": "problem_topic"})

    summary = (
        exploded
        .groupby("problem_topic", as_index=False)
        .agg(
            problem_reviews=("review_id", "count"),
            hard_problem_reviews=("is_hard_problem", "sum"),
            rich_problem_reviews=("is_rich_problem", "sum"),
            avg_rating=("rating", "mean"),
            score=("problem_score", "sum"),
            last_review_date=("date", "max"),
            unique_days=("date", "nunique"),
        )
    )

    total_problem_reviews = len(problem_df)

    summary["problem_share"] = summary["problem_reviews"] / total_problem_reviews
    summary["hard_share"] = summary["hard_problem_reviews"] / summary["problem_reviews"]
    summary["rich_share"] = summary["rich_problem_reviews"] / summary["problem_reviews"]
    summary["avg_rating"] = summary["avg_rating"].round(2)
    summary["score"] = summary["score"].round(2)

    q75 = summary["score"].quantile(0.75)
    q50 = summary["score"].quantile(0.50)

    summary["priority"] = np.select(
        [
            summary["score"] >= q75,
            summary["score"] >= q50,
        ],
        [
            "Высокий",
            "Средний",
        ],
        default="Низкий",
    )

    return summary.sort_values(
        ["priority", "score", "problem_reviews"],
        ascending=[True, False, False],
    )


def build_weekly_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    weekly = (
        df
        .groupby(["week_start", "source"], as_index=False)
        .agg(
            reviews=("review_id", "count"),
            avg_rating=("rating", "mean"),
            problem_reviews=("is_problem", "sum"),
            hard_problem_reviews=("is_hard_problem", "sum"),
            rich_problem_reviews=("is_rich_problem", "sum"),
            empty_positive_reviews=("is_empty_positive", "sum"),
        )
    )

    weekly["problem_share"] = weekly["problem_reviews"] / weekly["reviews"]
    weekly["hard_problem_share"] = weekly["hard_problem_reviews"] / weekly["reviews"]
    weekly["empty_positive_share"] = weekly["empty_positive_reviews"] / weekly["reviews"]
    weekly["avg_rating"] = weekly["avg_rating"].round(2)

    return weekly.sort_values("week_start")


def build_source_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    summary = (
        df
        .groupby("source", as_index=False)
        .agg(
            reviews=("review_id", "count"),
            avg_rating=("rating", "mean"),
            problem_reviews=("is_problem", "sum"),
            hard_problem_reviews=("is_hard_problem", "sum"),
            rich_problem_reviews=("is_rich_problem", "sum"),
            text_reviews=("has_text", "sum"),
            developer_responses=("has_developer_response", "sum"),
        )
    )

    summary["problem_share"] = summary["problem_reviews"] / summary["reviews"]
    summary["hard_problem_share"] = summary["hard_problem_reviews"] / summary["reviews"]
    summary["text_share"] = summary["text_reviews"] / summary["reviews"]
    summary["developer_response_share"] = summary["developer_responses"] / summary["reviews"]
    summary["avg_rating"] = summary["avg_rating"].round(2)

    return summary.sort_values("reviews", ascending=False)


def build_anomaly_days(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    daily = (
        df
        .groupby("date", as_index=False)
        .agg(
            reviews=("review_id", "count"),
            avg_rating=("rating", "mean"),
            problem_reviews=("is_problem", "sum"),
            hard_problem_reviews=("is_hard_problem", "sum"),
            positive_reviews=("is_positive", "sum"),
            empty_positive_reviews=("is_empty_positive", "sum"),
        )
        .sort_values("date")
    )

    daily["problem_share"] = daily["problem_reviews"] / daily["reviews"]
    daily["empty_positive_share"] = daily["empty_positive_reviews"] / daily["reviews"]
    daily["avg_rating"] = daily["avg_rating"].round(2)

    daily["reviews_roll_mean_14"] = (
        daily["reviews"]
        .rolling(14, min_periods=5)
        .mean()
    )

    daily["reviews_roll_std_14"] = (
        daily["reviews"]
        .rolling(14, min_periods=5)
        .std()
    )

    daily["activity_spike"] = (
        daily["reviews"]
        > daily["reviews_roll_mean_14"] + 2 * daily["reviews_roll_std_14"]
    )

    daily["problem_spike"] = (
        (daily["problem_share"] >= daily["problem_share"].quantile(0.90))
        & (daily["problem_reviews"] >= 2)
    )

    daily["positive_concentration"] = (
        (daily["empty_positive_share"] >= 0.70)
        & (daily["reviews"] >= daily["reviews"].quantile(0.75))
    )

    daily["requires_check"] = (
        daily["activity_spike"]
        | daily["problem_spike"]
        | daily["positive_concentration"]
    )

    return daily[daily["requires_check"]].sort_values("date", ascending=False)


# ============================================================
# LOAD DATA
# ============================================================

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

    frames = [
        frame
        for frame in [google_df, apple_df]
        if frame is not None and not frame.empty
    ]

    if not frames:
        df = empty_reviews_df()
    else:
        df = pd.concat(frames, ignore_index=True)

    df = prepare_reviews(df)

    warnings = [
        warning
        for warning in [google_warning, apple_warning]
        if warning
    ]

    return df, warnings


# ============================================================
# UI HELPERS
# ============================================================

def format_percent(value: float) -> str:
    if pd.isna(value):
        return "—"

    return f"{value * 100:.1f}%"


def render_review_cards(df: pd.DataFrame, limit: int = 20) -> None:
    preview = df.head(limit)

    for _, row in preview.iterrows():
        title = safe_text(row.get("review_title"))
        text = safe_text(row.get("review_description"))
        topics = safe_text(row.get("problem_topics_str"))
        patterns = safe_text(row.get("matched_patterns"))
        date = row.get("date")
        source = row.get("source")
        rating = row.get("rating")

        header = f"{date.date() if pd.notna(date) else '—'} | {source} | {rating:.0f}★"

        if topics:
            header += f" | {topics}"

        with st.expander(header):
            if title:
                st.markdown(f"**{title}**")

            st.write(text if text else "Текст отзыва отсутствует.")

            if patterns:
                st.caption(f"Сработавшие паттерны: {patterns}")

            if pd.notna(row.get("developer_response")):
                st.markdown("**Ответ разработчика:**")
                st.write(row.get("developer_response"))


# ============================================================
# MAIN APP
# ============================================================

def main() -> None:
    st.set_page_config(
        page_title="Операционная аналитика отзывов",
        layout="wide",
    )

    st.title("Операционная аналитика отзывов SmartMed / Storemed")
    st.caption(
        "Фокус: регулярный мониторинг проблем по отзывам 1–3★, "
        "словарная классификация паттернов, динамика и выгрузка для операционного директора."
    )

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

    min_date = df["date"].min().date()
    max_date = df["date"].max().date()

    with st.sidebar:
        st.header("Фильтры")

        date_range = st.date_input(
            "Период",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
        )

        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_date, end_date = date_range
        else:
            start_date, end_date = min_date, max_date

        selected_sources = st.multiselect(
            "Источник",
            options=sorted(df["source"].dropna().unique()),
            default=sorted(df["source"].dropna().unique()),
        )

        selected_countries = st.multiselect(
            "Страна магазина",
            options=sorted(df["store_country"].dropna().unique()),
            default=sorted(df["store_country"].dropna().unique()),
        )

        selected_ratings = st.multiselect(
            "Оценка",
            options=sorted(df["rating"].dropna().astype(int).unique()),
            default=sorted(df["rating"].dropna().astype(int).unique()),
        )

    mask = (
        (df["date"].dt.date >= start_date)
        & (df["date"].dt.date <= end_date)
        & df["source"].isin(selected_sources)
        & df["store_country"].isin(selected_countries)
        & df["rating"].astype(int).isin(selected_ratings)
    )

    df_selection = df.loc[mask].copy()

    if df_selection.empty:
        st.warning("По выбранным фильтрам данных нет.")
        st.stop()

    problem_df = df_selection[df_selection["is_problem"]].copy()
    topic_summary = build_topic_summary(df_selection)
    weekly_summary = build_weekly_summary(df_selection)
    source_summary = build_source_summary(df_selection)
    anomaly_days = build_anomaly_days(df_selection)

    tab_overview, tab_problems, tab_dynamics, tab_reviews, tab_export = st.tabs(
        [
            "Операционный обзор",
            "Проблемы 1–3★",
            "Динамика",
            "Отзывы",
            "Экспорт",
        ]
    )

    # ========================================================
    # TAB 1 — OVERVIEW
    # ========================================================

    with tab_overview:
        total_reviews = len(df_selection)
        problem_reviews = int(df_selection["is_problem"].sum())
        hard_problem_reviews = int(df_selection["is_hard_problem"].sum())
        avg_rating = df_selection["rating"].mean()
        problem_share = problem_reviews / total_reviews if total_reviews else 0
        hard_problem_share = hard_problem_reviews / total_reviews if total_reviews else 0
        text_share = df_selection["has_text"].mean()
        response_share = df_selection["has_developer_response"].mean()

        col1, col2, col3, col4 = st.columns(4)

        col1.metric("Всего отзывов", f"{total_reviews}")
        col2.metric("Средний рейтинг", f"{avg_rating:.2f}")
        col3.metric("Проблемные 1–3★", f"{problem_reviews}", format_percent(problem_share))
        col4.metric("Жёсткий негатив 1–2★", f"{hard_problem_reviews}", format_percent(hard_problem_share))

        col5, col6, col7, col8 = st.columns(4)

        col5.metric("Доля отзывов с текстом", format_percent(text_share))
        col6.metric("Ответы разработчика", format_percent(response_share))
        col7.metric("Проблемных тем", f"{len(topic_summary)}")
        col8.metric("Дней для проверки", f"{len(anomaly_days)}")

        st.markdown("---")

        if not topic_summary.empty:
            top_problem = topic_summary.iloc[0]

            st.info(
                f"Главный проблемный кластер: **{top_problem['problem_topic']}**. "
                f"Отзывов: **{int(top_problem['problem_reviews'])}**, "
                f"доля жёсткого негатива: **{format_percent(top_problem['hard_share'])}**, "
                f"приоритет: **{top_problem['priority']}**."
            )

        left, right = st.columns([1.2, 1])

        with left:
            st.subheader("Сравнение источников")

            if not source_summary.empty:
                source_plot = source_summary.copy()
                source_plot["problem_share_pct"] = source_plot["problem_share"] * 100

                fig_source = px.bar(
                    source_plot,
                    x="source",
                    y="problem_share_pct",
                    text="problem_share_pct",
                    title="Доля проблемных отзывов 1–3★ по источникам",
                )

                fig_source.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
                fig_source.update_layout(yaxis_title="Доля проблемных отзывов, %", xaxis_title="")
                st.plotly_chart(fig_source, use_container_width=True)

        with right:
            st.subheader("Сводка по источникам")
            st.dataframe(
                source_summary[
                    [
                        "source",
                        "reviews",
                        "avg_rating",
                        "problem_reviews",
                        "problem_share",
                        "hard_problem_share",
                        "text_share",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

        st.markdown("---")

        st.subheader("Последние проблемные отзывы")
        render_review_cards(
            problem_df.sort_values("review_date", ascending=False),
            limit=10,
        )

    # ========================================================
    # TAB 2 — PROBLEMS
    # ========================================================

    with tab_problems:
        st.subheader("Проблемные кластеры по отзывам 1–3★")

        if topic_summary.empty:
            st.warning("Проблемных отзывов по выбранному периоду нет.")
        else:
            topic_plot = topic_summary.sort_values("problem_reviews", ascending=True)

            fig_topics = px.bar(
                topic_plot,
                x="problem_reviews",
                y="problem_topic",
                orientation="h",
                color="priority",
                title="Количество проблемных отзывов по темам",
                hover_data=[
                    "hard_problem_reviews",
                    "rich_problem_reviews",
                    "avg_rating",
                    "score",
                ],
            )

            fig_topics.update_layout(xaxis_title="Отзывы 1–3★", yaxis_title="")
            st.plotly_chart(fig_topics, use_container_width=True)

            portfolio = topic_summary.copy()
            portfolio["hard_share_pct"] = portfolio["hard_share"] * 100

            fig_portfolio = px.scatter(
                portfolio,
                x="problem_reviews",
                y="hard_share_pct",
                size="score",
                color="priority",
                hover_name="problem_topic",
                title="Карта приоритета: частота проблемы × жёсткость негатива",
                hover_data=[
                    "avg_rating",
                    "rich_problem_reviews",
                    "last_review_date",
                ],
            )

            fig_portfolio.update_layout(
                xaxis_title="Количество проблемных отзывов",
                yaxis_title="Доля 1–2★ внутри темы, %",
            )

            st.plotly_chart(fig_portfolio, use_container_width=True)

            st.subheader("Таблица проблем")
            st.dataframe(
                topic_summary[
                    [
                        "problem_topic",
                        "priority",
                        "problem_reviews",
                        "problem_share",
                        "hard_problem_reviews",
                        "hard_share",
                        "rich_problem_reviews",
                        "avg_rating",
                        "score",
                        "last_review_date",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

            selected_topic = st.selectbox(
                "Показать примеры отзывов по теме",
                options=topic_summary["problem_topic"].tolist(),
            )

            topic_reviews = problem_df[
                problem_df["problem_topics"].apply(lambda topics: selected_topic in topics)
            ].sort_values(["is_hard_problem", "review_date"], ascending=[False, False])

            render_review_cards(topic_reviews, limit=15)

    # ========================================================
    # TAB 3 — DYNAMICS
    # ========================================================

    with tab_dynamics:
        st.subheader("Динамика отзывов и проблем")

        weekly_long = (
            df_selection
            .groupby(["week_start", "rating_group"], as_index=False)
            .agg(reviews=("review_id", "count"))
        )

        fig_weekly = px.bar(
            weekly_long,
            x="week_start",
            y="reviews",
            color="rating_group",
            title="Отзывы по неделям и группам рейтинга",
        )

        fig_weekly.update_layout(xaxis_title="Неделя", yaxis_title="Количество отзывов")
        st.plotly_chart(fig_weekly, use_container_width=True)

        if not weekly_summary.empty:
            weekly_total = (
                df_selection
                .groupby("week_start", as_index=False)
                .agg(
                    reviews=("review_id", "count"),
                    avg_rating=("rating", "mean"),
                    problem_reviews=("is_problem", "sum"),
                    hard_problem_reviews=("is_hard_problem", "sum"),
                )
                .sort_values("week_start")
            )

            weekly_total["problem_share"] = weekly_total["problem_reviews"] / weekly_total["reviews"]
            weekly_total["hard_problem_share"] = weekly_total["hard_problem_reviews"] / weekly_total["reviews"]

            fig_problem_share = px.line(
                weekly_total,
                x="week_start",
                y="problem_share",
                markers=True,
                title="Доля проблемных отзывов 1–3★ по неделям",
            )

            fig_problem_share.update_layout(
                xaxis_title="Неделя",
                yaxis_title="Доля проблемных отзывов",
            )

            st.plotly_chart(fig_problem_share, use_container_width=True)

            fig_rating = px.line(
                weekly_total,
                x="week_start",
                y="avg_rating",
                markers=True,
                title="Средний рейтинг по неделям",
            )

            fig_rating.update_layout(
                xaxis_title="Неделя",
                yaxis_title="Средний рейтинг",
            )

            st.plotly_chart(fig_rating, use_container_width=True)

        st.markdown("---")

        st.subheader("Дни, требующие проверки")

        if anomaly_days.empty:
            st.success("Аномальных дней по выбранным правилам не найдено.")
        else:
            st.dataframe(
                anomaly_days[
                    [
                        "date",
                        "reviews",
                        "avg_rating",
                        "problem_reviews",
                        "problem_share",
                        "empty_positive_reviews",
                        "empty_positive_share",
                        "activity_spike",
                        "problem_spike",
                        "positive_concentration",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

    # ========================================================
    # TAB 4 — REVIEWS
    # ========================================================

    with tab_reviews:
        st.subheader("Поиск и предпросмотр отзывов")

        col1, col2, col3 = st.columns(3)

        with col1:
            only_problems = st.checkbox("Только 1–3★", value=True)

        with col2:
            search_text = st.text_input("Поиск по тексту")

        with col3:
            available_topics = sorted(
                {
                    topic
                    for topics in df_selection["problem_topics"]
                    for topic in topics
                }
            )

            selected_topic_filter = st.selectbox(
                "Тема",
                options=["Все"] + available_topics,
            )

        review_view = df_selection.copy()

        if only_problems:
            review_view = review_view[review_view["is_problem"]]

        if search_text:
            review_view = review_view[
                review_view["review_text"]
                .str.lower()
                .str.contains(re.escape(search_text.lower()), na=False)
            ]

        if selected_topic_filter != "Все":
            review_view = review_view[
                review_view["problem_topics"]
                .apply(lambda topics: selected_topic_filter in topics)
            ]

        st.caption(f"Найдено отзывов: {len(review_view)}")

        render_review_cards(
            review_view.sort_values("review_date", ascending=False),
            limit=30,
        )

        with st.expander("Таблица"):
            st.dataframe(
                review_view[
                    [
                        "source",
                        "store_country",
                        "date",
                        "rating",
                        "rating_group",
                        "problem_topics_str",
                        "matched_patterns",
                        "user_name",
                        "review_title",
                        "review_description",
                        "developer_response",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

    # ========================================================
    # TAB 5 — EXPORT
    # ========================================================

    with tab_export:
        st.subheader("Экспорт аналитического пакета")

        export_columns = [
            "source",
            "store_country",
            "review_id",
            "date",
            "review_date",
            "rating",
            "rating_group",
            "is_problem",
            "is_hard_problem",
            "is_rich_problem",
            "problem_topics_str",
            "matched_patterns",
            "text_len",
            "user_name",
            "review_title",
            "review_description",
            "developer_response",
            "developer_response_date",
            "thumbs_up",
            "app_version",
        ]

        report_bytes = to_excel_report(
            filtered_df=df_selection[export_columns],
            problem_df=problem_df[export_columns],
            topic_summary=topic_summary,
            weekly_summary=weekly_summary,
            anomaly_days=anomaly_days,
            source_summary=source_summary,
        )

        st.download_button(
            label="Скачать Excel-отчёт",
            data=report_bytes,
            file_name="smartmed_operational_reviews_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown(
            """
            **Состав отчёта:**

            - `reviews_filtered` — все отзывы по фильтрам;
            - `problem_reviews` — только отзывы 1–3★;
            - `problem_topics` — сводка по проблемным темам;
            - `weekly_summary` — недельная динамика;
            - `anomaly_days` — дни, требующие проверки;
            - `source_summary` — сравнение Android / iOS.
            """
        )


if __name__ == "__main__":
    main()
