import io
import re
import zipfile
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st


GROUPS = [
    "Пашута ОПТ",
    "Родина, Самойлова",
    "Трошина, Гончарова",
    "Шакулова",
    "Суркова, Ромащенко, Бабушкина, Новожилова",
    "Селянкина, Королева",
]

DEFAULT_SELECTIONS = {
    "Пашута ОПТ": ["Пашута М.С.", "Пашута М.С. (Ростов)"],
    "Родина, Самойлова": ["Родина", "Самойлова", "Родина Е.В. (Ростов)"],
    "Трошина, Гончарова": ["Трошина Лариса"],
    "Шакулова": ["Шакулова Екатерина"],
    "Суркова, Ромащенко, Бабушкина, Новожилова": [
        "Суркова Н.",
        "Ромащенко Екатерина",
        "Бабушкина Виктория",
        "Новожилова М.",
    ],
    "Селянкина, Королева": ["Селянкина Татьяна", "Королева Светлана"],
}

MANAGER_BLOCKLIST = {
    "!!!",
    "<>",
    "андреева дарья",
    "антюфеева яна",
    "гордиенко",
    "дегтярев алексей",
    "дегтярева оксана александровна",
    "ермохина ирина",
    "клишко ю.н.",
    "никишова ольга",
    "пашута - сети",
    "пименова любовь",
    "стародубцева полина",
    "сотрудник авиаторов",
    "сотрудник санвэй",
    "существующие сотрудники",
    "яицкая ольга",
}

RETAIL_CLIENT_EXCLUDE_KEYWORDS = [
    "(сотрудник)",
    "ип",
    "(закрылась)",
    "(закрыта)",
    "ооо",
    "зао",
    "муп",
    "моу",
    "фгку",
    "ао",
    "мбоу",
    "оао",
    "гуп",
    "вдгоо",
    "тск",
    "гбпоу",
    "мбу",
    "гувд",
    "усзн",
    "ск-кристалл строительная компания",
    "стс-волгоград",
    "фгуп",
    "гуз",
]

EMAIL_EXCLUDE_FILES = [
    "Ne otkravali 300 dney roznica.xlsx",
    "Otpiski.xlsx",
    "Status_problemnie_otdel prodaj.xlsx",
]


@st.cache_data(show_spinner=False)
def parse_html_tables(html_bytes: bytes) -> List[pd.DataFrame]:
    """Читает HTML (cp1251) и возвращает список таблиц с заголовками из второй строки."""
    try:
        html_text = html_bytes.decode("cp1251", errors="replace")
        tables = pd.read_html(io.StringIO(html_text), header=1)
        for table in tables:
            table.columns = [str(col).strip() for col in table.columns]
        return tables
    except Exception as exc:  # noqa: BLE001 - показываем ошибку пользователю
        raise ValueError(f"Не удалось прочитать HTML: {exc}") from exc


@st.cache_data(show_spinner=False)
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Приводит все значения к строкам и нормализует заголовки колонок."""
    normalized = df.copy()
    normalized = normalized.fillna("")
    normalized = normalized.astype(str)
    normalized.columns = [
        re.sub(r"\s+", " ", str(col).replace("\xa0", " "))
        .replace("–", "-")
        .replace("—", "-")
        .replace("−", "-")
        .replace("‑", "-")
        .strip()
        .lower()
        for col in normalized.columns
    ]
    return normalized


def _find_column(columns: List[str], keywords: Tuple[str, ...]) -> str | None:
    """Ищет колонку по ключевым словам в нормализованных заголовках."""
    for col in columns:
        for keyword in keywords:
            if keyword in col:
                return col
    return None


def _find_header_row(df: pd.DataFrame) -> int | None:
    """Ищет строку, которая содержит заголовки Клиент/Email/Менеджер/Фамилия/Ответственный."""
    keywords = (
        "клиент",
        "e-mail",
        "email",
        "e mail",
        "менеджер",
        "фамилия",
        "ответственный",
        "отвественный",
    )
    for idx in range(len(df)):
        row_values = df.iloc[idx].astype(str).fillna("").tolist()
        normalized_cells = [
            re.sub(r"\s+", " ", value).strip().lower() for value in row_values
        ]
        if any(keyword in cell for cell in normalized_cells for keyword in keywords):
            return idx
    return None


def _promote_header_row(df: pd.DataFrame, header_idx: int) -> pd.DataFrame:
    """Поднимает строку header_idx в заголовок таблицы."""
    promoted = df.copy()
    new_columns = promoted.iloc[header_idx].astype(str).fillna("").tolist()
    promoted = promoted.drop(index=range(header_idx + 1)).reset_index(drop=True)
    promoted.columns = new_columns
    return promoted


def _normalize_manager(value: str) -> str:
    """Нормализует значения менеджеров для фильтрации."""
    cleaned = value.strip()
    cleaned = cleaned.replace("–", "-").replace("—", "-").replace("−", "-")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.casefold()


def _split_emails(value: str) -> List[str]:
    """Разбивает строку с email на список отдельных адресов."""
    prepared = value.replace("\r", "\n")
    prepared = re.sub(r"\s+\.\s+", " ", prepared)
    parts = re.split(r"[\s,;|/\\]+", prepared)
    return [part for part in parts if part]


def validate_email(email: str) -> bool:
    """Проверяет валидность email по заданным правилам."""
    if email.count("@") != 1:
        return False

    local_part, domain = email.split("@")
    if not local_part or not domain:
        return False
    if domain.count(".") == 0:
        return False
    if email.startswith(".") or email.endswith("."):
        return False

    allowed = re.compile(r"^[a-z0-9._%+-]+@[a-z0-9.-]+$")
    if not allowed.match(email):
        return False

    return True


def clean_and_expand_emails(df: pd.DataFrame, email_col: str) -> pd.DataFrame:
    """Очищает emails и разворачивает множественные адреса на отдельные строки."""
    rows = []
    for _, row in df.iterrows():
        raw_email = row[email_col]
        for part in _split_emails(raw_email):
            cleaned = part.replace(" ", "")
            cleaned = re.sub(r"[А-Яа-яЁё]", "", cleaned)
            cleaned = re.sub(r"[<>«»()\/|\\;,!?:“”]", "", cleaned)
            cleaned = cleaned.strip("-")
            cleaned = cleaned.lower()
            if not cleaned:
                continue
            if validate_email(cleaned):
                new_row = row.copy()
                new_row[email_col] = cleaned
                rows.append(new_row)
    return pd.DataFrame(rows, columns=df.columns)


def build_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """Собирает XLSX в памяти и возвращает байты."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()


def build_zip_archive(group_frames: Dict[str, pd.DataFrame]) -> bytes:
    """Формирует ZIP-архив с XLSX файлами и количеством email в названии."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zipf:
        for group_name, frame in group_frames.items():
            xlsx_bytes = build_xlsx_bytes(frame)
            email_count = len(frame)
            zipf.writestr(f"{group_name} ({email_count}).xlsx", xlsx_bytes)
    return buffer.getvalue()


@st.cache_data(show_spinner=False)
def load_excluded_emails() -> set[str]:
    """Загружает список email для исключения из внешних Excel-файлов рядом с app.py."""
    base_dir = Path(__file__).resolve().parent
    excluded: set[str] = set()
    for file_name in EMAIL_EXCLUDE_FILES:
        file_path = base_dir / file_name
        if not file_path.exists():
            continue
        try:
            file_df = pd.read_excel(file_path)
            file_df = normalize_columns(file_df)
            email_col = _find_column(list(file_df.columns), ("e-mail", "email", "e mail"))
            if not email_col:
                continue
            values = file_df[email_col].astype(str).str.strip().str.lower()
            excluded.update(value for value in values if "@" in value)
        except Exception:
            continue
    return excluded


st.set_page_config(page_title="Email Cleaner", layout="wide")

st.title("Обновление Email баз. Оптовые клиенты")

tab_opt, tab_corp, tab_retail_base, tab_retail_site = st.tabs(
    [
        "Оптовые клиенты",
        "Корпоративные клиенты",
        "Розничные клиенты (база)",
        "Розничные клиенты (сайт)",
    ]
)

with tab_opt:
    uploaded_file = st.file_uploader(
        "Загрузите HTML файл",
        type=["html", "htm"],
    )

    if uploaded_file is None:
        st.info("Загрузите файл, чтобы начать обработку.")
    else:
        log_messages: List[str] = []
        excluded_emails = load_excluded_emails()

        def _log(message: str) -> None:
            """Добавляет сообщение в журнал обработки."""
            log_messages.append(message)

        try:
            _log("Начинаем парсинг HTML и удаляем первую строку файла.")
            tables = parse_html_tables(uploaded_file.getvalue())
            _log(f"Найдено таблиц: {len(tables)}.")
        except ValueError as exc:
            st.error(str(exc))
            st.stop()

        if not tables:
            st.error("В HTML не найдено таблиц.")
            st.stop()

        if len(tables) > 1:
            _log("В HTML обнаружено несколько таблиц, ожидаем выбор пользователя.")
            st.warning("Найдено несколько таблиц. Выберите нужную.")
            table_options = []
            for idx, table in enumerate(tables, start=1):
                preview = table.head(3).to_string(index=False)
                table_options.append((idx, preview))

            selected = st.selectbox(
                "Таблица",
                options=table_options,
                format_func=lambda option: f"Таблица {option[0]}\n{option[1]}",
            )
            selected_table = tables[selected[0] - 1]
            _log(f"Выбрана таблица номер {selected[0]}.")
        else:
            selected_table = tables[0]
            _log("Используется единственная таблица в HTML.")

        normalized_table = normalize_columns(selected_table)
        columns = list(normalized_table.columns)
        _log(f"Нормализованные колонки: {', '.join(columns)}.")

        client_col = _find_column(columns, ("клиент",))
        email_col = _find_column(columns, ("e-mail", "email", "e mail"))
        manager_col = _find_column(columns, ("менеджер",))

        if not client_col or not email_col or not manager_col:
            _log("Обязательные колонки не найдены в заголовках, ищем ниже в строках.")
            header_idx = _find_header_row(selected_table.fillna("").astype(str))
            if header_idx is not None:
                _log(f"Заголовки найдены в строке {header_idx + 1}, поднимаем её.")
                selected_table = _promote_header_row(selected_table, header_idx)
                normalized_table = normalize_columns(selected_table)
                columns = list(normalized_table.columns)
                _log(f"Обновленные колонки: {', '.join(columns)}.")
                client_col = _find_column(columns, ("клиент",))
                email_col = _find_column(columns, ("e-mail", "email", "e mail"))
                manager_col = _find_column(columns, ("менеджер",))

        if not client_col or not email_col or not manager_col:
            st.error(
                "Не удалось найти все обязательные колонки: Клиент, E-mail, Менеджер. "
                "Проверьте заголовки таблицы."
            )
            st.stop()

        _log(
            "Найдены колонки: "
            f"Клиент -> {client_col}, Email -> {email_col}, Менеджер -> {manager_col}."
        )
        initial_count = len(normalized_table)
        _log(f"Исходных строк: {initial_count}.")

        step1 = normalized_table[normalized_table[email_col].str.contains("@", na=False)]
        step1_count = len(step1)
        _log(f"После фильтра Email осталось строк: {step1_count}.")

        filtered_manager = step1.copy()
        filtered_manager["_manager_norm"] = filtered_manager[manager_col].apply(_normalize_manager)
        step2 = filtered_manager[
            (filtered_manager["_manager_norm"] != "")
            & (~filtered_manager["_manager_norm"].isin(MANAGER_BLOCKLIST))
        ].drop(columns=["_manager_norm"])
        step2_count = len(step2)
        _log(f"После фильтра Менеджеров осталось строк: {step2_count}.")

        step3 = clean_and_expand_emails(step2, email_col)
        step3_count = len(step3)
        _log(f"После очистки и разбиения Email осталось строк: {step3_count}.")

        step4 = step3.drop_duplicates(subset=[email_col], keep="first")
        step4_count = len(step4)
        _log(f"После дедупликации осталось строк: {step4_count}.")

        if excluded_emails:
            before_exclude_count = len(step4)
            step4 = step4[~step4[email_col].str.lower().isin(excluded_emails)]
            _log(
                "После исключения email из внешних файлов осталось строк: "
                f"{len(step4)} (удалено {before_exclude_count - len(step4)})."
            )

        result_full = step4[[client_col, email_col, manager_col]].rename(
            columns={client_col: "Клиент", email_col: "Email", manager_col: "Менеджер"}
        )
        result_preview = result_full.drop(columns=["Менеджер"])

        st.subheader("Метрики обработки")
        metrics_cols = st.columns(5)
        metrics_cols[0].metric("Исходные строки", initial_count)
        metrics_cols[1].metric("После фильтра Email", step1_count, delta=step1_count - initial_count)
        metrics_cols[2].metric(
            "После фильтра Менеджеров", step2_count, delta=step2_count - step1_count
        )
        metrics_cols[3].metric("После очистки Email", step3_count, delta=step3_count - step2_count)
        metrics_cols[4].metric("После дедупликации", step4_count, delta=step4_count - step3_count)

        st.markdown("---")
        st.subheader("Распределение по группам")

        unique_managers = sorted(result_full["Менеджер"].unique())

        selected_managers = {}
        for group_name in GROUPS:
            key = f"group_{group_name}"
            if key not in st.session_state:
                defaults = [
                    name for name in DEFAULT_SELECTIONS.get(group_name, []) if name in unique_managers
                ]
                st.session_state[key] = defaults

            selected_in_other_groups = set()
            for other_group in GROUPS:
                other_key = f"group_{other_group}"
                if other_key == key:
                    continue
                selected_in_other_groups.update(st.session_state.get(other_key, []))

            current_selection = st.session_state.get(key, [])
            available_options = [
                manager
                for manager in unique_managers
                if manager not in selected_in_other_groups or manager in current_selection
            ]

            st.markdown(f"**{group_name}**")
            selected_managers[group_name] = st.multiselect(
                "Выбор менеджера",
                options=available_options,
                default=current_selection,
                key=key,
            )
        selected_all = set()
        for managers in selected_managers.values():
            selected_all.update(managers)

        remaining_managers = [manager for manager in unique_managers if manager not in selected_all]
        if remaining_managers:
            st.caption("Не выбраны: " + ", ".join(remaining_managers))
        else:
            st.caption("Выбраны все")

        group_frames = {}
        for group_name, managers in selected_managers.items():
            filtered = result_full[result_full["Менеджер"].isin(managers)].copy()
            group_frames[group_name] = filtered.drop(columns=["Менеджер"])

        zip_bytes = build_zip_archive(group_frames)
        st.download_button(
            label="Скачать архив XLSX",
            data=zip_bytes,
            file_name="email_groups.zip",
            mime="application/zip",
        )

        with st.expander("Журнал обработки", expanded=False):
            st.text("\n".join(log_messages))

with tab_corp:
    uploaded_file_corp = st.file_uploader(
        "Загрузите HTML файл",
        type=["html", "htm"],
        key="corp_uploader",
    )

    if uploaded_file_corp is None:
        st.info("Загрузите файл, чтобы начать обработку.")
    else:
        log_messages_corp: List[str] = []
        excluded_emails = load_excluded_emails()

        def _log_corp(message: str) -> None:
            """Добавляет сообщение в журнал обработки."""
            log_messages_corp.append(message)

        try:
            _log_corp("Начинаем парсинг HTML и удаляем первую строку файла.")
            tables_corp = parse_html_tables(uploaded_file_corp.getvalue())
            _log_corp(f"Найдено таблиц: {len(tables_corp)}.")
        except ValueError as exc:
            st.error(str(exc))
            st.stop()

        if not tables_corp:
            st.error("В HTML не найдено таблиц.")
            st.stop()

        if len(tables_corp) > 1:
            _log_corp("В HTML обнаружено несколько таблиц, ожидаем выбор пользователя.")
            st.warning("Найдено несколько таблиц. Выберите нужную.")
            table_options_corp = []
            for idx, table in enumerate(tables_corp, start=1):
                preview = table.head(3).to_string(index=False)
                table_options_corp.append((idx, preview))

            selected_corp = st.selectbox(
                "Таблица",
                options=table_options_corp,
                format_func=lambda option: f"Таблица {option[0]}\n{option[1]}",
                key="corp_table",
            )
            selected_table_corp = tables_corp[selected_corp[0] - 1]
            _log_corp(f"Выбрана таблица номер {selected_corp[0]}.")
        else:
            selected_table_corp = tables_corp[0]
            _log_corp("Используется единственная таблица в HTML.")

        normalized_table_corp = normalize_columns(selected_table_corp)
        columns_corp = list(normalized_table_corp.columns)
        _log_corp(f"Нормализованные колонки: {', '.join(columns_corp)}.")

        client_col_corp = _find_column(columns_corp, ("клиент",))
        email_col_corp = _find_column(columns_corp, ("e-mail", "email", "e mail"))
        manager_col_corp = _find_column(columns_corp, ("менеджер",))

        if not client_col_corp or not email_col_corp or not manager_col_corp:
            _log_corp("Обязательные колонки не найдены в заголовках, ищем ниже в строках.")
            header_idx_corp = _find_header_row(selected_table_corp.fillna("").astype(str))
            if header_idx_corp is not None:
                _log_corp(f"Заголовки найдены в строке {header_idx_corp + 1}, поднимаем её.")
                selected_table_corp = _promote_header_row(selected_table_corp, header_idx_corp)
                normalized_table_corp = normalize_columns(selected_table_corp)
                columns_corp = list(normalized_table_corp.columns)
                _log_corp(f"Обновленные колонки: {', '.join(columns_corp)}.")
                client_col_corp = _find_column(columns_corp, ("клиент",))
                email_col_corp = _find_column(columns_corp, ("e-mail", "email", "e mail"))
                manager_col_corp = _find_column(columns_corp, ("менеджер",))

        if not client_col_corp or not email_col_corp or not manager_col_corp:
            st.error(
                "Не удалось найти все обязательные колонки: Клиент, E-mail, Менеджер. "
                "Проверьте заголовки таблицы."
            )
            st.stop()

        _log_corp(
            "Найдены колонки: "
            f"Клиент -> {client_col_corp}, Email -> {email_col_corp}, Менеджер -> {manager_col_corp}."
        )
        initial_count_corp = len(normalized_table_corp)
        _log_corp(f"Исходных строк: {initial_count_corp}.")

        step1_corp = normalized_table_corp[
            normalized_table_corp[email_col_corp].str.contains("@", na=False)
        ]
        step1_count_corp = len(step1_corp)
        _log_corp(f"После фильтра Email осталось строк: {step1_count_corp}.")

        filtered_manager_corp = step1_corp.copy()
        filtered_manager_corp["_manager_norm"] = filtered_manager_corp[manager_col_corp].apply(
            _normalize_manager
        )
        step2_corp = filtered_manager_corp[
            (filtered_manager_corp["_manager_norm"] != "")
            & (~filtered_manager_corp["_manager_norm"].isin(MANAGER_BLOCKLIST))
        ].drop(columns=["_manager_norm"])
        step2_count_corp = len(step2_corp)
        _log_corp(f"После фильтра Менеджеров осталось строк: {step2_count_corp}.")

        step3_corp = clean_and_expand_emails(step2_corp, email_col_corp)
        step3_count_corp = len(step3_corp)
        _log_corp(f"После очистки и разбиения Email осталось строк: {step3_count_corp}.")

        step4_corp = step3_corp.drop_duplicates(subset=[email_col_corp], keep="first")
        step4_count_corp = len(step4_corp)
        _log_corp(f"После дедупликации осталось строк: {step4_count_corp}.")

        if excluded_emails:
            before_exclude_count_corp = len(step4_corp)
            step4_corp = step4_corp[~step4_corp[email_col_corp].str.lower().isin(excluded_emails)]
            _log_corp(
                "После исключения email из внешних файлов осталось строк: "
                f"{len(step4_corp)} (удалено {before_exclude_count_corp - len(step4_corp)})."
            )

        result_full_corp = step4_corp[[client_col_corp, email_col_corp, manager_col_corp]].rename(
            columns={
                client_col_corp: "Клиент",
                email_col_corp: "Email",
                manager_col_corp: "Менеджер",
            }
        )
        result_preview_corp = result_full_corp.drop(columns=["Менеджер"])

        st.subheader("Метрики обработки")
        metrics_cols_corp = st.columns(5)
        metrics_cols_corp[0].metric("Исходные строки", initial_count_corp)
        metrics_cols_corp[1].metric(
            "После фильтра Email", step1_count_corp, delta=step1_count_corp - initial_count_corp
        )
        metrics_cols_corp[2].metric(
            "После фильтра Менеджеров", step2_count_corp, delta=step2_count_corp - step1_count_corp
        )
        metrics_cols_corp[3].metric(
            "После очистки Email", step3_count_corp, delta=step3_count_corp - step2_count_corp
        )
        metrics_cols_corp[4].metric(
            "После дедупликации", step4_count_corp, delta=step4_count_corp - step3_count_corp
        )

        xlsx_bytes = build_xlsx_bytes(result_preview_corp)
        email_count_corp = len(result_preview_corp)
        st.download_button(
            label="Скачать файл XLSX",
            data=xlsx_bytes,
            file_name=f"corporate_emails ({email_count_corp}).xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Журнал обработки", expanded=False):
            st.text("\n".join(log_messages_corp))

with tab_retail_base:
    uploaded_file_retail_base = st.file_uploader(
        "Загрузите HTML файл",
        type=["html", "htm"],
        key="retail_base_uploader",
    )
    include_phone = st.checkbox("Номер телефона", key="retail_base_phone")

    if uploaded_file_retail_base is None:
        st.info("Загрузите файл, чтобы начать обработку.")
    else:
        log_messages_retail: List[str] = []
        excluded_emails = load_excluded_emails()

        def _log_retail(message: str) -> None:
            """Добавляет сообщение в журнал обработки."""
            log_messages_retail.append(message)

        try:
            _log_retail("Начинаем парсинг HTML и удаляем первую строку файла.")
            tables_retail = parse_html_tables(uploaded_file_retail_base.getvalue())
            _log_retail(f"Найдено таблиц: {len(tables_retail)}.")
        except ValueError as exc:
            st.error(str(exc))
            st.stop()

        if not tables_retail:
            st.error("В HTML не найдено таблиц.")
            st.stop()

        if len(tables_retail) > 1:
            _log_retail("В HTML обнаружено несколько таблиц, ожидаем выбор пользователя.")
            st.warning("Найдено несколько таблиц. Выберите нужную.")
            table_options_retail = []
            for idx, table in enumerate(tables_retail, start=1):
                preview = table.head(3).to_string(index=False)
                table_options_retail.append((idx, preview))

            selected_retail = st.selectbox(
                "Таблица",
                options=table_options_retail,
                format_func=lambda option: f"Таблица {option[0]}\n{option[1]}",
                key="retail_base_table",
            )
            selected_table_retail = tables_retail[selected_retail[0] - 1]
            _log_retail(f"Выбрана таблица номер {selected_retail[0]}.")
        else:
            selected_table_retail = tables_retail[0]
            _log_retail("Используется единственная таблица в HTML.")

        normalized_table_retail = normalize_columns(selected_table_retail)
        columns_retail = list(normalized_table_retail.columns)
        _log_retail(f"Нормализованные колонки: {', '.join(columns_retail)}.")

        client_col_retail = _find_column(columns_retail, ("клиент",))
        email_col_retail = _find_column(columns_retail, ("e-mail", "email", "e mail"))
        manager_col_retail = _find_column(columns_retail, ("менеджер",))
        place_col_retail = _find_column(columns_retail, ("место последней покупки",))
        birthday_col_retail = _find_column(columns_retail, ("день рождения",))
        phone_col_retail = _find_column(columns_retail, ("телефон",))

        if (
            not client_col_retail
            or not email_col_retail
            or not manager_col_retail
            or not place_col_retail
            or not birthday_col_retail
        ):
            _log_retail("Обязательные колонки не найдены в заголовках, ищем ниже в строках.")
            header_idx_retail = _find_header_row(selected_table_retail.fillna("").astype(str))
            if header_idx_retail is not None:
                _log_retail(f"Заголовки найдены в строке {header_idx_retail + 1}, поднимаем её.")
                selected_table_retail = _promote_header_row(
                    selected_table_retail, header_idx_retail
                )
                normalized_table_retail = normalize_columns(selected_table_retail)
                columns_retail = list(normalized_table_retail.columns)
                _log_retail(f"Обновленные колонки: {', '.join(columns_retail)}.")
                client_col_retail = _find_column(columns_retail, ("клиент",))
                email_col_retail = _find_column(columns_retail, ("e-mail", "email", "e mail"))
                manager_col_retail = _find_column(columns_retail, ("менеджер",))
                place_col_retail = _find_column(columns_retail, ("место последней покупки",))
                birthday_col_retail = _find_column(columns_retail, ("день рождения",))
                phone_col_retail = _find_column(columns_retail, ("телефон",))

        if (
            not client_col_retail
            or not email_col_retail
            or not manager_col_retail
            or not place_col_retail
            or not birthday_col_retail
        ):
            st.error(
                "Не удалось найти все обязательные колонки: Клиент, E-mail, Менеджер, "
                "Место последней покупки, День рождения. Проверьте заголовки таблицы."
            )
            st.stop()

        if include_phone and not phone_col_retail:
            st.error("Не удалось найти колонку Телефон. Проверьте заголовки таблицы.")
            st.stop()

        _log_retail(
            "Найдены колонки: "
            f"Клиент -> {client_col_retail}, Email -> {email_col_retail}, "
            f"Менеджер -> {manager_col_retail}, Место последней покупки -> {place_col_retail}, "
            f"День рождения -> {birthday_col_retail}."
        )
        initial_count_retail = len(normalized_table_retail)
        _log_retail(f"Исходных строк: {initial_count_retail}.")

        client_values_retail = normalized_table_retail[client_col_retail].astype(str).str.casefold()
        exclude_pattern_retail = r"|".join(
            rf"\b{keyword}\b" if keyword in {"ип", "ао"} else re.escape(keyword)
            for keyword in RETAIL_CLIENT_EXCLUDE_KEYWORDS
        )
        step1_retail = normalized_table_retail[
            normalized_table_retail[email_col_retail].str.contains("@", na=False)
            & ~client_values_retail.str.contains(exclude_pattern_retail, na=False, regex=True)
        ]
        step1_count_retail = len(step1_retail)
        _log_retail(f"После фильтра Email осталось строк: {step1_count_retail}.")

        filtered_manager_retail = step1_retail.copy()
        filtered_manager_retail["_manager_norm"] = filtered_manager_retail[
            manager_col_retail
        ].apply(_normalize_manager)
        step2_retail = filtered_manager_retail[
            filtered_manager_retail["_manager_norm"] == ""
        ].drop(columns=["_manager_norm"])
        step2_count_retail = len(step2_retail)
        _log_retail(f"После фильтра Менеджеров осталось строк: {step2_count_retail}.")

        step3_retail = clean_and_expand_emails(step2_retail, email_col_retail)
        step3_count_retail = len(step3_retail)
        _log_retail(f"После очистки и разбиения Email осталось строк: {step3_count_retail}.")

        step4_retail = step3_retail.drop_duplicates(subset=[email_col_retail], keep="first")
        step4_count_retail = len(step4_retail)
        _log_retail(f"После дедупликации осталось строк: {step4_count_retail}.")

        if excluded_emails:
            before_exclude_count_retail = len(step4_retail)
            step4_retail = step4_retail[
                ~step4_retail[email_col_retail].str.lower().isin(excluded_emails)
            ]
            _log_retail(
                "После исключения email из внешних файлов осталось строк: "
                f"{len(step4_retail)} (удалено {before_exclude_count_retail - len(step4_retail)})."
            )

        result_full_retail = step4_retail[
            [client_col_retail, email_col_retail, place_col_retail, birthday_col_retail]
        ].rename(
            columns={
                client_col_retail: "Клиент",
                email_col_retail: "Email",
                place_col_retail: "Место последней покупки",
                birthday_col_retail: "День рождения",
            }
        )

        if include_phone:
            result_full_retail["Телефон"] = step4_retail[phone_col_retail].values

        st.subheader("Метрики обработки")
        metrics_cols_retail = st.columns(5)
        metrics_cols_retail[0].metric("Исходные строки", initial_count_retail)
        metrics_cols_retail[1].metric(
            "После фильтра Email", step1_count_retail, delta=step1_count_retail - initial_count_retail
        )
        metrics_cols_retail[2].metric(
            "После фильтра Менеджеров",
            step2_count_retail,
            delta=step2_count_retail - step1_count_retail,
        )
        metrics_cols_retail[3].metric(
            "После очистки Email", step3_count_retail, delta=step3_count_retail - step2_count_retail
        )
        metrics_cols_retail[4].metric(
            "После дедупликации", step4_count_retail, delta=step4_count_retail - step3_count_retail
        )

        xlsx_bytes_retail = build_xlsx_bytes(result_full_retail)
        email_count_retail = len(result_full_retail)
        st.download_button(
            label="Скачать файл XLSX",
            data=xlsx_bytes_retail,
            file_name=f"retail_base_emails ({email_count_retail}).xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Журнал обработки", expanded=False):
            st.text("\n".join(log_messages_retail))

with tab_retail_site:
    uploaded_file_retail_site = st.file_uploader(
        "Загрузите XLS/XLSX файл",
        type=["xls", "xlsx"],
        key="retail_site_uploader",
    )

    if uploaded_file_retail_site is None:
        st.info("Загрузите файл, чтобы начать обработку.")
    else:
        log_messages_site: List[str] = []
        excluded_emails = load_excluded_emails()

        def _log_site(message: str) -> None:
            """Добавляет сообщение в журнал обработки."""
            log_messages_site.append(message)

        try:
            file_name_site = (uploaded_file_retail_site.name or "").lower()
            file_bytes_site = uploaded_file_retail_site.getvalue()
            is_xlsx = file_bytes_site[:2] == b"PK"
            primary_engine = "openpyxl" if is_xlsx else "xlrd"
            fallback_engine = "xlrd" if primary_engine == "openpyxl" else "openpyxl"
            _log_site(
                f"Начинаем парсинг XLS/XLSX файла (engine={primary_engine}, is_xlsx={is_xlsx})."
            )
            try:
                tables_site = [
                    pd.read_excel(io.BytesIO(file_bytes_site), engine=primary_engine)
                ]
            except Exception as exc:  # noqa: BLE001 - показываем ошибку пользователю
                _log_site(
                    f"Не удалось прочитать файл через {primary_engine}: {exc}. "
                    f"Пробуем {fallback_engine}."
                )
                tables_site = [
                    pd.read_excel(io.BytesIO(file_bytes_site), engine=fallback_engine)
                ]
            _log_site(f"Найдено таблиц: {len(tables_site)}.")
        except Exception as exc:  # noqa: BLE001 - показываем ошибку пользователю
            st.error(str(exc))
            st.stop()

        if not tables_site:
            st.error("В HTML не найдено таблиц.")
            st.stop()

        if len(tables_site) > 1:
            _log_site("В HTML обнаружено несколько таблиц, ожидаем выбор пользователя.")
            st.warning("Найдено несколько таблиц. Выберите нужную.")
            table_options_site = []
            for idx, table in enumerate(tables_site, start=1):
                preview = table.head(3).to_string(index=False)
                table_options_site.append((idx, preview))

            selected_site = st.selectbox(
                "Таблица",
                options=table_options_site,
                format_func=lambda option: f"Таблица {option[0]}\n{option[1]}",
                key="retail_site_table",
            )
            selected_table_site = tables_site[selected_site[0] - 1]
            _log_site(f"Выбрана таблица номер {selected_site[0]}.")
        else:
            selected_table_site = tables_site[0]
            _log_site("Используется единственная таблица в HTML.")

        normalized_table_site = normalize_columns(selected_table_site)
        columns_site = list(normalized_table_site.columns)
        _log_site(f"Нормализованные колонки: {', '.join(columns_site)}.")

        last_name_col_site = _find_column(columns_site, ("фамилия",))
        email_col_site = _find_column(columns_site, ("e-mail", "email", "e mail"))
        owner_col_site = _find_column(columns_site, ("ответственный", "отвественный"))

        if not last_name_col_site or not email_col_site or not owner_col_site:
            _log_site("Обязательные колонки не найдены в заголовках, ищем ниже в строках.")
            header_idx_site = _find_header_row(selected_table_site.fillna("").astype(str))
            if header_idx_site is not None:
                _log_site(f"Заголовки найдены в строке {header_idx_site + 1}, поднимаем её.")
                selected_table_site = _promote_header_row(selected_table_site, header_idx_site)
                normalized_table_site = normalize_columns(selected_table_site)
                columns_site = list(normalized_table_site.columns)
                _log_site(f"Обновленные колонки: {', '.join(columns_site)}.")
                last_name_col_site = _find_column(columns_site, ("фамилия",))
                email_col_site = _find_column(columns_site, ("e-mail", "email", "e mail"))
                owner_col_site = _find_column(columns_site, ("ответственный", "отвественный"))

        if not last_name_col_site or not email_col_site or not owner_col_site:
            st.error(
                "Не удалось найти все обязательные колонки: Фамилия, E-mail, Ответственный. "
                "Проверьте заголовки таблицы."
            )
            st.stop()

        _log_site(
            "Найдены колонки: "
            f"Фамилия -> {last_name_col_site}, Email -> {email_col_site}, "
            f"Ответственный -> {owner_col_site}."
        )
        initial_count_site = len(normalized_table_site)
        _log_site(f"Исходных строк: {initial_count_site}.")

        owner_values_site = normalized_table_site[owner_col_site].astype(str).str.strip()
        step1_site = normalized_table_site[
            normalized_table_site[email_col_site].str.contains("@", na=False)
            & (owner_values_site == "Интернет розница")
        ]
        step1_count_site = len(step1_site)
        _log_site(f"После фильтра Email/Ответственный осталось строк: {step1_count_site}.")

        step2_site = clean_and_expand_emails(step1_site, email_col_site)
        step2_count_site = len(step2_site)
        _log_site(f"После очистки и разбиения Email осталось строк: {step2_count_site}.")

        step3_site = step2_site.drop_duplicates(subset=[email_col_site], keep="first")
        step3_count_site = len(step3_site)
        _log_site(f"После дедупликации осталось строк: {step3_count_site}.")

        if excluded_emails:
            before_exclude_count_site = len(step3_site)
            step3_site = step3_site[~step3_site[email_col_site].str.lower().isin(excluded_emails)]
            _log_site(
                "После исключения email из внешних файлов осталось строк: "
                f"{len(step3_site)} (удалено {before_exclude_count_site - len(step3_site)})."
            )

        result_full_site = step3_site[[last_name_col_site, email_col_site]].rename(
            columns={last_name_col_site: "Фамилия", email_col_site: "E-Mail"}
        )

        st.subheader("Метрики обработки")
        metrics_cols_site = st.columns(4)
        metrics_cols_site[0].metric("Исходные строки", initial_count_site)
        metrics_cols_site[1].metric(
            "После фильтра", step1_count_site, delta=step1_count_site - initial_count_site
        )
        metrics_cols_site[2].metric(
            "После очистки Email", step2_count_site, delta=step2_count_site - step1_count_site
        )
        metrics_cols_site[3].metric(
            "После дедупликации", step3_count_site, delta=step3_count_site - step2_count_site
        )

        xlsx_bytes_site = build_xlsx_bytes(result_full_site)
        email_count_site = len(result_full_site)
        st.download_button(
            label="Скачать файл XLSX",
            data=xlsx_bytes_site,
            file_name=f"retail_site_emails ({email_count_site}).xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Журнал обработки", expanded=False):
            st.text("\n".join(log_messages_site))
