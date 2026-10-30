"""
Ежедневная выгрузка показателей рекламных кабинетов Meta Ads
в Google Sheets, с несколькими кабинетами и несколькими таблицами.

Структура таблицы:
    A - Дата
    B - Бюджет (spend)
    C - Кол-во заявок (тип определяется по objective/optimization_goal кампании,
        с фоллбэком на нейминг 'purchase'/'direct'/'директ'/'смс' для legacy-кампаний)
    K - CPM (взвешенный по всему кабинету: сумма расхода / сумма показов * 1000)
    L - Охват
    M - Клики по ссылке (сумма action_type "link_click" по всем кампаниям)
    N - Клики (все) (сумма поля "clicks" по всем кампаниям)
    D, E, F, G, I, H - НЕ ТРОГАЕМ, там формулы/ручные данные.

Режимы запуска:
    python meta_ads_to_sheets.py -> обычный режим: обновляет
    python meta_ads_to_sheets.py --backfill-month -> ручной разовый прогон:
        заполняет ВЕСЬ текущий месяц с 1-го числа по вчерашний день включительно..

Логика по месяцам:
- Каждому кабинету соответствует своя Google-таблица (см. accounts_config.py).
- Внутри таблицы данные лежат на листе с именем "MM.YYYY" (например "07.2026").
  Строка для конкретного дня месяца фиксирована: день 15 -> строка (header_row + 15).
- Если наступил новый месяц и листа для него ещё нет — скрипт дублирует лист
  предыдущего месяца (сохраняя все формулы и форматирование), подгоняет число
  дневных строк под длину нового месяца, и очищает ТОЛЬКО те колонки, которые
  сам заполняет (column_map). Остальные колонки (ваши формулы) не трогаются.

Один рекламный аккаунт может обслуживать НЕСКОЛЬКО услуг сразу (например,
разные направления бизнеса на одном кабинете) — тогда в accounts_config.py
у такого кабинета вместо одного spreadsheet_id указывается список "services",
каждая со своими ключевыми словами (ищутся в названии кампании без учёта
регистра) и своей таблицей. Meta запрашивается ОДИН РАЗ на кабинет за день,
а дальше кампании распределяются по услугам локально, без дублирования
запросов к API.

Установка зависимостей:
    pip install facebook-business gspread google-auth
"""

import argparse
import calendar
import datetime
import logging
import time
from typing import Any

import gspread
from facebook_business.adobjects.adaccount import AdAccount
from facebook_business.api import FacebookAdsApi
from google.oauth2.service_account import Credentials

import config
from accounts_config import ACCOUNTS

# ============================ ОБЩИЕ НАСТРОЙКИ ============================

FB_APP_ID = config.FB_APP_ID
FB_APP_SECRET = config.FB_APP_SECRET
FB_ACCESS_TOKEN = config.FB_ACCESS_TOKEN

GOOGLE_SERVICE_ACCOUNT_FILE = config.GOOGLE_SERVICE_ACCOUNT_FILE

# Актуализация за X дней
REFRESH_WINDOW_DAYS = 3

# Колонки таблицы
DEFAULT_COLUMN_MAP = {
    "date": "A",
    "spend": "B",
    "requests": "C",
    "cpm": "K",
    "reach": "L",
    "link_click": "M",
    "click": "N",
}

DEFAULT_HEADER_ROW = 1  # строка с заголовками; данные за 1-е число месяца начинаются со следующей строки

# --------------------- Классификация кампаний по названию ---------------------
# Если в названии кампании встречается одно из PURCHASE_KEYWORDS -> считаем
# "заявки" по PURCHASE_ACTION_TYPE. Если встречается одно из MESSAGE_KEYWORDS ->
# считаем "заявки" по MESSAGE_ACTION_TYPE. Названия проверяются без учёта регистра.

PURCHASE_KEYWORDS = ["purchase"]
MESSAGE_KEYWORDS = ["direct", "директ", "смс", "sms", "повідомлення", "пов"]
LEAD_KEYWORDS = ["lead", "лид", "квиз", "leds", "quiz"]

# Классификация по цели кампании (objective/optimization_goal) — основной сигнал
PURCHASE_OBJECTIVES = {"OUTCOME_SALES"}
MESSAGE_OBJECTIVES = {"OUTCOME_ENGAGEMENT"}
MESSAGE_OPTIMIZATION_GOALS = {"REPLIES", "CONVERSATIONS"}
LEAD_OBJECTIVES = {"OUTCOME_LEADS"}

PURCHASE_ACTION_TYPE = "purchase"
MESSAGE_ACTION_TYPE = "onsite_conversion.messaging_conversation_started_7d"

# Лиды в Meta могут приходить под разными action_type. Если лид мета или на строннем источнике
# Берём первый найденный и по нему считаем

# !!! ------  Доделать: оптимизировать под запуск на разные типы лидов ----------!!!

LEAD_ACTION_TYPE_PRIORITY = [
    "lead",
    "onsite_conversion.lead_grouped",
    "offsite_conversion.fb_pixel_lead",
]

# ==========================================================================


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)


def month_sheet_name(d: datetime.date) -> str:
    return d.strftime("%m.%Y")


def previous_month_date(d: datetime.date) -> datetime.date:
    first_of_this_month = d.replace(day=1)
    return first_of_this_month - datetime.timedelta(days=1)


def classify_campaign(
    campaign_name: str,
    objective: str | None = None,
    optimization_goal: str | None = None,
) -> str | None:
    """
    Возвращает 'purchase', 'message' или None (не удалось определить тип).
    """
    if objective in PURCHASE_OBJECTIVES:
        return "purchase"
    if objective in MESSAGE_OBJECTIVES and optimization_goal in MESSAGE_OPTIMIZATION_GOALS:
        return "message"
    if objective in LEAD_OBJECTIVES:
        return "leads"

    name_lower = (campaign_name or "").lower()
    if any(kw in name_lower for kw in PURCHASE_KEYWORDS):
        return "purchase"
    if any(kw in name_lower for kw in MESSAGE_KEYWORDS):
        return "message"
    if any(kw in name_lower for kw in LEAD_KEYWORDS):
        return "leads"
    return None


def get_action_value(actions: list[dict[str, Any]] | None, action_type: str) -> float:
    if not actions:
        return 0.0
    for a in actions:
        if a.get("action_type") == action_type:
            return float(a.get("value", 0) or 0)
    return 0.0


def get_action_value_priority(actions: list[dict[str, Any]] | None, priority_types: list[str]) -> float:
    """
    Как get_action_value, но проверяет список приоритетных action_type и
    возвращает значение ПЕРВОГО найденного — не складывает несколько типов
    вместе, чтобы не задвоить один и тот же лид, посчитанный под разными
    типами атрибуции одновременно.
    """
    if not actions:
        return 0.0
    for action_type in priority_types:
        for a in actions:
            if a.get("action_type") == action_type:
                return float(a.get("value", 0) or 0)
    return 0.0


def fetch_campaign_rows(fb_account_id: str, date_str: str) -> list[dict[str, Any]]:
    """Сырые данные по кампаниям кабинета за один день — без агрегации, чтобы можно было
    один и тот же набор строк агрегировать по-разному (разбить по услугам)."""
    account = AdAccount(fb_account_id)
    fields = [
        "campaign_name", "objective", "optimization_goal",
        "spend", "impressions", "clicks", "reach", "actions",
    ]
    params = {
        "level": "campaign",
        "time_range": {"since": date_str, "until": date_str},
    }
    insights = account.get_insights(fields=fields, params=params)
    return [dict(row) for row in insights]


def filter_rows_by_keywords(rows: list[dict[str, Any]], keywords: list[str]) -> list[dict[str, Any]]:
    """Оставляет только строки, у которых в campaign_name встречается одно из keywords (без учёта регистра)."""
    return [
        row for row in rows
        if any(kw.lower() in (row.get("campaign_name") or "").lower() for kw in keywords)
    ]


def split_rows_for_targets(rows: list[dict[str, Any]], targets: list[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    """
    Распределяет один и тот же набор строк кампаний по списку targets, в том же
    порядке. Логика (важен порядок применения):
    1. Целям с "keywords" достаются строки, где название кампании содержит
       одно из этих ключевых слов.
    2. Цели с "is_catch_all": True достаются ВСЕ строки, не подошедшие ни
       под одно ключевое слово ни одной из целей с keywords (даже если та
       цель не входит в текущую обработку — множество "занятых" строк
       считается по ВСЕМ keyword-целям кабинета).
    3. Обычные цели без keywords и без is_catch_all (простой кабинет без
       разбивки на услуги) получают все строки как есть.
    """
    result: list[list[dict[str, Any]]] = [None] * len(targets)  # type: ignore[list-item]
    matched_row_ids = set()

    for i, target in enumerate(targets):
        if target.get("keywords"):
            matched = filter_rows_by_keywords(rows, target["keywords"])
            result[i] = matched
            matched_row_ids.update(id(row) for row in matched)

    for i, target in enumerate(targets):
        if target.get("is_catch_all"):
            result[i] = [row for row in rows if id(row) not in matched_row_ids]

    for i, target in enumerate(targets):
        if result[i] is None:
            result[i] = rows

    return result


def aggregate_campaign_rows(
    rows: list[dict[str, Any]],
    custom_requests_action_type: str | None = None,
) -> dict[str, Any]:
    """
    Агрегирует уже полученный список строк кампаний в одну сводку за день:
    - spend: сумма расхода по всем строкам
    - requests: сумма "заявок" (purchase/message/leads, по классификации кампании,
      либо custom_requests_action_type, если задан)
    - cpm: взвешенный CPM (spend/impressions*1000, а не среднее по кампаниям —
      усреднять готовые CPM было бы математически неверно при разном расходе)
    - reach: сумма поля "reach" (охват)
    - link_click: сумма action_type "link_click"
    - click: сумма поля "clicks" (все клики)
    - has_data: False, если строк нет вообще (реклама не откручивалась / после
      фильтрации по ключевым словам ничего не осталось) — вызывающий код должен
      пропустить запись, а не писать нули
    """
    if not rows:
        return {"has_data": False}

    total_spend = 0.0
    total_impressions = 0.0
    total_clicks_all = 0.0
    total_link_clicks = 0.0
    total_reach = 0.0
    total_requests = 0.0
    unclassified_campaigns = []

    for row in rows:
        campaign_name = row.get("campaign_name", "")
        objective = row.get("objective")
        optimization_goal = row.get("optimization_goal")
        spend = float(row.get("spend", 0) or 0)
        impressions = float(row.get("impressions", 0) or 0)
        clicks_all = float(row.get("clicks", 0) or 0)
        reach = float(row.get("reach", 0) or 0)
        actions = row.get("actions")

        total_spend += spend
        total_impressions += impressions
        total_clicks_all += clicks_all
        total_reach += reach
        total_link_clicks += get_action_value(actions, "link_click")

        if custom_requests_action_type:
            # У этого кабинета/услуги отдельная кастомная конверсия вместо
            # классификации purchase/message — суммируем её по ВСЕМ
            # кампаниям сразу, без деления по названию/objective.
            total_requests += get_action_value(actions, custom_requests_action_type)
            continue

        category = classify_campaign(campaign_name, objective, optimization_goal)
        if category == "purchase":
            total_requests += get_action_value(actions, PURCHASE_ACTION_TYPE)
        elif category == "message":
            total_requests += get_action_value(actions, MESSAGE_ACTION_TYPE)
        elif category == "leads":
            total_requests += get_action_value_priority(actions, LEAD_ACTION_TYPE_PRIORITY)
        else:
            unclassified_campaigns.append(campaign_name)

    if unclassified_campaigns:
        logger.warning(
            "Не удалось определить тип %d кампани(й) (ни objective/optimization_goal, "
            "ни 'purchase'/'direct'/'директ'/'смс' в названии не подошли), их 'заявки' "
            "не посчитаны: %s",
            len(unclassified_campaigns), ", ".join(unclassified_campaigns),
        )

    cpm = (total_spend / total_impressions * 1000) if total_impressions > 0 else 0

    return {
        "has_data": True,
        "spend": round(total_spend, 2),
        "requests": round(total_requests),
        "cpm": round(cpm, 2),
        "reach": round(total_reach),
        "link_click": round(total_link_clicks),
        "click": round(total_clicks_all),
    }


def get_account_daily_insights(
    fb_account_id: str,
    date_str: str,
    custom_requests_action_type: str | None = None,
) -> dict[str, Any]:
    """Удобная обёртка: получить + агрегировать все кампании кабинета за день (без разбивки по услугам)."""
    rows = fetch_campaign_rows(fb_account_id, date_str)
    return aggregate_campaign_rows(rows, custom_requests_action_type)


def get_targets_for_account(acc: dict[str, Any]) -> list[dict[str, Any]]:
    """
    Превращает запись из ACCOUNTS в список "целей записи" (targets) — куда и
    как писать данные. Для обычного кабинета это один target на всю таблицу.

    Если у кабинета есть "services" (один fb_account_id, несколько услуг —
    например разные направления бизнеса на одном рекламном аккаунте), то для
    каждой услуги строится отдельный target со своими keywords/spreadsheet_id:
    строки кампаний перед агрегацией фильтруются по этим keywords (ищутся в
    названии кампании без учёта регистра), так что каждая услуга получает
    только "свои" кампании и пишется в свою таблицу.

    Одна из услуг может быть помечена "is_catch_all": True вместо "keywords" —
    такой target получает ВСЕ кампании, которые не подошли ни под одно
    ключевое слово ни одной другой услуги (см. split_rows_for_targets).
    """
    if "services" in acc:
        targets = []
        for service in acc["services"]:
            targets.append({
                "name": f"{acc['name']} — {service['name']}",
                "spreadsheet_id": service["spreadsheet_id"],
                "column_map": service.get("column_map", acc.get("column_map", DEFAULT_COLUMN_MAP)),
                "header_row": service.get("header_row", acc.get("header_row", DEFAULT_HEADER_ROW)),
                "custom_requests_action_type": service.get(
                    "custom_requests_action_type", acc.get("custom_requests_action_type")
                ),
                "keywords": service.get("keywords"),
                "is_catch_all": service.get("is_catch_all", False),
            })
        return targets

    return [{
        "name": acc["name"],
        "spreadsheet_id": acc["spreadsheet_id"],
        "column_map": acc.get("column_map", DEFAULT_COLUMN_MAP),
        "header_row": acc.get("header_row", DEFAULT_HEADER_ROW),
        "custom_requests_action_type": acc.get("custom_requests_action_type"),
        "keywords": None,
    }]


def clear_script_columns(ws: gspread.Worksheet, column_map: dict[str, str], header_row: int, days_in_month: int) -> None:
    """Очищает ТОЛЬКО значения в колонках скрипта, для строк с 1-го по последний день месяца."""
    first_row = header_row + 1
    last_row = header_row + days_in_month
    ranges = [f"{col}{first_row}:{col}{last_row}" for col in column_map.values()]
    ws.batch_clear(ranges)
    logger.info("Очищены колонки скрипта (%s) на диапазоне строк %d-%d", ", ".join(column_map.values()), first_row, last_row)


def adjust_days_in_month_rows(
    spreadsheet: gspread.Spreadsheet,
    ws: gspread.Worksheet,
    header_row: int,
    source_days_in_month: int,
    target_days_in_month: int,
) -> None:
    """
    Подгоняет количество дневных строк под фактическую длину нового месяца.
    Вставляет/удаляет строки ПРЯМО ПЕРЕД итоговой строкой — она сдвигается
    вниз/вверх вместе со своими формулами, ничего в ней не пересчитывается.
    """
    diff = target_days_in_month - source_days_in_month
    if diff == 0:
        return

    old_summary_row_number = header_row + source_days_in_month + 1
    start_index = old_summary_row_number - 1
    sheet_id = ws.id

    if diff > 0:
        request = {
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": start_index,
                    "endIndex": start_index + diff,
                },
                "inheritFromBefore": True,
            }
        }
        logger.info("Месяц длиннее предыдущего: вставляю %d стро(ку/ки) перед итоговой строкой", diff)
    else:
        remove_count = -diff
        request = {
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": start_index - remove_count,
                    "endIndex": start_index,
                }
            }
        }
        logger.info("Месяц короче предыдущего: удаляю %d лишн(юю/их) дневн(ую/ых) строк(у/и) перед итоговой строкой", remove_count)

    spreadsheet.batch_update({"requests": [request]})


def ensure_current_month_sheet(
    spreadsheet: gspread.Spreadsheet,
    target_date: datetime.date,
    column_map: dict[str, str],
    header_row: int,
) -> gspread.Worksheet:
    """Возвращает лист текущего месяца; при необходимости создаёт его дублированием предыдущего."""
    target_name = month_sheet_name(target_date)

    try:
        return spreadsheet.worksheet(target_name)
    except gspread.WorksheetNotFound:
        pass

    prev_date = previous_month_date(target_date)
    prev_name = month_sheet_name(prev_date)

    try:
        prev_ws = spreadsheet.worksheet(prev_name)
    except gspread.WorksheetNotFound:
        raise RuntimeError(
            f"Не найден лист '{prev_name}' для дублирования в таблице '{spreadsheet.title}'. "
            f"Создайте вручную первый лист месяца с нужной структурой, дальше скрипт будет "
            f"дублировать его сам."
        )

    logger.info("Новый месяц: дублирую лист '%s' -> '%s' в таблице '%s'", prev_name, target_name, spreadsheet.title)
    new_ws = spreadsheet.duplicate_sheet(source_sheet_id=prev_ws.id, new_sheet_name=target_name)

    days_in_prev_month = calendar.monthrange(prev_date.year, prev_date.month)[1]
    days_in_month = calendar.monthrange(target_date.year, target_date.month)[1]

    adjust_days_in_month_rows(spreadsheet, new_ws, header_row, days_in_prev_month, days_in_month)
    clear_script_columns(new_ws, column_map, header_row, days_in_month)

    return new_ws


def _coerce_numeric(value: Any) -> Any:
    """Приводит строковое число к настоящему Python float/int"""
    if isinstance(value, (int, float)):
        return value
    if value in (None, ""):
        return value
    try:
        normalized = str(value).strip().replace(" ", "").replace(",", ".")
        as_float = float(normalized)
        return int(as_float) if as_float.is_integer() else round(as_float, 2)
    except (TypeError, ValueError):
        return value


def write_daily_row(
    ws: gspread.Worksheet,
    target_date: datetime.date,
    data: dict[str, Any],
    column_map: dict[str, str],
    header_row: int,
) -> None:
    row_number = header_row + target_date.day
    data_with_date = dict(data)
    data_with_date["date"] = target_date.strftime("%d.%m.%Y")

    updates = []
    for field, col_letter in column_map.items():
        raw_value = data_with_date.get(field, "")
        value = raw_value if field == "date" else _coerce_numeric(raw_value)
        updates.append({"range": f"{col_letter}{row_number}", "values": [[value]]})

    ws.batch_update(updates, value_input_option="USER_ENTERED")
    logger.info("Записана строка %d (%s) в лист '%s'", row_number, data_with_date["date"], ws.title)


def get_filled_days(ws: gspread.Worksheet, column_map: dict[str, str], header_row: int, last_day: int) -> set:
    """
    Возвращает множество номеров дней месяца, для которых в таблице УЖЕ
    стоит значение в колонке "Бюджет" (используется как индикатор "день
    уже заполнен"). Используется в backfill, чтобы не дёргать Meta API
    за дни, которые заполнять не нужно.
    """
    check_col = column_map.get("spend") or next(iter(column_map.values()))
    first_row = header_row + 1
    last_row = header_row + last_day
    values = ws.get(f"{check_col}{first_row}:{check_col}{last_row}")

    filled_days = set()
    for i, row_vals in enumerate(values):
        day_number = i + 1
        if row_vals and str(row_vals[0]).strip() != "":
            filled_days.add(day_number)
    return filled_days


def write_target_data(target: dict[str, Any], filtered_rows: list[dict[str, Any]], ws: gspread.Worksheet, target_date: datetime.date) -> None:
    """Агрегирует уже ОТФИЛЬТРОВАННЫЕ под конкретную цель строки и пишет в лист ws."""
    name = target["name"]
    column_map = target["column_map"]
    header_row = target["header_row"]

    data = aggregate_campaign_rows(filtered_rows, target.get("custom_requests_action_type"))

    if not data.get("has_data"):
        logger.info(
            "'%s': за %s нет данных (реклама не откручивалась, либо под ключевые слова "
            "ничего не подошло) — пропускаю, ячейки останутся пустыми.",
            name, target_date.strftime("%d.%m.%Y"),
        )
        return

    write_daily_row(ws, target_date, data, column_map, header_row)


def process_one_day(gclient: gspread.Client, acc: dict[str, Any], target_date: datetime.date) -> None:
    """
    Обрабатывает один день для одного кабинета: тянет сырые данные из Meta
    ОДИН РАЗ (даже если у кабинета несколько услуг/таблиц), распределяет
    строки по целям (с учётом catch-all для "остального"), затем для каждой
    цели записи находит/создаёт нужный лист и пишет агрегированные данные.
    """
    date_str = target_date.strftime("%Y-%m-%d")
    rows = fetch_campaign_rows(acc["fb_account_id"], date_str)
    targets = get_targets_for_account(acc)
    rows_per_target = split_rows_for_targets(rows, targets)

    for target, filtered_rows in zip(targets, rows_per_target):
        try:
            spreadsheet = gclient.open_by_key(target["spreadsheet_id"])
            ws = ensure_current_month_sheet(spreadsheet, target_date, target["column_map"], target["header_row"])
            write_target_data(target, filtered_rows, ws, target_date)
        except Exception as e:
            logger.error("Ошибка при обработке '%s' за %s: %s", target["name"], target_date, e)


def get_gclient() -> gspread.Client:
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(GOOGLE_SERVICE_ACCOUNT_FILE, scopes=scopes)
    return gspread.authorize(creds)


def run_daily_refresh() -> None:
    """
    Обычный режим: обновляет последние REFRESH_WINDOW_DAYS дней (включая вчера)
    для всех кабинетов. Данные Meta по конверсиям донакапливаются в течение
    нескольких дней после показа — за счёт повторной перезаписи одних и тех
    же дней это и даёт "актуализацию задним числом".
    """
    FacebookAdsApi.init(FB_APP_ID, FB_APP_SECRET, FB_ACCESS_TOKEN)
    gclient = get_gclient()

    yesterday = datetime.date.today() - datetime.timedelta(days=1)
    dates_to_refresh = [
        yesterday - datetime.timedelta(days=offset)
        for offset in range(REFRESH_WINDOW_DAYS - 1, -1, -1)
    ]

    for acc in ACCOUNTS:
        logger.info("\n=== Кабинет: %s (%s) ===", acc["name"], acc["fb_account_id"])
        for d in dates_to_refresh:
            try:
                process_one_day(gclient, acc, d)
            except Exception as e:
                logger.error("Ошибка при обработке кабинета '%s' за %s: %s", acc["name"], d, e)
            time.sleep(1)

    logger.info("Готово. Обновлено %d последних дн(я/ей) для %d кабинетов.", REFRESH_WINDOW_DAYS, len(ACCOUNTS))

def get_month_date_range(month_param: str | None = None) -> tuple[datetime.date, datetime.date]:
    """
    Возвращает (start_date, end_date) для указанного месяца.
    - Если month_param не передан: текущий месяц (с 1-го числа по вчерашний день).
    - Если передан прошлый месяц (например '07'): весь месяц (с 1-го по 31-е число).
    - Если передан текущий месяц (например '08'): с 1-го по вчерашний день.
    """
    today = datetime.date.today()
    yesterday = today - datetime.timedelta(days=1)

    if not month_param:
        start_date = today.replace(day=1)
        end_date = yesterday
        return start_date, end_date

    month_str = str(month_param).strip()
    year = today.year

    if "." in month_str:      # Формат "07.2026"
        m_part, y_part = month_str.split(".")
        month, year = int(m_part), int(y_part)
    elif "-" in month_str:    # Формат "2026-07"
        y_part, m_part = month_str.split("-")
        year, month = int(y_part), int(m_part)
    else:                     # Формат "07" или "7"
        month = int(month_str)
        # Если указанный месяц больше текущего, считаем, что речь про прошлый год
        if month > today.month:
            year = today.year - 1

    # Если месяц совпадает с текущим
    if month == today.month and year == today.year:
        start_date = datetime.date(year, month, 1)
        end_date = yesterday
    else:
        # Для прошлых месяцев берем весь месяц целиком
        start_date = datetime.date(year, month, 1)
        _, last_day = calendar.monthrange(year, month)
        end_date = datetime.date(year, month, last_day)

    return start_date, end_date


def run_backfill_month(month_param: str | None = None, account_id_param: str | None = None) -> None:
    """
    Ручной разовый прогон: заполняет ТОЛЬКО ПУСТЫЕ дни текущего месяца
    с 1-го числа по вчерашний день.

    Заполненность проверяется ОТДЕЛЬНО для каждой целевой таблицы (если у
    кабинета несколько услуг — у каждой свой прогресс заполнения). Но запрос
    к Meta API за конкретную дату делается только ОДИН РАЗ на кабинет — и
    только если эта дата ещё нужна хотя бы одной из его таблиц.

    Заполняет пустые дни за указанный месяц (по умолчанию — текущий)
    для указанного кабинета (по умолчанию — для всех).

    Запуск: python meta_ads_to_sheets.py --backfill-month
            python meta_ads_to_sheets.py --backfill-month --07 --act_id
    """
    FacebookAdsApi.init(FB_APP_ID, FB_APP_SECRET, FB_ACCESS_TOKEN)
    gclient = get_gclient()

    start_date, end_date = get_month_date_range(month_param)

    # 1. Фильтрация кабинетов (если передан account_id_param)
    target_accounts = ACCOUNTS
    if account_id_param:
        clean_id = str(account_id_param).strip()
        target_accounts = [
            acc for acc in ACCOUNTS
            if acc["fb_account_id"] == clean_id 
            or acc["fb_account_id"] == f"act_{clean_id}"
            or acc["fb_account_id"].replace("act_", "") == clean_id
        ]

        if not target_accounts:
            logger.error("Кабинет с ID '%s' не найден в accounts_config.py!", account_id_param)
            return

    # 2. Формируем список дат для заполнения
    dates_to_fill = []
    d = start_date
    while d <= end_date:
        dates_to_fill.append(d)
        d += datetime.timedelta(days=1)

    logger.info(
        "Backfill: с %s по %s (%d дн.) для %d кабинетов",
        start_date.strftime("%d.%m.%Y"),
        end_date.strftime("%d.%m.%Y"),
        len(dates_to_fill),
        len(target_accounts),
    )

    for acc in target_accounts:
        logger.info("=== Кабинет: %s (%s) ===", acc["name"], acc["fb_account_id"])
        targets = get_targets_for_account(acc)

        target_states = []
        for i, target in enumerate(targets):
            try:
                spreadsheet = gclient.open_by_key(target["spreadsheet_id"])
                ws = ensure_current_month_sheet(spreadsheet, end_date, target["column_map"], target["header_row"])
                filled_days = get_filled_days(ws, target["column_map"], target["header_row"], end_date.day)
            except Exception as e:
                logger.error("Не удалось открыть лист для '%s': %s", target["name"], e)
                continue

            if filled_days:
                logger.info("'%s': уже заполнены дни %s — пропускаю их.", target["name"], sorted(filled_days))
            target_states.append({"target": target, "ws": ws, "filled_days": filled_days, "index": i})

        for d in dates_to_fill:
            needing = [ts for ts in target_states if d.day not in ts["filled_days"]]
            if not needing:
                continue  # все таблицы этого кабинета уже заполнены на этот день — к API не обращаемся

            try:
                rows = fetch_campaign_rows(acc["fb_account_id"], d.strftime("%Y-%m-%d"))
            except Exception as e:
                logger.error("Ошибка при запросе к Meta для кабинета '%s' за %s: %s", acc["name"], d, e)
                continue

            # Распределяем строки по ВСЕМ целям кабинета (не только по needing) —
            # иначе catch-all не будет знать, что часть строк уже "занята"
            # keyword-целями, которые сегодня не нуждались в записи.
            rows_per_target = split_rows_for_targets(rows, targets)

            for ts in needing:
                try:
                    filtered_rows = rows_per_target[ts["index"]]
                    write_target_data(ts["target"], filtered_rows, ts["ws"], d)
                except Exception as e:
                    logger.error("Ошибка при записи '%s' за %s: %s", ts["target"]["name"], d, e)

            time.sleep(1)

    logger.info("Backfill завершён для %d кабинетов.", len(target_accounts))


def main():
    parser = argparse.ArgumentParser(description="Выгрузка данных Meta Ads в Google Sheets")
    parser.add_argument("--backfill-month", action="store_true", help="Запустить режим бэкфилла")
    parser.add_argument("--month", "-m", type=str, default=None, help="Месяц (например: 07, 7, 07.2026)")
    parser.add_argument("--account", "--account-id", "-a", type=str, default=None, help="ID кабинета (например: act_12345 или 12345)")

    args, unknown = parser.parse_known_args()

    month_param = args.month
    account_id_param = args.account

    # Гибкий разбор дополнительных флагов (например, если передали --07 или act_12345 без ключей)
    for arg in unknown:
        clean_arg = arg.lstrip("-")
        if clean_arg.isdigit() and len(clean_arg) in (1, 2):
            month_param = clean_arg
        elif clean_arg.startswith("act_") or (clean_arg.isdigit() and len(clean_arg) > 5):
            account_id_param = clean_arg
    
    # Если вызван --backfill-month ИЛИ переданы месяц/аккаунт — запускаем бэкфилл
    if args.backfill_month or month_param or account_id_param:
        run_backfill_month(month_param=month_param, account_id_param=account_id_param)
    else:
        run_daily_refresh()


if __name__ == "__main__":
    main()