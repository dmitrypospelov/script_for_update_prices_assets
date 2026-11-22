import re
import shutil
import requests

from pathlib import Path
from datetime import datetime
from urllib.parse import quote

from openpyxl import load_workbook
from tkinter import Tk, filedialog


# ============================================================
# НАСТРОЙКИ
# ============================================================

MOEX_URL = "https://iss.moex.com/iss"
COINGECKO_URL = "https://api.coingecko.com/api/v3"

SHEET_NAME = "Data"
DASHBOARD_SHEET_NAME = "Dashboard"

# Названия инструментов находятся в столбце A
INSTRUMENT_COLUMN = 1

# Актуальные цены записываются только в столбец D
PRICE_COLUMN = 4

# Ячейка для курса USD на листе Dashboard
USD_RATE_CELL = "C6"

# Ячейка для даты обновления на листе Dashboard (текстовый формат)
UPDATE_DATE_CELL = "C1"
UPDATE_DATE_FORMAT = "%d.%m.%Y %H:%M:%S"

REQUEST_TIMEOUT = 30

# Создание резервной копии.
# При True рядом с исходным файлом создаётся копия.
# При False скрипт перезаписывает ТОЛЬКО исходный файл.
CREATE_BACKUP = False

# КРИПТОВАЛЮТЫ — ТОЛЬКО В РУБЛЯХ
CRYPTO_CURRENCY = "rub"


# ============================================================
# HTTP-СЕССИЯ
# ============================================================

session = requests.Session()

session.headers.update({
    "User-Agent": "InvestmentPortfolioUpdater/3.0"
})


# ============================================================
# СПРАВОЧНИК ОФЗ
# ============================================================

KNOWN_OFZ = {
    "26253": "SU26253RMFS3",
}


# ============================================================
# СПРАВОЧНИК КРИПТОВАЛЮТ
# ============================================================

CRYPTO_ALIASES = {
    "TON": "the-open-network",
    "TONCOIN": "the-open-network",

    "BITCOIN": "bitcoin",
    "BTC": "bitcoin",

    "ETHEREUM": "ethereum",
    "ETH": "ethereum",

    "SOLANA": "solana",
    "SOL": "solana",

    "XRP": "ripple",
    "RIPPLE": "ripple",

    "BNB": "binancecoin",

    "DOGE": "dogecoin",
    "DOGECOIN": "dogecoin",

    "ADA": "cardano",
    "CARDANO": "cardano",

    "TRX": "tron",
    "TRON": "tron",

    "LTC": "litecoin",
    "DOT": "polkadot",

    "AVAX": "avalanche-2",
    "LINK": "chainlink",

    "SHIB": "shiba-inu",
    "XLM": "stellar",

    "BCH": "bitcoin-cash",
    "ETC": "ethereum-classic",

    "NEAR": "near",
    "SUI": "sui",
    "APT": "aptos",
    "NOT": "notcoin",

    "USDT": "tether",
    "TETHER": "tether",

    "USDC": "usd-coin",
    "DAI": "dai",

    "ATOM": "cosmos",
    "UNI": "uniswap",
    "FIL": "filecoin",
    "ICP": "internet-computer",
    "HBAR": "hedera",

    "MATIC": "matic-network",
    "POL": "polygon-ecosystem-token",

    "AAVE": "aave",
    "PEPE": "pepe",
    "FLOKI": "floki",

    "OP": "optimism",
    "ARB": "arbitrum",
    "INJ": "injective-protocol",
    "SEI": "sei-network",
    "KAS": "kaspa",

    "WIF": "dogwifcoin",
    "BONK": "bonk",
}

crypto_cache = {}


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================

def normalize_name(name):

    return re.sub(
        r"[^A-Z0-9]",
        "",
        str(name).upper().strip()
    )


def get_json(url, params=None):

    response = session.get(
        url,
        params=params,
        timeout=REQUEST_TIMEOUT
    )

    response.raise_for_status()

    return response.json()


# ============================================================
# ПОИСК КРИПТОВАЛЮТЫ
# ============================================================

def get_crypto_id(name):

    name = str(name).strip()

    if name.upper().startswith("CRYPTO:"):
        name = name[7:].strip()

    normalized = normalize_name(name)

    if normalized in crypto_cache:
        return crypto_cache[normalized]

    if normalized in CRYPTO_ALIASES:

        crypto_id = CRYPTO_ALIASES[normalized]

        crypto_cache[normalized] = crypto_id

        return crypto_id

    data = get_json(
        f"{COINGECKO_URL}/search",
        params={
            "query": name
        }
    )

    coins = data.get("coins", [])

    if not coins:
        return None

    exact_name = [
        coin for coin in coins
        if normalize_name(coin.get("name", "")) == normalized
    ]

    exact_symbol = [
        coin for coin in coins
        if normalize_name(coin.get("symbol", "")) == normalized
    ]

    candidates = exact_name or exact_symbol

    if not candidates:
        return None

    candidates.sort(
        key=lambda coin: (
            coin.get("market_cap_rank") is None,
            coin.get("market_cap_rank") or 999999
        )
    )

    crypto_id = candidates[0]["id"]

    crypto_cache[normalized] = crypto_id

    return crypto_id


# ============================================================
# ПОЛУЧЕНИЕ ЦЕНЫ КРИПТОВАЛЮТЫ В РУБЛЯХ
# ============================================================

def get_crypto_price(name):

    crypto_id = get_crypto_id(name)

    if not crypto_id:

        print(f"[КРИПТО] {name}: монета не найдена")

        return None

    data = get_json(
        f"{COINGECKO_URL}/simple/price",
        params={
            "ids": crypto_id,
            "vs_currencies": CRYPTO_CURRENCY,
            "precision": "full"
        }
    )

    coin_data = data.get(crypto_id)

    if not coin_data:
        return None

    price_rub = coin_data.get("rub")

    if price_rub is None:
        return None

    price_rub = float(price_rub)

    if price_rub <= 0:
        return None

    return price_rub


# ============================================================
# ПОЛУЧЕНИЕ ЦЕНЫ АКЦИИ MOEX
# ============================================================

def get_stock_price(ticker):

    ticker = ticker.strip().upper()

    url = (
        f"{MOEX_URL}/engines/stock/markets/shares/"
        f"boards/TQBR/securities/{quote(ticker)}.json"
    )

    data = get_json(
        url,
        params={
            "iss.meta": "off",
            "iss.only": "marketdata",
            "marketdata.columns": "SECID,LAST"
        }
    )

    marketdata = data.get("marketdata", {})

    if not marketdata.get("data"):
        return None

    row = dict(zip(
        marketdata["columns"],
        marketdata["data"][0]
    ))

    price = row.get("LAST")

    if price is None:
        return None

    return float(price)


# ============================================================
# ПОИСК ОФЗ
# ============================================================

def search_ofz(name):

    name = str(name).strip().upper()

    if re.fullmatch(r"SU[A-Z0-9]+", name):
        return name

    number = re.search(r"\b\d{5}\b", name)

    if not number:
        print(f"[ОФЗ] Не найден номер выпуска: {name}")
        return None

    bond_number = number.group(0)

    fallback = KNOWN_OFZ.get(bond_number)

    data = get_json(
        f"{MOEX_URL}/securities.json",
        params={
            "q": bond_number,
            "iss.meta": "off",
            "iss.only": "securities"
        }
    )

    block = data.get("securities", {})

    columns = block.get("columns", [])
    rows = block.get("data", [])

    candidates = []

    for row in rows:

        security = dict(zip(columns, row))

        secid = str(
            security.get("SECID") or ""
        ).upper()

        shortname = str(
            security.get("SHORTNAME") or ""
        ).upper()

        fullname = str(
            security.get("NAME") or ""
        ).upper()

        description = shortname + " " + fullname

        if not secid.startswith("SU"):
            continue

        if (
            bond_number not in description
            and bond_number not in secid
        ):
            continue

        candidates.append(secid)

    candidates = list(set(candidates))

    if len(candidates) == 1:
        return candidates[0]

    if fallback:
        print(f"[ОФЗ] Резервный тикер: {fallback}")
        return fallback

    if len(candidates) > 1:
        print(f"[ОФЗ] Несколько вариантов: {candidates}")
    else:
        print(f"[ОФЗ] Выпуск не найден: {name}")

    return None


# ============================================================
# ПОЛУЧЕНИЕ ЦЕНЫ ОФЗ В РУБЛЯХ
# ============================================================

def get_ofz_price(name):

    ticker = search_ofz(name)

    if not ticker:
        return None

    url = (
        f"{MOEX_URL}/engines/stock/markets/bonds/"
        f"boards/TQOB/securities/{quote(ticker)}.json"
    )

    data = get_json(
        url,
        params={
            "iss.meta": "off",
            "iss.only": "marketdata,securities",
            "marketdata.columns": "LAST",
            "securities.columns": "FACEVALUE"
        }
    )

    marketdata = data.get("marketdata", {})
    securities = data.get("securities", {})

    if not marketdata.get("data"):
        print(f"[ОФЗ] Нет котировки: {name}")
        return None

    market = dict(zip(
        marketdata["columns"],
        marketdata["data"][0]
    ))

    security = {}

    if securities.get("data"):

        security = dict(zip(
            securities["columns"],
            securities["data"][0]
        ))

    price_percent = market.get("LAST")
    facevalue = security.get("FACEVALUE")

    if price_percent is None or facevalue is None:
        print(f"[ОФЗ] Недостаточно данных: {name}")
        return None

    price_rub = (
        float(price_percent)
        * float(facevalue)
        / 100
    )

    return round(price_rub, 2)


# ============================================================
# ОПРЕДЕЛЕНИЕ ТИПА ИНСТРУМЕНТА
# ============================================================

def is_ofz(instrument):

    name = instrument.upper()

    return (
        "ОФЗ" in name
        or bool(
            re.search(
                r"\bSU[A-Z0-9]+\b",
                name
            )
        )
    )


def is_crypto(instrument):

    name = instrument.strip()

    if name.upper().startswith("CRYPTO:"):
        return True

    return normalize_name(name) in CRYPTO_ALIASES


# ============================================================
# ПОЛУЧЕНИЕ КУРСА ДОЛЛАРА К РУБЛЮ (ЦБ РФ)
# ============================================================

def get_usd_rate():

    url = "https://www.cbr-xml-daily.ru/daily_json.js"

    try:

        data = get_json(url)

        valute = data.get("Valute", {})

        usd = valute.get("USD")

        if not usd:
            print("[КУРС USD] В ответе API нет данных по USD")
            return None

        value = usd.get("Value")

        if value is None:
            print("[КУРС USD] Не удалось извлечь значение курса")
            return None

        return float(value)

    except Exception as error:

        print(f"[КУРС USD] Ошибка получения курса: {error}")

        return None


# ============================================================
# СОЗДАНИЕ РЕЗЕРВНОЙ КОПИИ
# ============================================================

def create_backup(path):

    path = Path(path).resolve()

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    backup_path = path.with_name(
        f"{path.stem}_backup_{timestamp}{path.suffix}"
    )

    shutil.copy2(path, backup_path)

    return backup_path


# ============================================================
# ВЫБОР EXCEL-ФАЙЛА
# ============================================================

def select_file():

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    path = filedialog.askopenfilename(
        title="Выбери инвестиционный Excel-файл",
        filetypes=[
            ("Excel files", "*.xlsx")
        ]
    )

    root.destroy()

    return path


# ============================================================
# ОСНОВНАЯ ПРОГРАММА
# ============================================================

def main():

    print("\n======================================")
    print(" ОБНОВЛЕНИЕ ИНВЕСТИЦИОННОГО ПОРТФЕЛЯ")
    print("======================================")

    print("Криптовалюты: RUB")
    print("Акции: RUB")
    print("ОФЗ: RUB")

    print("======================================\n")

    path = select_file()

    if not path:
        print("Файл не выбран.")
        return

    original_path = Path(path).resolve()

    try:

        if CREATE_BACKUP:

            backup = create_backup(original_path)

            print(f"Резервная копия: {backup}\n")

        wb = load_workbook(filename=original_path)

    except Exception as error:

        print(f"Ошибка открытия файла: {error}")
        return

    if SHEET_NAME not in wb.sheetnames:

        print(f"Ошибка: лист {SHEET_NAME} не найден.")
        return

    ws = wb[SHEET_NAME]

    updated_stocks = 0
    updated_ofz = 0
    updated_crypto = 0
    failed = 0

    print("Начинаем обновление котировок...\n")

    for row in range(2, ws.max_row + 1):

        instrument = ws.cell(
            row=row,
            column=INSTRUMENT_COLUMN
        ).value

        if instrument is None:
            continue

        instrument = str(instrument).strip()

        if not instrument:
            continue

        try:

            # ОФЗ

            if is_ofz(instrument):

                price = get_ofz_price(instrument)

                if price is None:
                    failed += 1
                    continue

                updated_ofz += 1

                print(
                    f"[ОФЗ] {instrument}: "
                    f"{price:.2f} RUB"
                )

            # КРИПТОВАЛЮТЫ

            elif is_crypto(instrument):

                price = get_crypto_price(instrument)

                if price is None:
                    failed += 1
                    continue

                updated_crypto += 1

                print(
                    f"[КРИПТО] {instrument}: "
                    f"{price:,.4f} RUB"
                )

            # АКЦИИ

            else:

                price = get_stock_price(instrument)

                if price is None:

                    print(
                        f"[АКЦИЯ] {instrument}: "
                        "котировка отсутствует"
                    )

                    failed += 1
                    continue

                updated_stocks += 1

                print(
                    f"[АКЦИЯ] {instrument}: "
                    f"{price:.4f} RUB"
                )

            # ЗАПИСЫВАЕМ ТОЛЬКО В СТОЛБЕЦ D

            cell = ws.cell(
                row=row,
                column=PRICE_COLUMN
            )

            cell.value = price
            cell.number_format = '#,##0.00 "₽"'

        except requests.RequestException as error:

            print(
                f"[ИНТЕРНЕТ] {instrument}: {error}"
            )

            failed += 1

        except Exception as error:

            print(
                f"[ОШИБКА] {instrument}: {error}"
            )

            failed += 1

    # ========================================================
    # ОБНОВЛЕНИЕ ЛИСТА DASHBOARD
    # ========================================================

    if DASHBOARD_SHEET_NAME in wb.sheetnames:

        dashboard_ws = wb[DASHBOARD_SHEET_NAME]

        # --- Курс USD в C6 ---

        usd_rate = get_usd_rate()

        if usd_rate is not None:

            dashboard_ws[USD_RATE_CELL] = usd_rate
            dashboard_ws[USD_RATE_CELL].number_format = '#,##0.00 "₽"'

            print(
                f"\n[КУРС USD] {usd_rate:.2f} RUB "
                f"записан в {DASHBOARD_SHEET_NAME}!{USD_RATE_CELL}"
            )

        else:

            print("\n[КУРС USD] Не удалось получить курс, ячейка не обновлена.")

        # --- Дата обновления в C1 (текст) ---

        update_date = datetime.now().strftime(UPDATE_DATE_FORMAT)

        dashboard_ws[UPDATE_DATE_CELL] = update_date
        dashboard_ws[UPDATE_DATE_CELL].number_format = "@"  # текстовый формат

        print(
            f"[ДАТА] {update_date} "
            f"записана в {DASHBOARD_SHEET_NAME}!{UPDATE_DATE_CELL}"
        )

    else:

        print(f"\n[КУРС USD] Лист {DASHBOARD_SHEET_NAME} не найден, курс не обновлён.")
        print(f"[ДАТА] Лист {DASHBOARD_SHEET_NAME} не найден, дата не записана.")

    # ========================================================
    # СОХРАНЕНИЕ В ТОТ ЖЕ ФАЙЛ, КОТОРЫЙ БЫЛ ОТКРЫТ
    # ========================================================

    try:

        wb.save(filename=original_path)

        print("\nИсходный файл успешно обновлён!")

    except PermissionError:

        print("\nОШИБКА: Закрой Excel и попробуй ещё раз.")
        return

    except Exception as error:

        print(f"\nОшибка сохранения: {error}")
        return

    # ========================================================
    # ИТОГОВЫЙ ОТЧЁТ
    # ========================================================

    total = (
        updated_stocks
        + updated_ofz
        + updated_crypto
    )

    print("\n======================================")
    print("              РЕЗУЛЬТАТ")
    print("======================================")

    print(f"Акции:             {updated_stocks}")
    print(f"ОФЗ:               {updated_ofz}")
    print(f"Криптовалюты:      {updated_crypto}")

    print("--------------------------------------")

    print(f"Всего обновлено:   {total}")
    print(f"Ошибок:            {failed}")

    print("--------------------------------------")

    print("Валюта криптовалют: RUB")
    print("Обновлялся только столбец D.")
    print("Курс USD обновлён на Dashboard!C6.")
    print("Дата обновления записана на Dashboard!C1.")
    print("Исходный файл перезаписан.")

    print("--------------------------------------")

    print(f"Файл: {original_path}")

    print("======================================\n")


# ============================================================
# ЗАПУСК
# ============================================================

if __name__ == "__main__":
    main()
