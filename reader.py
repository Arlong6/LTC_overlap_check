import os
import re
import pandas as pd
from datetime import datetime
from pathlib import Path
import rich


def fullwidth_to_halfwidth(s):
    result = ""
    for char in s:
        code = ord(char)
        if code == 0x3000:
            code = 0x0020
        elif 0xFF01 <= code <= 0xFF5E:
            code -= 0xFEE0
        result += chr(code)
    return result


def parse_date_time(date_val, time_str):
    if isinstance(date_val, datetime):
        date_str = date_val.strftime("%Y-%m-%d")
    else:
        date_str = str(date_val).strip()
    time_str = time_str.strip()
    dt_str = f"{date_str} {time_str}"
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(dt_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"無法解析時間：'{dt_str}'（原始時間字串: '{time_str}'）")


def read_csv(csv_path):
    tables = []
    folders = os.listdir(csv_path)
    if not folders:
        rich.print(f"[yellow]CSV資料夾 {csv_path} 內沒有資料夾，將不會進行檢查[/yellow]")
        return None, None, None

    for path in folders:
        if not path.endswith(".xlsx"):
            continue
        path = os.path.join(csv_path, path)
        rich.print(f"[bold blue]正在讀取 {path}[/bold blue]")
        data = pd.read_excel(path, sheet_name=None)
        for table_name in data.keys():
            table = data[table_name]
            table = table.iloc[3:, [0, 5, 17, 18, 23, 25, 46]]
            table.columns = ["id", "patient", "company", "worker", "date", "time1", "address"]
            table = table.dropna(subset=["date", "time1"])
            table["source_file"] = os.path.basename(path)
            table["excel_row"] = table.index + 2
            table["address"] = table["address"].apply(fullwidth_to_halfwidth)
            if ":" in str(table["date"].iloc[0]):
                table["date"], table["time1"] = table["time1"], table["date"]
            tables.append(table)

    print("合併完成，開始分析")
    df = pd.concat(tables)
    df["service_time"] = df.apply(
        lambda x: (
            parse_date_time(x["date"], x["time1"].split("~")[0].strip()),
            parse_date_time(x["date"], x["time1"].split("~")[1].strip()),
        ),
        axis=1,
    )
    worker_list = list(set(df["worker"]))
    patient_list = list(set(df["patient"]))
    return df, worker_list, patient_list


def support_data_reconstruct(df):
    def convert_roc_to_ad(date_str):
        def repl(match):
            roc_year = int(match.group(1))
            return str(roc_year + 1911)
        return re.sub(r"^(\d{1,3})([-/])", lambda m: f"{repl(m)}{m.group(2)}", date_str)

    for col in ["start_time", "end_time"]:
        df[col] = df[col].str.replace("/", "-", regex=False)
        df[col] = df[col].apply(convert_roc_to_ad)
    df["end_time"] = df["end_time"].str.replace("24:00:00", "23:59:59", n=1)
    df["start_time"] = pd.to_datetime(df["start_time"], format="%Y-%m-%d %H:%M:%S")
    df["end_time"] = pd.to_datetime(df["end_time"], format="%Y-%m-%d %H:%M:%S")
    df["service_time"] = list(zip(df["start_time"], df["end_time"]))
    df = df.sort_values(["worker", "start_time"])
    return df


def detect_xlsx_type(path):
    """
    判斷 xlsx 是哪種類型，不依賴檔名。
    - 'cz'       : 單頁服務紀錄（欄數 >= 20）
    - 'schedule' : 支援 / 上課排班表（欄數 < 20）
    - 'unknown'  : 無法判斷
    """
    try:
        df = pd.read_excel(path, nrows=3)
        return "cz" if len(df.columns) >= 20 else "schedule"
    except Exception:
        return "unknown"


def make_support_class_file(support_path):
    rich.print(f"[blue]正在讀取支援資料：{support_path}[/blue]")
    data_support = pd.read_excel(Path(support_path), sheet_name=None)
    tables = []
    for table_name in data_support.keys():
        table = data_support[table_name]
        table = table.iloc[0:, [0, 1, 2, 4, 5]]
        table.columns = ["Support", "worker", "supported", "start_time", "end_time"]
        tables.append(table)
    df_support = pd.concat(tables)
    df_support = support_data_reconstruct(df_support)
    companys_support = list(set(df_support["supported"]))
    workers_support = list(set(df_support["worker"]))
    df_support["type"] = ["support"] * len(df_support)
    return df_support, companys_support, workers_support


def make_cz_file(cz_path):
    rich.print(f"[blue]正在讀取單頁資料：{cz_path}[/blue]")
    dataCZ = pd.read_excel(Path(cz_path), sheet_name=None)
    tables = []
    for table_name in dataCZ.keys():
        tableCZ = dataCZ[table_name]
        tableCZ = tableCZ.iloc[3:, [0, 5, 17, 18, 23, 25]]
        tableCZ.columns = ["id", "patient", "supported", "worker", "date", "time1"]
        tableCZ = tableCZ.dropna(subset=["date", "time1"])
        tableCZ["source_file"] = os.path.basename(cz_path)
        tableCZ["excel_row"] = tableCZ.index + 2
        if ":" in str(tableCZ["date"].iloc[0]):
            tableCZ["date"], tableCZ["time1"] = tableCZ["time1"], tableCZ["date"]
        tables.append(tableCZ)
    dfCZ = pd.concat(tables)

    def parse_time_row(x):
        date_str = (
            x["date"].strftime("%Y-%m-%d")
            if isinstance(x["date"], datetime)
            else str(x["date"]).split()[0].strip()
        )
        try:
            time_parts = x["time1"].split("~")
            start_time_str = time_parts[0].strip()
            end_time_str = time_parts[1].strip()
        except Exception:
            raise ValueError(f"時間格式錯誤: {x['time1']}")
        start_time = datetime.strptime(f"{date_str} {start_time_str}", "%Y-%m-%d %H:%M")
        end_time = datetime.strptime(f"{date_str} {end_time_str}", "%Y-%m-%d %H:%M")
        return pd.Series([start_time, end_time, (start_time, end_time)])

    dfCZ[["start_time", "end_time", "service_time"]] = dfCZ.apply(parse_time_row, axis=1)
    dfCZ["type"] = "cz"
    return dfCZ
