import numpy as np
import pandas as pd
import rich

from writer import save_worker_results, save_patient_results


# ── 個案重疊檢查 ─────────────────────────────────────────────────────────────

def check_overlap_patient(df, save_path, save_name):
    n = len(df)
    if n > 1000:
        overlap = _sweep_line_patient_overlap(df)
    else:
        overlap = _vectorized_patient_overlap(df)
    df["overlap"] = overlap
    save_patient_results(df, save_path, save_name)
    return int(overlap.sum())


def _vectorized_patient_overlap(df):
    n = len(df)
    overlap = np.zeros(n, dtype=bool)
    start_times = np.array([t[0] for t in df["service_time"]])
    end_times = np.array([t[1] for t in df["service_time"]])
    i_idx, j_idx = np.triu_indices(n, k=1)
    overlap_mask = (start_times[i_idx] < end_times[j_idx]) & (end_times[i_idx] > start_times[j_idx])
    for idx in np.where(overlap_mask)[0]:
        i, j = i_idx[idx], j_idx[idx]
        overlap[i] = overlap[j] = True
        rich.print(f"[red]時間重疊: {df['service_time'].iloc[i]} 與 {df['service_time'].iloc[j]}[/red]")
    return overlap


def _sweep_line_patient_overlap(df):
    n = len(df)
    overlap = np.zeros(n, dtype=bool)
    events = []
    for i, (start, end) in enumerate(df["service_time"]):
        events.append((start, "start", i))
        events.append((end, "end", i))
    events.sort(key=lambda x: (x[0], x[1] == "start"))
    active = set()
    for _, event_type, idx in events:
        if event_type == "start":
            for active_idx in active:
                overlap[idx] = overlap[active_idx] = True
                rich.print(f"[red]時間重疊: {df['service_time'].iloc[idx]} 與 {df['service_time'].iloc[active_idx]}[/red]")
            active.add(idx)
        else:
            active.discard(idx)
    return overlap


# ── 居服員重疊檢查 ───────────────────────────────────────────────────────────

def check_overlap_worker(df, save_path, save_name):
    n = len(df)
    overlap = np.zeros(n, dtype=bool)
    reasons = [[] for _ in range(n)]
    if n > 1000:
        overlap, reasons = _sweep_line_overlap_check(df, overlap, reasons)
    else:
        overlap, reasons = _vectorized_overlap_check(df, overlap, reasons)
    df["overlap"] = overlap
    df["reason_list"] = reasons
    save_worker_results(df, save_path, save_name, reasons, overlap)
    return int(overlap.sum())


def _vectorized_overlap_check(df, overlap, reasons):
    n = len(df)
    start_times = np.array([t[0] for t in df["service_time"]])
    end_times = np.array([t[1] for t in df["service_time"]])
    companies = df["company"].astype(str).values
    patients = df["patient"].astype(str).values
    addresses = df["address"].values
    i_idx, j_idx = np.triu_indices(n, k=1)

    # 時間重疊
    time_overlap_mask = (start_times[i_idx] < end_times[j_idx]) & (end_times[i_idx] > start_times[j_idx])
    for idx in np.where(time_overlap_mask)[0]:
        i, j = i_idx[idx], j_idx[idx]
        overlap[i] = overlap[j] = True
        reasons[i].append("時間重疊")
        reasons[j].append("時間重疊")
        rich.print(f"[red]時間重疊: {df['service_time'].iloc[i]} 與 {df['service_time'].iloc[j]}[/red]")

    # 相鄰時間衝突
    adjacent_mask = (end_times[i_idx] == start_times[j_idx])
    diff_address_mask = (addresses[i_idx] != addresses[j_idx])

    for idx in np.where(adjacent_mask & (companies[i_idx] != companies[j_idx]) & diff_address_mask)[0]:
        i, j = i_idx[idx], j_idx[idx]
        overlap[i] = overlap[j] = True
        reasons[i].append("不同公司")
        reasons[j].append("不同公司")
        rich.print(f"[red]不同公司: {df['service_time'].iloc[i]} 與 {df['service_time'].iloc[j]} (地址: {addresses[i]} vs {addresses[j]})[/red]")

    for idx in np.where(adjacent_mask & (patients[i_idx] != patients[j_idx]) & diff_address_mask)[0]:
        i, j = i_idx[idx], j_idx[idx]
        overlap[i] = overlap[j] = True
        reasons[i].append("不同個案")
        reasons[j].append("不同個案")
        rich.print(f"[red]不同個案: {df['service_time'].iloc[i]} 與 {df['service_time'].iloc[j]} (地址: {addresses[i]} vs {addresses[j]})[/red]")

    return overlap, reasons


def _sweep_line_overlap_check(df, overlap, reasons):
    events = []
    for i, (start, end) in enumerate(df["service_time"]):
        events.append((start, "start", i))
        events.append((end, "end", i))
    events.sort(key=lambda x: (x[0], x[1] == "start"))
    active = set()
    for _, event_type, idx in events:
        if event_type == "start":
            for active_idx in active.copy():
                if df["service_time"].iloc[idx][0] < df["service_time"].iloc[active_idx][1]:
                    overlap[idx] = overlap[active_idx] = True
                    reasons[idx].append("時間重疊")
                    reasons[active_idx].append("時間重疊")
                    rich.print(f"[red]時間重疊: {df['service_time'].iloc[idx]} 與 {df['service_time'].iloc[active_idx]}[/red]")
                if df["service_time"].iloc[active_idx][1] == df["service_time"].iloc[idx][0]:
                    _check_adjacent_conflicts(df, active_idx, idx, overlap, reasons)
            active.add(idx)
        else:
            active.discard(idx)
    return overlap, reasons


def _check_adjacent_conflicts(df, i, j, overlap, reasons):
    if (str(df["company"].iloc[i]) != str(df["company"].iloc[j]) and
            df["address"].iloc[i] != df["address"].iloc[j]):
        overlap[i] = overlap[j] = True
        reasons[i].append("不同公司")
        reasons[j].append("不同公司")
        rich.print(f"[red]不同公司: {df['service_time'].iloc[i]} 與 {df['service_time'].iloc[j]}[/red]")
    if (str(df["patient"].iloc[i]) != str(df["patient"].iloc[j]) and
            df["address"].iloc[i] != df["address"].iloc[j]):
        overlap[i] = overlap[j] = True
        reasons[i].append("不同個案")
        reasons[j].append("不同個案")
        rich.print(f"[red]不同個案: {df['service_time'].iloc[i]} 與 {df['service_time'].iloc[j]}[/red]")


# ── 支援 / 上課 重疊檢查 ─────────────────────────────────────────────────────

def plot_overlap_support(support_data, origin_data):
    data = pd.concat([
        support_data[["worker", "supported", "service_time", "start_time", "end_time", "type"]],
        origin_data[["worker", "supported", "service_time", "start_time", "end_time", "type"]],
    ])
    data = data.sort_values(["service_time", "type"], ascending=[True, True])
    date_list = sorted(set(data["service_time"].apply(lambda x: x[0].date())))

    out_of_range_rows = []
    for date in date_list:
        df = data[data["service_time"].apply(lambda x: x[0].date() == date)].copy()
        df["start_time"] = pd.to_datetime(df["service_time"].apply(lambda x: x[0]))
        df["end_time"] = pd.to_datetime(df["service_time"].apply(lambda x: x[1]))
        for _, row in df.iterrows():
            if row["type"] == "cz":
                out_of_range = True
                for _, support_row in df.iterrows():
                    if support_row["type"] == "support":
                        if (support_row["start_time"] <= row["start_time"] <= support_row["end_time"]
                                and support_row["start_time"] <= row["end_time"] <= support_row["end_time"]):
                            out_of_range = False
                            break
                if out_of_range:
                    out_of_range_rows.append(row)

    if out_of_range_rows:
        out_of_range_df = pd.DataFrame(out_of_range_rows)
    else:
        out_of_range_df = pd.DataFrame(columns=data.columns)
    return out_of_range_df[["worker", "supported", "start_time", "end_time"]]


def plot_overlap_class(class_data, origin_data):
    data = pd.concat([
        class_data[["worker", "supported", "service_time", "start_time", "end_time", "type"]],
        origin_data[["worker", "supported", "service_time", "start_time", "end_time", "type"]],
    ])
    data = data.sort_values(["service_time", "type"], ascending=[True, True])
    date_set = sorted(set(data["service_time"].apply(lambda x: x[0].date())))

    out_of_range_rows = []
    for date in date_set:
        df = data[data["service_time"].apply(lambda x: x[0].date() == date)].copy()
        df["start_time"] = pd.to_datetime(df["service_time"].apply(lambda x: x[0]))
        df["end_time"] = pd.to_datetime(df["service_time"].apply(lambda x: x[1]))
        for _, row in df.iterrows():
            if row["type"] == "cz":
                out_of_range = False
                for _, class_row in df.iterrows():
                    if class_row["type"] == "support":
                        if (class_row["start_time"] <= row["start_time"] <= class_row["end_time"]
                                or class_row["start_time"] <= row["end_time"] <= class_row["end_time"]):
                            out_of_range = True
                            break
                        elif row["start_time"] == class_row["end_time"]:
                            out_of_range = True
                            break
                        elif row["end_time"] == class_row["start_time"]:
                            out_of_range = True
                            break
                if out_of_range:
                    out_of_range_rows.append(row)

    if out_of_range_rows:
        out_of_range_df = pd.DataFrame(out_of_range_rows)
    else:
        out_of_range_df = pd.DataFrame(columns=data.columns)
    return out_of_range_df[["worker", "supported", "start_time", "end_time"]]
