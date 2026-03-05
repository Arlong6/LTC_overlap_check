import os
import time
import pandas as pd
import rich

from reader import read_csv, make_support_class_file, make_cz_file, detect_xlsx_type
from checker import check_overlap_patient, check_overlap_worker, plot_overlap_support, plot_overlap_class


class LTC_solution:
    def __init__(self, csv_path, save_path, support_path, class_path):
        self.csv_path = csv_path
        self.save_path = save_path
        self.support_path = support_path
        self.class_path = class_path

    def run_patient_overlap_check(self, on_progress=None, on_result=None):
        df, worker_list, patient_list = read_csv(self.csv_path)
        if df is None:
            rich.print("[red]沒有資料可處理，結束檢查。[/red]")
            return

        patient_conflict_count = 0
        for i, patient in enumerate(patient_list):
            sick = df[df["patient"] == patient].sort_values(["date", "time1"])
            if check_overlap_patient(sick, self.save_path, patient) > 0:
                patient_conflict_count += 1
            print(f"目前進度：{i + 1}/{len(patient_list)} {patient}")
            if on_progress:
                on_progress("個案", i + 1, len(patient_list), patient)

        rich.print(f"[cyan]個案檢查完成：{patient_conflict_count}/{len(patient_list)} 位個案有時間衝突[/cyan]")
        if on_result:
            patient_folder = os.path.join(self.save_path, "Patient")
            on_result("patient", f"{patient_conflict_count} / {len(patient_list)} 位有衝突", patient_conflict_count, patient_folder)

        worker_conflict_count = 0
        for i, worker in enumerate(worker_list):
            sick = df[df["worker"] == worker].sort_values(["date", "time1"])
            if check_overlap_worker(sick, self.save_path, worker) > 0:
                worker_conflict_count += 1
            print(f"目前進度：{i + 1}/{len(worker_list)} {worker}")
            if on_progress:
                on_progress("居服員", i + 1, len(worker_list), worker)

        rich.print(f"[cyan]居服員檢查完成：{worker_conflict_count}/{len(worker_list)} 位居服員有時間衝突[/cyan]")
        if on_result:
            worker_folder = os.path.join(self.save_path, "Worker")
            on_result("worker", f"{worker_conflict_count} / {len(worker_list)} 位有衝突", worker_conflict_count, worker_folder)

    def run_support_check(self, on_progress=None, on_result=None):
        if not self.support_path or not os.path.isdir(self.support_path):
            return
        folders = os.listdir(self.support_path)
        if not folders:
            rich.print(f"[yellow]支援資料夾 {self.support_path} 內沒有資料夾，將不會進行支援檢查[/yellow]")
            return

        for folder_name in folders:
            folder = os.path.join(self.support_path, folder_name)
            if not os.path.isdir(folder):
                continue
            files = os.listdir(folder)
            if not files:
                rich.print(f"[red]資料夾 {folder} 內沒有檔案，請檢查路徑[/red]")
                continue

            support_path = None
            cz_path = None
            for file in files:
                if file.startswith("~$") or not file.endswith(".xlsx"):
                    continue
                fpath = os.path.join(folder, file)
                kind = detect_xlsx_type(fpath)
                if kind == "schedule" and support_path is None:
                    support_path = fpath
                elif kind == "cz" and cz_path is None:
                    cz_path = fpath

            if not support_path or not cz_path:
                rich.print(f"[red]資料夾 {folder} 缺少排班表或單頁資料，請確認檔案是否放入[/red]")
                continue

            df_support, companys_support, _ = make_support_class_file(support_path)
            df_cz = make_cz_file(cz_path)

            out_of_range_rows = []
            workers_all = []
            for company in companys_support:
                support = df_support[df_support["supported"] == company]
                cz = df_cz[df_cz["supported"] == company]
                if len(cz) == 0:
                    rich.print(f"[yellow]cz裡面沒有 {company} 的資料[/yellow]")
                    continue
                workers = list(set(support["worker"]))
                workers_all.extend(workers)
                for i, worker in enumerate(workers):
                    support1 = support[support["worker"] == worker]
                    cz1 = cz[cz["worker"] == worker]
                    if len(support1) == 0:
                        rich.print(f"[yellow]支援資料沒有 {worker} 的資訊[/yellow]")
                    elif len(cz1) == 0:
                        rich.print(f"[yellow]單頁資料沒有 {worker} 的資訊[/yellow]")
                    else:
                        out_of_range_rows.append(plot_overlap_support(support1, cz1))
                    print(f"目前進度：{i + 1}/{len(workers)} {worker}")
                    if on_progress:
                        on_progress("支援", i + 1, len(workers), worker)

            non_empty = [r for r in out_of_range_rows if not r.empty]
            out_of_range_df_total = pd.concat(non_empty) if non_empty else pd.DataFrame()
            save_path = os.path.join(self.save_path, "support", folder_name)
            if not os.path.exists(save_path):
                os.makedirs(save_path)
            report_path = os.path.join(save_path, "output_support.txt")
            out_of_range_df_total.to_csv(report_path, sep="\t", index=False)
            rich.print(f"[green]已儲存 {report_path}[/green]")
            rich.print(f"[cyan]支援檢查完成 [{folder_name}]：共 {len(out_of_range_df_total)} 筆超出支援範圍[/cyan]")
            if on_result:
                on_result("support", folder_name, len(out_of_range_df_total), report_path)

    def run_class_check(self, on_progress=None, on_result=None):
        if not self.class_path or not os.path.isdir(self.class_path):
            return
        folders = os.listdir(self.class_path)
        if not folders:
            rich.print(f"[yellow]上課資料夾 {self.class_path} 內沒有資料夾，將不會進行上課檢查[/yellow]")
            return

        for folder_name in folders:
            folder = os.path.join(self.class_path, folder_name)
            if not os.path.isdir(folder):
                continue
            files = os.listdir(folder)
            if not files:
                rich.print(f"[red]資料夾 {folder} 內沒有檔案，請檢查路徑[/red]")
                continue

            name_list = []
            class_path = None
            cz_path = None
            for file in files:
                if file.startswith("~$"):
                    continue
                fpath = os.path.join(folder, file)
                if file.endswith(".txt"):
                    with open(fpath, "r", encoding="UTF-8") as f:
                        name_list = [line.strip() for line in f]
                elif file.endswith(".xlsx"):
                    kind = detect_xlsx_type(fpath)
                    if kind == "schedule" and class_path is None:
                        class_path = fpath
                    elif kind == "cz" and cz_path is None:
                        cz_path = fpath

            if not class_path or not cz_path:
                rich.print(f"[red]資料夾 {folder} 缺少排班表或單頁資料，請確認檔案是否放入[/red]")
                continue

            df_class, companys_class, _ = make_support_class_file(class_path)
            df_cz = make_cz_file(cz_path)

            out_of_range_rows = []
            for company in companys_class:
                classes = df_class[df_class["supported"] == company]
                cz = df_cz[df_cz["supported"] == company]
                if len(cz) == 0:
                    rich.print(f"[yellow]cz裡面沒有 {company} 的資料[/yellow]")
                    continue
                workers = list(set(classes["worker"]))
                for i, worker in enumerate(workers):
                    if worker not in name_list:
                        continue
                    class1 = classes[classes["worker"] == worker]
                    cz1 = cz[cz["worker"] == worker]
                    if len(class1) == 0:
                        rich.print(f"[yellow]上課資料沒有 {worker} 的資訊[/yellow]")
                    elif len(cz1) == 0:
                        rich.print(f"[yellow]單頁資料沒有 {worker} 的資訊[/yellow]")
                    else:
                        out_of_range_rows.append(plot_overlap_class(class1, cz1))
                    print(f"目前進度：{i + 1}/{len(workers)} {worker}")
                    if on_progress:
                        on_progress("上課", i + 1, len(workers), worker)

            non_empty = [r for r in out_of_range_rows if not r.empty]
            out_of_range_df_total = pd.concat(non_empty) if non_empty else pd.DataFrame()
            save_path = os.path.join(self.save_path, "class", folder_name)
            if not os.path.exists(save_path):
                os.makedirs(save_path)
            report_path = os.path.join(save_path, "output_class.txt")
            out_of_range_df_total.to_csv(report_path, sep="\t", index=False)
            rich.print(f"[green]已儲存 {report_path}[/green]")
            rich.print(f"[cyan]上課檢查完成 [{folder_name}]：共 {len(out_of_range_df_total)} 筆上課期間有服務紀錄[/cyan]")
            if on_result:
                on_result("class", folder_name, len(out_of_range_df_total), report_path)

    def run(self, on_progress=None, on_result=None):
        self.run_patient_overlap_check(on_progress, on_result)
        self.run_support_check(on_progress, on_result)
        self.run_class_check(on_progress, on_result)


if __name__ == "__main__":
    import sys
    if getattr(sys, "frozen", False):
        _base = os.path.dirname(sys.executable)   # 打包後：exe 所在目錄
    else:
        _base = os.path.dirname(os.path.abspath(__file__))  # 原始 .py
    csv_path = os.path.join(_base, "csv")
    save_path = os.path.join(_base, "save")
    support_path = os.path.join(_base, "support")
    class_path = os.path.join(_base, "class")

    for path in [csv_path, save_path, support_path, class_path]:
        if not os.path.exists(path):
            os.makedirs(path)
            rich.print(f"[yellow]路徑：{path} 不存在，請檢查路徑或建立資料夾，已自動新增，但裡面沒有檔案[/yellow]")

    ltc = LTC_solution(csv_path, save_path, support_path, class_path)
    start_time = time.time()
    ltc.run()
    end_time = time.time()
    print(f"總共花費時間: {end_time - start_time} 秒")
