import os
import rich


def save_overlap_to_txt(df, filepath, reason="Need to check"):
    df = df[df["overlap"] == True]
    if not df.empty:
        with open(filepath, "w", encoding="utf-8") as f:
            for idx, row in df.iterrows():
                f.write(f"Index: {idx}\n")
                f.write(f"服務時間: {row['service_time'][0]} ~ {row['service_time'][1]}\n")
                f.write(f"患者: {row.get('patient', 'N/A')}\n")
                f.write(f"居服員: {row.get('worker', 'N/A')}\n")
                f.write(f"公司: {row.get('company', 'N/A')}\n")
                f.write(f"地址: {row.get('address', 'N/A')}\n")
                try:
                    f.write(f"原因: {', '.join(row.get('reason_list', []))}\n")
                except Exception:
                    f.write(f"原因: {reason}\n")
                f.write(f"來源: {row.get('source_file', 'N/A')}  第 {row.get('excel_row', 'N/A')} 列\n")
                f.write("-" * 30 + "\n")


def save_worker_results(df, save_path, save_name, reasons, overlap):
    save_path = os.path.join(save_path, "Worker")
    if not os.path.exists(save_path):
        os.makedirs(save_path)

    all_reasons = set(r for sublist in reasons for r in sublist)
    reason_str = "_".join(sorted(r.replace(" ", "") for r in all_reasons)) or "no_reason"

    overlap_companies = set(
        str(df["company"].iloc[i]).replace(" ", "")
        for i in range(len(df)) if overlap[i]
    )
    company_str = "_".join(sorted(overlap_companies)) or "no_company"

    txt_path = os.path.join(save_path, f"{save_name}_worker_overlap_{reason_str}_{company_str}.txt")
    save_overlap_to_txt(df, txt_path)


def save_patient_results(df, save_path, save_name):
    save_path = os.path.join(save_path, "Patient")
    if not os.path.exists(save_path):
        os.makedirs(save_path)

    txt_path = os.path.join(save_path, f"{save_name}_patient_overlap.txt")
    save_overlap_to_txt(df, txt_path)
