import os

import pandas as pd

from .auth import authorize, get_user
from .config import load_config
from .events import emit_event
from .notifications import (
    notify_data_cleaned,
    notify_pipeline_completed,
    notify_pipeline_failed,
    notify_pipeline_started,
)


def _run_pipeline() -> None:
    config = load_config()
    user = get_user("analyst1@example.com")
    authorize(user, "run_pipeline")

    input_dir = config["paths"]["input_dir"]
    output_dir = config["paths"]["output_dir"]
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    input_ext = config["files"]["input_extension"]
    cleaned_output = os.path.join(output_dir, config["files"]["cleaned_output"])
    report_output = os.path.join(output_dir, config["files"]["report_output"])

    excel_files = sorted(
        [
            f for f in os.listdir(input_dir)
            if f.endswith(input_ext)
            and not f.startswith("~$")
            and os.path.isfile(os.path.join(input_dir, f))
        ]
    )

    if not excel_files:
        print("No Excel files found.")
        raise SystemExit(0)

    emit_event(
        event_type="PIPELINE_STARTED",
        user_id=user.id,
        payload={"files_found": len(excel_files)},
    )
    notify_pipeline_started()

    cleaned_frames = []

    for file in excel_files:
        file_path = os.path.join(input_dir, file)
        try:
            df = pd.read_excel(file_path)
        except Exception as exc:
            print(f"Skipping {file}: failed to read Excel file ({exc})")
            continue

        qty_col = config["columns"]["quantity"]
        price_col = config["columns"]["unit_price"]
        status_col = config["columns"]["status"]
        category_col = config["columns"]["category"]
        region_col = config["columns"]["region"]
        order_id_col = config.get("columns", {}).get("order_id")

        required_cols = [qty_col, price_col, status_col, category_col, region_col]
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            print(f"Skipping {file}: missing columns {missing}")
            continue

        df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce")
        df[price_col] = pd.to_numeric(df[price_col], errors="coerce")

        df = df[df[qty_col] >= config["filters"]["min_quantity"]]
        df = df[df[price_col] >= config["filters"]["min_unit_price"]]

        df[status_col] = df[status_col].astype(str).str.strip()
        exclude_status = set(s.strip() for s in config["filters"]["exclude_status"])
        include_status = set(s.strip() for s in config["filters"]["include_status"])
        df = df[~df[status_col].isin(exclude_status)]
        df = df[df[status_col].isin(include_status)]

        df["TotalRevenue"] = df[qty_col] * df[price_col]

        if df.empty:
            print(f"No valid rows after cleaning for {file}")
            continue

        if order_id_col and order_id_col in df.columns:
            df = df.drop_duplicates(subset=[order_id_col])
        else:
            df = df.drop_duplicates()

        cleaned_frames.append(df)

    if not cleaned_frames:
        print("No valid data to write after processing all files.")
        raise SystemExit(0)

    cleaned_all = pd.concat(cleaned_frames, ignore_index=True)

    if os.path.exists(cleaned_output):
        try:
            existing_cleaned = pd.read_excel(cleaned_output)
            cleaned_all = pd.concat([existing_cleaned, cleaned_all], ignore_index=True)
            cleaned_all = cleaned_all[cleaned_all[qty_col] >= config["filters"]["min_quantity"]]
            cleaned_all = cleaned_all[cleaned_all[price_col] >= config["filters"]["min_unit_price"]]
        except Exception as exc:
            print(f"Warning: failed reading existing cleaned output. Error: {exc}")

    cleaned_all.to_excel(cleaned_output, index=False)
    emit_event(
        event_type="DATA_CLEANED",
        user_id=user.id,
        payload={"rows_written": len(cleaned_all), "output_file": cleaned_output},
    )
    notify_data_cleaned(cleaned_output)

    if order_id_col and order_id_col in cleaned_all.columns:
        total_orders = cleaned_all[order_id_col].nunique()
    else:
        total_orders = len(cleaned_all)

    total_revenue = cleaned_all["TotalRevenue"].sum()
    avg_order_value = total_revenue / total_orders if total_orders else 0

    overall_summary = pd.DataFrame([{
        "TotalOrders": total_orders,
        "TotalRevenue": total_revenue,
        "AverageOrderValue": avg_order_value
    }])

    category_summary = (
        cleaned_all.groupby(category_col, dropna=False)
        .agg(TotalRevenue=("TotalRevenue", "sum"), TotalQuantity=(qty_col, "sum"))
        .reset_index()
    )

    if order_id_col and order_id_col in cleaned_all.columns:
        region_summary = (
            cleaned_all.groupby(region_col, dropna=False)
            .agg(TotalRevenue=("TotalRevenue", "sum"), TotalOrders=(order_id_col, "nunique"))
            .reset_index()
        )
    else:
        region_summary = (
            cleaned_all.groupby(region_col, dropna=False)
            .agg(TotalRevenue=("TotalRevenue", "sum"), TotalOrders=(region_col, "count"))
            .reset_index()
        )

    with pd.ExcelWriter(report_output, engine="openpyxl") as writer:
        overall_summary.to_excel(writer, sheet_name="Overall_Summary", index=False)
        category_summary.to_excel(writer, sheet_name="Category_Summary", index=False)
        region_summary.to_excel(writer, sheet_name="Region_Summary", index=False)

    emit_event(
        event_type="PIPELINE_COMPLETED",
        user_id=user.id,
        payload={
            "files_processed": len(cleaned_frames),
            "total_rows": len(cleaned_all)
        },
    )

    notify_pipeline_completed(cleaned_output, report_output)

    print("Cleaned data written to:", cleaned_output)
    print("Summary report written to:", report_output)


def main() -> None:
    try:
        _run_pipeline()
    except SystemExit:
        raise
    except Exception as exc:
        emit_event(
            event_type="PIPELINE_FAILED",
            user_id="system",
            payload={"error": str(exc)},
        )
        notify_pipeline_failed(str(exc))
        raise


if __name__ == "__main__":
    main()
