from flask import Flask, render_template, request, send_file
import pandas as pd
from io import BytesIO
import os

app = Flask(__name__)

# ============================================================
# FILE CONFIGURATION
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXHIBITIONS = [
    {
        "name": "Bharat Packaging",
        "slug": "bharat-packaging",
        "logo": "bharat-packaging-expo-logo.jpg",
        "excel": "MASTER SHEET BHARAT PACKAGING 2026  FINAL - CLEAN -DATA-WITH-CAMPAIGN.xlsx"
    },
    {
        "name": "Money Alpha",
        "slug": "money-alpha",
        "logo": "money-alpha-logo-img.jpg",
        "excel": "Money Alpha - Full Data.xlsx"
    },
    {
        "name": "Bharat Sustainability",
        "slug": "bharat-sustainability",
        "logo": "bharat-sustainability-expo-logo.jpg",
        "excel": "Bharat Sustainability Expo Complete Data.xlsx"
    },
    {
        "name": "Engineering Industry Tech Expo",
        "slug": "engineering",
        "logo": "logo.jpg",
        "excel": "Engineering Industry Tech Expo- Full Data.xlsx"
    },
    {
        "name": "Indian Manufacturing Show",
        "slug": "indian-manufacturing-show",
        "logo": "images.jpg",
        "excel": "IMS - Full Data.xlsx"
    },
    {
        "name": "Global Bhakti Fest",
        "slug": "global-bhakti-fest",
        "logo": "global-bhakti-fest.jpg",
        "excel": "Global Bhakti Fest - Master Data.xlsx"
    }
]


# ============================================================
# FIND EXHIBITION
# ============================================================

def get_exhibition(slug):
    return next(
        (exhibition for exhibition in EXHIBITIONS
         if exhibition["slug"] == slug),
        None
    )


# ============================================================
# LOAD EXCEL DATA
# ============================================================

def load_data(exhibition):
    excel_path = os.path.join(BASE_DIR, exhibition["excel"])

    if not os.path.exists(excel_path):
        raise FileNotFoundError(
            f"Excel file not found: {exhibition['excel']}"
        )

    df = pd.read_excel(excel_path)

    # Sector is our Category field
    if "Sector" not in df.columns:
        raise ValueError(
            f"The Excel file '{exhibition['excel']}' "
            f"does not contain a 'Sector' column."
        )

    # Clean Sector values (Sector is the dashboard's Category field).
    df["Sector"] = (
        df["Sector"]
        .fillna("")
        .astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    # Older master files may not have this column yet. Treat all their
    # records as undated without changing the source Excel file.
    if "Date of Update" not in df.columns:
        df["Date of Update"] = pd.NaT
    else:
        # Empty cells and unparseable values are treated as blank/undated.
        df["Date of Update"] = pd.to_datetime(
            df["Date of Update"], errors="coerce", dayfirst=True
        ).dt.normalize()

    return df


def filter_by_update_date(df, date_filter, selected_date):
    """Apply the dashboard date filter to a copy of the master data."""
    if date_filter == "blank":
        return df[df["Date of Update"].isna()].copy()

    if date_filter == "specific":
        if not selected_date:
            return df.iloc[0:0].copy()
        parsed_date = pd.to_datetime(selected_date, errors="coerce")
        if pd.isna(parsed_date):
            return df.iloc[0:0].copy()
        parsed_date = parsed_date.normalize()
        return df[df["Date of Update"] == parsed_date].copy()

    # All Dates includes both dated records and blank/legacy records.
    return df.copy()


# ============================================================
# MAIN HOME PAGE
# ============================================================

@app.route("/")
def home():
    return render_template(
        "home.html",
        exhibitions=EXHIBITIONS
    )


# ============================================================
# EXHIBITION DASHBOARD
# ============================================================

@app.route("/dashboard/<slug>")
def dashboard(slug):

    exhibition = get_exhibition(slug)

    if exhibition is None:
        return "Exhibition not found", 404

    try:
        df = load_data(exhibition)

    except FileNotFoundError as e:
        return render_template(
            "coming_soon.html",
            show_name=exhibition["name"],
            message=str(e)
        )

    except ValueError as e:
        return render_template(
            "coming_soon.html",
            show_name=exhibition["name"],
            message=str(e)
        )

    # Date filter: all records, blank/undated records, or one specific date.
    date_filter = request.args.get("date_filter", "all").strip().lower()
    if date_filter not in {"all", "blank", "specific"}:
        date_filter = "all"
    selected_date = request.args.get("update_date", "").strip()
    filtered_df = filter_by_update_date(df, date_filter, selected_date)

    # Records having Sector, calculated only from the selected date subset.
    sector_data = filtered_df[filtered_df["Sector"] != ""]

    # Category counts
    category_counts = (
        sector_data["Sector"]
        .value_counts()
        .rename_axis("sector")
        .reset_index(name="count")
    )

    category_counts = category_counts.sort_values(
        by=["count", "sector"],
        ascending=[False, True]
    )

    total_records = len(filtered_df)

    total_categories = len(category_counts)

    records_with_sector = len(sector_data)

    records_without_sector = int(
        (df["Sector"] == "").sum()
    )

    return render_template(
        "dashboard.html",

        show_name=exhibition["name"],

        logo=exhibition["logo"],

        total=total_records,

        category_count=total_categories,

        named_total=records_with_sector,

        blank_sector=records_without_sector,

        categories=category_counts.to_dict(
            orient="records"
        ),
        date_filter=date_filter,
        selected_date=selected_date,
        undated_total=int(df["Date of Update"].isna().sum()),
        dated_total=int(df["Date of Update"].notna().sum())
    )


# ============================================================
# EXPORT EXCEL
# ============================================================

@app.route("/export/<slug>")
def export(slug):

    exhibition = get_exhibition(slug)

    if exhibition is None:
        return "Exhibition not found", 404

    try:
        df = load_data(exhibition)

    except Exception as e:
        return str(e), 500

    # Apply the same date filter used by the dashboard.
    date_filter = request.args.get("date_filter", "all").strip().lower()
    if date_filter not in {"all", "blank", "specific"}:
        date_filter = "all"
    selected_date = request.args.get("update_date", "").strip()
    df = filter_by_update_date(df, date_filter, selected_date)

    # Optional category filter, applied after the date filter.
    category = request.args.get(
        "category",
        ""
    ).strip()

    if category:

        df = df[
            df["Sector"] == category
        ]

        safe_category = "".join(
            c if c.isalnum() else "_"
            for c in category
        )

        filename = (
            f"{exhibition['slug']}_{safe_category}.xlsx"
        )

    else:

        filename = (
            f"{exhibition['slug']}_Full_Data.xlsx"
        )

    # Create Excel in memory
    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Master Data"
        )

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# ============================================================
# RUN APPLICATION
# ============================================================

if __name__ == "__main__":

    app.run(
        host="127.0.0.1",
        port=5000,
        debug=True
    )