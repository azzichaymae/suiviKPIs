import pandas as pd
from django.http import HttpResponse
from django.template.loader import render_to_string
from django.shortcuts import render
from xhtml2pdf import pisa
import io
import matplotlib.pyplot as plt
import base64
import re
from datetime import datetime

# Use non-interactive backend for matplotlib (prevents GUI thread warnings in Django)
plt.switch_backend('Agg')

def extract_date_from_filename(filename):
    match = re.search(r'(\d{4})(\d{2})(\d{2})', filename) 
    if match:
        year, month, day = match.groups()
        return datetime(int(year), int(month), int(day)).strftime("%d/%m/%Y")

    match = re.search(r'(\d{2})-(\d{2})-(\d{4})', filename) 
    if match:
        day, month, year = match.groups()
        return datetime(int(year), int(month), int(day)).strftime("%d/%m/%Y")

    return None

def parse_total_time(value):
    """Convert hh:mm string or Excel time serial into decimal hours."""
    if pd.isna(value):
        return 0
    # Excel time serial = fraction of a day (e.g. 0.3333 = 8 hours)
    if isinstance(value, (int, float)):
        if value < 1:
            return round(value * 24, 2)
        return round(value, 2)
    if isinstance(value, str):
        try:
            h, m = map(int, value.split(":"))
            return h + m/60
        except ValueError:
            try:
                val = float(value)
                if val < 1:
                    return round(val * 24, 2)
                return round(val, 2)
            except ValueError:
                return 0
    if hasattr(value, "hour") and hasattr(value, "minute"):
        return value.hour + value.minute/60
    return 0

def format_french_date(date_str):
    if not date_str:
        return ""
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    months = {
        1: "Jan", 2: "Fév", 3: "Mar", 4: "Avr", 5: "Mai", 6: "Juin",
        7: "Juil", 8: "Août", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Déc"
    }
    return f"{dt.day} {months[dt.month]} {str(dt.year)[2:]}" 

def validate_required_columns(df, file_label, required_cols, ignore_cols=None):
    if ignore_cols is None:
        ignore_cols = []

    alerts = []
    for idx, row in df.iterrows():
        for col in required_cols:
            if col not in ignore_cols:
                val = row[col] if col in df.columns else None
                if pd.isna(val) or str(val).strip() == "":
                    alerts.append(f"⚠️ {file_label} : Champ vide dans la colonne '{col}' (ligne {idx+2})")
    return alerts

def check_required_columns(df, file_label, required_cols):
    errors = []
    df_columns = [c.strip() for c in df.columns.tolist()]
    for col in required_cols:
        if col not in df_columns:
            errors.append(f"❌ {file_label} : Colonne obligatoire manquante '{col}'")
    return errors

def upload_excel(request):
    alerts = []

    reporter_stats, bog_stats, zone_stats, employee_stats = {}, {}, {}, {}
    histogram_base64 = None
    extraction_date = None
    employee_actions = {}
    grouped = {}

    # Flags for template conditional rendering
    has_vpc = False
    has_bog = False
    has_overdue = False
    has_std = False

    if request.method == "POST":
        Sdate = request.POST.get("date_from")
        Edate = request.POST.get("date_to")
        departement = request.POST.get("departement")
        vpc_file = request.FILES.get("file_vpc")
        bog_file = request.FILES.get("file_bog")
        overdue_file = request.FILES.get("overdue_file")
        std_file = request.FILES.get("std_file")

        if vpc_file and vpc_file.name:
            df_vpc = pd.read_excel(vpc_file)
            df_vpc.columns = df_vpc.columns.str.strip()

            missing = check_required_columns(df_vpc, "VPC", ["VPC rapporté par", "Type de VPC effectué"])
            if missing:
                alerts.extend(missing)
            else:
                has_vpc = True
                alerts.extend(validate_required_columns(df_vpc, "VPC", ["VPC rapporté par", "Type de VPC effectué"]))
                df_vpc["ReporterName"] = df_vpc["VPC rapporté par"].str.split(",").str[0:2].str.join(",")

                for reporter in df_vpc["ReporterName"].unique():
                    sub_df = df_vpc[df_vpc["ReporterName"] == reporter]
                    vpc_count = (sub_df["Type de VPC effectué"] == "Axé sur le sujet").sum()
                    cvpc_count = (sub_df["Type de VPC effectué"] == "Critical VPC").sum()
                    reporter_stats[reporter] = {"VPC": vpc_count, "cVPC": cvpc_count}

        if bog_file and bog_file.name:
            bog_filename = bog_file.name
            extraction_date = extract_date_from_filename(bog_filename)
            df_bog = pd.read_excel(bog_file)
            df_bog.columns = df_bog.columns.str.strip()

            missing = check_required_columns(df_bog, "BOG", ["User", "TotalTime", "Zone", "TourValidStatus"])
            if missing:
                alerts.extend(missing)
            else:
                has_bog = True
                alerts.extend(validate_required_columns(df_bog, "BOG", ["User", "TotalTime", "Zone", "TourValidStatus"]))
                df_bog = df_bog[df_bog["TourValidStatus"] == "valid"]

                df_bog["TotalHours"] = df_bog["TotalTime"].apply(parse_total_time)

                bog_stats = df_bog.groupby("User")["TotalHours"].sum().to_dict()

                df_bog["ZoneNumber"] = df_bog["Zone"].str.extract(r'(\d+)').astype(int)
                zone_stats = dict(sorted(df_bog.groupby("ZoneNumber")["TotalHours"].sum().to_dict().items()))

                plt.figure(figsize=(6,4))

                zone_labels = [f"Zone {z}" for z in zone_stats.keys()]
                zone_values = list(zone_stats.values())

                plt.bar(zone_labels, zone_values, color="#4472c4", width=0.4, edgecolor="#4472c4")

                plt.title("Nombre d\'heures passées dans chaque zone", fontsize=12)
                plt.xlabel("")
                plt.ylabel("")
                plt.xticks(rotation=45, ha="right", fontsize=10)
                plt.yticks(fontsize=10)

                plt.grid(axis="y", color="#d9d9d9", linestyle="-", linewidth=0.8)

                for spine in ["top", "right"]:
                    plt.gca().spines[spine].set_visible(False)

                plt.tight_layout()

                buf = io.BytesIO()
                plt.savefig(buf, format="png")
                buf.seek(0)
                histogram_base64 = base64.b64encode(buf.read()).decode("utf-8")
                buf.close()


        if overdue_file and overdue_file.name:
            df_overdue = pd.read_excel(overdue_file)
            df_overdue.columns = df_overdue.columns.str.strip()

            missing = check_required_columns(df_overdue, "Overdue", ["Assigned To", "Action Summary", "Priority"])
            if missing:
                alerts.extend(missing)
            else:
                has_overdue = True
                alerts.extend(validate_required_columns(df_overdue, "Overdue", ["Assigned To", "Action Summary", "Priority"]))

                df_overdue["Employee"] = df_overdue["Assigned To"].str.split(",").str[0:2].apply(
                    lambda x: " ".join([p.strip() for p in x if p.strip()])
                )

                employee_actions = {}

                for _, row in df_overdue.iterrows():
                    emp = row["Employee"]
                    summary = row["Action Summary"]
                    priority = row["Priority"]

                    if emp not in employee_actions:
                        employee_actions[emp] = []

                    employee_actions[emp].append({"summary": summary, "priority": priority})

        if std_file and std_file.name:
            df_std = pd.read_excel(std_file)
            df_std.columns = df_std.columns.str.strip()

            missing = check_required_columns(df_std, "STD", ["Rapporté par", "Description du danger", "Statut", "Actions"])
            if missing:
                alerts.extend(missing)
            else:
                has_std = True
                alerts.extend(validate_required_columns(
                    df_std,
                    "STD",
                    ["Rapporté par", "Description du danger", "Statut"],
                    ignore_cols=["Assigné_nom"]  # allowed empty
                ))

                df_std["Rapporté_par_nom"] = df_std["Rapporté par"].str.split(",").str[0:2].str.join(" ").str.strip()

                df_std["Assigné_nom"] = (
                    df_std["Actions"]
                    .fillna("")                  
                    .str.split(",")
                    .str[2:4]
                    .str.join(" ")
                    .str.strip()
                )

                grouped = {}
                for _, row in df_std.iterrows():
                    emp = row["Rapporté_par_nom"]
                    if emp not in grouped:
                        grouped[emp] = []
                    desc = row["Description du danger"]
                    if pd.isna(desc):  
                        desc = ""
                        print(f"Warning: Missing description for employee {emp} in row {_}")

                    grouped[emp].append({
                        "description": desc,
                        "assigned_to": row["Assigné_nom"],
                        "status": row["Statut"]
                    })

        if alerts:
            return render(request, "upload_form.html", {"alerts": alerts})
        else:
            html_string = render_to_string("report.html", {
                "reporter_stats": reporter_stats,
                "bog_stats": bog_stats,
                "zone_stats": zone_stats,
                "employee_stats": employee_stats,
                "employee_actions": employee_actions,
                "histogram_base64": histogram_base64,
                "extraction_date": extraction_date,
                "Sdate": format_french_date(Sdate),
                "Edate": format_french_date(Edate),
                "departement": departement,
                "grouped": grouped,
                "has_vpc": has_vpc,
                "has_bog": has_bog,
                "has_overdue": has_overdue,
                "has_std": has_std,
            })

            pdf_buffer = io.BytesIO()
            pisa.CreatePDF(html_string, dest=pdf_buffer)

            response = HttpResponse(pdf_buffer.getvalue(), content_type="application/pdf")
            response["Content-Disposition"] = "attachment; filename=suivi hebdomadaire.pdf"
            return response

    return render(request, "upload_form.html")