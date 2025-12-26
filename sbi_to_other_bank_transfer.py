from flask import Flask, render_template, request, send_file
import pandas as pd
from datetime import datetime
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Predefined account number mapping
ACCOUNT_MAP = {
    "43557115725": {"name": "SBI SCHCT", "ifsc": "07773"},
    "41256726637": {"name": "SBI SCHCT", "ifsc": "07773"},
    "34889306900": {"name": "SCHCT Gorhe", "ifsc": "07773"},
    "40237416058": {"name": "SBI SCHCT", "ifsc": "05800"}
}

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        start_serial = request.form.get("start_serial")
        account_input = request.form.get("account_input")

        if not file or not start_serial or not start_serial.isdigit() or not account_input:
            return render_template(
                "index.html",
                error="Please upload file, enter valid start serial and account number.",
                account_map=ACCOUNT_MAP
            )

        account_number = account_input.strip()
        if account_number not in ACCOUNT_MAP:
            return render_template(
                "index.html",
                error="Invalid account number entered.",
                account_map=ACCOUNT_MAP
            )

        filename = secure_filename(file.filename)
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)

        try:
            processed_file = process_excel(file_path, int(start_serial), account_number)
            return send_file(
                processed_file,
                as_attachment=True,
                download_name=os.path.basename(processed_file)
            )
        except Exception as e:
            return render_template(
                "index.html",
                error=f"Error processing file: {e}",
                account_map=ACCOUNT_MAP
            )

    return render_template("index.html", account_map=ACCOUNT_MAP)


def process_excel(file_path, start_serial, account_number):
    account_name = ACCOUNT_MAP[account_number]["name"]
    account_ifsc = ACCOUNT_MAP[account_number]["ifsc"]

    date_str = datetime.today().strftime('%d.%m.%Y')
    today_date = datetime.today().strftime('%d/%m/%y')

    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".xlsx":
        df = pd.read_excel(file_path, engine="openpyxl", dtype={"Bank-A/C": str})
    elif ext == ".xls":
        df = pd.read_excel(file_path, engine="xlrd", dtype={"Bank-A/C": str})
    else:
        raise ValueError("Unsupported file format. Please upload .xls or .xlsx")

    # 🔒 FORCE SAFE DATA TYPES
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
    df["Bank-A/C"] = df["Bank-A/C"].astype(str)
    df["IFSC"] = df["IFSC"].astype(str)
    df["Vendor"] = df["Vendor"].astype(str)

    # Filter non-SBI transfers
    other_bank_transfer = df[~df["IFSC"].str.contains("SBIN", na=False)]

    other_bank_transfer = other_bank_transfer[
        ["Bank-A/C", "IFSC", "Amount", "Vendor"]
    ].rename(columns={
        "Bank-A/C": "Account Number",
        "IFSC": "IFSC Code",
        "Vendor": "Name"
    })

    other_bank_transfer.insert(2, "Date", today_date)

    # Calculate total safely
    total_amount = float(other_bank_transfer["Amount"].sum())

    # Default first row
    default_row = pd.DataFrame([{
        "Account Number": str(account_number),
        "IFSC Code": str(account_ifsc),
        "Date": today_date,
        "Amount": total_amount,
        "Name": account_name,
        "Serial": start_serial,
        "Mode": "",
        "Formula": ""
    }])

    other_bank_transfer = pd.concat(
        [default_row, other_bank_transfer],
        ignore_index=True
    )

    # Serial number generation
    other_bank_transfer["Serial"] = range(
        start_serial,
        start_serial + len(other_bank_transfer)
    )

    # Ensure columns exist
    if "Mode" not in other_bank_transfer.columns:
        other_bank_transfer.insert(6, "Mode", "")

    if "Formula" not in other_bank_transfer.columns:
        other_bank_transfer.insert(7, "Formula", "")

    # Formula for first row
    other_bank_transfer.loc[0, "Formula"] = (
        str(account_number) + "#" +
        str(account_ifsc) + "#" +
        today_date + "#" +
        str(total_amount) + "##" +
        str(start_serial) + "#" +
        account_name + "#"
    )

    # Formula for remaining rows
    other_bank_transfer.loc[1:, "Formula"] = (
        other_bank_transfer.loc[1:, "Account Number"].astype(str) + "#" +
        other_bank_transfer.loc[1:, "IFSC Code"].astype(str) + "#" +
        other_bank_transfer.loc[1:, "Date"].astype(str) + "##" +
        other_bank_transfer.loc[1:, "Amount"].astype(str) + "#" +
        other_bank_transfer.loc[1:, "Serial"].astype(str) + "#" +
        other_bank_transfer.loc[1:, "Name"].astype(str) + "#"
    )

    output_filename = f"SCHCT_{date_str}_Other_Bank_Transfer.xlsx"
    output_path = os.path.join(UPLOAD_FOLDER, output_filename)

    other_bank_transfer.to_excel(output_path, index=False)

    return output_path


if __name__ == "__main__":
    app.run(debug=True)
