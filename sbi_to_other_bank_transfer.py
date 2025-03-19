from flask import Flask, render_template, request, send_file
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        start_serial = request.form.get("start_serial")

        if file and start_serial.isdigit():
            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)

            try:
                processed_file = process_excel(file_path, int(start_serial))
                return send_file(processed_file, as_attachment=True)
            except Exception as e:
                return f"Error processing file: {e}"

    return render_template("index.html")

def process_excel(file_path, start_serial):
    date_str = datetime.today().strftime('%d.%m.%Y')
    today_date = datetime.today().strftime('%d/%m/%y')

    other_bank_transfer = pd.read_excel(file_path, engine="xlrd", dtype={'Bank-A/C': str})
    other_bank_transfer = other_bank_transfer[~other_bank_transfer['IFSC'].str.contains('SBIN', na=False)]
    other_bank_transfer = other_bank_transfer[['Bank-A/C', 'IFSC', 'Amount', 'Vendor']]
    other_bank_transfer = other_bank_transfer.rename(columns={'Bank-A/C': 'Account Number', 'IFSC': 'IFSC Code', 'Vendor': 'Name'})
    other_bank_transfer.insert(2, 'Date', today_date)

    total_amount = other_bank_transfer["Amount"].sum()

    default_row = pd.DataFrame([{
        "Account Number": "41256726637",
        "IFSC Code": "07773",
        "Date": today_date,
        "Amount": total_amount,
        "Name": "SCST GORHE",
        "Serial": start_serial,
        "Mode": "",
        "Formula": ""
    }])

    other_bank_transfer = pd.concat([default_row, other_bank_transfer], ignore_index=True)
    other_bank_transfer["Serial"] = range(start_serial, start_serial + len(other_bank_transfer))

    if "Mode" not in other_bank_transfer.columns:
        other_bank_transfer.insert(6, 'Mode', '')

    if "Formula" not in other_bank_transfer.columns:
        other_bank_transfer.insert(7, 'Formula', '')

    other_bank_transfer.loc[0, "Formula"] = f"{other_bank_transfer.loc[0, 'Account Number']}#" \
                                            f"{other_bank_transfer.loc[0, 'IFSC Code']}#" \
                                            f"{other_bank_transfer.loc[0, 'Date']}#" \
                                            f"{other_bank_transfer.loc[0, 'Amount']}##" \
                                            f"{other_bank_transfer.loc[0, 'Serial']}#" \
                                            f"{other_bank_transfer.loc[0, 'Name']}#"

    other_bank_transfer.loc[1:, "Formula"] = other_bank_transfer.loc[1:, "Account Number"] + "#" + \
                                            other_bank_transfer.loc[1:, "IFSC Code"] + "#" + \
                                            other_bank_transfer.loc[1:, "Date"] + "##" + \
                                            other_bank_transfer.loc[1:, "Amount"].astype(str) + "#" + \
                                            other_bank_transfer.loc[1:, "Serial"].astype(str) + "#" + \
                                            other_bank_transfer.loc[1:, "Name"] + "#"

    output_filename = f"SCHCT_{date_str}_Other_Bank_Transfer.xlsx"
    output_path = os.path.join(UPLOAD_FOLDER, output_filename)
    other_bank_transfer.to_excel(output_path, index=False)

    return output_path

if __name__ == "__main__":
    app.run(debug=True)
