from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "secretkey"  # For flash messages
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload_files():
    try:
        # Check if files are uploaded
        if "old_file" not in request.files or "new_file" not in request.files:
            flash("Please upload both old and new payroll files.", "error")
            return redirect(url_for("index"))

        old_file = request.files["old_file"]
        new_file = request.files["new_file"]

        # Validate file names
        if old_file.filename == "" or new_file.filename == "":
            flash("No file selected.", "error")
            return redirect(url_for("index"))

        # Save uploaded files
        old_file_path = os.path.join(app.config["UPLOAD_FOLDER"],
                                     secure_filename(old_file.filename))
        new_file_path = os.path.join(app.config["UPLOAD_FOLDER"],
                                     secure_filename(new_file.filename))
        old_file.save(old_file_path)
        new_file.save(new_file_path)

        # Process files
        old_data = pd.read_excel(old_file_path)
        new_data = pd.read_excel(new_file_path)

        # Ensure required columns are present
        required_columns = [
            "Employee Number", "Net Salary", "Employee Name", "Supplier ID",
            "Gross Salary", "Deductions"
        ]
        for col in required_columns:
            if col not in old_data.columns or col not in new_data.columns:
                flash(f"Missing required column: {col}", "error")
                return redirect(url_for("index"))

        # Merge the data on "Employee Number"
        merged_data = pd.merge(old_data,
                               new_data,
                               on="Employee Number",
                               suffixes=("_old", "_new"))

        # Calculate changes in Net Salary
        merged_data["Increase/Decrease"] = merged_data[
            "Net Salary_new"] - merged_data["Net Salary_old"]
        merged_data["Difference"] = merged_data["Increase/Decrease"].apply(
            lambda x: "Increase" if x > 0 else "Decrease"
            if x < 0 else "No Change")

        # Select and organize columns for the output
        output_columns = [
            "Employee Number",
            "Employee Name_new",
            "Supplier ID_new",
            "Gross Salary_new",
            "Deductions_new",
            "Net Salary_old",
            "Net Salary_new",
            "Increase/Decrease",
            "Difference",
        ]
        report_data = merged_data[output_columns]
        report_data.rename(
            columns=lambda x: x.replace("_new", "").replace("_old", " (Old)"),
            inplace=True)

        # Save the output to an Excel file
        output_file_path = os.path.join(app.config["UPLOAD_FOLDER"],
                                        "Payroll_Comparison_Report.xlsx")
        report_data.to_excel(output_file_path, index=False)

        flash("Report generated successfully! Click below to download.",
              "success")
        return redirect(
            url_for("download_file",
                    filename="Payroll_Comparison_Report.xlsx"))

    except Exception as e:
        flash(str(e), "error")
        return redirect(url_for("index"))


@app.route("/download/<filename>")
def download_file(filename):
    try:
        file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        if os.path.exists(file_path):
            response = send_file(file_path, as_attachment=True)
            # Clean up all files after download
            for file in os.listdir(app.config["UPLOAD_FOLDER"]):
                os.remove(os.path.join(app.config["UPLOAD_FOLDER"], file))
            return response
        else:
            flash("File not found. Please process files first.", "error")
            return redirect(url_for("index"))
    except Exception as e:
        flash(str(e), "error")
        return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True, host = '0.0.0.0', port = 6565)
