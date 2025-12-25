from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
from io import BytesIO

app = Flask(__name__)

DATA = {}

# ---------------- STEP 1 : UPLOAD ----------------
@app.route("/", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        DATA.clear()

        DATA["pob_df"] = pd.read_excel(request.files["pob"])
        DATA["portal_df"] = pd.read_excel(request.files["portal"])

        DATA["pob_cols"] = DATA["pob_df"].columns.tolist()
        DATA["portal_cols"] = DATA["portal_df"].columns.tolist()

        return redirect(url_for("select_columns"))

    return render_template("upload.html")


# ---------------- STEP 2 : COLUMN SELECTION ----------------
@app.route("/select_columns", methods=["GET", "POST"])
def select_columns():
    if request.method == "POST":
        DATA["pob_ned"] = request.form["pob_ned"]
        DATA["pob_name"] = request.form["pob_name"]
        DATA["portal_ned"] = request.form["portal_ned"]
        DATA["portal_name"] = request.form["portal_name"]

        return redirect(url_for("check_duplicates"))

    return render_template(
        "column_select.html",
        pob_cols=DATA["pob_cols"],
        portal_cols=DATA["portal_cols"]
    )


# ---------------- STEP 3 : DUPLICATE CHECK ----------------
@app.route("/check_duplicates")
def check_duplicates():
    pob_dup = DATA["pob_df"][DATA["pob_ned"]].duplicated().any()
    portal_dup = DATA["portal_df"][DATA["portal_ned"]].duplicated().any()

    if pob_dup or portal_dup:
        return render_template("duplicate_warning.html")

    return redirect(url_for("user_inputs"))


@app.route("/duplicate_decision", methods=["POST"])
def duplicate_decision():
    if request.form["decision"] == "reupload":
        return redirect(url_for("upload"))
    return redirect(url_for("user_inputs"))


# ---------------- STEP 4 : USER INPUTS ----------------
@app.route("/user_inputs", methods=["GET", "POST"])
def user_inputs():
    if request.method == "POST":
        DATA["inputs"] = request.form.to_dict()
        return redirect(url_for("generate"))

    return render_template("user_inputs.html")


# ---------------- STEP 5 : PROCESS DATA ----------------
@app.route("/generate")
def generate():
    pob = DATA["pob_df"]
    portal = DATA["portal_df"]

    # Missing logic
    missing_in_portal = pob[~pob[DATA["pob_ned"]].isin(portal[DATA["portal_ned"]])]
    missing_in_pob = portal[~portal[DATA["portal_ned"]].isin(pob[DATA["pob_ned"]])]

    # Save counts
    DATA["manifest_count"] = len(missing_in_portal)
    DATA["return_count"] = len(missing_in_pob)

    # RFM
    DATA["rfm"] = pd.DataFrame({
        "Passenger Category": DATA["inputs"]["rfm_category"],
        "NED Pass No.": missing_in_portal[DATA["pob_ned"]],
        "Travelling Vendor Code": DATA["inputs"]["vendor_code"],
        "Vendor Name": DATA["inputs"]["vendor_name"],
        "Vendor Employee Name": missing_in_portal[DATA["pob_name"]],
        "Gender": DATA["inputs"]["gender"],
        "Designation": "",
        "Originating Point": DATA["inputs"]["rfm_origin"],
        "Destination Point": DATA["inputs"]["rfm_destination"],
        "": "",
        "Charge": DATA["inputs"]["charge"]
    })

    # Manifest
    DATA["manifest"] = pd.DataFrame({
        "Passenger Weight": DATA["inputs"]["passenger_weight"],
        "Baggage Weight": DATA["inputs"]["baggage_weight"],
        "": "",
        " ": "",
        "  ": "",
        "Time Reported": DATA["inputs"]["time_reported"]
    }, index=DATA["rfm"].index)

    # Return Manifest
    DATA["return_manifest"] = pd.DataFrame({
        "Passenger Category": DATA["inputs"]["return_category"],
        "Smart Card No.": missing_in_pob[DATA["portal_ned"]],
        "Supplier": DATA["inputs"]["supplier"],
        "Vendor Employee Name": missing_in_pob[DATA["portal_name"]],
        "Gender": DATA["inputs"]["gender"],
        "Designation": "",
        "": "",
        "Charge": DATA["inputs"]["charge"],
        "Pax wt.": DATA["inputs"]["passenger_weight"],
        "Baggage": DATA["inputs"]["baggage_weight"],
        "Originating Point": DATA["inputs"]["return_origin"],
        "Destination Point": DATA["inputs"]["return_destination"]
    })

    return render_template(
        "result.html",
        manifest_count=DATA["manifest_count"],
        return_count=DATA["return_count"]
    )


# ---------------- STEP 6 : DOWNLOAD ----------------
@app.route("/download")
def download():
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        DATA["rfm"].to_excel(writer, index=False, sheet_name="RFM")
        DATA["manifest"].to_excel(writer, index=False, sheet_name="Manifest")
        DATA["return_manifest"].to_excel(writer, index=False, sheet_name="Return Manifest")

    output.seek(0)
    return send_file(output, download_name="Final_Output.xlsx", as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
