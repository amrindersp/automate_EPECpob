from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill

app = Flask(__name__)

DATA = {}


# --------- Utility: critical cleaning for NED columns ---------

def clean_ned_column(df, col_name):
    """
    Critical cleaning rules for NED Pass Numbers:
    - Treat strictly as text
    - Remove leading/trailing spaces
    - Remove empty / nan / none / null / na / n/a
    - Do not assume fixed NED pattern
    - Automatically ignore footer-type rows:
      * value has no digits at all
      * row lies in the last 15 rows of the file
    """
    s = df[col_name].astype(str).str.strip()

    # Normalize empties
    empty_markers = {"", "nan", "none", "null", "na", "n/a"}
    mask_empty_like = s.str.lower().isin(empty_markers)

    # Identify values that have no digits at all (pure text)
    mask_no_digits = ~s.str.contains(r"\d", regex=True)

    # Consider last 15 rows as footer zone
    last_n = 15
    footer_start_index = max(0, len(df) - last_n)
    approx_footer = df.index >= footer_start_index

    # Footer-type rows: pure text in footer area
    mask_footer_text = mask_no_digits & approx_footer

    keep_mask = ~(mask_empty_like | mask_footer_text)

    df_clean = df.loc[keep_mask].copy()
    df_clean[col_name] = s.loc[keep_mask]

    return df_clean


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

        # Apply critical cleaning to selected NED columns
        DATA["pob_df"] = clean_ned_column(DATA["pob_df"], DATA["pob_ned"])
        DATA["portal_df"] = clean_ned_column(DATA["portal_df"], DATA["portal_ned"])

        return redirect(url_for("check_duplicates"))

    return render_template(
        "column_select.html",
        pob_cols=DATA["pob_cols"],
        portal_cols=DATA["portal_cols"],
    )


# ---------------- STEP 3 : DUPLICATE CHECK ----------------

@app.route("/check_duplicates")
def check_duplicates():
    pob_col = DATA["pob_ned"]
    portal_col = DATA["portal_ned"]
    pob = DATA["pob_df"]
    portal = DATA["portal_df"]

    pob_dup_mask = pob[pob_col].duplicated(keep=False)
    portal_dup_mask = portal[portal_col].duplicated(keep=False)

    pob_dup = pob_dup_mask.any()
    portal_dup = portal_dup_mask.any()

    if pob_dup or portal_dup:
        # Build Excel file with duplicate NED cells highlighted
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pob.to_excel(writer, sheet_name="POB", index=False)
            portal.to_excel(writer, sheet_name="PORTAL", index=False)

            wb = writer.book
            pob_ws = wb["POB"]
            portal_ws = wb["PORTAL"]

            yellow_fill = PatternFill(
                start_color="FFFF00",
                end_color="FFFF00",
                fill_type="solid"
            )

            # POB duplicates
            pob_ned_idx = pob.columns.get_loc(pob_col) + 1
            for i, is_dup in enumerate(pob_dup_mask, start=2):  # row 1 header
                if is_dup:
                    pob_ws.cell(row=i, column=pob_ned_idx).fill = yellow_fill

            # PORTAL duplicates
            portal_ned_idx = portal.columns.get_loc(portal_col) + 1
            for i, is_dup in enumerate(portal_dup_mask, start=2):
                if is_dup:
                    portal_ws.cell(row=i, column=portal_ned_idx).fill = yellow_fill

        output.seek(0)
        DATA["duplicate_file"] = output

        return render_template(
            "duplicate_warning.html",
            pob_dup=pob_dup,
            portal_dup=portal_dup,
        )

    return redirect(url_for("user_inputs"))


@app.route("/duplicate_decision", methods=["POST"])
def duplicate_decision():
    if request.form["decision"] == "reupload":
        return redirect(url_for("upload"))
    return redirect(url_for("user_inputs"))


@app.route("/download_duplicates")
def download_duplicates():
    buf = DATA.get("duplicate_file")
    if buf is None:
        return redirect(url_for("upload"))
    buf.seek(0)
    return send_file(
        buf,
        download_name="Uploaded_with_Duplicates_Highlighted.xlsx",
        as_attachment=True,
    )


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

    missing_in_portal = pob[~pob[DATA["pob_ned"]].isin(portal[DATA["portal_ned"]])]
    missing_in_pob = portal[~portal[DATA["portal_ned"]].isin(pob[DATA["pob_ned"]])]

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
        "Charge": DATA["inputs"]["charge"],
    })

    # Manifest
    DATA["manifest"] = pd.DataFrame({
        "Passenger Weight": DATA["inputs"]["passenger_weight"],
        "Baggage Weight": DATA["inputs"]["baggage_weight"],
        "": "",
        " ": "",
        "  ": "",
        "Time Reported": DATA["inputs"]["time_reported"],
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
        "Destination Point": DATA["inputs"]["return_destination"],
    })

    return render_template(
        "result.html",
        manifest_count=DATA["manifest_count"],
        return_count=DATA["return_count"],
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
