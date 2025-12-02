def get_tariffs_for_scan(scan_type, charge_file):
    charges = pd.read_excel(charge_file, sheet_name=None)
    results = []

    for sheet_name, df in charges.items():
        # Find EXAMINATION column
        exam_col = None
        for col in df.columns:
            if "examination" in col.lower():
                exam_col = col
                break
        if not exam_col:
            continue

        # Find rows where EXAMINATION matches or is similar to scan_type
        df["__lower_exam"] = df[exam_col].astype(str).str.lower()
        user_lower = scan_type.lower()
        matched_rows = df[df["__lower_exam"].str.contains(user_lower)]

        if not matched_rows.empty:
            # Return all matching rows as dicts
            for _, row in matched_rows.iterrows():
                results.append(row.to_dict())
            return results  # Return on first sheet match

    return None

def fill_template_multiple(template_file, patient, medaid, scan, tariffs):
    wb = load_workbook(template_file)
    sheet = wb.active

    # Find cells for patient, medaid, scan (same as before)
    patient_cell = find_cell_any(sheet, ["patient", "name", "client"])
    medaid_cell  = find_cell_any(sheet, ["medical", "member", "aid", "scheme"])
    scan_cell    = find_cell_any(sheet, ["scan", "exam", "procedure", "service", "investigation"])
    
    # Find header row with DESCRIPTION
    header_cell = find_cell_any(sheet, ["description"])
    if not header_cell:
        st.error("Template missing DESCRIPTION header.")
        return None

    # Find column indices for needed headers relative to DESCRIPTION header row
    header_row = header_cell[0]
    headers = {}
    for cell in sheet[header_row]:
        if cell.value:
            val = str(cell.value).strip().lower()
            headers[val] = cell.column

    required_cols = ["description", "tarrif", "mod", "qty", "fees", "amount"]
    for col in required_cols:
        if col not in headers:
            st.error(f"Template missing column header: {col.upper()}")
            return None

    # Write patient, medaid, scan
    safe_write(sheet, patient_cell[0], patient_cell[1] + 1, patient)
    safe_write(sheet, medaid_cell[0], medaid_cell[1] + 1, str(medaid))
    safe_write(sheet, scan_cell[0], scan_cell[1] + 1, scan)

    # Start writing tariffs from row after header
    start_row = header_row + 1

    for i, tariff_row in enumerate(tariffs):
        row_idx = start_row + i

        safe_write(sheet, row_idx, headers["description"], str(tariff_row.get("EXAMINATION", "")))
        safe_write(sheet, row_idx, headers["tarrif"], tariff_row.get("TARIFF", ""))
        safe_write(sheet, row_idx, headers["mod"], tariff_row.get("MOD", ""))
        safe_write(sheet, row_idx, headers["qty"], tariff_row.get("QTY", 1))
        safe_write(sheet, row_idx, headers["fees"], tariff_row.get("FEES", ""))
        safe_write(sheet, row_idx, headers["amount"], tariff_row.get("AMOUNT", ""))

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
