import os
import re
from utilities import STRATASYS_ORDER

def esc(txt):
    import html
    return html.escape(txt or "")

def export_to_html(data, file_path):
    """Exports checklist and quote info data to an HTML file."""
    checklist = data.get("checklist", {})
    top_fields = checklist.get("top_fields", ["", "", ""])
    placeholders = ["Customer Name", "Opp Name", "Sales/CSR"]
    header_lines = []
    for i in range(3):
        value = top_fields[i].strip() if i < len(top_fields) else ""
        label = placeholders[i]
        display_value = value if value else "_____"
        header_lines.append(f"<b>{label}:</b> {esc(display_value)}")
    additional_notes = data.get("notes", "").strip()
    quote_rows = data.get("quote_info", [])
    vendor_quotes = data.get("vendor_quotes", [])

    # --- CSS for readability and native image size ---
    style = """
    <style>
      body {
        font-family: Segoe UI, Arial, sans-serif;
        margin: 10px;
      }

      h2 {
        margin-top: 28px;
      }

      .header {
        margin-bottom: 20px;
      }

      .additional-notes-label {
        font-weight: bold;
        margin-top: 8px;
      }

      .drawing-section,
      .vendor-section {
        overflow-x: auto;
        white-space: nowrap;
        padding-bottom: 12px;
      }

      .drawings-container,
      .vendors-container {
        display: inline-flex;
        gap: 32px;
        flex-wrap: nowrap;
      }

      .drawing-block,
      .vendor-block {
        border: 1px solid #ddd;
        border-radius: 8px;
        box-shadow: 0 1px 4px #ddd;
        padding: 18px 18px 16px 18px;
        min-width: 325px;
        max-width: 340px;
        flex: 0 0 auto;
        background: #fcfcfc;
        box-sizing: border-box;
        display: inline-block;
        vertical-align: top;
        overflow-wrap: anywhere;
        word-wrap: break-word;
        white-space: normal;
      }

      .drawing-link {
        font-size: 20px;
        font-weight: bold;
        text-align: center;
        display: block;
        margin-bottom: 7px;
        margin-top: -4px;
        color: #1a49ad;
        text-decoration: underline;
      }

      .drawing-block hr {
        border: none;
        border-top: 1.5px solid #eee;
        margin: 10px 0 5px 0;
      }

      .vendor-block {
        width: fit-content;
        min-width: 220px;
        max-width: 98vw;
        margin-bottom: 14px;
      }

      .vendor-block h3 {
        margin-top: 0;
        font-size: 20px;
        text-align: center;
        font-weight: bold;
      }

      img {
        display: block;
        margin: 12px auto;
        max-width: none;
        width: auto;
        height: auto;
        box-shadow: 0 2px 6px #eee;
      }

      table {
        width: 100%;
        border-collapse: collapse;
        table-layout: auto;
        word-break: break-word;
      }

      td, th {
        border: 1px solid #ccc;
        padding: 4px;
        font-size: 13px;
        word-break: break-word;
      }
    </style>
    """

    html = [f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>{esc(os.path.splitext(os.path.basename(file_path))[0])}</title>
{style}
</head>
<body>
"""]

    # --- Header ---
    html.append('<div class="header">')
    for line in header_lines:
        html.append(f"{line}<br>")
    if additional_notes:
        html.append(f'<div class="additional-notes-label">Additional Notes:</div>')
        html.append(f'<div>{esc(additional_notes).replace(chr(10), "<br>")}</div>')
    html.append('</div>')

    # --- Drawing(s) & 3D Print Quotes section ---
    html.append('<div class="drawing-section">')
    html.append('<h2>Drawing(s)</h2>')
    html.append('<div class="drawings-container">')

    def is_stratasys_data_present(s):
        keys = list(STRATASYS_ORDER) + ["Time (hrs)", "Material $", "3D Cost"]
        return any(
            s.get(k, "").strip() not in ("", "0", "$0.00", "$0", "0.00")
            for k in keys
        )

    def is_formlabs_data_present(f):
        keys = ["RS-F2", "Time (hrs)", "Material $", "3D Cost"]
        return any(
            f.get(k, "").strip() not in ("", "0", "$0.00", "$0", "0.00")
            for k in keys
        )

    for row in quote_rows:
        fields = row.get("fields", []) if isinstance(row, dict) else []
        if not fields or not fields[0]:
            continue
        part_num = esc(fields[0])
        pdf_path, display_name = None, part_num
        try:
            from utilities import find_latest_pdf_with_rev
            pdf_path, display_name = find_latest_pdf_with_rev(part_num)
        except Exception:
            pass
        link_html = (
            f'<a href="http://cfs-sageapps:8085/drawings/{esc(display_name)}.pdf" class="drawing-link" target="_blank">{esc(display_name)}</a>'
            if pdf_path else f'<span class="drawing-link">{esc(display_name)}</span>'
        )
        material = esc(fields[1]) if len(fields) > 1 else ""
        quantity = esc(fields[2]) if len(fields) > 2 else ""
        customer_pn = esc(fields[3]) if len(fields) > 3 else ""
        s = row.get("stratasys", {}) if isinstance(row, dict) else {}
        f = row.get("formlabs", {}) if isinstance(row, dict) else {}
        enable_3d = row.get("enable_3d", False) if isinstance(row, dict) else False


        html.append('<div class="drawing-block">')
        html.append(link_html)
        if material and material != "Material":
            html.append(f'<b>Material:</b> {material}<br>')
        if quantity and quantity != "Quantity":
            html.append(f'<b>Quantity:</b> {quantity}<br>')
        if customer_pn and customer_pn != "Customer Part Number":
            html.append(f'<b>Customer Part Number:</b> {customer_pn}<br>')
        html.append('<hr>')

        stratasys_rendered = False
        if enable_3d and is_stratasys_data_present(s):
            html.append('<b><u>Stratasys 3D Quote</u></b><br>')
            for mat in STRATASYS_ORDER:
                val = esc(s.get(mat, "")) if enable_3d else ""
                html.append(f"<b>{mat}:</b> {val}<br>")
            html.append(f"<b>Time (hrs):</b> {esc(s.get('Time (hrs)', ''))}<br>")
            html.append(f"<b>Material $:</b> {esc(s.get('Material $', ''))}<br>")
            html.append(f"<b>3D Cost:</b> {esc(s.get('3D Cost', ''))}<br>")
            stratasys_rendered = True

        if enable_3d and is_formlabs_data_present(f):
            if stratasys_rendered:
                html.append('<br>')
            html.append('<b><u>Formlabs 3D Quote</u></b><br>')
            html.append(f"<b>RS-F2:</b> {esc(f.get('RS-F2', ''))}<br>")
            html.append(f"<b>Time (hrs):</b> {esc(f.get('Time (hrs)', ''))}<br>")
            html.append(f"<b>Material $:</b> {esc(f.get('Material $', ''))}<br>")
            html.append(f"<b>3D Cost:</b> {esc(f.get('3D Cost', ''))}<br>")
        html.append('</div>')
    html.append('</div></div>')

    # --- Vendor Quotes: Render as separate cards/blocks ---
    if vendor_quotes:
        html.append('<div class="vendor-section"><h2>Vendor Quotes</h2>')
        html.append('<div class="vendors-container">')
        vendor_names = [f"Vendor {i+1}" for i in range(len(vendor_quotes))]
        for i, row in enumerate(vendor_quotes):
            if row and row[0]:
                vendor_names[i] = esc(row[0])
        for idx, row in enumerate(vendor_quotes):
            html.append('<div class="vendor-block">')
            html.append(f'<h3>{vendor_names[idx]}</h3>')
            quote_text = row[1] if len(row) > 1 else ""
            img_html = ""
            if len(row) > 2 and row[2]:
                for img_b64 in row[2]:
                    img_html += f'<img src="data:image/png;base64,{img_b64}">'
            html.append(f'<pre style="white-space: pre-wrap; margin: 0 0 10px 0; font-family: inherit; font-size: 15px;">{esc(quote_text)}</pre>{img_html}')
            html.append('</div>')

    html.append('\n</body></html>')

    with open(file_path, "w", encoding="utf-8") as f:
        f.write("".join(html))
