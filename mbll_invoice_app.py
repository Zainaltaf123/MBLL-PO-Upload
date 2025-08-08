import streamlit as st
import pandas as pd
import io
import zipfile
import os
from openpyxl import load_workbook
from openpyxl.writer.excel import save_virtual_workbook

st.set_page_config(page_title="MBLL Invoice App", layout="centered")

st.title("📦 MBLL Invoice Generator")
st.write("Upload the MBLL Order Summary and Invoice Template files below:")

# File uploads
order_file = st.file_uploader("📄 MBLL Order Summary (.xlsx)", type=["xlsx"])
template_file = st.file_uploader("📄 Invoice Template (.xlsx)", type=["xlsx"])

if order_file and template_file:
    with st.spinner("Processing invoices..."):
        # Load the order data
        df = pd.read_excel(order_file)

        # Pivot data
        pivot_df = df.pivot_table(
            index=["Store Name", "Supplier", "Supplier Reference", "PO Number", "TechPOS Sku"],
            values=["Total Units", "Unit Cost", "Total ($)"],
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        # Prepare summary
        summary_df = pivot_df.groupby(["Store Name", "Supplier Reference"]).agg(
            Total_SKUs=("TechPOS Sku", "nunique"),
            Total_Quantity=("Total Units", "sum"),
            Total_Cost=("Total ($)", "sum")
        ).reset_index()

        # Create zip in memory
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            # Loop through each store and supplier reference
            for (store, supplier_ref), group_df in pivot_df.groupby(["Store Name", "Supplier Reference"]):
                supplier_data = group_df[["TechPOS Sku", "Total Units", "Unit Cost"]]

                # Load template
                template_file.seek(0)
                wb = load_workbook(template_file)
                ws = wb.active

                # Set Supplier Reference
                ws["B8"] = supplier_ref

                # Write product data from A14
                start_row = 14
                for i, row in supplier_data.iterrows():
                    ws.cell(row=start_row, column=1).value = row["TechPOS Sku"]
                    ws.cell(row=start_row, column=2).value = row["Total Units"]
                    ws.cell(row=start_row, column=3).value = row["Unit Cost"]
                    start_row += 1

                # Save Excel to bytes
                invoice_buffer = io.BytesIO()
                wb.save(invoice_buffer)
                invoice_buffer.seek(0)

                # Create folder path inside ZIP
                safe_store = store.replace("/", "-").replace("\\", "-")
                filename = f"{safe_store} - {supplier_ref} MBLL Invoice.xlsx"
                folder_path = f"{safe_store}/{filename}"

                # Add to zip
                zip_file.writestr(folder_path, invoice_buffer.read())

        zip_buffer.seek(0)

        # --- Create summary Excel buffer ---
        summary_excel = io.BytesIO()
        with pd.ExcelWriter(summary_excel, engine='openpyxl') as writer:
            summary_df.to_excel(writer, index=False)
        summary_excel.seek(0)

        st.success("✅ Processing complete! Download your files below:")

        # --- Download Buttons ---
        st.download_button(
            label="📥 Download Summary Excel",
            data=summary_excel,
            file_name="MBLL_Invoice_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="📦 Download ZIP of Invoices",
            data=zip_buffer,
            file_name="MBLL_Invoices.zip",
            mime="application/zip"
        )
else:
    st.info("📂 Please upload both files to begin.")
