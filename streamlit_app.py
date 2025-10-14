# -*- coding: utf-8 -*-
import json
import io
import os
import streamlit as st
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def export_to_excel(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"

    thin = Side(border_style="thin", color="000000")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)

    # --- Header top ---
    ws["A1"], ws["B1"], ws["C1"] = "FILE", "DATE", "TYPE"
    ws.merge_cells("A1:A2")
    ws.merge_cells("B1:B2")
    ws.merge_cells("C1:C2")
    for c in ["A1", "B1", "C1"]:
        ws[c].alignment = Alignment(horizontal="center", vertical="center")
        ws[c].font = Font(bold=True)
        ws[c].border = border

    # --- Column headers ---
    headers = ["ORDER ID", "ITEM", "F/B", "SHIRT TYPE", "QUANT.", "COLOR", "SIZE", "Approved", "Note"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col, value=h)
        cell.fill = PatternFill(start_color="C9DAF8", end_color="C9DAF8", fill_type="solid")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    start_row = 5
    current_row = start_row

    orders = {}
    for item in data:
        order_id = item.get("order_external_id", "")
        idx_item = item.get("index_item", "")
        orders.setdefault(order_id, {}).setdefault(idx_item, []).append(item)

    total_all_films = 0
    total_all_shirts = 0

    current_date = datetime.now().strftime("%d-%m")

    for order_id, groups in orders.items():
        for idx_item, group_items in sorted(groups.items(), key=lambda x: int(x[0]) if str(x[0]).isdigit() else x[0]):
            item_count = len(group_items)
            labels = sorted({x.get("label", "").strip() for x in group_items if x.get("label", "")})
            fb_value = "/".join(labels) if labels else ""
            first = group_items[0]
            shirt_type = first.get("product_name", "").upper()
            color = first.get("product_color", "").strip()
            size = first.get("product_size", "").strip()
            quant = "1"

            row_vals = [order_id, item_count, fb_value, shirt_type, quant, color, size]
            for col, val in enumerate(row_vals, 1):
                cell = ws.cell(current_row, col, val)
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.cell(current_row, 8, "").border = border
            ws.cell(current_row, 9, "").border = border

            total_all_films += item_count
            total_all_shirts += 1
            current_row += 1

    # --- Footer tổng ---
    ws.cell(current_row, 1, "...")
    for c in range(1, 10):
        ws.cell(current_row, c).border = border
        ws.cell(current_row, c).alignment = Alignment(horizontal="center", vertical="center")
    current_row += 1

    ws.cell(current_row, 1, "TOTAL FILMS")
    ws.cell(current_row, 2, total_all_films)
    ws.cell(current_row, 4, "TOTAL SHIRT")
    ws.cell(current_row, 5, total_all_shirts)
    for c in range(1, 10):
        ws.cell(current_row, c).border = border
        ws.cell(current_row, c).alignment = Alignment(horizontal="center", vertical="center")

    # --- Ghi ngày hiện tại (in đậm, không nghiêng) ---
    ws["B3"] = current_date
    ws["B3"].alignment = Alignment(horizontal="center")
    ws["B3"].font = Font(bold=True, color="000000")  # In đậm, không nghiêng

    widths = [18, 8, 18, 20, 10, 16, 12, 10, 24]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + i)].width = w

    # Trả về file Excel trong memory
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer, total_all_shirts, total_all_films


# --- Streamlit App ---
st.set_page_config(page_title="JSON ➜ Excel Export Tool", layout="wide")

st.title("🧾 JSON ➜ Excel Export Tool (Multi-user Notepad)")
st.write("Công cụ cho phép **4 người (Tiên, Hải, Dung, Sơn)** cùng nhập JSON và xuất ra Excel.")

tabs = st.tabs(["Tiên", "Hải", "Dung", "Sơn"])

for tab, user_name in zip(tabs, ["Tiên", "Hải", "Dung", "Sơn"]):
    with tab:
        st.subheader(f"👤 Notepad của {user_name}")
        file_prefix = st.text_input(f"Tên file xuất Excel ({user_name}):", f"{user_name}_orders")

        json_input = st.text_area(f"Dán JSON của {user_name} vào đây:", height=250, key=f"text_{user_name}")

        if st.button(f"📤 Export Excel cho {user_name}", key=f"btn_{user_name}"):
            if not json_input.strip():
                st.warning("⚠️ Vui lòng dán dữ liệu JSON trước khi export.")
            else:
                try:
                    data = json.loads(json_input)
                    if not isinstance(data, list):
                        st.error("❌ JSON phải là danh sách (list) các object.")
                    else:
                        buffer, total_shirt, total_films = export_to_excel(data)
                        filename = f"{file_prefix}_{datetime.now().strftime('%Y%m%d')}_TOTAL_SHIRT_{total_shirt}_TOTAL_FILMS_{total_films}.xlsx"

                        st.success(f"✅ Xuất thành công cho {user_name} ({datetime.now().strftime('%Y-%m-%d')})!")
                        st.download_button(
                            label="⬇️ Tải Excel về",
                            data=buffer,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except json.JSONDecodeError as e:
                    st.error(f"❌ JSON không hợp lệ!\n{e}")
                except Exception as e:
                    st.error(f"❌ Lỗi khi xuất file: {e}")
