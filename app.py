# app.py
# ---------------------------------------------------------------------------
# Đây là file giao diện web, viết bằng thư viện Streamlit.
#
# Cách chạy app (gõ trong terminal, tại thư mục chứa file này):
#     streamlit run app.py
#
# Sau khi chạy, app sẽ tự mở trên trình duyệt ở địa chỉ localhost.
# ---------------------------------------------------------------------------

import streamlit as st
import pandas as pd

import database  # đây là file database.py ở cùng thư mục
import bulk_import  # import nhiều file raw một lúc
import reconciliation  # logic đối chiếu chứng từ + tính doanh thu
import export  # logic xuất file Excel

# Cấu hình chung cho trang web
st.set_page_config(page_title="App Kiểm kê Doanh thu", layout="wide")

# Tạo các bảng trong database nếu chưa có.
# Gọi mỗi lần mở app cũng không sao, vì dùng "CREATE TABLE IF NOT EXISTS".
database.create_tables()


# ---------------------------------------------------------------------------
# SIDE MENU (menu bên trái)
# Sau này muốn thêm chức năng mới, chỉ cần thêm tên vào danh sách bên dưới
# và viết thêm một hàm hiển thị tương ứng.
# ---------------------------------------------------------------------------
st.sidebar.title("Menu")
selected_feature = st.sidebar.radio(
    "Chọn chức năng:",
    ["Import data", "Tạo bảng doanh thu", "Fill doanh thu", "Export Excel"],
)


# ---------------------------------------------------------------------------
# CHỨC NĂNG: IMPORT DATA
# ---------------------------------------------------------------------------
def show_bulk_import_section():
    """Import nhiều file raw một lúc (xem bulk_import.py)."""
    hotel_label = st.selectbox("Chọn cơ sở:", list(bulk_import.HOTELS))
    skip_company = st.checkbox(
        "Tạm bỏ qua sao kê tài khoản công ty",
        value=True,
        help="Tick: không import bảng company_bank_transactions (giữ nguyên dữ liệu cũ). "
             "Bỏ tick: bắt buộc phải có file sao kê tài khoản công ty, nếu không sẽ báo lỗi.",
    )
    uploaded_files = st.file_uploader(
        "Kéo thả tất cả file raw vào đây (csv, xlsx, xls):",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
    )
    st.caption(
        "Tool tự nhận diện từng file, bỏ phần tiêu đề/tổng/footer thừa và import vào đúng bảng. "
        "Nếu có file sai format thì KHÔNG import gì cả (dữ liệu cũ giữ nguyên)."
    )

    # File thu chi tiền mặt mà tool chưa tự chọn được sheet -> cho người dùng chọn.
    uploaded_names = {f.name for f in (uploaded_files or [])}
    cash_sheets = {}
    for request in st.session_state.get("bulk_sheet_requests", []):
        if request["file"] not in uploaded_names:
            continue
        cash_sheets[request["file"]] = st.multiselect(
            f"Chọn sheet tiền mặt cần import từ file '{request['file']}':",
            request["sheets"],
            key=f"bulk_sheets_{request['file']}",
        )

    if st.button("Import", disabled=not uploaded_files):
        files = [(f.name, f.getvalue()) for f in uploaded_files]
        with st.spinner("Đang xử lý và import..."):
            result = bulk_import.run_bulk_import(
                files,
                hotel_label,
                skip_company=skip_company,
                cash_sheets={k: v for k, v in cash_sheets.items() if v},
            )
        st.session_state["bulk_result"] = result
        old_files = [r["file"] for r in st.session_state.get("bulk_sheet_requests", [])]
        st.session_state["bulk_sheet_requests"] = result.sheet_requests
        if [r["file"] for r in result.sheet_requests] != old_files:
            st.rerun()  # vẽ lại để hiện / ẩn ô chọn sheet phía trên nút Import

    result = st.session_state.get("bulk_result")
    if result is None:
        return

    st.subheader("Kết quả từng file")
    st.dataframe(
        pd.DataFrame(result.files).rename(columns={
            "file": "File", "kind": "Loại file", "table": "Bảng",
            "rows": "Số dòng", "status": "Trạng thái", "detail": "Ghi chú",
        }),
        width="stretch",
    )

    if result.errors:
        st.error(
            "Có lỗi format nên ĐÃ HỦY import, toàn bộ dữ liệu cũ được giữ nguyên. "
            "Các lỗi cần sửa:"
        )
        for file_name, message in result.errors:
            st.write(f"- **{file_name}**: {message}")
    elif result.sheet_requests:
        st.info(
            "Chưa import gì. Hãy chọn sheet tiền mặt cần import ở phía trên "
            "rồi bấm Import lại."
        )
    elif result.committed:
        st.success(
            "Import thành công: "
            + ", ".join(f"{t} ({n} dòng)" for t, n in result.tables.items())
        )
    for warning in result.warnings:
        st.warning(warning)


def show_import_data_page():
    st.title("Import data")

    mode = st.radio(
        "Chế độ import:",
        ["Import từng bảng", "Import nhiều file 1 lúc"],
        horizontal=True,
    )
    if mode == "Import nhiều file 1 lúc":
        show_bulk_import_section()
        return

    # --- Bước 1: chọn bảng cần import ---
    table_names = list(database.TABLES.keys())
    selected_table = st.selectbox("Bước 1 — Chọn bảng cần import dữ liệu:", table_names)

    # Dành sẵn một "ô trống" ở vị trí này để hiển thị dữ liệu hiện tại của bảng.
    # Ta sẽ điền dữ liệu vào ô này ở CUỐI hàm (sau khi import xong), để bảng
    # luôn hiển thị dữ liệu mới nhất mà vẫn nằm đúng ngay dưới phần chọn bảng.
    current_data_container = st.container()

    # --- Bước 2: upload file để import ---
    st.subheader("Bước 2 — Upload file để import (nhớ sửa lại file để bảng dữ liệu bắt đầu từ ô A1)")
    uploaded_file = st.file_uploader(
        "Chọn file CSV hoặc Excel (.xlsx, .xls):",
        type=["csv", "xlsx", "xls"],
    )

    if uploaded_file is not None:
        # Đọc file vào DataFrame, chọn cách đọc theo đuôi file:
        #   .csv  -> đọc bằng read_csv
        #   .xlsx -> đọc bằng read_excel với engine "openpyxl"
        #   .xls  -> đọc bằng read_excel với engine "xlrd" (định dạng Excel cũ)
        file_name = uploaded_file.name.lower()
        if file_name.endswith(".csv"):
            new_data = pd.read_csv(uploaded_file)
        elif file_name.endswith(".xlsx"):
            new_data = pd.read_excel(uploaded_file, engine="openpyxl")
        else:  # .xls
            new_data = pd.read_excel(uploaded_file, engine="xlrd")

        # Cho người dùng xem trước nội dung file trước khi import
        st.write("Xem trước dữ liệu trong file vừa upload:")
        st.dataframe(new_data, width="stretch")

        # --- Mapping thủ công (chỉ với các bảng trong MANUAL_MAPPING_TABLES) ---
        # Với những bảng này, người dùng tự gõ tên cột trong file cho từng cột DB.
        # custom_mapping = None nghĩa là dùng mapping cố định (cho các bảng booking).
        custom_mapping = None
        if selected_table in database.MANUAL_MAPPING_TABLES:
            st.write(
                "Nhập tên cột trong file tương ứng với mỗi cột bên dưới "
                "(để trống nếu file không có cột đó):"
            )
            custom_mapping = {}
            for db_col in database.TABLES[selected_table]:
                # key phải là duy nhất để Streamlit không nhầm lẫn giữa các ô
                file_col = st.text_input(
                    f"{db_col}  →  tên cột trong file",
                    key=f"manual_map_{selected_table}_{db_col}",
                )
                # Chỉ thêm vào mapping nếu người dùng có nhập (không để trống).
                # Lưu ý: dùng đúng văn bản đã gõ (không cắt khoảng trắng), vì
                # tên cột trong file đôi khi có khoảng trắng ở cuối.
                if file_col.strip() != "":
                    custom_mapping[db_col] = file_col

        # Nút bấm để xác nhận import (tránh import nhầm khi chỉ vừa chọn file)
        if st.button("Import vào bảng"):
            row_count, mapped_pairs, missing_file_cols = database.import_dataframe(
                selected_table, new_data, custom_mapping
            )

            if row_count == 0:
                st.warning(
                    "Không tìm thấy cột nào trong file khớp với mapping của bảng, "
                    "nên chưa import được dòng nào. Dữ liệu cũ được giữ nguyên."
                )
                if missing_file_cols:
                    st.info(
                        "Các cột mong đợi (theo mapping) nhưng không thấy trong file: "
                        f"{missing_file_cols}"
                    )
            else:
                st.success(
                    f"Đã import {row_count} dòng vào bảng '{selected_table}' "
                    "(đã ghi đè toàn bộ dữ liệu cũ)."
                )
                # Hiển thị các cặp đã map theo dạng: "tên cột file → tên cột DB"
                mapping_display = [f"{file_col} → {db_col}" for db_col, file_col in mapped_pairs]
                st.write("Các cột đã được map và import:", mapping_display)
                if missing_file_cols:
                    st.info(
                        "Các cột mong đợi (theo mapping) nhưng không thấy trong file "
                        f"(để trống): {missing_file_cols}"
                    )

    # --- Hiển thị dữ liệu hiện tại của bảng (luôn là dữ liệu mới nhất) ---
    # Phần này được điền vào "ô trống" đã dành sẵn ở Bước 1.
    # Vì nó chạy SAU đoạn import ở trên, nên nếu vừa import xong, dữ liệu hiển
    # thị ở đây đã là dữ liệu mới (đã được refresh).
    with current_data_container:
        st.subheader(f"Dữ liệu hiện tại của bảng: {selected_table}")
        current_data = database.get_table_data(selected_table)
        st.dataframe(current_data, width="stretch")
        st.caption(f"Tổng số dòng hiện tại: {len(current_data)}")


# ---------------------------------------------------------------------------
# CHỨC NĂNG: TẠO BẢNG DOANH THU
# ---------------------------------------------------------------------------
def show_create_revenue_page():
    st.title("Tạo bảng doanh thu")

    # Danh sách phòng cố định cho 2 cơ sở (khai báo ngay đây cho dễ đọc/sửa)
    rooms_signature = [
        "201", "202", "203", "204", "301", "302", "303", "304",
        "401", "402", "403", "404", "501", "502", "503", "504",
        "601", "602", "603", "604", "701", "702", "703", "704", "801",
    ]
    rooms_atelier = [
        "201", "202", "203", "204", "301", "302", "303", "304",
        "401", "402", "403", "404", "501", "502", "503", "504",
        "601", "602", "603", "701", "702", "703",
    ]

    # --- Nhập số ngày trong tháng và chọn tháng ---
    num_days = st.number_input(
        "Số ngày trong tháng:",
        min_value=1,
        max_value=31,
        value=30,
        step=1,
    )
    month = st.selectbox("Tháng cần kiểm kê:", list(range(1, 13)))
    year = st.number_input(
        "Năm:",
        min_value=2000,
        max_value=2100,
        value=2026,
        step=1,
    )

    # --- Chọn danh sách phòng (room name) ---
    room_choice = st.selectbox(
        "Chọn danh sách phòng:",
        ["cơ sở irest signature", "cơ sở irest atelier", "Tự nhập"],
    )

    if room_choice == "cơ sở irest signature":
        room_names = rooms_signature
    elif room_choice == "cơ sở irest atelier":
        room_names = rooms_atelier
    else:  # Tự nhập
        typed = st.text_input("Nhập danh sách phòng, cách nhau bởi dấu phẩy:")
        # Tách theo dấu phẩy, bỏ khoảng trắng thừa, bỏ phần tử rỗng
        room_names = [name.strip() for name in typed.split(",") if name.strip() != ""]

    # Hiển thị lại danh sách phòng đã chọn để người dùng kiểm tra
    if room_names:
        st.caption(f"Danh sách phòng ({len(room_names)} phòng): {room_names}")

    # --- Nút tạo bảng: chỉ bật khi đã có danh sách phòng ---
    can_create = len(room_names) > 0
    if not can_create:
        st.info("Hãy chọn (hoặc nhập) danh sách phòng để có thể tạo bảng doanh thu.")

    # Hàm nhỏ dùng lại: thực sự tạo bảng và báo thành công
    def do_create():
        database.create_revenue_table(room_names, int(num_days), int(month), int(year))
        st.success(
            f"Đã tạo bảng doanh thu cho tháng {month}/{int(year)} "
            f"({int(num_days)} ngày, {len(room_names)} phòng)."
        )

    if st.button("Tạo bảng doanh thu", disabled=not can_create):
        # Nếu bảng đang có dữ liệu -> chuyển sang trạng thái "chờ xác nhận"
        # (sẽ hiện cảnh báo + nút xác nhận bên dưới), chưa xóa gì cả.
        if database.revenue_table_has_data():
            st.session_state["confirm_recreate_revenue"] = True
        else:
            do_create()

    # --- Phần xác nhận khi bảng đã có dữ liệu ---
    if st.session_state.get("confirm_recreate_revenue", False):
        st.warning(
            "Bảng doanh thu đã tồn tại và đang có dữ liệu. "
            "Nếu tiếp tục, toàn bộ dữ liệu cũ sẽ bị xóa và tạo lại theo lựa chọn mới."
        )
        col_yes, col_no = st.columns(2)
        with col_yes:
            if st.button("Xác nhận xóa và tạo lại"):
                do_create()
                st.session_state["confirm_recreate_revenue"] = False
        with col_no:
            if st.button("Hủy"):
                st.session_state["confirm_recreate_revenue"] = False
                st.info("Đã hủy. Dữ liệu cũ được giữ nguyên.")

    # --- Hiển thị bảng doanh thu hiện tại (nếu đã có) ---
    if database.revenue_table_exists():
        st.subheader("Bảng doanh thu hiện tại")
        st.dataframe(database.get_revenue_table_data(), width="stretch")


# ---------------------------------------------------------------------------
# CHỨC NĂNG: FILL DOANH THU (hiện tại: điền Ezcloud_booking_id)
# ---------------------------------------------------------------------------
def show_fill_revenue_page():
    st.title("Fill doanh thu")

    # Phải có bảng doanh thu trước đã
    if not database.revenue_table_exists():
        st.warning(
            "Chưa có bảng doanh thu. Hãy sang tab 'Tạo bảng doanh thu' để tạo trước."
        )
        return

    year = database.get_revenue_meta_year()
    if year is None:
        st.warning(
            "Bảng doanh thu chưa lưu thông tin năm (có thể được tạo từ trước). "
            "Hãy tạo lại bảng ở tab 'Tạo bảng doanh thu' để có năm."
        )
        return

    st.write(f"Năm áp dụng cho bảng doanh thu hiện tại: **{year}**")
    st.caption(
        "Thao tác này sẽ điền Ezcloud_booking_id vào từng ô [phòng, ngày] dựa "
        "trên bảng ezcloud_booking_list. Ô có sẵn dữ liệu sẽ bị ghi đè."
    )

    if st.button("Fill Ezcloud_booking_id vào bảng doanh thu"):
        filled_count, multi_matches, parse_errors = database.fill_ezcloud_booking_ids(year)
        st.success(f"Đã điền {filled_count} ô.")

        # Liệt kê các ô có nhiều hơn 1 booking để người dùng rà soát
        if multi_matches:
            st.warning(f"Có {len(multi_matches)} ô khớp NHIỀU booking — hãy rà soát:")
            review_rows = [
                {"Phòng": room, "Ngày": day_col, "Các booking id": ", ".join(ids)}
                for room, day_col, ids in multi_matches
            ]
            st.dataframe(pd.DataFrame(review_rows), width="stretch")

        # Liệt kê các giá trị ngày giờ không đọc được (nếu có)
        if parse_errors:
            st.info(
                "Một số giá trị Checkin/Checkout không đọc được (đã bỏ qua), "
                f"hãy kiểm tra lại: {parse_errors}"
            )

    # --- Bước 2: Fill doanh thu đầy đủ (đối chiếu chứng từ + tính tiền) ---
    st.divider()
    st.subheader("Bước 2: Fill doanh thu (đầy đủ)")
    st.caption(
        "Thao tác này sẽ: tự điền lại ezcloud_booking_id, đối chiếu từng khoản "
        "thanh toán với chứng từ (bảng OTA, sao kê ngân hàng, tiền mặt), tính "
        "doanh thu từng đêm-phòng và điền theo format [id, doanh thu, company, "
        "pay_type]. Booking thiếu chứng từ sẽ hiển thị "
        '["Không tìm thấy chứng từ", id]; booking thanh toán bằng thẻ tín dụng '
        'hiển thị ["Thanh toán bằng Thẻ tín dụng", id]. Mỗi lần chạy sẽ reset '
        "lại toàn bộ is_done."
    )

    if st.button("Fill doanh thu (đầy đủ)"):
        summary = reconciliation.fill_revenue_full()
        st.success(f"Đã xử lý xong. Số booking hợp lệ: {summary['ok']}.")

        if summary["credit_card"]:
            st.info(
                f"{len(summary['credit_card'])} booking thanh toán bằng thẻ tín "
                f"dụng (đã bỏ qua): {summary['credit_card']}"
            )

        if summary["errors"]:
            st.warning(f"Có {len(summary['errors'])} booking lỗi — hãy rà soát:")
            err_rows = [
                {"Ezcloud_booking_id": bid, "Lý do": reason}
                for bid, reason in summary["errors"]
            ]
            st.dataframe(pd.DataFrame(err_rows), width="stretch")

    # Hiển thị bảng doanh thu hiện tại
    st.subheader("Bảng doanh thu hiện tại")
    st.dataframe(database.get_revenue_table_data(), width="stretch")


# ---------------------------------------------------------------------------
# CHỨC NĂNG: EXPORT EXCEL
# ---------------------------------------------------------------------------
def show_export_excel_page():
    st.header("Export Excel")
    st.caption(
        "Xuất bảng doanh thu ra file Excel: cột là các ngày trong tháng, dòng "
        "là phòng, giá trị ô là doanh thu đêm phòng. Màu nền mỗi ô theo Company "
        "(OTA) hoặc phương thức thanh toán — bạn có thể chỉnh màu bên dưới trước "
        "khi tải về."
    )

    if not database.revenue_table_exists():
        st.warning("Chưa có bảng doanh thu. Hãy tạo và fill doanh thu trước.")
        return

    st.subheader("Chỉnh màu cho từng loại")
    # Hiển thị color picker theo 3 cột cho gọn
    color_settings = {}
    cols = st.columns(3)
    for idx, key in enumerate(export.COLOR_KEYS):
        default_hex = "#" + export.DEFAULT_COLORS[key]
        with cols[idx % 3]:
            chosen = st.color_picker(key, default_hex, key=f"color_{key}")
        color_settings[key] = chosen

    st.divider()

    # Dựng file theo màu hiện tại và cho tải về
    xlsx_bytes = export.build_revenue_xlsx_bytes(color_settings)
    st.download_button(
        label="Tải file Excel",
        data=xlsx_bytes,
        file_name="doanh_thu.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# Hiển thị trang tương ứng với chức năng được chọn ở side menu
# ---------------------------------------------------------------------------
if selected_feature == "Import data":
    show_import_data_page()
elif selected_feature == "Tạo bảng doanh thu":
    show_create_revenue_page()
elif selected_feature == "Fill doanh thu":
    show_fill_revenue_page()
elif selected_feature == "Export Excel":
    show_export_excel_page()
