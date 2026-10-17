# database.py
# ---------------------------------------------------------------------------
# File này lo TOÀN BỘ việc liên quan đến database SQLite:
#   - Định nghĩa cấu trúc các bảng (có những cột gì)
#   - Tạo các bảng nếu chưa có
#   - Đọc dữ liệu của một bảng
#   - Thêm (import) dữ liệu vào một bảng
#
# Phần giao diện (app.py) chỉ gọi các hàm ở đây, không tự đụng vào database.
# Tách riêng như vậy để dễ đọc và dễ debug.
# ---------------------------------------------------------------------------

import sqlite3
import re
import math

import pandas as pd

import matching

# Tên file database. File này sẽ tự động được tạo ở cùng thư mục khi chạy app lần đầu.
DB_NAME = "data.db"


# ---------------------------------------------------------------------------
# ĐỊNH NGHĨA CẤU TRÚC CÁC BẢNG (ở MỘT chỗ duy nhất cho dễ sửa)
#
# Mỗi bảng là một danh sách các cột DỮ LIỆU (cột sẽ được import từ file).
# Lưu ý: cột "is_done" KHÔNG nằm trong danh sách này. Lý do: "is_done" là cờ
# đánh dấu do hệ thống tự quản lý (mặc định = 0), không phải dữ liệu lấy từ
# file, nên ta tạo nó tự động cho mọi bảng ở hàm create_tables() bên dưới.
# ---------------------------------------------------------------------------
TABLES = {
    "ezcloud_booking_list": [
        "Ezcloud_booking_id",
        "Customer_name",
        "Room_number",
        "Checkin_date",
        "Checkout_date",
        "Company",
        "OTA_id",
        "Room_price",
        "Service_price",
        "Discount",
        "Total_income",
    ],
    "ezcloud_payment_list": [
        "Ezcloud_booking_id",
        "Customer_name",
        "Room_number",
        "Content",
        "Pay_amount",
        "Pay_type",
        "Pay_time",
    ],
    "agoda_booking_list": [
        "Agoda_booking_id",
        "Net_income",
        "Gross_income",
    ],
    "booking_booking_list": [
        "Booking_booking_id",
        "Net_income",
        "Gross_income",
    ],
    "traveloka_booking_list": [
        "Traveloka_booking_id",
        "Net_income",
        "Gross_income",
    ],
    "expedia_booking_list": [
        "Expedia_booking_id",
        "Net_income",
        "Gross_income",  # Ghi chú của bạn: Expedia chỉ trả về giá này
    ],
    "ctrip_booking_list": [
        "Ctrip_booking_id",
        "Net_income",
        "Gross_income",
    ],
    "personal_bank_transactions": [
        "Time",
        "Income",
        "Outcome",
    ],
    "company_bank_transactions": [
        "Time",
        "Income",
        "Outcome",
    ],
    "cash_transactions": [
        "Time",
        "Content",
        "Income",
        "Outcome",
        "Note",
    ],
}


# ---------------------------------------------------------------------------
# BẢNG MAPPING: cột trong DB  <==  tên cột trong file CSV/Excel
#
# Cấu trúc:  COLUMN_MAPPING[tên_bảng][tên_cột_DB] = "tên cột trong file"
#
# Cách đọc một dòng, ví dụ:
#     "Customer_name": "Tên khách"
# nghĩa là: cột "Customer_name" trong DB sẽ lấy dữ liệu từ cột có tiêu đề
# "Tên khách" trong file.
#
# Lưu ý:
#   - Cột DB nào KHÔNG có trong mapping (ví dụ Net_income của Booking/Expedia,
#     hoặc is_done) thì sẽ KHÔNG được lấy từ file -> để trống (NULL).
#   - 3 bảng transaction để mapping rỗng {} vì bạn sẽ định nghĩa sau.
#   - Tên cột trong file phải gõ ĐÚNG CHÍNH XÁC như tiêu đề trong file
#     (đúng chữ hoa/thường, đúng dấu cách, đúng dấu tiếng Việt).
# ---------------------------------------------------------------------------
COLUMN_MAPPING = {
    "ezcloud_booking_list": {
        "Ezcloud_booking_id": "Mã đặt phòng",
        "Customer_name": "Tên khách",
        "Room_number": "Tên phòng",
        "Checkin_date": "Ngày đến",
        "Checkout_date": "Ngày đi",
        "Company": "Công ty",
        "OTA_id": "Mã OTA",
        "Room_price": "Tiền phòng",
        "Service_price": "Dịch vụ",
        "Discount": "Chiết khấu",
        "Total_income": "Tổng doanh thu",
    },
    "ezcloud_payment_list": {
        "Ezcloud_booking_id": "#Phòng",
        "Customer_name": "Tên khách",
        "Room_number": "Tên phòng",
        "Content": "Nội dung",
        "Pay_amount": "Số tiền ",
        "Pay_type": "Hình thức TT",
        "Pay_time": "Thời gian",
    },
    "agoda_booking_list": {
        "Agoda_booking_id": "BookingIDExternal_reference_ID",
        "Net_income": "Total_inclusive_rate",
        "Gross_income": "ReferenceSellInclusive",
    },
    "booking_booking_list": {
        "Booking_booking_id": "Mã số đặt phòng",
        # Net_income: không có cột trong file -> bỏ trống
        "Gross_income": "Giá",
    },
    "traveloka_booking_list": {
        "Traveloka_booking_id": "Itinerary ID",
        "Net_income": "Net Fare",
        "Gross_income": "Fare",
    },
    "expedia_booking_list": {
        "Expedia_booking_id": "Mã đặt phòng",
        # Net_income: không có cột trong file -> bỏ trống
        "Gross_income": "Khoản tiền đặt phòng",
    },
    "ctrip_booking_list": {
        "Ctrip_booking_id": "Mã Xác Nhận Của Khách Sạn",
        "Net_income": "Giá phòng thu về",
        "Gross_income": "Giá bán phòng",
    },
    # --- 3 bảng dưới đây dùng MAPPING THỦ CÔNG (xem MANUAL_MAPPING_TABLES) ---
    # Người dùng tự nhập tên cột trong file lúc import, nên ở đây để rỗng {}.
    "personal_bank_transactions": {},
    "company_bank_transactions": {},
    "cash_transactions": {},
}


# ---------------------------------------------------------------------------
# Các bảng dùng MAPPING THỦ CÔNG:
# Khác với các bảng booking (mapping cố định trong COLUMN_MAPPING), với các
# bảng dưới đây người dùng sẽ TỰ NHẬP tên cột trong file lúc import. Tên cột
# người dùng nhập sẽ được truyền vào import_dataframe(...) qua tham số
# custom_mapping.
# ---------------------------------------------------------------------------
MANUAL_MAPPING_TABLES = [
    "personal_bank_transactions",
    "company_bank_transactions",
    "cash_transactions",
]


def get_connection():
    """Mở một kết nối tới file database. Dùng xong nhớ gọi conn.close()."""
    return sqlite3.connect(DB_NAME)


def create_tables():
    """
    Tạo tất cả các bảng nếu chúng CHƯA tồn tại.

    Quy ước cho bước này:
      - Mọi cột dữ liệu đều để kiểu TEXT cho đơn giản (không lo lỗi sai kiểu
        dữ liệu khi import). Sau này khi làm phần mapping, nếu cần đổi kiểu
        (số, ngày...) thì ta sửa lại sau.
      - Mỗi bảng tự động có thêm cột "is_done" kiểu INTEGER, mặc định = 0.
    """
    conn = get_connection()
    cursor = conn.cursor()

    for table_name, columns in TABLES.items():
        # Tạo phần định nghĩa cột. Ví dụ: "Customer_name" TEXT, "Room_number" TEXT
        column_definitions = ", ".join(f'"{column}" TEXT' for column in columns)

        # Thêm cột is_done vào cuối mỗi bảng
        column_definitions += ', "is_done" INTEGER DEFAULT 0'

        sql = f'CREATE TABLE IF NOT EXISTS "{table_name}" ({column_definitions})'
        cursor.execute(sql)

        # --- Tự động bổ sung cột còn thiếu cho bảng ĐÃ tồn tại từ trước ---
        # (Quan trọng: CREATE TABLE IF NOT EXISTS không thêm cột mới vào bảng cũ.
        #  Nếu sau này ta thêm cột vào TABLES, đoạn này sẽ thêm cột đó vào DB cũ
        #  mà không làm mất dữ liệu.)
        existing_columns = [row[1] for row in cursor.execute(
            f'PRAGMA table_info("{table_name}")'
        ).fetchall()]

        for column in columns:
            if column not in existing_columns:
                cursor.execute(
                    f'ALTER TABLE "{table_name}" ADD COLUMN "{column}" TEXT'
                )
        if "is_done" not in existing_columns:
            cursor.execute(
                f'ALTER TABLE "{table_name}" ADD COLUMN "is_done" INTEGER DEFAULT 0'
            )

    conn.commit()
    conn.close()


def get_table_data(table_name):
    """Đọc toàn bộ dữ liệu của một bảng, trả về dưới dạng bảng (pandas DataFrame)."""
    conn = get_connection()
    df = pd.read_sql_query(f'SELECT * FROM "{table_name}"', conn)
    conn.close()
    return df


def normalize_column_name(name):
    """
    Chuẩn hóa tên cột TRƯỚC KHI SO SÁNH (chỉ dùng để so khớp, không làm đổi
    dữ liệu thật trong file).

    Việc chuẩn hóa gồm:
      - Đổi mọi ký tự xuống dòng (\\n, \\r) và tab thành dấu cách.
      - Gộp nhiều khoảng trắng liên tiếp thành 1 dấu cách.
      - Bỏ khoảng trắng thừa ở đầu và cuối.

    Nhờ vậy, tên cột trong file là dạng 2 dòng (ví dụ "Số bút toán\\nTransaction No")
    vẫn khớp với khi người dùng gõ trên 1 dòng ("Số bút toán Transaction No").
    """
    text = str(name)
    # \s khớp với mọi loại khoảng trắng (dấu cách, tab, xuống dòng).
    # Gộp mọi cụm khoảng trắng thành đúng 1 dấu cách.
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def import_dataframe(table_name, df, custom_mapping=None):
    """
    Đưa dữ liệu từ một DataFrame (đọc từ file CSV/Excel) vào bảng tương ứng,
    DỰA THEO MAPPING.

    Mapping được chọn như sau:
      - Nếu có truyền custom_mapping (mapping do NGƯỜI DÙNG tự nhập, dùng cho
        các bảng trong MANUAL_MAPPING_TABLES) thì dùng nó.
      - Nếu không, dùng mapping CỐ ĐỊNH đã khai báo trong COLUMN_MAPPING.

    Cách hoạt động:
      - Với mỗi cột DB trong mapping, tìm cột tương ứng trong file (theo tên
        cột trong file). Nếu file có cột đó thì lấy dữ liệu và đổi tên thành
        tên cột DB.
      - GHI ĐÈ TOÀN BỘ: trước khi thêm dữ liệu mới, bảng sẽ bị xóa sạch dữ
        liệu cũ. Sau khi import, bảng chỉ còn đúng dữ liệu vừa import.
      - Lưu ý an toàn: nếu KHÔNG map được cột nào (không tìm thấy cột nào
        trong file khớp mapping) thì KHÔNG xóa gì cả, để tránh lỡ tay xóa
        mất dữ liệu khi chọn nhầm file/nhầm bảng.

    Trả về một bộ 3 giá trị:
      (số dòng đã thêm,
       danh sách các cặp đã map - mỗi cặp là (tên_cột_DB, tên_cột_file),
       danh sách tên cột file mong đợi nhưng KHÔNG tìm thấy trong file)
    """
    # Chọn mapping để dùng. Dạng: {tên_cột_DB: "tên cột trong file"}
    if custom_mapping is not None:
        mapping = custom_mapping
    else:
        mapping = COLUMN_MAPPING.get(table_name, {})

    mapped_pairs = []        # các cặp (db_col, file_col) map thành công
    missing_file_cols = []   # các cột file mong đợi nhưng không có trong file

    # Tạo "bảng tra cứu" cho các cột trong file:
    #   tên cột đã chuẩn hóa  ->  tên cột GỐC trong file
    # (Ta cần tên gốc để lấy đúng dữ liệu df[tên_gốc].)
    normalized_to_original = {}
    for original_col in df.columns:
        normalized_to_original[normalize_column_name(original_col)] = original_col

    # Gom dữ liệu cần chèn, đặt tên cột theo tên cột DB
    data_for_insert = {}
    for db_col, file_col in mapping.items():
        # So khớp theo tên ĐÃ CHUẨN HÓA (bỏ qua xuống dòng/khoảng trắng thừa)
        normalized_target = normalize_column_name(file_col)
        if normalized_target in normalized_to_original:
            original_col = normalized_to_original[normalized_target]
            data_for_insert[db_col] = df[original_col]
            mapped_pairs.append((db_col, file_col))
        else:
            missing_file_cols.append(file_col)

    # Nếu không map được cột nào thì không xóa, không import
    if len(data_for_insert) == 0:
        return 0, [], missing_file_cols

    # TRƯỜNG HỢP ĐẶC BIỆT (bảng sao kê): người dùng có thể map CÙNG MỘT cột
    # file (vd "Số tiền GD") cho cả Income lẫn Outcome, phân biệt bằng dấu +/-:
    #   - giá trị bắt đầu bằng "+" (hoặc không dấu) -> vào Income
    #   - giá trị bắt đầu bằng "-"                  -> vào Outcome
    # và BỎ dấu +/- khi lưu.
    income_src = mapping.get("Income")
    outcome_src = mapping.get("Outcome")
    if (
        income_src and outcome_src
        and normalize_column_name(income_src) == normalize_column_name(outcome_src)
        and "Income" in data_for_insert and "Outcome" in data_for_insert
    ):
        shared = data_for_insert["Income"]  # cùng một cột, đang chứa cả dấu
        signed_parts = shared.map(_split_signed_amount)
        # phần tử là (is_negative, magnitude_str); None nếu ô trống
        data_for_insert["Income"] = signed_parts.map(
            lambda t: t[1] if t is not None and t[0] is False else None
        )
        data_for_insert["Outcome"] = signed_parts.map(
            lambda t: t[1] if t is not None and t[0] is True else None
        )

    # Tạo DataFrame mới với các cột đã được đổi tên sang tên cột DB
    df_to_insert = pd.DataFrame(data_for_insert)

    # QUAN TRỌNG — tránh mất chính xác với mã dài (OTA_id 16 chữ số):
    # pandas hay đọc cột mã thành số thực (float) khi cột có ô trống. Khi ghi
    # một số thực vào cột TEXT, SQLite tự chuyển float -> text bằng 15 chữ số
    # có nghĩa, làm tròn mất chữ số cuối (ví dụ ...089 -> ...090). Để tránh,
    # ta tự chuyển mỗi ô số thực thành CHUỖI số nguyên chính xác trước khi ghi:
    #   - số thực nguyên (vd 1616328635692089.0) -> "1616328635692089"
    #   - ô trống / NaN -> None
    #   - số thực có phần lẻ -> giữ nguyên giá trị
    #   - các kiểu khác (chuỗi, ngày dạng text...) -> giữ nguyên
    df_to_insert = df_to_insert.apply(
        lambda col: col.map(_cell_to_exact_text)
    )

    conn = get_connection()
    # Bước 1: xóa sạch dữ liệu cũ trong bảng (giữ nguyên cấu trúc bảng)
    conn.execute(f'DELETE FROM "{table_name}"')
    conn.commit()
    # Bước 2: thêm dữ liệu mới vào (append vào bảng vừa được xóa rỗng)
    df_to_insert.to_sql(table_name, conn, if_exists="append", index=False)
    conn.close()

    return len(df_to_insert), mapped_pairs, missing_file_cols


def _cell_to_exact_text(value):
    """
    Chuyển một ô dữ liệu sang dạng chuỗi chính xác để ghi vào DB (cột TEXT),
    tránh việc SQLite làm tròn số thực thành 15 chữ số có nghĩa.
      - NaN / None / chuỗi rỗng -> None
      - số thực nguyên (vd 1616328635692089.0) -> "1616328635692089" (không mất số)
      - số thực có phần lẻ -> chuỗi giữ nguyên giá trị
      - kiểu khác (chuỗi, ngày dạng text) -> giữ nguyên, không đụng tới
    """
    if value is None:
        return None
    # NaN của pandas/numpy
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))  # số nguyên chính xác, không có ".0" hay e+15
        return repr(value)          # số có phần lẻ -> giữ nguyên giá trị
    return value


def _split_signed_amount(value):
    """
    Tách một giá trị "có dấu" (vd "+4,559,000", "-1,123,000", hay số âm/dương)
    thành tuple (is_negative, magnitude_str):
      - is_negative = False -> khoản THU (Income)
      - is_negative = True  -> khoản CHI (Outcome)
      - magnitude_str = phần số đã BỎ dấu +/- (giữ nguyên dấu phẩy nếu có)
    Ô trống / không đọc được -> trả về None.
    Quy ước: không có dấu thì coi như dương (vào Income).
    """
    if value is None:
        return None
    # Số thực/nguyên: xét dấu trực tiếp
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, (int, float)):
        is_negative = value < 0
        return (is_negative, _cell_to_exact_text(abs(value)))
    # Chuỗi
    s = str(value).strip()
    if s == "" or s.lower() in ("none", "nan"):
        return None
    if s.startswith("-"):
        return (True, s[1:].strip())
    if s.startswith("+"):
        return (False, s[1:].strip())
    return (False, s)  # không dấu -> coi như khoản thu (Income)


# ---------------------------------------------------------------------------
# CHỨC NĂNG: TẠO BẢNG DOANH THU (total_income_by_rooms)
# ---------------------------------------------------------------------------

# Tên bảng doanh thu (cố định)
REVENUE_TABLE_NAME = "total_income_by_rooms"


def revenue_table_exists():
    """Kiểm tra bảng doanh thu đã tồn tại trong database hay chưa. Trả về True/False."""
    conn = get_connection()
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (REVENUE_TABLE_NAME,),
    ).fetchone()
    conn.close()
    return row is not None


def revenue_table_has_data():
    """
    Kiểm tra bảng doanh thu hiện CÓ dữ liệu hay không.
    Dùng để quyết định có cần hỏi xác nhận trước khi tạo lại hay không.
    Trả về True nếu bảng tồn tại VÀ có ít nhất 1 dòng.
    """
    if not revenue_table_exists():
        return False
    conn = get_connection()
    count = conn.execute(f'SELECT COUNT(*) FROM "{REVENUE_TABLE_NAME}"').fetchone()[0]
    conn.close()
    return count > 0


def build_day_column_names(num_days, month):
    """
    Tạo danh sách tên cột ngày, dạng "1/<tháng>", "2/<tháng>", ..., "<num_days>/<tháng>".
    Ví dụ num_days=30, month=6 -> ["1/6", "2/6", ..., "30/6"].
    """
    return [f"{day}/{month}" for day in range(1, num_days + 1)]


def create_revenue_table(room_names, num_days, month, year):
    """
    Tạo MỚI bảng doanh thu total_income_by_rooms.

    LƯU Ý: hàm này luôn XÓA bảng cũ (nếu có) rồi tạo lại từ đầu. Việc hỏi xác
    nhận người dùng được làm ở phần giao diện (app.py) TRƯỚC KHI gọi hàm này.

    Tham số "year" (năm) được lưu lại vào bảng phụ revenue_table_meta để bước
    fill doanh thu sau này biết các cột ngày thuộc năm nào (vì tên cột chỉ có
    ngày/tháng, không có năm).

    Cấu trúc bảng:
      - Cột đầu tiên: "room_name"
      - Tiếp theo là các cột ngày: "1/<tháng>", "2/<tháng>", ... theo num_days
    Các dòng được tạo:
      - Mỗi room_name một dòng: chỉ điền cột "room_name", các cột ngày để None.
      - Thêm 1 dòng "CHI" và 1 dòng "TỔNG" ở cuối (cũng chỉ điền room_name).
    """
    day_columns = build_day_column_names(num_days, month)

    conn = get_connection()
    cursor = conn.cursor()

    # Bước 1: xóa bảng cũ nếu có, để tạo lại hoàn toàn theo lựa chọn mới
    cursor.execute(f'DROP TABLE IF EXISTS "{REVENUE_TABLE_NAME}"')

    # Bước 2: tạo bảng mới. Cột room_name là TEXT, các cột ngày để TEXT cho đơn giản.
    column_definitions = '"room_name" TEXT'
    for day_col in day_columns:
        column_definitions += f', "{day_col}" TEXT'
    cursor.execute(f'CREATE TABLE "{REVENUE_TABLE_NAME}" ({column_definitions})')

    # Bước 3: tạo danh sách các dòng cần thêm (theo thứ tự):
    #   các room_name, rồi "CHI", rồi "TỔNG"
    all_row_labels = list(room_names) + ["CHI", "TỔNG"]

    # Mỗi dòng chỉ điền cột room_name; các cột ngày để None (NULL).
    # Câu lệnh INSERT chỉ ghi vào cột room_name, nên các cột còn lại tự là NULL.
    for label in all_row_labels:
        cursor.execute(
            f'INSERT INTO "{REVENUE_TABLE_NAME}" ("room_name") VALUES (?)',
            (label,),
        )

    conn.commit()
    conn.close()

    # Lưu lại năm để bước fill dùng
    set_revenue_meta_year(year)


def set_revenue_meta_year(year):
    """Lưu (hoặc cập nhật) năm của bảng doanh thu vào bảng phụ revenue_table_meta."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        'CREATE TABLE IF NOT EXISTS "revenue_table_meta" '
        '("key" TEXT PRIMARY KEY, "value" TEXT)'
    )
    # INSERT OR REPLACE: nếu đã có dòng key='year' thì ghi đè, chưa có thì thêm mới
    cursor.execute(
        'INSERT OR REPLACE INTO "revenue_table_meta" ("key", "value") VALUES (?, ?)',
        ("year", str(year)),
    )
    conn.commit()
    conn.close()


def get_revenue_meta_year():
    """
    Đọc năm đã lưu của bảng doanh thu.
    Trả về số nguyên (năm), hoặc None nếu chưa lưu (ví dụ bảng tạo từ trước
    khi có tính năng lưu năm).
    """
    conn = get_connection()
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='revenue_table_meta'"
    ).fetchone()
    if row is None:
        conn.close()
        return None
    result = conn.execute(
        'SELECT "value" FROM "revenue_table_meta" WHERE "key"=?', ("year",)
    ).fetchone()
    conn.close()
    if result is None:
        return None
    return int(result[0])


def get_revenue_table_data():
    """Đọc toàn bộ dữ liệu bảng doanh thu, trả về DataFrame."""
    conn = get_connection()
    df = pd.read_sql_query(f'SELECT * FROM "{REVENUE_TABLE_NAME}"', conn)
    conn.close()
    return df


# Các dòng KHÔNG phải phòng (bỏ qua khi fill)
NON_ROOM_LABELS = ["CHI", "TỔNG"]


def fill_ezcloud_booking_ids(year):
    """
    Điền Ezcloud_booking_id vào bảng doanh thu.

    Với mỗi ô [phòng, ngày]:
      - Tìm trong ezcloud_booking_list các booking CÙNG phòng và "phủ" đêm đó
        (theo luật trong matching.night_matches).
      - Nếu có booking khớp: ghi danh sách Ezcloud_booking_id (nối bằng ", ").
      - Bỏ qua 2 dòng CHI và TỔNG.

    Trả về:
      (số ô đã điền,
       danh sách ô có NHIỀU hơn 1 booking để rà soát — mỗi phần tử là
         (room_name, tên_cột_ngày, [danh sách id]),
       danh sách giá trị ngày giờ KHÔNG đọc được — để rà soát)
    """
    revenue_df = get_revenue_table_data()
    booking_df = get_table_data("ezcloud_booking_list")

    # Các cột ngày = mọi cột trừ "room_name"
    day_columns = [c for c in revenue_df.columns if c != "room_name"]

    # --- Chuẩn bị dữ liệu booking: gom theo phòng (đã chuẩn hóa) ---
    # Mỗi phần tử: {"id":..., "checkin": datetime|None, "checkout": datetime|None}
    bookings_by_room = {}
    parse_errors = []  # các chuỗi ngày giờ không đọc được

    for _, b in booking_df.iterrows():
        norm_room = matching.normalize_room(b["Room_number"])

        checkin = matching.parse_datetime(b["Checkin_date"])
        checkout = matching.parse_datetime(b["Checkout_date"])

        # Ghi nhận các giá trị có nội dung nhưng không đọc được (để báo cho user)
        if checkin is None and str(b["Checkin_date"]).strip() not in ("", "None", "nan"):
            parse_errors.append(str(b["Checkin_date"]))
        if checkout is None and str(b["Checkout_date"]).strip() not in ("", "None", "nan"):
            parse_errors.append(str(b["Checkout_date"]))

        bookings_by_room.setdefault(norm_room, []).append(
            {"id": b["Ezcloud_booking_id"], "checkin": checkin, "checkout": checkout}
        )

    filled_count = 0
    multi_matches = []  # ô có nhiều hơn 1 booking

    # --- Duyệt từng dòng (phòng) và từng cột (ngày) ---
    for row_index, row in revenue_df.iterrows():
        room_name = row["room_name"]
        if room_name in NON_ROOM_LABELS:
            continue  # bỏ qua CHI / TỔNG

        norm_room = matching.normalize_room(room_name)
        room_bookings = bookings_by_room.get(norm_room, [])
        if not room_bookings:
            continue  # phòng này không có booking nào

        for day_col in day_columns:
            # Tên cột dạng "15/5" -> ngày = 15, tháng = 5
            day_str, month_str = day_col.split("/")
            day = int(day_str)
            month = int(month_str)

            # Tìm tất cả booking khớp đêm này
            matched_ids = []
            for bk in room_bookings:
                if matching.night_matches(bk["checkin"], bk["checkout"], year, month, day):
                    matched_ids.append(str(bk["id"]))

            if matched_ids:
                revenue_df.at[row_index, day_col] = ", ".join(matched_ids)
                filled_count += 1
                if len(matched_ids) > 1:
                    multi_matches.append((room_name, day_col, matched_ids))

    # --- Ghi bảng doanh thu đã điền trở lại database (giữ nguyên cấu trúc) ---
    conn = get_connection()
    conn.execute(f'DELETE FROM "{REVENUE_TABLE_NAME}"')
    conn.commit()
    revenue_df.to_sql(REVENUE_TABLE_NAME, conn, if_exists="append", index=False)
    conn.close()

    # Bỏ trùng trong danh sách lỗi parse cho gọn
    parse_errors = sorted(set(parse_errors))

    return filled_count, multi_matches, parse_errors


# ---------------------------------------------------------------------------
# CHỨC NĂNG: FILL TIỀN PHÒNG OTA (bước tạm — chỉ lấy tiền phòng từ chứng từ OTA)
# ---------------------------------------------------------------------------

# Cấu hình cho từng hãng OTA:
#   key   = giá trị Company (đã chuyển về chữ thường) trong ezcloud_booking_list
#   value = {
#       "table":    tên bảng chứng từ OTA,
#       "id_col":   tên cột chứa mã booking trong bảng đó,
#       "income_col": tên cột dùng làm TIỀN PHÒNG,
#       "divide_by_1_15": True nếu cần tính Net = Gross / 1.15 (chỉ Expedia),
#   }
# Muốn đổi quy tắc lấy tiền phòng của một hãng thì sửa ở đây.
OTA_CONFIG = {
    "agoda": {
        "table": "agoda_booking_list",
        "id_col": "Agoda_booking_id",
        "income_col": "Net_income",
        "divide_by_1_15": False,
    },
    "booking": {
        "table": "booking_booking_list",
        "id_col": "Booking_booking_id",
        "income_col": "Gross_income",  # Booking lấy Gross_income
        "divide_by_1_15": False,
    },
    "traveloka": {
        "table": "traveloka_booking_list",
        "id_col": "Traveloka_booking_id",
        "income_col": "Net_income",
        "divide_by_1_15": False,
    },
    "expedia": {
        "table": "expedia_booking_list",
        "id_col": "Expedia_booking_id",
        "income_col": "Gross_income",  # lấy Gross rồi chia 1.15 ra Net
        "divide_by_1_15": True,
    },
    "ctrip": {
        "table": "ctrip_booking_list",
        "id_col": "Ctrip_booking_id",
        "income_col": "Net_income",
        "divide_by_1_15": False,
    },
}


def _get_ota_room_revenue(company_raw, ota_id_raw):
    """
    Lấy tiền phòng OTA cho một booking.

    Tham số:
      - company_raw: giá trị cột Company (chưa chuẩn hóa)
      - ota_id_raw:  giá trị cột OTA_id (chưa chuẩn hóa)

    Trả về (revenue, error):
      - Nếu thành công: (số tiền phòng dạng float, None)
      - Nếu lỗi:        (None, "mô tả lỗi")

    Riêng Expedia: tính Net = Gross / 1.15 và GHI ĐÈ vào cột Net_income của
    row tương ứng trong bảng expedia_booking_list.
    """
    company = str(company_raw).strip().lower()

    # Company không thuộc 5 OTA đã biết -> lỗi
    if company not in OTA_CONFIG:
        return None, f"Company lạ: '{company_raw}'"

    config = OTA_CONFIG[company]
    target_id = matching.normalize_id(ota_id_raw)
    if target_id is None:
        return None, f"OTA_id rỗng (company={company})"

    # Đọc bảng OTA, tìm row có id khớp (so theo dạng đã chuẩn hóa)
    ota_df = get_table_data(config["table"])
    id_col = config["id_col"]

    matched_row = None
    for _, row in ota_df.iterrows():
        if matching.normalize_id(row[id_col]) == target_id:
            matched_row = row
            break

    if matched_row is None:
        return None, f"Không tìm thấy OTA_id '{target_id}' trong {config['table']}"

    # Lấy tiền phòng từ cột income tương ứng
    raw_income = matched_row[config["income_col"]]

    if config["divide_by_1_15"]:
        # Expedia: lấy Gross rồi chia 1.15 ra Net, ghi đè vào DB
        gross = matching.normalize_number(raw_income)
        if gross is None:
            return None, f"Không đọc được Gross_income của Expedia id '{target_id}'"
        net = round(gross * 0.85)
        _update_expedia_net_income(target_id, net)
        return net, None
    else:
        revenue = matching.normalize_number(raw_income)
        if revenue is None:
            return None, (
                f"Không đọc được {config['income_col']} "
                f"trong {config['table']} cho id '{target_id}'"
            )
        return revenue, None


def _update_expedia_net_income(target_id, net_value):
    """Ghi đè giá trị Net_income (đã tính = Gross/1.15) vào row Expedia tương ứng."""
    conn = get_connection()
    # Đọc lại để tìm đúng row theo id đã chuẩn hóa, rồi update theo giá trị id gốc
    ota_df = pd.read_sql_query('SELECT rowid, * FROM "expedia_booking_list"', conn)
    for _, row in ota_df.iterrows():
        if matching.normalize_id(row["Expedia_booking_id"]) == target_id:
            conn.execute(
                'UPDATE "expedia_booking_list" SET "Net_income"=? WHERE rowid=?',
                (str(net_value), row["rowid"]),
            )
            break
    conn.commit()
    conn.close()


def fill_ota_room_revenue():
    """
    Bước TẠM: điền tiền phòng OTA vào bảng doanh thu.

    Cách làm:
      - Gom tất cả ezcloud_booking_id distinct đang có trong bảng doanh thu.
      - Với mỗi booking_id: tra Company/OTA_id trong ezcloud_booking_list.
          + Nếu Company rỗng (walk-in): giữ nguyên, bỏ qua ở bước này.
          + Nếu có Company: lấy tiền phòng OTA. Thành công -> ghi tạm
            "[id, tiền_phòng, company]" vào MỌI ô chứa booking_id đó.
            Lỗi -> ghi "[\"Không tìm thấy chứng từ\", id]" vào mọi ô đó.
      - CHƯA đánh dấu is_done (để bước cuối).

    Trả về:
      (số booking OTA xử lý thành công,
       danh sách lỗi - mỗi phần tử (booking_id, mô tả lỗi))
    """
    revenue_df = get_revenue_table_data()
    booking_df = get_table_data("ezcloud_booking_list")
    day_columns = [c for c in revenue_df.columns if c != "room_name"]

    # --- Tạo tra cứu: booking_id (chuẩn hóa) -> (Company, OTA_id) ---
    booking_info = {}
    for _, b in booking_df.iterrows():
        bid = matching.normalize_id(b["Ezcloud_booking_id"])
        if bid is not None:
            booking_info[bid] = {
                "company": b["Company"],
                "ota_id": b["OTA_id"],
            }

    # --- Gom các booking_id distinct đang xuất hiện trong bảng doanh thu ---
    # Mỗi ô có thể chứa nhiều id ngăn bởi ", "; bỏ qua dòng CHI/TỔNG.
    ids_in_cell = {}   # (row_index, day_col) -> danh sách id gốc trong ô
    distinct_ids = set()
    for row_index, row in revenue_df.iterrows():
        if row["room_name"] in NON_ROOM_LABELS:
            continue
        for day_col in day_columns:
            value = row[day_col]
            if value is None or str(value).strip() == "" or str(value).lower() == "nan":
                continue
            ids = [piece.strip() for piece in str(value).split(",") if piece.strip() != ""]
            if ids:
                ids_in_cell[(row_index, day_col)] = ids
                for one_id in ids:
                    distinct_ids.add(matching.normalize_id(one_id))

    # --- Xử lý từng booking_id distinct đúng một lần ---
    # Kết quả mỗi id -> chuỗi sẽ hiển thị trong ô (hoặc None nếu walk-in: giữ nguyên)
    rendered = {}
    errors = []
    success_count = 0

    for bid in distinct_ids:
        info = booking_info.get(bid)
        if info is None:
            # booking_id trong bảng doanh thu nhưng không có trong ezcloud_booking_list
            errors.append((bid, "Không thấy trong ezcloud_booking_list"))
            rendered[bid] = f'["Không tìm thấy chứng từ", {bid}]'
            continue

        company_raw = info["company"]
        company_str = str(company_raw).strip()

        # Walk-in (Company rỗng) -> giữ nguyên id, không xử lý ở bước này
        if company_str == "" or company_str.lower() in ("none", "nan"):
            rendered[bid] = None  # None = giữ nguyên id gốc
            continue

        revenue, error = _get_ota_room_revenue(company_raw, info["ota_id"])
        if error is not None:
            errors.append((bid, error))
            rendered[bid] = f'["Không tìm thấy chứng từ", {bid}]'
        else:
            company_norm = company_str.lower()
            rendered[bid] = f"[{bid}, {revenue}, {company_norm}]"
            success_count += 1

    # --- Ghi kết quả vào từng ô (gộp nhiều id trong cùng ô bằng ", ") ---
    for (row_index, day_col), ids in ids_in_cell.items():
        pieces = []
        for one_id in ids:
            norm = matching.normalize_id(one_id)
            r = rendered.get(norm)
            if r is None:
                # walk-in hoặc không đổi -> giữ nguyên id gốc
                pieces.append(one_id)
            else:
                pieces.append(r)
        revenue_df.at[row_index, day_col] = ", ".join(pieces)

    # --- Ghi bảng doanh thu trở lại DB ---
    conn = get_connection()
    conn.execute(f'DELETE FROM "{REVENUE_TABLE_NAME}"')
    conn.commit()
    revenue_df.to_sql(REVENUE_TABLE_NAME, conn, if_exists="append", index=False)
    conn.close()

    return success_count, errors
