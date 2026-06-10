# reconciliation.py
# ---------------------------------------------------------------------------
# File này lo phần PHỨC TẠP NHẤT: với mỗi ezcloud_booking_id trên bảng doanh
# thu, đối chiếu các khoản thanh toán với chứng từ (bảng OTA, sao kê ngân
# hàng, tiền mặt), rồi tính doanh thu từng ô và điền theo format cuối.
#
# Cách tổ chức để dễ đọc/sửa:
#   - matching.py  : các hàm chuẩn hóa (số phòng, mã, số tiền, ngày giờ) + luật khớp đêm.
#   - database.py  : truy cập DB (đọc bảng, cấu hình OTA_CONFIG, bảng doanh thu).
#   - reconciliation.py (file này): toàn bộ quy trình đối chiếu + tính tiền.
#
# Nguyên tắc an toàn về is_done:
#   - Đầu mỗi lần chạy: coi như toàn bộ is_done = 0 (reset).
#   - Trong khi chạy: mỗi booking gom các thay đổi is_done vào "tentative".
#     Chỉ khi booking HỢP LỆ mới ghi nhận (commit) các thay đổi đó; booking
#     lỗi thì bỏ hết (không đụng is_done), và các dòng sao kê đã tạm chiếm
#     được nhả lại cho booking khác.
#   - Cuối cùng mới ghi is_done xuống DB một lần.
# ---------------------------------------------------------------------------

from datetime import date, timedelta

import re

import pandas as pd

import database
import matching


# Hằng số dùng chung
CREDIT_LIMIT = 1000          # sai số cho phép khi so tiền: |a - b| <= 1000
BANK_WINDOW_HOURS = 48       # sao kê TRƯỚC Pay_time được chấp nhận trong vòng 48 giờ
BANK_AFTER_MINUTES = 30      # sao kê SAU Pay_time được chấp nhận trong vòng 30 phút
                             # (ngân hàng đôi khi ghi nhận trễ hơn lúc hệ thống ghi)
WEEKEND_SURCHARGE = 1.05     # đêm thứ 6 / thứ 7 đắt hơn ngày thường 5% (khách qua Company)

# Các OTA có thanh toán dạng "Công nợ" (booking thì xử lý như walk-in)
OTA_CREDIT_COMPANIES = ["agoda", "traveloka", "expedia", "ctrip"]

# Các bảng có cột is_done cần reset/ghi
IS_DONE_TABLES = [
    "ezcloud_booking_list",
    "ezcloud_payment_list",
    "agoda_booking_list",
    "booking_booking_list",
    "traveloka_booking_list",
    "expedia_booking_list",
    "ctrip_booking_list",
    "personal_bank_transactions",
    "company_bank_transactions",
    "cash_transactions",
]


def _norm_text(value):
    """Chuẩn hóa văn bản để so sánh: về chữ thường, bỏ khoảng trắng thừa."""
    if value is None:
        return ""
    s = str(value).strip().lower()
    if s in ("none", "nan"):
        return ""
    return s


def _read_with_rowid(table_name):
    """Đọc một bảng kèm cột rowid (khóa ngầm của SQLite) để sau này UPDATE đúng dòng."""
    conn = database.get_connection()
    df = pd.read_sql_query(f'SELECT rowid AS _rowid, * FROM "{table_name}"', conn)
    conn.close()
    return df


def _content_is_company(content_lower):
    """True nếu nội dung chuyển khoản nhắc tới công ty (chứa 'cty' hoặc 'công ty')."""
    return ("cty" in content_lower) or ("công ty" in content_lower)


def _format_money(value):
    """Hiển thị số tiền gọn: số nguyên thì bỏ phần thập phân (50000.0 -> '50000')."""
    if value is None:
        return "(không đọc được số tiền)"
    if value == int(value):
        return str(int(value))
    return str(value)


def _format_pay_date(dt):
    """Hiển thị ngày dạng d/m (vd 19/5). Trả về '(không đọc được thời gian)' nếu None."""
    if dt is None:
        return "(không đọc được thời gian)"
    return f"{dt.day}/{dt.month}"


def _format_pay_datetime(dt):
    """Hiển thị ngày giờ dạng d/m/Y HH:MM. Trả về thông báo nếu None."""
    if dt is None:
        return "(không đọc được thời gian)"
    return dt.strftime("%d/%m/%Y %H:%M")


def _parse_month_day(text):
    """
    Tách (tháng, ngày) từ một giá trị ngày, để đối chiếu tiền mặt.
    Hỗ trợ:
      - dạng đầy đủ có năm: "16/05/2026", "16/05/2026 23:37:08", "2026-05-16..."
      - dạng NGẮN không năm: "16/5" (ngày/tháng) — vì bảng cash_transactions
        thực tế lưu kiểu này.
    Trả về tuple (tháng, ngày), hoặc None nếu không đọc được.
    So sánh theo tháng+ngày (bỏ năm) vì cột Time tiền mặt có thể không có năm.
    """
    if text is None:
        return None
    s = str(text).strip()
    if s == "" or s.lower() in ("none", "nan"):
        return None
    # Thử dạng đầy đủ (có năm) trước
    dt = matching.parse_datetime_full(s)
    if dt is not None:
        return (dt.month, dt.day)
    # Dạng ngắn "ngày/tháng" như "16/5"
    parts = s.split("/")
    if len(parts) >= 2:
        try:
            day = int(parts[0])
            month = int(parts[1])
            return (month, day)
        except ValueError:
            return None
    return None


# ---------------------------------------------------------------------------
# HÀM CHÍNH
# ---------------------------------------------------------------------------
def fill_revenue_full():
    """
    Chạy toàn bộ quy trình đối chiếu + tính doanh thu + điền format cuối vào
    bảng doanh thu. Yêu cầu: bảng doanh thu đã được fill ezcloud_booking_id ở
    bước trước (các ô chứa id dạng "E1" hoặc "E1, E2").

    Trả về một dict tóm tắt:
      {
        "ok": số booking hợp lệ,
        "errors": [(booking_id, lý do), ...],
        "credit_card": [booking_id, ...],
      }
    """
    year = database.get_revenue_meta_year()

    # Tạo lại các ô id thô từ đầu (ezcloud_booking_id), để hàm này luôn làm việc
    # trên dữ liệu id thô — chạy lại nhiều lần vẫn cho kết quả ổn định, không bị
    # phân tích nhầm format cuối của lần chạy trước.
    database.fill_ezcloud_booking_ids(year)

    # --- Đọc toàn bộ dữ liệu cần dùng (kèm rowid) ---
    revenue_df = database.get_revenue_table_data()
    day_columns = [c for c in revenue_df.columns if c != "room_name"]

    booking_df = _read_with_rowid("ezcloud_booking_list")
    payment_df = _read_with_rowid("ezcloud_payment_list")

    # Gom booking theo id (một booking có thể nhiều dòng = nhiều phòng)
    booking_by_id = {}
    for _, r in booking_df.iterrows():
        bid = matching.normalize_id(r["Ezcloud_booking_id"])
        if bid is None:
            continue
        booking_by_id.setdefault(bid, []).append({
            "rowid": r["_rowid"],
            "company": r["Company"],
            "ota_id": r["OTA_id"],
            "room": matching.normalize_room(r["Room_number"]),
            "checkin": matching.parse_datetime(r["Checkin_date"]),
            "checkout": matching.parse_datetime(r["Checkout_date"]),
            "room_price": matching.normalize_number(r["Room_price"]),
            "service_price": matching.normalize_number(r["Service_price"]),
            "discount": matching.normalize_number(r["Discount"]),
        })

    # Gom các khoản thanh toán theo booking id
    payment_by_id = {}
    for _, r in payment_df.iterrows():
        bid = matching.normalize_id(r["Ezcloud_booking_id"])
        if bid is None:
            continue
        payment_by_id.setdefault(bid, []).append({
            "rowid": r["_rowid"],
            "pay_type": _norm_text(r["Pay_type"]),
            "pay_amount": matching.normalize_number(r["Pay_amount"]),
            "pay_time": matching.parse_datetime_full(r["Pay_time"]),
            "content": _norm_text(r["Content"]),
        })

    # --- Bảng sao kê / tiền mặt (đọc 1 lần thành danh sách dict) ---
    def load_bank(table_name):
        df = _read_with_rowid(table_name)
        rows = []
        for _, r in df.iterrows():
            rows.append({
                "rowid": r["_rowid"],
                "time": matching.parse_datetime_full(r["Time"]),
                "income": matching.normalize_number(r["Income"]),
            })
        return rows

    personal_bank = load_bank("personal_bank_transactions")
    company_bank = load_bank("company_bank_transactions")

    cash_df = _read_with_rowid("cash_transactions")
    cash_rows = []
    for _, r in cash_df.iterrows():
        cash_rows.append({
            "rowid": r["_rowid"],
            "month_day": _parse_month_day(r["Time"]),  # (tháng, ngày), bỏ năm
            "income": matching.normalize_number(r["Income"]),
        })

    # --- Chỉ mục các bảng OTA ---
    # Mỗi công ty có 2 chỉ mục:
    #   index       : id CHÍNH XÁC -> {rowid, net}   (ưu tiên khớp chính xác)
    #   index_15sig : id làm tròn 15 chữ số -> {rowid, net}  (dự phòng khi dữ
    #                 liệu cũ bị Excel làm tròn mất chữ số cuối)
    ota_indexes = {}
    expedia_net_updates = {}  # rowid -> net (để ghi đè Net_income của Expedia)
    for company in OTA_CREDIT_COMPANIES:
        cfg = database.OTA_CONFIG[company]
        df = _read_with_rowid(cfg["table"])
        index = {}
        index_15sig = {}
        for _, r in df.iterrows():
            raw_id = r[cfg["id_col"]]
            nid = matching.normalize_id(raw_id)
            if nid is None:
                continue
            income_val = r[cfg["income_col"]]
            if company == "expedia":
                gross = matching.normalize_number(income_val)
                net = round(gross * 0.85) if gross is not None else None
                if net is not None:
                    expedia_net_updates[r["_rowid"]] = net
            else:
                net = matching.normalize_number(income_val)
            entry = {"rowid": r["_rowid"], "net": net}
            index[nid] = entry
            fz = matching.id_15sig(raw_id)
            if fz is not None:
                # Chỉ giữ entry đầu tiên cho mỗi khóa 15-chữ-số (tránh ghi đè
                # khi hiếm hoi có 2 mã trùng 15 chữ số đầu).
                index_15sig.setdefault(fz, entry)
        ota_indexes[company] = {"exact": index, "fuzzy": index_15sig}

    # --- Trạng thái is_done ---
    # done_rowids: các dòng đã được commit là is_done=true (từ booking hợp lệ)
    done_rowids = {t: set() for t in IS_DONE_TABLES}

    # ===================================================================
    # STAGE A: xử lý từng booking_id distinct -> quyết định hợp lệ/lỗi
    # ===================================================================
    # Gom các booking_id distinct xuất hiện trong bảng doanh thu
    distinct_ids = set()
    for _, row in revenue_df.iterrows():
        if row["room_name"] in database.NON_ROOM_LABELS:
            continue
        for dc in day_columns:
            for one_id in _parse_cell_ids(row[dc]):
                distinct_ids.add(one_id)

    booking_result = {}  # bid -> {"status", "company", "pay_type", "reason"}

    # Duyệt theo thứ tự CỐ ĐỊNH (sắp xếp) để kết quả ổn định giữa các lần chạy.
    # (Nếu duyệt theo thứ tự ngẫu nhiên của set, việc tranh chấp dòng sao kê
    #  giữa các khoản cùng số tiền sẽ cho kết quả khác nhau mỗi lần chạy.)
    for bid in sorted(distinct_ids):
        status, reason, tentative = _process_one_booking(
            bid, booking_by_id, payment_by_id,
            ota_indexes, personal_bank, company_bank, cash_rows,
            done_rowids,
        )

        if status == "ok":
            # Commit các thay đổi is_done của booking hợp lệ
            for table_name, rowid in tentative:
                done_rowids[table_name].add(rowid)
            booking_result[bid] = {
                "status": "ok",
                "company": _company_label(booking_by_id[bid][0]["company"]),
                "pay_type": _pay_type_label(payment_by_id.get(bid, [])),
            }
        elif status == "credit_card":
            booking_result[bid] = {"status": "credit_card"}
        else:  # error
            booking_result[bid] = {"status": "error", "reason": reason}

    # ===================================================================
    # STAGE B: tính doanh thu từng ô và điền format cuối
    # ===================================================================
    # Lưu danh sách id (giữ thứ tự) của từng ô để dùng lại
    row_cell_ids = {}  # row_index -> list (theo day_columns) của list id
    for row_index, row in revenue_df.iterrows():
        if row["room_name"] in database.NON_ROOM_LABELS:
            continue
        row_cell_ids[row_index] = [_parse_cell_ids(row[dc]) for dc in day_columns]

    for row_index, per_day_ids in row_cell_ids.items():
        room_norm = matching.normalize_room(revenue_df.at[row_index, "room_name"])

        for col_i, dc in enumerate(day_columns):
            ids_here = per_day_ids[col_i]
            if not ids_here:
                continue

            pieces = []
            for bid in ids_here:
                result = booking_result.get(bid, {"status": "error", "reason": "?"})

                if result["status"] == "credit_card":
                    pieces.append(f'["Thanh toán bằng Thẻ tín dụng", {bid}]')
                elif result["status"] == "error":
                    pieces.append(f'["Không tìm thấy chứng từ", {bid}]')
                else:
                    # Tìm đúng row [booking_id, phòng] cho đêm của ô hiện tại
                    brow = _find_booking_row(booking_by_id, bid, room_norm, dc, year)

                    # Số đêm n và đêm muộn nhất được tính từ Checkin/Checkout của
                    # row đó (theo luật khớp đêm). Nếu thiếu/đọc không được ngày
                    # giờ -> coi ô này là lỗi (theo thỏa thuận).
                    nights = []
                    if brow is not None and brow["checkin"] is not None and brow["checkout"] is not None:
                        nights = _nights_covered(brow["checkin"], brow["checkout"])

                    if not nights:
                        pieces.append(f'["Không tìm thấy chứng từ", {bid}]')
                    else:
                        n = len(nights)
                        latest_night = nights[-1]  # đêm muộn nhất row này phủ

                        # Ngày của ô hiện tại (năm lấy từ meta, ngày/tháng từ tên cột)
                        day_str, month_str = dc.split("/")
                        current_night = date(year, int(month_str), int(day_str))
                        is_last = (current_night == latest_night)

                        room_price = brow["room_price"] if brow["room_price"] is not None else 0
                        discount = brow["discount"] if brow["discount"] is not None else 0
                        service = brow["service_price"] if brow["service_price"] is not None else 0
                        base = room_price - discount  # tổng tiền phòng sau giảm giá

                        if result["company"] == "walk-in":
                            # Khách lẻ (Company rỗng): chia đều cho các đêm
                            value = base / n
                        else:
                            # Khách có Company (booking / agoda / traveloka /
                            # expedia / ctrip): đêm thứ 6, thứ 7 đắt hơn ngày
                            # thường WEEKEND_SURCHARGE lần (5%).
                            # Giải: d*x + d*WEEKEND_SURCHARGE*y = base
                            #   x = số đêm thường, y = số đêm thứ 6/thứ 7 trong chặng
                            y = sum(1 for nd in nights if _is_weekend(nd))
                            x = n - y
                            d = base / (x + WEEKEND_SURCHARGE * y)
                            value = d * WEEKEND_SURCHARGE if _is_weekend(current_night) else d

                        if is_last:
                            value += service
                        value = round(value)

                        pieces.append(f"[{bid}, {value}, {result['company']}, {result['pay_type']}]")

            revenue_df.at[row_index, dc] = ", ".join(pieces)

    # ===================================================================
    # GHI KẾT QUẢ XUỐNG DB
    # ===================================================================
    conn = database.get_connection()
    # 1) Reset toàn bộ is_done = 0
    for t in IS_DONE_TABLES:
        conn.execute(f'UPDATE "{t}" SET is_done=0')
    # 2) Đánh dấu is_done = 1 cho các dòng đã commit
    for t, rowids in done_rowids.items():
        for rid in rowids:
            conn.execute(f'UPDATE "{t}" SET is_done=1 WHERE rowid=?', (int(rid),))
    # 3) Ghi đè Net_income đã tính cho Expedia
    for rid, net in expedia_net_updates.items():
        conn.execute(
            'UPDATE "expedia_booking_list" SET "Net_income"=? WHERE rowid=?',
            (str(net), int(rid)),
        )
    conn.commit()

    # 4) Ghi lại bảng doanh thu đã điền format cuối
    conn.execute(f'DELETE FROM "{database.REVENUE_TABLE_NAME}"')
    conn.commit()
    revenue_df.to_sql(database.REVENUE_TABLE_NAME, conn, if_exists="append", index=False)
    conn.close()

    # --- Tóm tắt cho giao diện ---
    ok_count = sum(1 for r in booking_result.values() if r["status"] == "ok")
    errors = [(bid, r["reason"]) for bid, r in booking_result.items() if r["status"] == "error"]
    credit_card = [bid for bid, r in booking_result.items() if r["status"] == "credit_card"]
    return {"ok": ok_count, "errors": errors, "credit_card": credit_card}


# ---------------------------------------------------------------------------
# XỬ LÝ MỘT BOOKING (Stage A)
# ---------------------------------------------------------------------------
def _process_one_booking(bid, booking_by_id, payment_by_id, ota_indexes,
                         personal_bank, company_bank, cash_rows, done_rowids):
    """
    Trả về (status, reason, tentative):
      - status: "ok" | "error" | "credit_card"
      - reason: mô tả lỗi (chỉ khi error)
      - tentative: danh sách (table_name, rowid) sẽ đánh is_done nếu booking hợp lệ
    """
    booking_rows = booking_by_id.get(bid)
    if not booking_rows:
        return "error", "Không thấy trong ezcloud_booking_list", []

    company = _norm_text(booking_rows[0]["company"])
    payments = payment_by_id.get(bid, [])

    # (1) Ưu tiên: nếu có khoản Thẻ tín dụng -> skip cả booking
    for p in payments:
        if p["pay_type"] == "thẻ tín dụng":
            return "credit_card", None, []

    # tentative: các (table, rowid) sẽ đánh dấu nếu booking hợp lệ
    tentative = []
    # các dòng sao kê/tiền mặt đã tạm chiếm trong booking này (để khoản sau không lấy lại)
    tentative_bank = {
        "personal_bank_transactions": set(),
        "company_bank_transactions": set(),
        "cash_transactions": set(),
    }

    company_empty = (company == "")

    # (2) Phân loại Company
    if (not company_empty) and (company != "booking") and (company not in OTA_CREDIT_COMPANIES):
        return "error", f"Company lạ: '{booking_rows[0]['company']}'", []

    if company in OTA_CREDIT_COMPANIES:
        # --- Khách OTA: lấy Net, đối chiếu Công nợ ---
        cfg = database.OTA_CONFIG[company]
        raw_ota = booking_rows[0]["ota_id"]
        ota_id = matching.normalize_id(raw_ota)
        index = ota_indexes[company]["exact"]
        index_15sig = ota_indexes[company]["fuzzy"]

        # Ưu tiên khớp CHÍNH XÁC; nếu trượt thì thử khớp theo 15 chữ số có nghĩa
        # (dự phòng cho dữ liệu cũ bị Excel làm tròn chữ số cuối).
        entry = index.get(ota_id)
        if entry is None:
            fz = matching.id_15sig(raw_ota)
            if fz is not None:
                entry = index_15sig.get(fz)

        if entry is None:
            return "error", f"Không tìm thấy OTA_id '{ota_id}' trong {cfg['table']}", []
        net = entry["net"]
        if net is None:
            return "error", f"Không đọc được tiền phòng của {company} (id {ota_id})", []

        congno = [p for p in payments if p["pay_type"] == "công nợ"]
        if not congno:
            return "error", "Khách OTA nhưng không có khoản Công nợ", []
        total_congno = sum(p["pay_amount"] for p in congno if p["pay_amount"] is not None)
        if abs(total_congno - net) > CREDIT_LIMIT:
            return "error", f"Tổng Công nợ ({total_congno}) không khớp tiền phòng ({net})", []

        # Công nợ khớp -> đánh dấu các row Công nợ + row OTA
        for p in congno:
            tentative.append(("ezcloud_payment_list", p["rowid"]))
        tentative.append((cfg["table"], entry["rowid"]))

        # Các khoản còn lại (không phải Công nợ) -> đối chiếu sao kê
        others = [p for p in payments if p["pay_type"] != "công nợ"]
        ok, reason = _reconcile_payment_list(
            others, personal_bank, company_bank, cash_rows,
            done_rowids, tentative_bank, tentative,
        )
        if not ok:
            return "error", reason, []
        return "ok", None, tentative

    else:
        # --- Walk-in (Company rỗng) hoặc "booking": không có Công nợ ---
        if any(p["pay_type"] == "công nợ" for p in payments):
            return "error", "Khách walk-in/booking nhưng lại có khoản Công nợ", []

        ok, reason = _reconcile_payment_list(
            payments, personal_bank, company_bank, cash_rows,
            done_rowids, tentative_bank, tentative,
        )
        if not ok:
            return "error", reason, []
        return "ok", None, tentative


def _reconcile_payment_list(payments, personal_bank, company_bank, cash_rows,
                            done_rowids, tentative_bank, tentative):
    """
    Đối chiếu từng khoản thanh toán (không phải Công nợ) với sao kê/tiền mặt.
    Mỗi khoản PHẢI tìm được chứng từ; nếu một khoản không tìm được -> trả về
    (False, lý do). Nếu tất cả đều tìm được -> (True, None).

    Khi khớp: ghi (bảng, rowid) vào tentative (cả dòng sao kê lẫn dòng payment).
    """
    for p in payments:
        pay_type = p["pay_type"]

        if pay_type == "chuyển khoản nh":
            # Chọn bảng sao kê theo Content
            if _content_is_company(p["content"]):
                bank_table = "company_bank_transactions"
                bank_rows = company_bank
                ck_label = "Chuyển khoản công ty"
            else:
                bank_table = "personal_bank_transactions"
                bank_rows = personal_bank
                ck_label = "Chuyển khoản cá nhân"

            rowid = _match_bank(
                bank_rows, p["pay_time"], p["pay_amount"],
                done_rowids[bank_table], tentative_bank[bank_table],
            )
            # Dự phòng: nếu cửa sổ theo Pay_time không thấy, thử các NGÀY ghi
            # trong Content (vd "cọc 9/5") -> tìm dòng sao kê cùng ngày đó.
            if rowid is None:
                content_dates = _extract_content_dates(p["content"])
                rowid = _match_bank_by_dates(
                    bank_rows, content_dates, p["pay_amount"],
                    done_rowids[bank_table], tentative_bank[bank_table],
                )
            if rowid is None:
                return False, (
                    f"Khoản {ck_label} không tìm thấy trong sao kê. "
                    f"Trong payment_list có khoản {_format_money(p['pay_amount'])} "
                    f"lúc {_format_pay_datetime(p['pay_time'])}"
                )
            tentative_bank[bank_table].add(rowid)
            tentative.append((bank_table, rowid))
            tentative.append(("ezcloud_payment_list", p["rowid"]))

        elif pay_type == "tiền mặt":
            rowid = _match_cash(
                cash_rows, p["pay_time"], p["pay_amount"],
                done_rowids["cash_transactions"], tentative_bank["cash_transactions"],
            )
            # Dự phòng: nếu ngày của Pay_time không thấy, thử các NGÀY ghi trong Content
            if rowid is None:
                content_dates = _extract_content_dates(p["content"])
                rowid = _match_cash_by_dates(
                    cash_rows, content_dates, p["pay_amount"],
                    done_rowids["cash_transactions"], tentative_bank["cash_transactions"],
                )
            if rowid is None:
                return False, (
                    "Khoản tiền mặt không tìm thấy trong cash_transactions. "
                    f"Trong payment_list có khoản {_format_money(p['pay_amount'])} "
                    f"ngày {_format_pay_date(p['pay_time'])}"
                )
            tentative_bank["cash_transactions"].add(rowid)
            tentative.append(("cash_transactions", rowid))
            tentative.append(("ezcloud_payment_list", p["rowid"]))

        else:
            # Pay_type lạ (không thuộc các loại đã biết) -> coi như không có chứng từ
            return False, f"Hình thức thanh toán lạ: '{pay_type}'"

    return True, None


def _match_bank(bank_rows, pay_time, pay_amount, done_set, tentative_set):
    """
    Tìm dòng sao kê ngân hàng khớp với một khoản thanh toán:
      - Thời điểm sao kê nằm trong khoảng quanh Pay_time:
          + TRƯỚC Pay_time: trong vòng 48 giờ (BANK_WINDOW_HOURS)
          + SAU  Pay_time: trong vòng 30 phút (BANK_AFTER_MINUTES) — vì ngân
            hàng đôi khi ghi nhận trễ hơn thời điểm hệ thống ghi.
      - chưa is_done (chưa nằm trong done_set / tentative_set)
      - |Income - Pay_amount| <= 1000
    Trong các dòng thỏa, chọn dòng có thời điểm GẦN Pay_time nhất (chênh lệch
    tuyệt đối nhỏ nhất). Trả về rowid, hoặc None nếu không có.
    """
    if pay_time is None or pay_amount is None:
        return None

    before_limit = timedelta(hours=BANK_WINDOW_HOURS)
    after_limit = timedelta(minutes=BANK_AFTER_MINUTES)

    candidates = []
    for r in bank_rows:
        t = r["time"]
        if t is None or r["income"] is None:
            continue
        # Kiểm tra cửa sổ thời gian (trước/sau Pay_time)
        if t <= pay_time:
            if (pay_time - t) > before_limit:
                continue
        else:
            if (t - pay_time) > after_limit:
                continue
        if r["rowid"] in done_set or r["rowid"] in tentative_set:
            continue
        if abs(r["income"] - pay_amount) <= CREDIT_LIMIT:
            diff_seconds = abs((t - pay_time).total_seconds())
            candidates.append((diff_seconds, r["rowid"]))

    if not candidates:
        return None
    # Gần Pay_time nhất = chênh lệch tuyệt đối nhỏ nhất
    candidates.sort(key=lambda x: x[0])
    return candidates[0][1]


def _match_cash(cash_rows, pay_time, pay_amount, done_set, tentative_set):
    """
    Tìm dòng tiền mặt khớp: cùng NGÀY/THÁNG với Pay_time (bỏ năm vì cột Time
    tiền mặt có thể không ghi năm), chưa is_done, |Income - Pay_amount| <= 1000.
    Trả về rowid hoặc None.
    """
    if pay_time is None or pay_amount is None:
        return None
    pay_md = (pay_time.month, pay_time.day)
    for r in cash_rows:
        if r["month_day"] is None or r["income"] is None:
            continue
        if r["month_day"] != pay_md:
            continue
        if r["rowid"] in done_set or r["rowid"] in tentative_set:
            continue
        if abs(r["income"] - pay_amount) <= CREDIT_LIMIT:
            return r["rowid"]
    return None


def _extract_content_dates(content):
    """
    Trích các NGÀY ghi trong cột Content, dạng "ngày/tháng" (vd "cọc 9/5" -> 9/5).
    Trả về danh sách (tháng, ngày), giữ thứ tự xuất hiện, bỏ trùng. Chỉ nhận
    các cụm hợp lệ (ngày 1..31, tháng 1..12).
    """
    if content is None:
        return []
    results = []
    for m_ in re.finditer(r"(\d{1,2})/(\d{1,2})", str(content)):
        day = int(m_.group(1))
        month = int(m_.group(2))
        if 1 <= day <= 31 and 1 <= month <= 12:
            md = (month, day)
            if md not in results:
                results.append(md)
    return results


def _match_bank_by_dates(bank_rows, md_list, pay_amount, done_set, tentative_set):
    """
    Dự phòng cho chuyển khoản NH: tìm dòng sao kê có Time đúng vào một trong các
    NGÀY (tháng, ngày) lấy từ Content, số tiền khớp (±1000), chưa is_done.
    KHÔNG áp giới hạn 48h/30 phút (vì Content chỉ ghi ngày). Thử lần lượt từng
    ngày; với mỗi ngày lấy dòng ĐẦU TIÊN khớp. Trả về rowid hoặc None.
    """
    if pay_amount is None or not md_list:
        return None
    for md in md_list:
        for r in bank_rows:
            t = r["time"]
            if t is None or r["income"] is None:
                continue
            if (t.month, t.day) != md:
                continue
            if r["rowid"] in done_set or r["rowid"] in tentative_set:
                continue
            if abs(r["income"] - pay_amount) <= CREDIT_LIMIT:
                return r["rowid"]
    return None


def _match_cash_by_dates(cash_rows, md_list, pay_amount, done_set, tentative_set):
    """
    Dự phòng cho tiền mặt: tìm dòng tiền mặt có Time đúng vào một trong các NGÀY
    (tháng, ngày) lấy từ Content, số tiền khớp (±1000), chưa is_done. Thử lần
    lượt từng ngày; lấy dòng ĐẦU TIÊN khớp. Trả về rowid hoặc None.
    """
    if pay_amount is None or not md_list:
        return None
    for md in md_list:
        for r in cash_rows:
            if r["month_day"] is None or r["income"] is None:
                continue
            if r["month_day"] != md:
                continue
            if r["rowid"] in done_set or r["rowid"] in tentative_set:
                continue
            if abs(r["income"] - pay_amount) <= CREDIT_LIMIT:
                return r["rowid"]
    return None


# ---------------------------------------------------------------------------
# CÁC HÀM PHỤ CHO STAGE B (tính tiền + nhãn)
# ---------------------------------------------------------------------------
def _parse_cell_ids(value):
    """
    Tách danh sách booking_id (đã chuẩn hóa) từ giá trị một ô bảng doanh thu.
    Ô chứa id thô dạng "E1" hoặc "E1, E2". Trả về list (giữ thứ tự, bỏ trùng).
    """
    if value is None:
        return []
    s = str(value).strip()
    if s == "" or s.lower() == "nan":
        return []
    result = []
    for piece in s.split(","):
        nid = matching.normalize_id(piece)
        if nid is not None and nid not in result:
            result.append(nid)
    return result


def _is_weekend(d):
    """True nếu ngày d là thứ Sáu (weekday=4) hoặc thứ Bảy (weekday=5)."""
    return d.weekday() in (4, 5)


def _nights_covered(checkin, checkout):
    """
    Trả về danh sách các ĐÊM (đối tượng date) mà một booking phủ, dựa trên
    Checkin/Checkout và đúng luật khớp đêm của hệ thống (matching.night_matches).
    Danh sách được sắp xếp tăng dần; phần tử cuối là đêm muộn nhất.

    Cách làm: duyệt các ngày ứng viên từ (ngày Checkin - 1) đến ngày Checkout,
    đếm những ngày D mà night_matches(checkin, checkout, D) = True. Khoảng này
    đủ bao quát cả 3 điều kiện (ĐK3 lùi 1 ngày trước Checkin, ĐK1/ĐK2 tới ngày
    Checkout). Tự động đúng cả khi lưu trú vắt sang tháng/năm khác.
    """
    nights = []
    start = checkin.date() - timedelta(days=1)
    end = checkout.date()
    d = start
    while d <= end:
        if matching.night_matches(checkin, checkout, d.year, d.month, d.day):
            nights.append(d)
        d += timedelta(days=1)
    return nights


def _find_booking_row(booking_by_id, bid, room_norm, day_col, year):
    """
    Tìm dòng trong ezcloud_booking_list ứng với (booking_id, phòng) cho ô ngày day_col.
    Nếu có nhiều dòng cùng [booking, phòng] thì chọn dòng mà ĐÊM đó "phủ" theo
    luật night_matches. Trả về dict dòng đó, hoặc None nếu không tìm được.
    """
    rows = [r for r in booking_by_id.get(bid, []) if r["room"] == room_norm]
    if not rows:
        return None
    if len(rows) == 1:
        return rows[0]

    day_str, month_str = day_col.split("/")
    day = int(day_str)
    month = int(month_str)
    for r in rows:
        if matching.night_matches(r["checkin"], r["checkout"], year, month, day):
            return r
    return rows[0]  # phòng hờ: không khớp đêm nào thì lấy dòng đầu


def _company_label(company_raw):
    """Nhãn company để hiển thị: rỗng -> 'walk-in', còn lại -> chữ thường (vd 'agoda')."""
    c = _norm_text(company_raw)
    if c == "":
        return "walk-in"
    return c


def _pay_type_label(payments):
    """
    Tính nhãn Pay_type cho cả booking từ danh sách khoản thanh toán, theo thứ
    tự ưu tiên đã thống nhất.
    """
    # 1) có Công nợ
    if any(p["pay_type"] == "công nợ" for p in payments):
        return "Công nợ"
    # 2) có Chuyển khoản NH + Content nhắc công ty
    if any(p["pay_type"] == "chuyển khoản nh" and _content_is_company(p["content"])
           for p in payments):
        return "Chuyển khoản công ty"
    # 3) có Thẻ tín dụng (hiếm khi tới đây vì thẻ đã skip booking)
    if any(p["pay_type"] == "thẻ tín dụng" for p in payments):
        return "Thẻ tín dụng"
    # 4) còn lại: xét chuyển khoản cá nhân và tiền mặt
    has_ck_personal = any(
        p["pay_type"] == "chuyển khoản nh" and not _content_is_company(p["content"])
        for p in payments
    )
    has_cash = any(p["pay_type"] == "tiền mặt" for p in payments)
    if has_ck_personal and has_cash:
        return "Chuyển khoản cá nhân + Tiền mặt"
    if has_ck_personal:
        return "Chuyển khoản cá nhân"
    if has_cash:
        return "Tiền mặt"
    return ""  # không rơi vào trường hợp nào (vd booking không có khoản nào)
