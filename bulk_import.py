# bulk_import.py
# ---------------------------------------------------------------------------
# Import NHIỀU file raw (chưa qua chỉnh sửa) trong một lần.
#
# Quy trình:
#   1. Với mỗi file: đọc -> nhận diện loại file (theo tiêu đề cột) -> làm sạch
#      (bỏ dòng tiêu đề thừa, dòng tổng, footer...) -> ra bảng sạch cho 1 bảng DB.
#   2. Gom các file cùng bảng lại (nối liên tiếp), áp mapping (database.prepare_import).
#   3. Nếu KHÔNG có lỗi format nào -> ghi tất cả vào DB trong MỘT transaction.
#      Có lỗi format -> không ghi gì (dữ liệu cũ giữ nguyên) và trả về danh sách lỗi.
#
# "Lỗi format" = file không đúng cấu trúc mong đợi (không nhận diện được, thiếu
# cột, sai kiểu ngày...). Lỗi về giá trị dữ liệu (ô trống, số lạ...) thì KHÔNG
# tính là lỗi format.
# ---------------------------------------------------------------------------

import io
import re
import unicodedata
import zipfile
from datetime import datetime

import pandas as pd

import database
import matching


# Cơ sở (hotel): nhãn hiển thị -> từ khóa nhận diện trong file
HOTELS = {
    "iRest Atelier": "atelier",
    "iRest Signature": "signature",
}

# Ngân hàng dùng cho tài khoản CÁ NHÂN của từng cơ sở
PERSONAL_BANK_BY_HOTEL = {
    "atelier": "vietinbank",
    "signature": "mb",
}
BANK_LABEL = {"vietinbank": "Vietinbank", "mb": "MB", "techcombank": "Techcombank"}

# Phòng chỉ có ở Signature (Atelier không có) - dùng nhận diện file tiền mặt
SIGNATURE_ONLY_ROOMS = {"604", "704", "801"}

MONTH_ABBR = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
}

# Bảng -> loại file (dùng để hiển thị)
KIND_LABEL = {
    "ezcloud_booking": "Ezcloud - báo cáo doanh thu",
    "ezcloud_payment": "Ezcloud - thanh toán hàng ngày",
    "ctrip": "Trip.com",
    "agoda": "Agoda",
    "expedia": "Expedia",
    "traveloka": "Traveloka",
    "booking": "Booking.com",
    "bank_vietinbank": "Sao kê Vietinbank",
    "bank_mb": "Sao kê MB",
    "bank_techcombank": "Sao kê Techcombank",
    "cash": "Thu chi tiền mặt",
}
KIND_TABLE = {
    "ezcloud_booking": "ezcloud_booking_list",
    "ezcloud_payment": "ezcloud_payment_list",
    "ctrip": "ctrip_booking_list",
    "agoda": "agoda_booking_list",
    "expedia": "expedia_booking_list",
    "traveloka": "traveloka_booking_list",
    "booking": "booking_booking_list",
    "cash": "cash_transactions",
}


class FormatError(Exception):
    """Lỗi format của một file (cấu trúc file không như mong đợi)."""


class BadDate(FormatError):
    """Một ô ngày trong file thu chi không đọc được (lỗi dữ liệu của dòng, không phải cả file)."""


class NeedSheetChoice(Exception):
    """Không tự chọn được sheet tiền mặt -> cần người dùng chọn trong danh sách sheet."""

    def __init__(self, sheet_names):
        super().__init__("Cần chọn sheet tiền mặt.")
        self.sheet_names = sheet_names


class SkipFile(Exception):
    """File hợp lệ nhưng không dùng cho lần import này (khác cơ sở, khác ngân hàng...)."""


# ---------------------------------------------------------------------------
# Đọc file
# ---------------------------------------------------------------------------
def _fix_styles(raw):
    """
    Một số file xlsx (vd sao kê Vietinbank) có màu trong stylesheet sai chuẩn
    nên openpyxl không mở được. Màu không ảnh hưởng dữ liệu -> bỏ các thẻ <color>.
    """
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(raw)) as zin, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                data = re.sub(rb"<color [^>]*/>", b"", data)
            zout.writestr(item, data)
    return out.getvalue()


def read_file(name, raw):
    """Đọc file thành dict {tên_sheet: DataFrame (không header, mọi ô là chuỗi)}."""
    lower = name.lower()
    try:
        if lower.endswith(".csv"):
            df = None
            for encoding in ("utf-8-sig", "utf-16", "cp1252"):
                try:
                    df = pd.read_csv(io.BytesIO(raw), header=None, dtype=str, encoding=encoding,
                                     keep_default_na=False, na_values=[""])
                    break
                except UnicodeError:
                    continue
            if df is None:
                raise FormatError("Không đọc được bảng mã của file csv.")
            sheets = {"csv": df}
        elif lower.endswith(".xlsx"):
            try:
                sheets = pd.read_excel(io.BytesIO(raw), sheet_name=None, header=None,
                                       dtype=str, engine="openpyxl",
                                       keep_default_na=False, na_values=[""])
            except ValueError:
                sheets = pd.read_excel(io.BytesIO(_fix_styles(raw)), sheet_name=None,
                                       header=None, dtype=str, engine="openpyxl",
                                       keep_default_na=False, na_values=[""])
        elif lower.endswith(".xls"):
            sheets = pd.read_excel(io.BytesIO(raw), sheet_name=None, header=None,
                                   dtype=str, engine="xlrd",
                                   keep_default_na=False, na_values=[""])
        else:
            raise FormatError("Định dạng file không được hỗ trợ (chỉ csv, xlsx, xls).")
    except FormatError:
        raise
    except Exception as e:  # file hỏng / không đọc được
        raise FormatError(f"Không đọc được file: {e}")
    return {k: _clean(v) for k, v in sheets.items()}


def _clean(df):
    """Bỏ khoảng trắng thừa, ô rỗng -> NaN, '123.0' -> '123'."""
    def fix(v):
        if not isinstance(v, str):
            return v
        s = v.strip()
        if s == "":
            return float("nan")
        if re.fullmatch(r"-?\d+\.0", s):
            return s[:-2]
        return s
    return df.astype(object).map(fix)


# ---------------------------------------------------------------------------
# Tiện ích dò header
# ---------------------------------------------------------------------------
def _norm(text):
    """Chuẩn hóa để so sánh: NFC, gộp khoảng trắng/xuống dòng, chữ thường."""
    s = unicodedata.normalize("NFC", str(text))
    return re.sub(r"\s+", " ", s).strip().casefold()


def _isblank(v):
    return v is None or (isinstance(v, float) and pd.isna(v))


def _row_cells(df, i):
    return [_norm(v) for v in df.iloc[i].tolist() if not _isblank(v)]


def _find_header(df, tokens, exact=False, max_rows=60):
    """
    Tìm dòng đầu tiên mà MỖI token đều xuất hiện trong ít nhất 1 ô.
    exact=True: ô phải bằng token; ngược lại: ô chứa token.
    Trả về chỉ số dòng, hoặc None.
    """
    tokens = [_norm(t) for t in tokens]
    for i in range(min(max_rows, len(df))):
        cells = _row_cells(df, i)
        if not cells:
            continue
        ok = True
        for t in tokens:
            if exact:
                if t not in cells:
                    ok = False
                    break
            elif not any(t in c for c in cells):
                ok = False
                break
        if ok:
            return i
    return None


def _find_col(header_cells, token, exact=False, startswith=False):
    """Chỉ số cột (vị trí) trong dòng header khớp token; None nếu không có."""
    token = _norm(token)
    for j, c in enumerate(header_cells):
        if _isblank(c):
            continue
        n = _norm(c)
        if (exact and n == token) or (startswith and n.startswith(token)) \
                or (not exact and not startswith and token in n):
            return j
    return None


def _table_from_header(df, hdr):
    """Dùng dòng hdr làm tên cột (đã chuẩn hóa khoảng trắng); phần dưới là dữ liệu."""
    names, seen = [], {}
    for j, v in enumerate(df.iloc[hdr].tolist()):
        n = f"_cột{j}" if _isblank(v) else database.normalize_column_name(v)
        if n in seen:
            seen[n] += 1
            n = f"{n}#{seen[n]}"
        else:
            seen[n] = 0
        names.append(n)
    body = df.iloc[hdr + 1:].copy()
    body.columns = names
    return body.reset_index(drop=True)


def _sheet_text(df, rows=8):
    """Gộp text của vài dòng đầu (để đọc tiêu đề báo cáo / tên khách sạn)."""
    parts = []
    for i in range(min(rows, len(df))):
        parts.extend(str(v) for v in df.iloc[i].tolist() if not _isblank(v))
    return _norm(" ".join(parts))


def _hotel_of_text(text):
    t = _norm(text)
    if "atelier" in t:
        return "atelier"
    if "signature" in t:
        return "signature"
    return None


def _filter_by_hotel(df, col, target):
    """
    Giữ các dòng thuộc cơ sở target dựa vào cột tên khách sạn.
    Nếu không dòng nào nhắc tới cơ sở nào -> giữ nguyên. Nếu có nhưng không
    dòng nào thuộc target -> SkipFile.
    """
    if col not in df.columns:
        return df
    tags = df[col].map(lambda v: None if _isblank(v) else _hotel_of_text(v))
    if tags.notna().sum() == 0:
        return df
    kept = df[tags == target]
    if len(kept) == 0:
        raise SkipFile("File thuộc cơ sở khác.")
    return kept.reset_index(drop=True)


# ---------------------------------------------------------------------------
# Nhận diện loại file
# ---------------------------------------------------------------------------
def classify(sheets):
    """Trả về mã loại file (khóa trong KIND_LABEL) hoặc None."""
    first = next(iter(sheets.values()))
    checks = [
        ("ezcloud_booking", ["mã đặt phòng", "mã ota", "tổng doanh thu"]),
        ("ezcloud_payment", ["#phòng", "hình thức tt", "#thanh toán"]),
        ("ctrip", ["mã xác nhận của khách sạn"]),
        ("agoda", ["bookingidexternal_reference_id"]),
        ("expedia", ["khoản tiền đặt phòng"]),
        ("traveloka", ["itinerary id"]),
        ("booking", ["mã số đặt phòng"]),
        ("bank_vietinbank", ["số tiền gd", "số dư"]),
        ("bank_mb", ["phát sinh nợ", "phát sinh có"]),
        ("bank_techcombank", ["requesting date", "credit"]),
    ]
    for kind, tokens in checks:
        if _find_header(first, tokens) is not None:
            return kind
    # Tiền mặt: workbook có sheet với header NỘI DUNG / THU / CHI / TIỀN MẶT
    # (ô tiêu đề cột NGÀY ở nhiều sheet bị bỏ trống nên không dùng để nhận diện)
    for df in sheets.values():
        if _find_header(df, ["nội dung", "thu", "chi", "tiền mặt"], exact=True) is not None:
            return "cash"
    return None


# ---------------------------------------------------------------------------
# Xử lý từng loại file. Mỗi hàm trả về DataFrame sạch của MỘT bảng DB.
#   - Bảng booking: giữ NGUYÊN tên cột của file (mapping cố định của database.py).
#   - Bảng sao kê / tiền mặt: trả về cột đã là tên cột DB (Time, Income, ...).
# ---------------------------------------------------------------------------
def _process_ezcloud(df, tokens, id_col, hotel, what):
    hdr = _find_header(df, tokens)
    if hdr is None:
        raise FormatError(f"Không tìm thấy dòng tiêu đề của {what}.")
    title_hotel = _hotel_of_text(_sheet_text(df, hdr))
    if title_hotel and title_hotel != hotel:
        raise SkipFile("File thuộc cơ sở khác.")
    t = _table_from_header(df, hdr)
    id_name = database.normalize_column_name(id_col)
    if id_name not in t.columns:
        raise FormatError(f"Thiếu cột '{id_col}'.")
    # Chỉ giữ dòng dữ liệu thật (mã là số): bỏ dòng "Phân hệ", "Khoản thu Tổng",
    # dòng tổng, footer "Người in..."
    t = t[t[id_name].map(lambda v: not _isblank(v) and re.fullmatch(r"\d+", str(v)) is not None)]
    return t.reset_index(drop=True)


def _process_ezcloud_booking(sheets, hotel):
    df = next(iter(sheets.values()))
    return _process_ezcloud(df, ["mã đặt phòng", "mã ota", "tổng doanh thu"],
                            "Mã đặt phòng", hotel, "báo cáo doanh thu ezcloud")


def _process_ezcloud_payment(sheets, hotel):
    df = next(iter(sheets.values()))
    return _process_ezcloud(df, ["#phòng", "hình thức tt", "#thanh toán"],
                            "#Phòng", hotel, "báo cáo thanh toán ezcloud")


def _process_flat(sheets, hotel, header_tokens, hotel_col=None):
    """File phẳng (Trip, Agoda, Expedia, Traveloka, Booking): header ở dòng đầu tìm được."""
    df = next(iter(sheets.values()))
    hdr = _find_header(df, header_tokens)
    if hdr is None:
        raise FormatError("Không tìm thấy dòng tiêu đề.")
    t = _table_from_header(df, hdr)
    t = t.dropna(how="all").reset_index(drop=True)
    if hotel_col:
        t = _filter_by_hotel(t, database.normalize_column_name(hotel_col), hotel)
    return t


def _parse_time_or_fail(times, what):
    bad = [t for t in times if matching.parse_datetime_full(t) is None]
    if bad:
        raise FormatError(
            f"{what}: {len(bad)} dòng có ngày giờ không đọc được (vd '{bad[0]}')."
        )


def _subject_has_company(subject):
    s = unicodedata.normalize("NFD", subject or "")
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn").lower().replace("đ", "d")
    return re.search(r"cong\s*ty", s) is not None


def _process_bank(sheets, kind, hotel, skip_company=False):
    """
    Trả về (bảng DB, DataFrame[Time, Income, Outcome]).
    Bảng được chọn theo CHỦ TÀI KHOẢN: có chữ "CÔNG TY" -> company_bank,
    ngược lại -> personal_bank (chỉ nhận ngân hàng tương ứng với cơ sở).
    """
    df = next(iter(sheets.values()))
    if kind == "bank_vietinbank":
        hdr = _find_header(df, ["số tiền gd", "số dư"])
        head = [df.iat[hdr, j] for j in range(df.shape[1])]
        c_stt = _find_col(head, "stt", exact=True)
        c_time = _find_col(head, "ngày", exact=True)
        c_amt = _find_col(head, "số tiền gd", exact=True)
        subject = ""
        for i in range(hdr):
            m = re.search(r"Số tài khoản:\s*\d+\s*-\s*(.+?)\s*-\s*Tài khoản", _row_text(df, i), re.I)
            if m:
                subject = m.group(1)
                break
    elif kind == "bank_mb":
        hdr = _find_header(df, ["phát sinh nợ", "phát sinh có"])
        head = [df.iat[hdr, j] for j in range(df.shape[1])]
        c_stt = _find_col(head, "stt", startswith=True)
        c_time = _find_col(head, "ngày giao dịch")
        c_in = _find_col(head, "phát sinh có")
        c_out = _find_col(head, "phát sinh nợ")
        subject = _find_label_value(df, hdr, "account name")
    else:  # techcombank
        hdr = _find_header(df, ["requesting date", "credit"])
        head = [df.iat[hdr, j] for j in range(df.shape[1])]
        c_stt = None
        c_time = _find_col(head, "requesting date")
        c_in = _find_col(head, "credit")
        c_out = _find_col(head, "debit")
        subject = _find_label_value(df, hdr, "account name")

    needed = {"ngày/time": c_time}
    if kind == "bank_vietinbank":
        needed.update({"Số tiền GD": c_amt, "STT": c_stt})
    elif kind == "bank_mb":
        needed.update({"Phát sinh có": c_in, "Phát sinh nợ": c_out, "STT": c_stt})
    else:
        needed.update({"Có/Credit": c_in, "Nợ/Debit": c_out})
    missing = [k for k, v in needed.items() if v is None]
    if missing:
        raise FormatError(f"Thiếu cột: {', '.join(missing)}.")
    if not subject:
        raise FormatError("Không đọc được tên chủ tài khoản trong file sao kê.")

    is_company = _subject_has_company(subject)
    table = "company_bank_transactions" if is_company else "personal_bank_transactions"
    if is_company and skip_company:
        raise SkipFile("Đã tick 'Tạm bỏ qua sao kê tài khoản công ty'.")
    if not is_company:
        expected = PERSONAL_BANK_BY_HOTEL[hotel]
        if kind.replace("bank_", "") != expected:
            raise SkipFile(
                f"Tài khoản cá nhân của {_hotel_name(hotel)} dùng sao kê "
                f"{BANK_LABEL[expected]}, file này là {BANK_LABEL[kind.replace('bank_', '')]}."
            )

    body = df.iloc[hdr + 1:]
    rows = []
    for _, r in body.iterrows():
        t = r.iloc[c_time]
        if _isblank(t) or matching.parse_datetime_full(t) is None:
            continue  # dòng tổng / footer / dòng trống
        if c_stt is not None and (_isblank(r.iloc[c_stt]) or not re.fullmatch(r"\d+", str(r.iloc[c_stt]))):
            continue
        if kind == "bank_vietinbank":
            parts = database._split_signed_amount(r.iloc[c_amt])
            inc = parts[1] if parts and parts[0] is False else None
            out = parts[1] if parts and parts[0] is True else None
        elif kind == "bank_mb":
            inc, out = _positive(r.iloc[c_in]), _positive(r.iloc[c_out])
        else:
            inc = None if _isblank(r.iloc[c_in]) else r.iloc[c_in]
            out = None if _isblank(r.iloc[c_out]) else str(r.iloc[c_out]).lstrip("-")
        rows.append((str(t), inc, out))
    if not rows:
        raise FormatError("Không có dòng giao dịch nào.")
    return table, pd.DataFrame(rows, columns=["Time", "Income", "Outcome"])


def _positive(v):
    """Số tiền > 0 -> giữ nguyên chuỗi; 0 / rỗng -> None."""
    n = matching.normalize_number(None if _isblank(v) else v)
    return str(v) if n is not None and n > 0 else None


def _row_text(df, i):
    return " ".join(str(v) for v in df.iloc[i].tolist() if not _isblank(v))


def _find_label_value(df, upto, label):
    """Tìm ô có chứa `label` ở các dòng phía trên header; lấy giá trị sau dấu ':' hoặc ô bên cạnh."""
    label = _norm(label)
    for i in range(upto):
        cells = df.iloc[i].tolist()
        for j, v in enumerate(cells):
            if _isblank(v) or label not in _norm(v):
                continue
            if ":" in str(v):
                tail = str(v).split(":", 1)[1].strip()
                if tail:
                    return tail
            for w in cells[j + 1:]:
                if not _isblank(w):
                    return str(w).strip()
    return ""


def _hotel_name(hotel):
    return next(k for k, v in HOTELS.items() if v == hotel)


# ---- Tiền mặt -------------------------------------------------------------
def _sheet_months(name):
    """Tên sheet -> (tập tháng, năm hoặc None); None nếu không phải sheet tháng."""
    m = re.fullmatch(r"\s*(\d{2})(\d{4})\s*", name)
    if m and 1 <= int(m.group(1)) <= 12:
        return {int(m.group(1))}, int(m.group(2))
    months = {MONTH_ABBR[w] for w in re.findall(r"[A-Za-z]{3}", name.upper()) if w in MONTH_ABBR}
    if months:
        return months, None
    return None


def _month_range(lo, hi):
    out, (y, m) = set(), (lo.year, lo.month)
    while (y, m) <= (hi.year, hi.month):
        out.add((y, m))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


CASH_HEADER = ["nội dung", "thu", "chi", "tiền mặt"]


def _cash_hotel(sheets, names):
    """
    Cơ sở của file thu chi, dựa vào số phòng trong các sheet `names` (phòng 604/704/801
    chỉ có ở Signature). Dùng tỉ lệ để một dòng gõ nhầm không làm sai kết quả.
    """
    total = sig = 0
    for name in names:
        df = sheets[name]
        if df.shape[1] > 1:
            for v in df.iloc[:, 1].dropna().tolist():
                for room in re.findall(r"[Pp]\s?(\d{3})", str(v)):
                    total += 1
                    sig += room in SIGNATURE_ONLY_ROOMS
    return "signature" if total and sig / total > 0.03 else "atelier"


def _cash_sheet_has_data(df):
    hdr = _find_header(df, CASH_HEADER, exact=True)
    head = [df.iat[hdr, j] for j in range(df.shape[1])]
    c_content = _find_col(head, "nội dung", exact=True)
    return df.iloc[hdr + 1:, c_content].notna().any()


def _process_cash(sheets, hotel, needed_months, warnings, fname, picked=None):
    """
    Chọn sheet theo thứ tự ưu tiên:
      1. `picked`: danh sách sheet do người dùng chọn.
      2. needed_months: tập (năm, tháng) suy ra từ khoảng ngày của file thanh toán ezcloud.
      3. Không biết -> ném NeedSheetChoice để giao diện hỏi người dùng.
    """
    header_sheets = [n for n, df in sheets.items()
                     if _find_header(df, CASH_HEADER, exact=True) is not None]
    if not header_sheets:
        raise FormatError("Không tìm thấy sheet thu chi nào (thiếu các cột NỘI DUNG, THU, CHI, TIỀN MẶT).")

    absent = []
    if picked:
        wrong = [n for n in picked if n not in header_sheets]
        if wrong:
            raise FormatError(f"Sheet đã chọn không phải sheet thu chi hoặc không tồn tại: {wrong}.")
        chosen_names = [n for n in header_sheets if n in picked]
    else:
        chosen_names = []
        if needed_months is not None:
            covered = set()
            for name in header_sheets:
                info = _sheet_months(name)
                if info is None:
                    continue
                months, year = info
                hit = {(y, m) for (y, m) in needed_months
                       if m in months and (year is None or y == year)}
                if hit:
                    chosen_names.append(name)
                    covered |= hit
            absent = sorted(needed_months - covered)

    # Cơ sở: kiểm tra trên sheet đã chọn (hoặc mọi sheet nếu chưa chọn được)
    if _cash_hotel(sheets, chosen_names or header_sheets) != hotel:
        raise SkipFile("File thu chi thuộc cơ sở khác.")

    if not chosen_names:
        with_data = [n for n in header_sheets if _cash_sheet_has_data(sheets[n])]
        raise NeedSheetChoice(with_data or header_sheets)

    if absent:
        warnings.append(
            f"{fname}: thiếu sheet tiền mặt cho tháng "
            + ", ".join(f"{m}/{y}" for y, m in absent) + "."
        )

    # Sắp sheet theo thời gian (sheet không đọc được tháng giữ nguyên thứ tự file)
    def order_key(name):
        info = _sheet_months(name)
        return (info[1] or 0, min(info[0])) if info else (0, 0)
    chosen_names.sort(key=order_key)

    # Xử lý từng sheet
    frames = []
    for name in chosen_names:
        info = _sheet_months(name)
        frames.append(_process_cash_sheet(sheets[name], name, info[0] if info else None,
                                          warnings, fname))
    result = pd.concat(frames, ignore_index=True)
    if len(result) == 0:
        raise FormatError("Không có dòng thu chi nào.")
    return result


_INVALID_DATE = object()  # đánh dấu nhóm dòng có ngày không đọc được


def _process_cash_sheet(df, sheet, months, warnings, fname):
    hdr = _find_header(df, CASH_HEADER, exact=True)
    head = [df.iat[hdr, j] for j in range(df.shape[1])]
    c_content = _find_col(head, "nội dung", exact=True)
    # Cột ngày: ô tiêu đề "NGÀY" (có thể bị bỏ trống) -> mặc định là cột liền trước NỘI DUNG
    c_date = _find_col(head, "ngày", exact=True)
    if c_date is None:
        c_date = c_content - 1
        if c_date < 0:
            raise FormatError(f"Sheet '{sheet}': không xác định được cột ngày.")
    c_in = _find_col(head, "tiền mặt", exact=True)
    c_out = _find_col(head, "chi", exact=True)
    c_thu = _find_col(head, "thu", exact=True)
    c_note = _find_col(head, "ghi chú", exact=True)

    current = None
    rows = []
    bad_dates = []
    for i in range(hdr + 1, len(df)):
        r = df.iloc[i]
        raw_date = r.iloc[c_date]
        if not _isblank(raw_date):
            try:
                current = _parse_cash_date(str(raw_date), months, sheet, i + 1)
            except BadDate:
                if current is None and not _looks_like_date(str(raw_date)):
                    continue  # nhãn chữ trước ngày đầu tiên (vd "TM T9" = số dư đầu kỳ)
                # Ngày gõ sai: các dòng thuộc ngày này để trống Time (không dùng để đối chiếu)
                # và báo cho người dùng, không hủy cả lần import.
                current = _INVALID_DATE
                bad_dates.append(f"dòng {i + 1} '{str(raw_date)[:20]}'")
        if current is None:
            continue  # dòng trước ngày đầu tiên (vd số dư đầu kỳ)
        content = r.iloc[c_content]
        vals = [r.iloc[c] for c in (c_thu, c_out, c_in) if c is not None]
        if _isblank(content) and all(_isblank(v) for v in vals):
            continue
        note = None if c_note is None or _isblank(r.iloc[c_note]) else r.iloc[c_note]
        rows.append((
            None if current is _INVALID_DATE else current,
            None if _isblank(content) else content,
            None if _isblank(r.iloc[c_in]) else r.iloc[c_in],
            None if _isblank(r.iloc[c_out]) else r.iloc[c_out],
            note,
        ))
    if bad_dates:
        warnings.append(
            f"{fname}, sheet '{sheet}': {len(bad_dates)} ô ngày không đọc được "
            f"({', '.join(bad_dates[:5])}{'...' if len(bad_dates) > 5 else ''}); "
            f"các dòng thuộc ngày đó được import với Time để trống."
        )
    return pd.DataFrame(rows, columns=["Time", "Content", "Income", "Outcome", "Note"])


def _looks_like_date(text):
    return re.match(r"\d{4}-\d{2}-\d{2}", text) is not None or \
        re.fullmatch(r"\d{1,2}/\d{1,2}(?:/\d{2,4})?", text) is not None


def _parse_cash_date(text, months, sheet, line):
    """
    Trả về chuỗi "ngày/tháng".
      - "17/3", "16/3/2026"  -> "17/3", "16/3"
      - "2026-01-10 00:00:00": Excel đọc nhầm "1/10" (ngày 1 tháng 10) thành ngày
        10/01. Đối chiếu với tháng ghi trên tên sheet để biết đảo hay không.
    """
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})(?:/\d{2,4})?", text)
    if m:
        return f"{int(m.group(1))}/{int(m.group(2))}"
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", text)
    if m:
        if not months:
            raise FormatError(
                f"Sheet '{sheet}' dòng {line}: ô ngày bị Excel đọc thành ngày/tháng đầy đủ "
                f"('{text[:10]}') nhưng không biết tháng của sheet để sửa. Hãy đặt tên sheet "
                f"theo dạng 102026 hoặc Oct."
            )
        a, b = int(m.group(2)), int(m.group(3))  # a = "tháng" Excel hiểu, b = "ngày" Excel hiểu
        if b in months:       # Excel đảo: thực tế là ngày a, tháng b
            return f"{a}/{b}"
        if a in months:       # Excel đọc đúng: ngày b, tháng a
            return f"{b}/{a}"
        raise BadDate(
            f"Sheet '{sheet}' dòng {line}: ngày '{text[:10]}' không thuộc tháng của sheet."
        )
    raise BadDate(f"Sheet '{sheet}' dòng {line}: không đọc được ngày '{text}'.")


# ---------------------------------------------------------------------------
# Điều phối
# ---------------------------------------------------------------------------
class BulkResult:
    def __init__(self):
        self.files = []        # [{file, kind, table, rows, status, detail}]
        self.errors = []       # [(tên file, nội dung lỗi)]
        self.warnings = []     # [nội dung]
        self.tables = {}       # {bảng: số dòng đã import}
        self.committed = False
        self.sheet_requests = []  # [{file, sheets}] file tiền mặt cần người dùng chọn sheet


def run_bulk_import(files, hotel_label, skip_company=False, cash_sheets=None):
    """
    files: list (tên_file, bytes). hotel_label: khóa của HOTELS.
    skip_company: True -> bỏ qua hoàn toàn bảng company_bank_transactions; False ->
        bắt buộc phải có file sao kê tài khoản công ty, nếu không báo lỗi.
    cash_sheets: {tên file thu chi: [tên sheet]} do người dùng chọn (khi tool không tự chọn được).
    Trả về BulkResult. Chỉ ghi DB khi không có lỗi format và không còn việc chờ người dùng chọn.
    """
    cash_sheets = cash_sheets or {}
    hotel = HOTELS[hotel_label]
    res = BulkResult()
    parts = {}  # bảng -> list DataFrame
    parsed = []  # (name, kind, sheets)

    # --- Bước 1: đọc + nhận diện ---
    for name, raw in files:
        try:
            sheets = read_file(name, raw)
            kind = classify(sheets)
            if kind is None:
                raise FormatError("Không nhận diện được đây là file gì (không khớp loại file nào).")
            parsed.append((name, kind, sheets))
        except FormatError as e:
            res.errors.append((name, str(e)))
            res.files.append({"file": name, "kind": "?", "table": "-", "rows": 0,
                              "status": "Lỗi format", "detail": str(e)})

    def add(name, kind, table, df, extra=""):
        parts.setdefault(table, []).append((name, kind, df))
        res.files.append({"file": name, "kind": KIND_LABEL[kind], "table": table,
                          "rows": len(df), "status": "OK", "detail": extra})

    def fail(name, kind, msg):
        res.errors.append((name, msg))
        res.files.append({"file": name, "kind": KIND_LABEL.get(kind, kind), "table": "-",
                          "rows": 0, "status": "Lỗi format", "detail": msg})

    def skip(name, kind, msg):
        res.warnings.append(f"{name}: bỏ qua - {msg}")
        res.files.append({"file": name, "kind": KIND_LABEL.get(kind, kind), "table": "-",
                          "rows": 0, "status": "Bỏ qua", "detail": msg})

    # --- Bước 2: xử lý file (tiền mặt xử lý sau cùng vì cần khoảng ngày của file thanh toán) ---
    deferred_cash = []
    for name, kind, sheets in parsed:
        try:
            if kind == "ezcloud_booking":
                add(name, kind, KIND_TABLE[kind], _process_ezcloud_booking(sheets, hotel))
            elif kind == "ezcloud_payment":
                add(name, kind, KIND_TABLE[kind], _process_ezcloud_payment(sheets, hotel))
            elif kind == "ctrip":
                add(name, kind, KIND_TABLE[kind],
                    _process_flat(sheets, hotel, ["mã xác nhận của khách sạn"], "Tên khách sạn"))
            elif kind == "agoda":
                add(name, kind, KIND_TABLE[kind],
                    _process_flat(sheets, hotel, ["bookingidexternal_reference_id"]))
            elif kind == "expedia":
                add(name, kind, KIND_TABLE[kind],
                    _process_flat(sheets, hotel, ["khoản tiền đặt phòng"]))
            elif kind == "traveloka":
                add(name, kind, KIND_TABLE[kind],
                    _process_flat(sheets, hotel, ["itinerary id"], "Hotel Name"))
            elif kind == "booking":
                raise SkipFile("Booking.com tạm thời không được import ở chế độ nhiều file.")
            elif kind.startswith("bank_"):
                table, df = _process_bank(sheets, kind, hotel, skip_company)
                add(name, kind, table, df)
            elif kind == "cash":
                deferred_cash.append((name, kind, sheets))
        except SkipFile as e:
            skip(name, kind, str(e))
        except FormatError as e:
            fail(name, kind, str(e))
        except Exception as e:  # lỗi không lường trước khi xử lý -> coi là lỗi format của file đó
            fail(name, kind, f"Lỗi khi xử lý file: {e}")

    needed = _needed_months(parts.get("ezcloud_payment_list"))
    for name, kind, sheets in deferred_cash:
        try:
            warn_before = len(res.warnings)
            df = _process_cash(sheets, hotel, needed, res.warnings, name,
                               cash_sheets.get(name))
            add(name, kind, KIND_TABLE[kind], df,
                "; ".join(res.warnings[warn_before:]))
        except NeedSheetChoice as e:
            res.sheet_requests.append({"file": name, "sheets": e.sheet_names})
            res.files.append({"file": name, "kind": KIND_LABEL[kind], "table": KIND_TABLE[kind],
                              "rows": 0, "status": "Chờ chọn sheet",
                              "detail": "Chưa xác định được sheet cần import, hãy chọn sheet."})
        except SkipFile as e:
            skip(name, kind, str(e))
        except FormatError as e:
            fail(name, kind, str(e))
        except Exception as e:
            fail(name, kind, f"Lỗi khi xử lý file: {e}")

    # --- Bước 3: gom theo bảng, sắp xếp sao kê theo thời gian, áp mapping ---
    prepared = {}
    for table, items in parts.items():
        frames = [df for _, _, df in items]
        merged = pd.concat(frames, ignore_index=True) if len(frames) > 1 else frames[0]
        custom = None
        if table in database.MANUAL_MAPPING_TABLES:
            if table == "cash_transactions":
                custom = {c: c for c in ("Time", "Content", "Income", "Outcome", "Note")}
            else:
                custom = {c: c for c in ("Time", "Income", "Outcome")}
            if table.endswith("_bank_transactions"):
                merged = _sort_by_time_desc(merged)
        try:
            df_ins, _, missing = database.prepare_import(table, merged, custom)
        except Exception as e:
            res.errors.append((", ".join(n for n, _, _ in items), f"Lỗi mapping bảng {table}: {e}"))
            continue
        if missing:
            res.errors.append((
                ", ".join(n for n, _, _ in items),
                f"Bảng {table}: thiếu cột {missing} (format file đã thay đổi?)."))
            continue
        if df_ins is None or len(df_ins) == 0:
            res.errors.append((", ".join(n for n, _, _ in items), f"Bảng {table}: không có dòng dữ liệu nào."))
            continue
        prepared[table] = df_ins

    # Sao kê tài khoản công ty: không tick "bỏ qua" thì bắt buộc phải có file
    if not skip_company and "company_bank_transactions" not in parts:
        bank_failed = any(f["status"] == "Lỗi format" and "Sao kê" in f["kind"] for f in res.files)
        if not bank_failed:
            res.errors.append((
                "-",
                "Không tìm thấy file sao kê tài khoản công ty (chủ tài khoản có chữ CÔNG TY). "
                "Nếu chưa có thì tick 'Tạm bỏ qua sao kê tài khoản công ty'."))

    # Booking.com không import ở chế độ này nên không nhắc trong cảnh báo
    untouched = [t for t in database.TABLES if t not in parts and t != KIND_TABLE["booking"]]
    if untouched:
        res.warnings.append(
            "Các bảng không có file trong lần import này (giữ nguyên dữ liệu cũ): "
            + ", ".join(untouched)
        )

    if not res.errors and not prepared and not res.sheet_requests:
        res.errors.append(("-", "Không có file nào được import."))

    # --- Bước 4: ghi DB trong 1 transaction nếu không có lỗi và không còn chờ chọn sheet ---
    if not res.errors and not res.sheet_requests:
        try:
            database.replace_tables_atomically(prepared)
            res.committed = True
            res.tables = {t: len(df) for t, df in prepared.items()}
        except Exception as e:
            res.errors.append(("-", f"Ghi database thất bại, đã rollback toàn bộ: {e}"))
    return res


def _needed_months(payment_parts):
    """Tập (năm, tháng) mà file thanh toán ezcloud trải qua; None nếu không có file."""
    if not payment_parts:
        return None
    times = []
    for _, _, df in payment_parts:
        col = database.normalize_column_name("Thời gian")
        if col in df.columns:
            times += [matching.parse_datetime_full(v) for v in df[col]]
    times = [t for t in times if t is not None]
    if not times:
        return None
    return _month_range(min(times), max(times))


def _sort_by_time_desc(df):
    keys = df["Time"].map(lambda v: matching.parse_datetime_full(v) or datetime.min)
    return df.assign(_k=keys).sort_values("_k", ascending=False, kind="stable").drop(columns="_k").reset_index(drop=True)
