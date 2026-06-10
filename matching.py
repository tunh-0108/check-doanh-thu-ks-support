# matching.py
# ---------------------------------------------------------------------------
# File này chứa "LUẬT" để quyết định một booking có thuộc về một đêm cụ thể
# (tại một phòng) hay không. Đây CHÍNH LÀ CHỖ để chỉnh sửa điều kiện sau này.
#
# Các hàm ở đây là logic thuần (không đụng tới database), nên rất dễ đọc, dễ
# sửa và dễ kiểm tra.
# ---------------------------------------------------------------------------

from datetime import datetime, date, timedelta


# ---------------------------------------------------------------------------
# HAI MỐC GIỜ QUAN TRỌNG — muốn đổi điều kiện thì sửa 2 số này là chính:
#   - LATE_CHECKOUT_HOUR = 16: nếu khách trả phòng ĐÚNG NGÀY đang xét nhưng
#     vào lúc >= 16:00 thì phải tính thêm 1 đêm cho ngày đó (ĐK2).
#   - EARLY_CHECKIN_HOUR = 9: nếu khách nhận phòng vào NGÀY HÔM SAU nhưng
#     trước 09:00 thì phải tính thêm 1 đêm cho ngày đang xét (ĐK3).
# ---------------------------------------------------------------------------
LATE_CHECKOUT_HOUR = 16
EARLY_CHECKIN_HOUR = 9


# Định dạng ngày giờ trong dữ liệu, ví dụ: "30/04/26 19:18"
#   %d = ngày, %m = tháng, %y = năm 2 chữ số, %H = giờ, %M = phút
DATETIME_FORMAT = "%d/%m/%y %H:%M"


def parse_datetime(text):
    """
    Chuyển một chuỗi ngày giờ trong dữ liệu thành đối tượng datetime.
    Nếu ô trống / không đọc được thì trả về None (để chỗ gọi tự bỏ qua).

    Ví dụ: "30/04/26 19:18" -> datetime(2026, 4, 30, 19, 18)
    """
    if text is None:
        return None
    s = str(text).strip()
    if s == "" or s.lower() in ("none", "nan"):
        return None
    try:
        return datetime.strptime(s, DATETIME_FORMAT)
    except ValueError:
        # Không đúng định dạng -> trả về None để chỗ gọi biết là "không đọc được"
        return None


# Các định dạng ngày giờ của Pay_time và cột Time trong sao kê (năm 4 chữ số):
#   "01/04/2026 07:06:19" (có giây)
#   "26/05/2026 21:53"     (không giây)
#   "27/04/2026"           (chỉ có ngày)
#   "2026-05-31 02:30:24"  (dạng ISO — bảng company_bank dùng kiểu này)
#   "2026-05-31 02:30"     (ISO không giây)
#   "2026-05-31"           (ISO chỉ ngày)
FULL_DATETIME_FORMATS = [
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%d",
]


def parse_datetime_full(text):
    """
    Chuyển chuỗi ngày giờ của Pay_time / cột Time (sao kê, tiền mặt) thành
    datetime. Thử lần lượt nhiều định dạng (có giây, không giây, chỉ ngày).
    Trả về None nếu trống hoặc không đọc được.
    """
    if text is None:
        return None
    s = str(text).strip()
    if s == "" or s.lower() in ("none", "nan"):
        return None
    for fmt in FULL_DATETIME_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def normalize_room(value):
    """
    Chuẩn hóa giá trị số phòng để so sánh cho khớp.
    Dữ liệu đôi khi lưu là số thực (ví dụ 201.0) trong khi bảng doanh thu lưu
    là "201". Hàm này đưa cả hai về cùng một dạng chuỗi.

    Ví dụ: "201.0" -> "201" ; 201.0 -> "201" ; "201" -> "201" ; "P201" -> "P201"
    """
    s = str(value).strip()
    try:
        number = float(s)
        # Nếu là số nguyên (vd 201.0) thì bỏ phần thập phân
        if number == int(number):
            return str(int(number))
        return str(number)
    except ValueError:
        # Không phải số -> giữ nguyên chuỗi
        return s


def normalize_id(value):
    """
    Chuẩn hóa một MÃ (id) về dạng chuỗi số CHÍNH XÁC, để so khớp giữa OTA_id và
    id trong bảng chứng từ OTA. Dữ liệu có thể lưu dạng số thực như 1714647764.0.

    Ví dụ: "1714647764.0" -> "1714647764" ; 1714647764.0 -> "1714647764"
    Nếu không phải số (mã có chữ) thì giữ nguyên chuỗi (đã bỏ khoảng trắng).
    Trả về None nếu giá trị rỗng.
    """
    if value is None:
        return None
    s = str(value).strip()
    if s == "" or s.lower() in ("none", "nan"):
        return None
    # Dùng lại normalize_room: bỏ ".0", xử lý dạng khoa học "1.12e+15", v.v.
    return normalize_room(s)


def id_15sig(value):
    """
    Trả về mã đã LÀM TRÒN về 15 chữ số có nghĩa (dùng cho khớp dự phòng).

    Lý do: Excel chỉ giữ 15 chữ số có nghĩa, nên mã 16 chữ số có thể bị làm
    tròn trong dữ liệu cũ (ví dụ ...089 -> ...090). Khi khớp CHÍNH XÁC bị trượt,
    ta thử khớp lại theo 15 chữ số: làm tròn cả hai phía về cùng một dạng. Mã
    ngắn (<=15 chữ số) giữ nguyên. Trả về None nếu không phải số nguyên.
    """
    base = normalize_id(value)
    if base is None:
        return None
    digits = base.lstrip("-")
    if not digits.isdigit():
        return None
    if len(digits) <= 15:
        return base
    try:
        return str(int(float(f"{int(base):.15g}")))
    except (ValueError, OverflowError):
        return None


def normalize_number(value):
    """
    Chuyển một giá trị tiền/số về dạng số thực (float) để tính toán.
    Xử lý được:
      - dấu phẩy ngăn cách nghìn: "1,000,000" -> 1000000.0
      - số âm: "-500" -> -500.0
      - khoảng trắng thừa.
    Trả về None nếu rỗng hoặc không đọc được (để chỗ gọi tự xử lý).

    LƯU Ý về dấu phẩy: ở đây coi dấu phẩy là NGĂN CÁCH NGHÌN (kiểu "1,000,000"),
    nên sẽ bị bỏ đi. Nếu dữ liệu của bạn dùng dấu phẩy làm dấu thập phân thì
    cần xem lại hàm này.
    """
    if value is None:
        return None
    s = str(value).strip()
    if s == "" or s.lower() in ("none", "nan"):
        return None
    # Bỏ dấu phẩy ngăn cách nghìn
    s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def night_matches(checkin, checkout, year, month, day):
    """
    Trả về True nếu một booking (với thời điểm checkin/checkout đã cho) được
    tính cho ĐÊM của ngày (day/month/year), theo 3 điều kiện đã thống nhất.

    Tham số:
      - checkin, checkout: đối tượng datetime (hoặc None nếu không có).
      - year, month, day: ngày ĐÊM cần xét (số nguyên).

    Gọi D = ngày đang xét (year/month/day), D+1 = ngày liền sau D (tính cả
    sang tháng/năm). Ba điều kiện (chỉ cần ĐÚNG 1 là khớp):

      ĐK1 (chỉ so phần NGÀY, bỏ giờ):
          ngày(Checkin) <= D  VÀ  D < ngày(Checkout)
      ĐK2 (khách trả phòng muộn đúng ngày D):
          ngày(Checkin) <= D  VÀ  ngày(Checkout) = D  VÀ  giờ(Checkout) >= 16:00
      ĐK3 (khách nhận phòng sớm ngày D+1):
          ngày(Checkin) = D+1  VÀ  ngày(Checkout) >= D+1  VÀ  giờ(Checkin) < 09:00

    Ghi chú: chấp nhận một booking có thể khớp 2 đêm ở ca biên (khách đến rất
    sớm và đi rất muộn), đúng theo thỏa thuận.
    """
    # D và D+1 ở dạng "chỉ có ngày" (bỏ giờ/phút) để so sánh phần ngày.
    d = date(year, month, day)
    d_next = d + timedelta(days=1)

    # Lấy phần NGÀY (bỏ giờ) của checkin/checkout, nếu có.
    checkin_date = checkin.date() if checkin is not None else None
    checkout_date = checkout.date() if checkout is not None else None

    # --- ĐK1: chỉ so phần ngày ---
    if checkin_date is not None and checkout_date is not None:
        if checkin_date <= d and d < checkout_date:
            return True

    # --- ĐK2: trả phòng muộn đúng ngày D (cần xét cả giờ) ---
    if checkin_date is not None and checkout is not None:
        if checkin_date <= d and checkout_date == d and checkout.hour >= LATE_CHECKOUT_HOUR:
            return True

    # --- ĐK3: nhận phòng sớm ngày D+1 (cần xét cả giờ) ---
    if checkin is not None and checkout_date is not None:
        if checkin_date == d_next and checkout_date >= d_next and checkin.hour < EARLY_CHECKIN_HOUR:
            return True

    return False
