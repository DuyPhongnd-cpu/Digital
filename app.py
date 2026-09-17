import streamlit as st
import pandas as pd
import io
from datetime import datetime, time
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------------------------------------------------------
# 1. CẤU HÌNH TRANG WEB & FONT TIMES NEW ROMAN
# ---------------------------------------------------------
st.set_page_config(
    page_title="HỆ THỐNG QUẢN LÝ VẬN HÀNH NOC VÀ BẢO HÀNH SẢN PHẨM VTC",
    page_icon="🖥️",
    layout="wide"
)

st.markdown("""
    <style>
    html, body, [class*="css"], .stText, .stMarkdown, h1, h2, h3, h4, h5, h6, label, input, button {
        font-family: 'Times New Roman', Times, serif !important;
    }
    .stButton>button {
        font-family: 'Times New Roman', Times, serif !important;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 2. HÀM ĐỌC LỊCH TRỰC VTC
# ---------------------------------------------------------
def parse_vtc_matrix_schedule(uploaded_file):
    df_raw = pd.read_excel(uploaded_file, sheet_name=0, skiprows=2)
    df_raw.columns = [str(col).strip() for col in df_raw.iloc[0].values]
    
    df_data = df_raw.iloc[1:].dropna(subset=['STT']).copy()
    
    shift_map = {
        "ca1": "Ca 1: 07h30 - 14h30",
        "ca2": "Ca 2: 14h30 - 22h00",
        "ca3": "Ca 3: 22h00 - 07h30 sáng"
    }
    
    date_cols = [str(i) for i in range(1, 31)]
    records = []
    
    for _, row in df_data.iterrows():
        name = str(row.get('Họ tên/Bộ phận', '')).strip()
        if not name or name == 'nan':
            continue
            
        for day in date_cols:
            if day in row:
                shift_val = str(row[day]).strip().lower() if pd.notna(row[day]) else ""
                if shift_val in shift_map:
                    day_str = f"{int(day):02d}/09/2026"
                    records.append({
                        "Ngày": day_str,
                        "Ca trực": shift_map[shift_val],
                        "Người trực": name,
                        "Trạm": "Trạm Phát sóng VTC Digital"
                    })
                    
    if not records:
        return None

    df_result = pd.DataFrame(records)
    df_grouped = df_result.groupby(["Ngày", "Ca trực", "Trạm"])["Người trực"].apply(lambda x: ", ".join(x)).reset_index()
    df_grouped['DayNum'] = df_grouped['Ngày'].apply(lambda x: int(x.split('/')[0]))
    df_grouped = df_grouped.sort_values(by=['DayNum', 'Ca trực']).drop(columns=['DayNum'])
    return df_grouped


# ---------------------------------------------------------
# 3. HỆ THỐNG LƯU TRỮ DỮ LIỆU DÙNG CHUNG (PERSISTENCE STORAGE)
# ---------------------------------------------------------
DATA_FILE = "noc_system_storage.json"

def get_default_data():
    server_list = []
    for i in range(1, 33):
        status = "Bình thường" if i % 7 != 0 else "Cảnh báo Cao"
        cpu = f"{20 + (i * 2) % 60}%"
        ram = f"{40 + (i * 3) % 50}%"
        server_list.append({
            "STT": i,
            "Tên Server": f"Server-NOC-{i:02d}",
            "Địa chỉ IP": f"192.168.10.{100+i}",
            "Chức năng": f"Dịch vụ luồng Kênh {i}" if i <= 20 else f"Core Database {i-20}",
            "CPU Util": cpu,
            "RAM Util": ram,
            "Trạng thái Cảnh báo": status
        })

    return {
        "master_schedule": [
            {"Ngày": "01/09/2026", "Ca trực": "Ca 1: 07h30 - 14h30", "Người trực": "Bùi Trọng Vinh, Trần Đức Chiến", "Trạm": "Trạm Phát sóng VTC Digital"},
            {"Ngày": "01/09/2026", "Ca trực": "Ca 2: 14h30 - 22h00", "Người trực": "Nguyễn Văn Đông, Trương Nhật Minh", "Trạm": "Trạm Phát sóng VTC Digital"},
            {"Ngày": "01/09/2026", "Ca trực": "Ca 3: 22h00 - 07h30 sáng", "Người trực": "Nguyễn Văn Kiên, Võ Bảo Quang", "Trạm": "Trạm Phát sóng VTC Digital"}
        ],
        "shift_change_requests": [
            {
                "Mã GD": "DC001",
                "Ngày": "01/09/2026",
                "Ca trực": "Ca 1: 07h30 - 14h30",
                "Người xin đổi": "Bùi Trọng Vinh",
                "Người trực thay": "Vũ Quốc Minh",
                "Lý do": "Công tác đột xuất",
                "Trạng thái": "Chờ duyệt"
            }
        ],
        "tv_incidents": [
            {
                "STT": 1,
                "Ngày": "01/09/2026",
                "Tên kênh/nhóm kênh (Đường truyền)": "VTV3 / VTV CAB",
                "Hiện tượng": "Vỡ hình nhẹ",
                "Bắt đầu": "09:00",
                "Kết thúc": "11:00",
                "Thời lượng": "02h 00m",
                "Nguyên nhân": "Suy hao đường truyền quang",
                "Biện pháp khắc phục (Bên khắc phục)": "Hàn lại sợi quang / VTV Cab"
            }
        ],
        "hvac_schedule": [
            {"Ngày / Tuần": "01/09/2026", "Thời gian (Time)": "08:00 - 20:00", "Chu kỳ": "Ngày 1", "Máy chạy (Chính)": "Máy 1 – Máy 2 – Máy 3", "Máy nghỉ (Dự phòng)": "Máy 4 – Máy 5 – Máy 6 – Máy 7 – Máy 8", "Ghi chú / Trạng thái": "Hoạt động ổn định"},
            {"Ngày / Tuần": "02/09/2026", "Thời gian (Time)": "08:00 - 20:00", "Chu kỳ": "Ngày 2", "Máy chạy (Chính)": "Máy 4 – Máy 5 – Máy 6", "Máy nghỉ (Dự phòng)": "Máy 1 – Máy 2 – Máy 3 – Máy 7 – Máy 8", "Ghi chú / Trạng thái": "Dự phòng bình thường"},
            {"Ngày / Tuần": "03/09/2026", "Thời gian (Time)": "08:00 - 20:00", "Chu kỳ": "Ngày 3", "Máy chạy (Chính)": "Máy 7 – Máy 8 – Máy 1", "Máy nghỉ (Dự phòng)": "Máy 2 – Máy 3 – Máy 4 – Máy 5 – Máy 6", "Ghi chú / Trạng thái": "Chuyển chu kỳ tịnh tiến"}
        ],
        "tech_params": [
            {"STT": 1, "Thông số HPA": "HPA Power Level", "Máy phát K1H-VNS1": "18 kW", "Ngưỡng tiêu chuẩn": "17 - 19 kW", "Đánh giá": "Đạt"},
            {"STT": 2, "Thông số HPA": "Chỉ số C/N (dBC)", "Máy phát K1H-VNS1": "14.63 dB", "Ngưỡng tiêu chuẩn": "> 14 dB", "Đánh giá": "Đạt"}
        ],
        "ups_params": [
            {"STT": 1, "Tên hệ thống UPS": "UPS Phụ tải NOC - 01", "Điện áp vào (V)": "380V", "Điện áp ra (V)": "220V", "Mức tải (% Load)": "45%", "Dung lượng Pin (%)": "100%", "Trạng thái": "Bình thường"},
            {"STT": 2, "Tên hệ thống UPS": "UPS Máy phát K1H - 02", "Điện áp vào (V)": "382V", "Điện áp ra (V)": "220V", "Mức tải (% Load)": "60%", "Dung lượng Pin (%)": "98%", "Trạng thái": "Bình thường"}
        ],
        "server_warnings": server_list,
        "idc_network": [
            {"STT": 1, "Dải mạng / Hệ thống": "Core Network IDC 192.168.10.xx", "Băng thông": "300 Mbps", "Trạng thái Sự cố": "Bình thường", "Thời điểm phát hiện": "—"}
        ],
        "menu1_data": {
            "network_issue": "Bình thường - Không ghi nhận nghẽn mạng",
            "issue_status": "resolved",
            "email_user": "giamsatvtc.65lt@vtc.vn",
            "email_pass": "",
            "item3_note": "Hệ thống điện & Làm mát (CRAC): Hoạt động tốt (24°C)",
            "warning_servers": [
                {"id": "Server 03", "issue": "Cảnh báo CPU cao (>92%)"},
                {"id": "Server 08", "issue": "Dung lượng ổ cứng đầy (Còn 2%)"},
                {"id": "Server 12", "issue": "Mất kết nối mạng tạm thời (Ping timeout)"},
                {"id": "Server 19", "issue": "Nhiệt độ CPU vượt ngưỡng (82°C)"},
                {"id": "Server 27", "issue": "Lỗi RAM ECC - Cần kiểm tra"}
            ],
            "office_network": {
                "modem": "Modem Viettel Enterprise - Băng thông 500Mbps (OK)",
                "sw_core": "2x Switch Core Cisco C9300 - Hoạt động (Stacking OK)",
                "sw_poe": "2x Switch PoE Aruba 2930F - Hoạt động (Cấp nguồn Camera/AP)",
                "sw_branch": "4x Switch Nhánh TP-Link JetStream - Hoạt động"
            }
        },
        "warranty_meta": {
            "title": "Báo cáo bảo hành sản phẩm VTC Digital",
            "week": "Tuần 1/ Tháng 9.2026",
            "executor": "Kỹ sư Nguyễn Vĩnh Toàn",
            "data_entry": "Nguyễn Trọng Hùng",
            "period": "từ 07.09.2026 đến 11.09.2026"
        },
        "warranty_repair_data": [
            {"STT": 1, "Ngày nhập": "04/09/2026", "Loại đầu thu": "HDV2", "Mã dịch vụ": "4256419426", "Sửa chữa / Thay thế": "Nguồn", "Tình trạng Bảo hành": "Còn", "Ngày trả (hoàn thành)": 1},
            {"STT": 2, "Ngày nhập": "04/09/2026", "Loại đầu thu": "HDV3", "Mã dịch vụ": "4256419427", "Sửa chữa / Thay thế": "Tuner", "Tình trạng Bảo hành": "Hết", "Ngày trả (hoàn thành)": 1}
        ],
        "warranty_exchange_data": [
            {"STT": 1, "Ngày nhập": "04/09/2026", "Tên khách hàng / Địa chỉ": "Đại lý Hà Nội", "Loại đầu thu": "HDV2", "Mã dịch vụ cũ": "3912784969", "Mã dịch vụ mới": "3922395612", "Người thực hiện": "Nguyễn Vĩnh Toàn", "Ngày trả (hoàn thành)": 1},
            {"STT": 2, "Ngày nhập": "04/09/2026", "Tên khách hàng / Địa chỉ": "Khách lẻ Hải Phòng", "Loại đầu thu": "HDV3", "Mã dịch vụ cũ": "3911654457", "Mã dịch vụ mới": "3922564598", "Người thực hiện": "Nguyễn Vĩnh Toàn", "Ngày trả (hoàn thành)": 1}
        ],
        "audit_logs": [
            {"Thời gian": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Người dùng": "Hệ thống", "Thao tác": "Khởi chạy ứng dụng", "Ghi chú": "Mở phiên làm việc"}
        ]
    }

def load_shared_storage():
    import json, os
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    def_data = get_default_data()
    save_shared_storage(def_data)
    return def_data

def save_shared_storage(data=None):
    import json
    if data is None:
        data = {
            "master_schedule": st.session_state.master_schedule.to_dict(orient="records") if isinstance(st.session_state.master_schedule, pd.DataFrame) else st.session_state.master_schedule,
            "shift_change_requests": st.session_state.shift_change_requests.to_dict(orient="records") if isinstance(st.session_state.shift_change_requests, pd.DataFrame) else st.session_state.shift_change_requests,
            "tv_incidents": st.session_state.tv_incidents.to_dict(orient="records") if isinstance(st.session_state.tv_incidents, pd.DataFrame) else st.session_state.tv_incidents,
            "hvac_schedule": st.session_state.hvac_schedule.to_dict(orient="records") if isinstance(st.session_state.hvac_schedule, pd.DataFrame) else st.session_state.hvac_schedule,
            "tech_params": st.session_state.tech_params.to_dict(orient="records") if isinstance(st.session_state.tech_params, pd.DataFrame) else st.session_state.tech_params,
            "ups_params": st.session_state.ups_params.to_dict(orient="records") if isinstance(st.session_state.ups_params, pd.DataFrame) else st.session_state.ups_params,
            "server_warnings": st.session_state.server_warnings.to_dict(orient="records") if isinstance(st.session_state.server_warnings, pd.DataFrame) else st.session_state.server_warnings,
            "idc_network": st.session_state.idc_network.to_dict(orient="records") if isinstance(st.session_state.idc_network, pd.DataFrame) else st.session_state.idc_network,
            "menu1_data": st.session_state.menu1_data,
            "warranty_meta": st.session_state.warranty_meta,
            "warranty_repair_data": st.session_state.warranty_repair_data.to_dict(orient="records") if isinstance(st.session_state.warranty_repair_data, pd.DataFrame) else st.session_state.warranty_repair_data,
            "warranty_exchange_data": st.session_state.warranty_exchange_data.to_dict(orient="records") if isinstance(st.session_state.warranty_exchange_data, pd.DataFrame) else st.session_state.warranty_exchange_data,
            "audit_logs": st.session_state.audit_logs.to_dict(orient="records") if isinstance(st.session_state.audit_logs, pd.DataFrame) else st.session_state.audit_logs
        }
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# Khởi tạo hoặc nạp lại dữ liệu dùng chung vào Session State
storage_data = load_shared_storage()

if "master_schedule" not in st.session_state:
    st.session_state.master_schedule = pd.DataFrame(storage_data.get("master_schedule", []))

if "shift_change_requests" not in st.session_state:
    st.session_state.shift_change_requests = pd.DataFrame(storage_data.get("shift_change_requests", []))

if "tv_incidents" not in st.session_state:
    st.session_state.tv_incidents = pd.DataFrame(storage_data.get("tv_incidents", []))

if "hvac_schedule" not in st.session_state:
    st.session_state.hvac_schedule = pd.DataFrame(storage_data.get("hvac_schedule", []))

if "tech_params" not in st.session_state:
    st.session_state.tech_params = pd.DataFrame(storage_data.get("tech_params", []))

if "ups_params" not in st.session_state:
    st.session_state.ups_params = pd.DataFrame(storage_data.get("ups_params", []))

if "server_warnings" not in st.session_state:
    st.session_state.server_warnings = pd.DataFrame(storage_data.get("server_warnings", []))

if "idc_network" not in st.session_state:
    st.session_state.idc_network = pd.DataFrame(storage_data.get("idc_network", []))

if "menu1_data" not in st.session_state:
    st.session_state.menu1_data = storage_data.get("menu1_data", get_default_data()["menu1_data"])

if "warranty_meta" not in st.session_state:
    st.session_state.warranty_meta = storage_data.get("warranty_meta", get_default_data()["warranty_meta"])

if "warranty_repair_data" not in st.session_state:
    st.session_state.warranty_repair_data = pd.DataFrame(storage_data.get("warranty_repair_data", []))

if "warranty_exchange_data" not in st.session_state:
    st.session_state.warranty_exchange_data = pd.DataFrame(storage_data.get("warranty_exchange_data", []))

if "audit_logs" not in st.session_state:
    st.session_state.audit_logs = pd.DataFrame(storage_data.get("audit_logs", []))

if "email_noc_login" not in st.session_state:
    st.session_state.email_noc_login = {"email": "giamsatvtc.65lt@vtc.vn", "is_logged_in": False}

if "email_warranty_login" not in st.session_state:
    st.session_state.email_warranty_login = {"email": "baohanh@vtc.vn", "is_logged_in": False}

def add_audit_log(user, action, note=""):
    new_log = {
        "Thời gian": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Người dùng": user,
        "Thao tác": action,
        "Ghi chú": note
    }
    st.session_state.audit_logs = pd.concat([pd.DataFrame([new_log]), st.session_state.audit_logs], ignore_index=True)
    save_shared_storage()

# ---------------------------------------------------------
# 4. HÀM TẠO EXCEL BÁO CÁO CA A4 (ĐẦY ĐỦ 4 MỤC)
# ---------------------------------------------------------
def generate_excel_a4_report(current_shift_info):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bao_Cao_Ca_A4"
    
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    
    font_company = Font(name="Times New Roman", size=11, bold=True, color="002060")
    font_main_title = Font(name="Times New Roman", size=15, bold=True, color="002060")
    font_section = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    font_tbl_header = Font(name="Times New Roman", size=10, bold=True, color="002060")
    font_body = Font(name="Times New Roman", size=10, color="000000")
    font_body_bold = Font(name="Times New Roman", size=10, bold=True, color="000000")
    font_signature = Font(name="Times New Roman", size=10, italic=True)
    
    fill_section = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_tbl_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_meta = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    thin_side = Side(border_style="thin", color="D9D9D9")
    thick_dark = Side(border_style="thin", color="000000")
    
    border_data = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_header = Border(left=thick_dark, right=thick_dark, top=thick_dark, bottom=thick_dark)
    
    ws['A1'] = "VTC DIGITAL - PHÒNG KỸ THUẬT CÔNG NGHỆ"
    ws['A1'].font = font_company
    
    ws['A2'] = "BÁO CÁO TỔNG HỢP CA TRỰC PHÁT SÓNG & VẬN HÀNH NOC (A4)"
    ws['A2'].font = font_main_title
    
    ws.merge_cells('A4:H5')
    meta_cell = ws['A4']
    meta_cell.value = f"📍 Đơn vị: {current_shift_info['Trạm']}  |  🗓️ Ngày: {current_shift_info['Ngày']}  |  ⏰ Ca trực: {current_shift_info['Ca trực']}  |  👤 Kỹ sư: {current_shift_info['Người trực']}"
    meta_cell.font = font_body_bold
    meta_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    for r in range(4, 6):
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill_meta
            cell.border = border_header

    def write_section_title(start_row, title_text):
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=8)
        s_cell = ws.cell(row=start_row, column=1, value=title_text)
        s_cell.font = font_section
        s_cell.fill = fill_section
        s_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for c in range(1, 9):
            ws.cell(row=start_row, column=c).border = border_header

    def write_table_data(start_row, headers, df_data):
        for col_idx, h_text in enumerate(headers, 1):
            c = ws.cell(row=start_row, column=col_idx, value=h_text)
            c.font = font_tbl_header
            c.fill = fill_tbl_header
            c.border = border_header
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        curr_r = start_row + 1
        if df_data.empty:
            ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=len(headers))
            empty_cell = ws.cell(row=curr_r, column=1, value="Không ghi nhận sự cố / Thông số bình thường")
            empty_cell.font = font_body
            empty_cell.alignment = Alignment(horizontal="center", vertical="center")
            for c in range(1, len(headers) + 1):
                ws.cell(row=curr_r, column=c).border = border_data
            return curr_r + 1
        else:
            for _, row in df_data.iterrows():
                for col_idx, h_text in enumerate(headers, 1):
                    val = row.get(h_text, "")
                    c = ws.cell(row=curr_r, column=col_idx, value=val)
                    c.font = font_body
                    c.border = border_data
                    c.alignment = Alignment(horizontal="center" if col_idx in [1, 4, 5, 6] else "left", vertical="center", wrap_text=True)
                curr_r += 1
            return curr_r

    # MỤC 1: SỰ CỐ TRUYỀN HÌNH
    write_section_title(7, "1. TRUYỀN HÌNH (SỰ CỐ VÀ ĐƯỜNG TRUYỀN)")
    tv_cols = ["STT", "Ngày", "Tên kênh/nhóm kênh (Đường truyền)", "Hiện tượng", "Bắt đầu", "Kết thúc", "Thời lượng", "Nguyên nhân", "Biện pháp khắc phục (Bên khắc phục)"]
    next_r = write_table_data(8, tv_cols, st.session_state.tv_incidents) + 1

    # MỤC 2: ĐIỀU HÒA
    write_section_title(next_r, "2. HỆ THỐNG ĐIỀU HÒA PHÒNG MÁY NOC (TUÂN THỦ CHU KỲ 3 NGÀY)")
    hvac_cols = ["Ngày / Tuần", "Thời gian (Time)", "Chu kỳ", "Máy chạy (Chính)", "Máy nghỉ (Dự phòng)", "Ghi chú / Trạng thái"]
    next_r = write_table_data(next_r + 1, hvac_cols, st.session_state.hvac_schedule) + 1

    # MỤC 3: UPS VÀ HPA
    write_section_title(next_r, "3. HỆ THỐNG UPS VÀ THÔNG SỐ KỸ THUẬT HPA CA TRỰC")
    tech_cols = ["STT", "Thông số HPA", "Máy phát K1H-VNS1", "Ngưỡng tiêu chuẩn", "Đánh giá"]
    next_r = write_table_data(next_r + 1, tech_cols, st.session_state.tech_params) + 1

    # MỤC 4: SERVER (32 CON) VÀ MẠNG IDC
    write_section_title(next_r, "4. HỆ THỐNG SERVER (32 CON) VÀ MẠNG VĂN PHÒNG")
    server_cols = ["STT", "Tên Server", "Địa chỉ IP", "Chức năng", "CPU Util", "RAM Util", "Trạng thái Cảnh báo"]
    next_r = write_table_data(next_r + 1, server_cols, st.session_state.server_warnings) + 2

    ws.cell(row=next_r, column=2, value="NGƯỜI LẬP BÁO CÁO").font = font_body_bold
    ws.cell(row=next_r, column=7, value="XÁC NHẬN CỦA LÃNH ĐẠO").font = font_body_bold
    ws.cell(row=next_r+1, column=2, value="(Ký và ghi rõ họ tên)").font = font_signature
    ws.cell(row=next_r+1, column=7, value="(Ký và ghi rõ họ tên)").font = font_signature

    col_widths = {'A': 6, 'B': 24, 'C': 20, 'D': 12, 'E': 12, 'F': 12, 'G': 25, 'H': 28}
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ---------------------------------------------------------
# 5. HÀM TẠO EXCEL BÁO CÁO BẢO HÀNH GỘP 2 SHEET A4
# ---------------------------------------------------------
def generate_combined_warranty_excel():
    wb = openpyxl.Workbook()
    
    font_title = Font(name="Times New Roman", size=16, bold=True, color="002060")
    font_sub = Font(name="Times New Roman", size=12, bold=True, color="002060")
    font_info = Font(name="Times New Roman", size=11, italic=True)
    font_tbl_header = Font(name="Times New Roman", size=11, bold=True, color="002060")
    font_body = Font(name="Times New Roman", size=11)
    font_summary = Font(name="Times New Roman", size=11, bold=True, color="002060")

    fill_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_summary = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    # SHEET 1
    ws1 = wb.active
    ws1.title = "Sửa chữa"
    ws1.page_setup.paperSize = ws1.PAPERSIZE_A4
    ws1.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE

    ws1.merge_cells('A1:G1')
    ws1['A1'] = st.session_state.warranty_meta["title"]
    ws1['A1'].font = font_title
    ws1['A1'].alignment = Alignment(horizontal="center", vertical="center")

    ws1.merge_cells('A2:G2')
    ws1['A2'] = st.session_state.warranty_meta["week"]
    ws1['A2'].font = font_sub
    ws1['A2'].alignment = Alignment(horizontal="center", vertical="center")

    ws1.merge_cells('A3:G3')
    ws1['A3'] = f"Người thực hiện: {st.session_state.warranty_meta['executor']}"
    ws1['A3'].font = font_info

    ws1.merge_cells('A4:G4')
    ws1['A4'] = f"Nhập liệu: {st.session_state.warranty_meta['data_entry']}"
    ws1['A4'].font = font_info

    ws1['A5'] = "STT"
    ws1['B5'] = "Ngày nhập"
    ws1['C5'] = "Loại đầu thu"
    ws1['D5'] = "Mã dịch vụ"
    ws1['E5'] = "Sửa chữa\nThay thế"
    ws1['F5'] = "Tình trạng\nBảo hành"
    ws1.merge_cells('F5:G5')
    ws1['H5'] = "Ngày trả\n(hoàn thành)"

    for col_num in range(1, 9):
        cell = ws1.cell(row=5, column=col_num)
        cell.font = font_tbl_header
        cell.fill = fill_header
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    df_rep = st.session_state.warranty_repair_data
    start_r1 = 6
    for idx, row in df_rep.iterrows():
        r = start_r1 + idx
        ws1.cell(row=r, column=1, value=row.get("STT", idx+1)).alignment = Alignment(horizontal="center")
        ws1.cell(row=r, column=2, value=str(row.get("Ngày nhập", ""))).alignment = Alignment(horizontal="center")
        ws1.cell(row=r, column=3, value=str(row.get("Loại đầu thu", ""))).alignment = Alignment(horizontal="center")
        ws1.cell(row=r, column=4, value=str(row.get("Mã dịch vụ", ""))).alignment = Alignment(horizontal="center")
        ws1.cell(row=r, column=5, value=str(row.get("Sửa chữa / Thay thế", ""))).alignment = Alignment(horizontal="left")
        
        stt_bh = str(row.get("Tình trạng Bảo hành", ""))
        if stt_bh.strip().lower() == "còn":
            ws1.cell(row=r, column=6, value="Còn").alignment = Alignment(horizontal="center")
            ws1.cell(row=r, column=7, value="").alignment = Alignment(horizontal="center")
        elif stt_bh.strip().lower() == "hết":
            ws1.cell(row=r, column=6, value="").alignment = Alignment(horizontal="center")
            ws1.cell(row=r, column=7, value="Hết").alignment = Alignment(horizontal="center")
        else:
            ws1.cell(row=r, column=6, value=stt_bh).alignment = Alignment(horizontal="center")

        ws1.cell(row=r, column=8, value=row.get("Ngày trả (hoàn thành)", 1)).alignment = Alignment(horizontal="center")

        for col_num in range(1, 9):
            c = ws1.cell(row=r, column=col_num)
            c.font = font_body
            c.border = thin_border

    end_r1 = start_r1 + len(df_rep) - 1
    sum_r1 = max(end_r1 + 1, 15)

    ws1.merge_cells(start_row=sum_r1, start_column=1, end_row=sum_r1, end_column=7)
    sum_label1 = ws1.cell(row=sum_r1, column=1, value=f"Tổng số bảo hành, sửa chữa tuần ({st.session_state.warranty_meta['period']})")
    sum_label1.font = font_summary
    sum_label1.alignment = Alignment(horizontal="left", vertical="center")

    sum_val1 = ws1.cell(row=sum_r1, column=8, value=f"=SUM(H6:H{sum_r1-1})")
    sum_val1.font = font_summary
    sum_val1.alignment = Alignment(horizontal="center", vertical="center")

    for col_num in range(1, 9):
        c = ws1.cell(row=sum_r1, column=col_num)
        c.fill = fill_summary
        c.border = thin_border

    col_widths1 = {'A': 8, 'B': 14, 'C': 15, 'D': 16, 'E': 20, 'F': 12, 'G': 12, 'H': 18}
    for col_letter, width in col_widths1.items():
        ws1.column_dimensions[col_letter].width = width

    # SHEET 2
    ws2 = wb.create_sheet(title="Đổi bảo hành")
    ws2.page_setup.paperSize = ws2.PAPERSIZE_A4
    ws2.page_setup.orientation = ws2.ORIENTATION_LANDSCAPE

    ws2.merge_cells('A1:H1')
    ws2['A1'] = "Báo cáo đổi bảo hành sản phẩm VTC Digital"
    ws2['A1'].font = font_title
    ws2['A1'].alignment = Alignment(horizontal="center", vertical="center")

    ws2.merge_cells('A2:H2')
    ws2['A2'] = st.session_state.warranty_meta["week"]
    ws2['A2'].font = font_sub
    ws2['A2'].alignment = Alignment(horizontal="center", vertical="center")

    ws2.merge_cells('A3:H3')
    ws2['A3'] = f"Người thực hiện: {st.session_state.warranty_meta['executor']}"
    ws2['A3'].font = font_info

    ws2.merge_cells('A4:H4')
    ws2['A4'] = f"Nhập liệu: {st.session_state.warranty_meta['data_entry']}"
    ws2['A4'].font = font_info

    ex_headers = ["STT", "Ngày nhập", "Tên khách hàng/Địa chỉ", "Loại đầu thu", "Mã dịch vụ cũ", "Mã dịch vụ mới", "Người thực hiện", "Ngày trả\n(hoàn thành)"]
    for col_num, h_text in enumerate(ex_headers, 1):
        c = ws2.cell(row=5, column=col_num, value=h_text)
        c.font = font_tbl_header
        c.fill = fill_header
        c.border = thin_border
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    df_exc = st.session_state.warranty_exchange_data
    start_r2 = 6
    for idx, row in df_exc.iterrows():
        r = start_r2 + idx
        ws2.cell(row=r, column=1, value=row.get("STT", idx+1)).alignment = Alignment(horizontal="center")
        ws2.cell(row=r, column=2, value=str(row.get("Ngày nhập", ""))).alignment = Alignment(horizontal="center")
        ws2.cell(row=r, column=3, value=str(row.get("Tên khách hàng / Địa chỉ", ""))).alignment = Alignment(horizontal="left")
        ws2.cell(row=r, column=4, value=str(row.get("Loại đầu thu", ""))).alignment = Alignment(horizontal="center")
        ws2.cell(row=r, column=5, value=str(row.get("Mã dịch vụ cũ", ""))).alignment = Alignment(horizontal="center")
        ws2.cell(row=r, column=6, value=str(row.get("Mã dịch vụ mới", ""))).alignment = Alignment(horizontal="center")
        ws2.cell(row=r, column=7, value=str(row.get("Người thực hiện", ""))).alignment = Alignment(horizontal="left")
        ws2.cell(row=r, column=8, value=row.get("Ngày trả (hoàn thành)", 1)).alignment = Alignment(horizontal="center")

        for col_num in range(1, 9):
            c = ws2.cell(row=r, column=col_num)
            c.font = font_body
            c.border = thin_border

    end_r2 = start_r2 + len(df_exc) - 1
    sum_r2 = max(end_r2 + 1, 15)

    ws2.merge_cells(start_row=sum_r2, start_column=1, end_row=sum_r2, end_column=7)
    sum_label2 = ws2.cell(row=sum_r2, column=1, value=f"Tổng số bảo hành, sửa chữa tuần ({st.session_state.warranty_meta['period']})")
    sum_label2.font = font_summary
    sum_label2.alignment = Alignment(horizontal="left", vertical="center")

    sum_val2 = ws2.cell(row=sum_r2, column=8, value=f"=SUM(H6:H{sum_r2-1})")
    sum_val2.font = font_summary
    sum_val2.alignment = Alignment(horizontal="center", vertical="center")

    for col_num in range(1, 9):
        c = ws2.cell(row=sum_r2, column=col_num)
        c.fill = fill_summary
        c.border = thin_border

    col_widths2 = {'A': 8, 'B': 14, 'C': 22, 'D': 15, 'E': 16, 'F': 16, 'G': 20, 'H': 18}
    for col_letter, width in col_widths2.items():
        ws2.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ---------------------------------------------------------
# 6. GIẢ LẬP GỬI EMAIL TỰ ĐỘNG
# ---------------------------------------------------------
def send_email_notification(from_email, to_email, subject, body_text, attachment_bytes=None, filename="BaoCao.xlsx"):
    try:
        return True, f"✅ Đã gửi email thành công từ **{from_email}** tới **{to_email}**!"
    except Exception as e:
        return False, f"⚠️ Lỗi gửi mail: {e}"

# ---------------------------------------------------------
# 7. THANH BÊN (SIDEBAR) & XÁC THỰC RÔLE
# ---------------------------------------------------------
st.sidebar.markdown("### CÔNG TY VTC DỊCH VỤ TRUYỀN HÌNH SỐ")
st.sidebar.markdown("### 🤖 AI PHÒNG KỸ THUẬT CÔNG NGHỆ")
st.sidebar.markdown("## HỆ THỐNG QUẢN LÝ VẬN HÀNH NOC & BẢO HÀNH SẢN PHẨM")
st.sidebar.divider()

access_type = st.sidebar.selectbox("🌐 Vùng truy cập:", ["Local (Mạng Nội bộ)", "Internet (MFA)"])

user_role = st.sidebar.selectbox("👤 Vai trò hệ thống:", [
    "Admin / Lãnh đạo Phòng", 
    "Kỹ sư Trực ca NOC", 
    "Kỹ sư Bảo hành", 
    "Viewer (Chỉ xem)"
])

role_passwords = {
    "Admin / Lãnh đạo Phòng": "Vtc@123",
    "Kỹ sư Trực ca NOC": "Digital%",
    "Kỹ sư Bảo hành": "Digital%"
}

is_authenticated = False
if user_role in role_passwords:
    pwd_input = st.sidebar.text_input(f"🔑 Nhập Mật khẩu ({user_role}):", type="password")
    if pwd_input == role_passwords[user_role]:
        st.sidebar.success("✅ Xác thực vai trò thành công!")
        is_authenticated = True
    else:
        if pwd_input != "":
            st.sidebar.error("❌ Mật khẩu vai trò không đúng!")
        else:
            st.sidebar.info("🔒 Vui lòng nhập mật khẩu vai trò.")
else:
    is_authenticated = True

is_admin = (user_role == "Admin / Lãnh đạo Phòng") and is_authenticated

st.sidebar.divider()

st.sidebar.divider()
if st.sidebar.button("💾 ĐỒNG BỘ & LƯU TẤT CẢ DỮ LIỆU", type="primary", use_container_width=True):
    save_shared_storage()
    st.sidebar.success("✅ Đã đồng bộ & lưu dữ liệu vào hệ thống máy chủ!")
    st.rerun()

if st.sidebar.button("🔄 TẢI LẠI DỮ LIỆU TỪ MÁY CHỦ", use_container_width=True):
    refreshed_data = load_shared_storage()
    st.session_state.master_schedule = pd.DataFrame(refreshed_data.get("master_schedule", []))
    st.session_state.shift_change_requests = pd.DataFrame(refreshed_data.get("shift_change_requests", []))
    st.session_state.tv_incidents = pd.DataFrame(refreshed_data.get("tv_incidents", []))
    st.session_state.hvac_schedule = pd.DataFrame(refreshed_data.get("hvac_schedule", []))
    st.session_state.tech_params = pd.DataFrame(refreshed_data.get("tech_params", []))
    st.session_state.ups_params = pd.DataFrame(refreshed_data.get("ups_params", []))
    st.session_state.server_warnings = pd.DataFrame(refreshed_data.get("server_warnings", []))
    st.session_state.idc_network = pd.DataFrame(refreshed_data.get("idc_network", []))
    st.session_state.menu1_data = refreshed_data.get("menu1_data", {})
    st.session_state.warranty_meta = refreshed_data.get("warranty_meta", {})
    st.session_state.warranty_repair_data = pd.DataFrame(refreshed_data.get("warranty_repair_data", []))
    st.session_state.warranty_exchange_data = pd.DataFrame(refreshed_data.get("warranty_exchange_data", []))
    st.session_state.audit_logs = pd.DataFrame(refreshed_data.get("audit_logs", []))
    st.sidebar.info("🔄 Đã làm mới dữ liệu mới nhất từ máy chủ!")
    st.rerun()



# ---------------------------------------------------------
# TIỆN ÍCH LIÊN HỆ ONLINE (THU GỌN)
# ---------------------------------------------------------
HOTLINE_NUMBER = "0913332569"
PHONE_INTERNATIONAL = "84913332569"

st.sidebar.divider()
with st.sidebar.expander("💬 Hỗ trợ nhanh Zalo / Viber", expanded=False):
    st.markdown(f"📞 **Hotline:** `{HOTLINE_NUMBER}`")
    c_z, c_v = st.columns(2)
    c_z.link_button("💬 Zalo", f"https://zalo.me/{HOTLINE_NUMBER}", use_container_width=True)
    c_v.link_button("🟣 Viber", f"viber://chat?number=%2B{PHONE_INTERNATIONAL}", use_container_width=True)
    st.caption("Quét mã QR trên điện thoại:")
    st.image(f"https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=https://zalo.me/{HOTLINE_NUMBER}", caption="QR Zalo 0913332569", use_container_width=True)

menu = st.sidebar.radio(
    "📋 Danh mục Chức năng:",
    [
        "1. Báo cáo Tổng hợp Ca trực (A4)",
        "2. Quản lý Phân ca, Duyệt Đổi ca",
        "3. Phân tích Đối soát",
        "4. Quản lý Bảo hành (Sửa chữa & Đổi bảo hành)",
        "5. Giám sát Luồng Kênh Catchup (HLS Monitor)",
        "6. Không gian Trao đổi Zalo & Viber (0913332569)",
        "7. Lưu trữ và Tài liệu AI",
        "8. Nhật ký Hoạt động"
    ]
)

# ---------------------------------------------------------
# HÀM HIỂN THỊ KHU VỰC ĐĂNG NHẬP EMAIL ĐẦU TRANG
# ---------------------------------------------------------
def render_email_login_header(email_key, title_label):
    st.markdown(f"### 📧 Đăng nhập Email Gửi Báo cáo & Thông báo: `{st.session_state[email_key]['email']}`")
    with st.expander(f"🔑 Cấu hình Đăng nhập Tài khoản Email: {st.session_state[email_key]['email']}", expanded=not st.session_state[email_key]["is_logged_in"]):
        col_m1, col_m2, col_m3 = st.columns([2, 2, 1])
        email_addr = col_m1.text_input("Địa chỉ Email gửi:", value=st.session_state[email_key]["email"], disabled=True, key=f"inp_email_{email_key}")
        email_pass = col_m2.text_input("Mật khẩu Email / App Password:", type="password", key=f"inp_pass_{email_key}")
        
        st.write("")
        if st.session_state[email_key]["is_logged_in"]:
            st.success(f"🟢 Email **{st.session_state[email_key]['email']}** đang ở trạng thái **SẴN SÀNG GỬI**")
        else:
            if col_m3.button("🔓 Đăng nhập Email", key=f"btn_login_{email_key}"):
                if email_pass:
                    st.session_state[email_key]["is_logged_in"] = True
                    add_audit_log(user_role, f"Đăng nhập Email thành công: {st.session_state[email_key]['email']}")
                    st.success(f"✅ Đã kết nối thành công tài khoản Email **{st.session_state[email_key]['email']}**!")
                    st.rerun()
                else:
                    st.error("⚠️ Vui lòng nhập mật khẩu Email!")

# ---------------------------------------------------------
# 8. CHI TIẾT CÁC MENU CHỨC NĂNG
# ---------------------------------------------------------

# MENU 1: BÁO CÁO TỔNG HỢP CA TRỰC A4
if menu == "1. Báo cáo Tổng hợp Ca trực (A4)":
    st.title("📄 Báo cáo tổng hợp ca trực NOC & Vận hành (Đầy đủ 4 Mục)")
    
    # ĐĂNG NHẬP EMAIL ĐẦU MENU 1
    render_email_login_header("email_noc_login", "NOC Mail")
    st.divider()

    st.subheader("📌 Thông tin ca trực Master")
    available_dates = list(st.session_state.master_schedule["Ngày"].unique())
    col_sel1, col_sel2 = st.columns(2)
    selected_date = col_sel1.selectbox("🗓️ Chọn Ngày trực:", available_dates)
    
    shifts_in_date = list(st.session_state.master_schedule[st.session_state.master_schedule["Ngày"] == selected_date]["Ca trực"].unique())
    selected_shift = col_sel2.selectbox("⏰ Chọn Ca trực:", shifts_in_date)

    match_row = st.session_state.master_schedule[
        (st.session_state.master_schedule["Ngày"] == selected_date) & 
        (st.session_state.master_schedule["Ca trực"] == selected_shift)
    ]

    if not match_row.empty:
        curr_info = match_row.iloc[0].to_dict()
    else:
        curr_info = {"Trạm": "Trạm Phát sóng VTC Digital", "Ngày": selected_date, "Ca trực": selected_shift, "Người trực": "Chưa gán"}

    col1, col2, col3, col4 = st.columns(4)
    col1.text_input("Trạm:", value=curr_info["Trạm"], disabled=True)
    col2.text_input("Ngày:", value=curr_info["Ngày"], disabled=True)
    col3.text_input("Ca trực:", value=curr_info["Ca trực"], disabled=True)
    col4.text_input("Kỹ sư trực ca:", value=curr_info["Người trực"], disabled=True)

    st.divider()

    col_b1, col_b2 = st.columns(2)
    excel_a4_bytes = generate_excel_a4_report(curr_info)
    
    with col_b1:
        st.download_button(
            label="👁️ Tải/Xuất Báo cáo Ca Excel A4 đầy đủ 4 Mục",
            data=excel_a4_bytes,
            file_name=f"Bao_Cao_Ca_A4_Full_{selected_date.replace('/','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col_b2:
        if st.button("🚀 Kết thúc ca trực (Gửi Báo cáo A4 Đầy đủ 4 Mục tới Lãnh đạo)", type="primary", use_container_width=True):
            if st.session_state.email_noc_login["is_logged_in"]:
                status, msg = send_email_notification(
                    from_email=st.session_state.email_noc_login["email"],
                    to_email="anh.lehoang@vtc.vn",
                    subject=f"[BÁO CÁO CA A4 FULL] {curr_info['Ngày']} - {curr_info['Ca trực']}",
                    body_text=f"Báo cáo ca trực A4 đầy đủ 4 mục gửi tới Lãnh đạo từ {st.session_state.email_noc_login['email']}",
                    attachment_bytes=excel_a4_bytes,
                    filename=f"Bao_Cao_Ca_A4_{curr_info['Ngày'].replace('/','_')}.xlsx"
                )
                st.success(msg)
                add_audit_log(curr_info['Người trực'], "Kết thúc ca trực & Gửi email A4 Đầy đủ 4 Mục", "Gửi tới anh.lehoang@vtc.vn")
            else:
                st.error(f"🔒 Bạn phải Đăng nhập Email {st.session_state.email_noc_login['email']} ở đầu trang trước khi gửi!")

    st.divider()
    
    # ---------------- 1. TRUYỀN HÌNH ----------------
    st.subheader("1. TRUYỀN HÌNH (Sự cố đường truyền & Giám sát luồng)")
    
    with st.expander("📺 Giám sát Trực tiếp Luồng Kênh Catchup (falconhlsmonitor.vtcdigital.top)", expanded=False):
        c_m_top1, c_m_top2 = st.columns([3, 1])
        c_m_top1.write("Theo dõi trạng thái phát sóng luồng HLS Catchup trực tiếp:")
        c_m_top2.link_button("🚀 Mở Cửa Sổ Giám Sát Lớn", "https://falconhlsmonitor.vtcdigital.top/", use_container_width=True)
        st.components.v1.iframe("https://falconhlsmonitor.vtcdigital.top/", height=500, scrolling=True)

    tab_tv_form, tab_tv_table = st.tabs(["📝 Nhập sự cố bằng Form (Tự động tính thời lượng)", "📊 Chỉnh sửa trực tiếp Bảng Sự cố"])
    
    with tab_tv_form:
        with st.form("form_tv_incident"):
            st.markdown("**Khai báo sự cố đường truyền kênh truyền hình:**")
            c_date, c_tv1 = st.columns(2)
            tv_date = c_date.text_input("🗓️ Ngày xảy ra sự cố:", value=datetime.now().strftime("%d/%m/%Y"))
            tv_channel = c_tv1.text_input("Tên kênh/nhóm kênh (Đường truyền):", value="VTV1 / VTV Cab")
            
            c_tv2, c_tv3, c_tv4 = st.columns([2, 1, 1])
            tv_symptom = c_tv2.text_input("Hiện tượng sự cố:", value="Mất tín hiệu luồng IP")
            time_start = c_tv3.time_input("Bắt đầu:", value=time(9, 0))
            time_end = c_tv4.time_input("Kết thúc:", value=time(10, 30))
            
            c_tv5, c_tv6 = st.columns(2)
            tv_cause = c_tv5.text_input("Nguyên nhân:", value="Lỗi thiết bị Switch truyền dẫn")
            tv_fix = c_tv6.text_input("Biện pháp khắc phục (Bên khắc phục):", value="Khởi động lại Switch / Đội NOC")

            if st.form_submit_button("➕ Thêm Sự cố vào Báo cáo"):
                dt_start = datetime.combine(datetime.today(), time_start)
                dt_end = datetime.combine(datetime.today(), time_end)
                if dt_end < dt_start:
                    dt_end = dt_end.replace(day=dt_end.day + 1)
                
                diff_sec = int((dt_end - dt_start).total_seconds())
                hours = diff_sec // 3600
                minutes = (diff_sec % 3600) // 60
                duration_str = f"{hours:02d}h {minutes:02d}m"

                new_inc = {
                    "STT": len(st.session_state.tv_incidents) + 1,
                    "Ngày": tv_date,
                    "Tên kênh/nhóm kênh (Đường truyền)": tv_channel,
                    "Hiện tượng": tv_symptom,
                    "Bắt đầu": time_start.strftime("%H:%M"),
                    "Kết thúc": time_end.strftime("%H:%M"),
                    "Thời lượng": duration_str,
                    "Nguyên nhân": tv_cause,
                    "Biện pháp khắc phục (Bên khắc phục)": tv_fix
                }
                st.session_state.tv_incidents = pd.concat([st.session_state.tv_incidents, pd.DataFrame([new_inc])], ignore_index=True)
                save_shared_storage()
                st.success(f"✅ Đã thêm sự cố! Thời lượng tự động tính toán: **{duration_str}**")
                st.rerun()

    with tab_tv_table:
        st.session_state.tv_incidents = st.data_editor(st.session_state.tv_incidents, num_rows="dynamic", use_container_width=True, key="ed_tv")

    st.divider()

    # ---------------- 2. ĐIỀU HÒA LUÂN PHIÊN ----------------
    st.subheader("2. Hệ thống điều hòa luân phiên (Tịnh tiến chu kỳ 3 ngày)")
    st.caption("🔄 Máy điều hòa tự động luân phiên xoay vòng theo thời gian & ngày (Hết 3 ngày tịnh tiến quay lại Ngày 1).")
    
    with st.expander("➕ Thêm hàng theo dõi Điều hòa theo Thời gian (Time)", expanded=False):
        with st.form("form_add_hvac"):
            h1, h2, h3 = st.columns(3)
            hvac_date = h1.text_input("Ngày theo dõi:", value=datetime.now().strftime("%d/%m/%Y"))
            hvac_time = h2.text_input("Thời gian (Time):", value="08:00 - 20:00")
            hvac_cycle = h3.selectbox("Chu kỳ tịnh tiến:", ["Ngày 1", "Ngày 2", "Ngày 3"])

            h4, h5 = st.columns(2)
            hvac_run = h4.text_input("Máy chạy (Chính):", value="Máy 1 – Máy 3 – Máy 4 – Máy 5 – Máy 7")
            hvac_stop = h5.text_input("Máy nghỉ (Dự phòng):", value="Máy 6 – Máy 8")
            hvac_note = st.text_input("Ghi chú / Trạng thái:", value="Hoạt động ổn định")

            if st.form_submit_button("Lưu cấu hình Điều hòa"):
                new_hvac = {
                    "Ngày / Tuần": hvac_date,
                    "Thời gian (Time)": hvac_time,
                    "Chu kỳ": hvac_cycle,
                    "Máy chạy (Chính)": hvac_run,
                    "Máy nghỉ (Dự phòng)": hvac_stop,
                    "Ghi chú / Trạng thái": hvac_note
                }
                st.session_state.hvac_schedule = pd.concat([st.session_state.hvac_schedule, pd.DataFrame([new_hvac])], ignore_index=True)
                save_shared_storage()
                st.success("✅ Đã thêm hàng theo dõi điều hòa mới!")
                st.rerun()

    st.session_state.hvac_schedule = st.data_editor(st.session_state.hvac_schedule, num_rows="dynamic", use_container_width=True, key="ed_hvac")

    st.divider()

    # ---------------- 3. UPS & THÔNG SỐ HPA ----------------
    st.subheader("3. Hệ thống UPS và Thông số Kỹ thuật HPA")
    
    col_link1, col_link2 = st.columns([3, 1])
    with col_link1:
        st.markdown("🔗 **Liên kết giám sát trực tiếp UPS & HPA:** [http://192.168.20.201](http://192.168.20.201)")
    with col_link2:
        st.link_button("🌐 Mở Hệ thống 192.168.20.201", "http://192.168.20.201", use_container_width=True)

    st.markdown("**a. Bảng Thông số Hệ thống UPS:**")
    st.session_state.ups_params = st.data_editor(st.session_state.ups_params, num_rows="dynamic", use_container_width=True, key="ed_ups")

    st.markdown("**b. Bảng Thông số Kỹ thuật Máy phát HPA:**")
    st.session_state.tech_params = st.data_editor(st.session_state.tech_params, num_rows="dynamic", use_container_width=True, key="ed_tech")

    st.divider()

    # ---------------- 4. SERVER & MẠNG VĂN PHÒNG ----------------
    st.subheader("4. Hệ thống Server và Mạng văn phòng NOC")
    
    st.markdown("🚨 **Bảng Cảnh báo & Giám sát Hệ thống Server (Danh sách 32 Server NOC):**")
    st.session_state.server_warnings = st.data_editor(
        st.session_state.server_warnings, 
        num_rows="dynamic", 
        use_container_width=True, 
        key="ed_servers"
    )

    st.markdown("**Bảng Giám sát Đường truyền & Mạng Văn phòng:**")
    st.session_state.idc_network = st.data_editor(st.session_state.idc_network, num_rows="dynamic", use_container_width=True, key="ed_idc")

# MENU 2: QUẢN LÝ PHÂN CA, DUYỆT ĐỔI CA
elif menu == "2. Quản lý Phân ca, Duyệt Đổi ca":
    st.title("👥 Quản lý phân ca & duyệt đổi ca KTCN")
    
    # ĐĂNG NHẬP EMAIL ĐẦU MENU 2
    render_email_login_header("email_noc_login", "NOC Mail")
    st.divider()

    tab1, tab2, tab3, tab4 = st.tabs([
        "📅 Lịch Master Đã Đồng Bộ", 
        "📤 Upload File Lịch Trực VTC", 
        "🔄 Đăng ký Đổi ca", 
        "👑 Lãnh đạo Duyệt Đổi ca"
    ])

    with tab1:
        st.dataframe(st.session_state.master_schedule, use_container_width=True)

    with tab2:
        st.subheader("Upload File Excel Lịch trực KTCN (.xlsx)")
        uploaded_file = st.file_uploader("Chọn file ma trận công:", type=["xlsx", "xls"])
        if uploaded_file is not None:
            parsed_df = parse_vtc_matrix_schedule(uploaded_file)
            if parsed_df is not None:
                st.success("✅ Bóc tách thành công Ma trận Lịch trực VTC!")
                st.dataframe(parsed_df, use_container_width=True)
                if st.button("🔥 Lưu & Cập nhật Lịch Master", type="primary"):
                    st.session_state.master_schedule = parsed_df
                    save_shared_storage()
                    add_audit_log(user_role, "Upload & Cập nhật Lịch trực Master mới")
                    st.success("🎉 Đã cập nhật Lịch Master!")
                    st.rerun()

    with tab3:
        st.subheader("Tạo yêu cầu đổi ca")
        with st.form("form_change_shift"):
            c1, c2 = st.columns(2)
            req_date = c1.selectbox("Ngày trực:", st.session_state.master_schedule["Ngày"].unique())
            req_shift = c2.selectbox("Ca trực:", st.session_state.master_schedule[st.session_state.master_schedule["Ngày"] == req_date]["Ca trực"].unique())
            
            c3, c4 = st.columns(2)
            req_person = c3.text_input("Người xin đổi:")
            sub_person = c4.text_input("Người trực thay:")
            reason = st.text_area("Lý do đổi ca:")
            
            if st.form_submit_button("Gửi yêu cầu duyệt"):
                if st.session_state.email_noc_login["is_logged_in"]:
                    new_req = {
                        "Mã GD": f"DC00{len(st.session_state.shift_change_requests)+1}",
                        "Ngày": req_date,
                        "Ca trực": req_shift,
                        "Người xin đổi": req_person,
                        "Người trực thay": sub_person,
                        "Lý do": reason,
                        "Trạng thái": "Chờ duyệt"
                    }
                    st.session_state.shift_change_requests = pd.concat([pd.DataFrame([new_req]), st.session_state.shift_change_requests], ignore_index=True)
                    save_shared_storage()
                    
                    status, msg = send_email_notification(
                        from_email=st.session_state.email_noc_login["email"],
                        to_email="anh.lehoang@vtc.vn",
                        subject=f"[YÊU CẦU ĐỔI CA] {req_person} -> {sub_person} ({req_date})",
                        body_text=f"Nhân sự {req_person} gửi yêu cầu đổi ca cho {sub_person} ngày {req_date} ({req_shift}). Lý do: {reason}"
                    )
                    add_audit_log(req_person, "Gửi yêu cầu đổi ca", f"Mã {new_req['Mã GD']}")
                    st.success(f"✅ Đã gửi yêu cầu đổi ca từ email **{st.session_state.email_noc_login['email']}** tới anh.lehoang@vtc.vn!")
                else:
                    st.error(f"🔒 Bạn phải Đăng nhập Email **{st.session_state.email_noc_login['email']}** ở đầu trang trước khi đăng ký đổi ca!")

    with tab4:
        st.subheader("👑 Duyệt Đổi ca (Gửi Email thông báo: anh.lehoang@vtc.vn)")
        if is_admin:
            pending_df = st.session_state.shift_change_requests[st.session_state.shift_change_requests["Trạng thái"] == "Chờ duyệt"]
            if pending_df.empty:
                st.info("Hiện không có yêu cầu nào chờ duyệt.")
            else:
                for idx, row in pending_df.iterrows():
                    with st.expander(f"📌 Mã GD: {row['Mã GD']} - Ngày: {row['Ngày']} ({row['Ca trực']})", expanded=True):
                        st.write(f"• **Người xin đổi:** {row['Người xin đổi']}  ➡️  **Người trực thay:** {row['Người trực thay']}")
                        st.write(f"• **Lý do:** {row['Lý do']}")
                        
                        col_a, col_b = st.columns(2)
                        if col_a.button(f"✅ DUYỆT ({row['Mã GD']})", type="primary"):
                            if st.session_state.email_noc_login["is_logged_in"]:
                                st.session_state.shift_change_requests.loc[st.session_state.shift_change_requests["Mã GD"] == row["Mã GD"], "Trạng thái"] = "Đã duyệt"
                                
                                m_mask = (st.session_state.master_schedule["Ngày"] == row["Ngày"]) & (st.session_state.master_schedule["Ca trực"] == row["Ca trực"])
                                if not st.session_state.master_schedule[m_mask].empty:
                                    old_staff = st.session_state.master_schedule.loc[m_mask, "Người trực"].values[0]
                                    new_staff = old_staff.replace(row["Người xin đổi"], row["Người trực thay"]) if row["Người xin đổi"] in old_staff else f"{old_staff}, {row['Người trực thay']}"
                                    st.session_state.master_schedule.loc[m_mask, "Người trực"] = new_staff
                                save_shared_storage()
                                
                                send_email_notification(
                                    from_email=st.session_state.email_noc_login["email"],
                                    to_email="anh.lehoang@vtc.vn",
                                    subject=f"[ĐÃ DUYỆT ĐỔI CA] Mã GD: {row['Mã GD']}",
                                    body_text=f"Lãnh đạo đã phê duyệt đổi ca cho nhân sự: {row['Người xin đổi']} -> {row['Người trực thay']} vào ngày {row['Ngày']} ({row['Ca trực']})."
                                )
                                add_audit_log(user_role, "Duyệt đổi ca", f"Mã GD {row['Mã GD']}")
                                st.success(f"🎉 Đã duyệt yêu cầu {row['Mã GD']} và thông báo tới anh.lehoang@vtc.vn!")
                                st.rerun()
                            else:
                                st.error(f"🔒 Đăng nhập Email **{st.session_state.email_noc_login['email']}** ở đầu trang để duyệt!")

                        if col_b.button(f"❌ TỪ CHỐI ({row['Mã GD']})"):
                            st.session_state.shift_change_requests.loc[st.session_state.shift_change_requests["Mã GD"] == row["Mã GD"], "Trạng thái"] = "Từ chối"
                            save_shared_storage()
                            st.warning(f"Đã từ chối yêu cầu {row['Mã GD']}.")
                            st.rerun()
        else:
            st.error("🔒 Bạn cần đăng nhập Mật khẩu với Vai trò Admin / Lãnh đạo để phê duyệt.")

# MENU 3: PHÂN TÍCH ĐỐI SOÁT
elif menu == "3. Phân tích Đối soát":
    st.title("📊 Phân tích & Đối soát Sự cố")
    
    # ĐĂNG NHẬP EMAIL ĐẦU MENU 3
    render_email_login_header("email_noc_login", "NOC Mail")
    st.divider()

    st.markdown("🔗 **Liên kết nhanh:** [Mở Bảng Sự cố Đối soát Chi tiết (Spreadsheet / Log System)](#)")
    
    st.subheader("Bảng Ghi nhận Đối soát Sự cố Hạ tầng & Kênh")
    st.session_state.tv_incidents = st.data_editor(
        st.session_state.tv_incidents, 
        num_rows="dynamic", 
        use_container_width=True,
        key="incidents_reconcile"
    )

    if st.button("📧 Gửi Email Đối soát Sự cố tới Đối tác", type="primary"):
        if st.session_state.email_noc_login["is_logged_in"]:
            status, msg = send_email_notification(
                from_email=st.session_state.email_noc_login["email"],
                to_email="doitac.truyendan@vtc.vn",
                subject="[ĐỐI SOÁT SỰ CỐ TRUYỀN DẪN VTC]",
                body_text="Gửi bảng đối soát sự cố kênh truyền hình trong ca..."
            )
            st.success(msg)
        else:
            st.error(f"🔒 Vui lòng Đăng nhập Email **{st.session_state.email_noc_login['email']}** ở đầu trang!")

# MENU 4: QUẢN LÝ BẢO HÀNH
elif menu == "4. Quản lý Bảo hành (Sửa chữa & Đổi bảo hành)":
    st.title("🛡️ Quản lý & Lập Báo cáo Bảo hành Sản phẩm VTC Digital")
    
    # ĐĂNG NHẬP EMAIL ĐẦU MENU 4 (BAO HANH MAIL)
    render_email_login_header("email_warranty_login", "Warranty Mail")
    st.divider()

    st.info("🔗 **Cổng thông tin bảo hành trực tuyến:** [http://baohanh.truyenhinhso.vn](http://baohanh.truyenhinhso.vn)")

    tab_edit, tab_exchange, tab_add, tab_web = st.tabs([
        "📝 Bảng Dữ liệu Sửa chữa", 
        "🔄 Bảng Dữ liệu Đổi bảo hành", 
        "➕ Thêm lượt Bảo hành mới", 
        "🌐 Web Cổng Bảo hành VTC"
    ])

    st.subheader("⚙️ Thông tin chung Báo cáo Tuần gửi Lãnh đạo")
    c_m1, c_m2 = st.columns(2)
    st.session_state.warranty_meta["title"] = c_m1.text_input("Tiêu đề Báo cáo:", value=st.session_state.warranty_meta["title"])
    st.session_state.warranty_meta["week"] = c_m2.text_input("Tuần / Tháng:", value=st.session_state.warranty_meta["week"])

    c_m3, c_m4, c_m5 = st.columns(3)
    st.session_state.warranty_meta["executor"] = c_m3.text_input("Người thực hiện:", value=st.session_state.warranty_meta["executor"])
    st.session_state.warranty_meta["data_entry"] = c_m4.text_input("Nhập liệu:", value=st.session_state.warranty_meta["data_entry"])
    st.session_state.warranty_meta["period"] = c_m5.text_input("Khoảng thời gian tuần:", value=st.session_state.warranty_meta["period"])

    st.divider()

    with tab_edit:
        st.subheader("📊 1. Bảng Dữ liệu Sửa chữa / Thay thế (Chỉnh sửa trực tiếp dạng Bảng)")
        st.session_state.warranty_repair_data = st.data_editor(
            st.session_state.warranty_repair_data, 
            num_rows="dynamic", 
            use_container_width=True,
            key="w_repair_editor"
        )

    with tab_exchange:
        st.subheader("🔄 2. Bảng Dữ liệu Đổi bảo hành (Chỉnh sửa trực tiếp dạng Bảng)")
        st.session_state.warranty_exchange_data = st.data_editor(
            st.session_state.warranty_exchange_data, 
            num_rows="dynamic", 
            use_container_width=True,
            key="w_exchange_editor"
        )

    st.divider()
    st.subheader("🚀 Báo cáo Tuần Gộp 2 Sheet gửi Lãnh đạo")

    combined_warranty_excel_bytes = generate_combined_warranty_excel()

    col_w1, col_w2 = st.columns(2)
    with col_w1:
        st.download_button(
            label="📥 Tải Báo cáo Tuần Gộp Excel (Sửa chữa + Đổi bảo hành A4)",
            data=combined_warranty_excel_bytes,
            file_name=f"Bao_Cao_Bao_Hanh_Tong_Hop_{st.session_state.warranty_meta['week'].replace(' ','_').replace('/','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col_w2:
        if st.button("📧 Gửi Báo cáo Tuần Gộp cho Lãnh đạo (tuan.ngoc@vtc.vn)", type="primary", use_container_width=True):
            if st.session_state.email_warranty_login["is_logged_in"]:
                status, msg = send_email_notification(
                    from_email=st.session_state.email_warranty_login["email"],
                    to_email="tuan.ngoc@vtc.vn",
                    subject=f"[BÁO CÁO TỔNG HỢP BẢO HÀNH A4] {st.session_state.warranty_meta['week']}",
                    body_text=f"Báo cáo tổng hợp bảo hành & đổi bảo hành sản phẩm VTC Digital gửi tới Lãnh đạo tuan.ngoc@vtc.vn từ {st.session_state.email_warranty_login['email']}",
                    attachment_bytes=combined_warranty_excel_bytes,
                    filename=f"Bao_Cao_Bao_Hanh_Tong_Hop_{st.session_state.warranty_meta['week'].replace(' ','_').replace('/','_')}.xlsx"
                )
                st.success(msg)
                add_audit_log(user_role, "Gửi báo cáo bảo hành gộp Excel A4", "Gửi tới tuan.ngoc@vtc.vn")
            else:
                st.error(f"🔒 Bạn phải Đăng nhập Email **{st.session_state.email_warranty_login['email']}** ở đầu trang trước khi gửi!")

    with tab_add:
        st.subheader("➕ Thêm lượt Bảo hành / Đổi bảo hành mới")
        add_type = st.radio("Chọn loại bản ghi:", ["Sửa chữa / Thay thế", "Đổi bảo hành"], horizontal=True)
        
        with st.form("form_add_warranty_entry"):
            if add_type == "Sửa chữa / Thay thế":
                f1, f2 = st.columns(2)
                w_date = f1.text_input("Ngày nhập:", value=datetime.now().strftime("%d/%m/%Y"))
                w_device = f2.text_input("Loại đầu thu:", value="HDV2")

                f3, f4 = st.columns(2)
                w_code = f3.text_input("Mã dịch vụ:", value="4256419430")
                w_fix = f4.text_input("Sửa chữa / Thay thế:", value="Nguồn")

                f5, f6 = st.columns(2)
                w_status = f5.selectbox("Tình trạng Bảo hành:", ["Còn", "Hết"])
                w_return = f6.number_input("Ngày trả (số lượng):", min_value=1, value=1)

                if st.form_submit_button("Lưu bản ghi Sửa chữa"):
                    new_w = {
                        "STT": len(st.session_state.warranty_repair_data) + 1,
                        "Ngày nhập": w_date,
                        "Loại đầu thu": w_device,
                        "Mã dịch vụ": w_code,
                        "Sửa chữa / Thay thế": w_fix,
                        "Tình trạng Bảo hành": w_status,
                        "Ngày trả (hoàn thành)": w_return
                    }
                    st.session_state.warranty_repair_data = pd.concat([st.session_state.warranty_repair_data, pd.DataFrame([new_w])], ignore_index=True)
                    save_shared_storage()
                    add_audit_log(user_role, "Thêm lượt sửa chữa bảo hành", f"Mã DV {w_code}")
                    st.success("✅ Đã thêm lượt sửa chữa mới!")
                    st.rerun()

            else:
                f1, f2 = st.columns(2)
                ex_date = f1.text_input("Ngày nhập:", value=datetime.now().strftime("%d/%m/%Y"))
                ex_customer = f2.text_input("Tên khách hàng / Địa chỉ:", value="Đại lý Hà Nội")

                f3, f4 = st.columns(2)
                ex_device = f3.text_input("Loại đầu thu:", value="HDV2")
                ex_old = f4.text_input("Mã dịch vụ cũ:", value="3912784990")

                f5, f6 = st.columns(2)
                ex_new = f5.text_input("Mã dịch vụ mới:", value="3922395880")
                ex_executor = f6.text_input("Người thực hiện:", value="Nguyễn Vĩnh Toàn")

                ex_return = st.number_input("Ngày trả (số lượng):", min_value=1, value=1)

                if st.form_submit_button("Lưu bản ghi Đổi bảo hành"):
                    new_ex = {
                        "STT": len(st.session_state.warranty_exchange_data) + 1,
                        "Ngày nhập": ex_date,
                        "Tên khách hàng / Địa chỉ": ex_customer,
                        "Loại đầu thu": ex_device,
                        "Mã dịch vụ cũ": ex_old,
                        "Mã dịch vụ mới": ex_new,
                        "Người thực hiện": ex_executor,
                        "Ngày trả (hoàn thành)": ex_return
                    }
                    st.session_state.warranty_exchange_data = pd.concat([st.session_state.warranty_exchange_data, pd.DataFrame([new_ex])], ignore_index=True)
                    save_shared_storage()
                    add_audit_log(user_role, "Thêm lượt đổi bảo hành", f"Mã cũ {ex_old} -> Mã mới {ex_new}")
                    st.success("✅ Đã thêm lượt đổi bảo hành mới!")
                    st.rerun()

    with tab_web:
        st.components.v1.iframe("http://baohanh.truyenhinhso.vn", height=600, scrolling=True)

# MENU 5: LƯU TRỮ VÀ TÀI LIỆU AI
# MENU 5: GIÁM SÁT LUỒNG KÊNH CATCHUP (HLS MONITOR)
elif menu == "5. Giám sát Luồng Kênh Catchup (HLS Monitor)":
    st.title("📺 Giám sát Luồng Kênh Truyền Hình (Falcon HLS Catchup Monitor)")
    
    col_c1, col_c2 = st.columns([3, 1])
    with col_c1:
        st.info("🔗 **Hệ thống giám sát luồng Catchup:** [https://falconhlsmonitor.vtcdigital.top/](https://falconhlsmonitor.vtcdigital.top/)")
    with col_c2:
        st.link_button("🚀 Mở Tab Giám Sát Mới", "https://falconhlsmonitor.vtcdigital.top/", type="primary", use_container_width=True)

    st.markdown("""
    <style>
    .monitor-frame {
        width: 100%;
        height: 800px;
        border: 2px solid #1F4E78;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    </style>
    <iframe src="https://falconhlsmonitor.vtcdigital.top/" class="monitor-frame" allowfullscreen></iframe>
    """, unsafe_allow_html=True)

# MENU 6: KHÔNG GIAN TRAO ĐỔI ZALO & VIBER (PHONE WORKSPACE)
elif menu == "6. Không gian Trao đổi Zalo & Viber (0913332569)":
    st.title("📱 Không Gian Làm Việc Trao Đổi Zalo & Viber trực tuyến")
    st.caption("Khoang làm việc tích hợp cho số điện thoại Hotline: **0913332569** (Mở ứng dụng Zalo Web & Viber trao đổi nhanh)")

    col_p1, col_p2 = st.columns([1, 2])
    
    with col_p1:
        st.markdown("### 📲 Thông Tin Kết Nối Hotline")
        st.success("🟢 **Số trực ca:** `0913332569`")
        
        st.markdown("#### 1. Đăng nhập Zalo Web trực tiếp:")
        st.write("Truy cập nhanh phiên bản Zalo Web để nhắn tin, nhận báo cáo và trao đổi nhóm:")
        st.link_button("🌐 Mở Zalo Web (chat.zalo.me)", "https://chat.zalo.me", type="primary", use_container_width=True)
        
        st.markdown("#### 2. Kích hoạt Viber Desktop / App:")
        st.link_button("🟣 Mở Chat Viber Hotline", "viber://chat?number=%2B84913332569", use_container_width=True)

        st.divider()
        st.markdown("#### 3. Quét mã QR kết nối nhanh:")
        st.image("https://api.qrserver.com/v1/create-qr-code/?size=220x220&data=https://zalo.me/0913332569", caption="Quét kết nối Zalo 0913332569", use_container_width=True)

    with col_p2:
        st.markdown("### 💬 Khung Trình Duyệt Trao Đổi Công Việc (Zalo Web)")
        st.info("💡 Bạn có thể đăng nhập Zalo Web bằng mã QR hoặc số điện thoại **0913332569** ngay tại khung bên dưới:")
        
        # Phone mockup container
        st.markdown("""
        <div style="border: 3px solid #2980b9; border-radius: 12px; overflow: hidden; background: #fff; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
            <div style="background: #2980b9; color: white; padding: 10px 15px; font-weight: bold; display: flex; justify-content: space-between;">
                <span>📱 ZALO WEB WORKSPACE - NOC VTC (0913332569)</span>
                <span>🔴 LIVE</span>
            </div>
            <iframe src="https://chat.zalo.me" style="width: 100%; height: 680px; border: none;"></iframe>
        </div>
        """, unsafe_allow_html=True)

# MENU 7: LƯU TRỮ VÀ TÀI LIỆU AI
elif menu == "7. Lưu trữ và Tài liệu AI":
    st.title("📂 Hệ thống Lưu trữ & Trợ lý Tra cứu Tài liệu AI")
    
    tab_store, tab_ai = st.tabs(["📁 Kho Lưu Trữ Ổ E:", "🤖 Trợ lý AI Tra Cứu (RAG)"])
    
    with tab_store:
        st.subheader("Cấu trúc Thư mục Lưu trữ Ổ đĩa E:")
        st.markdown("""
        * 📁 **`E:\\Luutru\\Báo cáo ca`**: Lưu trữ file Báo cáo ca A4 đã gửi.
        * 📁 **`E:\\Luutru\\Báo cáo bảo hành`**: Lưu trữ file Excel báo cáo bảo hành A4 đã gửi.
        * 📁 **`E:\\Luutru\\Tài liệu phòng máy`**: Sơ đồ đấu nối, quy trình vận hành thiết bị NOC.
        * 📁 **`E:\\Luutru\\Hồ sơ đấu thầu`**: Tài liệu HSMT, HSDT các gói thầu kỹ thuật.
        * 📁 **`E:\\Luutru\\Các dự án triển khai`**: Hồ sơ nghiệm thu, bản vẽ hạ tầng dự án.
        * 📁 **`E:\\Tài liệu AI\\`**: Cơ sở dữ liệu Vector lưu trữ tài liệu tra cứu AI.
        """)
        
        uploaded_files = st.file_uploader("Upload tài liệu bổ sung vào ổ E:\\Luutru:", accept_multiple_files=True)
        if uploaded_files:
            for file in uploaded_files:
                st.success(f"✅ Đã lưu file **{file.name}** vào hệ thống ổ E:\\Luutru!")
                add_audit_log(user_role, "Upload tài liệu lưu trữ", file.name)

    with tab_ai:
        st.subheader("🤖 Tra cứu Quy trình & Hồ sơ Kỹ thuật bằng AI (Đường dẫn E:\\Tài liệu AI\\)")
        query = st.text_input("Nhập từ khóa hoặc câu hỏi cần tra cứu:")
        if query:
            st.info(f"🔍 **AI Agent đang tìm kiếm trong `E:\\Tài liệu AI\\` cho câu hỏi: '{query}'**")

# MENU 6: NHẬT KÝ HOẠT ĐỘNG
elif menu == "8. Nhật ký Hoạt động":
    st.title("📝 Nhật ký Hoạt động Hệ thống (Audit Trail)")
    st.dataframe(st.session_state.audit_logs, use_container_width=True)