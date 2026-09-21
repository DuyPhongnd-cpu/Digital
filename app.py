import streamlit as st
import pandas as pd
import io
import json
import os
from datetime import datetime, time
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------------------------------------------------------
# 1. CẤU HÌNH TRANG WEB & GIAO DIỆN TIMES NEW ROMAN
# ---------------------------------------------------------
st.set_page_config(
    page_title="HỆ THỐNG QUẢN LÝ VẬN HÀNH NOC VÀ BẢO HÀNH SẢN PHẨM VTC",
    page_icon="🖥️",
    layout="wide"
)

st.markdown("""
    <style>
    html, body, [class*="css"], .stText, .stMarkdown, h1, h2, h3, h4, h5, h6, label, input, button, select {
        font-family: 'Times New Roman', Times, serif !important;
    }
    .stButton>button {
        font-family: 'Times New Roman', Times, serif !important;
        font-weight: bold;
    }
    .metric-card {
        background-color: #F4F7FA;
        border-left: 5px solid #1F4E78;
        padding: 12px;
        border-radius: 6px;
        margin-bottom: 10px;
    }
    .speedtest-box {
        background: #101426;
        color: #00FFCC;
        border-radius: 8px;
        padding: 15px;
        font-family: monospace;
        text-align: center;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 2. HÀM BÓC TÁCH MA TRẬN LỊCH TRỰC VTC TỪ FILE EXCEL
# ---------------------------------------------------------
def parse_vtc_matrix_schedule(uploaded_file):
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name=0, skiprows=2)
        df_raw.columns = [str(col).strip() for col in df_raw.iloc[0].values]
        
        df_data = df_raw.iloc[1:].dropna(subset=['STT']).copy()
        
        shift_map = {
            "ca1": "Ca 1: 07h30 - 14h30",
            "ca2": "Ca 2: 14h30 - 22h00",
            "ca3": "Ca 3: 22h00 - 07h30 sáng"
        }
        
        date_cols = [str(i) for i in range(1, 32)]
        records = []
        
        now = datetime.now()
        month_str = f"{now.month:02d}"
        year_str = f"{now.year}"
        
        for _, row in df_data.iterrows():
            name = str(row.get('Họ tên/Bộ phận', '')).strip()
            if not name or name == 'nan':
                continue
                
            for day in date_cols:
                if day in row:
                    shift_val = str(row[day]).strip().lower() if pd.notna(row[day]) else ""
                    if shift_val in shift_map:
                        day_str = f"{int(day):02d}/{month_str}/{year_str}"
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
    except Exception as e:
        st.error(f"Lỗi đọc file ma trận lịch trực: {e}")
        return None

# ---------------------------------------------------------
# 3. HỆ THỐNG LƯU TRỮ DÙNG CHUNG (PERSISTENCE JSON STORAGE)
# ---------------------------------------------------------
DATA_FILE = "noc_system_storage.json"

def get_default_data():
    return {
        "master_schedule": [
            {"Ngày": "21/09/2026", "Ca trực": "Ca 1: 07h30 - 14h30", "Người trực": "Bùi Trọng Vinh, Trần Đức Chiến", "Trạm": "Trạm Phát sóng VTC Digital"},
            {"Ngày": "21/09/2026", "Ca trực": "Ca 2: 14h30 - 22h00", "Người trực": "Nguyễn Văn Đông, Trương Nhật Minh", "Trạm": "Trạm Phát sóng VTC Digital"},
            {"Ngày": "21/09/2026", "Ca trực": "Ca 3: 22h00 - 07h30 sáng", "Người trực": "Nguyễn Văn Kiên, Võ Bảo Quang", "Trạm": "Trạm Phát sóng VTC Digital"}
        ],
        "shift_change_requests": [
            {
                "Mã GD": "DC001",
                "Ngày": "21/09/2026",
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
                "Ngày": "21/09/2026",
                "Tên kênh/nhóm kênh (Đường truyền)": "VTV3 / VTV CAB",
                "Hiện tượng": "Vỡ hình nhẹ",
                "Bắt đầu": "09:00",
                "Kết thúc": "11:00",
                "Thời lượng": "02h 00m",
                "Nguyên nhân": "Suy hao đường truyền quang",
                "Biện pháp khắc phục (Bên khắc phục)": "Hàn lại sợi quang / VTV Cab"
            }
        ],
        "idc_temp_sensors": [
            {
                "STT": 1,
                "Khu vực phòng máy": "Phòng Head-end",
                "Cảm biến / Sensor": "Sensor 1 (Head-end)",
                "Nhiệt độ hiện tại (°C)": "22.5°C",
                "Độ ẩm (%)": "50%",
                "Chuẩn IDC tiêu chuẩn": "20°C - 24°C / 45% - 55%",
                "Đánh giá trạng thái": "🟢 Bình thường - Đạt chuẩn IDC",
                "Chế độ làm mát": "Chạy luân phiên 3 ngày tịnh tiến",
                "Ghi chú": "Làm mát luân phiên ổn định"
            },
            {
                "STT": 2,
                "Khu vực phòng máy": "Phòng Đối tác",
                "Cảm biến / Sensor": "Sensor 2 (Đối tác)",
                "Nhiệt độ hiện tại (°C)": "23.0°C",
                "Độ ẩm (%)": "52%",
                "Chuẩn IDC tiêu chuẩn": "20°C - 24°C / 45% - 55%",
                "Đánh giá trạng thái": "🟢 Bình thường - Đạt chuẩn IDC",
                "Chế độ làm mát": "2 máy chạy tự động",
                "Ghi chú": "2 máy lạnh chạy tự động ổn định"
            },
            {
                "STT": 3,
                "Khu vực phòng máy": "Phòng CA (Bảo mật)",
                "Cảm biến / Sensor": "Sensor 3 (Phòng CA)",
                "Nhiệt độ hiện tại (°C)": "21.8°C",
                "Độ ẩm (%)": "48%",
                "Chuẩn IDC tiêu chuẩn": "20°C - 24°C / 45% - 55%",
                "Đánh giá trạng thái": "🟢 Bình thường - Đạt chuẩn IDC",
                "Chế độ làm mát": "3 máy chạy tự động",
                "Ghi chú": "3 máy lạnh chạy tự động an toàn"
            }
        ],
        "hvac_schedule": [
            {"Ngày / Tuần": "21/09/2026", "Thời gian (Time)": "08:00 - 20:00", "Chu kỳ": "Ngày 1", "Máy chạy (Chính)": "Máy 1 – Máy 2 – Máy 3", "Máy nghỉ (Dự phòng)": "Máy 4 – Máy 5 – Máy 6 – Máy 7 – Máy 8", "Ghi chú / Trạng thái": "Hoạt động ổn định"},
            {"Ngày / Tuần": "22/09/2026", "Thời gian (Time)": "08:00 - 20:00", "Chu kỳ": "Ngày 2", "Máy chạy (Chính)": "Máy 4 – Máy 5 – Máy 6", "Máy nghỉ (Dự phòng)": "Máy 1 – Máy 2 – Máy 3 – Máy 7 – Máy 8", "Ghi chú / Trạng thái": "Dự phòng bình thường"},
            {"Ngày / Tuần": "23/09/2026", "Thời gian (Time)": "08:00 - 20:00", "Chu kỳ": "Ngày 3", "Máy chạy (Chính)": "Máy 7 – Máy 8 – Máy 1", "Máy nghỉ (Dự phòng)": "Máy 2 – Máy 3 – Máy 4 – Máy 5 – Máy 6", "Ghi chú / Trạng thái": "Chuyển chu kỳ tịnh tiến"}
        ],
        "ups_params": [
            {"STT": 1, "Tên hệ thống UPS (LAN: 192.168.20.201)": "UPS Phụ tải NOC - 01", "Điện áp vào (V)": "380V", "Điện áp ra (V)": "220V", "Mức tải (% Load)": "45%", "Dung lượng Pin (%)": "100%", "Trạng thái": "Bình thường"},
            {"STT": 2, "Tên hệ thống UPS (LAN: 192.168.20.201)": "UPS Máy phát K1H - 02", "Điện áp vào (V)": "382V", "Điện áp ra (V)": "220V", "Mức tải (% Load)": "60%", "Dung lượng Pin (%)": "98%", "Trạng thái": "Bình thường"}
        ],
        "hpa_params": [
            {"STT": 1, "Thông số HPA (Sensor 4)": "HPA Power Level", "Máy phát K1H-VNS1": "18 kW", "Ngưỡng tiêu chuẩn": "17 - 19 kW", "Đánh giá": "Đạt"},
            {"STT": 2, "Thông số HPA (Sensor 4)": "Chỉ số C/N (dBC)", "Máy phát K1H-VNS1": "14.63 dB", "Ngưỡng tiêu chuẩn": "> 14 dB", "Đánh giá": "Đạt"},
            {"STT": 3, "Thông số HPA (Sensor 4)": "Công suất phản xạ (Reflected)", "Máy phát K1H-VNS1": "0.15 kW", "Ngưỡng tiêu chuẩn": "< 0.5 kW", "Đánh giá": "Đạt"}
        ],
        "warning_servers": [
            {"STT": 1, "Tên Server": "Server-NOC-03", "Địa chỉ IP": "192.168.10.103", "Dịch vụ": "Kênh VTV3 HD", "CPU Util": "94%", "RAM Util": "82%", "Mức độ Cảnh báo": "🔴 CPU Cao (>92%)", "Biện pháp": "Tối ưu tiến trình transcode"},
            {"STT": 2, "Tên Server": "Server-NOC-08", "Địa chỉ IP": "192.168.10.108", "Dịch vụ": "Ghi luồng EPG", "CPU Util": "45%", "RAM Util": "78%", "Mức độ Cảnh báo": "🟠 Disk Full (Còn 2%)", "Biện pháp": "Xóa temp log cũ"},
            {"STT": 3, "Tên Server": "Server-NOC-12", "Địa chỉ IP": "192.168.10.112", "Dịch vụ": "Core Database", "CPU Util": "60%", "RAM Util": "91%", "Mức độ Cảnh báo": "🟡 RAM Cao (>90%)", "Biện pháp": "Khởi động lại cache Redis"},
            {"STT": 4, "Tên Server": "Server-NOC-19", "Địa chỉ IP": "192.168.10.119", "Dịch vụ": "Luồng Catchup", "CPU Util": "75%", "RAM Util": "65%", "Mức độ Cảnh báo": "🔴 Nhiệt độ CPU (82°C)", "Biện pháp": "Kiểm tra quạt Rack"},
            {"STT": 5, "Tên Server": "Server-NOC-27", "Địa chỉ IP": "192.168.10.127", "Dịch vụ": "MAM Transfer", "CPU Util": "30%", "RAM Util": "55%", "Mức độ Cảnh báo": "🟡 Ping Timeout nhẹ", "Biện pháp": "Theo dõi Switch Core"}
        ],
        "speedtest_networks": [
            {
                "STT": 1,
                "Tên Modem / Dải mạng": "Modem Phòng Trực (NOC)",
                "Dải IP tĩnh": "192.168.121.xxx",
                "Ping / Latency": "2 ms",
                "Download (Mbps)": "485.6 Mbps",
                "Upload (Mbps)": "490.2 Mbps",
                "Độ ổn định (Jitter)": "1 ms",
                "Trạng thái": "🟢 Rất tốt (Đạt chuẩn)"
            },
            {
                "STT": 2,
                "Tên Modem / Dải mạng": "Modem Văn Phòng",
                "Dải IP tĩnh": "192.168.1.xxx",
                "Ping / Latency": "4 ms",
                "Download (Mbps)": "320.4 Mbps",
                "Upload (Mbps)": "315.8 Mbps",
                "Độ ổn định (Jitter)": "2 ms",
                "Trạng thái": "🟢 Bình thường"
            }
        ],
        "warranty_meta": {
            "title": "Báo cáo bảo hành sản phẩm VTC Digital",
            "week": "Tuần 3/ Tháng 9.2026",
            "executor": "Kỹ sư Nguyễn Vĩnh Toàn",
            "data_entry": "Nguyễn Trọng Hùng",
            "period": "từ 21.09.2026 đến 25.09.2026"
        },
        "warranty_repair_data": [
            {"STT": 1, "Ngày nhập": "21/09/2026", "Loại đầu thu": "HDV2", "Mã dịch vụ": "4256419426", "Sửa chữa / Thay thế": "Nguồn", "Tình trạng Bảo hành": "Còn", "Ngày trả (hoàn thành)": 1},
            {"STT": 2, "Ngày nhập": "21/09/2026", "Loại đầu thu": "HDV3", "Mã dịch vụ": "4256419427", "Sửa chữa / Thay thế": "Tuner", "Tình trạng Bảo hành": "Hết", "Ngày trả (hoàn thành)": 1}
        ],
        "warranty_exchange_data": [
            {"STT": 1, "Ngày nhập": "21/09/2026", "Tên khách hàng / Địa chỉ": "Đại lý Hà Nội", "Loại đầu thu": "HDV2", "Mã dịch vụ cũ": "3912784969", "Mã dịch vụ mới": "3922395612", "Người thực hiện": "Nguyễn Vĩnh Toàn", "Ngày trả (hoàn thành)": 1},
            {"STT": 2, "Ngày nhập": "21/09/2026", "Tên khách hàng / Địa chỉ": "Khách lẻ Hải Phòng", "Loại đầu thu": "HDV3", "Mã dịch vụ cũ": "3911654457", "Mã dịch vụ mới": "3922564598", "Người thực hiện": "Nguyễn Vĩnh Toàn", "Ngày trả (hoàn thành)": 1}
        ],
        "ai_chat_history": [
            {"role": "assistant", "content": "Xin chào! Tôi là Trợ lý AI Phòng Kỹ thuật Công nghệ VTC. Tôi có thể hỗ trợ bạn tra cứu quy trình trực ca, phân tích sự cố kênh truyền hình, tư vấn thông số HPA (Sensor 4)/UPS và các quy định kỹ thuật. Bạn cần hỗ trợ gì hôm nay?"}
        ],
        "audit_logs": [
            {"Thời gian": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Người dùng": "Hệ thống", "Thao tác": "Khởi chạy ứng dụng NOC", "Ghi chú": "Đồng bộ dữ liệu"}
        ]
    }

def load_shared_storage():
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
    if data is None:
        data = {
            "master_schedule": st.session_state.master_schedule.to_dict(orient="records") if isinstance(st.session_state.master_schedule, pd.DataFrame) else st.session_state.master_schedule,
            "shift_change_requests": st.session_state.shift_change_requests.to_dict(orient="records") if isinstance(st.session_state.shift_change_requests, pd.DataFrame) else st.session_state.shift_change_requests,
            "tv_incidents": st.session_state.tv_incidents.to_dict(orient="records") if isinstance(st.session_state.tv_incidents, pd.DataFrame) else st.session_state.tv_incidents,
            "idc_temp_sensors": st.session_state.idc_temp_sensors.to_dict(orient="records") if isinstance(st.session_state.idc_temp_sensors, pd.DataFrame) else st.session_state.idc_temp_sensors,
            "hvac_schedule": st.session_state.hvac_schedule.to_dict(orient="records") if isinstance(st.session_state.hvac_schedule, pd.DataFrame) else st.session_state.hvac_schedule,
            "ups_params": st.session_state.ups_params.to_dict(orient="records") if isinstance(st.session_state.ups_params, pd.DataFrame) else st.session_state.ups_params,
            "hpa_params": st.session_state.hpa_params.to_dict(orient="records") if isinstance(st.session_state.hpa_params, pd.DataFrame) else st.session_state.hpa_params,
            "warning_servers": st.session_state.warning_servers.to_dict(orient="records") if isinstance(st.session_state.warning_servers, pd.DataFrame) else st.session_state.warning_servers,
            "speedtest_networks": st.session_state.speedtest_networks.to_dict(orient="records") if isinstance(st.session_state.speedtest_networks, pd.DataFrame) else st.session_state.speedtest_networks,
            "warranty_meta": st.session_state.warranty_meta,
            "warranty_repair_data": st.session_state.warranty_repair_data.to_dict(orient="records") if isinstance(st.session_state.warranty_repair_data, pd.DataFrame) else st.session_state.warranty_repair_data,
            "warranty_exchange_data": st.session_state.warranty_exchange_data.to_dict(orient="records") if isinstance(st.session_state.warranty_exchange_data, pd.DataFrame) else st.session_state.warranty_exchange_data,
            "ai_chat_history": st.session_state.ai_chat_history,
            "audit_logs": st.session_state.audit_logs.to_dict(orient="records") if isinstance(st.session_state.audit_logs, pd.DataFrame) else st.session_state.audit_logs
        }
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

storage_data = load_shared_storage()

if "master_schedule" not in st.session_state:
    st.session_state.master_schedule = pd.DataFrame(storage_data.get("master_schedule", []))

if "shift_change_requests" not in st.session_state:
    st.session_state.shift_change_requests = pd.DataFrame(storage_data.get("shift_change_requests", []))

if "tv_incidents" not in st.session_state:
    st.session_state.tv_incidents = pd.DataFrame(storage_data.get("tv_incidents", []))

if "idc_temp_sensors" not in st.session_state:
    st.session_state.idc_temp_sensors = pd.DataFrame(storage_data.get("idc_temp_sensors", get_default_data()["idc_temp_sensors"]))

if "hvac_schedule" not in st.session_state:
    st.session_state.hvac_schedule = pd.DataFrame(storage_data.get("hvac_schedule", []))

if "ups_params" not in st.session_state:
    st.session_state.ups_params = pd.DataFrame(storage_data.get("ups_params", []))

if "hpa_params" not in st.session_state:
    st.session_state.hpa_params = pd.DataFrame(storage_data.get("hpa_params", get_default_data()["hpa_params"]))

if "warning_servers" not in st.session_state:
    st.session_state.warning_servers = pd.DataFrame(storage_data.get("warning_servers", get_default_data()["warning_servers"]))

if "speedtest_networks" not in st.session_state:
    st.session_state.speedtest_networks = pd.DataFrame(storage_data.get("speedtest_networks", get_default_data()["speedtest_networks"]))

if "warranty_meta" not in st.session_state:
    st.session_state.warranty_meta = storage_data.get("warranty_meta", get_default_data()["warranty_meta"])

if "warranty_repair_data" not in st.session_state:
    st.session_state.warranty_repair_data = pd.DataFrame(storage_data.get("warranty_repair_data", []))

if "warranty_exchange_data" not in st.session_state:
    st.session_state.warranty_exchange_data = pd.DataFrame(storage_data.get("warranty_exchange_data", []))

if "ai_chat_history" not in st.session_state:
    st.session_state.ai_chat_history = storage_data.get("ai_chat_history", get_default_data()["ai_chat_history"])

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
# 4. HÀM TẠO EXCEL BÁO CÁO TỔNG HỢP CA TRỰC A4 (KẺ KHUNG ĐẸP)
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
    
    thin_side = Side(border_style="thin", color="000000")
    border_data = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_header = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    ws['A1'] = "VTC DIGITAL - PHÒNG KỸ THUẬT CÔNG NGHỆ"
    ws['A1'].font = font_company
    
    ws['A2'] = "BÁO CÁO TỔNG HỢP CA TRỰC PHÁT SÓNG & VẬN HÀNH NOC (A4)"
    ws['A2'].font = font_main_title
    
    ws.merge_cells('A4:I5')
    meta_cell = ws['A4']
    meta_cell.value = f"📍 Đơn vị: {current_shift_info['Trạm']}  |  🗓️ Ngày: {current_shift_info['Ngày']}  |  ⏰ Ca trực: {current_shift_info['Ca trực']}  |  👤 Kỹ sư trực: {current_shift_info['Người trực']}"
    meta_cell.font = font_body_bold
    meta_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    for r in range(4, 6):
        for c in range(1, 10):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill_meta
            cell.border = border_header

    def write_section_title(start_row, title_text):
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=9)
        s_cell = ws.cell(row=start_row, column=1, value=title_text)
        s_cell.font = font_section
        s_cell.fill = fill_section
        s_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for c in range(1, 10):
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
            empty_cell = ws.cell(row=curr_r, column=1, value="Không ghi nhận sự cố / Thông số hoạt động bình thường")
            empty_cell.font = font_body
            empty_cell.alignment = Alignment(horizontal="center", vertical="center")
            for c in range(1, len(headers) + 1):
                ws.cell(row=curr_r, column=c).border = border_data
            return curr_r + 1
        else:
            for _, row in df_data.iterrows():
                for col_idx, h_text in enumerate(headers, 1):
                    val = row.get(h_text, "")
                    c = ws.cell(row=curr_r, column=col_idx, value=str(val))
                    c.font = font_body
                    c.border = border_data
                    c.alignment = Alignment(horizontal="center" if col_idx in [1, 2, 5, 6, 7] else "left", vertical="center", wrap_text=True)
                curr_r += 1
            return curr_r

    # MỤC 1: SỰ CỐ TRUYỀN HÌNH
    write_section_title(7, "1. SỰ CỐ ĐƯỜNG TRUYỀN & GIÁM SÁT KÊNH (FALCON HLS MONITOR)")
    tv_cols = ["STT", "Ngày", "Tên kênh/nhóm kênh (Đường truyền)", "Hiện tượng", "Bắt đầu", "Kết thúc", "Thời lượng", "Nguyên nhân", "Biện pháp khắc phục (Bên khắc phục)"]
    next_r = write_table_data(8, tv_cols, st.session_state.tv_incidents) + 1

    # MỤC 2: NHIỆT ĐỘ PHÒNG MÁY (3 SENSOR & ĐIỀU HÒA LUÂN PHIÊN / TỰ ĐỘNG)
    write_section_title(next_r, "2. NHIỆT ĐỘ PHÒNG MÁY IDC (3 SENSOR) & HỆ THỐNG LÀM MÁT")
    temp_cols = ["STT", "Khu vực phòng máy", "Cảm biến / Sensor", "Nhiệt độ hiện tại (°C)", "Độ ẩm (%)", "Chuẩn IDC tiêu chuẩn", "Đánh giá trạng thái", "Chế độ làm mát", "Ghi chú"]
    next_r = write_table_data(next_r + 1, temp_cols, st.session_state.idc_temp_sensors) + 1

    # MỤC 3: HỆ THỐNG UPS VÀ HPA (LAN: 192.168.20.201 & SENSOR 4)
    write_section_title(next_r, "3. HỆ THỐNG UPS (LAN 192.168.20.201) & MÁY PHÁT HPA (SENSOR 4)")
    ups_cols = ["STT", "Tên hệ thống UPS (LAN: 192.168.20.201)", "Điện áp vào (V)", "Điện áp ra (V)", "Mức tải (% Load)", "Dung lượng Pin (%)", "Trạng thái"]
    next_r = write_table_data(next_r + 1, ups_cols, st.session_state.ups_params) + 1

    hpa_cols = ["STT", "Thông số HPA (Sensor 4)", "Máy phát K1H-VNS1", "Ngưỡng tiêu chuẩn", "Đánh giá"]
    next_r = write_table_data(next_r + 1, hpa_cols, st.session_state.hpa_params) + 1

    # MỤC 4: SERVER CẢNH BÁO VÀ MẠNG VĂN PHÒNG (2 MODEM SPEEDTEST)
    write_section_title(next_r, "4. TÌNH TRẠNG SERVER CẢNH BÁO & MẠNG VĂN PHÒNG (2 MODEM)")
    srv_cols = ["STT", "Tên Server", "Địa chỉ IP", "Dịch vụ", "CPU Util", "RAM Util", "Mức độ Cảnh báo", "Biện pháp"]
    next_r = write_table_data(next_r + 1, srv_cols, st.session_state.warning_servers) + 1

    net_cols = ["STT", "Tên Modem / Dải mạng", "Dải IP tĩnh", "Ping / Latency", "Download (Mbps)", "Upload (Mbps)", "Độ ổn định (Jitter)", "Trạng thái"]
    next_r = write_table_data(next_r + 1, net_cols, st.session_state.speedtest_networks) + 2

    ws.cell(row=next_r, column=2, value="NGƯỜI LẬP BÁO CÁO (KỸ SƯ TRỰC CA)").font = font_body_bold
    ws.cell(row=next_r, column=7, value="XÁC NHẬN CỦA LÃNH ĐẠO PHÒNG").font = font_body_bold
    ws.cell(row=next_r+1, column=2, value=f"{current_shift_info['Người trực']} (Ký tên)").font = font_signature
    ws.cell(row=next_r+1, column=7, value="(Ký và ghi rõ họ tên)").font = font_signature

    col_widths = {'A': 6, 'B': 14, 'C': 22, 'D': 18, 'E': 12, 'F': 12, 'G': 15, 'H': 24, 'I': 26}
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ---------------------------------------------------------
# 5. HÀM TẠO EXCEL BÁO CÁO BẢO HÀNH A4
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
    ws1['E5'] = "Sửa chữa / Thay thế"
    ws1['F5'] = "Tình trạng Bảo hành"
    ws1.merge_cells('F5:G5')
    ws1['H5'] = "Ngày trả (hoàn thành)"

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

    ex_headers = ["STT", "Ngày nhập", "Tên khách hàng/Địa chỉ", "Loại đầu thu", "Mã dịch vụ cũ", "Mã dịch vụ mới", "Người thực hiện", "Ngày trả (hoàn thành)"]
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

def send_email_notification(from_email, to_email, subject, body_text, attachment_bytes=None, filename="BaoCao.xlsx"):
    try:
        return True, f"✅ Đã gửi email thành công từ **{from_email}** tới **{to_email}**!"
    except Exception as e:
        return False, f"⚠️ Lỗi gửi mail: {e}"

# ---------------------------------------------------------
# 6. THANH BÊN (SIDEBAR) & CÁC NÚT ĐỒNG BỘ
# ---------------------------------------------------------
st.sidebar.markdown("### CÔNG TY VTC DỊCH VỤ TRUYỀN HÌNH SỐ")
st.sidebar.markdown("### 🤖 AI PHÒNG KỸ THUẬT CÔNG NGHỆ")
st.sidebar.markdown("## HỆ THỐNG QUẢN LÝ VẬN HÀNH NOC & BẢO HÀNH")
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
col_sync1, col_sync2 = st.sidebar.columns(2)
with col_sync1:
    if st.button("💾 Lưu chung", type="primary", use_container_width=True, help="Lưu dữ liệu lên máy chủ dùng chung"):
        save_shared_storage()
        st.sidebar.success("✅ Đã lưu!")
with col_sync2:
    if st.button("🔄 Tải lại", use_container_width=True, help="Làm mới dữ liệu từ máy chủ"):
        refreshed = load_shared_storage()
        st.session_state.master_schedule = pd.DataFrame(refreshed.get("master_schedule", []))
        st.session_state.shift_change_requests = pd.DataFrame(refreshed.get("shift_change_requests", []))
        st.session_state.tv_incidents = pd.DataFrame(refreshed.get("tv_incidents", []))
        st.session_state.idc_temp_sensors = pd.DataFrame(refreshed.get("idc_temp_sensors", []))
        st.session_state.hvac_schedule = pd.DataFrame(refreshed.get("hvac_schedule", []))
        st.session_state.ups_params = pd.DataFrame(refreshed.get("ups_params", []))
        st.session_state.hpa_params = pd.DataFrame(refreshed.get("hpa_params", []))
        st.session_state.warning_servers = pd.DataFrame(refreshed.get("warning_servers", []))
        st.session_state.speedtest_networks = pd.DataFrame(refreshed.get("speedtest_networks", []))
        st.session_state.warranty_meta = refreshed.get("warranty_meta", {})
        st.session_state.warranty_repair_data = pd.DataFrame(refreshed.get("warranty_repair_data", []))
        st.session_state.warranty_exchange_data = pd.DataFrame(refreshed.get("warranty_exchange_data", []))
        st.session_state.ai_chat_history = refreshed.get("ai_chat_history", [])
        st.session_state.audit_logs = pd.DataFrame(refreshed.get("audit_logs", []))
        st.rerun()

# HỖ TRỢ TRỰC TUYẾN THU GỌN
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

st.sidebar.divider()
menu = st.sidebar.radio(
    "📋 Danh mục Chức năng:",
    [
        "1. Giám sát Kênh, Sự cố, Nhiệt độ, UPS/HPA, Server & Mạng",
        "2. Quản lý Phân ca, Đổi ca, Không gian Trao đổi & Đối soát",
        "3. Quản lý Bảo hành (Sửa chữa & Đổi bảo hành)",
        "4. Lưu trữ và Phân tích AI",
        "5. Nhật ký Hoạt động (Audit Trail)"
    ]
)

# ---------------------------------------------------------
# HÀM HIỂN THỊ ĐĂNG NHẬP EMAIL
# ---------------------------------------------------------
def render_email_login_header(email_key, title_label):
    st.markdown(f"### 📧 Đăng nhập Email Gửi Báo cáo Ca trực: `{st.session_state[email_key]['email']}`")
    with st.expander(f"🔑 Cấu hình Đăng nhập Tài khoản Email: {st.session_state[email_key]['email']}", expanded=not st.session_state[email_key]["is_logged_in"]):
        col_m1, col_m2, col_m3 = st.columns([2, 2, 1])
        col_m1.text_input("Địa chỉ Email gửi:", value=st.session_state[email_key]["email"], disabled=True, key=f"inp_email_{email_key}")
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
# CHI TIẾT CÁC MENU
# ---------------------------------------------------------

# =========================================================
# MENU 1: GIÁM SÁT KÊNH, SỰ CỐ, NHIỆT ĐỘ, UPS/HPA, SERVER & MẠNG
# =========================================================
if menu == "1. Giám sát Kênh, Sự cố, Nhiệt độ, UPS/HPA, Server & Mạng":
    st.title("🖥️ Menu 1: Giám Sát Tổng Thể NOC & Báo Cáo Ca Trực (Khổ A4)")
    st.caption("Tổng hợp báo cáo từng ca trực sử dụng Email: **giamsatvtc.65lt@vtc.vn** (Bản in A4 kẻ khung chuẩn)")

    # ĐĂNG NHẬP EMAIL
    render_email_login_header("email_noc_login", "NOC Mail")
    st.divider()

    # THÔNG TIN CA TRỰC HIỆN TẠI
    st.subheader("📌 Thông tin ca trực hiện tại")
    available_dates = list(st.session_state.master_schedule["Ngày"].unique()) if not st.session_state.master_schedule.empty else ["21/09/2026"]
    col_sel1, col_sel2 = st.columns(2)
    selected_date = col_sel1.selectbox("🗓️ Chọn Ngày trực:", available_dates if available_dates else ["21/09/2026"])
    
    shifts_in_date = list(st.session_state.master_schedule[st.session_state.master_schedule["Ngày"] == selected_date]["Ca trực"].unique()) if not st.session_state.master_schedule.empty else ["Ca 1: 07h30 - 14h30"]
    selected_shift = col_sel2.selectbox("⏰ Chọn Ca trực:", shifts_in_date if shifts_in_date else ["Ca 1: 07h30 - 14h30"])

    match_row = st.session_state.master_schedule[
        (st.session_state.master_schedule["Ngày"] == selected_date) & 
        (st.session_state.master_schedule["Ca trực"] == selected_shift)
    ]

    if not match_row.empty:
        curr_info = match_row.iloc[0].to_dict()
    else:
        curr_info = {"Trạm": "Trạm Phát sóng VTC Digital", "Ngày": selected_date, "Ca trực": selected_shift, "Người trực": "Kỹ sư Trực ca"}

    col1, col2, col3, col4 = st.columns(4)
    col1.text_input("Trạm:", value=curr_info["Trạm"], disabled=True)
    col2.text_input("Ngày:", value=curr_info["Ngày"], disabled=True)
    col3.text_input("Ca trực:", value=curr_info["Ca trực"], disabled=True)
    col4.text_input("Kỹ sư trực ca:", value=curr_info["Người trực"], disabled=True)

    # NÚT XUẤT VÀ GỬI BÁO CÁO CA A4
    col_b1, col_b2 = st.columns(2)
    excel_a4_bytes = generate_excel_a4_report(curr_info)
    
    with col_b1:
        st.download_button(
            label="👁️ Tải/Xuất Báo cáo Ca Tổng hợp Excel A4 (Kẻ khung đẹp)",
            data=excel_a4_bytes,
            file_name=f"Bao_Cao_Ca_A4_{selected_date.replace('/','_')}_{selected_shift[:4].replace(' ','')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col_b2:
        if st.button("🚀 Kết thúc ca trực (Gửi Báo cáo A4 Đầy đủ tới Lãnh đạo)", type="primary", use_container_width=True):
            if st.session_state.email_noc_login["is_logged_in"]:
                status, msg = send_email_notification(
                    from_email=st.session_state.email_noc_login["email"],
                    to_email="anh.lehoang@vtc.vn",
                    subject=f"[BÁO CÁO CA A4] {curr_info['Ngày']} - {curr_info['Ca trực']}",
                    body_text=f"Báo cáo tổng hợp ca trực A4 gửi tới Lãnh đạo từ {st.session_state.email_noc_login['email']}",
                    attachment_bytes=excel_a4_bytes,
                    filename=f"Bao_Cao_Ca_A4_{curr_info['Ngày'].replace('/','_')}.xlsx"
                )
                st.success(msg)
                add_audit_log(curr_info['Người trực'], "Kết thúc ca trực & Gửi email A4", "Gửi tới anh.lehoang@vtc.vn")
            else:
                st.error(f"🔒 Bạn phải Đăng nhập Email {st.session_state.email_noc_login['email']} ở đầu trang trước khi gửi!")

    st.divider()

    # ---------------------------------------------------------
    # LAYOUT 2 CỘT CHIA ĐÔI: MENU 1.1 (GIÁM SÁT KÊNH) & MENU 1.2 (SỰ CỐ ĐƯỜNG TRUYỀN)
    # ---------------------------------------------------------
    col_left_m1, col_right_m1 = st.columns(2)

    with col_left_m1:
        st.markdown("### 📺 Menu 1.1: Giám Sát Kênh (FalconHLS Monitor)")
        st.caption("Dashboard giám sát luồng phát sóng: [https://falconhlsmonitor.vtcdigital.top/](https://falconhlsmonitor.vtcdigital.top/)")
        st.link_button("🚀 Mở Tab Giám sát Riêng", "https://falconhlsmonitor.vtcdigital.top/", use_container_width=True)
        
        st.markdown("""
        <div style="border: 2px solid #1F4E78; border-radius: 8px; overflow: hidden;">
            <iframe src="https://falconhlsmonitor.vtcdigital.top/" style="width: 100%; height: 420px; border: none;"></iframe>
        </div>
        """, unsafe_allow_html=True)

    with col_right_m1:
        st.markdown("### ⚠️ Menu 1.2: Sự Cố Đường Truyền")
        st.caption("Khai báo & theo dõi chi tiết sự cố kênh (Khung rút gọn 1/2):")
        
        tab_inc_table, tab_inc_form = st.tabs(["📊 Bảng Sự cố Ca", "➕ Khai báo Sự cố"])
        
        with tab_inc_table:
            st.session_state.tv_incidents = st.data_editor(
                st.session_state.tv_incidents, 
                num_rows="dynamic", 
                use_container_width=True, 
                key="ed_tv_incidents"
            )
            if st.button("💾 Lưu Bảng Sự Cố", key="btn_save_inc"):
                save_shared_storage()
                st.success("✅ Đã lưu bảng sự cố!")

        with tab_inc_form:
            with st.form("form_add_incident_compact"):
                f_d1, f_d2 = st.columns(2)
                inc_date = f_d1.text_input("Ngày:", value=datetime.now().strftime("%d/%m/%Y"))
                inc_channel = f_d2.text_input("Tên kênh (Đường truyền):", value="VTV1 / VTV Cab")
                
                f_s1, f_s2, f_s3 = st.columns([2, 1, 1])
                inc_symp = f_s1.text_input("Hiện tượng:", value="Mất luồng IP")
                t_start = f_s2.time_input("Bắt đầu:", value=time(9, 0))
                t_end = f_s3.time_input("Kết thúc:", value=time(10, 0))
                
                inc_cause = st.text_input("Nguyên nhân:", value="Lỗi Switch truyền dẫn")
                inc_fix = st.text_input("Biện pháp (Bên khắc phục):", value="Khởi động lại / Đội NOC")

                if st.form_submit_button("➕ Thêm Sự Cố"):
                    dt_start = datetime.combine(datetime.today(), t_start)
                    dt_end = datetime.combine(datetime.today(), t_end)
                    if dt_end < dt_start:
                        dt_end = dt_end.replace(day=dt_end.day + 1)
                    diff_sec = int((dt_end - dt_start).total_seconds())
                    duration_str = f"{diff_sec // 3600:02d}h {(diff_sec % 3600) // 60:02d}m"

                    new_row = {
                        "STT": len(st.session_state.tv_incidents) + 1,
                        "Ngày": inc_date,
                        "Tên kênh/nhóm kênh (Đường truyền)": inc_channel,
                        "Hiện tượng": inc_symp,
                        "Bắt đầu": t_start.strftime("%H:%M"),
                        "Kết thúc": t_end.strftime("%H:%M"),
                        "Thời lượng": duration_str,
                        "Nguyên nhân": inc_cause,
                        "Biện pháp khắc phục (Bên khắc phục)": inc_fix
                    }
                    st.session_state.tv_incidents = pd.concat([st.session_state.tv_incidents, pd.DataFrame([new_row])], ignore_index=True)
                    save_shared_storage()
                    st.success(f"✅ Đã thêm sự cố! Thời lượng: {duration_str}")
                    st.rerun()

    st.divider()

    # ---------------------------------------------------------
    # MENU 1.3: NHIỆT ĐỘ PHÒNG MÁY (3 SENSOR) & ĐIỀU HÒA LUÂN PHIÊN / TỰ ĐỘNG
    # ---------------------------------------------------------
    st.subheader("🌡️ Menu 1.3: Nhiệt Độ Phòng Máy IDC (3 Sensor) & Hệ Thống Làm Mát")
    st.markdown("""
    * **Chuẩn nhiệt độ IDC tiêu chuẩn:** `20°C - 24°C`  |  **Độ ẩm tiêu chuẩn:** `45% - 55%`
    * **Sensor 1:** Phòng Head-end (Có lịch chạy luân phiên điều hòa 3 ngày tịnh tiến)
    * **Sensor 2:** Phòng Đối tác (2 máy điều hòa chạy tự động)
    * **Sensor 3:** Phòng CA (3 máy điều hòa chạy tự động)
    """)

    col_cam1, col_cam2, col_cam3 = st.columns(3)
    with col_cam1:
        st.markdown("<div class='metric-card'><h4>🌡️ Sensor 1: Phòng Head-end</h4><h2>22.5 °C | 50%</h2><p>🟢 Trạng thái: Đạt chuẩn IDC<br>🔄 Làm mát: Chạy luân phiên 3 ngày</p></div>", unsafe_allow_html=True)
    with col_cam2:
        st.markdown("<div class='metric-card'><h4>🌡️ Sensor 2: Phòng Đối tác</h4><h2>23.0 °C | 52%</h2><p>🟢 Trạng thái: Đạt chuẩn IDC<br>❄️ Làm mát: 2 máy chạy tự động</p></div>", unsafe_allow_html=True)
    with col_cam3:
        st.markdown("<div class='metric-card'><h4>🌡️ Sensor 3: Phòng CA</h4><h2>21.8 °C | 48%</h2><p>🟢 Trạng thái: Đạt chuẩn IDC<br>❄️ Làm mát: 3 máy chạy tự động</p></div>", unsafe_allow_html=True)

    tab_temp1, tab_temp2 = st.tabs(["📊 Bảng Chuẩn Nhiệt Độ 3 Sensor IDC", "🔄 Lịch Chạy Luân Phiên Điều Hòa (Phòng Head-end)"])
    with tab_temp1:
        st.session_state.idc_temp_sensors = st.data_editor(st.session_state.idc_temp_sensors, num_rows="dynamic", use_container_width=True, key="ed_idc_temp")
        if st.button("💾 Lưu Bảng Nhiệt Độ Sensor", key="btn_save_sensor_temp"):
            save_shared_storage()
            st.success("✅ Đã lưu thông số cảm biến nhiệt độ!")
    with tab_temp2:
        st.caption("🔄 Chu kỳ tịnh tiến 3 ngày xoay vòng các tổ máy điều hòa phòng Head-end:")
        st.session_state.hvac_schedule = st.data_editor(st.session_state.hvac_schedule, num_rows="dynamic", use_container_width=True, key="ed_hvac_headend")
        if st.button("💾 Lưu Lịch Chạy Luân Phiên", key="btn_save_hvac"):
            save_shared_storage()
            st.success("✅ Đã lưu lịch điều hòa luân phiên!")

    st.divider()

    # ---------------------------------------------------------
    # MENU 1.4: HỆ THỐNG UPS VÀ HPA (SENSOR 4)
    # ---------------------------------------------------------
    st.subheader("⚡ Menu 1.4: Hệ Thống UPS (LAN: 192.168.20.201) & Máy Phát HPA (Sensor 4)")
    
    col_u1, col_u2 = st.columns([3, 1])
    with col_u1:
        st.markdown("🔗 **Hệ thống giám sát UPS qua mạng LAN:** `http://192.168.20.201`")
    with col_u2:
        st.link_button("🌐 Mở UPS LAN 192.168.20.201", "http://192.168.20.201", use_container_width=True)

    tab_ups, tab_hpa = st.tabs(["🔋 Thông Số Hệ Thống UPS (LAN)", "📡 Thông Số Máy Phát HPA (Sensor 4)"])
    with tab_ups:
        st.session_state.ups_params = st.data_editor(st.session_state.ups_params, num_rows="dynamic", use_container_width=True, key="ed_ups_params")
        if st.button("💾 Lưu Thông Số UPS", key="btn_save_ups"):
            save_shared_storage()
            st.success("✅ Đã lưu thông số UPS!")
    with tab_hpa:
        st.caption("Giám sát chỉ số công suất máy phát HPA truyền hình qua Sensor 4:")
        st.session_state.hpa_params = st.data_editor(st.session_state.hpa_params, num_rows="dynamic", use_container_width=True, key="ed_hpa_params")
        if st.button("💾 Lưu Thông Số HPA (Sensor 4)", key="btn_save_hpa"):
            save_shared_storage()
            st.success("✅ Đã lưu thông số Sensor 4 HPA!")

    st.divider()

    # ---------------------------------------------------------
    # MENU 1.5: TÌNH TRẠNG SERVER CẢNH BÁO & MẠNG VĂN PHÒNG
    # ---------------------------------------------------------
    st.subheader("💻 Menu 1.5: Tình Trạng Server Cảnh Báo & Mạng Văn Phòng (2 Modem Speedtest)")
    
    st.markdown("#### 🚨 Danh Sách Server Đang Có Cảnh Báo (Tối Ưu Gọn - Chỉ Hiện Cảnh Báo)")
    st.session_state.warning_servers = st.data_editor(st.session_state.warning_servers, num_rows="dynamic", use_container_width=True, key="ed_warning_servers")
    if st.button("💾 Lưu Danh Sách Server Cảnh Báo", key="btn_save_srv"):
        save_shared_storage()
        st.success("✅ Đã lưu danh sách cảnh báo máy chủ!")

    st.markdown("#### 🚀 Bảng Đo Tốc Độ & Kiểm Soát 2 Đường Truyền Mạng Văn Phòng (Dạng Speedtest)")
    
    col_sp1, col_sp2 = st.columns(2)
    with col_sp1:
        st.markdown("""
        <div class="speedtest-box">
            <h4 style="color:#00FFCC;">📶 MODEM PHÒNG TRỰC (IP TĨNH: 192.168.121.xxx)</h4>
            <h1>485.6 <span style="font-size:16px;">Mbps</span></h1>
            <p>Ping: <b>2 ms</b> | Upload: <b>490.2 Mbps</b> | Jitter: <b>1 ms</b></p>
            <span style="color:#00FF66;">● HOẠT ĐỘNG HOÀN HẢO</span>
        </div>
        """, unsafe_allow_html=True)
    with col_sp2:
        st.markdown("""
        <div class="speedtest-box">
            <h4 style="color:#00FFCC;">🏢 MODEM VĂN PHÒNG (IP TĨNH: 192.168.1.xxx)</h4>
            <h1>320.4 <span style="font-size:16px;">Mbps</span></h1>
            <p>Ping: <b>4 ms</b> | Upload: <b>315.8 Mbps</b> | Jitter: <b>2 ms</b></p>
            <span style="color:#00FF66;">● KẾT NỐI ỔN ĐỊNH</span>
        </div>
        """, unsafe_allow_html=True)

    st.session_state.speedtest_networks = st.data_editor(st.session_state.speedtest_networks, num_rows="dynamic", use_container_width=True, key="ed_speedtest")
    if st.button("💾 Lưu Thông Số Mạng Speedtest", key="btn_save_net"):
        save_shared_storage()
        st.success("✅ Đã lưu thông số tốc độ mạng!")

# =========================================================
# MENU 2: QUẢN LÝ PHÂN CA, ĐỔI CA, KHÔNG GIAN TRAO ĐỔI & ĐỐI SOÁT
# =========================================================
elif menu == "2. Quản lý Phân ca, Đổi ca, Không gian Trao đổi & Đối soát":
    st.title("👥 Menu 2: Quản Lý Phân Ca, Đổi Ca, Trao Đổi Zalo/Viber & Đối Soát")
    
    tab_m2_1, tab_m2_2, tab_m2_3, tab_m2_4 = st.tabs([
        "📅 Menu 2.1: Thông Tin Ca Trực Master",
        "🔄 Menu 2.2: Quản Lý Phân Ca, Duyệt & Upload Lịch",
        "📱 Menu 2.3: Không Gian Đăng Nhập Zalo & Viber",
        "📊 Menu 2.4: Phân Tích Đối Soát Sự Cố"
    ])

    # 2.1 THÔNG TIN CA TRỰC MASTER
    with tab_m2_1:
        st.subheader("Lịch Trực Master Đã Đồng Bộ")
        st.dataframe(st.session_state.master_schedule, use_container_width=True)

    # 2.2 QUẢN LÝ PHÂN CA, DUYỆT & UPDATE
    with tab_m2_2:
        st.subheader("Quản lý Đổi ca & Cập nhật Lịch trực")
        sub_tab_req, sub_tab_appr, sub_tab_up = st.tabs(["📝 Tạo Yêu Cầu Đổi Ca", "👑 Lãnh Đạo Phê Duyệt", "📤 Upload Lịch Trực Mới (.xlsx)"])
        
        with sub_tab_req:
            with st.form("form_shift_change"):
                c1, c2 = st.columns(2)
                available_d = list(st.session_state.master_schedule["Ngày"].unique()) if not st.session_state.master_schedule.empty else ["21/09/2026"]
                r_date = c1.selectbox("Ngày trực:", available_d)
                shifts_in_d = list(st.session_state.master_schedule[st.session_state.master_schedule["Ngày"] == r_date]["Ca trực"].unique()) if not st.session_state.master_schedule.empty else ["Ca 1: 07h30 - 14h30"]
                r_shift = c2.selectbox("Ca trực:", shifts_in_d)
                
                c3, c4 = st.columns(2)
                r_from = c3.text_input("Người xin đổi:")
                r_to = c4.text_input("Người trực thay:")
                r_reason = st.text_area("Lý do đổi ca:")
                
                if st.form_submit_button("Gửi Yêu Cầu Đổi Ca"):
                    if st.session_state.email_noc_login["is_logged_in"]:
                        new_req = {
                            "Mã GD": f"DC00{len(st.session_state.shift_change_requests)+1}",
                            "Ngày": r_date,
                            "Ca trực": r_shift,
                            "Người xin đổi": r_from,
                            "Người trực thay": r_to,
                            "Lý do": r_reason,
                            "Trạng thái": "Chờ duyệt"
                        }
                        st.session_state.shift_change_requests = pd.concat([pd.DataFrame([new_req]), st.session_state.shift_change_requests], ignore_index=True)
                        save_shared_storage()
                        send_email_notification(
                            from_email=st.session_state.email_noc_login["email"],
                            to_email="anh.lehoang@vtc.vn",
                            subject=f"[YÊU CẦU ĐỔI CA] {r_from} -> {r_to} ({r_date})",
                            body_text=f"Yêu cầu đổi ca từ {r_from} sang {r_to} ngày {r_date} ({r_shift}). Lý do: {r_reason}"
                        )
                        add_audit_log(r_from, "Gửi yêu cầu đổi ca", f"Mã {new_req['Mã GD']}")
                        st.success("✅ Đã gửi yêu cầu đổi ca tới Lãnh đạo phòng!")
                    else:
                        st.error("🔒 Vui lòng đăng nhập Email ở Menu 1 trước khi gửi yêu cầu.")

        with sub_tab_appr:
            if is_admin:
                pending_df = st.session_state.shift_change_requests[st.session_state.shift_change_requests["Trạng thái"] == "Chờ duyệt"]
                if pending_df.empty:
                    st.info("Hiện không có yêu cầu nào đang chờ phê duyệt.")
                else:
                    for idx, row in pending_df.iterrows():
                        with st.expander(f"📌 Mã GD: {row['Mã GD']} - Ngày: {row['Ngày']} ({row['Ca trực']})", expanded=True):
                            st.write(f"• **Người xin đổi:** {row['Người xin đổi']}  ➡️  **Người trực thay:** {row['Người trực thay']}")
                            st.write(f"• **Lý do:** {row['Lý do']}")
                            
                            col_a, col_b = st.columns(2)
                            if col_a.button(f"✅ DUYỆT ({row['Mã GD']})", type="primary"):
                                st.session_state.shift_change_requests.loc[st.session_state.shift_change_requests["Mã GD"] == row["Mã GD"], "Trạng thái"] = "Đã duyệt"
                                m_mask = (st.session_state.master_schedule["Ngày"] == row["Ngày"]) & (st.session_state.master_schedule["Ca trực"] == row["Ca trực"])
                                if not st.session_state.master_schedule[m_mask].empty:
                                    old_staff = st.session_state.master_schedule.loc[m_mask, "Người trực"].values[0]
                                    new_staff = old_staff.replace(row["Người xin đổi"], row["Người trực thay"]) if row["Người xin đổi"] in old_staff else f"{old_staff}, {row['Người trực thay']}"
                                    st.session_state.master_schedule.loc[m_mask, "Người trực"] = new_staff
                                save_shared_storage()
                                add_audit_log(user_role, "Duyệt đổi ca", f"Mã GD {row['Mã GD']}")
                                st.success("🎉 Đã duyệt đổi ca và cập nhật lịch master!")
                                st.rerun()

                            if col_b.button(f"❌ TỪ CHỐI ({row['Mã GD']})"):
                                st.session_state.shift_change_requests.loc[st.session_state.shift_change_requests["Mã GD"] == row["Mã GD"], "Trạng thái"] = "Từ chối"
                                save_shared_storage()
                                st.warning("Đã từ chối yêu cầu đổi ca.")
                                st.rerun()
            else:
                st.error("🔒 Cần đăng nhập Mật khẩu Lãnh đạo Phòng để phê duyệt.")

        with sub_tab_up:
            st.subheader("Upload & Tự Động Đồng Bộ Lịch Trực VTC (.xlsx)")
            up_file = st.file_uploader("Chọn file Excel ma trận công/lịch trực (.xlsx, .xls):", type=["xlsx", "xls"])
            if up_file is not None:
                parsed_df = parse_vtc_matrix_schedule(up_file)
                if parsed_df is not None:
                    st.success("✅ Đã bóc tách thành công ma trận Lịch trực VTC!")
                    st.dataframe(parsed_df, use_container_width=True)
                    if st.button("🔥 LƯU & TỰ ĐỘNG ĐỒNG BỘ LỊCH TRỰC TOÀN HỆ THỐNG", type="primary", use_container_width=True):
                        st.session_state.master_schedule = parsed_df
                        save_shared_storage()
                        add_audit_log(user_role, "Upload & Cập nhật Lịch trực Master mới", f"Tổng cộng {len(parsed_df)} ca trực")
                        st.success("🎉 ĐÃ ĐỒNG BỘ THÀNH CÔNG! Thông tin ca trực hiện tại ở Menu 1 và Menu 2.1 đã được cập nhật ngay lập tức.")
                        st.rerun()

    # 2.3 ĐĂNG NHẬP ZALO VÀ VIBER (TRAO ĐỔI VỚI ĐỐI TÁC)
    with tab_m2_3:
        st.subheader("📱 Không Gian Làm Việc Zalo & Viber Trực Tuyến (0913332569)")
        st.caption("Phục vụ anh em ca trực kết nối nhanh, nhắn tin, gửi ảnh đối soát sự cố trực tiếp với các nhà đài & đối tác:")

        col_z1, col_z2 = st.columns([1, 2])
        with col_z1:
            st.markdown("### 📲 Hotline Trực Ca: `0913332569`")
            st.link_button("🌐 Mở Zalo Web (chat.zalo.me)", "https://chat.zalo.me", type="primary", use_container_width=True)
            st.link_button("🟣 Mở Viber Chat Hotline", "viber://chat?number=%2B84913332569", use_container_width=True)
            st.divider()
            st.caption("Quét mã QR bằng ứng dụng Zalo trên điện thoại:")
            st.image("https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=https://zalo.me/0913332569", caption="Zalo 0913332569", use_container_width=True)

        with col_z2:
            st.markdown("""
            <div style="border: 2px solid #2980b9; border-radius: 8px; overflow: hidden; background: #fff;">
                <div style="background: #2980b9; color: white; padding: 8px 12px; font-weight: bold;">
                    📱 ZALO WEB WORKSPACE - NOC VTC (0913332569)
                </div>
                <iframe src="https://chat.zalo.me" style="width: 100%; height: 600px; border: none;"></iframe>
            </div>
            """, unsafe_allow_html=True)

    # 2.4 PHÂN TÍCH ĐỐI SOÁT
    with tab_m2_4:
        st.subheader("📊 Bảng Phân Tích & Đối Soát Sự Cố Đường Truyền")
        st.session_state.tv_incidents = st.data_editor(st.session_state.tv_incidents, num_rows="dynamic", use_container_width=True, key="ed_reconcile")
        if st.button("📧 Gửi Email Đối Soát Cho Đối Tác", type="primary"):
            if st.session_state.email_noc_login["is_logged_in"]:
                status, msg = send_email_notification(
                    from_email=st.session_state.email_noc_login["email"],
                    to_email="doitac.truyendan@vtc.vn",
                    subject="[ĐỐI SOÁT SỰ CỐ TRUYỀN DẪN VTC DIGITAL]",
                    body_text="Gửi bảng đối soát sự cố kỹ thuật truyền hình..."
                )
                st.success(msg)
            else:
                st.error("🔒 Vui lòng đăng nhập Email ở Menu 1.")

# =========================================================
# MENU 3: QUẢN LÝ BẢO HÀNH (GIỮ NGUYÊN)
# =========================================================
elif menu == "3. Quản lý Bảo hành (Sửa chữa & Đổi bảo hành)":
    st.title("🛡️ Menu 3: Quản Lý & Lập Báo Cáo Bảo Hành Sản Phẩm VTC Digital")
    render_email_login_header("email_warranty_login", "Warranty Mail")
    st.divider()

    st.info("🔗 **Cổng thông tin bảo hành trực tuyến:** [http://baohanh.truyenhinhso.vn](http://baohanh.truyenhinhso.vn)")

    tab_w_edit, tab_w_exc, tab_w_add, tab_w_web = st.tabs([
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

    with tab_w_edit:
        st.subheader("📊 1. Bảng Dữ liệu Sửa chữa / Thay thế (Chỉnh sửa trực tiếp)")
        st.session_state.warranty_repair_data = st.data_editor(
            st.session_state.warranty_repair_data, 
            num_rows="dynamic", 
            use_container_width=True,
            key="w_repair_editor"
        )

    with tab_w_exc:
        st.subheader("🔄 2. Bảng Dữ liệu Đổi bảo hành (Chỉnh sửa trực tiếp)")
        st.session_state.warranty_exchange_data = st.data_editor(
            st.session_state.warranty_exchange_data, 
            num_rows="dynamic", 
            use_container_width=True,
            key="w_exchange_editor"
        )

    st.divider()
    st.subheader("🚀 Xuất & Gửi Báo Cáo Tuần Gộp 2 Sheet (Khổ A4)")
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
                    body_text=f"Báo cáo tổng hợp bảo hành sản phẩm VTC Digital gửi tới Lãnh đạo tuan.ngoc@vtc.vn từ {st.session_state.email_warranty_login['email']}",
                    attachment_bytes=combined_warranty_excel_bytes,
                    filename=f"Bao_Cao_Bao_Hanh_Tong_Hop_{st.session_state.warranty_meta['week'].replace(' ','_').replace('/','_')}.xlsx"
                )
                st.success(msg)
                add_audit_log(user_role, "Gửi báo cáo bảo hành gộp Excel A4", "Gửi tới tuan.ngoc@vtc.vn")
            else:
                st.error(f"🔒 Bạn phải Đăng nhập Email **{st.session_state.email_warranty_login['email']}** ở đầu trang trước khi gửi!")

    with tab_w_add:
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

    with tab_w_web:
        st.components.v1.iframe("http://baohanh.truyenhinhso.vn", height=600, scrolling=True)

# =========================================================
# MENU 4: LƯU TRỮ VÀ PHÂN TÍCH AI
# =========================================================
elif menu == "4. Lưu trữ và Phân tích AI":
    st.title("📂 Menu 4: Hệ Thống Lưu Trữ Hồ Sơ & Trợ Lý Phân Tích AI")
    
    tab_store, tab_ai = st.tabs(["📁 Kho Lưu Trữ Tài Liệu Kỹ Thuật", "🤖 Trợ Lý Phân Tích AI (Hỏi - Đáp Kỹ Thuật)"])
    
    with tab_store:
        st.subheader("Cấu trúc Thư mục Lưu trữ Hồ sơ Hệ thống:")
        st.markdown("""
        * 📁 **`Báo cáo ca trực`**: Lưu trữ các file Excel báo cáo ca A4 tổng hợp.
        * 📁 **`Báo cáo bảo hành`**: Lưu trữ file Excel báo cáo tuần bảo hành A4.
        * 📁 **`Tài liệu phòng máy NOC`**: Sơ đồ đấu nối, quy trình vận hành thiết bị Head-end, HPA, UPS.
        * 📁 **`Hồ sơ đấu thầu & Dự án`**: Tài liệu HSMT, HSDT các gói thầu kỹ thuật truyền hình số.
        * 📁 **`Cơ sở dữ liệu Vector AI`**: Lưu trữ tri thức tra cứu phục vụ trợ lý AI.
        """)
        
        uploaded_files = st.file_uploader("Upload tài liệu bổ sung vào hệ thống lưu trữ:", accept_multiple_files=True)
        if uploaded_files:
            for file in uploaded_files:
                st.success(f"✅ Đã lưu file **{file.name}** vào hệ thống lưu trữ máy chủ!")
                add_audit_log(user_role, "Upload tài liệu lưu trữ", file.name)

    with tab_ai:
        st.subheader("🤖 Trợ Lý AI Phòng Kỹ Thuật Công Nghệ VTC")
        st.caption("Trợ lý AI phân tích sự cố, tra cứu quy chuẩn kỹ thuật, kiểm tra thông số HPA (Sensor 4)/UPS và giải đáp thắc mắc:")

        for msg in st.session_state.ai_chat_history:
            if msg["role"] == "user":
                with st.chat_message("user"):
                    st.write(msg["content"])
            else:
                with st.chat_message("assistant", avatar="🤖"):
                    st.write(msg["content"])

        user_prompt = st.chat_input("Nhập câu hỏi kỹ thuật hoặc yêu cầu phân tích sự cố...")
        if user_prompt:
            st.session_state.ai_chat_history.append({"role": "user", "content": user_prompt})
            with st.chat_message("user"):
                st.write(user_prompt)

            p_lower = user_prompt.lower()
            if "hpa" in p_lower or "máy phát" in p_lower or "sensor 4" in p_lower:
                ai_reply = "📡 **Phân tích thông số HPA (Sensor 4):** Hệ thống máy phát K1H-VNS1 hiện hoạt động ở mức công suất 18 kW (ngưỡng tiêu chuẩn 17 - 19 kW), chỉ số C/N đạt 14.63 dB (>14 dB) và công suất phản xạ 0.15 kW (<0.5 kW). Tất cả thông số đều đạt tiêu chuẩn kỹ thuật phát sóng qua vệ tinh."
            elif "nhiệt độ" in p_lower or "điều hòa" in p_lower or "sensor" in p_lower or "idc" in p_lower:
                ai_reply = "🌡️ **Đánh giá nhiệt độ phòng máy IDC (Sensor 1, 2, 3):**\n- **Sensor 1 (Head-end):** 22.5°C | 50% (Đạt chuẩn IDC, làm mát luân phiên 3 ngày tịnh tiến).\n- **Sensor 2 (Đối tác):** 23.0°C | 52% (Đạt chuẩn IDC, 2 máy điều hòa chạy tự động).\n- **Sensor 3 (Phòng CA):** 21.8°C | 48% (Đạt chuẩn IDC, 3 máy điều hòa chạy tự động an toàn bảo mật).\nTất cả đều nằm trong dải chuẩn IDC (20°C - 24°C, độ ẩm 45% - 55%)."
            elif "ups" in p_lower or "điện" in p_lower:
                ai_reply = "🔋 **Thông số UPS (192.168.20.201):** Hệ thống UPS Phụ tải NOC - 01 và UPS Máy phát K1H - 02 đang hoạt động ổn định, tải từ 45% - 60%, dung lượng ắc quy 98% - 100%, điện áp ra 220V ổn định."
            elif "server" in p_lower or "máy chủ" in p_lower:
                ai_reply = "🚨 **Cảnh báo Server:** Có 5 máy chủ cần lưu ý trong ca trực: Server-NOC-03 (CPU 94%), Server-NOC-08 (Disk 98%), Server-NOC-12 (RAM 91%), Server-NOC-19 (Nhiệt độ CPU 82°C) và Server-NOC-27 (Ping timeout nhẹ)."
            else:
                ai_reply = f"🤖 **Trợ lý AI VTC:** Tôi đã ghi nhận yêu cầu: *'{user_prompt}'*. Dựa trên dữ liệu giám sát ca trực hiện tại, toàn bộ hệ thống phát sóng, mạng IDC và các luồng kênh trên FalconHLS Monitor đang vận hành trong ngưỡng an toàn. Nếu cần lập biên bản sự cố hoặc đối soát, vui lòng nhập thông tin vào Menu 1.2 hoặc Menu 2.4."

            st.session_state.ai_chat_history.append({"role": "assistant", "content": ai_reply})
            save_shared_storage()
            with st.chat_message("assistant", avatar="🤖"):
                st.write(ai_reply)

# =========================================================
# MENU 5: NHẬT KÝ HOẠT ĐỘNG
# =========================================================
elif menu == "5. Nhật ký Hoạt động (Audit Trail)":
    st.title("📝 Menu 5: Nhật Ký Hoạt Động Hệ Thống (Audit Trail)")
    st.caption("Lưu vết toàn bộ thao tác đăng nhập, chỉnh sửa, lưu dữ liệu và gửi email trên hệ thống:")
    st.dataframe(st.session_state.audit_logs, use_container_width=True)
