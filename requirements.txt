import streamlit as st
import pandas as pd
from datetime import datetime

# ======================================
# CẤU HÌNH TRANG
# ======================================
st.set_page_config(
    page_title="Quản lý tiền phòng trọ",
    page_icon="🏠",
    layout="wide"
)

st.title("🏠 PHẦN MỀM QUẢN LÝ THU TIỀN DÃY TRỌ_VŨ ĐỨC BÌNH")

# ======================================
# SIDEBAR - ĐƠN GIÁ & CẤU HÌNH
# ======================================
st.sidebar.header("⚙️ ĐƠN GIÁ ĐIỆN NƯỚC")

gia_dien = st.sidebar.number_input(
    "⚡ Giá điện (đ/kWh)",
    min_value=0,
    value=3500,
    step=100
)

gia_nuoc = st.sidebar.number_input(
    "💧 Giá nước (đ/m³)",
    min_value=0,
    value=15000,
    step=500
)

phi_khac = st.sidebar.number_input(
    "📦 Phí khác (wifi, vệ sinh...)",
    min_value=0,
    value=150000,
    step=10000
)

so_phong = st.sidebar.slider(
    "Số lượng phòng",
    min_value=1,
    max_value=100,
    value=10,
    step=1
)

# ======================================
# CHỌN THÁNG LÀM VIỆC (ĐÃ SỬA THÀNH CHUNG MỘT DÒNG)
# ======================================

# Tạo 2 cột bên trong Sidebar để đặt Năm và Tháng nằm song song chung 1 dòng
col_nam, col_thang = st.sidebar.columns(2)

with col_nam:
    nam_hien_tai = datetime.now().year
    danh_sach_nam = list(range(nam_hien_tai - 1, nam_hien_tai + 6))
    nam_chon = st.selectbox(
        "Chọn năm:",
        danh_sach_nam,
        index=danh_sach_nam.index(nam_hien_tai)
    )

with col_thang:
    thang_hien_tai = datetime.now().month
    danh_sach_thang = list(range(1, 13))
    idx_thang_chon = st.selectbox(
        "Chọn tháng:",
        danh_sach_thang,
        index=thang_hien_tai - 1,
        format_func=lambda x: f"Tháng {x}"
    )

# Chuỗi hiển thị tiêu đề trực quan
thang_chon = f"Tháng {idx_thang_chon}/{nam_chon}"

# ======================================
# KHỔI TẠO & ĐỒNG BỘ DỮ LIỆU THEO THÁNG VÀ NĂM
# ======================================
if "phong_data_theo_thang" not in st.session_state:
    st.session_state.phong_data_theo_thang = {}

# Khởi tạo nhánh cho Năm nếu chưa có
if nam_chon not in st.session_state.phong_data_theo_thang:
    st.session_state.phong_data_theo_thang[nam_chon] = {}

# Khởi tạo nhánh cho Tháng trong Năm đó nếu chưa có
if idx_thang_chon not in st.session_state.phong_data_theo_thang[nam_chon]:
    st.session_state.phong_data_theo_thang[nam_chon][idx_thang_chon] = {}

for i in range(1, so_phong + 1):
    if i not in st.session_state.phong_data_theo_thang[nam_chon][idx_thang_chon]:
        
        # TỰ ĐỘNG LẤY SỐ CŨ TỪ SỐ MỚI CỦA THÁNG TRƯỚC
        dien_cu_mac_dinh = 0
        nuoc_cu_mac_dinh = 0
        
        # Xác định tháng trước và năm trước (Xử lý trường hợp tháng 1)
        if idx_thang_chon == 1:
            thang_truoc = 12
            nam_truoc = nam_chon - 1
        else:
            thang_truoc = idx_thang_chon - 1
            nam_truoc = nam_chon
        
        # Nếu có dữ liệu tháng trước, kế thừa số điện/nước mới của tháng trước
        if nam_truoc in st.session_state.phong_data_theo_thang:
            if thang_truoc in st.session_state.phong_data_theo_thang[nam_truoc]:
                if i in st.session_state.phong_data_theo_thang[nam_truoc][thang_truoc]:
                    data_thang_truoc = st.session_state.phong_data_theo_thang[nam_truoc][thang_truoc][i]
                    dien_cu_mac_dinh = data_thang_truoc.get("dien_moi", 0)
                    nuoc_cu_mac_dinh = data_thang_truoc.get("nuoc_moi", 0)

        st.session_state.phong_data_theo_thang[nam_chon][idx_thang_chon][i] = {
            "gia_phong": 2500000,
            "dien_cu": dien_cu_mac_dinh,
            "dien_moi": dien_cu_mac_dinh, 
            "nuoc_cu": nuoc_cu_mac_dinh,
            "nuoc_moi": nuoc_cu_mac_dinh, 
            "ghi_chu": ""
        }

# ======================================
# GIAO DIỆN NHẬP DỮ LIỆU
# ======================================
st.header(f"📋 Nhập số liệu - {thang_chon}")

phong = st.selectbox(
    "Chọn phòng cần nhập dữ liệu:",
    range(1, so_phong + 1),
    key=f"selectbox_phong_{nam_chon}_{idx_thang_chon}" # Đảm bảo không lỗi cache khi đổi thời gian
)

# Lấy dữ liệu của phòng được chọn trong tháng và năm được chọn
data = st.session_state.phong_data_theo_thang[nam_chon][idx_thang_chon][phong]

st.subheader(f"🏠 Cập nhật: Phòng {phong} ({thang_chon})")
c1, c2, c3, c4, c5 = st.columns(5)

with c1:
    data["gia_phong"] = st.number_input(
        "Giá phòng (VNĐ)",
        min_value=0,
        value=int(data["gia_phong"]),
        step=50000,
        key=f"gia_{nam_chon}_{idx_thang_chon}_{phong}"
    )

with c2:
    data["dien_cu"] = st.number_input(
        "Số điện CŨ",
        min_value=0,
        value=int(data["dien_cu"]),
        key=f"diencu_{nam_chon}_{idx_thang_chon}_{phong}"
    )

with c3:
    data["dien_moi"] = st.number_input(
        "Số điện MỚI",
        min_value=0,
        value=int(data["dien_moi"]),
        key=f"dienmoi_{nam_chon}_{idx_thang_chon}_{phong}"
    )

with c4:
    data["nuoc_cu"] = st.number_input(
        "Số nước CŨ",
        min_value=0,
        value=int(data["nuoc_cu"]),
        key=f"nuoccu_{nam_chon}_{idx_thang_chon}_{phong}"
    )

with c5:
    data["nuoc_moi"] = st.number_input(
        "Số nước MỚI",
        min_value=0,
        value=int(data["nuoc_moi"]),
        key=f"nuocmoi_{nam_chon}_{idx_thang_chon}_{phong}"
    )

data["ghi_chu"] = st.text_input(
    "📝 Ghi chú (Ví dụ: Đã thanh toán, chưa thanh toán...)",
    value=data["ghi_chu"],
    key=f"ghichu_{nam_chon}_{idx_thang_chon}_{phong}"
)

# Lưu ngược lại session_state theo năm, tháng và phòng
st.session_state.phong_data_theo_thang[nam_chon][idx_thang_chon][phong] = data

# Cập nhật nhanh số cũ cho tháng sau (Xử lý đồng bộ dữ liệu thời gian thực bao gồm cả khi chuyển giao năm)
if idx_thang_chon == 12:
    thang_sau = 1
    nam_sau = nam_chon + 1
else:
    thang_sau = idx_thang_chon + 1
    nam_sau = nam_chon

if nam_sau in st.session_state.phong_data_theo_thang:
    if thang_sau in st.session_state.phong_data_theo_thang[nam_sau]:
        if phong in st.session_state.phong_data_theo_thang[nam_sau][thang_sau]:
            st.session_state.phong_data_theo_thang[nam_sau][thang_sau][phong]["dien_cu"] = data["dien_moi"]
            st.session_state.phong_data_theo_thang[nam_sau][thang_sau][phong]["nuoc_cu"] = data["nuoc_moi"]

# ======================================
# XỬ LÝ & TÍNH TOÁN DỮ LIỆU
# ======================================
ket_qua = []
tong_doanh_thu = 0

for i in range(1, so_phong + 1):
    # Đảm bảo phòng i có tồn tại trong bộ nhớ tháng này
    if i in st.session_state.phong_data_theo_thang[nam_chon][idx_thang_chon]:
        d = st.session_state.phong_data_theo_thang[nam_chon][idx_thang_chon][i]
    else:
        continue

    so_dien = max(d["dien_moi"] - d["dien_cu"], 0)
    so_nuoc = max(d["nuoc_moi"] - d["nuoc_cu"], 0)

    tien_dien = so_dien * gia_dien
    tien_nuoc = so_nuoc * gia_nuoc
    tong_tien = d["gia_phong"] + tien_dien + tien_nuoc + phi_khac

    tong_doanh_thu += tong_tien

    # Đã lược bỏ các cột Điện tiêu thụ và Nước tiêu thụ theo cấu trúc mong muốn
    ket_qua.append({
        "Phòng": f"Phòng {i}",
        "Giá phòng": d["gia_phong"],
        "Số điện cũ": d["dien_cu"],
        "Số điện mới": d["dien_moi"],
        "Tiền điện": tien_dien,
        "Số nước cũ": d["nuoc_cu"],
        "Số nước mới": d["nuoc_moi"],
        "Tiền nước": tien_nuoc,
        "Phí khác": phi_khac,
        "Tổng tiền": tong_tien,
        "Ghi chú": d["ghi_chu"]
    })

# ======================================
# HIỂN THỊ BẢNG KẾT QUẢ & XUẤT FILE
# ======================================
st.divider()
st.header(f"📊 Bảng tổng hợp tiền trọ - {thang_chon}")

if ket_qua:
    df = pd.DataFrame(ket_qua)

    # Đã loại bỏ chữ "VNĐ" ở thuộc tính định dạng của tất cả các cột tiền tệ
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Giá phòng": st.column_config.NumberColumn(format="%d"),
            "Tiền điện": st.column_config.NumberColumn(format="%d"),
            "Tiền nước": st.column_config.NumberColumn(format="%d"),
            "Phí khác": st.column_config.NumberColumn(format="%d"),
            "Tổng tiền": st.column_config.NumberColumn(format="%d"),
        }
    )

    # Hiển thị tổng doanh thu (Vẫn giữ VNĐ ở ô Metric bên ngoài để dễ nhìn)
    st.metric(
        f"💰 TỔNG THU NHẬP CÁC PHÒNG ({thang_chon})",
        f"{tong_doanh_thu:,.0f} VNĐ"
    )

    # Nút xuất file dữ liệu CSV kèm tên tháng và năm
    csv = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label=f"📥 Xuất Báo Cáo {thang_chon} (CSV)",
        data=csv,
        file_name=f"Bang_Tien_Phong_{thang_chon.replace('/', '_')}.csv",
        mime="text/csv"
    )
else:
    st.info("Chưa có dữ liệu phòng nào cho tháng này.")
