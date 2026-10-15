import streamlit as st
import pandas as pd
from unidecode import unidecode
import io
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter
from streamlit_autorefresh import st_autorefresh
from supabase import create_client, Client

# ==========================================
# 1. CẤU HÌNH TRANG WEB & TỰ ĐỘNG LÀM MỚI (REAL-TIME)
# ==========================================
st.set_page_config(page_title="Hệ thống Kiểm phiếu tự động", layout="wide")
st.title("📊 Hệ thống Kiểm phiếu Tự động (Đồng bộ đa máy - Siêu tốc)")

# Tự động quét giao diện sau mỗi 3 giây. Chu kỳ này cũng được dùng để đẩy dữ liệu đệm lên Cloud
st_autorefresh(interval=3000, key="datarefresh")

# Khởi tạo bộ nhớ đệm cục bộ trên trình duyệt để hỗ trợ bấm dồn dập không trễ
if "phieu_cho_xu_ly" not in st.session_state:
    st.session_state.phieu_cho_xu_ly = {}

# ==========================================
# 2. KẾT NỐI CƠ SỞ DỮ LIỆU ĐÁM MÂY SUPABASE
# ==========================================
@st.cache_resource
def init_supabase() -> Client:
    url: str = st.secrets["supabase"]["url"].strip("/")
    key: str = st.secrets["supabase"]["key"]
    return create_client(url, key)

supabase: Client = init_supabase()

# Đọc dữ liệu cấu hình thông số tổng
try:
    response_ts = supabase.table("thong_so").select("*").eq("id", 1).execute()
    if response_ts.data:
        ts_data = response_ts.data[0]
    else:
        insert_res = supabase.table("thong_so").insert({"id": 1, "vong": "Vòng đề cử", "phiu_phat": 0, "phiu_thu": 0, "phiu_huy": 0, "phiu_hop_le": 0, "phiu_trang": 0}).execute()
        ts_data = insert_res.data[0]
except Exception as e:
    st.error(f"Lỗi kết nối bảng thong_so: {e}")
    st.stop()

# ĐỒNG BỘ NGẦM: Trước khi đọc bảng ứng viên, kiểm tra xem máy này có phiếu đệm nào chưa đẩy lên Cloud không
if st.session_state.phieu_cho_xu_ly:
    try:
        for ten_uv, so_luong_them in list(st.session_state.phieu_cho_xu_ly.items()):
            if so_luong_them > 0:
                # Lấy số phiếu hiện tại trên Cloud của ứng viên đó
                res_check = supabase.table("ung_vien").select("so_phieu").eq("ten", ten_uv).execute()
                if res_check.data:
                    phieu_hien_tai = res_check.data[0]["so_phieu"]
                    # Cộng dồn lượng phiếu đệm tích lũy và cập nhật một lần duy nhất lên Cloud
                    supabase.table("ung_vien").update({
                        "so_phieu": phieu_hien_tai + so_luong_them,
                        "vong": ts_data.get('vong', 'Vòng đề cử')
                    }).eq("ten", ten_uv).execute()
        # Đẩy dữ liệu thành công thì xóa sạch hàng đợi chờ xử lý của máy này
        st.session_state.phieu_cho_xu_ly = {}
    except Exception:
        pass # Nếu mạng lỗi, giữ lại hàng đệm để chu kỳ sau đẩy tiếp

# Đọc danh sách ứng viên mới nhất từ hệ thống trực tuyến
try:
    response_uv = supabase.table("ung_vien").select("*").execute()
    danh_sach_uv = response_uv.data if response_uv.data else []
except Exception as e:
    st.error(f"Lỗi kết nối bảng ung_vien: {e}")
    st.stop()

# ==========================================
# 3. MỤC 1 - THIẾT LẬP THÔNG SỐ & TẢI DANH SÁCH ỨNG VIÊN
# ==========================================
st.header("1. Thiết lập thông số và Tải danh sách ứng viên")

col1, col2, col3, col4, col5, col6 = st.columns(6)
with col1:
    vong_list = ["Vòng đề cử", "Vòng 1", "Vòng 2", "Vòng 3"]
    current_vong = ts_data.get('vong', 'Vòng đề cử')
    idx_vong = vong_list.index(current_vong) if current_vong in vong_list else 0
    vong = st.selectbox("Chọn vòng tín nhiệm:", vong_list, index=idx_vong)
with col2:
    phiu_phat_ra = st.number_input("Số phiếu phát ra:", min_value=0, value=int(ts_data.get('phiu_phat', 0)), step=1)
with col3:
    phiu_thu_vao = st.number_input("Số phiếu thu vào:", min_value=0, value=int(ts_data.get('phiu_thu', 0)), step=1)
with col4:
    phiu_bat_hop_le = st.number_input("Phiếu bất hợp lệ:", min_value=0, value=int(ts_data.get('phiu_huy', 0)), step=1)
with col5:
    phiu_hop_le = st.number_input("Số phiếu hợp lệ:", min_value=0, value=int(ts_data.get('phiu_hop_le', 0)), step=1)
with col6:
    phiu_trang = st.number_input("Số phiếu trắng (hợp lệ):", min_value=0, value=int(ts_data.get('phiu_trang', 0)), step=1)

tong_mau_so = phiu_hop_le + phiu_trang

col_save_ts, col_reset_system = st.columns(2)
with col_save_ts:
    if st.button("💾 Lưu cấu hình thông số", use_container_width=True, type="primary"):
        supabase.table("thong_so").update({
            "vong": vong, 
            "phiu_phat": phiu_phat_ra, 
            "phiu_thu": phiu_thu_vao, 
            "phiu_huy": phiu_bat_hop_le, 
            "phiu_hop_le": phiu_hop_le, 
            "phiu_trang": phiu_trang
        }).eq("id", 1).execute()
        st.success("Đã đồng bộ thông số hệ thống lên Supabase!")
        st.rerun()

with col_reset_system:
    if st.button("🧹 BẮT ĐẦU CUỘC BỎ PHIẾU MỚI", use_container_width=True, type="secondary"):
        supabase.table("ung_vien").delete().neq("ten", "").execute()
        supabase.table("thong_so").update({
            "vong": "Vòng đề cử", "phiu_phat": 0, "phiu_thu": 0, "phiu_huy": 0, "phiu_hop_le": 0, "phiu_trang": 0
        }).eq("id", 1).execute()
        st.session_state.phieu_cho_xu_ly = {}
        st.toast("Đã xóa hết dữ liệu cũ! Hãy tải file Excel mới.", icon="🗑️")
        st.rerun()

st.write("")

if not danh_sach_uv:
    uploaded_file = st.file_uploader("Tải lên danh sách ứng viên từ Excel (Cột đầu là Họ tên):", type=["xlsx", "xls"])
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            ten_ung_vien = df.iloc[:, 0].dropna().tolist()
            
            for ten in ten_ung_vien:
                supabase.table("ung_vien").upsert({"ten": str(ten).strip(), "so_phieu": 0, "vong": vong}).execute()
            st.success(f"Đã tải thành công {len(ten_ung_vien)} ứng viên lên hệ thống Cloud!")
            st.rerun()
        except Exception as e:
            st.error(f"Lỗi đọc dữ liệu file Excel hoặc đẩy lên Cloud: {e}")
else:
    st.info("💡 Hệ thống đang có dữ liệu dữ liệu trên đám mây. Hãy bấm **'🧹 BẮT ĐẦU CUỘC BỎ PHIẾU MỚI'** nếu muốn đổi danh sách.")

col_test, _ = st.columns(2)
with col_test:
    if not danh_sach_uv:
        if st.button("✨ Tạo dữ liệu thử nghiệm", use_container_width=True):
            test_names = [f"Nguyễn Văn {chr(65+i)}" for i in range(10)] + ["Nguyễn Văn Tuấn", "Trần Minh Hoàng", "Lê Thị Diễm"]
            for ten in test_names:
                supabase.table("ung_vien").upsert({"ten": ten, "so_phieu": 0, "vong": vong}).execute()
            st.rerun()

# ==========================================
# 4. MỤC 2 - KIỂM PHIẾU VÀ BẢNG XẾP HẠNG
# ==========================================
if danh_sach_uv:
    st.write("---")
    st.header("2. Kiểm phiếu & Bảng xếp hạng (Dữ liệu chung trực tuyến)")
    
    col_nhap_lieu, col_bxh = st.columns(2)
    
    # Cập nhật hiển thị số phiếu bằng cách cộng thêm lượng phiếu đệm đang chờ xử lý của riêng máy này
    df_kq = pd.DataFrame(danh_sach_uv)
    df_kq.columns = ["ten", "so_phieu", "vong"]
    
    for idx, row in df_kq.iterrows():
        ten_uv = row["ten"]
        if ten_uv in st.session_state.phieu_cho_xu_ly:
            df_kq.at[idx, "so_phieu"] += st.session_state.phieu_cho_xu_ly[ten_uv]
    
    if tong_mau_so > 0:
        df_kq["Tỷ Lệ %"] = (df_kq["so_phieu"] * 100 / tong_mau_so).round(2)
    else:
        df_kq["Tỷ Lệ %"] = 0.0
        
    df_kq["Số Phiếu Tín Nhiệm"] = df_kq["so_phieu"].astype(str) + " / " + str(tong_mau_so)
    df_kq["Đánh Giá"] = df_kq["Tỷ Lệ %"].apply(lambda x: "🟢 ĐẠT" if x >= 50.0 else "🔴 CHƯA ĐẠT")
    
    df_kq = df_kq.sort_values(by="so_phieu", ascending=False).reset_index(drop=True)
    df_kq.index += 1
    
    df_hien_thi = df_kq[["ten", "Số Phiếu Tín Nhiệm", "vong", "Tỷ Lệ %", "Đánh Giá"]]
    df_hien_thi.columns = ["Họ và Tên", "Số Phiếu Tín Nhiệm", "Số Vòng", "Tỷ Lệ %", "Ghi Chú (Đánh giá)"]

    with col_nhap_lieu:
        with st.container(border=True):
            st.subheader("⌨️ Ô nhập phiếu nhanh")
            tu_khoa = st.text_input("Gõ tên ứng viên (Không dấu để tìm kiếm):", placeholder="Ví dụ: tuan, hoang, diem...")
            
            if tu_khoa:
                tu_khoa_khong_dau = unidecode(tu_khoa).lower().strip()
                goi_y_khop = [uv for uv in danh_sach_uv if tu_khoa_khong_dau in unidecode(uv["ten"]).lower()]
                
                if goi_y_khop:
                    st.write("*Click chọn đúng tên bên dưới để cộng phiếu:*")
                    for uv_goi_y in goi_y_khop[:5]:
                        ten_goi_y = uv_goi_y["ten"]
                        
                        if st.button(f"👤 {ten_goi_y}", key=f"btn_{ten_goi_y}", use_container_width=True):
                            # TỐI ƯU HÓA: Ghi nhận phiếu vào bộ nhớ tạm cục bộ ngay lập tức (Không đợi mạng Internet)
                            if ten_goi_y not in st.session_state.phieu_cho_xu_ly:
                                st.session_state.phieu_cho_xu_ly[ten_goi_y] = 0
                            st.session_state.phieu_cho_xu_ly[ten_goi_y] += 1
                            
                            st.toast(f"✅ Đã lưu tạm 1 phiếu cho: {ten_goi_y}! (Hệ thống sẽ đồng bộ lên Cloud sau ít giây)", icon="🎉")
                            st.rerun()
                else:
                    st.caption("❌ Không tìm thấy ứng viên nào khớp.")
                
    with col_bxh:
        with st.container(border=True):
            st.subheader("🏆 Bảng xếp hạng hiển thị")
            st.dataframe(df_hien_thi, use_container_width=True, height=400)

    # ==========================================
    # 5. CHỨC NĂNG XUẤT FILE EXCEL ĐẸP
    # ==========================================
    def export_to_excel_pretty(df_source):
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "KetQuaKiemPhieu"
        
        font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        font_data = Font(name="Arial", size=11)
        fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        
        thin_border = Border(
            left=Side(style='thin', color='BFBFBF'),
            right=Side(style='thin', color='BFBFBF'),
            top=Side(style='thin', color='BFBFBF'),
            bottom=Side(style='thin', color='BFBFBF')
        )
        
        headers = ["STT"] + list(df_source.columns)
        ws.append(headers)
        
        for idx, row in enumerate(df_source.itertuples(), start=1):
            ws.append([idx] + list(row[1:]))
            
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=1):
            for cell in row:
                cell.border = thin_border
                if row_idx == 1:
                    cell.font = font_header
                    cell.fill = fill_header
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                else:
                    cell.font = font_data
                    if cell.column == 2:
                        cell.alignment = Alignment(vertical="center", horizontal="left")
                    else:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for row in range(1, ws.max_row + 1):
                val = ws.cell(row=row, column=col_idx).value
                if val:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
            
        wb.save(output)
        return output.getvalue()

    st.write("")
    excel_data = export_to_excel_pretty(df_hien_thi)
    st.download_button(
        label="📥 Tải về Báo cáo Kết quả (Định dạng Excel có khung kẻ bảng)",
        data=excel_data,
        file_name="Ket_qua_kiem_phieu_chuan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
# ==========================================
# 6. MỤC 3 - PHẦN CHỈNH SỬA VÀ ĐIỀU CHỈNH DỮ LIỆU
# ==========================================
if danh_sach_uv:
    st.write("---")
    st.header("3. Điều chỉnh dữ liệu")
    
    # Bước 1: Tạo danh sách từ khóa ngầm là "tên không dấu" để selectbox tự động lọc được khi gõ
    # Để tránh trùng lặp nếu có người trùng tên, ta đính kèm thêm ID ẩn bằng ký tự đặc biệt "||"
    danh_sach_ngam_selectbox = [""]
    ban_do_dich_nguoc_ten = {"": ""} # Bộ từ điển dùng để dịch ngược chuỗi không dấu thành tên gốc có dấu
    
    for uv in danh_sach_uv:
        ten_co_dau = uv["ten"]
        ten_khong_dau_ngam = unidecode(ten_co_dau).lower().strip()
        
        # Tạo chuỗi định danh ngầm
        chuoi_khoa_ngam = f"{ten_khong_dau_ngam}||{ten_co_dau}"
        
        danh_sach_ngam_selectbox.append(chuoi_khoa_ngam)
        ban_do_dich_nguoc_ten[chuoi_khoa_ngam] = ten_co_dau

    # Bước 2: Hàm định dạng hiển thị. Khi hiển thị lên droplist, hàm này sẽ ép selectbox
    # CẮT BỎ hoàn toàn từ khóa ngầm và chỉ hiện duy nhất họ tên tiếng Việt có dấu sạch sẽ.
    def hien_thi_ten_co_dau_sach_se(chuoi_khoa):
        return ban_do_dich_nguoc_ten.get(chuoi_khoa, chuoi_khoa)

    # Bước 3: Hiển thị 1 O^ CHỌN DUY NHẤT. Bạn gõ không dấu thoải mái, màn hình chỉ hiện tên có dấu sạch 100%
    lua_chon_khoa_ngam = st.selectbox(
        "Gõ tên ứng viên (không dấu) hoặc cuộn chuột để chọn người cần sửa/xóa:",
        danh_sach_ngam_selectbox,
        format_func=hien_thi_ten_co_dau_sach_se,
        key="select_sua"
    )
    
    # Bước 4: Tiến hành hiển thị chức năng Sửa/Xóa khi đã chọn được ứng viên cụ thể
    if len(lua_chon_khoa_ngam) > 0:
        # Lấy lại tên gốc có dấu chính xác từ bản đồ dịch ngược
        ten_can_sua = ban_do_dich_nguoc_ten[lua_chon_khoa_ngam]
        
        st.write("")
        info_hien_tai = next(item for item in danh_sach_uv if item["ten"] == ten_can_sua)
        col_info, col_action = st.columns(2)
        
        with col_info:
            st.info(f"Ứng viên đang chọn: **{ten_can_sua}**  \nSố phiếu hiện tại trên hệ thống Cloud: **{info_hien_tai['so_phieu']}** phiếu.")
            so_phieu_moi = st.number_input("Nhập chính xác số phiếu thay thế mới:", min_value=0, value=int(info_hien_tai["so_phieu"]), step=1)
            
        with col_action:
            st.write("")
            col_btn_sua, col_btn_xoa = st.columns(2)
            with col_btn_sua:
                if st.button("💾 Lưu thay đổi số phiếu", use_container_width=True, type="primary"):
                    supabase.table("ung_vien").update({"so_phieu": so_phieu_moi}).eq("ten", ten_can_sua).execute()
                    if ten_can_sua in st.session_state.phieu_cho_xu_ly:
                        del st.session_state.phieu_cho_xu_ly[ten_can_sua]
                    st.success(f"Đã cập nhật số phiếu của {ten_can_sua} thành {so_phieu_moi} lên Cloud!")
                    st.rerun()
            with col_btn_xoa:
                if st.button("🗑️ Xóa ứng viên này", use_container_width=True):
                    supabase.table("ung_vien").delete().eq("ten", ten_can_sua).execute()
                    if ten_can_sua in st.session_state.phieu_cho_xu_ly:
                        del st.session_state.phieu_cho_xu_ly[ten_can_sua]
                    st.warning(f"Đã xóa hoàn toàn ứng viên {ten_can_sua} khỏi cơ sở dữ liệu trực tuyến!")
                    st.rerun()
