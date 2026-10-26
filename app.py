from datetime import datetime
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Hệ Thống Đăng Ký Đề Tài & GVHD", page_icon="🎓", layout="centered"
)

# Khai báo file
FILE_HV = "DS HV AI.xlsx"
FILE_DE_TAI_CHUNG = "LIST_De_tai_Chung.xlsx"
FILE_DE_TAI_LUAT = "LIST_De_tai_Luat_ok.xlsx"
FILE_GV = "DS GVHD& ListEmail.xlsx"
FILE_MAU_DK = "DS Mau Dang ky.xlsx"
FILE_KET_QUA = "DS_Dang_Ky_Thanh_Cong.xlsx"


@st.cache_data
def load_and_prepare_data():
    try:
        df_hv = pd.read_excel(FILE_HV)
        df_dt_chung = pd.read_excel(FILE_DE_TAI_CHUNG)
        df_dt_luat = pd.read_excel(FILE_DE_TAI_LUAT)

        # Đọc 2 sheet từ file GV và ghép theo Họ và tên
        xls_gv = pd.ExcelFile(FILE_GV)
        # Tự động lấy tên sheet đầu tiên và sheet chứa chữ email
        sheet_gv = xls_gv.sheet_names[0]
        sheet_email = [s for s in xls_gv.sheet_names if "email" in s.lower()][0]

        df_gv_info = pd.read_excel(FILE_GV, sheet_name=sheet_gv)
        df_gv_email = pd.read_excel(FILE_GV, sheet_name=sheet_email)

        # Merge thông tin GV và Email
        df_gv_full = pd.merge(
            df_gv_info,
            df_gv_email[["Họ và tên", "Email"]],
            on="Họ và tên",
            how="left",
        )
        return df_hv, df_dt_chung, df_dt_luat, df_gv_full
    except Exception as e:
        st.error(f"Lỗi khi đọc file dữ liệu: {e}")
        return None, None, None, None


df_hv, df_dt_chung, df_dt_luat, df_gv = load_and_prepare_data()

# Khởi tạo file kết quả nếu chưa có
if not os.path.exists(FILE_KET_QUA):
    if os.path.exists(FILE_MAU_DK):
        df_init = pd.read_excel(FILE_MAU_DK)
    else:
        df_init = pd.DataFrame(
            columns=[
                "Mã học viên",
                "Họ và tên",
                "Ngành",
                "Chuyên ngành",
                "Tên đề tài",
                "Bộ môn",
                "GVHD 1",
                "Email GVHD",
                "Thời gian đăng ký",
            ]
        )
    df_init.to_excel(FILE_KET_QUA, index=False)

# GIAO DIỆN CHÍNH
st.title("🎓 HỆ THỐNG ĐĂNG KÝ ĐỀ TÀI & GVHD")
st.markdown("---")

if df_hv is not None and df_gv is not None:
    # 1. Tra cứu học viên
    st.subheader("1. Tra cứu thông tin học viên")
    ma_hv = st.text_input(
        "Nhập Mã Học Viên (Ví dụ: HV001, HV002, HV003):"
    ).strip()

    if ma_hv:
        hv_row = df_hv[
            df_hv["Mã học viên"].astype(str).str.lower() == ma_hv.lower()
        ]

        if hv_row.empty:
            st.error("❌ Không tìm thấy mã học viên trong hệ thống!")
        else:
            hv = hv_row.iloc[0]
            ten_hv = hv["Họ và tên"]
            nganh = str(hv["Ngành"]).strip()
            chuyen_nganh = (
                str(hv["Chuyên ngành"]).strip()
                if pd.notna(hv["Chuyên ngành"])
                else ""
            )

            st.success(f"Học viên: **{ten_hv}**")
            col1, col2 = st.columns(2)
            col1.info(f"**Ngành:** {nganh}")
            col2.info(f"**Chuyên ngành:** {chuyen_nganh}")

            st.markdown("---")
            st.subheader("2. Lựa chọn Đề tài và Giảng viên hướng dẫn")

            # XỬ LÝ QUY TẮC NGHIỆP VỤ:
            is_luat_bdkh = (
                "luật kinh tế" in nganh.lower()
                and "biến đổi khí hậu" in chuyen_nganh.lower()
            )
            is_luat_chung = "luật" in nganh.lower() and not is_luat_bdkh

            # A. Lọc danh sách đề tài
            if is_luat_bdkh:
                st.caption(
                    "💡 *Quyền ưu tiên đặc thù:* Học viên được chọn đề tài và GVHD ở tất cả các bộ môn."
                )
                df_de_tai_chon = pd.concat(
                    [df_dt_chung, df_dt_luat], ignore_index=True
                )
            elif is_luat_chung:
                st.caption(
                    "💡 *Ngành Luật:* Bạn có thể chọn đề tài ở tất cả các bộ môn trong danh mục đề tài Luật."
                )
                df_de_tai_chon = df_dt_luat
            else:
                # Ngành riêng biệt: Chỉ lọc đề tài thuộc ngành của học viên
                df_de_tai_chon = df_dt_chung[
                    df_dt_chung["Ngành"].str.strip().str.lower()
                    == nganh.lower()
                ]

            if df_de_tai_chon.empty:
                st.warning("Hiện tại chưa có đề tài mở phù hợp với ngành này.")
            else:
                # Chọn đề tài
                list_ten_de_tai = df_de_tai_chon["Tên đề tài"].tolist()
                ten_de_tai = st.selectbox("Chọn Tên đề tài:", list_ten_de_tai)

                # Lấy bộ môn của đề tài đã chọn
                info_dt = df_de_tai_chon[
                    df_de_tai_chon["Tên đề tài"] == ten_de_tai
                ].iloc[0]
                bo_mon_de_tai = str(info_dt["Bộ môn"]).strip()
                st.write(f"📌 *Đề tài thuộc Bộ môn:* **{bo_mon_de_tai}**")

                # B. Lọc danh sách Giảng viên hướng dẫn
                if is_luat_bdkh:
                    # Được chọn GVHD của bộ môn khác
                    df_gv_chon = df_gv
                else:
                    # Chỉ chọn GVHD của bộ môn có đề tài
                    df_gv_chon = df_gv[
                        df_gv["Bộ môn"].str.strip().str.lower()
                        == bo_mon_de_tai.lower()
                    ]

                if df_gv_chon.empty:
                    st.error(
                        f"Bộ môn '{bo_mon_de_tai}' hiện chưa có GVHD trong danh sách!"
                    )
                else:
                    list_ten_gv = df_gv_chon["Họ và tên"].tolist()
                    gv_huong_dan = st.selectbox(
                        "Chọn Giảng viên hướng dẫn 1:", list_ten_gv
                    )

                    # Lấy email của GV đã chọn
                    gv_info = df_gv_chon[
                        df_gv_chon["Họ và tên"] == gv_huong_dan
                    ].iloc[0]
                    email_gv = (
                        gv_info["Email"] if pd.notna(gv_info["Email"]) else ""
                    )
                    st.write(f"✉️ *Email GVHD:* `{email_gv}`")

                    # 3. Ghi dữ liệu kết quả
                    st.markdown("---")
                    if st.button("Xác Nhận Đăng Ký Đề Tài", type="primary"):
                        df_kq = pd.read_excel(FILE_KET_QUA)

                        # Nếu học viên đăng ký lại -> Xóa bản ghi cũ
                        df_kq = df_kq[
                            df_kq["Mã học viên"].astype(str)
                            != str(hv["Mã học viên"])
                        ]

                        ban_ghi_moi = {
                            "Mã học viên": hv["Mã học viên"],
                            "Họ và tên": ten_hv,
                            "Ngành": nganh,
                            "Chuyên ngành": chuyen_nganh,
                            "Tên đề tài": ten_de_tai,
                            "Bộ môn": bo_mon_de_tai,
                            "GVHD 1": gv_huong_dan,
                            "Email GVHD": email_gv,
                            "Thời gian đăng ký": datetime.now().strftime(
                                "%d/%m/%Y %H:%M:%S"
                            ),
                        }

                        df_kq = pd.concat(
                            [df_kq, pd.DataFrame([ban_ghi_moi])],
                            ignore_index=True,
                        )
                        df_kq.to_excel(FILE_KET_QUA, index=False)

                        st.balloons()
                        st.success(
                            f"🎉 Chúc mừng học viên **{ten_hv}** đã đăng ký thành công đề tài!"
                        )

# Dành cho Quản trị viên
st.markdown("---")
with st.expander("🛠️ Quản Trị Viên - Tải Kết Quả Đăng Ký"):
    if os.path.exists(FILE_KET_QUA):
        with open(FILE_KET_QUA, "rb") as f:
            st.download_button(
                label="📥 Tải file kết quả hoàn chỉnh (.xlsx)",
                data=f,
                file_name=f"Ket_Qua_Dang_Ky_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )