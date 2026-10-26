from datetime import datetime
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Hệ Thống Đăng Ký Đề Tài & GVHD", page_icon="🎓", layout="centered"
)

FILE_HV = "DS HV AI.xlsx"
FILE_DE_TAI_CHUNG = "LIST_De_tai_Chung.xlsx"
FILE_DE_TAI_LUAT = "LIST_De_tai_Luat_ok.xlsx"
FILE_GV = "DS GVHD& ListEmail.xlsx"
FILE_KET_QUA = "DS_Dang_Ky_Thanh_Cong.xlsx"


# Hàm đọc file Excel thông minh: Tự động bỏ qua các dòng tiêu đề to của trường
def doc_excel_thong_minh(file_path, tu_khoa_nhan_dien):
    for skip in range(10):  # Quét từ dòng 0 đến dòng 9
        try:
            df = pd.read_excel(file_path, header=skip)
            cols_str = " ".join([str(c).lower() for c in df.columns])
            if any(tk.lower() in cols_str for tk in tu_khoa_nhan_dien):
                return df
        except Exception:
            continue
    return pd.read_excel(file_path)


# Hàm tìm cột linh hoạt
def tim_cot(df, keywords):
    for col in df.columns:
        col_clean = str(col).strip().lower()
        if any(kw.lower() in col_clean for kw in keywords):
            return col
    return None


@st.cache_data
def load_all_data():
    try:
        # 1. Đọc file Học viên (tự tìm dòng tiêu đề chứa chữ 'mã' hoặc 'học viên')
        df_hv = doc_excel_thong_minh(
            FILE_HV, ["mã", "học viên", "sinh viên", "họ và tên", "họ tên"]
        )

        # 2. Đọc file Đề tài chung
        df_dt_chung = doc_excel_thong_minh(
            FILE_DE_TAI_CHUNG, ["đề tài", "tên đề tài", "bộ môn"]
        )

        # 3. Đọc file Đề tài Luật
        df_dt_luat = doc_excel_thong_minh(
            FILE_DE_TAI_LUAT, ["đề tài", "tên đề tài", "bộ môn"]
        )

        # 4. Đọc file Giảng viên & Email
        xls_gv = pd.ExcelFile(FILE_GV)
        sheet_email = [
            s
            for s in xls_gv.sheet_names
            if "email" in s.lower() or "mail" in s.lower()
        ]

        if sheet_email:
            sheet_em = sheet_email[0]
            sheet_info = [s for s in xls_gv.sheet_names if s != sheet_em][0]
            df_info = doc_excel_thong_minh(
                xls_gv, ["họ", "tên", "giảng viên", "bộ môn"]
            )
            df_email = pd.read_excel(xls_gv, sheet_name=sheet_em)

            c_ten1 = tim_cot(df_info, ["họ và tên", "họ tên", "tên"])
            c_bm = tim_cot(df_info, ["bộ môn", "khoa"])
            c_ten2 = tim_cot(df_email, ["họ và tên", "họ tên", "tên"])
            c_mail = tim_cot(df_email, ["email", "mail"])

            df_info = df_info.rename(
                columns={c_ten1: "Họ và tên", c_bm: "Bộ môn"}
            )
            df_email = df_email.rename(
                columns={c_ten2: "Họ và tên", c_mail: "Email"}
            )
            df_info["Họ và tên"] = df_info["Họ và tên"].astype(str).str.strip()
            df_email["Họ và tên"] = (
                df_email["Họ và tên"].astype(str).str.strip()
            )
            df_gv = pd.merge(
                df_info[["Họ và tên", "Bộ môn"]],
                df_email[["Họ và tên", "Email"]],
                on="Họ và tên",
                how="left",
            )
        else:
            df_gv = doc_excel_thong_minh(
                FILE_GV, ["họ", "tên", "giảng viên", "email"]
            )
            c_ten = tim_cot(df_gv, ["họ và tên", "họ tên", "tên"])
            c_bm = tim_cot(df_gv, ["bộ môn", "khoa"])
            c_mail = tim_cot(df_gv, ["email", "mail"])
            df_gv = df_gv.rename(
                columns={c_ten: "Họ và tên", c_bm: "Bộ môn", c_mail: "Email"}
            )

        # Đổi tên chuẩn cho bảng Đề tài
        c_dt_chung = tim_cot(df_dt_chung, ["tên đề tài", "đề tài"])
        c_ng_chung = tim_cot(df_dt_chung, ["ngành"])
        c_bm_chung = tim_cot(df_dt_chung, ["bộ môn", "khoa"])
        df_dt_chung = df_dt_chung.rename(
            columns={
                c_dt_chung: "Tên đề tài",
                c_ng_chung: "Ngành",
                c_bm_chung: "Bộ môn",
            }
        )

        c_dt_luat = tim_cot(df_dt_luat, ["tên đề tài", "đề tài"])
        c_bm_luat = tim_cot(df_dt_luat, ["bộ môn", "khoa"])
        df_dt_luat = df_dt_luat.rename(
            columns={c_dt_luat: "Tên đề tài", c_bm_luat: "Bộ môn"}
        )

        return df_hv, df_dt_chung, df_dt_luat, df_gv
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
        return None, None, None, None


df_hv, df_dt_chung, df_dt_luat, df_gv = load_all_data()

# File kết quả
if not os.path.exists(FILE_KET_QUA):
    pd.DataFrame(
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
    ).to_excel(FILE_KET_QUA, index=False)

st.title("🎓 HỆ THỐNG ĐĂNG KÝ ĐỀ TÀI & GVHD")
st.markdown("---")

if df_hv is not None and df_gv is not None:
    st.subheader("1. Tra cứu thông tin học viên")
    ma_hv_nhap = st.text_input("Nhập Mã Học Viên (Ví dụ: 252010006):").strip()

    if ma_hv_nhap:
        # TÌM KIẾM HỌC VIÊN TRÊN TOÀN BỘ CÁC CỘT (Không bao giờ lo sai tên cột)
        dong_tim_thay = None
        for idx, row in df_hv.iterrows():
            row_str = " ".join([str(val).strip() for val in row.values])
            if ma_hv_nhap.lower() in row_str.lower():
                dong_tim_thay = row
                break

        if dong_tim_thay is None:
            st.error(
                f"❌ Không tìm thấy mã học viên '{ma_hv_nhap}' trong danh sách!"
            )
        else:
            # Nhận diện thông tin học viên từ dòng tìm được
            cot_ten = tim_cot(df_hv, ["họ và tên", "họ tên", "tên"])
            cot_nganh = tim_cot(df_hv, ["ngành"])
            cot_chuyennganh = tim_cot(
                df_hv, ["chuyên ngành", "chuyen nganh", "chuyên sâu"]
            )

            ten_hv = (
                str(dong_tim_thay[cot_ten])
                if cot_ten
                else "Học viên " + ma_hv_nhap
            )
            nganh = str(dong_tim_thay[cot_nganh]) if cot_nganh else "Luật"
            chuyen_nganh = (
                str(dong_tim_thay[cot_chuyennganh])
                if cot_chuyennganh and pd.notna(dong_tim_thay[cot_chuyennganh])
                else ""
            )

            st.success(f"Học viên: **{ten_hv}**")
            c1, c2 = st.columns(2)
            c1.info(f"**Ngành:** {nganh}")
            c2.info(f"**Chuyên ngành:** {chuyen_nganh if chuyen_nganh else 'Không có'}")

            st.markdown("---")
            st.subheader("2. Lựa chọn Đề tài và Giảng viên hướng dẫn")

            # Xử lý quy tắc nghiệp vụ
            is_luat_bdkh = (
                "luật kinh tế" in nganh.lower()
                and "biến đổi khí hậu" in chuyen_nganh.lower()
            )
            is_luat_chung = "luật" in nganh.lower() and not is_luat_bdkh

            if is_luat_bdkh:
                st.caption(
                    "💡 *Ưu tiên đặc thù:* Bạn được chọn đề tài và GVHD ở tất cả các bộ môn."
                )
                df_dt_chon = pd.concat(
                    [df_dt_chung, df_dt_luat], ignore_index=True
                )
            elif is_luat_chung:
                st.caption(
                    "💡 *Ngành Luật:* Bạn có thể chọn đề tài ở tất cả bộ môn trong danh mục Luật."
                )
                df_dt_chon = df_dt_luat
            else:
                df_dt_chon = df_dt_chung[
                    df_dt_chung["Ngành"].astype(str).str.strip().str.lower()
                    == nganh.lower()
                ]

            if df_dt_chon.empty:
                st.warning("Hiện tại chưa có đề tài mở phù hợp với ngành.")
            else:
                list_dt = df_dt_chon["Tên đề tài"].dropna().unique().tolist()
                ten_dt = st.selectbox("Chọn Tên đề tài:", list_dt)

                info_dt = df_dt_chon[df_dt_chon["Tên đề tài"] == ten_dt].iloc[0]
                bo_mon_dt = str(info_dt.get("Bộ môn", "")).strip()
                st.write(f"📌 *Đề tài thuộc Bộ môn:* **{bo_mon_dt}**")

                # Lọc GVHD
                if is_luat_bdkh:
                    df_gv_chon = df_gv
                else:
                    df_gv_chon = df_gv[
                        df_gv["Bộ môn"].astype(str).str.strip().str.lower()
                        == bo_mon_dt.lower()
                    ]

                if df_gv_chon.empty:
                    st.error(f"Bộ môn '{bo_mon_dt}' hiện chưa có GVHD.")
                else:
                    list_gv = df_gv_chon["Họ và tên"].dropna().unique().tolist()
                    gv_chon = st.selectbox(
                        "Chọn Giảng viên hướng dẫn 1:", list_gv
                    )

                    gv_info = df_gv_chon[
                        df_gv_chon["Họ và tên"] == gv_chon
                    ].iloc[0]
                    email_gv = str(gv_info.get("Email", ""))
                    st.write(
                        f"✉️ *Email GVHD:* `{email_gv if email_gv != 'nan' else 'Chưa có email'}`"
                    )

                    st.markdown("---")
                    if st.button("Xác Nhận Đăng Ký Đề Tài", type="primary"):
                        df_kq = pd.read_excel(FILE_KET_QUA)
                        df_kq = df_kq[
                            df_kq["Mã học viên"].astype(str) != str(ma_hv_nhap)
                        ]

                        ban_ghi = {
                            "Mã học viên": ma_hv_nhap,
                            "Họ và tên": ten_hv,
                            "Ngành": nganh,
                            "Chuyên ngành": chuyen_nganh,
                            "Tên đề tài": ten_dt,
                            "Bộ môn": bo_mon_dt,
                            "GVHD 1": gv_chon,
                            "Email GVHD": (
                                email_gv if email_gv != "nan" else ""
                            ),
                            "Thời gian đăng ký": datetime.now().strftime(
                                "%d/%m/%Y %H:%M:%S"
                            ),
                        }
                        df_kq = pd.concat(
                            [df_kq, pd.DataFrame([ban_ghi])], ignore_index=True
                        )
                        df_kq.to_excel(FILE_KET_QUA, index=False)

                        st.balloons()
                        st.success(
                            f"🎉 Chúc mừng học viên **{ten_hv}** đã đăng ký thành công!"
                        )

# Quản trị viên
st.markdown("---")
with st.expander("🛠️ Dành cho Quản trị viên (Tải danh sách đăng ký)"):
    if os.path.exists(FILE_KET_QUA):
        with open(FILE_KET_QUA, "rb") as f:
            st.download_button(
                label="📥 Tải file kết quả (.xlsx)",
                data=f,
                file_name=f"Ket_Qua_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
