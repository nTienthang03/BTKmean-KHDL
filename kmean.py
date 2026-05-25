import os

# Tránh warning KMeans (MKL) trên Windows khi dữ liệu nhỏ.
# Cần set trước khi import numpy/sklearn.
os.environ.setdefault("OMP_NUM_THREADS", "1")

import re
import warnings
from datetime import date, datetime

import matplotlib.pyplot as plt
import pandas as pd
from sklearn.cluster import KMeans


# Ẩn warning đã biết của scikit-learn trên Windows + MKL.
warnings.filterwarnings(
    "ignore",
    message=r"KMeans is known to have a memory leak on Windows with MKL.*",
    category=UserWarning,
)


GPA_COL = "Điểm TK (4)"
INPUT_XLSX = "TỔNG HỢP ĐIỂM K58KTP.xlsx"
EXCLUDED_COURSE_CODES = {"B103BC1", "B103BR1", "BAS0109"}


def _project_root() -> str:
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def _classify_gpa(gpa: float) -> int:
    if gpa >= 3.2:
        return 1
    if gpa >= 2.5:
        return 2
    return 3


def _find_header_row(excel_path: str, required_col: str, max_scan_rows: int = 50) -> int | None:
    raw = pd.read_excel(excel_path, header=None)
    scan_rows = min(len(raw), max_scan_rows)
    required_norm = required_col.replace("\n", " ").strip().lower()

    for row_idx in range(scan_rows):
        for cell in raw.iloc[row_idx].tolist():
            if pd.isna(cell):
                continue
            cell_norm = str(cell).replace("\n", " ").strip().lower()
            if cell_norm == required_norm or required_norm in cell_norm:
                return row_idx
    return None


def _coerce_grade(value) -> float | None:
    if pd.isna(value):
        return None

    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()
    if isinstance(value, (datetime, date)):
        # Excel đôi khi tự hiểu nhầm điểm (vd 3.5) thành ngày (vd 03/05) -> pandas đọc thành date.
        # Ta thử suy ra lại điểm theo 2 cách và chỉ nhận giá trị hợp lệ trong [0, 4].
        day = float(value.day)
        month = float(value.month)

        candidate_a = day + month / 10.0
        candidate_b = month + day / 10.0

        valid_a = 0.0 <= candidate_a <= 4.0
        valid_b = 0.0 <= candidate_b <= 4.0

        if valid_a and not valid_b:
            return candidate_a
        if valid_b and not valid_a:
            return candidate_b
        if valid_a and valid_b:
            # Trường hợp mơ hồ (vd 02/03): ưu tiên theo quy ước day + month/10 như cũ.
            return candidate_a
        return None

    text = str(value).strip()
    if not text:
        return None

    text = text.replace(",", ".")
    while ".." in text:
        text = text.replace("..", ".")

    # Chỉ giữ lại chữ số và dấu chấm
    text = re.sub(r"[^0-9.]", "", text)
    if not text:
        return None

    try:
        return float(text)
    except ValueError:
        return None


def _extract_students_from_matrix(excel_path: str) -> pd.DataFrame:
    raw = pd.read_excel(excel_path, header=None)

    # Tìm hàng MSSV và cột nhãn (thường là cột 2)
    mssv_row = None
    mssv_label_col = None
    for i in range(min(len(raw), 20)):
        for j in range(min(raw.shape[1], 10)):
            cell = raw.iat[i, j]
            if pd.isna(cell):
                continue
            if str(cell).replace("\n", " ").strip().lower() == "mssv":
                mssv_row = i
                mssv_label_col = j
                break
        if mssv_row is not None:
            break

    if mssv_row is None or mssv_label_col is None:
        raise ValueError("Không tìm thấy hàng 'MSSV' trong file Excel.")

    # Hàng tên SV
    name_row = None
    for i in range(min(len(raw), 25)):
        cell = raw.iat[i, mssv_label_col]
        if pd.isna(cell):
            continue
        if "tên" in str(cell).lower():
            name_row = i
            break

    if name_row is None:
        raise ValueError("Không tìm thấy hàng tên sinh viên (ô có chữ 'Tên').")

    # Hàng STT sinh viên (ngay phía trên MSSV là trường hợp phổ biến)
    stt_row = mssv_row - 1 if mssv_row - 1 >= 0 else None

    # Các cột sinh viên bắt đầu từ sau cột nhãn
    start_col = mssv_label_col + 1
    last_col = raw.iloc[mssv_row, start_col:].last_valid_index()
    if last_col is None:
        raise ValueError("Không đọc được danh sách MSSV theo cột.")

    student_cols = list(range(start_col, int(last_col) + 1))

    stt_vals = (
        raw.iloc[stt_row, student_cols].tolist() if stt_row is not None else [None] * len(student_cols)
    )
    mssv_vals = raw.iloc[mssv_row, student_cols].tolist()
    name_vals = raw.iloc[name_row, student_cols].tolist()

    # Xác định vùng môn học: tìm hàng tiêu đề "Mã Môn học" / "Tên Môn học"
    course_header_row = None
    for i in range(min(len(raw), 40)):
        c0 = str(raw.iat[i, 0]).strip().lower() if not pd.isna(raw.iat[i, 0]) else ""
        c1 = str(raw.iat[i, 1]).strip().lower() if raw.shape[1] > 1 and not pd.isna(raw.iat[i, 1]) else ""
        c2 = str(raw.iat[i, 2]).strip().lower() if raw.shape[1] > 2 and not pd.isna(raw.iat[i, 2]) else ""
        if c0 == "stt" and "mã" in c1 and "môn" in c1 and "tên" in c2 and "môn" in c2:
            course_header_row = i
            break

    if course_header_row is None:
        raise ValueError("Không tìm thấy dòng tiêu đề môn học (STT | Mã Môn học | Tên Môn học).")

    course_rows = []
    for i in range(course_header_row + 1, len(raw)):
        ma_mon = raw.iat[i, 1] if raw.shape[1] > 1 else None
        ten_mon = raw.iat[i, 2] if raw.shape[1] > 2 else None
        if pd.isna(ma_mon) or pd.isna(ten_mon):
            continue

        ma_mon_norm = str(ma_mon).strip().upper()
        if ma_mon_norm in EXCLUDED_COURSE_CODES:
            continue

        course_rows.append(i)

    if not course_rows:
        raise ValueError("Không tìm thấy dữ liệu môn học để tính GPA.")

    grade_matrix_raw = raw.iloc[course_rows, student_cols]
    grade_matrix = grade_matrix_raw.apply(lambda col: col.map(_coerce_grade))
    gpa_series = grade_matrix.mean(axis=0, skipna=True)

    df_students = pd.DataFrame(
        {
            "STT": pd.to_numeric(pd.Series(stt_vals), errors="coerce"),
            "MSSV": pd.Series(mssv_vals, dtype="string"),
            "Tên sinh viên": pd.Series(name_vals, dtype="string"),
            GPA_COL: gpa_series.values,
        }
    )

    df_students[GPA_COL] = pd.to_numeric(df_students[GPA_COL], errors="coerce")
    return df_students.dropna(subset=[GPA_COL]).copy()


def main() -> None:
    root = _project_root()
    data_path = os.path.join(root, "data", INPUT_XLSX)
    output_dir = os.path.join(root, "output")
    os.makedirs(output_dir, exist_ok=True)

    # Ưu tiên đọc theo dạng bảng thường (có cột Điểm TK (4)). Nếu không có, fallback sang dạng ma trận.
    header_row = _find_header_row(data_path, required_col=GPA_COL)
    df_try = pd.read_excel(data_path, header=header_row if header_row is not None else 0)
    df_try.columns = df_try.columns.astype(str).str.replace("\n", " ").str.strip()

    if GPA_COL in df_try.columns:
        # Có thể gặp trường hợp ô GPA bị Excel lưu kiểu date -> cần coerce trước.
        df_try[GPA_COL] = df_try[GPA_COL].map(_coerce_grade)
        df_try[GPA_COL] = pd.to_numeric(df_try[GPA_COL], errors="coerce")
        df_diem = df_try.dropna(subset=[GPA_COL]).copy()
    else:
        df_diem = _extract_students_from_matrix(data_path)

    df_diem[GPA_COL] = pd.to_numeric(df_diem[GPA_COL], errors="coerce").round(2)

    # K-means (K=3) để minh hoạ phân cụm dữ liệu theo GPA hệ 4
    X = df_diem[[GPA_COL]]
    kmeans = KMeans(n_clusters=3, random_state=42, n_init=10)
    kmeans.fit(X)

    # Phân loại theo tiêu chí ngưỡng GPA (đúng theo đề bài)
    df_diem["Nhom"] = df_diem[GPA_COL].apply(_classify_gpa)

    def _reset_stt_for_output(df: pd.DataFrame) -> pd.DataFrame:
        df_out = df.copy()
        if "STT" in df_out.columns:
            df_out["STT"] = range(1, len(df_out) + 1)
            cols = ["STT"] + [c for c in df_out.columns if c != "STT"]
            return df_out[cols]
        df_out.insert(0, "STT", range(1, len(df_out) + 1))
        return df_out

    df_nhom_1 = _reset_stt_for_output(df_diem[df_diem["Nhom"] == 1].drop(columns=["Nhom"]))
    df_nhom_2 = _reset_stt_for_output(df_diem[df_diem["Nhom"] == 2].drop(columns=["Nhom"]))
    df_nhom_3 = _reset_stt_for_output(df_diem[df_diem["Nhom"] == 3].drop(columns=["Nhom"]))

    df_nhom_1.to_excel(
        os.path.join(output_dir, "nhom_1_sinh_vien_hoc_tot.xlsx"), index=False
    )
    df_nhom_2.to_excel(
        os.path.join(output_dir, "nhom_2_sinh_vien_on_dinh.xlsx"), index=False
    )
    df_nhom_3.to_excel(
        os.path.join(output_dir, "nhom_3_sinh_vien_can_ho_tro.xlsx"), index=False
    )

    # Vẽ biểu đồ
    plt.figure(figsize=(10, 6))

    x_col = "STT" if "STT" in df_diem.columns else None
    x1 = df_nhom_1[x_col] if x_col else df_nhom_1.index
    x2 = df_nhom_2[x_col] if x_col else df_nhom_2.index
    x3 = df_nhom_3[x_col] if x_col else df_nhom_3.index

    plt.scatter(x1, df_nhom_1[GPA_COL], label="Nhóm 1 - Sinh viên học tốt")
    plt.scatter(x2, df_nhom_2[GPA_COL], label="Nhóm 2 - Sinh viên ổn định")
    plt.scatter(
        x3, df_nhom_3[GPA_COL], label="Nhóm 3 - Sinh viên cần hỗ trợ"
    )

    plt.axhline(y=3.2, linestyle="--", label="Ngưỡng GPA 3.2")
    plt.axhline(y=2.5, linestyle="--", label="Ngưỡng GPA 2.5")

    plt.xlabel("STT sinh viên" if x_col else "Dòng dữ liệu")
    plt.ylabel("Điểm hệ 4")
    plt.title("Phân nhóm học lực theo GPA hệ 4 (K=3)")
    plt.legend()
    plt.grid(True)

    plt.savefig(os.path.join(output_dir, "bieu_do_phan_cum.png"), dpi=300)
    plt.show()

    print(
        "Đã tạo xong 3 file Excel và 1 biểu đồ trong thư mục output. "
        f"(Nhóm 1: {len(df_nhom_1)}, Nhóm 2: {len(df_nhom_2)}, Nhóm 3: {len(df_nhom_3)})"
    )


if __name__ == "__main__":
    main()