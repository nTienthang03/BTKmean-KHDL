
#  BÀI THỰC HÀNH buổi 02 KHOA HỌC DỮ LIỆU
 Link youtube : 

## Thông tin sinh viên

* **Họ và tên:** Nguyễn Tiến Thắng
* **MSSV:** K225480106058
* **Lớp:** K58KTP

## Yêu cầu bài tập

1. Mỗi sinh viên nộp **bảng điểm lớp** cho lớp trưởng.
2. Sử dụng thuật toán **K-Means Clustering** để phân cụm lớp thành **3 nhóm **.
3. Mỗi cá nhân nộp:

   * Link source code GitHub
   * Video chạy chương trình
   * Link video YouTube phải được đính kèm trong GitHub README


---

# 1) Tiêu Chí đặt ra 
Xây dựng chương trình phân loại học lực sinh viên theo **GPA hệ 4** thành 3 nhóm:

| Nhóm   | Điều kiện GPA   | Ý nghĩa              |
| ------ | --------------- | -------------------- |
| Nhóm 1 | GPA ≥ 3.2       | Sinh viên học tốt    |
| Nhóm 2 | 2.5 ≤ GPA < 3.2 | Sinh viên ổn định    |
| Nhóm 3 | GPA < 2.5       | Sinh viên cần hỗ trợ |

Ngoài việc phân loại theo ngưỡng GPA, chương trình còn sử dụng thuật toán:

```text
K-Means Clustering (K = 3)
```

để minh hoạ quá trình phân cụm dữ liệu sinh viên theo GPA.

---
#  Cấu trúc thư mục dự án

```text
Kmean-HocLucK58KTP/
├── data/
│   └── TỔNG HỢP ĐIỂM K58KTP.xlsx
│
├── output/
│   ├── nhom_1_sinh_vien_hoc_tot.xlsx
│   ├── nhom_2_sinh_vien_on_dinh.xlsx
│   ├── nhom_3_sinh_vien_can_ho_tro.xlsx
│   └── bieu_do_phan_cum.png
│
├── src/
│   └── kmean.py
│
├── requirements.txt
├── README.md
└── .gitignore
```

---
# 2) Dữ liệu đầu vào

File dữ liệu được đặt tại:

```text
data/TỔNG HỢP ĐIỂM K58KTP.xlsx
```

Chương trình hỗ trợ 2 kiểu dữ liệu Excel:

---

## (A) Dạng bảng GPA

Nếu file có cột:

```text
Điểm TK (4)
```

thì chương trình sẽ đọc trực tiếp GPA của từng sinh viên.

---

## (B) Dạng ma trận điểm môn học

Nếu không tồn tại cột GPA, chương trình sẽ tự động:

* Nhận diện danh sách sinh viên
* Trích xuất điểm từng môn
* Tính GPA trung bình hệ 4

Dữ liệu có cấu trúc:

```text
STT | Mã môn | Tên môn học | Điểm sinh viên
```

Công thức tính:

[
GPA = mean(các\ môn\ hệ\ 4)
]

Các ô trống sẽ được bỏ qua.

---

# 3) Xử lý lỗi dữ liệu Excel 
- Lý do: Khi copy điểm từ web https://portal.tnut.edu.vn/ rồi paste vào Excel, dữ liệu có thể giữ nguyên định dạng cũ, khiến một số điểm bị chuyển từ kiểu Number sang kiểu Date. vậy nên code canf thêm mục xử lý dữ liệu .

<img width="1342" height="660" alt="image" src="https://github.com/user-attachments/assets/7179939c-887f-4f6c-bc3b-e0ad1554100a" />


Trong quá trình nhập liệu, Excel đôi khi tự động chuyển:

```text
3.5 → 03/05/2026
3.7 → 07/03/2026
```

khiến pandas đọc thành kiểu ngày tháng.

Chương trình đã xử lý bằng cách:

* Kiểm tra kiểu dữ liệu DateTime
* Suy ngược lại điểm theo:

  * day/month
  * month/day
* Chỉ chấp nhận giá trị hợp lệ trong khoảng:

```text
0 → 4
```

Ví dụ:

```text
03/05/2026 → 3.5
07/03/2026 → 3.7
```

---

# 4) Thuật toán sử dụng


Ý tưởng chính là: **lấy GPA hệ 4 cho từng sinh viên** (đọc trực tiếp hoặc tự tính từ ma trận điểm), sau đó **chia 3 nhóm theo ngưỡng GPA**.

### 4.1. Các bước xử lý

1. Xác định đường dẫn dự án, đọc file Excel đầu vào trong `data/`.
2. Thử đọc theo **dạng bảng**:
   - Tự dò dòng tiêu đề để tìm cột `Điểm TK (4)`.
   - Nếu có cột này: chuẩn hoá giá trị GPA (đổi dấu phẩy, bỏ ký tự lạ, xử lý lỗi “điểm bị đọc thành ngày”).
3. Nếu **không có cột `Điểm TK (4)`** thì chuyển sang **dạng ma trận**:
   - Tìm hàng có nhãn `MSSV` để xác định các cột sinh viên.
   - Tìm hàng có chữ `Tên` để lấy tên sinh viên.
   - Tìm vùng môn học theo tiêu đề `STT | Mã Môn học | Tên Môn học`.
   - Lấy các dòng môn học (bỏ qua một số mã môn không tính GPA theo cấu hình), rồi đọc ma trận điểm.
   - Chuẩn hoá từng ô điểm về số thực trong [0, 4], sau đó tính GPA cho mỗi sinh viên bằng trung bình các môn (bỏ qua ô trống).
4. Làm tròn GPA 2 chữ số thập phân.
5. (Minh hoạ) Chạy **K-Means (K=3)** trên 1 biến GPA để “phân cụm” dữ liệu.
6. Chia nhóm **theo ngưỡng GPA của đề bài**:
   - Nhóm 1 nếu GPA ≥ 3.2
   - Nhóm 2 nếu 2.5 ≤ GPA < 3.2
   - Nhóm 3 nếu GPA < 2.5
7. Xuất kết quả:
   - 3 file Excel theo từng nhóm vào `output/`.
   - 1 biểu đồ scatter GPA theo STT (nếu có) và vẽ 2 đường ngưỡng 3.2, 2.5.
---
Gần chuẩn rồi, nhưng có vài chỗ nên sửa để đúng ký hiệu toán và dễ thuyết trình hơn:

---

## 4.2 Công thức sử dụng
## 4.2 Công thức sử dụng

### (1) Tính GPA từ ma trận điểm theo môn

Gọi $g_{i,c}$ là điểm hệ 4 của sinh viên $i$ ở môn $c$.

Nếu sinh viên $i$ có $m_i$ môn hợp lệ (không rỗng, không bị loại theo danh sách mã môn), thì:

$$
\text{GPA}_i = \frac{1}{m_i} \sum_{c \in C_i} g_{i,c}
$$

Trong đó:

- $C_i$: tập các môn hợp lệ của sinh viên $i$
- $m_i = |C_i|$: số môn hợp lệ

---

### (2) K-Means (minh hoạ, không dùng để chia nhóm cuối)

Với dữ liệu đầu vào $x_i$ (ở đây $x_i = \text{GPA}_i$), K-Means tìm $k$ tâm cụm $\mu_1,\mu_2,\dots,\mu_k$ sao cho hàm mục tiêu nhỏ nhất:

$$
\min_{\{\mu_j\}} \sum_{i=1}^{n} \min_{j \in \{1,\dots,k\}} \|x_i - \mu_j\|^2
$$

Hai bước lặp cơ bản:

#### • Gán cụm (Assignment)

$$
c_i = \arg\min_{j \in \{1,\dots,k\}} \|x_i - \mu_j\|^2
$$

Trong đó $c_i$ là cụm của sinh viên $i$.

#### • Cập nhật tâm cụm (Update centroid)

$$
\mu_j = \frac{1}{|S_j|} \sum_{i \in S_j} x_i
$$

Trong đó:

$$
S_j = \{ i \mid c_i = j \}
$$

là tập các phần tử thuộc cụm $j$.

---

### (3) Chia nhóm theo ngưỡng (kết quả chính của bài)

$$
\text{Nhóm}(\text{GPA}) =
\begin{cases}
1, & \text{GPA} \ge 3.2 \\
2, & 2.5 \le \text{GPA} < 3.2 \\
3, & \text{GPA} < 2.5
\end{cases}
$$

Trong đó:

- Nhóm 1: Sinh viên học tốt
- Nhóm 2: Sinh viên ổn định
- Nhóm 3: Sinh viên cần hỗ trợ
# 5) Kết quả đầu ra

Sau khi chạy chương trình, thư mục `output/` sẽ được tạo:

```text
output/
```

Bao gồm:

```text
nhom_1_sinh_vien_hoc_tot.xlsx
nhom_2_sinh_vien_on_dinh.xlsx
nhom_3_sinh_vien_can_ho_tro.xlsx
bieu_do_phan_cum.png
```

---



# 7) Cài đặt môi trường

## Bước 1 — Tạo môi trường ảo

```bash
python -m venv venv
```

---

## Bước 2 — Kích hoạt môi trường

### Windows PowerShell

```bash
.\venv\Scripts\Activate
```

---

## Bước 3 — Cài thư viện

```bash
pip install -r requirements.txt
```

---

# 8) Chạy chương trình

```bash
python src/kmean.py
```

Hoặc:

```bash
.\venv\Scripts\Activate; python .\src\kmean.py
```

---

# 9) Kết quả hiển thị

Terminal sẽ hiển thị:

```text
Số sinh viên nhóm học tốt
Số sinh viên nhóm ổn định
Số sinh viên cần hỗ trợ
```

Đồng thời sinh biểu đồ phân cụm GPA.

---




