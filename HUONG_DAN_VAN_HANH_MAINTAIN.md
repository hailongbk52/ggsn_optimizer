# HƯỚNG DẪN VẬN HÀNH & BẢO TRÌ HỆ THỐNG GGSN LOAD CALCULATOR

Tài liệu này hướng dẫn chi tiết cách cấu hình, khởi chạy, dừng tiến trình, cập nhật dữ liệu bảng và tùy chỉnh giao diện/công thức tính toán của hệ thống **GGSN Load Calculator** khi triển khai thực tế (Production).

---

## 1. Kiến Trúc Hệ Thống Tổng Quan

Hệ thống được xây dựng theo mô hình client-server tách biệt:

* **Backend**:
  * Công nghệ: Python (FastAPI), SQLite, `openpyxl`.
  * Nhiệm vụ: Cung cấp API lưu trữ dữ liệu cấu hình, truy vấn database trực tiếp, upload file và xuất kết quả báo cáo Excel dựa trên template có sẵn.
  * Database: File SQLite (`ggsn_persistent_store.db`) nằm tại thư mục `backend/`.
* **Frontend**:
  * Công nghệ: React, Vite, TypeScript.
  * Nhiệm vụ: Hiển thị giao diện tương tác trực quan, chỉnh sửa bảng, tính toán các công thức Excel mô phỏng thời gian thực phía Client (qua module `formulaEvaluator.ts`).

---

## 2. Cấu Hình Địa Chỉ IP & Cổng (Port)

### 2.1. Cấu hình Backend (API)

Mặc định trong code phát triển (`backend/main.py`), FastAPI chạy ở địa chỉ `127.0.0.1` cổng `8000`. Khi deploy thực tế, bạn cần mở cấu hình để cho phép truy cập từ mạng nội bộ hoặc Internet.

* **File cấu hình chính**: [backend/main.py](file:///c:/Users/longvh3/Downloads/Tinh_tai_UCTT/ggsn-load-calculator/backend/main.py)
* **Đoạn mã khởi chạy**:

  ```python
  if __name__ == "__main__":
      import uvicorn
      uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=False)
  ```

  * *Lưu ý*: Đổi `host="127.0.0.1"` thành `host="0.0.0.0"` để lắng nghe từ mọi card mạng (IP của server). Tắt chế độ `reload=True` bằng cách chuyển thành `reload=False` để tối ưu hiệu năng chạy thực tế.

### 2.2. Cấu hình Frontend (API Client)

Frontend giao tiếp với API thông qua hằng số `API_BASE`.

* **File cấu hình chính**: [frontend/src/App.tsx](file:///c:/Users/longvh3/Downloads/Tinh_tai_UCTT/ggsn-load-calculator/frontend/src/App.tsx)
* **Dòng cấu hình**:

  ```typescript
  const API_BASE = "http://<IP_SERVER_BACKEND>:8000/api";
  ```

  * Hãy thay đổi `<IP_SERVER_BACKEND>` thành địa chỉ IP tĩnh của máy chủ Backend hoặc domain được cấu hình.
  * *Khuyên dùng*: Sử dụng biến môi trường hoặc cấu hình proxy của Nginx để tránh hardcode IP trực tiếp.

---

## 3. Quy Trình Khởi Chạy & Dừng Tiến Trình (Start/Stop)

### 3.1. Môi trường Development (Phát triển & Thử nghiệm)

Thực hiện chạy bằng terminal thông thường:

* **Khởi chạy Backend**:
  ```powershell
  cd ggsn-load-calculator/backend
  # Cài đặt thư viện nếu chưa có: pip install fastapi uvicorn pydantic openpyxl
  & "C:\Users\longvh3\AppData\Local\Programs\Python\Python312\python.exe" main.py

  #Nếu muốn chạy ngầm
  Start-Process -FilePath "C:\Users\longvh3\AppData\Local\Programs\Python\Python312\python.exe" -ArgumentList "main.py" -WindowStyle Hidden
  ```
* **Khởi chạy Frontend**:
  ```powershell
  cd ggsn-load-calculator/frontend
  # Cài đặt thư viện nếu chưa có: npm install
  npm run dev
  ```

### 3.2. Môi trường Production (Triển khai thực tế)

#### A. Đối với Backend (Python FastAPI)

Nên sử dụng công cụ quản lý tiến trình như **PM2** (yêu cầu cài đặt Node.js) hoặc chạy như một **Systemd Service** (trên Linux) để đảm bảo tự khởi động lại khi server crash hoặc restart.

* **Cách 1: Khởi chạy bằng PM2 (Khuyên dùng cả trên Windows/Linux)**:

  ```bash
  cd ggsn-load-calculator/backend
  # Khởi chạy tiến trình uvicorn dưới sự quản lý của PM2
  pm2 start "uvicorn main:app --host 0.0.0.0 --port 8000" --name "ggsn-backend"

  # Kiểm tra trạng thái
  pm2 status

  # Dừng tiến trình
  pm2 stop ggsn-backend

  # Khởi động lại
  pm2 restart ggsn-backend
  ```
* **Cách 2: Chạy ngầm bằng nohup (Linux đơn giản)**:

  ```bash
  cd ggsn-load-calculator/backend
  nohup uvicorn main:app --host 0.0.0.0 --port 8000 > backend.log 2>&1 &

  # Để dừng tiến trình, tìm PID và kill:
  ps -ef | grep uvicorn
  kill -9 <PID>
  ```

#### B. Đối với Frontend (React Vite)

Không chạy `npm run dev` trên môi trường Production. Phải build code thành file tĩnh và deploy qua Web Server (Nginx, Apache, IIS).

1. **Biên dịch code (Build)**:

   ```bash
   cd ggsn-load-calculator/frontend
   npm run build
   ```

   * Sau khi chạy, một thư mục tên là `dist` sẽ được sinh ra chứa toàn bộ mã nguồn HTML/JS/CSS đã được nén và tối ưu hóa.
2. **Cấu hình Nginx phục vụ Frontend**:
   Sao chép thư mục `dist` lên server và cấu hình Nginx block tương tự như sau:

   ```nginx
   server {
       listen 80;
       server_name ggsn-calculator.internal;

       location / {
           root /var/www/ggsn-load-calculator/frontend/dist;
           index index.html;
           try_files $uri $uri/ /index.html;
       }

       # Cấu hình proxy ngược cho API backend để tránh lỗi CORS
       location /api {
           proxy_pass http://127.0.0.1:8000/api;
           proxy_set_header Host $host;
           proxy_set_header X-Real-IP $remote_addr;
       }
   }
   ```

---

## 4. Cập Nhật Dữ Liệu Bảng & Cơ Sở Dữ Liệu

Hệ thống lưu trữ 5 bảng cấu hình quan trọng trong SQLite:
1.  `table_license`: Thông số license (License Bear, License Throughput, License Bear UCTT, License Throughput VHKT).
2.  `table_current`: Số liệu tải thực tế hiện tại (Bear sử dụng, Throughput, Bear IMS, IPv4 Internet, IPv4 IMS...).
3.  `table_weight`: Trọng số định tuyến (Weight cũ, Weight mới và trạng thái ON/OFF của node).
4.  `table_ims_routing`: Cấu hình mapping IMS Site cho từng GGSN node (phục vụ cột IMS (Q) trong bảng mô phỏng).
5.  `table_hw_site`: Cấu hình mapping vị trí phần cứng HW/NFVI Site cho từng node (phục vụ cột HW/NFVI Site (AM)).

### 4.1. Cập nhật qua giao diện người dùng (UI)
*   Hệ thống có mục chỉnh sửa trực tiếp trên lưới dữ liệu (editable tables) cho cả 5 bảng trên. Sau khi thay đổi thông số, ấn nút **"Lưu dữ liệu"** để hệ thống gọi API `save-table-data` lưu thẳng vào SQLite.
*   **Truy vấn SQL trực tiếp**: Hệ thống cung cấp một bảng điều khiển SQL Query trên giao diện để quản trị viên có thể chạy lệnh `UPDATE`, `INSERT`, `DELETE` thủ công trên cả 5 bảng.

### 4.2. Khởi tạo/Cập nhật cơ sở dữ liệu vật lý
*   **File database**: `backend/ggsn_persistent_store.db`.
*   **Sao lưu (Backup)**: Chỉ cần copy file `.db` này cất vào thư mục lưu trữ an toàn định kỳ.
*   **Thay đổi dữ liệu mặc định**: Nếu cần thay đổi cấu trúc bảng hoặc danh sách node GGSN mặc định ban đầu, chỉnh sửa hàm `init_db()` trong file [backend/main.py](file:///c:/Users/longvh3/Downloads/Tinh_tai_UCTT/ggsn-load-calculator/backend/main.py).

### 4.3. Quản lý Lập lịch SQL Định kỳ (Scheduler)
Hệ thống sử dụng bảng `table_schedules` để lưu trữ các tác vụ chạy SQL tự động cập nhật dữ liệu nguồn:
*   **Cấu trúc bảng**: `table_key TEXT PRIMARY KEY, query TEXT, schedule_type TEXT, is_active INTEGER, last_run TEXT`.
*   **Cơ chế hoạt động**: Khi chạy Backend, tiến trình ngầm `schedule_runner_loop` sẽ kiểm tra mỗi 30 giây. Nếu cấu hình lập lịch của bảng nào được bật (`is_active = 1`) và thời gian kể từ `last_run` vượt quá chu kỳ (`5m` - 5 phút, `30m` - 30 phút, `1h` - 1 tiếng, `24h` - 24 tiếng), backend sẽ tự động thực thi truy vấn đó và cập nhật lại bảng dữ liệu tương ứng.
*   **Điều chỉnh code Scheduler**: Nếu cần thêm các chu kỳ mới hoặc thay đổi logic xử lý ánh xạ cột kết quả SQL, chỉnh sửa hàm `execute_and_apply_query` và `schedule_runner_loop` tại [backend/main.py](file:///c:/Users/longvh3/Downloads/Tinh_tai_UCTT/ggsn-load-calculator/backend/main.py).

### 4.4. Quy tắc Kiểm tra File Excel khi Import
Khi import file `.xlsx` thủ công qua giao diện, hệ thống sẽ tự động đối chiếu các tiêu đề cột trong file với các cột được định nghĩa trước ở cấu trúc bảng tương ứng:
*   **Quy tắc**: File import bắt buộc phải chứa đủ các cột trong danh sách mẫu (không phân biệt chữ hoa, chữ thường và khoảng trắng thừa).
*   **Cảnh báo**: Nếu thiếu bất kỳ cột nào, Frontend sẽ chặn hành động import và hiển thị thông báo lỗi chi tiết dạng: `File import sai cấu trúc. Thiếu cột: [Tên cột]`.

---

## 5. Thay Đổi Công Thức & Hàm Tính Toán

Toàn bộ logic tính toán công thức Excel được xử lý phía Client để đảm bảo tốc độ phản hồi nhanh tức thì khi người dùng thay đổi trọng số hoặc thông số tải.

* **File quản lý logic**: [frontend/src/formulaEvaluator.ts](file:///c:/Users/longvh3/Downloads/Tinh_tai_UCTT/ggsn-load-calculator/frontend/src/formulaEvaluator.ts)
* **Cách thức bổ sung hàm Excel mới**:
  * Hàm `evaluateFormula` thực hiện bóc tách và phân tích chuỗi công thức Excel (như `SUM`, `SUMIF`, `VLOOKUP`, `IFERROR`, `IFNA`).
  * Nếu trong file template Excel mới của bạn có thêm các hàm Excel khác (ví dụ: `AVERAGE`, `ROUND`, `AND`, `OR`), bạn cần khai báo các hàm phân tích cú pháp tương tự như `resolveSums` hoặc `resolveVlookups` và tích hợp vào chuỗi xử lý của `evaluateFormula`.
* *Lưu ý*: Mọi chỉnh sửa ở file này cần phải chạy lệnh `npm run build` lại ở frontend để các thay đổi có hiệu lực trên môi trường production.

---

## 6. Cập Nhật Giao Diện Người Dùng (UI) & Thuộc Tính Bảng

* **Giao diện chính**: [frontend/src/App.tsx](file:///c:/Users/longvh3/Downloads/Tinh_tai_UCTT/ggsn-load-calculator/frontend/src/App.tsx)
* **Các thành phần UI mới**:
  * **Thu phóng (Zoom)**: Hệ thống cho phép điều chỉnh tỷ lệ bảng mô phỏng qua thanh công cụ **Thu Phóng** (`zoomScale` state). Tác động trực tiếp lên CSS `fontSize` của thẻ `<table>` từ 75% đến 150%.
  * **Bộ lọc cột (Filter)**: Hàng lọc dưới tiêu đề cột truyền từ khóa vào `columnFilters` state, dùng để sàng lọc động dữ liệu hiển thị (`displayRows`) mà không làm ảnh hưởng đến dữ liệu gốc trong SQLite.
  * **Toàn màn hình (Full View)**: Nút **"Toàn màn hình"** (icon Maximize) ở góc phải toolbar bảng mô phỏng cho phép phóng to bảng thành overlay cố định chiếm toàn bộ viewport. Khi kích hoạt (`isFullView = true`):
    * Khung bảng chuyển sang `position: fixed; inset: 16px; z-index: 50` — tức là chiếm gần như toàn màn hình.
    * Chiều cao bảng mở rộng từ `600px` lên `calc(100vh - 160px)` để tận dụng tối đa không gian màn hình.
    * Backdrop mờ đen được phủ lên phần còn lại của trang để tập trung vào bảng.
    * **Cách thoát**: Nhấn nút **"Thu nhỏ"**, bấm phím `Escape`, hoặc click vào vùng tối bên ngoài bảng.
* **CSS Styles**: [frontend/src/index.css](file:///c:/Users/longvh3/Downloads/Tinh_tai_UCTT/ggsn-load-calculator/frontend/src/index.css) và [frontend/src/App.css](file:///c:/Users/longvh3/Downloads/Tinh_tai_UCTT/ggsn-load-calculator/frontend/src/App.css).
* **Quy trình cập nhật**:
  1. Chỉnh sửa giao diện, nhãn hiển thị hoặc thứ tự các cột trong `App.tsx`.
  2. Kiểm tra cục bộ bằng cách chạy `npm run dev` để kiểm tra độ tương thích và hiển thị trực quan.
  3. Build lại frontend bằng `npm run build`.
  4. Copy đè thư mục `dist` mới lên thư mục lưu trữ tĩnh của Web Server (Nginx/IIS).
