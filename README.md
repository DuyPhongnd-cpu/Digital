# 🖥️ HỆ THỐNG QUẢN LÝ VẬN HÀNH NOC VÀ BẢO HÀNH SẢN PHẨM VTC DIGITAL

Ứng dụng Dashboard Streamlit phục vụ quản lý ca trực NOC, đối soát sự cố truyền hình, giám sát hạ tầng máy chủ / mạng / điều hòa và quản lý bảo hành sản phẩm.

---

## 🚀 HƯỚNG DẪN TRIỂN KHAI TRÊN COOLIFY VỚI DOCKER

### Cách 1: Triển khai trực tiếp từ GitHub Repository (Khuyên dùng)
1. **Tạo Repository trên GitHub:**
   - Tạo một Repository mới trên GitHub (ví dụ: `vtc-noc-management`).
   - Đẩy toàn bộ mã nguồn lên GitHub (xem các lệnh bên dưới).

2. **Cấu hình trên giao diện Coolify:**
   - Đăng nhập vào Coolify Dashboard.
   - Chọn **Projects** -> Chọn hoặc tạo **Environment**.
   - Bấm **+ New Resource** -> Chọn **Public Repository** (hoặc **Private Repository** nếu cần kết nối GitHub App/Deploy Key).
   - Dán URL Repository GitHub vào (ví dụ: `https://github.com/<username>/vtc-noc-management`).
   - Chọn nhánh (Branch): `main` hoặc `master`.
   - **Build Pack:** Chọn `Dockerfile` (hoặc `Docker Compose`).
   - **Port Configuration:** Đặt cổng **`8501`**.
   - Cấu hình Domain / FQDN (ví dụ: `https://noc.yourdomain.com`).
   - Bấm **Deploy**.

---

### Cách 2: Chạy cục bộ bằng Docker / Docker Compose

```bash
# Chạy với docker-compose
docker compose up -d --build

# Hoặc build và run Dockerfile thủ công
docker build -t vtc-noc-app .
docker run -d -p 8501:8501 --name vtc-noc-app vtc-noc-app
```
Sau đó truy cập qua trình duyệt tại: `http://localhost:8501`

---

## 📌 CÁC LỆNH GIT ĐẨY LÊN GITHUB

```bash
git init
git add app.py requirements.txt Dockerfile docker-compose.yml .dockerignore .gitignore README.md
git commit -m "feat: init VTC NOC & Warranty management app with Docker support"
git branch -M main
git remote add origin https://github.com/<YOUR_USERNAME>/<YOUR_REPO_NAME>.git
git push -u origin main
```
