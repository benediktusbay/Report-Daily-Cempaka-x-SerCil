# Daily Sales Report V2 — Team Dashboard

Dashboard Flask multi-user untuk upload Billing Detail harian dan melihat Sales Achievement, BO/QVO, Speed Distribution, dan KPI SKU.

## Logika KPI
- Device = Mobile Phones + Tablet
- Macbook = Computer
- ACC = Mobile Accessories + Computer Accessories + Tablets Accessories + Wearable + Audio
- BO = dealer memiliki pembelian Device dan/atau Macbook > Rp0
- QVO = Device + Macbook + ACC >= Rp46.000.000
- Default BO Target = 25 / salesman
- Speed Distribution cumulative: W1 19, W2 23, W3 25, W4 25 BO
- KPI SKU: 13 BO (2–3 SKU), 6 BO (4–6), 5 BO (7–10), 2 BO (>10)
- SKU Device memakai Article Description dengan model/storage dan mengabaikan warna umum
- Upload berikutnya menambah data; baris yang sama dideduplikasi dengan hash

## Jalankan lokal
```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
python app.py
```
Buka `http://localhost:5000`.

Default login pertama: `admin` / `admin123`. Untuk production, wajib ubah melalui environment variables `ADMIN_USERNAME`, `ADMIN_PASSWORD`, dan `SECRET_KEY` sebelum database dibuat.

## Agar bisa diakses tim
Deploy folder ini ke Render/Railway/Fly.io/VPS. Gunakan PostgreSQL lewat `DATABASE_URL` agar data terpusat dan persisten. Setelah online, Admin membuat akun anggota tim melalui menu **Users**.

## Format upload
Workbook harus memiliki sheet `Export` dan kolom:
- Billing Date
- Salesman Name
- Sold to Party Code
- Sold to Party Name
- Item Group Desc
- Article Description
- Quantity
- Total Nett Amount with Tax

## Catatan target omzet
Target Device/Macbook/ACC dibuat editable per bulan dan per salesman dari menu **Targets** karena nominal target tidak ditentukan pada spesifikasi awal.
