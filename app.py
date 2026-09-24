Gunakan file TERBARU sebagai source of truth:

- app.py
- templates/dashboard.html

Billing histori Agustus dan September SUDAH ADA di database/dashboard.
Jangan meminta file billing tambahan.

SKU Penetration juga SUDAH DIPERBAIKI.
JANGAN ubah lagi logic SKU.

Kerjakan HANYA revisi berikut.

\==================================================

1. FORMAT ANGKA / DESIMAL KONSISTEN
   \==================================================

Rapikan formatting seluruh Dashboard.

Percentage gunakan format Indonesia:

59,9%
73,1%

Bukan:

59.9%

Rupiah short format harus konsisten:

Rp20,9 M
Rp751,3 Jt
Rp297 Rb

Gunakan:
M
Jt
Rb

Jangan campur:
juta
Jt

\==================================================
2\. NEGATIVE RUPIAH FORMAT
==========================

Semua negative Rupiah gunakan:

-Rp4,48 M
-Rp751,3 Jt

JANGAN:

Rp-4,48 M

Gunakan formatter/helper existing sebanyak mungkin agar konsisten.

\==================================================
3\. TERMINOLOGY CONSISTENCY
===========================

Normalisasi user-facing terminology utama:

Timegone / Time Gone
→ Time Gone

Macbook / MacBook
→ MacBook

Pada KPI/card/heading utama:

ACC / Accessories
→ Accessories

ACC boleh tetap dipakai di area/table yang sangat compact jika diperlukan.

\==================================================
4\. TARGET DAILY
================

JANGAN ubah.

Pertahankan wording:

Target Daily

Pertahankan calculation existing.

\==================================================
5\. DATA FRESHNESS DI BAGIAN ATAS
=================================

Tambahkan informasi data freshness compact di area atas Dashboard.

Contoh:

Data s/d 23 Sep 2026 · Updated 10:37 WIB

Gunakan data dynamic existing.

"Data s/d":
gunakan latest billing date pada active period.

"Updated":
hanya tampilkan timestamp jika timestamp yang benar memang tersedia.

Jika timestamp tidak tersedia:
cukup tampilkan:

Data s/d 23 Sep 2026

Jangan hardcode tanggal/jam.

Jangan membuat card besar.
Cukup badge/subtext compact.

\==================================================
6\. MTD VS LAST MONTH — 4 KPI UTAMA
===================================

Tambahkan automatic comparison ke:

- Achievement
- Device
- MacBook
- Accessories

Gunakan SAME-PERIOD MTD.

Contoh active period:

latest billing = 23 Sep 2026

maka compare:

1–23 Sep 2026
VS
1–23 Aug 2026

Data Agustus dan September sudah ada di database.
Gunakan histori billing existing.

JANGAN compare current MTD dengan full previous month.

\==================================================
7\. MTD GROWTH CALCULATION
==========================

Formula:

(Current MTD - Previous Same-Period MTD)
/
Previous Same-Period MTD

- 100

Gunakan category mapping existing untuk:

- Total/Achievement
- Device
- MacBook
- Accessories

Jangan membuat category definition baru.

Handle:

- positive growth
- negative growth
- previous = 0
- no previous month data
- January → December tahun sebelumnya

Jika previous = 0 / tidak ada data:

—
MTD vs Last Month

Jangan divide by zero.

\==================================================
8\. KPI CARD DISPLAY
====================

Tambahkan comparison di BAGIAN PALING BAWAH setiap KPI card existing.

Contoh:

↑ +14,6%
MTD vs Last Month (Rp17,9 M)

Jika turun:

↓ -6,2%
MTD vs Last Month (Rp356 Jt)

Nominal dalam kurung adalah value same-period MTD bulan sebelumnya.

Semua 4 card harus konsisten.

JANGAN membuat card MTD baru.

JANGAN redesign KPI existing.

Mobile harus tetap compact.

\==================================================
9\. DEALER BELUM BELANJA — TOTAL TARGET BELUM TERGARAP
======================================================

Tambahkan aggregate total target pada card Dealer Belum Belanja.

Contoh:

3 Dealer · Rp745 Jt Target Belum Tergarap

Nilai harus dynamic.

Calculation:

sum target dealer yang masuk population Dealer Belum Belanja
dalam current dashboard scope.

Harus mengikuti:

- month
- depo
- salesman
- mobileSalesmanFocus existing

Jangan hardcode Rp745 Jt.

Main card tetap clean.

\==================================================
10\. DEALER BELUM BELANJA — LAST PURCHASE
=========================================

Pada:

Lihat Semua

tambahkan:

Last Purchase

atau:

Pembelian Terakhir

untuk setiap dealer.

Gunakan historical billing yang SUDAH ADA di database.

Tampilkan tanggal transaksi terakhir dealer.

Jika tidak ada histori:

—

Desktop:
boleh menjadi kolom tambahan.

Mobile:
masukkan compact di dealer card detail.

Main card Dealer Belum Belanja tidak perlu menampilkan Last Purchase.

\==================================================
11\. PERFORMANCE AUDIT — JANGAN REFACTOR BESAR
==============================================

Audit apakah Dealer Achievement/detail dealer memuat terlalu banyak duplicate DOM
karena desktop table + mobile cards.

JANGAN langsung refactor besar.

Jika bisa dioptimalkan dengan perubahan aman:
boleh lakukan.

Jika berisiko merusak:

- filtering
- salesman sync
- Download JPG
- popup SKU
- mobile/desktop behavior

maka JANGAN ubah.

Cukup laporkan temuan dan rekomendasi.

Functional safety lebih penting daripada optimization.

\==================================================
12\. DO NOT TOUCH
=================

JANGAN ubah:

- SKU Penetration logic — SUDAH FIX
- BO
- QVO
- Speed Distribution
- Target Daily
- Time Gone calculation
- target
- salesman mapping
- QVO Potential logic
- Dealer Achievement visual
- Stock
- Pricelist
- Program
- Incentive
- PJP
- Target Master
- Users
- auth
- navigation
- logo

\==================================================
13\. TEST MTD VS LAST MONTH
===========================

Gunakan database existing.

Untuk September dengan latest billing 23 Sep:

compare:

1–23 Sep
vs
1–23 Aug

Test keempat KPI:

Achievement
Device
MacBook
Accessories

Pastikan cutoff mengikuti latest billing date,
bukan today's date.

\==================================================
14\. MOBILE TEST
================

Test:

390px
414px
430px

Pastikan:

- MTD comparison tidak overflow
- 4 KPI card tetap aligned
- Dealer Belum Belanja tetap readable
- Last Purchase detail readable
- tidak ada horizontal page overflow baru

\==================================================
15\. FINAL REPORT
=================

Setelah selesai laporkan:

1. file yang berubah
2. formatting yang dinormalisasi
3. implementasi Data s/d
4. hasil MTD vs Last Month untuk 4 KPI
5. current vs previous values
6. implementasi total target Dealer Belum Belanja
7. implementasi Last Purchase
8. hasil performance audit
9. mobile regression
10. desktop regression
11. syntax checks
12. konfirmasi Target Daily tidak berubah
13. konfirmasi SKU Penetration tidak diubah
14. konfirmasi business logic lain tidak berubah
