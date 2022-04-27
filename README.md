# Excel-SUMIF-InsertFunction


## SUM-IF melalui Insert Function
 - Digunakan untuk menggali informasi yang ada di data dengan satu acuan
 - Insert Function kelebihannya dapat mencegah error pada rumus atau fungsi karena hasilnya langsung diketahui sebelum di Klik OK.
 - Kita akan menjawab pertanyaan siapa mitra yang memiliki kontribusi penjualan terbesar?
 - Jadi acuannya satu yaitu total penjualan


#### Dengan cara sebagai berikut :
- 1.) Pertama kita copy lagi worksheet penjualan dengan cara klik kiri pada nama worksheet, kemudian klik “move or copy...”
###
![1](https://user-images.githubusercontent.com/86678205/165497699-3c92f0ad-745d-48b8-a828-d1432825f261.png)

- 2.) Setelah itu muncul user interface atau jendela baru tampilan seperti gambar di bawah, kemudian klik penjualan, dan centang create a copy, lalu klik OK.
###
![2](https://user-images.githubusercontent.com/86678205/165497732-deeb8887-bb43-40fc-83a6-cfd0356c096c.png)

- 3.) Maka akan muncul tambahan worksheet dengan data penjualan (3), kemudian klik worksheet tersebut. Setelah itu untuk menjawab pertanyaannya maka kita buat tabel baru seperti berikut.
###
![3](https://user-images.githubusercontent.com/86678205/165497759-a2bc8e37-71ea-46ef-8b17-44d53b21ddfd.png)

- 4.) Ada cara lain dalam penggunaan fungsi atau rumus dalam excel, salah satunya yaitu dengan tombol/klik icon/lambing/logo insert function. Kemudian muncul tampilan jendela baru insert function, kita akan menggunakan SUM IF jadi pada tulisan Or select a category, klik pilih Math & Trig, jika kita sering menggunakan SUM IF maka SUM IF juga terdapat di Most Recently Used, kemudian klik OK.
###
![4](https://user-images.githubusercontent.com/86678205/165497787-19597e19-8db5-422f-8be1-e7322d86cf7a.png)

- 5.) Setelah klik OK, maka akan muncul jendela baru tampilan Fuction Arguments seperti gambar di bawah ini. 
###
![5](https://user-images.githubusercontent.com/86678205/165497814-0d6bf26c-c2d3-43c6-9d42-5448bb4e7ad0.png)

Kemudian market kita ada 3 yaitu PT. ABC, PT.XYZ, dan PT.QAZ berarti Range yang akan dipakai yaitu Mitra sehingga blok seluruh baris cell mitra dan jangan lupa diberi nilai absolut ($) (blok & tekan f4 untuk memberi nilai absolut). Kemudian Criteria diisi K4 karena yang akan kita cari yaitu cell PT.ABC terlebih dahulu dan jangan lupa tidak usah diberi nilai absolut agar dapat ditarik kebawah nantinya untuk mempercepat pencarian di cell PT.XYZ & PT.QAZ. Setelah itu Sum_range diisi blok seluruh baris yang ada di cell total karena kita mencari total penjualan & jangan lupa diberi nilai absolut ($) (blok & tekan f4 untuk memberi nilai absolut).

- 6.) Setelah klik OK, maka akan muncul hasilnya dan tarik ke bawah untuk mengetahui hasil semua cell. 
###
![6](https://user-images.githubusercontent.com/86678205/165497839-dc900108-bd0d-4599-911b-c782a2ca9b93.PNG)

Jadi total penjualan terbaik mitra per tahun yaitu PT.QAZ, disusul PT.QAZ, dan terakhir PT.ABC.
