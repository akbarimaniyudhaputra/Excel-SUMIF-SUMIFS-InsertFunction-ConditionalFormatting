# Excel-SUMIF-SUMIFS-InsertFunction-ConditionalFormatting


## SUM-IF melalui Insert Function
 - Digunakan untuk menggali informasi yang ada di data dengan satu acuan
 - Insert Function kelebihannya dapat mencegah error pada rumus atau fungsi karena hasilnya langsung diketahui sebelum di Klik OK.
 - Kita akan menjawab pertanyaan siapa mitra yang memiliki kontribusi penjualan terbesar?
 - Jadi acuannya satu yaitu mitra


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

- 4.) Ada cara lain dalam penggunaan fungsi atau rumus dalam excel, salah satunya yaitu dengan tombol/klik icon/lambang/logo insert function. Kemudian muncul tampilan jendela baru insert function, kita akan menggunakan SUM IF jadi pada tulisan Or select a category, klik pilih Math & Trig, jika kita sering menggunakan SUM IF maka SUM IF juga terdapat di Most Recently Used, kemudian klik OK.
###
![4](https://user-images.githubusercontent.com/86678205/165497787-19597e19-8db5-422f-8be1-e7322d86cf7a.png)

- 5.) Setelah klik OK, maka akan muncul jendela baru tampilan Fuction Arguments seperti gambar di bawah ini. 
###
![5](https://user-images.githubusercontent.com/86678205/165497814-0d6bf26c-c2d3-43c6-9d42-5448bb4e7ad0.png)

```http
Kemudian market kita ada 3 yaitu PT. ABC, PT.XYZ, dan PT.QAZ berarti Range yang akan dipakai yaitu Mitra sehingga blok seluruh baris cell mitra dan jangan lupa diberi nilai absolut ($) (blok & tekan f4 untuk memberi nilai absolut). 
```

```http
Kemudian Criteria diisi K4 karena yang akan kita cari yaitu cell PT.ABC terlebih dahulu dan jangan lupa tidak usah diberi nilai absolut agar dapat ditarik kebawah nantinya untuk mempercepat pencarian di cell PT.XYZ & PT.QAZ. 
```

```http
Setelah itu Sum_range diisi blok seluruh baris yang ada di cell total karena kita mencari total penjualan & jangan lupa diberi nilai absolut ($) (blok & tekan f4 untuk memberi nilai absolut).
```

- 6.) Setelah klik OK, maka akan muncul hasilnya dan tarik ke bawah untuk mengetahui hasil semua cell. 
###
![6](https://user-images.githubusercontent.com/86678205/165497839-dc900108-bd0d-4599-911b-c782a2ca9b93.PNG)

Jadi total penjualan terbaik mitra per tahun yaitu PT.QAZ, disusul PT.QAZ, dan terakhir PT.ABC.

## SUM-IFS melalui Insert Function
 - Digunakan untuk menggali informasi yang ada di data dengan lebih dari satu acuan
 - Insert Function kelebihannya dapat mencegah error pada rumus atau fungsi karena hasilnya langsung diketahui sebelum di Klik OK.
 - Kita akan menjawab pertanyaan siapa mitra yang memiliki kontribusi penjualan terbesar pada setiap lokasi? Mencari total penjualan di setiap lokasi pada setiap mitra? 
 - Acuannya disini ada 2 yaitu lokasi & mitra

#### Dengan cara sebagai berikut :
- 1.) Langkah pertama yaitu tambahkan tabel seperti gambar di bawah ini.
###
![1](https://user-images.githubusercontent.com/86678205/167100286-24cbde95-2302-448f-8848-66127d3448d9.png)

- 2.) Kemudian kita akan menggunakan fungsi/rumus SUM IFS karena dalam kasus ini mempunyai 2 parameter yaitu mitra & Lokasi. Saya menggunakan insert function (tidak menuliskan rumus langsung di cell).
###
![2 0](https://user-images.githubusercontent.com/86678205/167100337-2d89427c-2381-4479-b198-a75ace1d653b.png)

- 3.) Perbedaan SUM IF & SUM IFS, SUM IFS adalah sum_range ada di awal atau depan. Kalau SUM IF adalah sum_range nya di akhir atau belakang. Dalam kasus ini menggunakan SUM IFS, maka 
```http
  Sum_range : diisi/blok seluruh baris/cell di kolom total, jangan lupa diabsolutkan ($) (blok kemudian tekan f4)
```
```http
  Criteria_range1 : diisi/blok seluruh baris/cell di kolom mitra jangan lupa diabsolutkan ($) (blok kemudian tekan f4)
```
```http
  Criteria1 : diisi/blok satu saja mitra (PT.ABC, PT.XYZ, PT.QAZ) (blok cell k4, k5, k6) yang berubah hanya angkanya maka yang diabsolutkan    hurufnya/K-nya saja
```
```http
  Criteria_range2 : diisi/blok seluruh baris/cell di kolom lokasi jangan lupa diabsolutkan ($) (blok kemudian tekan f4)
```
```http
  Criteria2 : diisi/blok satu saja lokasi (Jabodetabek, Jateng, Jatim) (blok cell m3, n3, o3) yang berubah hanya hurufnya maka yang diabsolutkan angkanya/3-nya saja 
```
###
![2 1](https://user-images.githubusercontent.com/86678205/167100379-5308e2dd-cca1-4d7d-81f7-261a2e0698f4.png)

- 4.) Menambahkan conditional formatting color scales, untuk mempercantik & mempermudah dalam menganalisis data maka kita tambahkan conditional formatting dengan cara blok seluruh baris/cell di kolom total penjualan > klik home > conditional formatting > silahkan pilih yang diinginkan (disini menggunakan color scales).
###
![3 0](https://user-images.githubusercontent.com/86678205/167100425-2ef05418-8030-466f-b31e-2fd78327f6b8.png)

Jika tadi pada total penjualan, maka yang selanjutnya lakukan langkah yang sama pada data penjualan per lokasi, jadi bloknya ke kanan per market. Sehingga hasilnya seperti ini.
###
![3 1](https://user-images.githubusercontent.com/86678205/167100473-5b39e1ff-7d2b-411e-9f3b-ce4b002325ad.png)

Jadi informasi yang didapatkan yaitu: 
- Pada mitra PT.ABC urutan penjualan terbesarnya di Jabodetabek, Jateng, Jatim
- Pada mitra PT.XYZ urutan penjualan terbesarnya di Jatim, Jabodetabek, Jateng
- Pada mitra PT.QAZ urutan penjualan terbesarnya di Jateng, Jatim, Jabodetabek

- Di lokasi Jabodetabek dimenangkan oleh PT.QAZ, padahal di PT.QAZ Jabodetabek adalah total penjualan terendah
- Di lokasi Jateng dimenangkan oleh PT.QAZ 
- Di lokasi Jatim dimenangkan oleh PT.XYZ
###
![3 2](https://user-images.githubusercontent.com/86678205/167100496-7ce41710-f85f-4c0f-9b37-9cd9f9a0edf9.PNG)
