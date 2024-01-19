# penyata-sesi-hr
Penjana Penyata Akhir Merit Demerit Homeroom.

Sistem Penjana Penyata ini berintegrasi terus dengan Sistem Merit Demerit HR yang sedia ada di MRSM Gemencheh. Apabila file excel yang mengandungi data merit demerit HR diletakkan dalam sebuah folder yang akan dinyatakan kemudian, Sistem Penjana Penyata akan menghasilkan penyata merit demerit berdasarkan data yang sedia ada dalam file excel tersebut.

Sistem Penjana Penyata mula dibangunkan pada April 2023.

## Cara Penggunaan

Sebuah dokumentasi penuh penggunaan boleh didapati di [sini](https://docs.google.com/document/d/1EO1ZJwPavTDKv6M_ybXg-wm25_HHItIb3t_a_2tttqc/edit?usp=sharing). Dokumentasi yang terkandung dalam link tersebut lebih lengkap dan lebih mudah difahami. Dokumentasi yang terkandung dalam README ini hanya ringkasan sahaja.

### Pastikan anda mempunyai Node.js

Pastikan anda mempunyai Node.js versi 12.0.0 atau ke atas. Anda boleh memuat turun Node.js di [sini](https://nodejs.org/en/download/).

### Langkah 1: Muat naik fail Excel ke folder 'src'

Fail Excel mestilah bernama 'Data_Merit_Demerit_HR_TAHUN.xlsx' dan terdapat di dalam folder 'src'.

Contoh Nama Fail: 'Data_Merit_Demerit_HR_2020.xlsx'

### Langkah 2: Jalankan penjana untuk mengambil data dari fail Excel

Buka terminal dan jalankan penjana dengan menaipkan perintah berikut:

```
node index.js 
```

### Langkah 3: Pastikan kewujudan fail template penyata akhir merit dan demerit homeroom

Fail template penyata akhir merit dan demerit homeroom mestilah bernama 'Template_Penyata_Akhir_Merit_Demerit_HR_TAHUN.xlsx' dan terdapat di dalam folder 'src'.

Contoh: 'Template_Penyata_Akhir_Merit_Demerit_HR_2020.xlsx'

### Langkah 4: Jalankan penjana untuk menghasilkan penyata akhir merit dan demerit homeroom

Buka terminal dan jalankan penjana dengan menaipkan perintah berikut:

```
node write.js
```

## Todo

- [ ] Major: Implement testing using Jest.
- [ ] Major: Implement a GUI for the generator.
- [ ] Minor: Implement a more interactive CLI for the generator.
- [ ] Minor: Implement a better way to handle errors.


## Menyumbang kepada Projek

Sumbangan anda amatlah dialu-alukan, sama ada untuk membetulkan sebarang ralat ataupun menginovasikan projek ini kepada yang lebih baik. Anda boleh membantu dengan cara berikut:

1. Fork projek ini.
2. Buat branch baru.
3. Buat perubahan yang dikehendaki.
4. Commit perubahan anda.
5. Push ke branch yang baru di fork anda.
6. Buat pull request.

Nama anda akan disenaraikan di dalam senarai kontributor projek ini. Terima kasih!