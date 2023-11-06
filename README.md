# penyata-sesi-hr
Penjana Penyata Akhir Merit Demerit Homeroom

## Cara Penggunaan

Terdapat beberapa langkah untuk menghasilkan penyata akhir merit dan demerit homeroom menggunakan aplikasi ini. 

### Pastikan anda mempunyai Node.js

Pastikan anda mempunyai Node.js versi 12.0.0 atau ke atas. Anda boleh memuat turun Node.js di [sini](https://nodejs.org/en/download/).

### Langkah 1: Muat naik fail Excel ke folder 'src'

Pastikan fail Excel yang dimuat naik mengandungi data yang betul. Fail Excel mestilah bernama 'Data_Merit_Demerit_HR_TAHUN.xlsx'

Contoh: 'Data_Merit_Demerit_HR_2020.xlsx'

### Langkah 2: Jalankan aplikasi untuk mengambil data dari fail Excel

Buka terminal dan jalankan aplikasi dengan menaipkan perintah berikut:

```
node index.js Data_Merit_Demerit_HR_TAHUN.xlsx
```

Contoh: 
```
node index.js Data_Merit_Demerit_HR_2020.xlsx
```

### Langkah 3: Muat naik fail template penyata akhir merit dan demerit homeroom

Pastikan fail template penyata akhir merit dan demerit homeroom yang dimuat naik mengandungi data yang betul. Fail template penyata akhir merit dan demerit homeroom mestilah bernama 'Template_Penyata_Akhir_Merit_Demerit_HR_TAHUN.xlsx'

Contoh: 'Template_Penyata_Akhir_Merit_Demerit_HR_2020.xlsx'

### Langkah 4: Jalankan aplikasi untuk menghasilkan penyata akhir merit dan demerit homeroom

Buka terminal dan jalankan aplikasi dengan menaipkan perintah berikut:

```
node index.js Template_Penyata_Akhir_Merit_Demerit_HR_TAHUN.xlsx
```

Contoh: 
```
node index.js Template_Penyata_Akhir_Merit_Demerit_HR_2020.xlsx
```