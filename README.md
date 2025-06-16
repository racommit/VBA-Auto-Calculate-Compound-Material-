# VBA-Auto-Calculate-Compound-Material-

## Fitur Tambahan

Repositori ini kini memiliki sub `permanenkan_terakhir` pada *CoreModule*.
Sub tersebut dapat dipanggil (misalnya melalui tombol) setelah proses submit
berhasil untuk menghapus seluruh histori undo dari `HISTORY_UNDO` yang
terkait dengan Action ID terakhir. Gunakan fitur ini apabila perubahan sudah
dianggap final dan tidak ingin dapat di-*undo* kembali.
## Logging dan Notifikasi

Khusus pesan informasi, sekarang menggunakan helper `ShowInfo` yang akan
menampilkan `MsgBox` hanya jika konstanta `SHOW_INFO` bernilai `True`.
Pesan debug telah dialihkan ke `DebugLog` yang dapat diaktifkan dengan
mengubah konstanta `DEBUG_MODE` pada `DebugModule`.

## Validasi Input Persentase

Semua prosedur penggantian material kini memeriksa nilai persentase yang
diberikan. Jika pengguna memasukkan nilai di luar rentang 0–100%, proses
akan dibatalkan dan pesan kesalahan ditampilkan.
