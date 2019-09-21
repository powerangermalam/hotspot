<?php
//koneksi ke database, username,password  dan namadatabase menyesuaikan
$koneksi = mysqli_connect("localhost","root","","pegawai1");
mysql_select_db('pegawai1');
//memanggil file excel_reader
require "excel_reader.php";
//jika tombol import ditekan
if(isset($_POST['submit'])){
    $target = basename($_FILES['filepegawaiall']['name']) ;
    move_uploaded_file($_FILES['filepegawaiall']['tmp_name'], $target);
// tambahkan baris berikut untuk mencegah error is not readable
    chmod($_FILES['filepegawaiall']['name'],0777);
    $data = new Spreadsheet_Excel_Reader($_FILES['filepegawaiall']['name'],false);
//    menghitung jumlah baris file xls
    $baris = $data->rowcount($sheet_index=0);
//    jika kosongkan data dicentang jalankan kode berikut
    $drop = isset( $_POST["drop"] ) ? $_POST["drop"] : 0 ;
    if($drop == 1){
//             kosongkan tabel pegawai
             $truncate ="TRUNCATE TABLE pegawai";
             mysql_query($truncate);
    };
//    import data excel mulai baris ke-2 (karena tabel xls ada header pada baris 1)
    for ($i=2; $i<=$baris; $i++)
    {
//       membaca data (kolom ke-1 sd terakhir)
      $nama           = $data->val($i, 1);
      $tempat_lahir   = $data->val($i, 2);
      $tanggal_lahir  = $data->val($i, 3);
//      setelah data dibaca, masukkan ke tabel pegawai sql
      $query = "INSERT into pegawai1 (nama,tempat_lahir,tanggal_lahir)values('$nama','$tempat_lahir','$tanggal_lahir')";
      $hasil = mysql_query($query);
    }
    if(!$hasil){
//          jika import gagal
          die(mysql_error());
      }else{
//          jika impor berhasil
          echo "Data berhasil diimpor.";
    }
//    hapus file xls yang udah dibaca
    unlink($_FILES['filepegawaiall']['name']);
}
?>
<form name="myForm" id="myForm" onSubmit="return validateForm()" action="import.php" method="post" enctype="multipart/form-data">
    <input type="file" id="filepegawaiall" name="filepegawaiall" />
    <input type="submit" name="submit" value="Import" /><br/>
    <label><input type="checkbox" name="drop" value="1" /> <u>Kosongkan tabel sql terlebih dahulu.</u> </label>
</form>
<script type="text/javascript">
//    validasi form (hanya file .xls yang diijinkan)
    function validateForm()
    {
        function hasExtension(inputID, exts) {
            var fileName = document.getElementById(inputID).value;
            return (new RegExp('(' + exts.join('|').replace(/\./g, '\\.') + ')$')).test(fileName);
        }
        if(!hasExtension('filepegawaiall', ['.xls'])){
            alert("Hanya file XLS (Excel 2003) yang diijinkan.");
            return false;
        }
    }
</script>
