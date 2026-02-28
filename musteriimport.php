<?php
header('Content-Type: text/html; charset=utf-8');
error_reporting(E_ALL ^ E_NOTICE); 
ob_start();
session_start();
extract($_GET);
extract($_POST);
ob_flush();
include '../db/dbclass.php'; // İçinde dbClass'ın olduğu varsayılıyor
$dba = new dbClass();
$dba->connect();

// session başlatma fonksiyonun varsa (s_start gibi)
if (function_exists('s_start')) { s_start(); } 

require_once('vendor/excel_reader2.php');
require_once('vendor/SpreadsheetReader.php');

$message = "";
$type = "";

if (isset($_POST["import"])) {
    $allowedFileType = ['application/vnd.ms-excel','text/xls','text/xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
    
    if (in_array($_FILES["file"]["type"], $allowedFileType)) {
        $targetPath = 'uploads/' . time() . '_' . $_FILES['file']['name'];
        move_uploaded_file($_FILES['file']['tmp_name'], $targetPath);
        
        try {
            $Reader = new SpreadsheetReader($targetPath);
            foreach ($Reader->Sheets() as $index => $name) {
                $Reader->ChangeSheet($index);
                $satir = 0;

                foreach ($Reader as $Row) {
                    if ($satir == 0) { $satir++; continue; } // Başlık satırı

                    $musterikod   = $Row[0] ?? "";
                    $musteriunvan = $Row[1] ?? "";
                    $grupkod      = $Row[2] ?? "";
                    $ekgrupkod    = $Row[3] ?? "";
                    $ilgilikisi   = $Row[4] ?? "";
                    $vd           = $Row[5] ?? "";
                    $vn           = $Row[6] ?? "";
                    $telno        = $Row[7] ?? "";
                    $adres        = $Row[8] ?? "";
                    $aciklama     = $Row[9] ?? ""; // Düzeltme: Açıklama genellikle sonraki sütundur

                    if (!empty($musterikod) || !empty($musteriunvan)) {
                        // 1. GÜVENLİ KONTROL (Prepared Statement)
                        $checkSql = "SELECT COUNT(id) as sayi FROM musteri WHERE musterikod = ?";
                        $stmt = $dba->safeQuery($checkSql, [$musterikod]);
                        $sonuc = $dba->fetch_object($stmt);

                        if ($sonuc->sayi > 0) {
                            $message .= "{$musterikod} zaten var, atlandı.<br>";
                        } else {
                            // 2. GÜVENLİ EKLEME
                            $insertSql = "INSERT INTO musteri (musterikod, musteriunvan, grupkod, ekgrupkod, ilgilikisi, vd, vn, telno, adres, aciklama) 
                                          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
                            $params = [$musterikod, $musteriunvan, $grupkod, $ekgrupkod, $ilgilikisi, $vd, $vn, $telno, $adres, $aciklama];
                            
                            if ($dba->safeQuery($insertSql, $params)) {
                                $type = "success";
                                $message .= "{$musterikod} başarıyla eklendi.<br>";
                            }
                        }
                    }
                    $satir++;
                }
            }
        } catch (Exception $e) {
            $type = "error";
            $message = "Hata: " . $e->getMessage();
        }
    } else {
        $type = "error";
        $message = "Geçersiz dosya formatı.";
    }
}
?>

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Müşteri Aktarım</title>
    <style>
        body { font-family: Arial; width: 600px; margin: 20px auto; }
        .outer-container { background: #f0f0f0; padding: 30px; border-radius: 5px; border: 1px solid #ddd; }
        .btn-submit { background: #333; color: #fff; padding: 10px 20px; border: none; cursor: pointer; }
        .success { background: #d4edda; padding: 10px; border: 1px solid #c3e6cb; margin-top: 10px; }
        .error { background: #f8d7da; padding: 10px; border: 1px solid #f5c6cb; margin-top: 10px; }
        .data-table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        .data-table td, .data-table th { border: 1px solid #ddd; padding: 8px; font-size: 12px; }
    </style>
</head>
<body>
    <h2>Müşteri Aktarım</h2>
    <div class="outer-container">
        <form action="" method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xls,.xlsx" required>
            <button type="submit" name="import" class="btn-submit">Yükle</button>
        </form>
    </div>

    <?php if($message): ?>
        <div class="<?= $type ?>"><?= $message ?></div>
    <?php endif; ?>

    <h3>Kayıtlı Müşteriler</h3>
    <table class="data-table">
        <thead>
            <tr><th>Kod</th><th>Unvan</th><th>Şehir/Grup</th></tr>
        </thead>
        <tbody>
            <?php
            $stmt = $dba->query("SELECT * FROM musteri ORDER BY id DESC LIMIT 20");
            while($m = $dba->fetch_object($stmt)): ?>
                <tr>
                    <td><?= htmlspecialchars($m->musterikod) ?></td>
                    <td><?= htmlspecialchars($m->musteriunvan) ?></td>
                    <td><?= htmlspecialchars($m->grupkod) ?></td>
                </tr>
            <?php endwhile; ?>
        </tbody>
    </table>
</body>
</html>