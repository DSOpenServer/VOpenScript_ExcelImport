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
//s_start();

require_once('vendor/excel_reader2.php');
require_once('vendor/SpreadsheetReader.php');

$type    = '';
$message = '';

if (isset($_POST['import'])) {

    $allowedFileType = [
        'application/vnd.ms-excel',
        'text/xls',
        'text/xlsx',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ];

    $fileType = $_FILES['file']['type'] ?? '';

    if (in_array($fileType, $allowedFileType, true)) {

        $uploadDir  = 'uploads/';
        $fileName   = basename($_FILES['file']['name']);
        $targetPath = $uploadDir . $fileName;

        if (!move_uploaded_file($_FILES['file']['tmp_name'], $targetPath)) {
            $type    = 'error';
            $message = 'Dosya yüklenemedi. Lütfen tekrar deneyin.<br>';
        } else {
            $Reader     = new SpreadsheetReader($targetPath);
            $sheetCount = count($Reader->sheets());

            for ($i = 0; $i < $sheetCount; $i++) {
                $Reader->ChangeSheet($i);
                $satir = 0;

                foreach ($Reader as $Row) {
                    // Başlık satırını atla
                    if ($satir === 0) {
                        $satir++;
                        continue;
                    }

                    $stokkod    = trim((string)($Row[0] ?? ''));
                    $stokad     = trim((string)($Row[1] ?? ''));
                    $grupkod    = trim((string)($Row[2] ?? ''));
                    $ekgrupkod  = trim((string)($Row[3] ?? ''));
                    $birim      = trim((string)($Row[4] ?? ''));
                    $kdvoran    = trim((string)($Row[5] ?? ''));
                    $aciklama   = trim((string)($Row[6] ?? ''));
                    $alisfiyat  = trim((string)($Row[7] ?? ''));
                    $satisfiyat = trim((string)($Row[8] ?? ''));

                    if (empty($stokkod) && empty($stokad)) {
                        $satir++;
                        continue;
                    }

                    // Ürün zaten var mı? — prepared statement
                    $kontrolStmt = $dba->safeQuery(
                        "SELECT COUNT(id) AS sayi FROM stok WHERE stokkod = ?",
                        [$stokkod]
                    );
                    $sonuc = $dba->fetch_object($kontrolStmt);

                    if ((int)$sonuc->sayi > 0) {
                        $message .= htmlspecialchars($stokkod, ENT_QUOTES, 'UTF-8')
                            . ' Kodlu ürün sistemde zaten var! Bu yüzden eklenemedi.<br>';
                    } else {
                        // Ekle — prepared statement
                        $dba->safeQuery(
                            "INSERT INTO stok
                             (stokkod, stokad, grupkod, ekgrupkod, birim, kdvoran, aciklama, alisfiyat, satisfiyat)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                            [$stokkod, $stokad, $grupkod, $ekgrupkod, $birim, $kdvoran, $aciklama, $alisfiyat, $satisfiyat]
                        );

                        $insertId = $dba->insert_id();

                        if (!empty($insertId)) {
                            $type     = 'success';
                            $message .= htmlspecialchars($stokkod, ENT_QUOTES, 'UTF-8')
                                . ' koduyla stok tablosuna kaydedildi.<br>';
                        } else {
                            $type     = 'error';
                            $message .= 'Excel verisi eklenirken hata oluştu.<br>';
                        }
                    }

                    $satir++;
                }
            }

            // Geçici yüklenen dosyayı sil
            if (file_exists($targetPath)) {
                unlink($targetPath);
            }
        }

    } else {
        $type    = 'error';
        $message = 'Geçersiz dosya türü. Lütfen .xls veya .xlsx dosyası yükleyin.<br>';
    }
}

// Mevcut stokları çek
$stokListesi = [];
$sqlSelect   = $dba->safeQuery("SELECT * FROM stok");
while ($row = $dba->fetch_object($sqlSelect)) {
    $stokListesi[] = $row;
}
?>
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Stok Aktarım</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            width: 750px;
            margin: 20px auto;
        }
        .outer-container {
            background: #F0F0F0;
            border: #e0dfdf 1px solid;
            padding: 40px 20px;
            border-radius: 2px;
        }
        .btn-submit {
            background: #333;
            border: #1d1d1d 1px solid;
            border-radius: 2px;
            color: #f0f0f0;
            cursor: pointer;
            padding: 5px 20px;
            font-size: 0.9em;
        }
        .btn-submit:hover { background: #555; }
        .tutorial-table {
            margin-top: 40px;
            font-size: 0.8em;
            border-collapse: collapse;
            width: 100%;
        }
        .tutorial-table th {
            background: #f0f0f0;
            border-bottom: 1px solid #dddddd;
            padding: 8px;
            text-align: left;
        }
        .tutorial-table td {
            background: #FFF;
            border-bottom: 1px solid #dddddd;
            padding: 8px;
            text-align: left;
        }
        #response {
            padding: 10px;
            margin-top: 10px;
            border-radius: 2px;
            display: none;
        }
        .success { background: #c7efd9; border: #bbe2cd 1px solid; }
        .error   { background: #fbcfcf; border: #f3c6c7 1px solid; }
        .display-block { display: block; }
    </style>
</head>
<body>

<h2>Stok Aktarım</h2>
<p>
    Stokları aktarmak için örnek Excel kullanabilirsiniz.
    Örnek Excel indirmek için <a href="stoklar.xlsx">Tıklayın</a>.
</p>

<div class="outer-container">
    <form action="" method="post" name="frmExcelImport" id="frmExcelImport" enctype="multipart/form-data">
        <div>
            <label for="file">Excel Dosyası Seçin:</label>
            <input type="file" name="file" id="file" accept=".xls,.xlsx">
            <button type="submit" name="import" class="btn-submit">İçe Aktar</button>
        </div>
    </form>
</div>

<div id="response" class="<?= !empty($type) ? htmlspecialchars($type, ENT_QUOTES, 'UTF-8') . ' display-block' : '' ?>">
    <?= $message ?>
</div>

<?php if (!empty($stokListesi)): ?>
<table class="tutorial-table">
    <thead>
        <tr>
            <th>Stok Kodu</th>
            <th>Stok Adı</th>
            <th>Grup Kodu</th>
            <th>Ek Grup Kodu</th>
            <th>Birim</th>
            <th>KDV Oranı</th>
            <th>Açıklama</th>
            <th>Alış Fiyatı</th>
            <th>Satış Fiyatı</th>
        </tr>
    </thead>
    <tbody>
        <?php foreach ($stokListesi as $stok): ?>
        <tr>
            <td><?= htmlspecialchars((string)($stok->stokkod    ?? ''), ENT_QUOTES, 'UTF-8') ?></td>
            <td><?= htmlspecialchars((string)($stok->stokad     ?? ''), ENT_QUOTES, 'UTF-8') ?></td>
            <td><?= htmlspecialchars((string)($stok->grupkod    ?? ''), ENT_QUOTES, 'UTF-8') ?></td>
            <td><?= htmlspecialchars((string)($stok->ekgrupkod  ?? ''), ENT_QUOTES, 'UTF-8') ?></td>
            <td><?= htmlspecialchars((string)($stok->birim      ?? ''), ENT_QUOTES, 'UTF-8') ?></td>
            <td><?= htmlspecialchars((string)($stok->kdvoran    ?? ''), ENT_QUOTES, 'UTF-8') ?></td>
            <td><?= htmlspecialchars((string)($stok->aciklama   ?? ''), ENT_QUOTES, 'UTF-8') ?></td>
            <td><?= htmlspecialchars((string)($stok->alisfiyat  ?? ''), ENT_QUOTES, 'UTF-8') ?></td>
            <td><?= htmlspecialchars((string)($stok->satisfiyat ?? ''), ENT_QUOTES, 'UTF-8') ?></td>
        </tr>
        <?php endforeach; ?>
    </tbody>
</table>
<?php endif; ?>

</body>
</html>
