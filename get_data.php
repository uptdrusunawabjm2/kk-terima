<?php
header('Content-Type: application/json; charset=utf-8');

/**
 * get_data.php
 * Membaca file Excel “dbtbkksewa excel.xlsx”
 * dan mengirim data JSON ke index.html
 * dengan format tanggal DD/MM/YYYY
 */

// ===================================================
// Konfigurasi file Excel
// ===================================================
$excelFile = __DIR__ . '/dbtbkksewa excel.xlsx';

if (!file_exists($excelFile)) {
    echo json_encode(["error" => "❌ File 'dbtbkksewa excel.xlsx' tidak ditemukan di folder yang sama dengan get_data.php"]);
    exit;
}

// ===================================================
// Library SimpleXLSX (tanpa Composer)
// ===================================================
class SimpleXLSX {
    private $sheets;
    public static function parse($filename) {
        $zip = new ZipArchive();
        if ($zip->open($filename) !== true) throw new Exception('Tidak bisa membuka file Excel');
        $sharedStrings = [];
        if ($xml = $zip->getFromName('xl/sharedStrings.xml')) {
            $dom = new DOMDocument();
            $dom->loadXML($xml);
            foreach ($dom->getElementsByTagName('t') as $t) {
                $sharedStrings[] = $t->nodeValue;
            }
        }

        $sheets = [];
        for ($i = 1; $xml = $zip->getFromName('xl/worksheets/sheet'.$i.'.xml'); $i++) {
            $dom = new DOMDocument();
            $dom->loadXML($xml);
            $rows = [];
            foreach ($dom->getElementsByTagName('row') as $row) {
                $cells = [];
                foreach ($row->getElementsByTagName('c') as $c) {
                    $type = $c->getAttribute('t');
                    $v = $c->getElementsByTagName('v')->item(0);
                    if ($v) {
                        $val = $v->nodeValue;
                        if ($type === 's') {
                            $val = $sharedStrings[$val] ?? $val;
                        }
                        $cells[] = $val;
                    } else {
                        $cells[] = '';
                    }
                }
                $rows[] = $cells;
            }
            $sheets[] = $rows;
        }
        $zip->close();
        $xlsx = new self();
        $xlsx->sheets = $sheets;
        return $xlsx;
    }

    public function rows($sheetIndex = 0) {
        return $this->sheets[$sheetIndex];
    }
}

// ===================================================
// Fungsi bantu konversi tanggal Excel ke DD/MM/YYYY
// ===================================================
function excelDateToPHP($excelDate) {
    if (!is_numeric($excelDate)) return $excelDate;
    $unix_date = ((int)$excelDate - 25569) * 86400;
    if ($unix_date < 0) return $excelDate;
    return gmdate("d/m/Y", $unix_date);
}

// ===================================================
// Baca dan kirim data
// ===================================================
try {
    $xlsx = SimpleXLSX::parse($excelFile);
    $rows = $xlsx->rows();

    if (count($rows) < 2) {
        echo json_encode(["error" => "File Excel tidak memiliki data yang valid"]);
        exit;
    }

    $headers = array_map(function($h) {
        $key = strtolower(trim(str_replace([' ', '.', '/', '-'], '_', $h)));
        return $key;
    }, $rows[0]);

    $data = [];
    for ($i = 1; $i < count($rows); $i++) {
        $rowData = [];
        foreach ($headers as $index => $key) {
            $value = isset($rows[$i][$index]) ? $rows[$i][$index] : '';

            // ubah kolom yang mengandung kata 'tanggal'
            if (stripos($key, 'tanggal') !== false && is_numeric($value)) {
                $value = excelDateToPHP($value);
            }

            $rowData[$key] = $value;
        }
        $data[] = $rowData;
    }

    echo json_encode($data, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);

} catch (Exception $e) {
    echo json_encode(["error" => "Gagal membaca Excel: " . $e->getMessage()]);
}
