<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Table;
use PhpOffice\PhpSpreadsheet\Exception as SpreadsheetException;
use PhpOffice\PhpSpreadsheet\Reader\Exception as ReaderException;
use App\Services\CableService;


class ExcelCalculatorController extends Controller
{
    protected $cableService;

    public function __construct(CableService $cableService)
    {
        $this->cableService = $cableService;
    }

    public function index()
    {
        return view('mainpage');
    }

    public function process(Request $request)
    {
        try {
            // Perpanjang batas eksekusi di level PHP agar tidak langsung fatal ketika data besar.
            // Kita tetap punya timeout internal (di CableService) untuk menangani kasus data terlalu berat.
            set_time_limit(60); // detik

            $request->validate([
                'excel_file' => 'required|mimes:xlsx,xls|max:10240',
                'output_name' => 'nullable|string|max:100',
            ]);

            $file = $request->file('excel_file');

            // Load Excel Input
            $reader = IOFactory::createReader('Xlsx');
            $reader->setReadDataOnly(true);
            $spreadsheetInput = $reader->load($file->getPathname());

            // Buat Excel Output Baru
            $spreadsheetOutput = new Spreadsheet();
            $spreadsheetOutput->removeSheetByIndex(0); // Hapus sheet default

            $sheetIndex = 0;
            $globalSummaryData = [];

            foreach ($spreadsheetInput->getSheetNames() as $sheetName) {
                $sheetIn = $spreadsheetInput->getSheetByName($sheetName);

                // Baca Header (Baris 2, karena baris 1 biasanya judul di format Anda)
                // Asumsi data mulai baris 2 seperti kode Python: headers = [cell.value for cell in sheet_in[1]]
                // Tapi iter_rows min_row=2. Mari kita ambil data mentah array.

                $rawRows = $sheetIn->toArray(null, true, true, true);
                // Format toArray: ['A' => val, 'B' => val]
                // Asumsi header ada di baris 2 (indeks array 2 di PHPExcel jika start row 1? Tidak, array PHP mulai indexnya sesuai row number jika formatnya begini)

                // Kita cari baris pertama yang valid sebagai header.
                // Sesuai Python: headers = [cell.value for cell in sheet_in[1]] -> Ini berarti baris ke-2 excel (index 1 python 0-based).

                // Mari kita ambil semua data dulu
                $dataPerSheet = [];
                $headers = [];

                // Ambil header di row 2 (sesuai kode python sheet_in[1])
                $headerRow = $sheetIn->rangeToArray('A2:B2', null, true, true, true)[2];
                $headers = array_values($headerRow); // [0 => Label, 1 => Length]

                if (empty($headers[0]) || empty($headers[1])) continue;

                // Ambil data mulai row 3 (sesuai python min_row=2 tapi python row start 1, jadi ini row data)
                // Python: iter_rows(min_row=2). Jika Python sheet[1] adalah header, berarti row 2 Excel adalah Header.
                // Data dimulai dari row 3 Excel.

                $maxRow = $sheetIn->getHighestRow();
                for ($r = 3; $r <= $maxRow; $r++) {
                    $valA = $sheetIn->getCell("A$r")->getValue();
                    $valB = $sheetIn->getCell("B$r")->getValue();

                    if ($valA === null || $valB === null) continue;

                    $dataPerSheet[] = [
                        $headers[0] => $valA,
                        $headers[1] => (int)$valB,
                        'Roll/Drum' => null
                    ];
                }

                if (empty($dataPerSheet)) continue;

                // Validasi: Pastikan semua data panjang adalah angka
                foreach ($dataPerSheet as $row) {
                    if (!is_numeric($row[$headers[1]])) {
                        throw new \Exception('Data panjang harus berupa angka. Periksa kolom kedua di sheet "' . $sheetName . '".');
                    }
                }

                // --- PROSES LOGIC (Panggil Service) ---
                try {
                    // Kirim batas waktu (detik) agar jika terlalu lama, kita bisa catch dan kembali ke halaman utama
                    $processedData = $this->cableService->calculateRolls($dataPerSheet, $headers, 20);
                } catch (\Exception $e) {
                    throw new \Exception('Gagal memproses data di sheet "' . $sheetName . '". ' . $e->getMessage());
                }

                // Hitung Unique Roll untuk summary global
                $rolls = array_filter(array_column($processedData, 'Roll/Drum'));
                $totalRollsSheet = count(array_unique($rolls));

                $globalSummaryData[] = [
                    'sheet_name' => $sheetName,
                    'total_roll' => $totalRollsSheet
                ];

                // --- BUAT SHEET OUTPUT ---
                // Sanitize dan truncate nama sheet agar sesuai batas Excel (maks 31 karakter)
                $maxSheetNameLength = 31;
                $prefix = "Schedule ";
                $availableLength = $maxSheetNameLength - strlen($prefix);
                $sanitizedSheetName = preg_replace('/[\\/\\?\\*\\[\\]:]/', '_', $sheetName); // Ganti karakter terlarang dengan _
                $truncatedSheetName = substr($sanitizedSheetName, 0, $availableLength);
                if (empty($truncatedSheetName)) {
                    $truncatedSheetName = 'Sheet'; // Fallback jika nama sheet kosong setelah truncate
                }
                $finalSheetName = $prefix . $truncatedSheetName;

                $sheetOut = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheetOutput, $finalSheetName);
                $spreadsheetOutput->addSheet($sheetOut);

                // Judul Header
                $sheetOut->mergeCells('A1:C3');
                $sheetOut->setCellValue('A1', "DC SCHEDULE PROGRAM\n" . $sheetName);
                $sheetOut->getStyle('A1')->getFont()->setBold(true)->setSize(16)->setName('Calibri');
                $sheetOut->getStyle('A1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER)->setVertical(Alignment::VERTICAL_CENTER)->setWrapText(true);

                // Insert Images
                $this->cableService->insertImageFromUrl('https://pradita-s3.s3.ap-southeast-3.amazonaws.com/prod/alumni/logo_mitra/1693555482_ATW%20SOLAR%20PNG%20-%20HRD%20ATW%20Solar%20Indonesia.png', $sheetOut, 'A1', 60);

                // Warna Tab
                $sheetIndex++;
                $tabColor = ($sheetIndex % 2 == 1) ? 'FF0000' : '000000';
                $sheetOut->getTabColor()->setRGB($tabColor);

                // --- TABEL DETAIL ---
                $currentRow = 6; // Mulai setelah header + spasi
                $headerDetail = array_keys($processedData[0]);
                $sheetOut->fromArray($headerDetail, null, "A$currentRow");

                $startRowDetail = $currentRow;
                $currentRow++;

                foreach ($processedData as $rowVal) {
                    $sheetOut->fromArray(array_values($rowVal), null, "A$currentRow");
                    $currentRow++;
                }
                $endRowDetail = $currentRow - 1;

                // Spasi antar tabel
                $currentRow += 4;

                // --- TABEL SUMMARY ---
                // Grouping data logic
                $rollsSummary = [];
                foreach ($processedData as $item) {
                    $rollNum = $item['Roll/Drum'];
                    if ($rollNum !== null) {
                        if (!isset($rollsSummary[$rollNum])) {
                            $rollsSummary[$rollNum] = ['labels' => [], 'total' => 0];
                        }
                        $rollsSummary[$rollNum]['labels'][] = $item[$headers[0]];
                        $rollsSummary[$rollNum]['total'] += $item[$headers[1]];
                    }
                }
                ksort($rollsSummary);

                $headerSummary = ['Nomor Roll/Drum', 'Total Length (m)', 'Detail String Label'];
                $sheetOut->fromArray($headerSummary, null, "A$currentRow");

                $startRowSummary = $currentRow;
                $currentRow++;

                foreach ($rollsSummary as $rNum => $det) {
                    $labelsStr = implode(", ", $det['labels']);
                    $sheetOut->fromArray([$rNum, $det['total'], $labelsStr], null, "A$currentRow");
                    $currentRow++;
                }
                $endRowSummary = $currentRow - 1;

                // Create Excel Tables
                $safeTitle = preg_replace('/[^a-zA-Z0-9]/', '', $sheetName);
                $uniqueSuffix = "_" . $sheetIndex;

                // Table 1
                $table1 = new Table("A{$startRowDetail}:C{$endRowDetail}", "TableDetail_{$safeTitle}{$uniqueSuffix}");
                $table1Style = new \PhpOffice\PhpSpreadsheet\Worksheet\Table\TableStyle();
                $table1Style->setTheme(\PhpOffice\PhpSpreadsheet\Worksheet\Table\TableStyle::TABLE_STYLE_MEDIUM15);
                $table1Style->setShowRowStripes(true);
                $table1->setStyle($table1Style);
                $sheetOut->addTable($table1);

                // Table 2
                $table2 = new Table("A{$startRowSummary}:C{$endRowSummary}", "TableSummary_{$safeTitle}{$uniqueSuffix}");
                $table2->setStyle($table1Style); // Reuse style
                $sheetOut->addTable($table2);

                // Formatting Widths & Wrap
                // 1. Auto fit dulu untuk merapikan Kolom A dan B
                $this->cableService->autoFitColumn($sheetOut);

                // 2. KHUSUS KOLOM C: Kita timpa aturannya
                // Matikan auto size agar tidak melebar otomatis
                $sheetOut->getColumnDimension('C')->setAutoSize(false);

                // Set lebar manual (Angka 45 atau 50 biasanya pas agar tidak terlalu panjang)
                $sheetOut->getColumnDimension('C')->setWidth(60);

                // 3. Pastikan Text Wrapping aktif (agar teks panjang turun ke baris baru)
                $sheetOut->getStyle("C{$startRowSummary}:C{$endRowSummary}")
                    ->getAlignment()
                    ->setWrapText(true) // Ini kuncinya agar teks tidak memanjang ke samping
                    ->setVertical(Alignment::VERTICAL_TOP)
                    ->setHorizontal(Alignment::HORIZONTAL_LEFT);
            }

            // Validasi: Jika tidak ada data valid di sheet manapun, throw error
            if (empty($globalSummaryData)) {
                throw new \Exception('Data tidak ditemukan. Pastikan file Excel memiliki header di baris 2 dan data di baris 3 ke bawah.');
            }

            // --- GLOBAL SUMMARY ---
            if (!empty($globalSummaryData)) {
                $this->cableService->createSummarySheet($spreadsheetOutput, $globalSummaryData);
            }

            // --- DOWNLOAD ---
            // Nama default (jika user tidak mengisi) diambil dari file yang diunggah.
            // Contoh unggahan: "data_input.xlsx" -> output: "data_input 20260315_150102.xlsx"
            $uploadedName = $file->getClientOriginalName();
            $baseUploadedName = pathinfo($uploadedName, PATHINFO_FILENAME);

            // Jika user mengisi nama output, gunakan itu; jika kosong, pakai nama file input.
            $outputName = trim($request->input('output_name', ''));
            // Hapus karakter yang tidak valid di nama file Windows
            $outputName = str_replace(["\\", '/', ':', '*', '?', '"', '<', '>', '|'], '', $outputName);
            $outputName = preg_replace('/\s+/', ' ', $outputName);

            if ($outputName === '') {
                $outputName = $baseUploadedName;
            }

            // Tambahkan timestamp agar setiap download unik
            $timestamp = date('Ymd_His');
            $outputName = trim($outputName) . ' ' . $timestamp;

            // Pastikan ekstensi .xlsx
            if (!str_ends_with(strtolower($outputName), '.xlsx')) {
                $outputName .= '.xlsx';
            }

            $writer = IOFactory::createWriter($spreadsheetOutput, 'Xlsx');

            // Stream download langsung ke browser
            return response()->streamDownload(function() use ($writer) {
                $writer->save('php://output');
            }, $outputName);
        } catch (ReaderException $e) {
            return redirect('/')->withErrors(['excel_file' => 'Pastikan file Excel (.xlsx) yang Anda Kirim!']);
        } catch (SpreadsheetException $e) {
            return redirect('/')->withErrors(['excel_file' => 'Pastikan file Excel (.xlsx) yang Anda Kirim!']);
        } catch (\Exception $e) {
            return redirect('/')->withErrors(['excel_file' => 'Pastikan file Excel (.xlsx) yang Anda Kirim!']);
        }
    }
}
