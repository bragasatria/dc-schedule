<?php

namespace App\Services;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;


class CableService
{
    protected $max_length = 1000;
    protected $url_logo_atw = 'https://pradita-s3.s3.ap-southeast-3.amazonaws.com/prod/alumni/logo_mitra/1693555482_ATW%20SOLAR%20PNG%20-%20HRD%20ATW%20Solar%20Indonesia.png';

    /**
     * Algoritma Utama: Menghitung Roll (Konversi dari Python logic)
     */
    public function calculateRolls(array $dataList, array $headerNames, int $maxSeconds = 20)
    {
        $startTime = microtime(true);
        $cablesToProcess = $dataList;
        $rollCounter = 1;

        // Header key mapping
        $keyName = $headerNames[0];
        $keyLength = $headerNames[1];

        // Selama masih ada yang 'Roll/Drum' nya null, dan batasi loop agar tidak infinite
        $maxRolls = 100; // Limit maksimal roll untuk menghindari stuck
        while ($this->hasUnassigned($cablesToProcess) && $rollCounter <= $maxRolls) {
            // Timeout safety: jika sudah lebih dari batas maksimal waktu, stop dan error
            if ((microtime(true) - $startTime) > $maxSeconds) {
                throw new \Exception('Proses terlalu lama (lebih dari ' . $maxSeconds . ' detik). Periksa kembali isi file Anda.');
            }
            // Ambil item yang belum punya roll
            $unassignedCables = [];
            foreach ($cablesToProcess as $index => $item) {
                if ($item['Roll/Drum'] === null) {
                    $unassignedCables[] = ['index' => $index, 'length' => $item[$keyLength]];
                }
            }

            if (empty($unassignedCables)) break;

            // DP Logic (Knapsack variation)
            // Format DP: [total_length => ['sum' => int, 'indices' => []]]
            $dp = [0 => ['sum' => 0, 'indices' => []]];

            foreach ($unassignedCables as $cable) {
                $originalIndex = $cable['index'];
                $length = $cable['length'];

                // Kita copy array DP saat ini untuk iterasi agar tidak loop infinite pada item sendiri
                $currentDp = $dp;

                foreach ($currentDp as $w => $data) {
                    $newW = $w + $length;

                    if ($newW <= $this->max_length) {
                        $currentSum = $data['sum'];
                        $currentIndices = $data['indices'];

                        // Jika bobot baru belum ada, atau kita nemu kombinasi yang lebih besar (tapi <= max)
                        // Logika Python: if new_w not in dp or dp[new_w][0] < current_sum + length:
                        if (!isset($dp[$newW]) || $dp[$newW]['sum'] < ($currentSum + $length)) {
                            $newIndices = $currentIndices;
                            $newIndices[] = $originalIndex;
                            $dp[$newW] = [
                                'sum' => $currentSum + $length,
                                'indices' => $newIndices
                            ];
                        }
                    }
                }
            }

            // Cari nilai max dari DP
            if (count($dp) > 1) {
                // Sort by key (total length) descending untuk ambil yang paling optimal
                krsort($dp);
                $bestSolution = reset($dp); // Ambil elemen pertama setelah sort
                $bestIndices = $bestSolution['indices'];
            } else {
                continue;
            }

            // Assign Roll Counter
            foreach ($bestIndices as $idx) {
                $cablesToProcess[$idx]['Roll/Drum'] = $rollCounter;
            }

            $rollCounter++;
        }

        return $cablesToProcess;
    }

    private function hasUnassigned($list)
    {
        foreach ($list as $item) {
            if ($item['Roll/Drum'] === null) return true;
        }
        return false;
    }

    /**
     * Membuat Sheet Summary (Konversi logic create_summary_sheet)
     */
    public function createSummarySheet(Spreadsheet $spreadsheet, array $summaryData)
    {
        // Cek sheet, buat jika belum ada di index 0
        $sheet = $spreadsheet->getSheetByName('Summary DC Schedule');
        if (!$sheet) {
            $sheet = new Worksheet($spreadsheet, 'Summary DC Schedule');
            $spreadsheet->addSheet($sheet, 0);
        }
        $spreadsheet->setActiveSheetIndex(0);

        // Styling
        $thinBorder = [
            'borders' => [
                'allBorders' => ['borderStyle' => Border::BORDER_THIN],
            ],
        ];
        $grayFill = [
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['argb' => 'FFD9D9D9'],
            ],
        ];
        $headerFont = ['font' => ['bold' => true, 'name' => 'Calibri']];

        // Insert Logo ATW
        $this->insertImageFromUrl($this->url_logo_atw, $sheet, 'A2', 50, 150);

        // Jarak baris
        $row = 6;
        $headers = ["Sheet Name", "For", "Total Roll/Drum"];
        $sheet->fromArray($headers, null, "A{$row}");

        // Style Header
        $sheet->getStyle("A{$row}:C{$row}")->applyFromArray(array_merge($thinBorder, $grayFill, $headerFont));
        $sheet->getStyle("A{$row}:C{$row}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $row++;
        $grandTotal = 0;

        // Sort data based on sheet name
        usort($summaryData, function($a, $b) {
            return strcmp($a['sheet_name'], $b['sheet_name']);
        });

        foreach ($summaryData as $item) {
            $sheetName = $item['sheet_name'];
            $total = $item['total_roll'];
            $grandTotal += $total;

            $kabelType = "DC Cable";
            if (str_contains($sheetName, "(+)") || str_contains($sheetName, "Positif")) $kabelType = "Positif";
            elseif (str_contains($sheetName, "(-)") || str_contains($sheetName, "Negatif")) $kabelType = "Negatif";
            elseif (str_contains($sheetName, "HR") || str_contains($sheetName, "Homerun")) $kabelType = "Homerun";

            $sheet->setCellValue("A{$row}", $sheetName);
            $sheet->setCellValue("B{$row}", $kabelType);
            $sheet->setCellValue("C{$row}", $total);

            $sheet->getStyle("A{$row}:C{$row}")->applyFromArray($thinBorder);
            $sheet->getStyle("C{$row}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $row++;
        }

        // Grand Total Row
        $sheet->setCellValue("A{$row}", "Grand Total Roll/Drum");
        $sheet->setCellValue("C{$row}", $grandTotal);
        $sheet->mergeCells("A{$row}:B{$row}");

        $sheet->getStyle("A{$row}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle("A{$row}")->getFont()->setBold(true);
        $sheet->getStyle("C{$row}")->getFont()->setBold(true);
        $sheet->getStyle("C{$row}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle("C{$row}")->applyFromArray($thinBorder);

        // Column Widths
        $sheet->getColumnDimension('A')->setWidth(25);
        $sheet->getColumnDimension('B')->setWidth(15);
        $sheet->getColumnDimension('C')->setWidth(20);
    }

    /**
     * Helper Insert Image dari URL dengan aman
     */
    public function insertImageFromUrl($url, $sheet, $coordinates, $height, $width = null)
    {
        try {
            // Kita download dulu kontennya ke temp memory
            $contents = file_get_contents($url);
            if ($contents !== false) {
                $tempPath = tempnam(sys_get_temp_dir(), 'img');
                file_put_contents($tempPath, $contents);

                $drawing = new Drawing();
                $drawing->setPath($tempPath);
                $drawing->setCoordinates($coordinates);
                $drawing->setHeight($height);
                if ($width) $drawing->setWidth($width);
                $drawing->setWorksheet($sheet);

                // Unlink nanti ditangani sistem OS temp cleanup atau biarkan saja
            }
        } catch (\Exception $e) {
            // Ignore error image like in python code
        }
    }

    public function autoFitColumn($sheet)
    {
        foreach ($sheet->getColumnIterator() as $column) {
            $sheet->getColumnDimension($column->getColumnIndex())->setAutoSize(true);
        }
    }
}
