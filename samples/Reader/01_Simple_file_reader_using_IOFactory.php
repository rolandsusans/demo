<?php
error_reporting(-1);
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\NamedRange;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

require __DIR__ . '/../Header.php';

$inputFileName = __DIR__ . '/sampleData/excel_array.xlsx';
$helper->log('Loading file ' . pathinfo($inputFileName, PATHINFO_BASENAME) . ' using IOFactory to identify the format');
$spreadsheet = IOFactory::load($inputFileName);
$sheetData   = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);

$payload = [
    'members' => [
        [
            'name' => 'John',
            'age'  => 30,
        ],
        [
            'name' => 'Rolands',
            'age'  => 27,
        ],
    ],
];

$worksheet = $spreadsheet->getSheetByName('members');
$row    = 2;
foreach ($payload['members'] as $member) {
    $header = array_keys($member);
    $worksheet->setCellValue('A'.$row, $member[$header[0]]);
    $worksheet->setCellValue('B'.$row, $member[$header[1]]);
    $row++;
}

$worksheet->setCellValue('A1', $header[0]);
$worksheet->setCellValue('B1', $header[1]);
var_dump('------------------------------------------------------------------------');

$namedRange = new NamedRange('members', $worksheet, 'A2:B3');
var_dump('------------------------------------------------------------------------');
//$spreadsheet->addSheet($worksheet);
$spreadsheet->addNamedRange($namedRange);
var_dump('------------------------------------------------------------------------');
//$worksheet->

$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
var_dump('----------------------------- before save--------------------------------------');
$writer->save(__DIR__ . '/sampleData/out.xlsx');
var_dump('------------------------------------------------------------------------');
var_dump($sheetData);
