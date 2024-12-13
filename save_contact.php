<?php
// Include PhpSpreadsheet library
require 'vendor/autoload.php';

// Grab form data
$name = $_POST['name'];
$email = $_POST['email'];
$subject = $_POST['subject'];
$message = $_POST['message'];

// Create a new spreadsheet
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Set headers
$sheet->setCellValue('A1', 'Name');
$sheet->setCellValue('B1', 'Email');
$sheet->setCellValue('C1', 'Subject');
$sheet->setCellValue('D1', 'Message');

// Add data from the form
$sheet->setCellValue('A2', $name);
$sheet->setCellValue('B2', $email);
$sheet->setCellValue('C2', $subject);
$sheet->setCellValue('D2', $message);

// Check if the file exists, if not create it
$filePath = 'contact_data.xlsx';
if (file_exists($filePath)) {
    $existingData = \PhpOffice\PhpSpreadsheet\IOFactory::load($filePath);
    $sheet = $existingData->getActiveSheet();
    $rowCount = $sheet->getHighestRow() + 1;
    $sheet->setCellValue('A' . $rowCount, $name);
    $sheet->setCellValue('B' . $rowCount, $email);
    $sheet->setCellValue('C' . $rowCount, $subject);
    $sheet->setCellValue('D' . $rowCount, $message);
} else {
    // Create the file for the first time
    $writer = new Xlsx($spreadsheet);
    $writer->save($filePath);
}

// Save the spreadsheet
$writer = new Xlsx($spreadsheet);
$writer->save($filePath);

// Redirect back to the portfolio page or display success message
header('Location: index.html'); // Adjust this to your page location

exit();
?>
