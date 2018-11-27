<?php

use PhpOffice\PhpSpreadsheet\IOFactory;

use PhpOffice\PhpSpreadsheet\Helper\Sample;

require_once $_SERVER['DOCUMENT_ROOT'] . '/demo/src/Bootstrap.php';

$helper = new Sample();


// require __DIR__ . '/../Header.php';

$inputFileType = 'Xlsx';
$inputFileName = __DIR__ . '/sampleData/AmeriCorps.xlsx';

// $helper->log('Loading file ' . pathinfo($inputFileName, PATHINFO_BASENAME) . ' using IOFactory with a defined reader type of ' . $inputFileType);
$reader = IOFactory::createReader($inputFileType);
// $helper->log('Turning Formatting off for Load');
$reader->setReadDataOnly(true);
$spreadsheet = $reader->load($inputFileName);

$sheetData = $spreadsheet->getActiveSheet()->toArray( null, true, true, true);
// var_dump($sheetData);

//create connection
$conn =  new mysqli('localhost','root','','StudentTest');



//check connection
// if ($conn->connection_error){
//     die("Connection failed: " . $conn->connection_error);
// }
$sql = '';
for($row=2; $row <= count($sheetData); $row++){
    $xx = "'" . implode("','", $sheetData[$row]) . "'" ;
    $sql = "INSERT INTO StudentData (StudentId, LastName, FirstName, Location) VALUES ($xx);";
    if ($conn->query($sql) === TRUE) {
        echo "Row $row INSERTED successfull" . "<br>";
    } else {
        echo "Error: " . $sql . "<br>" . $conn->error;
    }

}