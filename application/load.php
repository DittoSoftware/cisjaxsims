<!DOCTYPE html>
<html lang="eng">
<head>
    <meta charset="UTF-8">
    <title>Data Preview</title>
    <link rel="stylesheet" type="text/css" href="load.css">
</head>
<body style="background: url(loginBackground.jpg);">

<center>
<h1>Data Preview</h1>
<div class='div1'><table class='table1'><tr ><th>StudentId</th><th> LastName</th><th >FirstName</th><th >Location</th></tr></table></div><div class='div2'><table class='table2'>
<?php


use PhpOffice\PhpSpreadsheet\IOFactory;

use PhpOffice\PhpSpreadsheet\Helper\Sample;

require_once $_SERVER['DOCUMENT_ROOT'] . '/demo/src/Bootstrap.php';

$helper = new Sample();

// echo $_SESSION['file'];

$inputFileType = 'Xlsx';
session_start();
if(isset($_SESSION['file'])){
  echo 'set';
}else{
$inputFileName = 'AmeriCorps.xlsx';
}
// $inputFileName = $_SESSION['file'];

$reader = IOFactory::createReader($inputFileType);
$reader->setReadDataOnly(true);
$spreadsheet = $reader->load($inputFileName);

$sheetData = $spreadsheet->getActiveSheet()->toArray();
// var_dump($sheetData);

//create connection
$conn =  new mysqli('localhost','root','','StudentTest');

$trunc = "TRUNCATE TABLE StudentData";
$conn->query($trunc);

$sql = '';
for($row=2; $row <= count($sheetData) - 1; $row++){
    $a = $sheetData[$row];
    $id = $a[0];
    $lname = $a[1];
    $fname = $a[2];
    $loc = $a[3];

    $sql = "INSERT INTO StudentData (StudentId, LastName, FirstName, Location)
            VALUES (
              ". "'" . mysqli_real_escape_string($conn,$id) . "',
              ". "'" . mysqli_real_escape_string($conn,$lname) . "',
              ". "'" . mysqli_real_escape_string($conn,$fname) . "',
              ". "'" . mysqli_real_escape_string($conn,$loc) . "'
            );";
    // $conn->query($sql);
    // echo "<tr><td>$id</td><td>$lname</td><td>$fname</td><td>$loc</td></tr>";
    if ($conn->query($sql) === TRUE) {
        echo "<tr><td id='id' >$id</td><td >$lname</td><td >$fname</td><td >$loc</td></tr>";
    } else {
        // echo "Error: " . $sql . "<br>" . $conn->error;
        echo "<tr><td>Error</td><td>Error</td><td>Error</td><td>Error</td></tr>";
    }

}
// $sql2 = "SELECT * FROM StudentTest";
// $result = mysqli_query($conn,$sql2);
// $rowcount = mysqli_num_rows($result);
//
// mysqli_close($conn);

?>
</table>
</div>
<br>
<table>
  <tr>
    <td>
      <form method="post" action="export.php">
        <input type="submit" action="export.php" name="export_excel" class="btn btn-success" value="Download">
    </form>
      </td>
      <td>
      <button type="button" onclick="location.href='home.php'">Home</button>
      </td>
  </tr>
</table>




</center>
</body>
</html>
