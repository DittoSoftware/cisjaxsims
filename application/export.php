<?php

  $con = new mysqli('localhost', 'root', '');
  mysqli_select_db($con, 'StudentTest');

  $sql = "SELECT * FROM StudentData";
  $result = mysqli_query($con,$sql);

  $columnHeader = '';
  $columnHeader = "StudentId" . "\t" . "LastName" . "\t" . "FirstName" . "\t" . "Location" . "\t";

  $setData = '';

  while ($rec = mysqli_fetch_row($result)) {
      $rowData = '';
      foreach ($rec as $value) {
          $value = '"' . $value . '"' . "\t";
          $rowData .= $value;
      }
      $setData .= trim($rowData) . "\n";
  }

header("Content-type: application/xls");
header("Content-Disposition: attachment; filename=combined.xls");
header("Pragma: no-cache");
header("Expires: 0");

echo ucwords($columnHeader) . "\n" . $setData . "\n";

 ?>
