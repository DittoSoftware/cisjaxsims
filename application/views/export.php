<?php
require_once ($libpath . 'phpspreadsheet/vendor/autoload.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;

$spreadsheet = new Spreadsheet();  /*----Spreadsheet object-----*/
$Excel_writer = new Xls($spreadsheet);  /*----- Excel (Xls) Object*/

$spreadsheet->setActiveSheetIndex(0);
$activeSheet = $spreadsheet->getActiveSheet();

$activeSheet->setCellValue('A1' , 'New file content')->getStyle('A1')->getFont()->setBold(true);


header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="'. $filename .'.xls"'); /*-- $filename is  xsl filename ---*/
header('Cache-Control: max-age=0');

ob_end_clean();
// $objWriter->save('php://output', 'xls');

$Excel_writer->save('php://output');






// $conn =  mysqli_connect('localhost','root','','StudentTest');
// $output = '';
// if(isset($_POST["export_excel"]))
// {
//   $sql = "SELECT * FROM StudentData";
//   $result = mysqli_query($conn, $sql);
//   if(mysqli_num_rows($result)>0)
//   {
//     $output .= '
//         <table class="table">
//           <tr>
//             <th>StudentId</th>
//             <th>LastName</th>
//             <th>FirstName</th>
//             <th>Location</th>
//           </tr>
//     ';
//     while($row = mysqli_fetch_array($result))
//     {
//         $output .='
//           <tr>
//             <td>'.$row["id"].'</td>
//             <td>'.$row["ln"].'</td>
//             <td>'.$row["fn"].'</td>
//             <td>'.$row["loc"].'</td>
//           </tr>
//         ';
//     }
//     $output .= '</table>';
//     header("Content-Type: application/xls");
//     header("Content-Disposition: attachment, filename=download.xls");
//     echo $output;
//   }
// }



 ?>
