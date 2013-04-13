<?php
// error_reporting(E_ALL);
// ini_set('display_errors', TRUE);
// ini_set('display_startup_errors', TRUE);

if(empty($_POST['contenttoexcel'])) exit;
// Config
//if(!defined('PHPExcel_PATH')) define('PHPExcel_PATH', '../../php/PHPExcel/');
if(!defined('PHPExcel_PATH')) define('PHPExcel_PATH', './PHPExcel/');
require_once PHPExcel_PATH . 'Classes/PHPExcel.php';
$data = utf8_encode($_POST['contenttoexcel']);
$data = str_replace('%%SQUOT%%', '\'', $data);
$data = str_replace('%%DQUOT%%', '\"', $data);
$data = json_decode(html_entity_decode($data));
$title = $data->title;
$content = $data->items;
$highestRow = count($content);
$highestColumn = count($title);
$objPHPExcel = new PHPExcel();
$objPHPExcel->getProperties()->setCreator("Generate Excel Tool")
							 ->setTitle("ECM documents")
							 ->setDescription("ECM documents exprot by Generate Excel Tool");
$objPHPExcel->setActiveSheetIndex(0);
$objWorksheet = $objPHPExcel->getActiveSheet();
// **Row begins with index 1, and column begins with index 0, 
//		so the index coordinate of cell A1's  is (0, 1).

// Set title
for($row = 1, $col = 0; $col < $highestColumn; ++$col) {
	$objWorksheet->setCellValueByColumnAndRow($col, $row, $title[$col]);
}
// Set content
for ($row = 1; $row <= $highestRow; ++$row) {
	for ($col = 0; $col < $highestColumn; ++$col) {
		$objWorksheet->setCellValueByColumnAndRow($col, $row + 1, $content[$row - 1][$col]);
	}
}
//Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="output_documents.xls"');
header('Cache-Control: max-age=0');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
exit;