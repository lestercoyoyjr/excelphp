<?php
    // this library is not supported anymore but we tried it for
    // learning purposes
    
    // We include libraries and files for connection
	require 'Classes/PHPExcel.php';
	require 'conexion.php';

    // Query
	$sql = "SELECT id, nombre, precio, existencia FROM productos";
	$resultado = $mysqli->query($sql);
    // We stablish the row where it will begin
    $fila = 2;

    // Logo
    $gdImage = imagecreatefrompng('images/logo.jpg');

    // PHPExcel Object
	$objPHPExcel  = new PHPExcel();
    // Documents properties
	$objPHPExcel->getProperties()->setCreator("Lester Coyoy")->setDescription("Reporte de Productos");

    // We stablish the first tab and its title
	$objPHPExcel->setActiveSheetIndex(0);
	$objPHPExcel->getActiveSheet()->setTitle("Productos");

    // Headers
    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
    $objPHPExcel->getActiveSheet()->setCellValue('B1', 'NOMBRE');
    $objPHPExcel->getActiveSheet()->setCellValue('C1', 'PRECIO');
    $objPHPExcel->getActiveSheet()->setCellValue('D1', 'EXISTENCIA');
    $objPHPExcel->getActiveSheet()->setCellValue('E1', 'TOTAL');

    while($row = $resultado->fetch_assoc()){
        $objPHPExcel->getActiveSheet()->setCellValue('A'.$fila, $row['id']);
		$objPHPExcel->getActiveSheet()->setCellValue('B'.$fila, $row['nombre']);
		$objPHPExcel->getActiveSheet()->setCellValue('C'.$fila, $row['precio']);
		$objPHPExcel->getActiveSheet()->setCellValue('D'.$fila, $row['existencia']);
		// the last one is an operation
        $objPHPExcel->getActiveSheet()->setCellValue('E'.$fila, '=C'.$fila.'*D'.$fila);
		
		// We increase 1 to pass to next row
        $fila++; 
    }

    // We save our document
    // format
    header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    // name and extension
	header('Content-Disposition: attachment;filename="Productos.xlsx"');
    // cache
	header('Cache-Control: max-age=0');
    // version
    $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
    $objWriter->save('php://output');
?>