
<?php
	//Agregamos la libreria para leer
	require 'Classes/PHPExcel/IOFactory.php';
	
	// Creamos un objeto PHPExcel
	$objPHPExcel = new PHPExcel();
	$objReader = PHPExcel_IOFactory::createReader('Excel2007');
	$objPHPExcel = $objReader->load('ejemplo.xlsx');
	// Indicamos que se pare en la hoja uno del libro
	$objPHPExcel->setActiveSheetIndex(0);
	
	//Modificamos los valoresde las celdas A2, B2 Y C2
	$objPHPExcel->getActiveSheet()->SetCellValue('A2', 'Nuevos Dulces');
	$objPHPExcel->getActiveSheet()->SetCellValue('B2', 8.30);
	$objPHPExcel->getActiveSheet()->SetCellValue('C2', 10);
	
	//Agregamos nuevos valores en las celdas A7, B7 y C7
	$objPHPExcel->getActiveSheet()->SetCellValue('A7', 'Nuevo Producto');
	$objPHPExcel->getActiveSheet()->SetCellValue('B7', 10);
	$objPHPExcel->getActiveSheet()->SetCellValue('C7', 2);
	
	//Guardamos los cambios
	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    // Estos cambios se guardan en un nuevo archivo
	$objWriter->save("Archivo_salida.xlsx");

    // this error
    // "Notice: Trying to access array offset on value of type int in C:\xampp\htdocs\excelphp\Classes\PHPExcel\Cell\DefaultValueBinder.php on line 82"
    // is because this library is not supported anymore. Instead, we must use PHPSpreadSheet
?>