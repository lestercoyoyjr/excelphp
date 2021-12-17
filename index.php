<?php
    // note this worked, however PHPExcel has no support anymore
    // so you must use PHPSpreadSheet
    // here is the link https://github.com/PHPOffice/PhpSpreadsheet
    require_once 'Classes/PHPExcel.php';
    $objPHPExcel = new PHPExcel();

    // metadata
    $objPHPExcel->getProperties()
        ->setCreator("Códigos de Programación")
        ->setLastModifiedBy("Códigos de Programación")
        ->setTitle("Excel en PHP")
        ->setSubject("Documento de prueba")
        ->setDescription("Documento generado con PHPExcel")
        ->setKeywords("excel phpexcel php")
        ->setCategory("Ejemplos");

    // sheet info
    $objPHPExcel->setActiveSheetIndex(0);
    $objPHPExcel->getActiveSheet()->setTitle('Hoja 1');

    // cells content
    
    // Agregar en celda A1 valor string
    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'PHPExcel');
	
    // Agregar en celda A2 valor numerico
    $objPHPExcel->getActiveSheet()->setCellValue('A2', 12345.6789);
	
    // Agregar en celda A3 valor boleano
    $objPHPExcel->getActiveSheet()->setCellValue('A3', TRUE);
	
    // Agregar a Celda A4 una formula
    $objPHPExcel->getActiveSheet()->setCellValue('A4', '=CONCATENATE(A1, " ", A2)');

    // Download file
    // xls
    // header('Content-Type: application/vnd.ms-excel');
    // header('Content-Disposition: attachment;filename="Excel.xls"');
    // header('Cache-Control: max-age=0');
	
    // $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
    // $objWriter->save('php://output');

    // xlsx
    
    header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    header('Content-Disposition: attachment;filename="Excel.xlsx"');
    header('Cache-Control: max-age=0');
 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save('php://output');

?>