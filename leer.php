
<?php
	//Agregamos la librería
    require 'Classes/PHPExcel/IOFactory.php';  
    //Agregamos la conexión
	require 'conexion.php'; 
	
	//Variable con el nombre del archivo
	$nombreArchivo = 'ejemplo.xlsx';
	// Cargo la hoja de cálculo
	$objPHPExcel = PHPExcel_IOFactory::load($nombreArchivo);
	
	//Asigno la hoja de calculo activa
	$objPHPExcel->setActiveSheetIndex(0);
	//Obtengo el numero de filas del archivo
	$numRows = $objPHPExcel->setActiveSheetIndex(0)->getHighestRow();
	
    // Impresion de los encabezados
    // if it already has a header, it will print it for two. Just start the for bucle
    // i=2
	echo '<table border=1><tr><td>Producto</td><td>Precio</td><td>Existencia</td></tr>';
	
    // Para ir leyendo las filas
	for ($i = 2; $i <= $numRows; $i++) {
		
		// nombre
        $nombre = $objPHPExcel->getActiveSheet()->getCell('A'.$i)->getCalculatedValue();
        // Precio
		$precio = $objPHPExcel->getActiveSheet()->getCell('B'.$i)->getCalculatedValue();
        // Existencia
		$existencia = $objPHPExcel->getActiveSheet()->getCell('C'.$i)->getCalculatedValue();
		
        // impresion de las columnas y sus valores para html
		echo '<tr>';
		echo '<td>'. $nombre.'</td>';
		echo '<td>'. $precio.'</td>';
		echo '<td>'. $existencia.'</td>';
		echo '</tr>';
		
        // query para insersion
		$sql = "INSERT INTO productos (nombre, precio, existencia) VALUES('$nombre','$precio','$existencia')";
		$result = $mysqli->query($sql);
	}
	
	echo '<table>';

    // this error
    // "Notice: Trying to access array offset on value of type int in C:\xampp\htdocs\excelphp\Classes\PHPExcel\Cell\DefaultValueBinder.php on line 82"
    // is because this library is not supported anymore. Instead, we must use PHPSpreadSheet
?>
