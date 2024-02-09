<?php
require 'vendor/autoload.php'; // Reemplaza 'vendor/autoload.php' con la ubicación real del archivo autoload.php de PhpSpreadsheet

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;

// Ruta de la carpeta donde se encuentran los archivos Excel
$ruta_carpeta = 'C:\xampp\htdocs\pdfToText\excel'; // Cambia esto a la ubicación real de la carpeta

// Obtener la lista de archivos en la carpeta
$archivos = scandir($ruta_carpeta);
$archivos = array_diff($archivos, array('.', '..')); // Eliminar los directorios "." y ".."

// Crear una instancia de Spreadsheet para el archivo final
$spreadsheet = new Spreadsheet();

// Obtener la hoja activa del archivo final
$sheet = $spreadsheet->getActiveSheet();

// Iterar sobre cada archivo Excel y copiar su contenido al archivo final
foreach ($archivos as $archivo) {
    $archivo_completo = $ruta_carpeta . '\\' . $archivo;

    $reader = IOFactory::createReader('Xlsx');
    $reader->setReadDataOnly(true);
    $excel = $reader->load($archivo_completo);

    // Obtener la hoja activa del archivo Excel
    $hoja = $excel->getActiveSheet();

    // Copiar los datos de la hoja activa del archivo Excel al archivo final
    $sheet->fromArray($hoja->toArray(), null, 'A' . ($sheet->getHighestRow() + 1));
}

// Establecer el ancho predeterminado de las columnas y la altura predeterminada de las filas
$sheet->getDefaultColumnDimension()->setWidth(30);
$sheet->getDefaultRowDimension()->setRowHeight(30);

// Guardar el archivo final
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('archivo_final.xlsx');

echo 'El archivo final se ha generado correctamente.';
?>
