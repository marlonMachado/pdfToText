<?php

require 'vendor/autoload.php'; // Carga el archivo autoload.php de Composer

use Smalot\PdfParser\Parser;
use Shuchkin\SimpleXLSX;


use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Crea una instancia de PhpSpreadsheet
$spreadsheet = new Spreadsheet();

// Obtén la hoja de cálculo activa
$sheet = $spreadsheet->getActiveSheet();
$pdfFilePath = 'D:\Documentos del Sistema\Desktop\CC79227.pdf';

// Crear un nuevo analizador PdfParser
$parser = new Parser();

$pdf = $parser->parseFile($pdfFilePath);

$text = $pdf->getText();

//echo $text;
echo 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';
echo '<br>';

echo 'Datos extraidos';
echo '<br>';


// $pattern = '/Fundación Clínica Nelson Mandela(.*?)Fundación Clínica Nelson Mandela/s';
$pattern =  '/Fundación Clínica Nelson Mandela(.*?)RECOMENDACIONES GENERALES PARA EL CUIDADO Y ADHERENCIA MEDICA/s';
$sheet->setCellValue('A1', 'fecha consulta');
$sheet->setCellValue('B1', 'nombre');
$sheet->setCellValue('C1', 'identifircacion');
$sheet->setCellValue('D1', 'colesterol total');
$sheet->setCellValue('E1', 'colesterolldl');
$sheet->setCellValue('F1', 'trigliceridos');
$sheet->setCellValue('G1', 'albuminuriaCreatinuria');
$sheet->setCellValue('H1', 'albuminaSerica');
$sheet->setCellValue('I1', 'fosforo');
$sheet->setCellValue('J1', 'pth');
$sheet->setCellValue('K1', 'hemoglobina');
$sheet->setCellValue('L1', 'uroanalisis');
$sheet->setCellValue('M1', 'hemoglobina en ayunas');
$sheet->getColumnDimension('A')->setWidth(30);
$sheet->getColumnDimension('B')->setWidth(30);
$sheet->getColumnDimension('C')->setWidth(30);
$sheet->getColumnDimension('D')->setWidth(30);
$sheet->getColumnDimension('E')->setWidth(30);
$sheet->getColumnDimension('F')->setWidth(30);
$sheet->getColumnDimension('G')->setWidth(30);
$sheet->getColumnDimension('H')->setWidth(30);
$sheet->getColumnDimension('I')->setWidth(30);
$sheet->getColumnDimension('J')->setWidth(30);
$sheet->getColumnDimension('K')->setWidth(30);
$sheet->getColumnDimension('L')->setWidth(30);
$sheet->getColumnDimension('M')->setWidth(30);
$contador = 2;
$datos = [];
if (preg_match_all($pattern, $text, $matches)) {
    foreach ($matches[1] as $match) {
        $patron_fecha_consulta = '/.*Edad/m';
        $patron_paciente = '/Paciente:\s*(.*?)\s*Identificación:/';
        $patron_identificacion = '/Identificación:\s*([^:]+(\sTel?))/';

        $patron_colesterolT = '/Colesterol\s+Total\s\d+/';
        $patron_colesterolHDL = '/Colesterol\s+HDL\s\d+/';
        $patron_trigliceridos = '/Trigliceridos\s\d+/';
        $patron_colesterolLDL = '/Colesterol\s+LDL\s\d+/';
        $patron_albuminuriaCreatinuria_actualizado = '/Albuminuria\s\/\s+Creatinuria\s+-\s\(en\s+este\s+campo\s+resgitrar\s+el\s+dato\s+mas\s+actualizado\)\s+\d+/m';
        $patron_albuminaSerica = '/Albumina\s+Serica\s+\d+/';
        $patron_fosforo = '/Fosforo\s+\(p\)\s+\d+/';
        $patron_uroanalisis = '/UROANALIS\s+O\s+PARCIAL\s+DE\s+ORINA\s+\d+/';
        $patron_pth = '/PTH\s+.Paratohormona.\s+\d+/';
        $patron_hemoglobina = '/Hemoglobina\s+\d+/';
        $patron_glicemia_ayuno = '/Glicemia\s+de\s+Ayuno\s+\d+/m';
        $textoCapturado = trim($match);
        //echo "Texto capturado: $textoCapturado\n";
        echo '<br>';

        $paciente = 'ND';
        if (preg_match($patron_paciente, $textoCapturado, $matches_paciente)) {
            $paciente = $matches_paciente[1];
            //echo "Paciente: $paciente\n";
        }
        $identificacion = 'ND';
        if (preg_match($patron_identificacion, $textoCapturado, $matches_identificacion)) {
            $identificacion = $matches_identificacion[1];
            $identificacion = str_replace("Tel", "", $identificacion);
            // echo "Identificación: $identificacion\n";
        }
        $fechaConsulta = 'ND';


        if (preg_match($patron_fecha_consulta, $textoCapturado, $matches_fcha)) {
            $fechaConsulta = $matches_fcha[0];
            $fechaConsulta = str_replace("Edad", "", $fechaConsulta);
            $fechaConsulta = str_replace("Edad", "", $fechaConsulta);
            $fechaConsulta = str_replace("Fecha", "", $fechaConsulta);
            $fechaConsulta = str_replace("de", "", $fechaConsulta);
            $fechaConsulta = str_replace("consulta:", "", $fechaConsulta);
            $posicion23 = strpos($fechaConsulta, "23"); // Encuentra la posición de "23"
            $fechaConsulta = substr($fechaConsulta, 0, $posicion23 + 2);
            // echo "fechaConsulta: $fechaConsulta\n";

        }

        $colesterolT = 'ND';
        if (preg_match($patron_colesterolT, $textoCapturado, $matches_colesterol)) {
            //print_r( $matches_colesterol);
            $colesterolT = $matches_colesterol[0];
            // echo "colesterolT: $colesterolT\n";
        }
        $colesterolldl = 'ND';
        if (preg_match($patron_colesterolHDL, $textoCapturado, $matches_colesterolHDL)) {
            $colesterolldl = $matches_colesterolHDL[0];
            //echo "colesterolldl: $colesterolldl\n";
        }
        $trigliceridos = 'ND';
        if (preg_match($patron_trigliceridos, $textoCapturado, $matches_trigliceridos)) {
            $trigliceridos = $matches_trigliceridos[0];
            // echo "trigliceridos: $trigliceridos\n";
        }
        $albuminuriaCreatinuria = 'ND';
        if (preg_match($patron_albuminuriaCreatinuria_actualizado, $textoCapturado, $matches_albuminuriaCreatinuria_actualizado)) {
            $albuminuriaCreatinuria = $matches_albuminuriaCreatinuria_actualizado[0];
            //   echo "albuminuriaCreatinuria: $albuminuriaCreatinuria\n";
        }
        $albuminaSerica = 'ND';
        if (preg_match($patron_albuminaSerica, $textoCapturado, $matches_albuminaSerica)) {
            $albuminaSerica = $matches_albuminaSerica[0];
            // echo "_albuminaSerica: $albuminaSerica\n";
        }
        $fosforo = 'ND';
        if (preg_match($patron_fosforo, $textoCapturado, $matches_fosforo)) {
            $fosforo = $matches_fosforo[0];
            //echo "_fosforo: $fosforo\n";
        }
        $pth = 'ND';
        if (preg_match($patron_pth, $textoCapturado, $matches_pth)) {
            $pth = $matches_pth[0];
            //echo "_pth: $pth\n";
        }
        $hemoglobina = 'ND';
        if (preg_match($patron_hemoglobina, $textoCapturado, $matches_hemoglobina)) {
            $hemoglobina = $matches_hemoglobina[0];
            //echo "_hemoglobina: $hemoglobina\n";
        }
        $uroanalisis = 'ND';
        if (preg_match($patron_uroanalisis, $textoCapturado, $matches_uroanalisis)) {
            $uroanalisis = $matches_uroanalisis[0];
            //echo "_uroanalisis: $uroanalisis\n";
        }
        $glicemiaAyuno = 'ND';
        if (preg_match($patron_glicemia_ayuno, $textoCapturado, $matches_glicemiaAyuno)) {
            $glicemiaAyuno = $matches_glicemiaAyuno[0];
            // echo "_glicemiaAyuno: $glicemiaAyuno\n";
        }
        echo '<br>';
        $datos[] = array(
            'Fecha_Consulta' => $fechaConsulta,
            'Paciente' => $paciente,
            'Identificacion' => $identificacion,
            'ColesterolTotal' => $colesterolT,
            'ColesterolLDL' => $colesterolldl,
            'Triglicéridos' => $trigliceridos,
            'AlbuminuriaCreatinuria' => $albuminuriaCreatinuria,
            'AlbuminaSerica' => $albuminaSerica,
            'Fósforo' => $fosforo,
            'PTH' => $pth,
            'Hemoglobina' => $hemoglobina,
            'Uroanálisis' => $uroanalisis,
            'Glicemia en Ayuno' => $glicemiaAyuno
        );
    }
    echo 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX  XXXXXXXXXXXXXXXXXXXx';
    $fechasMaximas = [];
    $identificacionesRepetidas = array_count_values(array_column($datos, 'Identificacion'));
    foreach ($datos as $key => $registro) {
        $identificacion = $registro["Identificacion"];
        $fechaConsulta = $registro["Fecha_Consulta"];
        if ($identificacionesRepetidas[$identificacion] > 1) {
            if (array_key_exists($identificacion, $fechasMaximas)) {
                if ($fechaConsulta > $fechasMaximas[$identificacion]) {
                    $fechasMaximas[$identificacion] = $fechaConsulta;
                } else {
                    unset($datos[$key]);
                }
            } else {
                $fechasMaximas[$identificacion] = $fechaConsulta;
            }
        }
    }
    foreach ($datos as $registro) {
        $sheet->setCellValue('A'.$contador, $registro['Fecha_Consulta']);
        $sheet->setCellValue('B'.$contador, $registro['Paciente']);
        $sheet->setCellValue('C'.$contador, $registro['Identificacion']);
        $sheet->setCellValue('D'.$contador, $registro['ColesterolTotal']);
        $sheet->setCellValue('E'.$contador, $registro['ColesterolLDL']);
        $sheet->setCellValue('F'.$contador, $registro['Triglicéridos']);
        $sheet->setCellValue('G'.$contador, $registro['AlbuminuriaCreatinuria']);
        $sheet->setCellValue('H'.$contador, $registro['AlbuminaSerica']);
        $sheet->setCellValue('I'.$contador, $registro['Fósforo']);
        $sheet->setCellValue('J'.$contador, $registro['PTH']);
        $sheet->setCellValue('K'.$contador, $registro['Hemoglobina']);
        $sheet->setCellValue('L'.$contador, $registro['Uroanálisis']);
        $sheet->setCellValue('M'.$contador, $registro['Glicemia en Ayuno']);
        
        // Incrementar el contador de filas
        $contador++;
    };
    // Configurar el tipo de archivo y el nombre de descarga

    $writer = new Xlsx($spreadsheet);

   // $writer->save('ejemplo1.xlsx');

    echo "El archivo ejemplo.xlsx se ha generado correctamente.";
} else {
    echo "No se encontró ningún texto entre las instancias de Fundación Clínica Nelson Mandela.";
}
