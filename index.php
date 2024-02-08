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
echo'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';
echo '<br>';

echo 'Datos extraidos';
echo'<br>';


// $pattern = '/Fundación Clínica Nelson Mandela(.*?)Fundación Clínica Nelson Mandela/s';
$pattern =  '/Fundación Clínica Nelson Mandela(.*?)RECOMENDACIONES GENERALES PARA EL CUIDADO Y ADHERENCIA MEDICA/s';
$sheet->setCellValue('A1', 'Fecha Consulta');
$sheet->setCellValue('B1', 'Nombre');
$sheet->setCellValue('C1', 'Identificacion');
$sheet->setCellValue('D1', 'Colesterol Total');
$sheet->setCellValue('E1', 'Colesterol Fecha');
$sheet->setCellValue('F1', 'Colesterol Ldl');
$sheet->setCellValue('G1', 'Fecha Colesterol Ldl');
$sheet->setCellValue('H1', 'Colesterol Hdl');
$sheet->setCellValue('I1', 'Fecha Hdl');
$sheet->setCellValue('J1', 'Trigliceridos');
$sheet->setCellValue('K1', 'Trigliceridos fecha');
$sheet->setCellValue('L1', 'Albuminuria/Creatinuria');
$sheet->setCellValue('M1', 'Albuminuria/Creatinuria fecha');
$sheet->setCellValue('N1', 'Creatinina actual');
$sheet->setCellValue('O1', 'Fecha creatinina');
$sheet->setCellValue('P1', 'Albumina Serica');
$sheet->setCellValue('Q1', 'Fecha Albumina Serica');
$sheet->setCellValue('R1', 'Fosforo');
$sheet->setCellValue('S1', 'Fosforo fecha');
$sheet->setCellValue('T1', 'Pth');
$sheet->setCellValue('U1', 'Fecha Pth');
$sheet->setCellValue('V1', 'Hemoglobina');
$sheet->setCellValue('W1', 'Hemoglobina Fecha');
$sheet->setCellValue('X1', 'Hemoglobina Glicosilada');
$sheet->setCellValue('Y1', 'Fecha Hemoglobina Glicosilada');
$sheet->setCellValue('Z1', 'Uroanalisis');
$sheet->setCellValue('AA1', 'Fecha  uroanalisis');
$sheet->setCellValue('AB1', 'Glicemia en ayuno');
$sheet->setCellValue('AC1', 'Glicemia en ayuno fecha');
$sheet->setCellValue('AD1', 'Fecha última atencion presencial en IPS por Medicina Interna');
$sheet->setCellValue('AE1', 'Fecha última atencion presencial en IPS por Endocrinologia');
$sheet->setCellValue('AF1', 'Fecha última atencion presencial en IPS por Cardiologia');
$sheet->setCellValue('AG1', 'Fecha última atencion presencial en IPS pór Oftalmologia');
$sheet->setCellValue('AH1', 'Fecha última atencion presencial en IPS por Nefrologia');
$sheet->setCellValue('AI1', 'Fecha de última Valoracion presencial en IPS por Psicologia');
$sheet->setCellValue('AJ1', 'Fecha de última valoracion presencial en IPS por Nutricion');
$sheet->setCellValue('AK1', 'Fecha de última valoracion presencial en IPS por Trabajo Social');
$sheet->setCellValue('AL1', 'Fecha de ultima consulta por medico general en la IPS ');
$sheet->setCellValue('AM1', 'Fecha deTeleconsulta por medico general');
$sheet->setCellValue('AN1', 'Fecha de Visita domiciliaria por medico general ');
$sheet->setCellValue('AO1', 'Fecha de Teleconsulta por medicina especializada ');
$sheet->setCellValue('AP1', 'Fecha de Telemedicina por especialidad ');  
$sheet->setCellValue('AQ1', 'Fecha de Visita domiciliaria por medicina especializada ');
$sheet->setCellValue('AR1', 'Fecha de Visita domiciliaria por promotor de salud');
$sheet->setCellValue('AS1', 'Fecha Seguimiento telefónico por auxiliar de enfermeria');
$sheet->setCellValue('AT1', 'Fecha de Visita domiciliaria por auxiliar de enfermera ');
$sheet->setCellValue('AU1', 'Fecha Seguimiento telefónico por enfermeria');
$sheet->setCellValue('AV1', 'Fecha de Visita domiciliaria por enfermera');
$sheet->setCellValue('AW1', 'Fecha de Visita domiciliaria por otro profesional ');
$sheet->setCellValue('AX1', 'Fecha de Visita domiciliaria por equipo interdisciplinario');
$sheet->setCellValue('AY1', 'Fecha deToma de laboratorios en domicilio');
$sheet->setCellValue('AZ1', 'Fecha de Entrega de medicamentos -domicilio');
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
$sheet->getColumnDimension('N')->setWidth(30);
$sheet->getColumnDimension('O')->setWidth(30);
$sheet->getColumnDimension('P')->setWidth(30);
$sheet->getColumnDimension('Q')->setWidth(30);
$sheet->getColumnDimension('R')->setWidth(30);
$sheet->getColumnDimension('S')->setWidth(30);
$sheet->getColumnDimension('T')->setWidth(30);
$sheet->getColumnDimension('U')->setWidth(30);
$sheet->getColumnDimension('V')->setWidth(30);
$sheet->getColumnDimension('W')->setWidth(30);
$sheet->getColumnDimension('X')->setWidth(30);
$sheet->getColumnDimension('Y')->setWidth(30);
$sheet->getColumnDimension('Z')->setWidth(30);
$sheet->getColumnDimension('AA')->setWidth(30);
$sheet->getColumnDimension('AB')->setWidth(30);
$sheet->getColumnDimension('AC')->setWidth(30);
$sheet->getColumnDimension('AD')->setWidth(30);
$sheet->getColumnDimension('AE')->setWidth(30);
$sheet->getColumnDimension('AF')->setWidth(30);
$sheet->getColumnDimension('AG')->setWidth(30);
$sheet->getColumnDimension('AH')->setWidth(30);
$sheet->getColumnDimension('AI')->setWidth(30);
$sheet->getColumnDimension('AJ')->setWidth(30);
$sheet->getColumnDimension('AK')->setWidth(30);
$sheet->getColumnDimension('AL')->setWidth(30);
$sheet->getColumnDimension('AM')->setWidth(30);
$sheet->getColumnDimension('AN')->setWidth(30);
$sheet->getColumnDimension('AO')->setWidth(30);
$sheet->getColumnDimension('AP')->setWidth(30);
$sheet->getColumnDimension('AQ')->setWidth(30);
$sheet->getColumnDimension('AR')->setWidth(30);
$sheet->getColumnDimension('AS')->setWidth(30);
$sheet->getColumnDimension('AT')->setWidth(30);
$sheet->getColumnDimension('AU')->setWidth(30);
$sheet->getColumnDimension('AV')->setWidth(30);
$sheet->getColumnDimension('AW')->setWidth(30);
$sheet->getColumnDimension('AX')->setWidth(30);
$sheet->getColumnDimension('AY')->setWidth(30);
$sheet->getColumnDimension('AZ')->setWidth(30);

$contador=2;
if (preg_match_all($pattern, $text, $matches)) {
    foreach ($matches[1] as $match) {
        //$patron_fecha_consulta = '/Fecha\s+de\s+consulta.\s+\d{1,2}\/\d{1,2}\/\d{4}\s+\d+:\d+\s+Edad/m';
        //$patron_fecha_consulta = '/Fecha de consulta: 29\/05\/2023 10:33 Edad/m';
       // $patron_fecha_consulta = '/Fecha de consulta: (\d{2}\/\d{2}\/\d{4})/';
        $patron_fecha_consulta = '/.*Edad/m';
        $patron_paciente = '/Paciente:\s*(.*?)\s*Identificación:/';
        $patron_identificacion = '/Identificación:\s*([^:]+(\sTel?))/';
        
        //$patron_colesterolT ='/Colesterol\s+Total\s\d+/';-----
        $patron_colesterolT ='/Colesterol\s+Total\s(\d+\.\d+|\d+)/';
        // $patron_colesterolHDL = '/Colesterol\s+HDL\s\d+/';-----
        $patron_colesterolHDL = '/Colesterol\s+HDL\s(\d+\.\d+|\d+)/';
        //$patron_trigliceridos = '/Trigliceridos\s\d+/';----
        $patron_trigliceridos = '/Trigliceridos\s(\d+\.\d+|\d+)/';
        //$patron_colesterolLDL = '/Colesterol\s+LDL\s\d+/';--------
        $patron_colesterolLDL = '/Colesterol\s+LDL\s(\d+\.\d+|\d+)/';
        //$patron_albuminuriaCreatinuria_actualizado ='/Albuminuria\s\/\s+Creatinuria\s+-\s\(en\s+este\s+campo\s+resgitrar\s+el\s+dato\s+mas\s+actualizado\)\s+\d+/m';-----
        $patron_albuminuriaCreatinuria_actualizado ='/Albuminuria\s\/\s+Creatinuria\s+-\s\(en\s+este\s+campo\s+resgitrar\s+el\s+dato\s+mas\s+actualizado\)\s+(\d+\.\d+|\d+)/m';
        //$patron_creatinina_actual = '/Creatinina\s+ACTUAL\s+(\d+.\d+|\d)/m';----
        $patron_creatinina_actual = '/Creatinina\s+ACTUAL\s+(\d+.\d+|\d)/m';
        // $patron_albuminaSerica ='/Albumina\s+Serica\s+\d+/';-------
        $patron_albuminaSerica ='/Albumina\s+Serica\s+(\d+\.\d+|\d+)/';
        // $patron_fosforo = '/Fosforo\s+\(p\)\s+\d+/';---------
        $patron_fosforo = '/Fosforo\s+\(p\)\s+(\d+\.\d+|\d+)/';
        // $patron_uroanalisis ='/UROANALIS\s+O\s+PARCIAL\s+DE\s+ORINA\s+\d+/';-----
        $patron_uroanalisis ='/UROANALIS\s+O\s+PARCIAL\s+DE\s+ORINA\s+(\d+\.\d+|\d+)/';
        //$patron_pth = '/PTH\s+.Paratohormona.\s+\d+/';
        $patron_pth = '/PTH\s+.Paratohormona.\s+(\d+\.\d+|\d+)/m';
        // $patron_hemoglobina ='/Hemoglobina\s+\d+/';---
        $patron_hemoglobina ='/Hemoglobina\s+(\d+\.\d+|\d+)/';
        $patron_hemoglobina_glico ='/Hemoglobina\s+Glicosilada\s+\(HbA1c\)\s+(\d+.\d+|\d)/m';
        // $patron_glicemia_ayuno = '/Glicemia\s+de\s+Ayuno\s+\d+/m';---
        $patron_glicemia_ayuno = '/Glicemia\s+de\s+Ayuno\s+(\d+\.\d+|\d+)/m';
        $patron_ultima_atencion_p = '/Fecha\s+última\s+atencion\s+presencial\s+en\s+IPS\s+por\s+Medicina\s+Interna\s+(\d{4}-\d+-\d+)/m';
        $patrorn_medicina_interna = '/Fecha\s+última\s+atencion\s+presencial\s+en\s+IPS\s+por\s+Medicina\s+Interna\s+(\d{4}-\d+-\d+)/';
        $patrorn_endo = '/Fecha\s+última\s+atencion\s+presencial\s+en\s+IPS\s+por\s+Endocrinologia\s+(\d{4}-\d+-\d+)/m';
        $patrorn_cardio = '/Fecha\s+última\s+atencion\s+presencial\s+en\s+IPS\s+por\s+Cardiologia\s+(\d{4}-\d+-\d+)/m';
        $patrorn_ofta = '/(\d+.\d+.\d+)\s+Fecha\s+última\s+atencion\s+presencial\s+en\s+IPS\s+por\s+Nefrologia/m';
        $patrorn_nefro = '/Fecha\s+última\s+atencion\s+presencial\s+en\s+IPS\s+por\s+Nefrologia\s+(\d+-\d+-\d+)/m';
        $patrorn_psico = '/Fecha\s+de\s+última\s+Valoracion\s+presencial\s+en\s+IPS\s+por\s+Psicologia\s+(\d{4}-\d+-\d+)/m';
        $patrorn_nutricion = '/Fecha\s+de\s+última\s+valoracion\s+presencial\sen\sIPS\s+por\s+Nutricion\s+(\d{4}-\d+-\d+)/m';
        $patrorn_trabajo_social = '/Fecha\s+de\s+última\s+valoracion\s+presencial\s+en\s+IPS\s+por\s+Trabajo\s+Social\s+(\d{4}-\d+-\d+)/m';
        $patrorn_medico_general = '/Fecha\s+de\s+ultima\s+consulta\s+por\s+medico\s+general\s+en\s+la\s+IPS\s+(\d{4}-\d+-\d+)/m';
        $patrorn_teleconsulta_medico_hgeneral = '/Fecha\s+deTeleconsulta\s+por\s+medico\sgeneral\s+(\d{4}-\d+-\d+)/m';
        $patrorn_domociliario_medico_general = '/Fecha\s+de\s+Visita\s+domiciliaria\s+por\s+medico\s+general\s+(\d{4}-\d+-\d+)/m';
        $patrorn_telc_medicina_esp = '/Fecha\s+de\s+Teleconsulta\s+por\s+medicina\s+especializada\s+(\d{4}-\d+-\d+)/m';
        $patrorn_telemedicina_esp = '/Fecha\s+de\s+Telemedicina\s+por\s+especialidad\s+(\d{4}-\d+-\d+)/m';
        $patrorn_domiciliario_medicina_espec = '/Fecha\s+de\s+Visita\s+domiciliaria\s+por\s+medicina\s+especializada\s+(\d{4}-\d+-\d+)/m';
        $patrorn_domociliario_promotor_salud = '/Fecha\s+de\s+Visita\s+domiciliaria\s+por\s+promotor\s+de\s+salud\s+(\d{4}-\d+-\d+)/m';
        $patrorn_seguimiento_tel_enfermeria = '/Fecha\s+Seguimiento\s+telefónico\s+por\s+auxiliar\s+de\s+enfermeria\s+(\d{4}-\d+-\d+)/m';
        $patrorn_visita_domiciliaria_aux_enfermeria = '/Fecha\s+de\s+Visita\s+domiciliaria\s+por\s+auxiliar\s+de\s+enfermera\s+(\d{4}-\d+-\d+)/m';
        $patrorn_segumiento_tel_enfermera = '/Fecha\s+Seguimiento\s+telefónico\s+por\s+enfermeria\s+(\d{4}-\d+-\d+)/m';
        $patrorn_visita_dom_enfermeria = '/Fecha\s+de\s+Visita\s+domiciliaria\s+por\s+enfermera\s+(\d{4}-\d+-\d+)/m';
        $patrorn_visita_dom_otro_profesional = '/Fecha\s+de\s+Visita\s+domiciliaria\s+por\s+otro\s+profesional\s+(\d{4}-\d+-\d+)/m';
        $patrorn_visita_dom_equipo_inter = '/Fecha\s+de\s+Visita\s+domiciliaria\s+por\s+equipo\s+interdisciplinario\s+(\d{4}-\d+-\d+)/m';
        $patrorn_toma_laboratorio_dom = '/Fecha\s+deToma\s+de\s+laboratorios\s+en\s+domicilio\s+(\d{4}-\d+-\d+)/m';
        $patrorn_entrega_medicamentos_dom = '/Fecha\s+de\s+Entrega\s+de\s+medicamentos\s+-domicilio\s+(\d{4}-\d+-\d+)/m';
        ///fechas examenes
        $patron_fecha_toma_colesterol = '/Fecha\s+Toma\s+CT\s+(\d+-\d+-\d+)/m';
        $patron_fecha_toma_colesterol_hdl = '/Fecha\s+Toma\s+HDL\s+(\d+-\d+-\d+)/m';
        $patron_fecha_toma_trigliceridos = '/Fecha\s+Toma\s+trigliceridos\s+(\d+-\d+-\d+)/m';
        $patron_fecha_toma_colesterol_ldl = '/Fecha\s+Toma\s+LDL\s+(\d+-\d+-\d+)/m';
        $patron_fecha_toma_albimnuriaCreatinuria ='/(\d+-\d+-\d+)\s+Albuminuria\s+.\s+Creatinuria\s+.\s+.anterior/m';
        $patron_fecha_toma_creatinina = '/Fecha\s+de\s+la\s+.ltima\s+creatinina\s+(\d+-\d+-\d+)/m';
        $patron_fecha_toma_albumina_serica = '/Fecha\s+de\s+la\s+última\s+Albumina\s+Serica\s+(\d+-\d+-\d+)/m';
        //$patron_fecha_toma_fosforo = '/Fecha\s+de\s+la\s+.ltimo\s+f.sforo\s+(\d+-\d+-\d+)/m';
        $patron_fecha_toma_fosforo = '/sforo\s+(\d+-\d+-\d+)/m';;
        $patron_fecha_toma_uroanailis = '/Fecha\s+del\s+uroanalisis\s+o\s+parcial\s+de\s+orina\s+(\d+-\d+-\d+)/m';
        $patron_fecha_toma_pth = '/Fecha\s+de\s+la\s+última\s+PTH\s+(\d+-\d+-\d+)/m';
        $patron_fecha_creatinina_actual = '/Fecha\s+de\s+la\s+ultima\s+creatinina\s+(\d+-\d+-\d+)/m';
        $patron_fecha_toma_hemoglobinaGlico = '/Fecha Toma la hemoglobina glicosilada \(HbA1c\)\s+\s+(\d+-\d+-\d+)/m';
        $patron_fecha_toma_hemoglobina = '/Fecha\s+de\s+la\s+última\s+hemoglobina\s+(\d+-\d+-\d+)/m';
        $patron_fecha_toma_glicemia_ayuno = '/fecha\s+de\s+la\s+Toma\s+Glicemia\s+de\s+ayuno\s+(\d+-\d+-\d+)/m';

        $textoCapturado = trim($match);
        //echo "Texto capturado: $textoCapturado\n";
        echo'<br>';
        $paciente='ND';
        if (preg_match($patron_paciente, $textoCapturado, $matches_paciente)) {
            $paciente = $matches_paciente[1];
            //echo "Paciente: $paciente\n";
        }
        $identificacion='ND';
        if (preg_match($patron_identificacion, $textoCapturado, $matches_identificacion)) {
            $identificacion = $matches_identificacion[1];
            $identificacion = str_replace("Tel", "", $identificacion);
            //echo "Identificación: $identificacion\n";
        }
        $fechaConsulta='ND';
        
        
        if (preg_match($patron_fecha_consulta, $textoCapturado, $matches_fcha)) {
            $fechaConsulta = $matches_fcha[0];
            $fechaConsulta = str_replace("Edad", "", $fechaConsulta);
            $fechaConsulta = str_replace("Edad", "", $fechaConsulta);
            $fechaConsulta = str_replace("Fecha", "", $fechaConsulta);
            $fechaConsulta = str_replace("de", "", $fechaConsulta);
            $fechaConsulta = str_replace("consulta:", "", $fechaConsulta);
            $posicion23 = strpos($fechaConsulta, "23"); // Encuentra la posición de "23"
            $fechaConsulta = substr($fechaConsulta, 0, $posicion23 + 2);
        }
        
        $colesterolT ='ND';
        if (preg_match($patron_colesterolT, $textoCapturado, $matches_colesterol)) {
            //print_r( $matches_colesterol);
            $colesterolT = $matches_colesterol[0];
            //echo "colesterolT: $colesterolT\n";
        }
        $colesterolldl='ND';
        if (preg_match($patron_colesterolLDL, $textoCapturado, $matches_colesterolLDL)) {
            $colesterolldl = $matches_colesterolLDL[0];
            //echo "colesterolldl: $colesterolldl\n";
        }
        $colesterolhdl='ND';
        if (preg_match($patron_colesterolHDL, $textoCapturado, $matches_colesterolHDL)) {
            $colesterolhdl = $matches_colesterolHDL[0];
            
        }
        $trigliceridos='ND';
        if (preg_match($patron_trigliceridos, $textoCapturado, $matches_trigliceridos)) {
            $trigliceridos = $matches_trigliceridos[0];
            //echo "trigliceridos: $trigliceridos\n";
        }
        $albuminuriaCreatinuria='ND';
        if( preg_match($patron_albuminuriaCreatinuria_actualizado, $textoCapturado, $matches_albuminuriaCreatinuria_actualizado)){
            $albuminuriaCreatinuria = $matches_albuminuriaCreatinuria_actualizado[0];
            //echo "albuminuriaCreatinuria: $albuminuriaCreatinuria\n";
        }
        $creatininaActual='ND';
        if( preg_match($patron_creatinina_actual, $textoCapturado, $matches__creatinina_actual)){
            $creatininaActual = $matches__creatinina_actual[1];
            echo "creatininaActual: $creatininaActual\n";
        }
       $albuminaSerica='ND';
       if (preg_match($patron_albuminaSerica, $textoCapturado, $matches_albuminaSerica)) {
           $albuminaSerica = $matches_albuminaSerica[0];
        //echo "_albuminaSerica: $albuminaSerica\n";
    }
        $fosforo='ND';
        if (preg_match($patron_fosforo, $textoCapturado, $matches_fosforo)) {
            $fosforo = $matches_fosforo[0];
            //echo "_fosforo: $fosforo\n";
        }
        $pth='ND';
        if (preg_match($patron_pth, $textoCapturado, $matches_pth)) {
            $pth = $matches_pth[0];
            //echo "_pth: $pth\n";
        }
        $hemoglobina='ND';
        if (preg_match($patron_hemoglobina, $textoCapturado, $matches_hemoglobina)) {
            $hemoglobina = $matches_hemoglobina[0];
            //echo "_hemoglobina: $hemoglobina\n";
        }
        $hemoglobina_glico = 'ND';
        if (preg_match($patron_hemoglobina_glico, $textoCapturado, $matches_hemoglobina_glico)) {
            $hemoglobina_glico = $matches_hemoglobina_glico[1];
            //echo "_hemoglobina: $hemoglobina\n";
        }
        $uroanalisis='ND';
        if (preg_match($patron_uroanalisis, $textoCapturado, $matches_uroanalisis)) {
            $uroanalisis = $matches_uroanalisis[0];
            //echo "_uroanalisis: $uroanalisis\n";
        }
        $glicemiaAyuno='ND';
        if (preg_match($patron_glicemia_ayuno, $textoCapturado, $matches_glicemiaAyuno)) {
            $glicemiaAyuno = $matches_glicemiaAyuno[0];
            //echo "_glicemiaAyuno: $glicemiaAyuno\n";
        }
        $ultima_atencion_p ='ND';
        if (preg_match($patron_ultima_atencion_p, $textoCapturado, $matches_ultima_a)) {
            //print_r( $matches_colesterol);
            $ultima_atencion_p = $matches_ultima_a[1];
            //echo "patron_ultima_atencion_p: $ultima_atencion_p\n";
        }
        $ultima_medicinaIn ='ND';
        if (preg_match($patrorn_medicina_interna, $textoCapturado, $matches_medicina)) {
            //print_r( $matches_colesterol);
            $ultima_medicinaIn = $matches_medicina[1];
            //echo " ultima_medicinaIn: $ultima_medicinaIn\n";
        }
        $endo ='ND';
        if (preg_match($patrorn_endo, $textoCapturado, $matches_endo)) {
            //print_r( $matches_colesterol);
            $endo = $matches_endo[1];
            //echo " endo: $endo\n";
        }
        $cardio ='ND';
        if (preg_match($patrorn_cardio, $textoCapturado, $matches_cardio)) {
            //print_r( $matches_colesterol);
            $cardio = $matches_cardio[1];
            //echo " cardio: $cardio\n";
        }
        $ofta ='ND';//no me toma
        if (preg_match($patrorn_ofta, $textoCapturado, $matches_ofta)) {
            //print_r( $matches_colesterol);
            $ofta = $matches_ofta[1];
            //echo " ofta: $ofta\n";
        }
        $nefro ='ND';
        if (preg_match($patrorn_nefro, $textoCapturado, $matches_nefro)) {
            //print_r( $matches_colesterol);
            $nefro = $matches_nefro[1];
            //echo " nefro: $nefro\n";
        }
        $psico ='ND';
        if (preg_match($patrorn_psico, $textoCapturado, $matches_psico)) {
            //print_r( $matches_colesterol);
            $psico = $matches_psico[1];
            //echo " psico: $psico\n";
        }
        $nutricion='ND';
        if (preg_match($patrorn_nutricion, $textoCapturado, $matches_nutricion)) {
            //print_r( $matches_colesterol);
            $nutricion = $matches_nutricion[1];
            //echo " nutricion: $nutricion\n";
        }
        $trabajo_social='ND';
        if (preg_match($patrorn_trabajo_social, $textoCapturado, $matches_trabajo_social)) {
            //print_r( $matches_colesterol);
            $trabajo_social = $matches_trabajo_social[1];
            //echo " trabajo_social: $trabajo_social\n";
        }
        $medico_general='ND';
        if (preg_match($patrorn_medico_general, $textoCapturado, $matches_medico_general)) {
            //print_r( $matches_colesterol);
            $medico_general = $matches_medico_general[1];
            //echo " medico_general: $medico_general\n";
        }
        $teleconsulta_medico_hgeneral='ND';
        if (preg_match($patrorn_teleconsulta_medico_hgeneral, $textoCapturado, $matches_teleconsulta_medico_hgeneral)) {
            //print_r( $matches_colesterol);
            $teleconsulta_medico_hgeneral = $matches_teleconsulta_medico_hgeneral[1];
            //echo " teleconsulta_medico_hgeneral: $teleconsulta_medico_hgeneral\n";
        }
        $domociliario_medico_general='ND';
        if (preg_match($patrorn_domociliario_medico_general, $textoCapturado, $matches_domociliario_medico_general)) {
            //print_r( $matches_colesterol);
            $domociliario_medico_general = $matches_domociliario_medico_general[1];
            //echo " domociliario_medico_general: $domociliario_medico_general\n";
        }
        $_telc_medicina_esp='ND';
        if (preg_match($patrorn_telc_medicina_esp, $textoCapturado, $matches__telc_medicina_esp)) {
            //print_r( $matches_colesterol);
            $_telc_medicina_esp = $matches__telc_medicina_esp[1];
            //echo " _telc_medicina_esp: $_telc_medicina_esp\n";
        }
        $telemedicina_esp='ND';
        if (preg_match($patrorn_telemedicina_esp, $textoCapturado, $matches__telemedicina_esp)) {
            //print_r( $matches_colesterol);
            $telemedicina_esp = $matches__telemedicina_esp[1];
            //echo " telemedicina_esp: $telemedicina_esp\n";
        }
        $domiciliario_medicina_espec='ND';
        if (preg_match($patrorn_domiciliario_medicina_espec, $textoCapturado, $matches__domiciliario_medicina_espec)) {
            //print_r( $matches_colesterol);
            $domiciliario_medicina_espec = $matches__domiciliario_medicina_espec[1];
            //echo " domiciliario_medicina_espec: $domiciliario_medicina_espec\n";
        }
        $domociliario_promotor_salud='ND';
        if (preg_match($patrorn_domociliario_promotor_salud, $textoCapturado, $matches__domociliario_promotor_salud)) {
            //print_r( $matches_colesterol);
            $domociliario_promotor_salud = $matches__domociliario_promotor_salud[1];
            //echo " domociliario_promotor_salud: $domociliario_promotor_salud\n";
        }
        $seguimiento_tel_enfermeria='ND';
        if (preg_match($patrorn_seguimiento_tel_enfermeria, $textoCapturado, $matches__seguimiento_tel_enfermeria)) {
            //print_r( $matches_colesterol);
            $seguimiento_tel_enfermeria = $matches__seguimiento_tel_enfermeria[1];
            //echo " seguimiento_tel_enfermeria: $seguimiento_tel_enfermeria\n";
        }
        $visita_domiciliaria_aux_enfermeria='ND';
        if (preg_match($patrorn_visita_domiciliaria_aux_enfermeria, $textoCapturado, $matches__visita_domiciliaria_aux_enfermeria)) {
            //print_r( $matches_colesterol);
            $visita_domiciliaria_aux_enfermeria = $matches__visita_domiciliaria_aux_enfermeria[1];
            //echo " visita_domiciliaria_aux_enfermeria: $visita_domiciliaria_aux_enfermeria\n";
        }
        $_segumiento_tel_enfermera='ND';
        if (preg_match($patrorn_segumiento_tel_enfermera, $textoCapturado, $matches___segumiento_tel_enfermera)) {
            //print_r( $matches_colesterol);
            $_segumiento_tel_enfermera = $matches___segumiento_tel_enfermera[1];
            //echo " _segumiento_tel_enfermera: $_segumiento_tel_enfermera\n";
        }
        $_visita_dom_enfermeria='ND';
        if (preg_match($patrorn_visita_dom_enfermeria, $textoCapturado, $matches_visita_dom_enfermeria)) {
            //print_r( $matches_colesterol);
            $_visita_dom_enfermeria = $matches_visita_dom_enfermeria[1];
            //echo " _visita_dom_enfermeria: $_visita_dom_enfermeria\n";
        }
        $_visita_dom_otro_profesional='ND';
        if (preg_match($patrorn_visita_dom_otro_profesional, $textoCapturado, $matches_visita_dom_otro_profesional)) {
            //print_r( $matches_colesterol);
            $_visita_dom_otro_profesional = $matches_visita_dom_otro_profesional[1];
            //echo " _visita_dom_otro_profesional: $_visita_dom_otro_profesional\n";
        }
        $visita_dom_equipo_inter='ND';
        if (preg_match($patrorn_visita_dom_equipo_inter, $textoCapturado, $matches_visita_dom_equipo_inter)) {
            //print_r( $matches_colesterol);
            $visita_dom_equipo_inter = $matches_visita_dom_equipo_inter[1];
            //echo " visita_dom_equipo_inter: $visita_dom_equipo_inter\n";
        }
        $toma_laboratorio_dom='ND';
        if (preg_match($patrorn_toma_laboratorio_dom, $textoCapturado, $matches_toma_laboratorio_dom)) {
            //print_r( $matches_colesterol);
            $toma_laboratorio_dom = $matches_toma_laboratorio_dom[1];
            //echo " toma_laboratorio_dom: $toma_laboratorio_dom\n";
        }
        $_entrega_medicamentos_dom='ND';
        if (preg_match($patrorn_entrega_medicamentos_dom, $textoCapturado, $matches_entrega_medicamentos_dom)) {
            //print_r( $matches_colesterol);
            $_entrega_medicamentos_dom = $matches_entrega_medicamentos_dom[1];
            //echo " _entrega_medicamentos_dom: $_entrega_medicamentos_dom\n";
        }
        $fecha_toma_coleterol='ND';
        if (preg_match($patron_fecha_toma_colesterol, $textoCapturado, $matches_fecha_toma_colesterol)) {
            //print_r( $matches_colesterol);
            $fecha_toma_coleterol = $matches_fecha_toma_colesterol[1];
            echo " _fecha_toma_coleterol: $fecha_toma_coleterol\n";
        }
        $fecha_toma_colesterol_hdl='ND';
        if (preg_match($patron_fecha_toma_colesterol_hdl, $textoCapturado, $matches_fecha_toma_coleterol_hdl)) {
            //print_r( $matches_colesterol);
            $fecha_toma_colesterol_hdl = $matches_fecha_toma_coleterol_hdl[1];
            echo " fecha_toma_coleterol_hdl: $fecha_toma_colesterol_hdl\n";
        }
        $fecha_toma_trigliceridos='ND';
        if (preg_match($patron_fecha_toma_trigliceridos, $textoCapturado, $matches_fecha_toma_trigliceridos)) {
            //print_r( $matches_colesterol);
            $fecha_toma_trigliceridos = $matches_fecha_toma_trigliceridos[1];
            echo " fecha_toma_trigliceridos: $fecha_toma_trigliceridos\n";
        }
        $fecha_toma_colesterol_ldl='ND';
        if (preg_match($patron_fecha_toma_colesterol_ldl, $textoCapturado, $matches_fecha_toma_colesterol_ldl)) {
            //print_r( $matches_colesterol);
            $fecha_toma_colesterol_ldl = $matches_fecha_toma_colesterol_ldl[1];
            echo " fecha_toma_colesterol_ldl: $fecha_toma_colesterol_ldl\n";
        }
        $fecha_toma_albimnuriaCreatinuria='ND';
        if (preg_match($patron_fecha_toma_albimnuriaCreatinuria, $textoCapturado, $matches_fecha_toma_albimnuriaCreatinuria)) {
            //print_r( $matches_colesterol);
            $fecha_toma_albimnuriaCreatinuria = $matches_fecha_toma_albimnuriaCreatinuria[1];
            echo " fecha_toma_albimnuriaCreatinuria: $fecha_toma_albimnuriaCreatinuria\n";
        }
        $fecha_toma_creatinina='ND';
        if (preg_match($patron_fecha_creatinina_actual, $textoCapturado, $matches_fecha_toma_creatinina)) {
            //print_r( $matches_colesterol);
            $fecha_toma_creatinina = $matches_fecha_toma_creatinina[1];
            echo " fecha_toma_creatinina: $fecha_toma_creatinina\n";
        }
        $fecha_toma_albumina_serica='ND';
        if (preg_match($patron_fecha_toma_albumina_serica, $textoCapturado, $matches_fecha_toma_albumina_serica)) {
            //print_r( $matches_colesterol);
            $fecha_toma_albumina_serica = $matches_fecha_toma_albumina_serica[1];
            echo " fecha_toma_albumina_serica: $fecha_toma_albumina_serica\n";
        }
        $fecha_toma_fosforo='ND';
        if (preg_match($patron_fecha_toma_fosforo, $textoCapturado, $matches_fecha_toma_fosforo)) {
            //print_r( $matches_colesterol);
            $fecha_toma_fosforo = $matches_fecha_toma_fosforo[1];
            echo " fecha_toma_fosforo: $fecha_toma_fosforo\n";
        }
        $fecha_toma_uroanailis='ND';
        if (preg_match($patron_fecha_toma_uroanailis, $textoCapturado, $matches_fecha_toma_uroanailis)) {
            //print_r( $matches_colesterol);
            $fecha_toma_uroanailis = $matches_fecha_toma_uroanailis[1];
            echo " fecha_toma_uroanailis: $fecha_toma_uroanailis\n";
        }
        $fecha_toma_pth='ND';
        if (preg_match($patron_fecha_toma_pth, $textoCapturado, $matches_fecha_toma_pth)) {
            //print_r( $matches_colesterol);
            $fecha_toma_pth = $matches_fecha_toma_pth[1];
            echo " fecha_toma_pth: $fecha_toma_pth\n";
        }
        $fecha_toma_hemoglobinaGlico='ND';
        if (preg_match($patron_fecha_toma_hemoglobinaGlico, $textoCapturado, $matches_fecha_toma_hemoglobinaGlico)) {
            //print_r( $matches_colesterol);
            $fecha_toma_hemoglobinaGlico = $matches_fecha_toma_hemoglobinaGlico[1];
            echo " fecha_toma_hemoglobinaGlico: $fecha_toma_hemoglobinaGlico\n";
        }
        $fecha_toma_hemoglobina='ND';
        if (preg_match($patron_fecha_toma_hemoglobina, $textoCapturado, $matches_fecha_toma_hemoglobina)) {
            //print_r( $matches_colesterol);
            $fecha_toma_hemoglobina = $matches_fecha_toma_hemoglobina[1];
            echo " fecha_toma_hemoglobina: $fecha_toma_hemoglobina\n";
        }
        $fecha_toma_glicemia_ayuno='ND';
        if (preg_match($patron_fecha_toma_glicemia_ayuno, $textoCapturado, $matches_fecha_toma_glicemia_ayuno)) {
            //print_r( $matches_colesterol);
            $fecha_toma_glicemia_ayuno = $matches_fecha_toma_glicemia_ayuno[1];
            echo " fecha_toma_glicemia_ayuno: $fecha_toma_glicemia_ayuno\n";
        }
        echo'<br>';
        
        $sheet->setCellValue('A'.$contador, $fechaConsulta);
        $sheet->setCellValue('B'.$contador, $paciente);
        $sheet->setCellValue('C'.$contador, $identificacion);
        $sheet->setCellValue('D'.$contador, $colesterolT);
        $sheet->setCellValue('E'.$contador, $fecha_toma_coleterol);
        $sheet->setCellValue('F'.$contador, $colesterolldl);
        $sheet->setCellValue('G'.$contador, $fecha_toma_colesterol_ldl);
        $sheet->setCellValue('H'.$contador, $colesterolhdl);
        $sheet->setCellValue('I'.$contador, $fecha_toma_colesterol_hdl);
        $sheet->setCellValue('J'.$contador, $trigliceridos);
        $sheet->setCellValue('K'.$contador, $fecha_toma_trigliceridos);
        $sheet->setCellValue('L'.$contador, $albuminuriaCreatinuria);
        $sheet->setCellValue('M'.$contador, $fecha_toma_albimnuriaCreatinuria);
        $sheet->setCellValue('N'.$contador, $creatininaActual);
        $sheet->setCellValue('O'.$contador, $fecha_toma_creatinina);
        $sheet->setCellValue('P'.$contador, $albuminaSerica);
        $sheet->setCellValue('Q'.$contador, $fecha_toma_albumina_serica);
        $sheet->setCellValue('R'.$contador, $fosforo);
        $sheet->setCellValue('S'.$contador, $fecha_toma_fosforo);
        $sheet->setCellValue('T'.$contador, $pth);
        $sheet->setCellValue('U'.$contador, $fecha_toma_pth);
        $sheet->setCellValue('V'.$contador, $hemoglobina);
        $sheet->setCellValue('W'.$contador, $fecha_toma_hemoglobina);
        $sheet->setCellValue('X'.$contador, $hemoglobina_glico);
        $sheet->setCellValue('Y'.$contador, $fecha_toma_hemoglobinaGlico);
        $sheet->setCellValue('Z'.$contador, $uroanalisis);
        $sheet->setCellValue('AA'.$contador, $fecha_toma_uroanailis);
        $sheet->setCellValue('AB'.$contador, $glicemiaAyuno);
        $sheet->setCellValue('AC'.$contador, $fecha_toma_glicemia_ayuno);
        $sheet->setCellValue('AD'.$contador, $ultima_atencion_p);
        $sheet->setCellValue('AE'.$contador, $endo);
        $sheet->setCellValue('AF'.$contador, $cardio);
        $sheet->setCellValue('AG'.$contador, $ofta);
        $sheet->setCellValue('AH'.$contador, $nefro);
        $sheet->setCellValue('AI'.$contador, $psico);
        $sheet->setCellValue('AJ'.$contador, $nutricion);
        $sheet->setCellValue('AK'.$contador, $trabajo_social);
        $sheet->setCellValue('AL'.$contador, $medico_general);
        $sheet->setCellValue('AM'.$contador, $teleconsulta_medico_hgeneral);
        $sheet->setCellValue('AN'.$contador, $domociliario_medico_general);
        $sheet->setCellValue('AO'.$contador, $_telc_medicina_esp);
        $sheet->setCellValue('AP'.$contador, $telemedicina_esp);
        $sheet->setCellValue('AQ'.$contador, $domiciliario_medicina_espec);
        $sheet->setCellValue('AR'.$contador, $domociliario_promotor_salud);
        $sheet->setCellValue('AS'.$contador, $seguimiento_tel_enfermeria);
        $sheet->setCellValue('AT'.$contador, $visita_domiciliaria_aux_enfermeria);
        $sheet->setCellValue('AU'.$contador, $_segumiento_tel_enfermera);
        $sheet->setCellValue('AV'.$contador, $_visita_dom_enfermeria);
        $sheet->setCellValue('AW'.$contador, $_visita_dom_otro_profesional);
        $sheet->setCellValue('AX'.$contador, $visita_dom_equipo_inter);
        $sheet->setCellValue('AY'.$contador, $toma_laboratorio_dom);
        $sheet->setCellValue('AZ'.$contador, $_entrega_medicamentos_dom);
        
        echo'<br>';
        echo $contador;
        echo'<br>';
        $contador=$contador+1;

        $writer = new Xlsx($spreadsheet);
        //$writer->save('ejemplo.xlsx');vamos a ver 
        $nombreArchivoPDF = basename($pdfFilePath);
        $nombreArchivoExcel = str_replace('.pdf', '.xlsx', $nombreArchivoPDF); 
        $writer->save($nombreArchivoExcel);
        echo "El archivo $nombreArchivoExcel se ha generado correctamente.";
    }
} else {
    echo "No se encontró ningún texto entre las instancias de Fundación Clínica Nelson Mandela.";
}

?>
