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
$sheet->setCellValue('N1', 'Fecha última atencion presencial en IPS por Medicina Interna');
$sheet->setCellValue('O1', 'Fecha última atencion presencial en IPS por Endocrinologia');
$sheet->setCellValue('P1', 'Fecha última atencion presencial en IPS por Cardiologia');
$sheet->setCellValue('Q1', 'Fecha última atencion presencial en IPS pór Oftalmologia');
$sheet->setCellValue('R1', 'Fecha última atencion presencial en IPS por Nefrologia');
$sheet->setCellValue('S1', 'Fecha de última Valoracion presencial en IPS por Psicologia');
$sheet->setCellValue('T1', 'Fecha de última valoracion presencial en IPS por Nutricion');
$sheet->setCellValue('U1', 'Fecha de última valoracion presencial en IPS por Trabajo Social');
$sheet->setCellValue('V1', 'Fecha de ultima consulta por medico general en la IPS ');
$sheet->setCellValue('W1', 'Fecha deTeleconsulta por medico general');
$sheet->setCellValue('X1', 'Fecha de Visita domiciliaria por medico general ');
$sheet->setCellValue('Y1', 'Fecha de Teleconsulta por medicina especializada ');
$sheet->setCellValue('Z1', 'Fecha de Telemedicina por especialidad ');  
$sheet->setCellValue('AA1', 'Fecha de Visita domiciliaria por medicina especializada ');
$sheet->setCellValue('AB1', 'Fecha de Visita domiciliaria por promotor de salud');
$sheet->setCellValue('AC1', 'Fecha Seguimiento telefónico por auxiliar de enfermeria');
$sheet->setCellValue('AD1', 'Fecha de Visita domiciliaria por auxiliar de enfermera ');
$sheet->setCellValue('AE1', 'Fecha Seguimiento telefónico por enfermeria');
$sheet->setCellValue('AF1', 'Fecha de Visita domiciliaria por enfermera');
$sheet->setCellValue('AG1', 'Fecha de Visita domiciliaria por otro profesional ');
$sheet->setCellValue('AH1', 'Fecha de Visita domiciliaria por equipo interdisciplinario');
$sheet->setCellValue('AI1', 'Fecha deToma de laboratorios en domicilio');
$sheet->setCellValue('AJ1', 'Fecha de Entrega de medicamentos -domicilio');
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
$contador=2;
if (preg_match_all($pattern, $text, $matches)) {
    foreach ($matches[1] as $match) {
        //$patron_fecha_consulta = '/Fecha\s+de\s+consulta.\s+\d{1,2}\/\d{1,2}\/\d{4}\s+\d+:\d+\s+Edad/m';
        //$patron_fecha_consulta = '/Fecha de consulta: 29\/05\/2023 10:33 Edad/m';
       // $patron_fecha_consulta = '/Fecha de consulta: (\d{2}\/\d{2}\/\d{4})/';
        $patron_fecha_consulta = '/.*Edad/m';
        $patron_paciente = '/Paciente:\s*(.*?)\s*Identificación:/';
        $patron_identificacion = '/Identificación:\s*([^:]+(\sTel?))/';
        
        $patron_colesterolT ='/Colesterol\s+Total\s\d+/';
        $patron_colesterolHDL = '/Colesterol\s+HDL\s\d+/';
        $patron_trigliceridos = '/Trigliceridos\s\d+/';
        $patron_colesterolLDL = '/Colesterol\s+LDL\s\d+/';
        $patron_albuminuriaCreatinuria_actualizado ='/Albuminuria\s\/\s+Creatinuria\s+-\s\(en\s+este\s+campo\s+resgitrar\s+el\s+dato\s+mas\s+actualizado\)\s+\d+/m';
        $patron_albuminaSerica ='/Albumina\s+Serica\s+\d+/';
        $patron_fosforo = '/Fosforo\s+\(p\)\s+\d+/';
        $patron_uroanalisis ='/UROANALIS\s+O\s+PARCIAL\s+DE\s+ORINA\s+\d+/';
        $patron_pth = '/PTH\s+.Paratohormona.\s+\d+/';
        $patron_hemoglobina ='/Hemoglobina\s+\d+/';
        $patron_glicemia_ayuno = '/Glicemia\s+de\s+Ayuno\s+\d+/m';
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
        if (preg_match($patron_colesterolHDL, $textoCapturado, $matches_colesterolHDL)) {
            $colesterolldl = $matches_colesterolHDL[0];
            //echo "colesterolldl: $colesterolldl\n";
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
            echo "patron_ultima_atencion_p: $ultima_atencion_p\n";
        }
        $ultima_medicinaIn ='ND';
        if (preg_match($patrorn_medicina_interna, $textoCapturado, $matches_medicina)) {
            //print_r( $matches_colesterol);
            $ultima_medicinaIn = $matches_medicina[1];
            echo " ultima_medicinaIn: $ultima_medicinaIn\n";
        }
        $endo ='ND';
        if (preg_match($patrorn_endo, $textoCapturado, $matches_endo)) {
            //print_r( $matches_colesterol);
            $endo = $matches_endo[1];
            echo " endo: $endo\n";
        }
        $cardio ='ND';
        if (preg_match($patrorn_cardio, $textoCapturado, $matches_cardio)) {
            //print_r( $matches_colesterol);
            $cardio = $matches_cardio[1];
            echo " cardio: $cardio\n";
        }
        $ofta ='ND';//no me toma
        if (preg_match($patrorn_ofta, $textoCapturado, $matches_ofta)) {
            //print_r( $matches_colesterol);
            $ofta = $matches_ofta[1];
            echo " ofta: $ofta\n";
        }
        $nefro ='ND';
        if (preg_match($patrorn_nefro, $textoCapturado, $matches_nefro)) {
            //print_r( $matches_colesterol);
            $nefro = $matches_nefro[1];
            echo " nefro: $nefro\n";
        }
        $psico ='ND';
        if (preg_match($patrorn_psico, $textoCapturado, $matches_psico)) {
            //print_r( $matches_colesterol);
            $psico = $matches_psico[1];
            echo " psico: $psico\n";
        }
        $nutricion='ND';
        if (preg_match($patrorn_nutricion, $textoCapturado, $matches_nutricion)) {
            //print_r( $matches_colesterol);
            $nutricion = $matches_nutricion[1];
            echo " nutricion: $nutricion\n";
        }
        $trabajo_social='ND';
        if (preg_match($patrorn_trabajo_social, $textoCapturado, $matches_trabajo_social)) {
            //print_r( $matches_colesterol);
            $trabajo_social = $matches_trabajo_social[1];
            echo " trabajo_social: $trabajo_social\n";
        }
        $medico_general='ND';
        if (preg_match($patrorn_medico_general, $textoCapturado, $matches_medico_general)) {
            //print_r( $matches_colesterol);
            $medico_general = $matches_medico_general[1];
            echo " medico_general: $medico_general\n";
        }
        $teleconsulta_medico_hgeneral='ND';
        if (preg_match($patrorn_teleconsulta_medico_hgeneral, $textoCapturado, $matches_teleconsulta_medico_hgeneral)) {
            //print_r( $matches_colesterol);
            $teleconsulta_medico_hgeneral = $matches_teleconsulta_medico_hgeneral[1];
            echo " teleconsulta_medico_hgeneral: $teleconsulta_medico_hgeneral\n";
        }
        $domociliario_medico_general='ND';
        if (preg_match($patrorn_domociliario_medico_general, $textoCapturado, $matches_domociliario_medico_general)) {
            //print_r( $matches_colesterol);
            $domociliario_medico_general = $matches_domociliario_medico_general[1];
            echo " domociliario_medico_general: $domociliario_medico_general\n";
        }
        $_telc_medicina_esp='ND';
        if (preg_match($patrorn_telc_medicina_esp, $textoCapturado, $matches__telc_medicina_esp)) {
            //print_r( $matches_colesterol);
            $_telc_medicina_esp = $matches__telc_medicina_esp[1];
            echo " _telc_medicina_esp: $_telc_medicina_esp\n";
        }
        $telemedicina_esp='ND';
        if (preg_match($patrorn_telemedicina_esp, $textoCapturado, $matches__telemedicina_esp)) {
            //print_r( $matches_colesterol);
            $telemedicina_esp = $matches__telemedicina_esp[1];
            echo " telemedicina_esp: $telemedicina_esp\n";
        }
        $domiciliario_medicina_espec='ND';
        if (preg_match($patrorn_domiciliario_medicina_espec, $textoCapturado, $matches__domiciliario_medicina_espec)) {
            //print_r( $matches_colesterol);
            $domiciliario_medicina_espec = $matches__domiciliario_medicina_espec[1];
            echo " domiciliario_medicina_espec: $domiciliario_medicina_espec\n";
        }
        $domociliario_promotor_salud='ND';
        if (preg_match($patrorn_domociliario_promotor_salud, $textoCapturado, $matches__domociliario_promotor_salud)) {
            //print_r( $matches_colesterol);
            $domociliario_promotor_salud = $matches__domociliario_promotor_salud[1];
            echo " domociliario_promotor_salud: $domociliario_promotor_salud\n";
        }
        $seguimiento_tel_enfermeria='ND';
        if (preg_match($patrorn_seguimiento_tel_enfermeria, $textoCapturado, $matches__seguimiento_tel_enfermeria)) {
            //print_r( $matches_colesterol);
            $seguimiento_tel_enfermeria = $matches__seguimiento_tel_enfermeria[1];
            echo " seguimiento_tel_enfermeria: $seguimiento_tel_enfermeria\n";
        }
        $visita_domiciliaria_aux_enfermeria='ND';
        if (preg_match($patrorn_visita_domiciliaria_aux_enfermeria, $textoCapturado, $matches__visita_domiciliaria_aux_enfermeria)) {
            //print_r( $matches_colesterol);
            $visita_domiciliaria_aux_enfermeria = $matches__visita_domiciliaria_aux_enfermeria[1];
            echo " visita_domiciliaria_aux_enfermeria: $visita_domiciliaria_aux_enfermeria\n";
        }
        $_segumiento_tel_enfermera='ND';
        if (preg_match($patrorn_segumiento_tel_enfermera, $textoCapturado, $matches___segumiento_tel_enfermera)) {
            //print_r( $matches_colesterol);
            $_segumiento_tel_enfermera = $matches___segumiento_tel_enfermera[1];
            echo " _segumiento_tel_enfermera: $_segumiento_tel_enfermera\n";
        }
        $_visita_dom_enfermeria='ND';
        if (preg_match($patrorn_visita_dom_enfermeria, $textoCapturado, $matches_visita_dom_enfermeria)) {
            //print_r( $matches_colesterol);
            $_visita_dom_enfermeria = $matches_visita_dom_enfermeria[1];
            echo " _visita_dom_enfermeria: $_visita_dom_enfermeria\n";
        }
        $_visita_dom_otro_profesional='ND';
        if (preg_match($patrorn_visita_dom_otro_profesional, $textoCapturado, $matches_visita_dom_otro_profesional)) {
            //print_r( $matches_colesterol);
            $_visita_dom_otro_profesional = $matches_visita_dom_otro_profesional[1];
            echo " _visita_dom_otro_profesional: $_visita_dom_otro_profesional\n";
        }
        $visita_dom_equipo_inter='ND';
        if (preg_match($patrorn_visita_dom_equipo_inter, $textoCapturado, $matches_visita_dom_equipo_inter)) {
            //print_r( $matches_colesterol);
            $visita_dom_equipo_inter = $matches_visita_dom_equipo_inter[1];
            echo " visita_dom_equipo_inter: $visita_dom_equipo_inter\n";
        }
        $toma_laboratorio_dom='ND';
        if (preg_match($patrorn_toma_laboratorio_dom, $textoCapturado, $matches_toma_laboratorio_dom)) {
            //print_r( $matches_colesterol);
            $toma_laboratorio_dom = $matches_toma_laboratorio_dom[1];
            echo " toma_laboratorio_dom: $toma_laboratorio_dom\n";
        }
        $_entrega_medicamentos_dom='ND';
        if (preg_match($patrorn_entrega_medicamentos_dom, $textoCapturado, $matches_entrega_medicamentos_dom)) {
            //print_r( $matches_colesterol);
            $_entrega_medicamentos_dom = $matches_entrega_medicamentos_dom[1];
            echo " _entrega_medicamentos_dom: $_entrega_medicamentos_dom\n";
        }
        echo'<br>';
        $sheet->setCellValue('A'.$contador, $fechaConsulta);
        $sheet->setCellValue('B'.$contador, $paciente);
        $sheet->setCellValue('C'.$contador, $identificacion);
        $sheet->setCellValue('D'.$contador, $colesterolT);
        $sheet->setCellValue('E'.$contador, $colesterolldl);
        $sheet->setCellValue('F'.$contador, $trigliceridos);
        $sheet->setCellValue('G'.$contador, $albuminuriaCreatinuria);
        $sheet->setCellValue('H'.$contador, $albuminaSerica);
        $sheet->setCellValue('I'.$contador, $fosforo);
        $sheet->setCellValue('J'.$contador, $pth);
        $sheet->setCellValue('K'.$contador, $hemoglobina);
        $sheet->setCellValue('L'.$contador, $uroanalisis);
        $sheet->setCellValue('M'.$contador, $glicemiaAyuno);
        $sheet->setCellValue('N'.$contador, $ultima_atencion_p);
        $sheet->setCellValue('O'.$contador, $endo);
        $sheet->setCellValue('P'.$contador, $cardio);
        $sheet->setCellValue('Q'.$contador, $ofta);
        $sheet->setCellValue('R'.$contador, $nefro);
        $sheet->setCellValue('S'.$contador, $psico);
        $sheet->setCellValue('T'.$contador, $nutricion);
        $sheet->setCellValue('U'.$contador, $trabajo_social);
        $sheet->setCellValue('V'.$contador, $medico_general);
        $sheet->setCellValue('W'.$contador, $teleconsulta_medico_hgeneral);
        $sheet->setCellValue('X'.$contador, $domociliario_medico_general);
        $sheet->setCellValue('Y'.$contador, $_telc_medicina_esp);
        $sheet->setCellValue('Z'.$contador, $telemedicina_esp);
        $sheet->setCellValue('AA'.$contador, $domiciliario_medicina_espec);
        $sheet->setCellValue('AB'.$contador, $domociliario_promotor_salud);
        $sheet->setCellValue('AC'.$contador, $seguimiento_tel_enfermeria);
        $sheet->setCellValue('AD'.$contador, $visita_domiciliaria_aux_enfermeria);
        $sheet->setCellValue('AE'.$contador, $_segumiento_tel_enfermera);
        $sheet->setCellValue('AF'.$contador, $_visita_dom_enfermeria);
        $sheet->setCellValue('AG'.$contador, $_visita_dom_otro_profesional);
        $sheet->setCellValue('AH'.$contador, $visita_dom_equipo_inter);
        $sheet->setCellValue('AI'.$contador, $toma_laboratorio_dom);
        $sheet->setCellValue('AJ'.$contador, $_entrega_medicamentos_dom);
        
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
