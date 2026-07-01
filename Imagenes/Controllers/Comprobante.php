<?php

namespace App\Controllers;
use App\Models\FacartModel;
use App\Controllers\BaseController;
use Luecano\NumeroALetras\NumeroALetras;
use Mike42\Escpos\EscposImage;
use Mike42\Escpos\PrintConnectors\WindowsPrintConnector;
use Mike42\Escpos\Printer;


class Comprobante extends BaseController
{
    public function doc()
    {
        $session = session();      
        $local = $this->request->uri->getSegment(3);
        $tipmov = $this->request->uri->getSegment(4);
        $numser = $this->request->uri->getSegment(5);
        $numfac = $this->request->uri->getSegment(6);
        $fecha = $this->request->uri->getSegment(7);
        $this->Facart = new FacartModel();
        $Facart = $this->Facart->get_comprobante($numser,$numfac,$tipmov,date('d/m/Y', strtotime($fecha)),$local);
        $locales = array(1=>"CENTRO",2=>"JUANJUICILLO",3=>"PEÑAMEZA");
        if(count($Facart)>0){

            if($local==1 || $session->get('user_id')=='ADMIN'){
                $connector = new WindowsPrintConnector("smb://asesor:159357@ventas2/6-EPSON TM-T20II Receipt5");
            }elseif($local==2){                
                $connector = new WindowsPrintConnector("smb://asesor:159357@server02/6-EPSON TM-T20II Receipt");                
            }elseif($local==3){
                $connector = new WindowsPrintConnector("smb://asesor:159357@medinaimpresora/6-EPSON TM-T20II Receipt");                
            }
            if($Facart[0]->FAR_FBG=='B'){
                $documento= "BOLETA : B0";
                $tipoDoc=empty(trim($Facart[0]->CLI_RUC_ESPOSA))?'':'DNI:'.$Facart[0]->CLI_RUC_ESPOSA."\n";
                $direccion="";
            }elseif ($Facart[0]->FAR_FBG=='F') {
                $documento= "FACTURA : F0";
                $tipoDoc='RUC : '.TRIM($Facart[0]->CLI_RUC_ESPOSO)."\n";
                $direccion ='DIRECCION : '.TRIM($Facart[0]->CLI_CASA_DIREC).' '.TRIM($Facart[0]->CLI_CASA_NUM)."\n";
            }elseif ($Facart[0]->FAR_FBG=='G') {
                $documento= "GUIA : G0";
                $tipoDoc='';
                $direccion="";
            }

            $logo = EscposImage::load(FCPATH.'dist\img\medinafarma-black.jpg', false);
            $printer = new Printer($connector);
            $copia = 0 ; // 'COP_SEC' => 1, 'COP_DESCRIP' => 'CLIENTE'
            foreach ($Facart as $val) { 
                
            
            if($copia != $val->COP_SEC){$copia = $val->COP_SEC;}
            
            $printer ->setFont(Printer::FONT_B);   
            $printer -> setJustification(Printer::JUSTIFY_CENTER);            
            $printer -> graphics($logo);        
            $printer -> feed();
            $printer -> text("INVERSIONES SAN MARTIN S.C.R.L.\n");
            $printer -> text("RUC: 20450337839\n");
            $printer -> text("Jr. Huallaga Nro 601 - Juanjuí - Mcal Cacéres - San Martín\n");
            $printer->setJustification(Printer::JUSTIFY_LEFT);
            $printer->text("----------------------------------------------------------------" . "\n");
            
            $printer->text($documento.": ");
            $printer->text(TRIM($val->FAR_NUMSER)."-".$val->FAR_NUMFAC."\n");
            $printer->text("Fecha Emisión: ");
            $printer->text(date('d-m-Y', strtotime($val->FAR_FECHA))." ".date("h:i A", strtotime($val->FAR_HORA)) . "\n");
            $printer->text("Responsable: ");
            $printer->text($val->VEM_NOMBRE."\n");

            $printer->text("----------------------------------------------------------------" . "\n");
            $printer->text("Cliente: ");
            $printer->text(empty(trim($val->CLI_NOMBRE))?trim($val->FAR_CLIENTE."\n"):trim($val->CLI_NOMBRE)."\n");
            $printer->text($tipoDoc);
            $printer->text($direccion);
            $printer->text("----------------------------------------------------------------" . "\n");
            $printer->text("PRODUCTO                         CANT.     P/U.         IMPORTE\n");
            $printer->text("----------------------------------------------------------------" . "\n");
            foreach ($Facart as $val2) { 
                if($copia == $val2->COP_SEC){
                $concepto = str_pad(substr(trim($val2->ART_NOMBRE), 0, 64),64, " ");                
                $monto = str_pad($val2->FAR_CANTIDAD_P/$val2->FAR_EQUIV, 40, " ", STR_PAD_LEFT); 
                $monto.= " ";
                $monto.= str_pad(substr($val2->FAR_DESCRI, 0, 5), 5, " ", STR_PAD_LEFT); 
                $monto.= str_pad(number_format((float)round($val2->FAR_PRECIO ,2, PHP_ROUND_HALF_DOWN),2,'.',','), 8, " ", STR_PAD_LEFT); 
                $monto.= str_pad(number_format((float)round($val2->FAR_PRECIO*$val2->FAR_CANTIDAD_P/$val2->FAR_EQUIV ,2, PHP_ROUND_HALF_DOWN),2,'.',','), 8, " ", STR_PAD_LEFT); 
                $printer-> text($concepto."\n");
                $printer-> text($monto."\n");
                }

            }


            $printer->text("----------------------------------------------------------------" . "\n");              
            $printer->text(str_pad(substr("Subtotal :", 0, 15), 46, " ", STR_PAD_LEFT) );   
            $printer->text(" S/. ");
            $printer->text(str_pad(number_format((float)round($val->FAR_BRUTO,2, PHP_ROUND_HALF_DOWN),2,'.',','), 12, " ", STR_PAD_LEFT). "\n");

            $printer->text(str_pad(substr("DESCUENTO :", 0, 15), 46, " ", STR_PAD_LEFT) );               
            $printer->text(" S/. ");
            $printer->text(str_pad(number_format((float)round(0.00,2, PHP_ROUND_HALF_DOWN),2,'.',','), 12, " ", STR_PAD_LEFT). "\n");

            $printer->text(str_pad(substr("ISC :", 0, 15), 46, " ", STR_PAD_LEFT) );               
            $printer->text(" S/. ");
            $printer->text(str_pad(number_format((float)round(0.00,2, PHP_ROUND_HALF_DOWN),2,'.',','), 12, " ", STR_PAD_LEFT). "\n");

            $printer->text(str_pad(substr("IGV (18%) :", 0, 15), 46, " ", STR_PAD_LEFT) );               
            $printer->text(" S/. ");
            $printer->text(str_pad(number_format((float)round(0.00,2, PHP_ROUND_HALF_DOWN),2,'.',','), 12, " ", STR_PAD_LEFT). "\n");


            $printer->text(str_pad(substr("IMPORTE TOTAL :", 0, 15), 46, " ", STR_PAD_LEFT) );               
            $printer->text(" S/. ");
            $printer->setTextSize(2, 2);
            $printer->text(str_pad(number_format((float)round($val->FAR_BRUTO,2, PHP_ROUND_HALF_DOWN),2,'.',','), 6, " ", STR_PAD_LEFT). "\n");

            $printer->setTextSize(1, 1);  
            
            $formatter = new NumeroALetras();

            $printer->setJustification(Printer::JUSTIFY_LEFT);
            $printer->text("SON: ".$formatter->toMoney($val->FAR_BRUTO, 2, 'SOLES', 'CENTIMOS')."\n");
            
            $printer->text("Forma de pago: CONTADO \n");
            if ($val->FAR_FBG=='F') {
            $printer -> qrCode("20450337839|01|F0".$val->FAR_NUMSER."|".$val->FAR_NUMFAC."|0.00|".$val->FAR_BRUTO."|".date('Y-m-d', strtotime($val->FAR_FECHA))."|6|".$val->CLI_RUC_ESPOS0, Printer::QR_ECLEVEL_L, 5);
            }
            $printer->text("\n");

            $printer->setJustification(Printer::JUSTIFY_CENTER);
            $printer->text("-- ".$val->COP_DESCRIP." --\n");
            $printer->text("Representación impresa de la "."\n");
            $printer->text("Obligado a ser Emisor Electrónico mediante la Resolución de Superintendecia"."\n");
            $printer->text("N° 155-2017/SUNAT-Anexo IV  \n");
            $printer->text(" " . "\n");
            $printer->setJustification(Printer::JUSTIFY_CENTER);
            $printer->text("GRACIAS POR SU VISITA !  \n");
           
            $printer->feed(3);
            $printer->cut();
        }
            $printer->pulse();
            $printer -> close();
        }
        
    }
}
