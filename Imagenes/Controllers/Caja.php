<?php

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

namespace App\Controllers;

use App\Models\AllogModel;
use App\Models\CajaModel;
use App\Models\CajaMovimientosModel;
use App\Models\VemaestModel;
use CodeIgniter\I18n\Time;

use Mike42\Escpos\EscposImage;
use Mike42\Escpos\PrintConnectors\WindowsPrintConnector;
use Mike42\Escpos\Printer;

/**
 * Description of Caja
 *
 * @author José Luis
 */
class Caja extends BaseController {

    private $motivos = [
        '1' => 'PAGO COMPRA',
        '2' => 'DEVOLUCIÓN PRODUCTO',
        '3' => 'CAMBIO DE PRODUCTO',
        '4' => 'SOBRANTE CAJA',
        '5' => 'FALTANTE CAJA',
        '6' => 'ADELANTO',
        '7' => 'CREDITO',
        '8' => 'PAGO CON TARJETA',
        '9' => 'PAGO CON TUNKY, YAPE, PLIN',
        '10' => 'DEPOSITO EN CUENTA',
        '11' => 'GASTOS BOTICA',
        '12' => 'DELIVERY',
        '13' => 'PAGO A MEDICO'
    ];

    public function index() {  
        $session = session();
        if($session->get('caja')){
            $caja = $session->get('caja');
            switch($caja){
                case '1': $data['color']='success'; break;
                case '2': $data['color']='danger'; break;
                case '3': $data['color']='info'; break;
            }
        }else{
            $caja = 1;
            $session->set('caja', $caja);
            $data['color']='success';
        }   
        $smes=$session->get('mes_caja')?$session->get('mes_caja'):date('m');
        $session->set('mes_caja', $smes);
        $this->Allog = new AllogModel();
        $this->Caja = new CajaModel();
        $mes = date('m');
        $anio = date('Y');
        $movimientos = $this->Caja->get_ventas_dia_det('', $smes, $anio,'', $caja);
        foreach($movimientos as $mov){
            $dia[date("d", strtotime($mov->ALL_FECHA_DIA))] = $dia[date("d", strtotime($mov->ALL_FECHA_DIA))] + $mov->TOT_VENTAS; //PARA GRAFICO
        }
        foreach ($dia as $key => $value) {
            $data['dias'][] = $key;
            $data['centro'][] = $value;
            $data['pmeza'][] = 0;
        }
        $data['ventas'] = $movimientos;
        return view('caja/index', $data);
    }
    public function reporte_cajas(){
        $session = session();
        if($session->get('caja')){
            $caja = $session->get('caja');
            switch($caja){
                case '1': $data['color']='success'; break;
                case '2': $data['color']='danger'; break;
                case '3': $data['color']='info'; break;
            }
        }else{
            $caja = 1;
            $data['color']='success';
            $session->set('caja', $caja);
        }   
        $smes=$session->get('mes_caja')?$session->get('mes_caja'):date('m');
        $session->set('mes_caja', $smes);
        $this->Allog = new AllogModel();
        $this->Caja = new CajaModel();
        $mes = date('m');
        $anio = date('Y');
        $this->Emp = new VemaestModel();
        $data['empleados'] = $emp = $this->Emp->get_empleado('');
        $movimientos = $this->Caja->get_ventas_dia_det('', $smes, $anio, '',$caja);
        //var_export($movimientos);
        $dia = array();
        foreach($movimientos as $mov){
            $d = date("d", strtotime($mov->ALL_FECHA_DIA));
            $dia[$d]['DIA'] = $d;
            foreach($emp as $e){
                $dia[$d][trim($e->VEM_NOMBRE)] = trim($e->VEM_NOMBRE)==trim($mov->VEM_NOMBRE)?$mov->TOT_EFECTIVO+$mov->TOT_MOVIM:($dia[$d][trim($e->VEM_NOMBRE)]?$dia[$d][trim($e->VEM_NOMBRE)]:0); 
            }
            $dia[$d]['TOT_DIA'] = $dia[$d]['TOT_DIA']+$mov->TOT_EFECTIVO+$mov->TOT_MOVIM;
        }
        $data['movimientos']=$dia;
        

        return view('caja/reporte_cajas',$data);
    }

    public function diario() {
        $session = session();        
        if($session->get('caja')){
            $caja = $session->get('caja');
            switch($caja){
                case '1': $data['color']='success'; break;
                case '2': $data['color']='danger'; break;
                case '3': $data['color']='info'; break;
            }
        }else{
            $caja = 1;
            $data['color']='success';
            $session->set('caja', $caja);
        }
            $smes=$session->get('mes_caja')?$session->get('mes_caja'):date('m');
            $session->set('mes_caja', $smes); 


        $this->Caja = new CajaModel();
        $mes = date('m');
        $anio = date('Y');
        $data['ventas'] = $this->Caja->get_ventas_dia('', $smes, $anio,'',$caja);
        $data['caja'] = $this->Caja->get_caja('',  '', $anio, 1, '','',$caja);


        $this->Emp = new VemaestModel();
        $data['empleados'] = $this->Emp->get_empleado('');
        $data['motivo_gasto'] = $this->motivos;
        $this->CajaMov = new CajaMovimientosModel();
        $data['movimientos'] = $this->CajaMov->get_movimientos($data['caja'][0]->CAJ_NRO?$data['caja'][0]->CAJ_NRO:0,$caja);
        return view('caja/index_diario', $data);
    }
public function dia(){
    $session = session();        
        if($session->get('caja')){
            $caja = $session->get('caja');
            switch($caja){
                case '1': $data['color']='success'; break;
                case '2': $data['color']='danger'; break;
                case '3': $data['color']='info'; break;
            }
        }else{
            $caja = 0;
            $data['color']='';
            $session->set('caja', $caja);
        }
            $smes=$session->get('mes_caja')?$session->get('mes_caja'):date('m');
            $session->set('mes_caja', $smes); 
            $session->set('anio_caja', date('Y')); 

        $this->Emp = new VemaestModel();
        $data['empleados'] = $this->Emp->get_empleado('');
        $data['motivo_gasto'] = $this->motivos;
    return view('caja/index_caja',$data);
}
public function get_cajas_dia(){
    $ani = $this->request->getVar('anio')?$this->request->getVar('anio'):date('Y');
    $mes = $this->request->getVar('mes')?$this->request->getVar('mes'):date('m');
    $dia = $this->request->getVar('dia')?$this->request->getVar('dia'):date('d');
    $est = $this->request->getVar('est');
    $loc = $this->request->getVar('loc');
    $caj = $this->request->getVar('caj');
    $this->Caja = new CajaModel();
    $cajas = $this->Caja->get_ventas_dia($dia,$mes,$ani,$caj,$loc);
    return $this->response->setJSON($cajas);
}
    public function diarios(){
        return view('caja/index_diarios', $data);
    }
    public function abrircaja() {
        $session = session();        
        $caja = $session->get('caja');
        $request = service('request');
        $data = array(
            'CAJ_CODVEN' => $request->getPost('VEM_CODVEN'),
            'CAJ_FECHA' => date('d/m/Y', strtotime($request->getPost('CAJ_FECHA2'))),
            'CAJ_ESTADO' => 1
        );
        $this->Caja = new CajaModel();
        $this->Caja->crear_caja($data,$caja);
        return redirect()->to('caja/diario');
    }
    public function abrircaja2() {
        $lcal = $this->request->getVar('local');
        $resp = $this->request->getVar('resp');
        $data = array(
            'CAJ_CODVEN' => $resp,
            'CAJ_FECHA' => date('d/m/Y'),
            'CAJ_ESTADO' => 1
        );
        $this->Caja = new CajaModel();
        $desc= $this->Caja->crear_caja($data,$lcal);
        return $this->response->setJSON($desc);
    }

    public function cerracaja() {
        $session = session();        
        $caja = $session->get('caja');
        $request2 = service('request');        
        $caja = $request2->getPost('LOCAL')?$request2->getPost('LOCAL'):$caja;
        $nrocaja = $request2->getPost('CAJ_NRO');
        $numser = $request2->getPost('CAJ_NUMSER');
        $numfact = $request2->getPost('CAJ_NUMFAC');
        $efectiv = $request2->getPost('CAJ_EFECTIVO');
        $fecha = $request2->getPost('CAJ_FECHA');
        $hoy = Time::parse($fecha, 'America/Lima');
        //actualizar en allog
        $this->Caja = new CajaModel();
        $cajaAnt = $this->Caja->get_caja('', '', '', 0, $nrocaja,'',$caja);
//var_export($cajaAnt); die();
        
        $ant = Time::parse($cajaAnt[0]->CAJ_FECHA, 'America/Lima');
        $diff = $hoy->difference($ant);
        //Obtener all_numoper 
        $nrop = $diff->getDays() == 0 ? ($cajaAnt[0]->CAJ_NUMOPER + 1) : 2;
        $this->Allog = new AllogModel();
        $numop2 = $this->Allog->get_nro_oper($hoy->toLocalizedString('d-M-Y'), $numfact,$caja);
        $this->Allog->set_cierre_caja($hoy->toLocalizedString('d-M-Y'), $numop2->ALL_NUMOPER, $nrop, $nrocaja,$caja);
        $this->Caja->cerrar_caja($numop2->ALL_NUMOPER, $numser, $numfact, $efectiv, $nrocaja,$caja);
        return redirect()->to('caja/diario');
    }
    public function cerrar_caja2(){
        $local = $this->request->getVar('local');
        $cajan = $this->request->getVar('CAJ_NRO'); /* ALL_CAJA_NRO */
        $numse = $this->request->getVar('CAJ_NUMSER');
        $numfa = $this->request->getVar('CAJ_NUMFAC'); /* ALL_NUMFAC */
        $efect = $this->request->getVar('CAJ_EFECTIVO');
        $fecha = $this->request->getVar('CAJ_FECHA');
        $hoy = Time::parse($fecha, 'America/Lima');
        /** nro de operacion de cierre */
        $this->Allog = new AllogModel();
        $numopAllog = $this->Allog->get_nro_oper($hoy->toLocalizedString('d-M-Y'), $numfa,$local);
        //var_export($numopAllog);
        /** datos de posible caja anterior */
        $this->Caja = new CajaModel();
        $cajaAnt = $this->Caja->get_caja_anterior($hoy->toLocalizedString('d-M-Y'),$numopAllog->ALL_NUMOPER,$local);
        $numopAnt=(count((array)$cajaAnt)>0)?$cajaAnt->CAJ_NUMOPER+1:2;
        try {
            $this->Allog->set_cierre_caja($hoy->toLocalizedString('d-M-Y'), $numopAllog->ALL_NUMOPER, $numopAnt, $cajan,$local);
            $this->Caja->cerrar_caja($numopAllog->ALL_NUMOPER, $numse, $numfa, $efect, $cajan,$local);
        } catch (\Exception $e) {
            exit($e->getMessage());
        }
        return true;
    }

    public function editarcaja(){
        $session = session();        
        $caja = $session->get('caja');
        if($session->get('user_id')=='ADMIN'){
            $nrocaja = $this->request->uri->getSegment(3);
            $this->Caja = new CajaModel();
            $this->Caja->editar_caja($nrocaja,$caja);
        }else{            
            $session->setFlashdata('item', 'No tiene permisos para está operación');
        } 
        
        return redirect()->to('caja/diario');
    }
    public function imprimircaja(){
        $session = session();      
        $local = $this->request->uri->getSegment(4);
        $nrocaja = $this->request->uri->getSegment(3);
        $this->Caja = new CajaModel();
        $anio = date('Y');
        $fecha = date('d/m/Y');
        $locales = array(1=>"CENTRO",2=>"JUANJUICILLO",3=>"PEÑAMEZA");
        $revisado = '';
        $cajac = $this->Caja->get_caja('',  '', $anio, 0, '',$nrocaja,$local);
        if($session->get('user_id')=='ADMIN'||date('d/m/Y', strtotime($cajac[0]->CAJ_FECHA))==$fecha){
            $cajar = $this->Caja->get_ventas_dia_det(date('d', strtotime($cajac[0]->CAJ_FECHA)),date('m', strtotime($cajac[0]->CAJ_FECHA)),date('Y', strtotime($cajac[0]->CAJ_FECHA)),$nrocaja,$local);
            $this->CajaMov = new CajaMovimientosModel();
            $movimientos = $this->CajaMov->get_movimientos($cajar[0]->ALL_CAJA_NRO?$cajar[0]->ALL_CAJA_NRO:0,$local);
            if($local==1 || $session->get('user_id')=='ADMIN'){
                $connector = new WindowsPrintConnector("smb://asesor:159357@ventas2/6-EPSON TM-T20II Receipt5");  
                $revisado = $session->get('user_id')=='ADMIN'?'**REVISADO**':'';              
            }elseif($local==2){                
                $connector = new WindowsPrintConnector("smb://asesor:159357@server02/6-EPSON TM-T20II Receipt");
            }elseif($local==3){
                $connector = new WindowsPrintConnector("smb://asesor:159357@medinaimpresora/6-EPSON TM-T20II Receipt");
            }
            
            $printer = new Printer($connector);
            $printer->setFont();  
            $printer -> setJustification(Printer::JUSTIFY_CENTER);
            $logo = EscposImage::load(FCPATH.'dist\img\medinafarma-black.jpg', false);
            $printer -> graphics($logo);        
            $printer -> feed();
            $printer->setJustification(Printer::JUSTIFY_CENTER);
            $printer -> text("█████▓▒░░ REPORTE DE CIERRE DE CAJA ░░▒▓█████ \n");
            $printer -> setTextSize(4, 4);
            $printer -> text($locales[$local]."\n");
            $printer -> setTextSize(1,1);
            $printer -> setJustification(Printer::JUSTIFY_LEFT);
            $printer -> text("╔═══════╦════════════╗  ╔══════╦═══════╗\n");
            $printer -> text("║ FECHA ║ ".date('d-m-Y', strtotime($cajar[0]->ALL_FECHA_DIA))." ║  ║ CAJA ║ ".str_pad($nrocaja, 5, " ", STR_PAD_LEFT)." ║\n");
            $printer -> text("╚═══════╩════════════╝  ╚══════╩═══════╝\n");        
            $printer -> text("╔═════════╦══════════╗  ╔══════╦═══════════════╗\n");
            $printer -> text("║ CLIENTE ║ ".str_pad($cajac[0]->CAJ_NUMOPER, 8, " ", STR_PAD_LEFT)." ║  ║ DOCU ║ ".str_pad($cajac[0]->CAJ_NUMFAC, 13, " ", STR_PAD_LEFT)." ║\n");  
            $printer -> text("╚═════════╩══════════╝  ╚══════╩═══════════════╝\n");
            $printer -> text("╔═══════╦══════════════╗ \n");
            $printer -> text("║ MONTO ║ S/. " . str_pad(number_format((float)round( $cajac[0]->CAJ_EFECTIVO ,2, PHP_ROUND_HALF_DOWN),2,'.',','), 8, " ", STR_PAD_LEFT)." ║ $revisado\n");
            $printer -> text("╚═══════╩══════════════╝ \n");
            $printer -> text("┌─────────────┐ \n");
            $printer -> text("│ MOVIMIENTOS │ \n");
            $printer -> text("├─────────────┴───────────────────┬────────────┐\n");
            $printer -> text("│ CONCEPTO                        │    MONTO   │\n");
            $printer -> text("├─────────────────────────────────┼────────────┤\n");

            foreach ($movimientos as $val) { 
                // 31 array ( 0 => (object) array( 'CMV_NRO' => 322, 'CMV_CAJA' => 313, 'CMV_TIPO' => 1, 'CMV_CODVEN' => 18, 'CMV_DESCRIPCION' => 'DEPOSITO', 'CMV_MONTO' => '100.00', ), ) -->
                $concep = $this->motivos[$val->CMV_TIPO]." : ".$val->CMV_DESCRIPCION;
                $find = array('á','é','í','ó','ú','â','ê','î','ô','û','ã','õ','ç','ñ','Á','É','Í','Ó','Ú','Â','Ê','Î','Ô','Û','Ã','Õ','Ç','Ñ');
                $repl = array('a','e','i','o','u','a','e','i','o','u','a','o','c','n','A','E','I','O','U','A','E','I','O','U','A','O','C','N');
                $concep = str_replace($find, $repl, $concep);
                $concepto = str_pad(substr($concep, 0, 31),31);
                $monto = "S/." . str_pad($val->CMV_MONTO, 7, " ", STR_PAD_LEFT); 
                $printer -> text("│ ".$concepto." │ ".$monto." │\n");
                }   

            $printer -> text("└─────────────────────────────────┴────────────┘\n");
            $printer -> text("                                  ┌────────────┐\n");
            $printer -> text("              MONTO TOTAL DE CAJA │ S/." . str_pad(number_format((float)round(($cajar[0]->TOT_MOVIM+$cajar[0]->TOT_EFECTIVO),2, PHP_ROUND_HALF_DOWN),2,'.',''), 7, " ", STR_PAD_LEFT)." │\n");
            $printer -> text("                                  ├────────────┤\n");
            $printer -> text("                MONTO DEL SISTEMA │ S/." . str_pad(number_format((float)round( $cajar[0]->TOT_VENTAS ,2, PHP_ROUND_HALF_DOWN),2,'.',''), 7, " ", STR_PAD_LEFT)." │\n");
            $printer -> text("                                  ├────────────┤\n");
            $printer -> text("                       DIFERENCIA │ S/." . str_pad(number_format((float)round(($cajar[0]->TOT_MOVIM+$cajar[0]->TOT_EFECTIVO)-$cajar[0]->TOT_VENTAS,2, PHP_ROUND_HALF_DOWN),2,'.',''), 7, " ", STR_PAD_LEFT)." │\n");
            $printer -> text("                                  └────────────┘\n");
            $printer -> text("┌────────────────────────────┬─────────────────┐\n");
            $printer -> text("│ CAJERO:                    │ FIRMA:          │\n");
            $printer -> text("│ ".str_pad(trim($cajar[0]->VEM_NOMBRE), 26)." │                 │\n");
            $printer -> text("│                            │                 │\n");
            $printer -> text("└────────────────────────────┴─────────────────┘\n");
            $printer -> text("OBSERVACIONES:\n");
            /* Footer */
            $printer -> feed(2);
            $printer -> setJustification(Printer::JUSTIFY_CENTER);
            /* Barcodes - see barcode.php for more detail */
            $printer->setBarcodeHeight(80);
            $printer->setBarcodeTextPosition(Printer::BARCODE_TEXT_BELOW);
            $printer->barcode($nrocaja);

            
            $printer -> text("Fecha de Impresión:".date('d-m-Y h:i:s a', time())."\n");

            /* Cut the receipt and open the cash drawer */
            $printer -> cut();
            $printer -> pulse();
            $printer -> close();
        }else{            
            $session->setFlashdata('item', 'No tiene permisos para está operación');
        } 

return redirect()->to('caja/diario');
        
        
    }
    public function agregar_movimiento() {
        $session = session();        
        $caja = $session->get('caja');
        $request = service('request');
        $data = array(
            'CMV_CAJA' => $request->getPost('cmv_caja'),
            'CMV_TIPO' => $request->getPost('cmv_tipo'),
            'CMV_CODVEN' => $request->getPost('cmv_codven'),
            'CMV_DESCRIPCION' => strtoupper($request->getPost('cmv_descri')),
            'CMV_MONTO' => $request->getPost('cvm_monto')
        );
        $this->CajaMov = new CajaMovimientosModel();
        $this->CajaMov->crear_movimiento($data,$caja);
        $movimientos = $this->CajaMov->get_movimientos($request->getPost('cmv_caja'),$caja);
        foreach ($movimientos as $val) {
            $moti = $this->motivos[$val->CMV_TIPO];
            $tr = "<tr><td>$moti</td><td>$val->CMV_DESCRIPCION</td><td>$val->CMV_MONTO</td>";            
            $tr.= "<td><a href='#' class='nav-link' title='Eliminar'><span class='float-right badge bg-danger'><i class='fas fa-trash' onclick='quitar_mov($val->CMV_NRO)'></i></span></a></td>";
            echo $tr;
        }
    }
    public function listar_movimientos(){
        $session = session();        
        $caja = $session->get('caja');
        $nroc = $this->request->getVar('nro_caja');
        $this->CajaMov = new CajaMovimientosModel();        
        $movimientos = $this->CajaMov->get_movimientos($nroc,$caja);
        return$this->response->setJSON($movimientos);
    }
    public function get_nro_doc(){
        $session = session();        
        $caja = $session->get('caja');
        $this->Allog = new AllogModel();
        $allog = $this->Allog->get_nro_doc(date('d/m/Y'),$caja);
        echo json_encode($allog);
    }
    public function eliminar_movimiento(){
        $session = session();        
        $caja = $this->request->getVar('local')?$this->request->getVar('local'):$session->get('caja');
        $data = $this->request->getVar('cmv_nro');
        $this->CajaMov = new CajaMovimientosModel();
        $this->CajaMov->delete_movimiento($data,$caja);
    }
    public function editar_caja2(){
        $local = $this->request->getVar('local');
        $caja = $this->request->getVar('caja');
        $resp = $this->request->getVar('resp');
        $this->Caja = new CajaModel();
        $this->Caja->editar_caja2($resp,$caja,$local);
    }
	public function set_caja(){
		$data = $this->request->getVar('caja');
		$session = session();
        $session->set('caja', $data);        
		return 'ok'; 
	}
	public function set_mes(){
		$data = $this->request->getVar('mes');
		$session = session();
        $session->set('mes_caja', $data);        
		return 'ok'; 
	}
}