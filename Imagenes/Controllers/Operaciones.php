<?php

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

namespace App\Controllers;

use App\Models\TablasModel;
use App\Models\AllogModel;
use App\Models\ArtiModel;
use App\Models\ArticuloModel;
use App\Models\FacartModel;
use App\Models\PreciosModel;
use App\Models\Server03Model;
use App\Models\Server03PruebaModel;

/**
 * Description of Operaciones
 *
 * @author José Luis
 */
class Operaciones extends BaseController {

    public function index() {
        $this->Tablas_model = new TablasModel();
        $operacion = $this->Tablas_model->get_by_tipreg('4', '1');
        $data['operacion'] = $operacion;
        $data['operaciones'];
        return view('operaciones/index', $data);
    }

    public function exportacion_guias() {

        return view('operaciones/index_exportacion');
    }

    public function ventas() {
        $session = session();        
        if($session->get('caja')){
            $local = $session->get('caja');
            switch($local){
                case '1': $data['color']='success'; break;
                case '2': $data['color']='danger'; break;
                case '3': $data['color']='info'; break;
            }
        }else{
            $local = 1;
            $data['color']='success';
        }
            $data['local']=$local;
            $smes=$session->get('mes_caja')?$session->get('mes_caja'):date('m');
            $session->set('mes_caja', $smes); 

        return view('operaciones/index_ventas',$data);
    }

    public function get_operaciones() {
        $fecha = $this->request->getVar('fecha');
        $opera = $this->request->getVar('operacion');
        $factu = $this->request->getVar('factura');
        $porciones = explode("-", $fecha);
        $fecha01 = date('d/m/Y', strtotime($porciones[0])); 
        $fecha02 = date('d/m/Y', strtotime($porciones[1]));

        $this->Allog_model = new AllogModel();
        $operaciones = $this->Allog_model->get_operaciones(0, 0, 0, $fecha01, $fecha02, $opera,$factu);
        //var_export($operaciones);
        return $this->response->setJSON($operaciones);
    }

    public function get_guias() {
        $session = session();        
        $local = $session->get('caja');
        $request = service('request');
        $fecha = $request->getPost('fecha');
        $factu = $this->request->getVar('factura');
        $porciones = explode("-", $fecha);
        $fecha01 = date('d/m/Y', strtotime($porciones[0]));
        $fecha02 = date('d/m/Y', strtotime($porciones[1]));
        $opcion = $request->getPost('operacion');
        $this->Allog_model = new AllogModel();
        $operaciones = $this->Allog_model->get_lista_ventas($fecha01, $fecha02, $opcion,$factu,$local);
        //var_export($operaciones);
        return $this->response->setJSON($operaciones);
    }

    public function migrate_guias() {
        $request = service('request');
        $fecha = date('d/m/Y', strtotime($request->getPost('fecha')));
        $serie = trim($request->getPost('serie'));
        $factu = trim($request->getPost('factura'));
        $serve = trim($request->getPost('server'));
        $TxtLo = $serve==3?'PEÑAMEZA':'JUANJUICILLO';
        $this->serv03 = new Server03Model();
        $allog = $this->serv03->get_max_numfac($serve);
        $FAR_NUMFAC = intval($allog->FAR_NUMFAC);
        $FAR_NUMOPER = intval($allog->FAR_NUMOPER);
        $this->facart = new FacartModel();
        $guia = $this->facart->get_guia($fecha, $serie, $factu);
        $fec_hoy = date('d/m/Y');
        $comiT = array();
        $c = 0;
        foreach ($guia as $detalle) {
            $comi['FAR_TIPMOV'] = $all['ALL_TIPMOV'] = $act['TIPMOV'] = 6;
            $comi['FAR_CODCIA'] = $all['ALL_CODCIA'] = $act['CODCIA'] = "25";
            $comi['FAR_NUMSER'] = $all['ALL_NUMSER'] = $act['NUMSER'] = "1";
            $comi['FAR_FBG'] = ' ';
            $comi['FAR_NUMFAC'] = $all['ALL_NUMFAC'] = $act['NUMFAC'] = $FAR_NUMFAC + 1;
            $comi['FAR_NUMSEC'] = $detalle->FAR_NUMSEC;
            $comi['FAR_FECHA'] = $all['ALL_FECHA_DIA'] = $act['FECHA'] = $fec_hoy;
            $comi['FAR_NUMOPER'] = $all['ALL_NUMOPER'] = $act['NUMOPER'] = $FAR_NUMOPER + 1;
            $comi['FAR_CODCLIE'] = 0;
            $comi['FAR_CODART'] = $detalle->FAR_CODART;
            $comi['FAR_ESTADO'] = "N";
            $comi['FAR_DIAS'] = 0;
            $comi['FAR_SIGNO_ARM'] = $all['ALL_SIGNO_ARM'] = 1;
            $comi['FAR_PRECIO'] = $detalle->FAR_PRECIO;
            $comi['FAR_STOCK'] = 0; //CALCULAR DESPUES DE INGRESAR
            $comi['FAR_COSPRO'] = $detalle->FAR_COSPRO;
            $comi['FAR_IMPTO'] = 0;
            $comi['FAR_TOT_DESCTO'] = 0;
            $comi['FAR_DESCTO'] = 0;
            $comi['FAR_GASTOS'] = 0;
            $comi['FAR_BRUTO'] = $all['ALL_IMPORTE_AMORT'] = $all['ALL_NETO'] = $all['ALL_BRUTO'] = $detalle->FAR_BRUTO;
            $comi['FAR_EQUIV'] = $detalle->FAR_EQUIV;
            $comi['FAR_PORDESCTO1'] = 0;
            $comi['FAR_TIPO_CAMBIO'] = 0;
            $comi['FAR_NUMSER_C'] = $all['ALL_NUMSER_C'] = 0;
            $comi['FAR_NUMFAC_C'] = $all['ALL_NUMFAC_C'] = 0;
            $comi['FAR_NUMDOC'] = 0;
            $comi['FAR_LIMCRE_ANT'] = 0;
            $comi['FAR_LIMCRE_ACT'] = 0;
            $comi['FAR_KEY_DIRCLI'] = 0;
            $comi['FAR_PRECIO_NETO'] = 0;
            $comi['FAR_CODVEN'] = 0;
            $comi['FAR_UNIDADES'] = 0;
            $comi['FAR_LITRO'] = 0;
            $comi['FAR_FECHA_COMPRA'] = $all['ALL_FECHA_VCTO'] = $fec_hoy;
            $comi['FAR_NUM_LOTE'] = $detalle->FAR_NUM_LOTE;
            $comi['FAR_CANTIDAD'] = $detalle->FAR_CANTIDAD;
            $comi['FAR_SIGNO_LOT'] = 0;
            $comi['FAR_CONCEPTO'] = $all['ALL_CONCEPTO'] = "INGRESO A MEDINAFARMA - $TxtLo G/R " . $serie . "-" . $factu;
            $comi['FAR_COD_SUNAT'] = 3;
            $comi['FAR_DESCRI'] = $detalle->FAR_DESCRI;
            $comi['FAR_PESO'] = $detalle->FAR_PESO;
            $comi['FAR_EX_IGV'] = "A";
            $comi['FAR_NUM_PRECIO'] = 0;
            $comi['FAR_SUBTRA'] = $all['ALL_SUBTRA'] = "MIGRACION WEB";
            $comi['FAR_CODUSU'] = $all['ALL_CODUSU'] = "WEB";
            $comi['FAR_COSPRO_ANT'] = $detalle->FAR_COSPRO_ANT;
            $comi['FAR_HORA'] = $detalle->FAR_HORA;
            $comi['FAR_CANTIDAD_P'] = $detalle->FAR_CANTIDAD_P;
            $comi['FAR_TURNO'] = $detalle->FAR_TURNO;
            $comi['FAR_ESTADO2'] = "N";
            $comi['FAR_FLAG_SO'] = "X";
            $comi['FAR_NUMOPER2'] = $all['ALL_NUMOPER2'] = $FAR_NUMOPER + 1;
            $comi['FAR_FECHA_PRO'] = $all['ALL_FECHA_PRO'] = $all['ALL_FECHA_SUNAT'] = $fec_hoy;
            $comi['FAR_FECHA_CAN'] = $all['ALL_FECHA_CAN'] = $fec_hoy;
            $comi['FAR_SUBTOTAL'] = $detalle->FAR_SUBTOTAL;
            $comi['FAR_CODLOT'] = "(*)";
            $comi['FAR_TRANSITO'] = " ";
            $all['ALL_CODTRA'] = $all['ALL_CODTRA_EXT'] = 2403;
            $all['ALL_CHESER'] = 'i_c';
            $all['ALL_CODSUNAT'] = 3;
            $all['ALL_CODUSU'] = 'ADMIN';
            $all['ALL_CANTIDAD'] = $c + $detalle->FAR_CANTIDAD;
            $all['ALL_AUTOCON'] = "MIGRACION WEB : INGRESO A MEDINAFARMA - $TxtLo G/R " . trim($serie) . "-" . $factu;
            array_push($comiT, $comi);
        }
        $new = $FAR_NUMFAC + 1;
        $allog2 = $this->serv03->crear_guia_allog($all,$serve);
        if($allog2){
            $facart = $this->serv03->crear_guia_facart($comiT,$serve);
            if($facart>0){
                $this->serv03->actualiza_stock($act,$serve);
                $this->Allog_model = new AllogModel();
                $this->Allog_model->set_guia_migrada($fecha, $serie, $factu,$new);
            }else{
                $msn = "No se actualizo el stock";
            }
            $msn="1_".$new;
        }else{
            $msn = "No se ingreso la guia";
        }        
       echo $msn;
    }

    public function get_operacion() {
        $fecha = $this->request->getVar('fecha');
        $serie = $this->request->getVar('serie');
        $factu = $this->request->getVar('factura');
        $opera = $this->request->getVar('operacion');
        $this->ArtiModel = new ArtiModel();
        $operacion = $this->ArtiModel->get_operacion($opera,0,$serie,$factu,$fecha);
        //var_export($operaciones);
        return $this->response->setJSON($operacion);       
    }

    public function get_precios_edit() {
        $key = $this->request->getVar('key');
       
        $this->PreciosModel = new PreciosModel();
        $operacion = $this->PreciosModel->get_precios_edit($key);
        //var_export($operaciones);
        //echo json_encode($operacion);
        return $this->response->setJSON($operacion);       
    }
    public function set_precios(){
        $codi = $this->request->getVar('codi');
        $ter2 = $this->request->getVar('dat2');
        $term = $this->request->getVar('dat1');
        parse_str($term, $out);
        $this->PreciosModel = new PreciosModel();
        $operacion = $this->PreciosModel->set_precios($codi,0,$out['PRE_POR1_1'],$out['PRE_POR2_1'],$out['PRE_POR3_1'],$out['PRE_POR4_1'],$out['PRE_POR5_1'],$out['PRE_PRE1_1'],$out['PRE_PRE2_1'],$out['PRE_PRE3_1'],$out['PRE_PRE4_1'],$out['PRE_PRE5_1']);

        parse_str($ter2, $out);
        if($out['PRE_PRE1_2']>0)
        $operacion = $this->PreciosModel->set_precios($codi,1,$out['PRE_POR1_2'],$out['PRE_POR2_2'],$out['PRE_POR3_2'],$out['PRE_POR4_2'],$out['PRE_POR5_2'],$out['PRE_PRE1_2'],$out['PRE_PRE2_2'],$out['PRE_PRE3_2'],$out['PRE_PRE4_2'],$out['PRE_PRE5_2']);
        return $operacion;
    }
    public function set_costo(){
        $key = $this->request->getVar('key');
        $costo = $this->request->getVar('costo');
        $this->ArticuloModel = new ArticuloModel();
        $operacion = $this->ArticuloModel->set_costo($key,$costo);
        return $operacion;
    }

}
