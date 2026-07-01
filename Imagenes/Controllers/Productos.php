<?php

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

namespace App\Controllers;
use App\Models\ArticuloModel;
use App\Models\Server03Model;
use App\Models\ArtiModel;
use App\Models\PreciosModel;
use App\Models\PreciosDigemidModel;
use App\Models\LoteModel;
/**
 * Description of Productos
 *
 * @author José Luis
 */
class Productos extends BaseController {

    public function index() {
        return view('productos/index');
    }
    public function get_productos(){
        $busqueda = $this->request->getVar('busqueda');
        $this->ArticuloModel = new ArticuloModel();
        $articulos = $this->ArticuloModel->get_articulos_det($busqueda);
        return $this->response->setJSON($articulos);
    }
    public function get_stock(){
        $artkey = $this->request->getVar('artkey');
        $artsbg = $this->request->getVar('artsbg');
        $local = $this->request->getVar('local');
        $btn = $this->request->getVar('btn');
        if(!empty($artkey) || $artsbg > 0){
            $this->Server03Model = new Server03Model();
            $articulos = $this->Server03Model->get_articulos_selec($artkey,$artsbg,$local);
            $locales = array('','CENTRO','JCILLO','PMEZA');
            foreach ($articulos as $val) {           
                $class = $val->StockGen>0?'even bg-success':'';
                $RNT='';
                $RENTABLE = array(113,444,521,456);
                if (in_array($val->art_familia, $RENTABLE)) {
                    $RNT='<i class="fas fa-chart-line"></i>';
                    $class = 'even bg-gray-dark';
                }
                $pqt = round($val->ART_PQT, 0);
                $und = round($val->ART_UNID, 0);
                $pre = round($val->PRE_PRE1, 2);
                $btn = $btn?"<td><button id='selprod' class='btn btn-block bg-gradient-primary btn-xs'><i class='fas fa-pills'></i> Seleccionar</button></td>":'';
                $tr = "<tr data-id='$val->ARM_CODART' data-toggle='modal' data-target='#modal-overlay' class='$class'><td>$locales[$local]</td></td></t><td class='descrip'>$RNT $val->ART_NOMBRE</td><td>$pqt</td><td>$und</td><td>$pre</td>$btn</tr>";
                echo $tr;
            }
        }
    }

    public function get_inventario(){
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
        }
        return view('productos/index_inventario',$data);
    }

    public function get_stock_articulos(){
        $session = session();        
        $caja = $session->get('caja'); 
        $stock = $this->request->getVar('stock');
        $this->ArticuloModel = new ArticuloModel();
        $articulos = $this->ArticuloModel->get_stock_articulos($stock,$caja);
        return $this->response->setJSON($articulos);
    }

    public function actualizar_productos(){
        return view('productos/index_actualizar');
    }
    public function crea_productos(){
        $serve =  $this->request->getVar('serve');
        $this->Server03Model = new Server03Model();
        $get_max_numart = $this->Server03Model->get_max_numart($serve);
        $this->ArtiModel = new ArtiModel();
        $get_arti_act=$this->ArtiModel->get_arti_act($get_max_numart);
        echo $get_max_numart.'|'. $get_arti_act; 
    }
    public function exporta_nuevos_prod(){
        $serve = $this->request->getVar('serve');
        $keyar = $this->request->getVar('keyval');
        $total = $this->crea_arti($keyar,$serve);
        if($total>0){$total = $this->crea_articulo($keyar,$serve);};
        if($total>0){$total = $this->crea_precios($keyar,$serve);};
        if($total>0){$total = $this->crea_lote($keyar,$serve);};
        return $total;
    }

    public function actualiza_precios(){
        $serve = $this->request->getVar('serve');
        $this->ArticuloModel = new ArticuloModel();
        $data = $this->ArticuloModel->get_precios(1);
        $this->Server03Model = new Server03Model();
        $c=0;
        foreach ($data as $item) {
            $oper = $this->Server03Model->update_precios_precio($serve,$item->PRE_CODCIA,$item->PRE_CODART,$item->PRE_SECUENCIA,$item->PRE_PRE1,$item->PRE_PRE2,$item->PRE_POR1,$item->PRE_POR2);
            $this->Server03Model->update_precios_articulo($server,$item->PRE_CODCIA,$item->PRE_CODART,$item->ARM_COSPRO);
            if ($oper){
                $comi['CANTIDAD'] = $c++;
            }else{
                $comi['PRE_CODART'][] = $item->PRE_CODART;
            }            
        }
        return json_encode($comi);
    }

    public function crea_arti($keyar,$serve){
        $this->ArtiModel = new ArtiModel();
        $data = $this->ArtiModel->get_arti($keyar);
        $comiT = array();
        foreach ($data as $item) {
            $comi['ART_KEY'] = $item->ART_KEY;
            $comi['ART_CODCIA'] = $item->ART_CODCIA;
            $comi['ART_NOMBRE'] = $item->ART_NOMBRE;
            $comi['ART_COSTO'] = $item->ART_COSTO;
            $comi['ART_TIPO'] = $item->ART_TIPO;
            $comi['ART_ESTADO'] = $item->ART_ESTADO;
            $comi['ART_CALIDAD'] = $item->ART_CALIDAD;
            $comi['ART_UNIDAD'] = $item->ART_UNIDAD;
            $comi['ART_EX_IGV'] = $item->ART_EX_IGV;
            $comi['ART_DECIMALES'] = $item->ART_DECIMALES;
            $comi['ART_FAMILIA'] = $item->ART_FAMILIA;
            $comi['ART_ALTERNO'] = $item->ART_ALTERNO;
            $comi['ART_CP'] = $item->ART_CP;
            $comi['ART_FLAG_STOCK'] = $item->ART_FLAG_STOCK;
            $comi['ART_FLAG_CAMBIO'] = $item->ART_FLAG_CAMBIO;
            $comi['ART_FECHAHORA'] = $item->ART_FECHAHORA;
            $comi['ART_GRUPOP'] = $item->ART_GRUPOP;
            $comi['ART_CODREL'] = $item->ART_CODREL;
            array_push($comiT, $comi);
        }
        $this->Server03Model = new Server03Model();
        $total = $this->Server03Model->crear_prod_arti($comiT,$serve);
        return $total;
    }
    public function crea_articulo($keyar,$serve){
        $this->ArticuloModel = new ArticuloModel();
        $data = $this->ArticuloModel->get_articulos($keyar);
        $comiT = array();
        foreach ($data as $item) {
            $comi['ARM_CODART'] = $item->ARM_CODART;
            $comi['ARM_CODCIA'] = $item->ARM_CODCIA;
            $comi['ARM_STOCK'] = $item->ARM_STOCK;
            $comi['ARM_INGRESOS'] = $item->ARM_INGRESOS;
            $comi['ARM_SALIDAS'] = $item->ARM_SALIDAS;
            $comi['ARM_STOCK_INI'] = $item->ARM_STOCK_INI;
            $comi['ARM_COSPRO'] = $item->ARM_COSPRO;
            $comi['ARM_STOCK2'] = $item->ARM_STOCK2;
            $comi['ARM_COSTO_ULT'] = $item->ARM_COSTO_ULT;
            $comi['ARM_FECHA_ULT'] = date('d/m/Y');
            $comi['ARM_SALDO_S'] = $item->ARM_SALDO_S;
            $comi['ARM_SALDO_S2'] = $item->ARM_SALDO_S2;
            $comi['ARM_SALDO_N'] = $item->ARM_SALDO_N;
            $comi['ARM_SALDO_N2'] = $item->ARM_SALDO_N2;
            $comi['ARM_STOCK_T'] = $item->ARM_STOCK_T;
            $comi['ARM_AJUSTA'] = $item->ARM_AJUSTA;
            array_push($comiT, $comi);
        }
        $this->Server03Model = new Server03Model();
        $total = $this->Server03Model->crear_prod_articulo($comiT,$serve);
        return $total;
    }
    public function crea_precios($keyar,$serve){
        $this->PreciosModel = new PreciosModel();
        $data = $this->PreciosModel->get_precios($keyar);
        $comiT = array();
        foreach ($data as $item) {
            $comi['PRE_CODCIA'] = $item->PRE_CODCIA;
            $comi['PRE_CODART'] = $item->PRE_CODART;
            $comi['PRE_SECUENCIA'] = $item->PRE_SECUENCIA;
            $comi['PRE_POR1'] = $item->PRE_POR1;
            $comi['PRE_POR2'] = $item->PRE_POR2;
            $comi['PRE_PRE1'] = $item->PRE_PRE1;
            $comi['PRE_PRE2'] = $item->PRE_PRE2;
            $comi['PRE_UNIDAD'] = $item->PRE_UNIDAD;
            $comi['PRE_EQUIV'] = $item->PRE_EQUIV;
            $comi['PRE_FLAG_UNIDAD'] = $item->PRE_FLAG_UNIDAD;
            array_push($comiT, $comi);
        }
        $this->Server03Model = new Server03Model();
        $total = $this->Server03Model->crear_prod_precios($comiT,$serve);
        return $total;
    }
    public function crea_lote($keyar,$serve){
        $this->LoteModel = new LoteModel();
        $data = $this->LoteModel->get_lotes($keyar);
        $comiT = array();
        foreach ($data as $item) {
            $comi['LOT_CODCIA'] = $item->LOT_CODCIA;
            $comi['LOT_CODART'] = $item->LOT_CODART;
            $comi['LOT_NROLOTE'] = $item->LOT_NROLOTE;
            $comi['LOT_CODCLIE'] = $item->LOT_CODCLIE;
            $comi['LOT_FECHA_VCTO'] = date('d/m/Y');
            $comi['LOT_SALDOS'] = $item->LOT_SALDOS;
            $comi['LOT_CODUSU'] = $item->LOT_CODUSU;
            array_push($comiT, $comi);
        }
        $this->Server03Model = new Server03Model();
        $total = $this->Server03Model->crear_prod_lote($comiT,$serve);
        return $total;
    }
    public function get_precios_digemid(){        
        $key = $this->request->getVar('artkey');
        $this->PreciosDigemidModel = new PreciosDigemidModel();
        $data = $this->PreciosDigemidModel->get_precios_digemid($key);
        return json_encode($data);
    }
}