<?php

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

namespace App\Controllers;
use App\Models\ClientesModel;

/**
 * Description of Personas
 *
 * @author José Luis
 */
class Personas extends BaseController {

    public function index() {     

        return view('personas/index');
    }
    public function get_personas(){
        $busqueda = $this->request->getVar('busqueda');
        $tipoCli = $this->request->getVar('tipoCli');
        
        $this->ClientesModel = new ClientesModel();
        $clientes = $this->ClientesModel->get_personas('','',$tipoCli,$busqueda);
        return $this->response->setJSON($clientes);
    }

    public function save_persona(){
        $cod = $this->request->getVar('cod');
        $ruc = $this->request->getVar('ruc');
        $nom = strtoupper($this->request->getVar('nom'));
        $dir = strtoupper($this->request->getVar('dir'));
        $tel = $this->request->getVar('tel');
        $est = $this->request->getVar('est');
        $nac = $this->request->getVar('nac')?$this->request->getVar('nac'):'';
        $his = $this->request->getVar('his')=='on'?1:0;
        $tps = $this->request->getVar('tps');
        $this->ClientesModel = new ClientesModel();
        if(empty($cod)){
            $cod = $this->ClientesModel->get_max_id();
        }
        if($est=='nuevo'){
            $data = array(
                    'CLI_CODCLIE' => $cod,          'CLI_CODCIA' => 25,         'CLI_CP' => $tps,            'CLI_NOMBRE' => substr($nom, 0, 60),
                    'CLI_NOMBRE_ESPOSO' => substr($nom, 0, 60),    'CLI_NOMBRE_ESPOSA' => '',  'CLI_NOMBRE_EMPRESA' => '', 'CLI_123' => 1,
                    'CLI_TELEF1' => $tel,           'CLI_TELEF2' => '',         'CLI_CASA_DIREC' => substr($dir, 0, 30),   'CLI_CASA_NUM' => 0,
                    'CLI_CASA_ZONA' => 0,           'CLI_CASA_SUBZONA' => 0,    'CLI_TRAB_DIREC' => substr($dir, 0, 30),   'CLI_TRAB_NUM' => 0,
                    'CLI_TRAB_ZONA' => 0,           'CLI_TRAB_SUBZONA' => 0,    'CLI_TRAB_PROV' => 0,       'CLI_RUC_ESPOSO' => strlen($ruc)==11?$ruc:'',
                    'CLI_RUC_ESPOSA' => strlen($ruc)==8?$ruc:'',                'CLI_RUC_EMPRESA' => '',    'CLI_CASA1' => '',
                    'CLI_CASA2' => '',              'CLI_REGPUB1' => '',        'CLI_REGPUB2' => '',        'CLI_AUTOAVALUO' => '',
                    'CLI_PRENDA' => '',             'CLI_AUTO1' => '07',           'CLI_AUTO2' => '',          'CLI_IGV_INCLUIDO' =>'',
                    'CLI_OTRO_CONTR' => 1,          'CLI_LETRA' => 0,           'CLI_LIMCRE' => 0,          'CLI_FECHA_FAC' => 0,
                    'CLI_TIPO_BLOQ1' => 1,          'CLI_TIPO_BLOQ2' => '',     'CLI_TIPO_BLOQ3' => '',     'CLI_TIPO_BLOQ4' => '',
                    'CLI_DET_TOT' =>  '',           'CLI_NOM_LET1' => '',       'CLI_NOM_LET2' => '',       'CLI_GRUPO' => 1,
                    'CLI_SUBGRUPO' => 0,            'CLI_DIVISION' => 0,        'CLI_ESTADO' => 'A',        'CLI_MONEDA' => 'S',
                    'CLI_CODART' => '',             'CLI_NUCLEO' =>  '',        'CLI_CUENTA_CONTAB' =>'',   'CLI_CIA_REF' =>'',
                    'CLI_PORDESCTO' => 0,           'CLI_SALDO' => 0,           'CLI_PRECIOS' =>'',         'CLI_DIA_VISITA' => 3,
                    'CLI_ZONA_NEW' => 0,            'CLI_PROGRAMADO' =>'',      'CLI_LUGAR_CASA' => 1,      'CLI_LUGAR_TRAB' => 1,
                    'CLI_CUENTA_CONTAB2' => 1,      'CLI_DIAS_CRED' => 0,       'CLI_DIAS_FAC' => 2,        'CLI_CUENTA_CONTAB22' =>'',
                    'CLI_LIMCRE2' => 0,             'CLI_TIPO' =>'',            'CLI_FECHAHORA' => 0,       'CLI_CIARELA' =>'',
                    'CLI_MARCAID' =>'',         'CLI_TIPOCLI' => 7,         'CLI_FECHA_NAC' => $nac,
                    'CLI_HISTORIA' => $his
                );
            $data2 = array(
                'CODCIA' =>	25,
                'CODCLI' => $cod,
                'CP' =>	$tps,
                'DIREC' => substr($dir, 0, 50),
                'REF' => '',
                'CLI_LUGAR_TRAB' => 0,
                'CLI_TRAB_ZONA' => 0,
                'CLI_CASA_SUBZONA' => 0,
                'CLI_TRAB_SUBZONA' => 0,
                'NUMERO' => 0,
                'DIRCOMP' => substr($dir, 0, 100)
            );
            $this->ClientesModel = new ClientesModel();
            $clientes = $this->ClientesModel->set_persona($data);
            $clientes2 = $this->ClientesModel->set_dir_persona($data2);      
        }
        if($est=='editar'){
            $data = array(
                        'CLI_CODCLIE' => $cod,                          'CLI_NOMBRE' => substr($nom, 0, 60),
                        'CLI_NOMBRE_ESPOSO' =>substr($nom, 0, 60),      'CLI_TELEF1' => $tel,           
                        'CLI_CASA_DIREC' => substr($dir, 0, 30),        'CLI_TRAB_DIREC' => substr($dir, 0, 30),   
                        'CLI_FECHA_NAC' => $nac,
                        'CLI_RUC_ESPOSO' => strlen($ruc)==11?$ruc:'',
                        'CLI_RUC_ESPOSA' => strlen($ruc)==8?$ruc:'',
                        'CLI_CP' => $tps
                    );
            $this->ClientesModel = new ClientesModel();
            $clientes = $this->ClientesModel->editar_persona($data,0);
        }        
        return $cod;     
    }

    public function get_persona_sunat(){
            
        $ruc = $this->request->getVar('ruc');
        $tam = strlen($ruc);
        $tip = $tam==8?1:($tam==11?2:0);
        $this->ClientesModel = new ClientesModel();
        $codcli = $this->ClientesModel->get_pos_id($ruc,$tip);
        if(empty($codcli)){
            $new_id = $this->ClientesModel->get_max_id();
        }else{
            //$this->Clientes2Model = new ClientesModel();
            $clientes = $this->ClientesModel->get_personas($codcli,'','','');
            $tip = 3;
        }
        
        if($tip==1){
            ob_start(); 
            $data = file_get_contents("http://dniruc.apisperu.com/api/v1/dni/".$ruc."?token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJlbWFpbCI6IlBFUkVaVklMTEFMVEFAR01BSUwuQ09NIn0.cE-L-BK4BuFEIbyFwYT8ACFSobwLeVK7Xc4u9QAuvn0");
            $info = json_decode($data, true);            
            if($data==='[]'){
                $datos = array(0 => 'nada');
                echo json_encode($datos);
                }else{
                $datos = array(
                0 => $info['dni'], 
                1 => $info['nombres']." ".$info['apellidoPaterno']." ".$info['apellidoMaterno'],
                2 => '',
                3 => '',
                4 => $new_id,
                5 => 'nuevo',
                6 => ''
                );
                echo json_encode($datos);
            }
            ob_end_flush();
        }elseif($tip==2){
            ob_start(); 
            $data = file_get_contents("http://dniruc.apisperu.com/api/v1/ruc/".$ruc."?token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJlbWFpbCI6IlBFUkVaVklMTEFMVEFAR01BSUwuQ09NIn0.cE-L-BK4BuFEIbyFwYT8ACFSobwLeVK7Xc4u9QAuvn0");
            $info = json_decode($data, true);            
            if($data==='[]'){
                $datos = array(0 => 'nada');
                echo json_encode($datos);
                }else{
                $datos = array(
                0 => $info['ruc'], 
                1 => $info['razonSocial'],
                2 => $info['direccion'],
                3 => $info['nombreComercial'],
                4 => $new_id,
                5 => 'nuevo',
                6 => ''
                );
                echo json_encode($datos);
            }
            ob_end_flush();
        }elseif($tip==3){
            $client = $clientes[0];
            $datos = array(
                0 => $tam==8?$client->CLI_RUC_ESPOSA:($tam==11?$client->CLI_RUC_ESPOSO:0), 
                1 => $client->CLI_NOMBRE,
                2 => $client->CLI_CASA_DIREC,
                3 => $client->CLI_TELEF1,
                4 => $client->CLI_CODCLIE,
                5 => 'existe',
                6 => $client->CLI_FECHA_NAC
                );
                echo json_encode($datos);
        }
        
         
        
    }


}