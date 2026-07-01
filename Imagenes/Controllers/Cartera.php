<?php

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

namespace App\Controllers;
use App\Models\CarteraModel;

/**
 * Description of Personas
 *
 * @author José Luis
 */
class Cartera extends BaseController
{
public function index(){
    return view('cartera/index');
}


	public function get_ctas_x_pagar()
	{
        
        $this->CarteraModel = new CarteraModel();
        $clientes = $this->CarteraModel->get_por_pagar('','');
        echo json_encode($clientes);

	}
}