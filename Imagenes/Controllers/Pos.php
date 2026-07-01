<?php

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

namespace App\Controllers;

use CodeIgniter\I18n\Time;

/**
 * Description of Caja
 *
 * @author José Luis
 */
class Pos extends BaseController {
public function index(){
    return view('pos/index');
}
}