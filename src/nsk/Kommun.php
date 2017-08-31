<?php
/**
 * Created by PhpStorm.
 * User: tobias
 * Date: 2017-08-31
 * Time: 01:05
 */

namespace nsk;


class Kommun
{
    public $KommunId;
    public $Foerening;

    public function __construct($kommunId) {
        $this->KommunId = $kommunId;
    }
}