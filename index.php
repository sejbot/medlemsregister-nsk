<?php
require('vendor/autoload.php');
$app = new \nsk\App();
$inFile = $argv[1];
$app->run($inFile);
