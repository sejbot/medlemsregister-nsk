<?php
/**
 * Created by PhpStorm.
 * User: tobias
 * Date: 2017-08-30
 * Time: 21:14
 */

namespace nsk;


use Sabre\Xml\Service;
use Sabre\Xml\Writer;

class App
{

    private $deltagarLista;
    private $ledarLista;
    public function run($fileName) {
        ini_set("default_charset", "utf-8");
        $fileNameInfo = pathinfo($fileName);
        $excelObject = \PHPExcel_IOFactory::load(__DIR__."/../../data/$fileName");
        $sheetData = $excelObject->getActiveSheet();
        $service = $this->setupXMLService();
        $specList = $this->createSpecListFromFile($fileNameInfo);
        $this->deltagarLista = $this->createDeltagarListaFromExcel($sheetData);
        $this->ledarLista = $this->createLedarListaFromExcel($sheetData);
        $aktivitetskort = new Aktivitetskort();
        $kommun = new Kommun("1480");
        $foerening = $this->createFoerening();
        $naervarokort = $this->createNaervaroKortListaFromExcel($sheetData, $specList);
        $foerening->Naervarokort = $naervarokort;
        $kommun->Foerening = $foerening;
        $aktivitetskort->Kommun = $kommun;
        $aktivitetskort->DeltagarRegister = new DeltagarRegister();
        $aktivitetskort->DeltagarRegister->Deltagare = $this->deltagarLista;
        $aktivitetskort->LedarRegister = new LedarRegister();
        $aktivitetskort->LedarRegister->Ledare= $this->ledarLista;
        $xml = $service->writeValueObject($aktivitetskort);
        $xml = str_replace(" xmlns=\"\"", "", $xml);
        $xml = str_replace("<?xml version=\"1.0\"?>", "<?xml version=\"1.0\" encoding=\"utf-8\"?>", $xml);

        file_put_contents(__DIR__."/../../data/".$fileNameInfo['filename'].".xml", $xml);
    }

    private function setupXMLService() {
        $service = new Service();
        $service->mapValueObject("{}Aktivitetskort", Aktivitetskort::class);
        $service->mapValueObject("{}Kommun", Kommun::class);
        $service->mapValueObject("{}Foerening", Foerening::class);
        $service->mapValueObject("{}Naervarokort", Naervarokort::class);
        $service->mapValueObject("{}Sammankomster", Sammankomster::class);
        $service->mapValueObject("{}Sammankomst", Sammankomst::class);
        $service->mapValueObject("{}DeltagarLista", DeltagarLista::class);
        $service->mapValueObject("{}Deltagare", DeltagarStatus::class);
        $service->mapValueObject("{}LedarLista", LedarLista::class);
        $service->mapValueObject("{}Ledare", LedarStatus::class);
        $service->mapValueObject("{}DeltagarRegister", DeltagarRegister::class);
        $service->mapValueObject("{}DeltagarRegister/Deltagare", Deltagare::class);
        $service->mapValueObject("{}LedarRegister", LedarRegister::class);
        $service->mapValueObject("{}LedarRegister/Ledare", Ledare::class);

        $service->classMap[Aktivitetskort::class] = function(Writer $writer, Aktivitetskort $aktivitetskort) {
            $writer->writeAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
            $writer->writeAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema");
            $writer->writeAttribute("xmlns", "http://aktivitetskort.net.umea.se/importSchema.xsd");
            foreach(get_object_vars($aktivitetskort) as $key=> $value ) {
                $writer->writeElement($key, $value);
            }
        };
        $service->classMap[Kommun::class] = function(Writer $writer, Kommun $kommun) {
            $writer->writeAttribute("kommunID", $kommun->KommunId);
            $writer->writeElement("Foerening", $kommun->Foerening);
        };

        $service->classMap[Foerening::class] = function(Writer $writer, Foerening $foerening) {
            $writer->writeAttribute("foereningsID", $foerening->foereningsID);
            $writer->writeAttribute("foereningsNamn", $foerening->foereningsNamn);
            $writer->writeAttribute("organisationsnummer", $foerening->organisationsnummer);
            foreach($foerening->Naervarokort as $naervarokort) {
                $writer->writeElement("Naervarokort", $naervarokort);
            }
            $writer->writeElement("BorttagnaSammankomster", $foerening->BorttagnaSammankomster);
        };

        $service->classMap[Naervarokort::class] = function(Writer $writer, Naervarokort $naervarokort) {
            $writer->writeAttribute("NaervarokortNummer", $naervarokort->NaervarokortNummer);
            foreach(get_object_vars($naervarokort) as $key=>$value ) {
                if($key !== "NaervarokortNummer") {
                    $writer->writeElement($key, $value);
                }
            }
        };
        $service->classMap[Sammankomst::class] = function(Writer $writer, Sammankomst $sammankomst) {
            $writer->writeAttribute("Datum", $sammankomst->Datum);
            $writer->writeAttribute("kod", $sammankomst->kod);
            foreach(get_object_vars($sammankomst) as $key=> $value ) {
                if($key != "Datum" && $key != "kod") {
                    $writer->writeElement($key, $value);
                }
            }
        };
        $service->classMap[DeltagarStatus::class] = function(Writer $writer, DeltagarStatus $deltagarStatus) {
            $writer->writeAttribute("id", $deltagarStatus->id);
            foreach(get_object_vars($deltagarStatus) as $key=> $value ) {
                if($key != "id") {
                    $writer->writeElement($key, $value);
                }
            }
        };

        $service->classMap[LedarStatus::class] = function(Writer $writer, LedarStatus $ledarStatus) {
            $writer->writeAttribute("id", $ledarStatus->id);
            foreach(get_object_vars($ledarStatus) as $key=> $value ) {
                if($key != "id") {
                    $writer->writeElement($key, $value);
                }
            }
        };

        $service->classMap[Deltagare::class] = function(Writer $writer, Deltagare $deltagare) {
            $writer->writeAttribute("id", $deltagare->id);
            foreach(get_object_vars($deltagare) as $key=> $value ) {
                if($key != "id") {
                    $writer->writeElement($key, $value);
                }
            }
        };

        $service->classMap[Ledare::class] = function(Writer $writer, Ledare $ledare) {
            $writer->writeAttribute("id", $ledare->id);
            foreach(get_object_vars($ledare) as $key=> $value ) {
                if($key != "id") {
                    $writer->writeElement($key, $value);
                }
            }
        };


        return $service;
    }

    private function createSpecListFromFile($fileNameInfo) {
      $specFile = $fileNameInfo['filename'].'-spec.csv';
      $specPath = __DIR__."/../../data/".$specFile;
      $specList = array_map('str_getcsv', file($specPath));
      return $specList;
    }

    public function createFoerening() {
        $foerening = new Foerening();
        $foerening->foereningsID = "1"; //TODO: Har klubben något riktigt id?
        $foerening->foereningsNamn = "Nolereds Schackklubb";
        $foerening->organisationsnummer = "802472-9371";
        return $foerening;
    }

    public function createNaervarokortListaFromExcel(\PHPExcel_Worksheet $sheetData, $specList) {
        $naervarokortLista = [];
        foreach($specList as $spec) {
          $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, $spec[0], $spec[1], $spec[2], $spec[3]);
        }
        /*$naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 1, 8, 23);
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 2, 24, 24);
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 3, 25, 41);
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 4, 42, 58);
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 5, 59, 59);
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 6, 60, 62);
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 7, 63, 64, "Taevling");
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 8, 65, 66, "Taevling");
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 9, 67, 67, "Taevling");
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 10, 68, 68, "Taevling");
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 11, 69, 69, "Taevling");
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 12, 70, 70, "Taevling");
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 13, 71, 71, "Taevling");
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 14, 72, 73, "Taevling");
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 15, 74, 74, "Taevling");
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 16, 75, 75, "Taevling");
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 17, 76, 76, "Taevling");
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 18, 77, 77, "Taevling");
        $naervarokortLista[] = $this->createNaervarokortFromExcel($sheetData, 19, 78, 81, "Taevling");
        */
        return $naervarokortLista;
    }

    private function createNaervarokortFromExcel(\PHPExcel_Worksheet $sheetData, $number, $naervarokortStartColumn, $naervarokortEndColumn, $type = "Traening") {
        $properties = ["Grupp","Lokal"];
        $naervarokortStartRow = 2;
        $naervarokortEndRow = 3;
        $sammankomster = new Sammankomster();
        $naervarokort = new Naervarokort();
        $naervarokort->NaervarokortNummer = $number;
        $naervarokort->Aktivitet = "Schack"; //TODO: Samma för alla?
        $naervarokort->Grupp = $sheetData->getCellByColumnAndRow($naervarokortStartColumn, 2)->getValue();
        $naervarokort->Lokal = $sheetData->getCellByColumnAndRow($naervarokortStartColumn, 3)->getValue();
        for($col = $naervarokortStartColumn; $col <= $naervarokortEndColumn; $col++) {

            for ($row = $naervarokortStartRow; $row <= $naervarokortEndRow; $row++) {
                $propertyName = $properties[$row-$naervarokortStartRow];
                $propertyValue = $sheetData->getCellByColumnAndRow($col, $row);
            }
            $year = "2017";
            $month = $sheetData->getCellByColumnAndRow($col, 4)->getValue();
            $day = $sheetData->getCellByColumnAndRow($col, 5)->getValue();
            $code = $year;
            $startDateTime = \DateTime::createFromFormat("Y-n-j", "$year-$month-$day");
            $datum = $startDateTime->format("Y-m-d");
            $startHour =  $sheetData->getCellByColumnAndRow($naervarokortStartColumn, 6)->getValue();
            $endHour = $sheetData->getCellByColumnAndRow($naervarokortStartColumn, 7)->getValue();

            $sammankomst = new Sammankomst();
            $sammankomst->kod = $col;
            $sammankomst->Datum = $datum;
            $sammankomst->StartTid = \DateTime::createFromFormat("G", $startHour)->format("H:00:00.0000000+01:00");
            $sammankomst->StoppTid = \DateTime::createFromFormat("G", $endHour)->format("H:00:00.0000000+01:00");
            $sammankomst->SlutDatum = $datum;
            $sammankomst->Aktivitet = $naervarokort->Aktivitet;
            $sammankomst->Grupp = $naervarokort->Grupp;
            $sammankomst->Lokal = $naervarokort->Lokal;
            $sammankomst->Typ = $type;
            $sammankomst->Metod = "Add";
            $sammankomst->DeltagarLista = new DeltagarLista();
            $sammankomst->LedarLista = new LedarLista();

            $deltagarStartRow = 10;
            $deltagarEndRow = 36;
            for($deltagarRow = $deltagarStartRow; $deltagarRow <= $deltagarEndRow; $deltagarRow++) {
                $deltagarPersonNummer = $sheetData->getCellByColumnAndRow(5, $deltagarRow)->getValue();
                $isPresent = $sheetData->getCellByColumnAndRow($col,$deltagarRow)->getValue() == "x" ? "true" : "false";
                $deltagarStatus = new DeltagarStatus();
                $deltagare = $this->deltagarLista[$deltagarPersonNummer];
                $deltagarStatus->id = $deltagare->id;
                $deltagarStatus->Handikapp = "false";
                $deltagarStatus->Naervarande = $isPresent;
                $sammankomst->DeltagarLista->Deltagare[] = $deltagarStatus;
            }
            $ledarStartRow = 37;
            $ledarEndRow = 42;
            for($ledarRow = $ledarStartRow; $ledarRow <= $ledarEndRow; $ledarRow++) {
                $ledarPersonNummer = $sheetData->getCellByColumnAndRow(5, $ledarRow)->getValue();
                $isPresent = $sheetData->getCellByColumnAndRow($col,$ledarRow)->getValue() == "x" ? "true" : "false";
                $ledarStatus = new LedarStatus();
                $ledare = $this->ledarLista[$ledarPersonNummer];
                $ledarStatus->id = $ledare->id;
                $ledarStatus->Handikapp = "false";
                $ledarStatus->Naervarande = $isPresent;
                $sammankomst->LedarLista->Ledare[] = $ledarStatus;
            }

            $sammankomster->Sammankomst[] = $sammankomst;
        }
        $naervarokort->Sammankomster = $sammankomster;
        return $naervarokort;
    }

    public function createDeltagarListaFromExcel(\PHPExcel_Worksheet $sheetData) {
        $deltagarLista = [];
        $properties = ["Namn", "Postnr", "Postadress", "Kommun", "Personnummer", "Kon"];
        $municipalityCodes = ["Göteborgs Kommun" => "1480", "Stenungsunds Kommun" => "1415", "Öckerö Kommun" => "1407", "Partille Kommun" => "1402"];
        $deltagarStartRow = 10;
        $deltagarEndRow = 36;
        $deltagarStartColumn = 1;
        $deltagarEndColumn = 6;
        for($row = $deltagarStartRow; $row <= $deltagarEndRow; $row++) {
            $deltagare = new Deltagare();
            $deltagare->id = $row;
            for ($col = $deltagarStartColumn; $col <= $deltagarEndColumn; $col++) {
                $propertyName = $properties[$col-1];
                $propertyValue = $sheetData->getCellByColumnAndRow($col, $row)->getValue();
                if($propertyName === "Namn") {
                    $nameChunks = explode(" ", $propertyValue, 2);
                    $deltagare->Foernamn = $nameChunks[0];
                    $deltagare->Efternamn = $nameChunks[1];
                }
                else {
                    if ($propertyName === "Kon") {
                        $propertyValue = $propertyValue == "1" ? "Kvinna" : "Man";
                    }
                    else if ($propertyName === "Kommun") {
                        $propertyValue = $municipalityCodes[$propertyValue];
                    }
                    $deltagare->$propertyName = $propertyValue;
                }
            }

            $deltagarLista[$deltagare->Personnummer] = $deltagare;
        }
        return $deltagarLista;
    }

    public function createLedarListaFromExcel(\PHPExcel_Worksheet $sheetData) {
        $deltagarLista = [];
        $properties = ["Namn", "Postnr", "Postadress", "Kommun", "Personnummer", "Kon"];
        $municipalityCodes = ["Göteborgs Kommun" => "1480", "Stenungsunds Kommun" => "1415", "Öckerö Kommun" => "1407", "Partille Kommun" => "1402"];
        $deltagarStartRow = 37;
        $deltagarEndRow = 42;
        $deltagarStartColumn = 1;
        $deltagarEndColumn = 6;
        for($row = $deltagarStartRow; $row <= $deltagarEndRow; $row++) {
            $deltagare = new Ledare();
            $deltagare->id = $row;
            for ($col = $deltagarStartColumn; $col <= $deltagarEndColumn; $col++) {
                $propertyName = $properties[$col-1];
                $propertyValue = $sheetData->getCellByColumnAndRow($col, $row)->getValue();
                if($propertyName === "Namn") {
                    $nameChunks = explode(" ", $propertyValue, 2);
                    $deltagare->Foernamn = $nameChunks[0];
                    $deltagare->Efternamn = $nameChunks[1];
                }
                else {
                    if ($propertyName === "Kon") {
                        $propertyValue = $propertyValue == "1" ? "Kvinna" : "Man";
                    }
                    else if ($propertyName === "Kommun") {
                        $propertyValue = $municipalityCodes[$propertyValue];
                    }
                    $deltagare->$propertyName = $propertyValue;
                }
            }

            $deltagarLista[$deltagare->Personnummer] = $deltagare;
        }
        return $deltagarLista;
    }
}
