<?php

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 * Description of PHPExcelHandler
 *
 * @author Thejan
 */
include './PHPExcel.php';

class PHPExcelHandler {

    public $objPHPExcel;
    private $objPHPWriter;
    private $currentRow;

    public function __construct() {
        $this->objPHPExcel = new PHPExcel();
        $this->currentRow = 1;
    }

    /**
     * Set Page Setup

     *'orientation' => PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE,

     *  'paperSize' => PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4,

     *  'fitToHeight' => 1,

     * 'FirstPageNumber' => 1

     *
     * @param array $pageSetupArray
     */
    public function setPageSetup($pageSetupArray) {
        if (array_key_exists('orientation', $pageSetupArray))
            $this->objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation($pageSetupArray['orientation']);

        if (array_key_exists('paperSize', $pageSetupArray))
            $this->objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize($pageSetupArray['paperSize']);

        if (array_key_exists('fitToHeight', $pageSetupArray))
            $this->objPHPExcel->getActiveSheet()->getPageSetup()->setFitToHeight($pageSetupArray['fitToHeight']);

        if (array_key_exists('FirstPageNumber', $pageSetupArray))
            $this->objPHPExcel->getActiveSheet()->getPageSetup()->setFirstPageNumber($pageSetupArray['FirstPageNumber']);
    }

    /**
     * Set Metadata of the file

     * 'creator' => "Creator",

    'lastModifiedBy' => "Last Modified",

    'title' => "Title",

    'subject' => "Subject",

    'description' => "Description",

    'keywords' => "Keywords test",

    'catagory' => "Catagory",

    'company' => "Company"

     * @param type $metadataArray
     */
    public function setMetadata($metadataArray) {
        if (array_key_exists('creator', $metadataArray))
            $this->objPHPExcel->getProperties()->setCreator($metadataArray['creator']);

        if (array_key_exists('lastModifiedBy', $metadataArray))
            $this->objPHPExcel->getProperties()->setLastModifiedBy($metadataArray['lastModifiedBy']);

        if (array_key_exists('title', $metadataArray))
            $this->objPHPExcel->getProperties()->setTitle($metadataArray['title']);

        if (array_key_exists('subject', $metadataArray))
            $this->objPHPExcel->getProperties()->setSubject($metadataArray['subject']);

        if (array_key_exists('description', $metadataArray))
            $this->objPHPExcel->getProperties()->setDescription($metadataArray['description']);

        if (array_key_exists('keywords', $metadataArray))
            $this->objPHPExcel->getProperties()->setKeywords($metadataArray['keywords']);

        if (array_key_exists('catagory', $metadataArray))
            $this->objPHPExcel->getProperties()->setCategory($metadataArray['catagory']);

        if (array_key_exists('company', $metadataArray))
            $this->objPHPExcel->getProperties()->setCompany($metadataArray['company']);
    }

    public function setHeaderStyle($range, $styleArray) {
        $this->objPHPExcel->getActiveSheet()->getStyle($range)->applyFromArray($styleArray);
    }

    public function setHeaderStyleByColumnAndRow($range = array(0, 0, 0, 0), $styleArray) {
        //range[c1,r1,c2,r2]
        $this->objPHPExcel->getActiveSheet()->getStyleByColumnAndRow($range[0], $range[1], $range[2], $range[3])->applyFromArray($styleArray);
    }

    public function setColumnWidth($columnWidthArray) {
        foreach ($columnWidthArray as $key => $value) {
            $this->objPHPExcel->getActiveSheet()->getColumnDimensionByColumn($key)->setWidth($value);
        }
    }

    public function setHeader($headerTextArray, $row = 1) {
        $this->currentRow = $row;
        foreach ($headerTextArray as $column => $value) {
            $this->objPHPExcel->getActiveSheet()->getCellByColumnAndRow($column, $row)->setValue($value);
        }
        return ++$this->currentRow;
    }

    public function setData($dataArray, $startRow = 1) {
        $row = $startRow;
        foreach ($dataArray as $rowKey => $dataRow) {
            $column = 0;
            foreach ($dataRow as $columnKey => $value) {
                $this->objPHPExcel->getActiveSheet()->getCellByColumnAndRow($column++, $row)->setValue($value);
            }
            $row++;
        }
        return $this->currentRow = $row;
    }

    public function save($filename = NULL) {
        if ($filename == NULL)
            $filename = date('Y-m-d-G-i-s') . '.xlsx';
        $this->objPHPWriter = new PHPExcel_Writer_Excel2007($this->objPHPExcel);
        $this->objPHPWriter->save($filename);
        return $filename;
    }

    public function download($filename = NULL) {
        $filename = $this->save($filename);
        header("Location:" . $filename);
    }

}
