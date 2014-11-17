<?php
include './PHPExcelHandler.php';
$objPHPExcelHandler = new PHPExcelHandler();
$cRow;
    $pageSetupArray = array(
        'orientation' => PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE,
        'paperSize' => PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4,
        'fitToHeight' => 1,
        'fitToWidth' => 1,
        'FirstPageNumber' => 1,
        'pageMargins' => array(
            'top' => 1,
            'right' => 1,
            'bottom' => 1,
            'left' => 1,
        )
    );
$objPHPExcelHandler->setPageSetup($pageSetupArray);

$metadataArray = array(
    'creator' => "Creator",
    'lastModifiedBy' => "Last Modified",
    'title' => "Title",
    'subject' => "Subject",
    'description' => "Description",
    'keywords' => "Keywords test",
    'catagory' => "Catagory",
    'company' => "Company"
);
$objPHPExcelHandler->setMetadata($metadataArray);

$styleArrayHeader = array(
    'font' => array(
        'bold' => true,
    ),
    'alignment' => array(
        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
    ),
    'borders' => array(
        'top' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
        ),
    ),
    'fill' => array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'startcolor' => array(
            'argb' => 'FFA0A0A0',
        ),
    ),
);
$objPHPExcelHandler->setHeaderStyle("A1:C1", $styleArrayHeader);

//setting column width
$columnWidthArray = array(10, 50, 30);
$objPHPExcelHandler->setColumnWidth($columnWidthArray);

//Setting the header
$headerArray = array('Column1', 'Column2', 'Column3');
$cRow = $objPHPExcelHandler->setHeader($headerArray);

$data = array(
    array(1, 'Row 1 Col 2', 'Row 1 Col 3'),
    array(2, 'Row 2 Col 2', 'Row 2 Col 3'),
);
$cRow = $objPHPExcelHandler->setData($data, $cRow);

$objPHPExcelHandler->setHeaderStyleByColumnAndRow(array(0, $cRow, 3, $cRow), $styleArrayHeader);
$cRow = $objPHPExcelHandler->setHeader($headerArray, $cRow);

$objPHPExcelHandler->download();
?>