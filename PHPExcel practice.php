<?php

require 'PHPExcel-1.8\Classes\PHPExcel.php';
require 'PHPExcel-1.8\Classes\PHPExcel\Calculation.php';
require 'PHPExcel-1.8\Classes\PHPExcel\Cell.php';


    $objPHPExcelInputCsv = PHPExcel_IOFactory::load("inputcsv/place_id_original.csv");

    $placeIdsGC = $objPHPExcelInputCsv->getActiveSheet()->toArray(null,true,true,true);

    if(!empty($placeIdsGC)){
        $objPHPExcelGC = new PHPExcel();
        $objPHPExcelGC->getProperties()->setCreator("Placeid")
                    ->setLastModifiedBy("Automatic")
                    ->setCategory("Store Placeid Result");
        $objPHPExcelGC->getActiveSheet()->setTitle('Place List');
        
        $objPHPExcelGC->setActiveSheetIndex(0) 
                    ->setCellValue('A1', 'store_id'); 
                   
        $objPHPExcelGC->getActiveSheet()->setCellValue('B1', 'lat_GeoC');
        $objPHPExcelGC->getActiveSheet()->setCellValue('C1', 'lng_GeoC');
        
        $style = array(
            'alignment' => array(
                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            )
        );
        
        $sheet = $objPHPExcelGC->getActiveSheet();
        $sheet->getStyle("A1:E1")->applyFromArray($style);// applying the styles as stored in the array '$style' i.e. the Alignment is Horizontally centered.
        $objPHPExcelGC->getActiveSheet()->getStyle("A1:E1")->getFont()->setBold( true );// ->getColor()->setRGB('f00000'); // setting the styles of the columns A1 to AF1 as BOLD and text as RED
        $objPHPExcelGC->getActiveSheet()->getStyle('A1:E1')->getAlignment()->setWrapText(true); // Set alignments
        
        function cellColor($cells,$color){
            global $objPHPExcelGC;
        
            $objPHPExcelGC->getActiveSheet()->getStyle($cells)->getFill()->applyFromArray(array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,//  FILL_GRADIENT_LINEAR
                'startcolor' => array(
                     'rgb' => $color
                )
            ));
        }
        //cellColor('A5:E5', 'ff0000');	//Red color
        //cellColor('A5', 'ff0000');	//Red color
        
      

        echo '<pre>';
        print_r($placeIdsGC);
        //die;
        $geoclenseLocationList = array();
        $i = 2;
        foreach($placeIdsGC as $data){
            if($i>0){
                $tem['lat'] = $data['C'];
                $objPHPExcelGC->getActiveSheet()->getColumnDimension('B')->setWidth(12);
                $objPHPExcelGC->getActiveSheet()->getColumnDimension('C')->setWidth(12);
                $objPHPExcelGC->getActiveSheet()->setCellValue('B'.$i,$data['C']);
                $objPHPExcelGC->getActiveSheet()->setCellValue('C'.$i,$data['D']);
                $objPHPExcelGC->getActiveSheet()->getColumnDimension('A')->setWidth(10);
                $objPHPExcelGC->getActiveSheet()->setCellValue('A'.$i, $data['A']);
			
                $tem['long'] = $data['D'];
                $geoclenseLocationList[$data['A']] = $tem;
            }
            $i++;
        }
        
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcelGC, 'Excel2007');
	$filename = 'GClense_'.date('d-m-Y',time()).'.csv';
    $objWriter->save('C:\xampp\htdocs\store_placeid\store_placeid\outputcsv\ '.$filename); 
}


?>
