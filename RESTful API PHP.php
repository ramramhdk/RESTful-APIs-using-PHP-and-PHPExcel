<?php

require 'PHPExcel-1.8\Classes\PHPExcel.php';
require 'PHPExcel-1.8\Classes\PHPExcel\Calculation.php';
require 'PHPExcel-1.8\Classes\PHPExcel\Cell.php';


$objPHPExcelInputCsv = PHPExcel_IOFactory::load("inputcsv/place_id_original.csv");//taking file as INPUT like importing File
$placeIds = $objPHPExcelInputCsv->getActiveSheet()->toArray(null,true,true,true);
// $placeIds is now an array and I can use it with no problem.
//echo '<pre>';
//toArray(NULL,TRUE,TRUE); -> will return all the cell values in the worksheet (calculated and formatted) exactly as they appear in Excel itself.
unset($placeIds[1]);




$apiKey = 'AUzaSyDS2Jver6g51df6v1df6vd8BKDarc';


if(!empty($placeIds)){
	$objPHPExcel = new PHPExcel();// Create new PHPExcel object
	$objPHPExcel->getProperties()->setCreator("Placeid") // Set document properties
				->setLastModifiedBy("Automatic")
				->setCategory("Store Placeid Result");
	$objPHPExcel->getActiveSheet()->setTitle('Place List');
	
	$objPHPExcel->setActiveSheetIndex(0) // Create a first sheet, representing sales data
				->setCellValue('A1', 'store_id') // setting the value 'store_id' to the cell 'A1'
				->setCellValue('B1', 'place_id');
	$objPHPExcel->getActiveSheet()->mergeCells('C1:M1');
	$objPHPExcel->getActiveSheet()->setCellValue('C1', 'address_component');
	$objPHPExcel->getActiveSheet()->setCellValue('C2', 'subpremise');
	$objPHPExcel->getActiveSheet()->setCellValue('D2', 'premise');
	$objPHPExcel->getActiveSheet()->setCellValue('E2', 'neighborhood');
	$objPHPExcel->getActiveSheet()->setCellValue('F2', 'sublocality_level_3');
	$objPHPExcel->getActiveSheet()->setCellValue('G2', 'sublocality_level_2');
	$objPHPExcel->getActiveSheet()->setCellValue('H2', 'sublocality_level_1');
	$objPHPExcel->getActiveSheet()->setCellValue('I2', 'locality');
	$objPHPExcel->getActiveSheet()->setCellValue('J2', 'adminstrative_area_level_2');
	$objPHPExcel->getActiveSheet()->setCellValue('K2', 'administrative_area_level_1');
	$objPHPExcel->getActiveSheet()->setCellValue('L2', 'country');
	$objPHPExcel->getActiveSheet()->setCellValue('M2', 'postal_code');
	$objPHPExcel->getActiveSheet()->setCellValue('N1','formatted_address');
	$objPHPExcel->getActiveSheet()->mergeCells('O1:P1');
	$objPHPExcel->getActiveSheet()->setCellValue('O1', 'geometry');
	$objPHPExcel->getActiveSheet()->setCellValue('O2', 'location_lat');
	$objPHPExcel->getActiveSheet()->setCellValue('P2', 'location_long');
	$objPHPExcel->getActiveSheet()->setCellValue('Q1', 'icon');
	$objPHPExcel->getActiveSheet()->setCellValue('R1', 'id');
	$objPHPExcel->getActiveSheet()->setCellValue('S1', 'name');
	$objPHPExcel->getActiveSheet()->mergeCells('T1:U1');
	$objPHPExcel->getActiveSheet()->setCellValue('T1', 'plus_code');
	$objPHPExcel->getActiveSheet()->setCellValue('T2', 'compound_code');
	$objPHPExcel->getActiveSheet()->setCellValue('U2', 'global_code');
	$objPHPExcel->getActiveSheet()->setCellValue('V1', 'url');
	$objPHPExcel->getActiveSheet()->setCellValue('W1', 'utc_offset');
	$objPHPExcel->getActiveSheet()->setCellValue('X1', 'vicinity');
	$objPHPExcel->getActiveSheet()->setCellValue('Y1', 'scope');
	$objPHPExcel->getActiveSheet()->setCellValue('Z1', 'rating');
	$objPHPExcel->getActiveSheet()->setCellValue('AA1', 'international_phone_number');
	$objPHPExcel->getActiveSheet()->setCellValue('AB1', 'Website');
	$objPHPExcel->getActiveSheet()->mergeCells('AC1:AF1');
	$objPHPExcel->getActiveSheet()->setCellValue('AC1', 'Viewport');
	$objPHPExcel->getActiveSheet()->setCellValue('AC2', 'northeast_lat');
	$objPHPExcel->getActiveSheet()->setCellValue('AD2', 'northeast_long');

	$objPHPExcel->getActiveSheet()->setCellValue('AE2', 'southwest_lat');
	$objPHPExcel->getActiveSheet()->setCellValue('AF2', 'southwest_long');
	

	$style = array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        )
	);
	//the above multi-dimensional assosiative array is similar to: -
	//$style = array('alignment' => 
	//				 array('horizontal' => 54,));
	// where echo $style['alignment']['horizontal']; outputs = 54
	// OR
	//$style = array(
	//				 'x' => 
	//				 array('y' => 54,));
	// where echo $style['x']['y']; outputs = 54

	$sheet = $objPHPExcel->getActiveSheet();
	$sheet->getStyle("A1:V1")->applyFromArray($style);// applying the styles as stored in the array '$style' i.e. the Alignment is Horizontally centered.
	$sheet->getStyle("A2:V2")->applyFromArray($style);
	$objPHPExcel->getActiveSheet()->getStyle("A1:AH1")->getFont()->setBold( true );// ->getColor()->setRGB('f00000'); // setting the styles of the columns A1 to AF1 as BOLD and text as RED
	//$variable_name->getActiveSheet()->getStyle("Column_name(s)")->getFont()->setBold(true)
	//->setName('Verdana')
	//->setSize(10)
	//->getColor()->setRGB('6F6F6F');
	$objPHPExcel->getActiveSheet()->getStyle("A2:AF2")->getFont()->setBold( true );// ->getColor()->setRGB('FFFFFF');// setting the styles of the columns A2 to AF2 as BOLD
	$objPHPExcel->getActiveSheet()->getStyle('A1:AF1')->getAlignment()->setWrapText(true); // Set alignments
	$objPHPExcel->getActiveSheet()->getStyle('A2:AF2')->getAlignment()->setWrapText(true); 
	
	
	function cellColor($cells,$color){
		global $objPHPExcel;
	
		$objPHPExcel->getActiveSheet()->getStyle($cells)->getFill()->applyFromArray(array(
			'type' => PHPExcel_Style_Fill::FILL_SOLID,//  FILL_GRADIENT_LINEAR
			'startcolor' => array(
				 'rgb' => $color
			)
		));
	}
	
	//cellColor('B5', 'F28A8C');
	//cellColor('G5', 'F28A8C');
	cellColor('A5:AF5', 'ff0000');	//Red color
	
	
	$i = 3;
	foreach($placeIds as $place){
		$baseUrl = 'https://maps.googleapis.com/maps/api/place/details/json';
		$baseUrl .= '?placeid='.$place['B'];
		$baseUrl .= '&key='.$apiKey;
		//$baseUrl .= '&fields=address_component,adr_address,alt_id,formatted_address,geometry,icon,id,name,permanently_closed,';
		//$baseUrl .= 'place_id,plus_code,scope,type,url,utc_offset,vicinity';
		$ch = curl_init($baseUrl); // curl_init — Initializes a new session and return a cURL handle for use with the curl_setopt(), curl_exec(),and curl_close() functions. 
		// Description = curl_init([ string $url = NULL] ) where url - If provided, the CURLOPT_URL option will be set to its value. You can manually set this using the curl_setopt() function. 
		curl_setopt($ch, CURLOPT_CUSTOMREQUEST, "GET"); //HTTPGET , //curl_setopt = Sets an option on the given cURL session handle.                                                                 
		// CURLOPT_CUSTOMREQUEST = A custom request method to use instead of "GET" or "HEAD" when doing a HTTP request. 
		//This is useful for doing "DELETE" or other, more obscure HTTP requests.
		//Valid values are things like "GET", "POST", "CONNECT" and so on; i.e.
		// Do not enter a whole HTTP request line here.
		// For instance,entering "GET /index.html HTTP/1.0\r\n\r\n"would be incorrect. 
		// Description = curl_setopt( resource $ch, int $option, mixed $value) : bool
		curl_setopt($ch, CURLOPT_RETURNTRANSFER, true); // TRUE to return the transfer as a string of thereturn value of curl_exec() instead of outputting it directly.
		curl_setopt($ch, CURLOPT_HTTPHEADER, array( //An array of HTTP header fields to set, in the format array('Content-type: text/plain', 'Content-length: 100')
			'Content-Type: application/json',
			'Connection: Keep-Alive',
			"User-Agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.98 Safari/537.36"
			)
		);

		curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, 0);// 1 to check the existence of a common name in theSSL peer certificate.
		// 2 to check the existence ofa common name and also verify that it matches the hostnameprovided. 0 to not check the names. 
		//In production environments the value of this optionshould be kept at 2 (default value). 
        curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, 0);    // FALSE to stop cURL from verifying the peer'scertificate. TRUE, by default.                                      
		$result = curl_exec($ch); //curl_exec — Perform a cURL session
		//Execute the given cURL session. 
		//This function should be called after initializing a cURL session and all the options for the session are set. 
		// Description -> curl_exec( resource $ch) where ch= A cURL handle returned by curl_init().
		//Returns TRUE on success or FALSE on failure. 
		//However, if the CURLOPT_RETURNTRANSFERoption is set, it will returnthe result on success, FALSE on failure. 
		$resultData = json_decode($result);//Takes a JSON encoded string and converts it into a PHP variable. 

		curl_close($ch);// Closes a cURL session and frees all resources. The cURL handle, ch, is also deleted.
		
		if($resultData->status == 'OK'){
			$objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $place['A']);
			$objPHPExcel->getActiveSheet()->setCellValue('B'.$i, $place['B']);
			if(!empty($resultData->result->address_components)){
				//print_r($resultData->result->address_components);
				foreach($resultData->result->address_components as $add){
					// it will traverse through each component that has been stored in $result->data, and will do the following actions:
					switch($add->types[0]){
						case 'subpremise' : 
							$objPHPExcel->getActiveSheet()->setCellValue('C'.$i, $add->long_name);
							break;

						case 'premise' : 
						$objPHPExcel->getActiveSheet()->setCellValue('D'.$i, $add->long_name);
						break;

						case 'neighborhood' : 
						$objPHPExcel->getActiveSheet()->setCellValue('E'.$i, $add->long_name);
						break;

						case 'sublocality_level_3' : 
						$objPHPExcel->getActiveSheet()->setCellValue('F'.$i, $add->long_name);
						break;

						case 'sublocality_level_2' : 
						$objPHPExcel->getActiveSheet()->setCellValue('G'.$i, $add->long_name);
						break;

						case 'sublocality_level_1' : 
						$objPHPExcel->getActiveSheet()->setCellValue('H'.$i, $add->long_name);
						break;

						case 'locality' : 
						$objPHPExcel->getActiveSheet()->setCellValue('I'.$i, $add->long_name);
						break;

						case 'administrative_area_level_2' : 
						$objPHPExcel->getActiveSheet()->setCellValue('J'.$i, $add->long_name);
						break;

						case 'administrative_area_level_1' : 
						$objPHPExcel->getActiveSheet()->setCellValue('K'.$i, $add->long_name);
						break;

						case 'country' : 
						$objPHPExcel->getActiveSheet()->setCellValue('L'.$i, $add->long_name);
						break;

						case 'postal_code' : 
							$objPHPExcel->getActiveSheet()->setCellValue('M'.$i, $add->long_name);
						break;
					}
				}
				$objPHPExcel->getActiveSheet()->setCellValue('N'.$i, $resultData->result->formatted_address); //formatted_address is a string containing the human-readableaddress of this location
				$objPHPExcel->getActiveSheet()->setCellValue('O'.$i, $resultData->result->geometry->location->lat);
				$objPHPExcel->getActiveSheet()->setCellValue('P'.$i, $resultData->result->geometry->location->lng);
				$objPHPExcel->getActiveSheet()->setCellValue('Q'.$i, $resultData->result->icon);// like Q3,Q4,Q5.etc
				$objPHPExcel->getActiveSheet()->setCellValue('R'.$i, $resultData->result->id);
				$objPHPExcel->getActiveSheet()->setCellValue('S'.$i, $resultData->result->name);
				$objPHPExcel->getActiveSheet()->setCellValue('T'.$i, $resultData->result->plus_code->compound_code);// compound_code is a 6 character or longer local code with an explicit location(CWC8+R9, Mountain View, CA, USA).
				$objPHPExcel->getActiveSheet()->setCellValue('U'.$i, $resultData->result->plus_code->global_code);// global_code is a 4 character area code and 6 character or longer local code(849VCWC8+R9).
				$objPHPExcel->getActiveSheet()->setCellValue('V'.$i, $resultData->result->url);
				$objPHPExcel->getActiveSheet()->setCellValue('W'.$i, $resultData->result->utc_offset);
				$objPHPExcel->getActiveSheet()->setCellValue('X'.$i, $resultData->result->vicinity);
				$objPHPExcel->getActiveSheet()->setCellValue('Y'.$i, $resultData->result->scope);
				$objPHPExcel->getActiveSheet()->setCellValue('Z'.$i, $resultData->result->rating);
				$objPHPExcel->getActiveSheet()->setCellValue('AA'.$i, $resultData->result->international_phone_number);
				$objPHPExcel->getActiveSheet()->setCellValue('AB'.$i, $resultData->result->website);
				$objPHPExcel->getActiveSheet()->setCellValue('AC'.$i, $resultData->result->geometry->viewport->northeast->lat);
				$objPHPExcel->getActiveSheet()->setCellValue('AD'.$i, $resultData->result->geometry->viewport->northeast->lng);
				$objPHPExcel->getActiveSheet()->setCellValue('AE'.$i, $resultData->result->geometry->viewport->southwest->lat);
				$objPHPExcel->getActiveSheet()->setCellValue('AF'.$i, $resultData->result->geometry->viewport->southwest->lng);
				
			}
		}
		$i++;
	}
	
	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');// Save Excel 2007 file
	$filename = 'place_'.date('d-m-Y',time()).'.csv';// format of the file name that would be generated.
	$objWriter->save('C:\xampp\htdocs\store_placeid\store_placeid\outputcsv\ '.$filename); //where to save that file
}




	?>
