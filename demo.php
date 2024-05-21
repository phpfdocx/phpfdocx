<?php
/*
 * PhpFDocx 2.0.0
 *
 * @link https://github.com/phpfdocx/phpfdocx
 * @author Humberto Fornazier - Brazil
 * contact: phpfdocx@gmail.com
 * @since 05.03.2020 
*/

/*
*/
include 'PhpFDocx.php';
//$doc  = 'test1.docx';
//$doc  = 'test2.docx';
//$doc  = 'test3.docx';
//$doc  = 'test4.docx';
//$doc  = 'test5.docx';
//$doc  = 'test6.docx';

$doc = 'test1.docx';

$aDataSearch = Array (  
                        'name'                    ,					
			'NAME'                    ,
			'address'                 ,
			'ADDRESS'                 ,
			'city'                    ,
			'CITY'                    ,						
			'region'                  ,
			'REGION'                  ,						
			'country'                 ,
			'COUNTRY'                 ,						
			'zip'                     ,
			'personalnumber'          ,
			'personal-number'         ,						
			'organizationnumber'      ,				   
			'email'                   ,				   
			'today'                   , 				
                        'documenttitle'           ,
			'[test$variable]'         ,				
			'test_complex!and@long-variable|for#tests'					
		      );

$aDataChange = Array (  
                        'Greyce, MacKensie X.'      ,
                        'GREYCE, MACKENSIE X.'      ,						
			'Ap #288-1785-B5 Morbi Road', 
			'AP #288-1785-B5 MORBI ROAD',
			'Saint-Georges'             , 
			'SAINT-GEORGES'             , 						
			'Konya'                     , 
			'KONYA'                     ,						
			'Austria'                   ,
			'AUSTRIA'                   ,						
			'741488'                    ,
			'16911214 2808'             ,						
			'16911214 2808'             ,							
			'16911468 1404'             ,				   
			'placerat.orci@quam.edu'    ,
			date('Y/m/d')               ,				
			'Document Title Changed'    ,
			'Teste Variable Changed'    ,
			'test variable complex and long changed'					
		      );				
						
$result = PhpFDocx( $doc , $aDataSearch , $aDataChange );	

echo '<a href="'.$result.'">View Generated Document = <b>'.$result.'</b></a><br /><br />';		
?>				   
