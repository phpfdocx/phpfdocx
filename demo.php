<?php
/*
 * PhpFDocx 1.0.0
 *
 * @link https://github.com/hfornazier/PhpFDocx
 * @author Humberto Fornazier
 * contact: phpfdocx@gmail.com
 * @since 05.03.2020 
*/

/* 
    ************************* VERY IMPORTANT: ***********************
	1) Variables to be searched for and replaced must be in lowercase
	
    Do not use numbers in the variables to be replaced.
    Example: ${phone3}, ${name1}, {$email4) ...

	Variables cannot have spaces, symbols or special characters. 	
	Example:
	my home = must be myhome
	my_home = must be myhome
 	

*/
include 'PhpFDocx.php';

$doc          = 'test1.docx';
//$doc        = 'test2.docx';
//$doc        = 'test3.docx';
//$doc        = 'test4.docx';
//$doc        = 'test5.docx';
//$doc        = 'test6.docx';

$aDataSearch = Array (  'name'                    ,
						'address'                 ,
						'city'                    ,
						'region'                  ,
						'country'                 ,
						'zip'                     ,
						'personalnumber'          ,
						'organizationnumber'      ,				   
						'email'                   ,				   
						'today'                   ,
                        'header'                  ,						
                        'titledocumentchange'     );

$aDataChange = Array (  'Greyce, MacKensie X.'    ,
						'Ap #537-6485 Morbi Road' , 
						'Saint-Georges'           , 
						'Konya'                   , 
						'Austria'                 ,
						'741488'                  ,
						'16911214 2808'           ,
						'16911468 1404'           ,				   
						'placerat.orci@quam.edu4' ,
						date('Y/m/d')             ,	
						'New header Changed'      ,						
						'Document title changed'  );
						
$result = PhpFDocx( $doc , $aDataSearch , $aDataChange );	

echo '<a href="'.$result.'">View Generated Document = <b>'.$result.'</b></a><br /><br />';	
	
?>				   
