<?php
/*
 * PhpFDocx 3.0.0
 *
 * @link https://github.com/phpfdocx/phpfdocx
 * @author Humberto Fornazier
 * contact: phpfdocx@gmail.com
 * @since 05.03.2020 
*/

/* 	
*/
function PhpFDocx( $doc , $aDataSearch , $aDataChange ) {	

	$newdoc    = removeSymbolsStr( $doc ) . '-' . generatePrefix() . '.docx'; 	
	$docsPath  = 'doc/'.$doc;
	$template  = $docsPath;
	$path_Doc = 'tmp/' . $newdoc;
	
	try {
		
		$content_XML = '';
   
		copy( $template , $path_Doc );
		
		$zip_val = new ZipArchive;
		if($zip_val->open( $path_Doc ) == true) {
			$key_XML     = 'word/document.xml';
			$content_XML = $zip_val->getFromName($key_XML);
		}
    
    } catch (Exception $exc) {            
     
	    return "Error creating the Word Document";
		
    }	

    preg_match_all("/\{(.*?)\}/", $content_XML, $matches);
    $blocks = $matches[1];	

    for ($k = 0; $k < count($aDataSearch); $k++) {
		
        foreach ($blocks as $block) {

            $processedBlock = processBlock("{" . $block . "}");
	    if( strpos( $processedBlock , '<w:t>{'.$aDataSearch[ $k ].'}</w:t>' ) !== false ) {  					       
                $content_XML = str_replace("{" . $block . "}", $processedBlock, $content_XML);				
                $content_XML = str_replace('<w:t>{'.$aDataSearch[$k].'}</w:t>', '<w:t>'.$aDataChange[$k].'</w:t>', $content_XML);			
	    }	
		
        }
    }
	
    $zip_val->addFromString($key_XML, $content_XML);
    $zip_val->close();			

    return( $path_Doc ); 		
	
}	


/*
*/
function processBlock($block) {
    $text = "";
	
	foreach (explode("<w:t>", $block) as $part) {
		$text .= strip_tags(trim($part));
    }
	
    return "<w:t>" . $text . "</w:t>";
}

/*
*/
function generatePrefix() {  
    $numbers     = (((date('Ymd') / 12) * 24) + mt_rand(800, 9999));
    $numbers    .= 123456789;
    $characters  = $numbers . 'abcdefghjmnoprstvxz';
    $ret         = substr(str_shuffle($characters), 0, 6);     
    return $ret;      
}   

/*
*/
function removeSymbolsStr($str) {
    return( preg_replace("/[^a-zA-Z0-9]/", "", $str));
}
?>
