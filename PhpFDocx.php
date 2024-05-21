<?php
/*
 * PhpFDocx 2.0.0
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

    $content_XML = change01( $content_XML ,$aDataSearch , $aDataChange );
    $content_XML = change02( $content_XML ,$aDataSearch , $aDataChange );	
    $content_XML = change03( $content_XML ,$aDataSearch , $aDataChange );		
	
    $zip_val->addFromString($key_XML, $content_XML);
    $zip_val->close();			

    return( $path_Doc ); 		
	
}	

/*
*/
function change01( $content_XML ,$aDataSearch , $aDataChange ) {
	
    for ($k = 0; $k < count($aDataSearch); $k++) {
        $content_XML = str_replace('<w:t>{'.$aDataSearch[$k].'}</w:t>', '<w:t>'.$aDataChange[$k].'</w:t>', $content_XML);
        $content_XML = str_replace('{'.$aDataSearch[$k].'}', $aDataChange[$k], $content_XML);
    }
	
    return $content_XML;
}

/*
*/
function change02( $content_XML ,$aDataSearch , $aDataChange ) {	
	$posIni  = 0;
	$ct      = 0;
	$t       = 0;
	$cBlock  = '';

	for( $i = 0 ; $i <= strlen( $content_XML ) ; $i++ ) {
		
		$char = substr( $content_XML , $i , 1 ); 
		
		if( $char == '{' AND $t == 0 ) {
			$iniKey = $i; 
			$t = 1;		
		}	

		if( $char == '}' AND $t == 1 ) {
			
			$endKey    = $i; 
			$cBlock    = substr( $content_XML , $iniKey , ( $endKey - $iniKey ) + 1 );
			$t         = 0; 	
			$cBlockTmp = $cBlock;
			
			for( $k = 0 ; $k < count( $aDataSearch ) ; $k++ ) {
				
				if( strpos( $cBlockTmp , '<w:t>'.$aDataSearch[ $k ].'</w:t>' ) !== false ) {
					$cBlockTmp   = str_replace( '<w:t>'.$aDataSearch[ $k ].'</w:t>' , '<w:t>'.$aDataChange[ $k ].'</w:t>'  , $cBlockTmp );
					$cBlockTmp   = str_replace( '{' , '' , $cBlockTmp );				
					$cBlockTmp   = str_replace( '}' , '' , $cBlockTmp );
					$content_XML = str_replace( $cBlock , $cBlockTmp , $content_XML );	
				}			
				
			}		
			
		}

		$ct++;  
		$posIni++;	
		
	}	

    return $content_XML;
	
}	

/*
*/
function change03( $content_XML ,$aDataSearch , $aDataChange ) {
         $posIni = 0;
	$ct     = 0;
	$t      = 0;
	$cBlock = '';

	for( $i = 0 ; $i <= strlen( $content_XML ) ; $i++ ) {
		
		$char = substr( $content_XML , $i , 1 ); 
		
		if( $char == '{' AND $t == 0 ) {
			$chaveIni = $i; 
			$t = 1;		
		}	

		if( $char == '}' AND $t == 1 ) {
			
			$chaveEnd  = $i;  
			$cBlock    = substr( $content_XML , $chaveIni , ( $chaveEnd - $chaveIni ) + 1 );
			$t         = 0;
			$lChange   = false;
			$cBlockTmp = $cBlock;
			
			for( $k = 0 ; $k < count( $aDataSearch ) ; $k++ ) {
				
				if( strpos( $cBlockTmp , '<w:t>{'.$aDataSearch[ $k ] ) !== false ) {
					$cBlockTmp   = str_replace( '<w:t>{'.$aDataSearch[ $k ] , '<w:t>'.$aDataChange[ $k ]  , $cBlockTmp );				
					$cBlockTmp   = str_replace( '{' , '' , $cBlockTmp );				
					$cBlockTmp   = str_replace( '}' , '' , $cBlockTmp );
					$content_XML = str_replace( $cBlock , $cBlockTmp , $content_XML );
					$lChange     = true;					
				}			
				if( !$lChange AND strpos( $cBlockTmp , $aDataSearch[ $k ] ) !== false ) {
					$cBlockTmp   = str_replace( '<w:t>'.$aDataSearch[ $k ].'<w:t>}' , '<w:t>'.$aDataChange[ $k ].'</w:t>'  , $cBlockTmp );						
					$cBlockTmp   = str_replace( '<w:t>'.$aDataSearch[ $k ].'}' , '<w:t>'.$aDataChange[ $k ]  , $cBlockTmp );				
					$cBlockTmp   = str_replace( '{'.$aDataSearch[ $k ].'</w:t>' , $aDataChange[ $k ].'</w:t>'  , $cBlockTmp ); //ÚLTIMO 21/05 - 11:05
					$cBlockTmp   = str_replace( '{' , '' , $cBlockTmp );				
					$cBlockTmp   = str_replace( '}' , '' , $cBlockTmp );
					$content_XML = str_replace( $cBlock , $cBlockTmp , $content_XML );
					$lChange     = true;					
				}		
				
			}
			
			if( !$lChange )	{

				$aContent2 = getContents( $cBlockTmp .'</w:t>' , '<w:t>' , '</w:t>' ); 												
				$cNewVar   = '';					
				$aResto    = getContents( $cBlockTmp , '{' , '</w:t>' ); 
				for( $q = 0 ; $q < count( $aResto ) ; $q++ ) {					
					$cNewVar   .= $aResto[ $q ];
					$cBlockTmp = str_replace(  '<w:t>'.$aResto[ $q ].'</w:t>' , '<w:t></w:t>' , $cBlockTmp );				    	
				}
				
				for( $q = 0 ; $q < count( $aContent2 ) ; $q++ ) {					
					$cNewVar   .= $aContent2[ $q ];
					$cBlockTmp = str_replace(  '<w:t>'.$aContent2[ $q ].'</w:t>' , '<w:t></w:t>' , $cBlockTmp );				    	
				}

				$cNewVar = str_replace( '{' , '' , $cNewVar );				
				$cNewVar = str_replace( '}' , '' , $cNewVar );	
				
				if( !empty( $cNewVar ) ) {	
				
					$pont = array_search($cNewVar,$aDataSearch);
					if( $pont > 0 ) {
						$cDataChange = $aDataChange[$pont];							
						
						$aContentTmp = getContents( $cBlockTmp , '{' , '</w:t>' ); 														
						for( $q = 0 ; $q < count( $aContentTmp ) ; $q++ ) {					
							$cBlockTmp = str_replace( '{'.$aContentTmp[ $q ] , '' , $cBlockTmp );							
						}
						
						$cBlockTmp   = str_replace( '{' , '' , $cBlockTmp );				
						$cBlockTmp   = str_replace( '}' , '' , $cBlockTmp );				
						$cBlockTmp   = str_replace( '<w:t></w:t>' , '' , $cBlockTmp );
						
						$aContentTmp = getContents( $cBlockTmp.'</w:t>' , '<w:t>' , '</w:t>' ); 
						for( $z = 0 ; $z < count( $aContentTmp ) ; $z++ ) {
					             $cBlockTmp   = str_replace( '<w:t>'.$aContentTmp[ $z ] , '<w:t>' , $cBlockTmp );							
						}
						
						$content_XML = str_replace( $cBlock , $cBlockTmp . '<w:t>'.$cDataChange.'</w:t>'  , $content_XML ); 																										
						$lChange     = true;				
                                         }
					
				   }									
			  }							
		}

		$ct++;  
		$posIni++;	
		
	}
	
    return $content_XML;
	
}	

/*
*/
function getContents($str, $startDelimiter, $endDelimiter) {
    $contents = array();
    $startDelimiterLength = strlen($startDelimiter);
    $endDelimiterLength = strlen($endDelimiter);
    $startFrom = $contentStart = $contentEnd = 0;
    while (false !== ($contentStart = strpos($str, $startDelimiter, $startFrom))) {
        $contentStart += $startDelimiterLength;	 
        $contentEnd = strpos($str, $endDelimiter, $contentStart);
        if (false === $contentEnd) {
           break;
        }         	
        $contents[] = substr($str, $contentStart, $contentEnd - $contentStart);
        $startFrom = $contentEnd + $endDelimiterLength;
    }

    return $contents;  
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
