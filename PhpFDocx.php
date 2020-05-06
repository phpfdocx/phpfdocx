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
function PhpFDocx( $doc , $aDataSearch , $aDataChange ) {
    
	$newdoc    = generatePrefix().'_'.$doc;
	$docsPath  = 'doc/'.$doc;
	$template  = $docsPath;
	$fileName  = $newdoc;
	$folterTmp = 'tmp/';
	$path_Doc = $folterTmp . '/' . $fileName;
	
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

    $aContent = getContents( $content_XML , '${', '}');			
    $aContentSeg = $aContent;							

    for( $x = 0 ; $x <= count( $aContent ) ; $x++ ) { 
		
        for( $k = 0 ; $k <= count( $aDataSearch ) ; $k++ ) { 		
			
            if( ( isset( $aContent[ $x ] ) ) and ( isset( $aDataSearch[ $k ] ) ) ) {
				    
                if( strpos( $aContent[ $x ] , '<w:t>'.$aDataSearch[ $k ].'</w:t>' ) !== false ) {
					$tmp =  $aContent[ $x ];
					$tmp =  str_replace(  $aDataSearch[ $k ] , $aDataChange[ $k ] , $tmp ) ;
					$content_XML = str_replace( $aContent[ $x ] , $tmp , $content_XML );
					$aContentSeg[ $x ] = 'T';
				}

                if( strpos( $content_XML , '${'.$aDataSearch[ $k ].'}' ) !== false ) { 							
                    $content_XML = str_replace( '${'.$aDataSearch[ $k ].'}' , $aDataChange[ $k ] , $content_XML );

                    for( $n = 0 ; $n <= count( $aContentSeg ) ; $n++ ) {
					
                        if( isset( $aContentSeg[ $n ] ) ) {								
                   
			                $compareStr = getOnlyLowercase( $aContentSeg[ $n ] );
		                    if( $compareStr == $aDataSearch[ $k ] )  {						
			                    $aContentSeg[ $n ] = 'T'; 	
			                } 
			            }
		            }					
	            }				

		        if( strpos( $content_XML ,  '<w:t>'. $aDataSearch[ $k ].'</w:t>' ) !== false ) { 							
			
			        $content_XML = str_replace( '<w:t>'. $aDataSearch[ $k ].'</w:t>' , '<w:t>'. $aDataChange[ $k ].'</w:t>' , $content_XML );					

		            for( $n = 0 ; $n <= count( $aContentSeg ) ; $n++ ) {   				
				
                        if( isset( $aContentSeg[ $n ] ) ) {
					
                            $compareStr = getOnlyLowercase( $aContentSeg[ $n ] );
		                    if( $compareStr == $aDataSearch[ $k ] )  {
			                    $aContentSeg[ $n ] = 'T'; 								 
			                } 
			            }
		            }					
		        }
	        }										
	    }				
    }			
		
    for( $x = 0 ; $x <= count( $aContentSeg ) ; $x++ ) {
        if( isset( $aContentSeg[ $x ] ) ) {
            if( $aContentSeg[ $x ] !== 'T' ) {	
	            $onlyLowerChars = getOnlyLowercase( $aContentSeg[ $x ] );
	            if( !empty($onlyLowerChars) ) {
		            $posArray = array_search( $onlyLowerChars , $aDataSearch );
		            $subst    = '';
		            if( $posArray > 0 ) {
		                $subst = $aDataChange[ $posArray ];
                    }					
                    $pos = strpos(  $aContentSeg[ $x ] , substr( $onlyLowerChars , 0 , 1 ) ); 				
		            $content_XML = str_replace( $aContentSeg[ $x ] , $subst , $content_XML );
	            }
	        }
        }			
    }	

	$newStr = '';
	for( $r = 0 ; $r <= count( $aContentSeg ) ; $r++ ) {
	    if( isset( $aContentSeg[ $r ] ) ) {
			
			if( $aContentSeg[ $r ] !== 'T' ) {	
			    
				$onlyLowerChars = getOnlyLowercase( $aContentSeg[ $r ] );
	            if( !empty($onlyLowerChars) ) {
		            $posArray = array_search( $onlyLowerChars , $aDataSearch );		        
					$subst    = '';
		            if( $posArray > 0 ) {
						$subst = $aDataChange[ $posArray ];
					} 
				}			
			
			    $newStr = oBorgSpecialSearchAndReplace( $aContentSeg[ $r ] , $subst );
				$content_XML = str_replace( $aContentSeg[ $r ] , $subst , $content_XML );
				
			}
		}
    }			
	
    $content_XML = str_replace( '${'            , '' , $content_XML);		
    $content_XML = str_replace( '<w:t>$</w:t>'  , '' , $content_XML);		
    $content_XML = str_replace( '<w:t>${</w:t>' , '' , $content_XML);		
    $content_XML = str_replace( '<w:t>}</w:t>'  , '' , $content_XML);
    $content_XML = str_replace( '<w:t>{</w:t>'  , '' , $content_XML);	
    $content_XML = str_replace( '{</w:t>'       , '</w:t>' , $content_XML);	
    $content_XML = str_replace( '}</w:t>'       , '</w:t>' , $content_XML);	

    $zip_val->addFromString($key_XML, $content_XML);
    $zip_val->close();			

    return( $folterTmp . $fileName ); 		
	
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
function getOnlyLowercase( $str ) {
    $ret = '';

    $str = str_replace( 'w:cs'           , '' , $str );
    $str = str_replace( 'w:rFonts'       , '' , $str );
    $str = str_replace( '<w:t>'          , '' , $str );
    $str = str_replace( '</w:t>'         , '' , $str );	
    $str = str_replace( '<w:r>'          , '' , $str );		
    $str = str_replace( '</w:r>'         , '' , $str );
    $str = str_replace( '<w:r w:rsidR='  , '' , $str );	
    $str = str_replace( '<w:rPr>'        , '' , $str );		
    $str = str_replace( '<w:b/>'         , '' , $str );		
    $str = str_replace( '<w:bCs/>'       , '' , $str );
    $str = str_replace( '</w:rPr>'       , '' , $str );
    $str = str_replace( '<w:r w:rsidRPr' , '' , $str );
    $str = str_replace( '<w:r w:rsidR'   , '' , $str );
    $str = str_replace( 'w:rsidRPr'      , '' , $str );	
		
    $ct = getContents( $str , '<' , '>' );
	for( $x = 0 ; $x <= count( $ct ) ; $x++ ) { 
        if( isset( $ct[ $x ] ) ) { 	
	        $str = str_replace( $ct[ $x ] , '' , $str );				  
		}
	}	  
    $str = str_replace( '<' , '' , $str );					  
    $str = str_replace( '>' , '' , $str );				  
    $str = str_replace( '/' , '' , $str );	    
	
	$n_caracteres = strlen( $str );
	
	for( $i=0; $i < $n_caracteres ; $i++ ){
		if( ctype_lower($str[$i]) ) {
			$ret .= $str[$i];
		}	
	}    
	
    return $ret;
}

/*
*/
function oBorgSpecialSearchAndReplace( $str , $subst ) {
	
    $firstPos = strpos ( $str , '<' );
	$LastPos  = strrpos( $str , '>' );
	$Part     = substr( $str , $firstPos , $LastPos-$firstPos );
	$ct       = getContents( $str , '<' , '>' );
	
    for( $x = 0 ; $x <= count( $ct ) ; $x++ ) { 			  
	    if( isset(  $ct[ $x ] ) ) {
            $str = str_replace( $ct[ $x ] , '' , $str );				  
		}
    }	  
			  
    $Part = str_replace( '<w:t>'  , '' , $Part );
    $Part = str_replace( '</w:t>' , '' , $Part );
    $Part = str_replace( '<w:t'   , '' , $Part );		  
    $newStr = $subst.'</w:t>'.$Part;	
	
    return $newStr; 
	
}

/*
*/
function generatePrefix() {    	
    $chars = 'abcdxyswz0123456789';
    $max = strlen($chars) - 1;
    $prefix = null;
    for($i=0;$i < 4; $i++) {
        $prefix .= $chars{mt_rand(0,$max)};
    }
    return rand(10,99).$prefix;          
}   
?>	

	
