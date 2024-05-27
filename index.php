<?php
$dirPath = "doc";
    
// Method 2: Using glob()
echo "--------------------------------<br>";
echo "Documents test <br>";
echo "--------------------------------<br>";
$files = glob($dirPath . "/*.docx");
foreach ($files as $file) {
    if (is_file($file)) {
        echo '<a href="demo.php?f=' . basename($file) . '"</a>'.basename($file).'<br>';
    }
}

?>