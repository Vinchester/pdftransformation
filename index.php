<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
</head>
<body>
    <?php

    require_once ('vendor/autoload.php');

    use PhpOffice\PhpWord\TemplateProcessor;
    use PhpOffice\PhpWord\IOFactory;
    use PhpOffice\PhpWord\Settings;

    $name = htmlspecialchars($_POST['Name']);
    $file = 'names.docx';

    $phpword = new \PhpOffice\PhpWord\TemplateProcessor($file);
    
    $phpword->setValue('{name}','' . $name . '');
    
    $wordPdf = $phpword->saveAs('edited.docx');

    require_once 'vendor/autoload.php';

    Settings::setPdfRendererName(Settings::PDF_RENDERER_DOMPDF);
    
    Settings::setPdfRendererPath('.');

    $phpWord = IOFactory::load('edited.docx', 'Word2007');
    $phpWord->save('document.pdf', 'PDF');
    ?>
    <embed src="document.pdf" width="80%" height="800px" />
    <a href="index2.php">Back</a>
</body>
</html>