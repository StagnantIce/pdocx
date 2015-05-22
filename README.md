# pdocx

<h1> PHP Docx file generation </h1>

This class help you replace some text, images and tables in docx files for MS Office without any troubles.

1) Usage 
```php
        $data = array('var' => 'Hello', 'table' => array('t1_name' => 'Petr', 't1_age' => 10), 'images' => array(1 => 'image1.png'))
	$docx = new DOCx($_SERVER['DOCUMENT_ROOT'] . '/template.docx');
	foreach($data['images'] as $i => $image) {
		$docx->setImage('image'.$i.'.png', $image);
	}
	$docx->setValues($data);
	$fileName = $docx->save();
	ob_clean();
	ob_get_clean();

	header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
	header('Content-Disposition: attachment; filename="' . $fileName .'"');
	header('Content-Transfer-Encoding: binary');
	header('Content-Length: ' . filesize($fileName));
	readfile($fileName);
	exit();
}
```

2) Notice

<p> data['images'] should have realy names of images from docx. You need rename file template.docx to template.zip and found in word/document.xml image names.</p>
