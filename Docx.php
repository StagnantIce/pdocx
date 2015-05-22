<?php

/**
 * @author StagnantIce (tria-aa@mail.ru)
 * @url https://github.com/StagnantIce/pdocx
 */

define('DOCX_BASE_PATH', dirname(__FILE__) . '/');

/**
 * The DOCx class.
 */
class DOCx {
  private $_filename     = NULL;
  private $_filepath     = NULL;
  private $_tempFilename = NULL;
  private $_tempFilepath = NULL;
  private $_zipArchive   = NULL;

  private $_documents    = array();
  private $_footers      = array();
  private $_headers      = array();

  /**
   * Initiate a new instance of the DOCx class.
   *
   * @param string $filepath
   *   The .docx file to load.
   */
  public function __construct($filepath = NULL) {
    if ($filepath !== NULL) {
      $this->load($filepath);
    }
  }

  /**
   * Cleans out the excessive '${' and '}' tags in the document.
   */
  public function cleanTagVars() {
    $this->setValueModifier($this->_documents, '${', '');
    $this->setValueModifier($this->_documents, '}', '');

    $this->setValueModifier($this->_footers, '${', '');
    $this->setValueModifier($this->_footers, '}', '');

    $this->setValueModifier($this->_headers, '${', '');
    $this->setValueModifier($this->_headers, '}', '');
  }

  /**
   * Resets the class for a new run.
   */
  public function close() {
    $this->_filename = NULL;
    $this->_filepath = NULL;
    $this->_tempFilename = NULL;
    $this->_tempFilepath = NULL;
    $this->_zipArchive = NULL;

    $this->_documents = array();
    $this->_footers = array();
    $this->_headers = array();
  }

  const EMU_PER_PT = 12700;
  const PX_PER_PT = 0.75;

  public $images = array();
  public function setImage($name, $path) {
    list( $width, $height ) = getimagesize($path);
    list( $w, $h ) = array($width * self::EMU_PER_PT * self::PX_PER_PT, $height * self::EMU_PER_PT * self::PX_PER_PT);
    foreach($this->_documents as &$document) {
        preg_match_all('/<wp:extent cx="(\d+)" cy="(\d+)"\/>/', $document['content'], $matches);
        if ($document['localName'] == 'word/document.xml') {
          if ($matches && isset($matches[2][count($this->images)]) && isset($matches[1][count($this->images)]))  {
            list($x, $y) = array( $matches[1][count($this->images)], $matches[2][count($this->images)]);
            $document['content'] = str_replace('<wp:extent cx="'.$x.'" cy="'.$y.'"/>', '<wp:extent cx="'.$w.'" cy="'.$h.'"/>', $document['content']);
            $document['content'] = str_replace('<a:ext cx="'.$x.'" cy="'.$y.'"/>', '<a:ext cx="'.$w.'" cy="'.$h.'"/>', $document['content']);
          } else {
            throw new Exception("Image not found in Docx");
          }
        }
    }
    $this->images[] = $path;
    $this->_zipArchive->addFile( $path, 'word/media/' . $name);
  }

  /**
   * Loads a .docx file into memory and unzips it.
   *
   * @param string $filename
   *   The .docx file to load.
   */
  public function load($filepath) {
    $this->_filename = (strpos($filepath, '/') !== FALSE ? substr($filepath, strrpos($filepath, '/') + 1) : $filepath);
    $this->_filepath = $filepath;

    $this->_tempFilename = '.' . time() . '.temp.docx';
    $this->_tempFilepath = DOCX_BASE_PATH . $this->_tempFilename;

    copy(
      $this->_filepath,
      $this->_tempFilepath
    );

    $this->_zipArchive = new ZipArchive();
    $this->_zipArchive->open($this->_tempFilepath);

    $iterators = array(
      '',
    );

    for ($i = 0; $i < 100; $i++) {
      $iterators[] = $i;
    }

    for ($i = 0; $i < count($iterators); $i++) {
      $xmlDocument = $this->_zipArchive->getFromName('word/document' . $iterators[$i] . '.xml');
      $xmlFooter = $this->_zipArchive->getFromName('word/footer' . $iterators[$i] . '.xml');
      $xmlHeader = $this->_zipArchive->getFromName('word/header' . $iterators[$i] . '.xml');

      if (is_string($xmlDocument) && !empty($xmlDocument)) {
        $this->_documents[] = array(
          'content'   => $xmlDocument,
          'localName' => 'word/document' . $iterators[$i] . '.xml'
        );
      }

      if (is_string($xmlFooter) && !empty($xmlFooter)) {
        $this->_footers[] = array(
          'content'   => $xmlFooter,
          'localName' => 'word/footer' . $iterators[$i] . '.xml'
        );
      }

      if (is_string($xmlHeader) && !empty($xmlHeader)) {
        $this->_headers[] = array(
          'content'   => $xmlHeader,
          'localName' => 'word/headers' . $iterators[$i] . '.xml'
        );
      }
    }
  }

  /**
   * Saves the temporary buffer to disk.
   *
   * @param string $filepath
   *   The file to save to. If none is given the temp file is used.
   *
   * @return string
   *   The filepath of the save file.
   */
  public function save($filepath = NULL) {
    if (count($this->_documents)) {
      foreach ($this->_documents as $document) {
        $this->_zipArchive->addFromString($document['localName'], $document['content']);
      }
    }

    if (count($this->_footers)) {
      foreach ($this->_footers as $footer) {
        $this->_zipArchive->addFromString($footer['localName'], $footer['content']);
      }
    }

    if (count($this->_headers)) {
      foreach ($this->_headers as $header) {
        $this->_zipArchive->addFromString($header['localName'], $header['content']);
      }
    }

    $this->_zipArchive->close();

    if ($filepath !== NULL) {
      copy(
        $this->_tempFilepath,
        $filepath
      );

      return $filepath;
    }
    else {
      return $this->_tempFilepath;
    }
  }

  /**
   * Does a global search and replace with the given values.
   *
   * @param string $search
   *   The tag to search for, represented as ${TAGNAME} in the file.
   * @param string $replace
   *   The text to replace it with.
   */
  public function setValue($search, $replace) {
      $this->setValueDocument($search, $replace);
      $this->setValueFooter($search, $replace);
      $this->setValueHeader($search, $replace);
  }

  /**
   * Does a global search and replace with an array of values.
   *
   * @param array $values
   *   A keyed array with search and replaces values.
   */
  public function setValues($values, $add="") {
    if (is_array($values) &&
      count($values)) {
      foreach ($values as $key => $value) {
        if (is_array($value)) {
          if ($key === "table") {
            $this->setValue(rtrim($add, '_'), $value);
          } else {
            $this->setValues($value, (string) $add . (string)$key . '_');
          }
        } else {
          $this->setValue((string)$add . (string)$key, (string)$value);
        }
      }
    }
  }

  /**
   * Does a search and replace in the 'document' part of the file.
   *
   * @param string $search
   *   The tag to search for, represented as ${TAGNAME} in the file.
   * @param string $replace
   *   The text to replace it with.
   */
  public function setValueDocument($search, $replace) {
      $this->setValueModifier($this->_documents, $search, $replace);
  }

  /**
   * Does a search and replace in the 'footer' part of the file.
   *
   * @param string $search
   *   The tag to search for, represented as ${TAGNAME} in the file.
   * @param string $replace
   *   The text to replace it with.
   */
  public function setValueFooter($search, $replace) {
      $this->setValueModifier($this->_footers, $search, $replace);
  }

  /**
   * Does a search and replace in the 'header' part of the file.
   *
   * @param string $search
   *   The tag to search for, represented as ${TAGNAME} in the file.
   * @param string $replace
   *   The text to replace it with.
   */
  public function setValueHeader($search, $replace) {
      $this->setValueModifier($this->_headers, $search, $replace);
  }

  /**
   * Does a search and replace in the given array.
   *
   * @param array $array
   *   The DOCx array to modify.
   * @param string $search
   *   The tag to search for, represented as ${TAGNAME} in the file.
   * @param string $replace
   *   The text to replace it with.
   * @param bool $validateSearchVar
   *   Wether or not to validate the search tag as ${TAGNAME}.
   */
  private function setValueModifier(&$array, $search, $replace, $validateSearchVar = TRUE) {
    $rowSearch = '${'.$search.'}';
    if (count($array)) {
      for ($i = 0; $i < count($array); $i++) {
        //prepare
        $array[$i]['content'] = preg_replace('/\$\{<[^\}]*?>([a-z_0-9]+)/', '${$1', $array[$i]['content']);
        $array[$i]['content'] = preg_replace('/\$\{([a-z_0-9]+)<[^\}]*?>([a-z_0-9]+)/', '${$1$2', $array[$i]['content']);
        $array[$i]['content'] = preg_replace('/\$\{([a-z_0-9]+)<[^\}]*?>([a-z_0-9]+)/', '${$1$2', $array[$i]['content']);
        $array[$i]['content'] = preg_replace('/\$\{([a-z_0-9]+)<[^\}]*?>(\})/', '${$1$2', $array[$i]['content']);
        if (is_array($replace) && count($replace) > 0 && strpos($array[$i]['content'], $rowSearch) !== false) {
          preg_match_all('/<\/w:tr>(<w:tr .+?<\/w:tr>)<\/w:tbl>/', $array[$i]['content'], $matches);
          foreach($matches[1] as $match) {
            $match = explode('<w:tr ', $match);
            $match = count($match) > 1 ? '<w:tr ' . $match[count($match) - 1] : '<w:tr ' . $match[0];
            if (strpos($match, $rowSearch) !== false) {
            $rows = array();
            foreach($replace as $row) {
              $res = $match;
              foreach($row as $key => $val) {
                $res = str_replace('${'.$key.'}', $val, $res);
              }
              $rows[] = $res;
            }
            $array[$i]['content'] = preg_replace('/('.str_replace(array('/', '$'), array('\/','\$'), $match).')<\/w:tbl>/', implode(',', $rows).'</w:tbl>', $array[$i]['content']);
          }
        }
       } else {
        $array[$i]['content'] = str_replace($rowSearch, $replace, $array[$i]['content']);
       }
    }
  }

 }
}
