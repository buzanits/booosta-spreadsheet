<?php
namespace booosta\spreadsheet;

use \booosta\Framework as b;
b::init_module('spreadsheet');

class Spreadsheet extends \booosta\base\Module
{ 
  use moduletrait_spreadsheet;

  protected $spreadsheet;
  protected $filename;
  protected $data;
  protected $mapping;
  protected $extract_hyperlinks = false;
  protected $autoresize = true;


  public function __construct($filename = null)
  {
    parent::__construct();

    $this->filename = $filename;

    if(is_readable($filename)) $this->spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filename);
    else $this->spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
  }

  public function set_mapping($mapping) { $this->mapping = $mapping; }

  public function __call($name, $arguments) { return $this->spreadsheet->$name(...$arguments); }
  public function get_writer_xlsx() { return new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($this->spreadsheet); }
  public function get_writer_ods() { return new \PhpOffice\PhpSpreadsheet\Writer\Ods($this->spreadsheet); }
  public function extract_hyperlinks($flag = true) { $this->extract_hyperlinks = $flag; }
  public function set_autoresize($flag = true) { $this->autoresize = $flag; }

  public function save($filename = null, $format = null)
  {
    if($filename === null) $filename = $this->filename;

    if($format === null):
      $ext = array_pop(explode('.', $filename));
      $format = $ext ?: 'xlsx';  // xlsx is default
    endif;

    $writer = $this->{"get_writer_$format"}();
    $writer->save($filename);
  }

  public function set_data($data)
  {
    if(!is_array($data)) $data = [$data];
    $this->spreadsheet->getActiveSheet()->fromArray($data);
    if($this->autoresize) $this->autoresize();
  }

  public function get_indexed_data($convert2utf8 = false, $header = true)
  {
    $sheet = $this->spreadsheet->getActiveSheet();
    $numRows = $sheet->getHighestRow();
    $numCols = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($sheet->getHighestDataColumn());
    #\booosta\debug("rows: $numRows, cols: $numCols");
    $indexname = [];

    // read first line with index names if $header === true
    if($header):
      $indexfound = false;
      $h = 1;

      while(!$indexfound && $h <= $numRows):
        for($j = 1; $j <= $numCols; $j++):
          $indexname[$j] = $sheet->getCellByColumnAndRow($j, $h)->getValue();
          if($convert2utf8) $indexname[$j] = \booosta\Framework::to_utf8($indexname[$j]);
          if($indexname[$j] != '') $indexfound = true;
        endfor;
        $h++;
      endwhile;
    else:
      $h = 0;
    endif;

    // then read data
    $result = [];
    for($i = 0; $i <= $numRows - $h; $i++):
      for($j = 1; $j <= $numCols; $j++):
        if($indexname[$j] == '') $indexname[$j] = $j;
        #$result[$i][$j] = $sheet->getCellByColumnAndRow($j, $i + $h)->getValue();
        $result[$i][$indexname[$j]] = $sheet->getCellByColumnAndRow($j, $i + $h)->getValue();
        if($convert2utf8) $result[$i][$indexname[$j]] = \booosta\Framework::to_utf8($result[$i][$indexname[$j]]);

        if($this->extract_hyperlinks && $sheet->getCellByColumnAndRow($j, $i + $h)->hasHyperlink())
          $result[$i]["{$indexname[$j]}_hyperlink"] = $sheet->getCellByColumnAndRow($j, $i + $h)->getHyperlink()->getUrl();
      endfor;
    endfor;

    #\booosta\debug($result);
    return $result;
  }

  public function get_mapped_data($convert2utf8 = false, $header = true)
  {
    $data = $this->get_indexed_data($convert2utf8, $header);
    if(!is_array($this->mapping)) return $data;

    $result = [];
    foreach($data as $ridx=>$row):
      foreach($row as $cidx=>$col):
        $idx = $this->mapping[$cidx] ? $this->mapping[$cidx] : $cidx;
        $result[$ridx][$idx] = $data[$ridx][$cidx];
      endforeach;
    endforeach;

    return $result;
  }

  public function get_header() 
  {
    $sheet = $this->spreadsheet->getActiveSheet();
    $numRows = $sheet->getHighestRow();
    $numCols = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($sheet->getHighestDataColumn());
    $indexname = [];
    
    // read first line with index names if $header === true
    $indexfound = false;
    $h = 1;

    while(!$indexfound && $h <= $numRows):
      for($j = 1; $j <= $numCols; $j++):
        $indexname[$j] = $sheet->getCellByColumnAndRow($j, $h)->getValue();
        if($convert2utf8) $indexname[$j] = \booosta\Framework::to_utf8($indexname[$j]);
        if($indexname[$j] != '') $indexfound = true;
      endfor;
      $h++;
    endwhile;

    #\booosta\debug($indexname);
    return $indexname;
  }

  protected function autoresize()
  {
    $sheet = $this->spreadsheet->getActiveSheet();
    foreach ($sheet->getColumnIterator() as $column) $sheet->getColumnDimension($column->getColumnIndex())->setAutoSize(true);
  }
}
