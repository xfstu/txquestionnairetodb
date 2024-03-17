<?php

namespace xfstu\tools;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

/**
 * @method export(array $field, array $data) 简单封装导出xlsx
 * @method getFileArray(string $fileName, int $start = 2, string $getHighestColumn = 'auto') 简单封装读取xlsx
 */
class ToolsXlsx
{
    public $Spreadsheet;

    public function __construct()
    {
        $this->Spreadsheet = new Spreadsheet();
    }

    /**
     * 导出二维数组到xlsx
     * @param array $field 字段名称
     * @param array $data 
     * @param callable $callback 回调
     */
    public function export($field, $data, $callback = null)
    {
        $Spreadsheet = $this->Spreadsheet;
        $sheet = $Spreadsheet->getActiveSheet();

        $fieldABC = array_slice(range('A', 'Z'), 0, count($field));
        $sheet->getStyle('A1:' . end($fieldABC) . '' . (count($data) + 1))->getFont()->setName('宋体');
        foreach ($field as $k => $v) {
            $sheet->setCellValue($fieldABC[$k] . '1', $v);
        }
        for ($i = 0; $i < count($data); $i++) {
            if ($callback) {
                $row = $callback($data[$i]);
            } else {
                $row = $data[$i];
            }
            for ($j = 0; $j < count($field); $j++) {
                // dump($fieldABC[$j] . ($i + 2), $row[$field[$j]]);
                $sheet->setCellValue($fieldABC[$j] . ($i + 2), $row[$field[$j]])->getColumnDimension($fieldABC[$j])->setAutoSize(TRUE);
            }
        }
        $writer = new Xlsx($Spreadsheet);
        return $writer->save('php://output');
    }

    /**
     * 读取xlsx
     * @param string $fileName
     * @param int $start
     * @param string $getHighestColumn
     * @return array
     */
    public function getFileArray($fileName, $start = 2, $callback = null, $getHighestColumn = 'auto')
    {
        $spreadsheet = IOFactory::load($fileName);
        $worksheet = $spreadsheet->getActiveSheet();
        $highestRow = $worksheet->getHighestRow();
        $highestColumn = $getHighestColumn == 'auto' ? $worksheet->getHighestColumn() : $getHighestColumn;
        $data = [];
        for ($start; $start <= $highestRow; ++$start) {
            $rows = $worksheet->rangeToArray('A' . $start . ':' . $highestColumn . $start, null);
            if ($callback) {
                $data[] = $callback($rows[0]);
            } else {
                $data[] = $rows[0];
            }
        }
        return $data;
    }
}
