<?php

namespace snippet\excel;

/**
 * Class ExcelReader
 * @package snippet\excel
 */
class ExcelReader
{
    /**
     * @var \PHPExcel
     */
    private $_excel;

    public function __construct($fileName)
    {
        $reader = \PHPExcel_IOFactory::createReader("Excel5");
        $this->_excel = $reader->load($fileName);
    }

    public function read($sheetName, $fieldMappers = [], $rowNumber)
    {
        $rows = [];

        $sheet = $this->_excel->getSheetByName($sheetName);

        $size = $sheet->getHighestRow();
        for ($rowIndex = intval($rowNumber); $rowIndex <= $size; $rowIndex++) {
            $row = [];
            foreach ($fieldMappers as $colIndex => $field) {
                $value = $sheet->getCell($colIndex . $rowIndex)->getValue();
                if (!empty($value)) {
                    $row[$field] = $value;
                }
            }
            if (!empty($row)) {
                $rows[] = $row;
            }
        }
        return $rows;
    }
}