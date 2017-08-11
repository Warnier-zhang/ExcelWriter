<?php

namespace snippet\excel;

/**
 * Class ExcelWriter
 * @package snippet\excel
 */
class ExcelWriter
{
    /**
     * @var \PHPExcel
     */
    private $_excel;

    public function __construct()
    {
        $this->_excel = new \PHPExcel();
    }

    public function write($sheetName, $head = [], $body = [])
    {
        $sheet = $this->_excel->setActiveSheetIndex(0);
        if (empty($sheetName)) {
            $sheet->setTitle("未命名文档");
        } else {
            $sheet->setTitle($sheetName);
        }

        $rowNumber = 1;
        $rowNumber = $this->writeHead($sheet, $head, $rowNumber);
        $rowNumber = $this->writeBody($sheet, $head, $body, $rowNumber);
        return $this;
    }

    private function writeHead(\PHPExcel_Worksheet $sheet, $head = [], $rowNumber)
    {
        $rowIndex = $rowNumber;

        foreach ($head as $headRow) {
            $colIndex = 0;
            foreach ($headRow as $headCol) {
                do {
                    $cell = $sheet->getCellByColumnAndRow($colIndex, $rowIndex);
                    $colIndex++;
                } while ($cell->isInMergeRange());
                $colIndex = $colIndex - 1;
                $range = \PHPExcel_Cell::stringFromColumnIndex($colIndex) . $rowIndex;

                if (isset($headCol["title"])) {
                    $sheet->setCellValue($range, $headCol["title"]);
                }

                $colspan = isset($headCol["colspan"]) ? intval($headCol["colspan"]) : 1;
                $rowspan = isset($headCol["rowspan"]) ? intval($headCol["rowspan"]) : 1;
                if ($colspan || $rowspan) {
                    if ($colspan > 0) {
                        $colspan = $colspan - 1;
                    }
                    if ($rowspan > 0) {
                        $rowspan = $rowspan - 1;
                    }
                    $mergeRange = \PHPExcel_Cell::stringFromColumnIndex($colIndex + $colspan) . ($rowIndex + $rowspan);
                    $sheet->mergeCells($range . ':' . $mergeRange);

                    $colIndex += $colspan;
                }
                $colIndex++;
            }
            $rowIndex++;
        }
        return $rowIndex;
    }

    private function writeBody(\PHPExcel_Worksheet $sheet, $head = [], $body = [], $rowNumber)
    {
        if ($rowNumber == 1 || empty($body)) {
            return 1;
        }

        $colIndex = 0;

        $limit = 0;
        foreach ($head[0] as $headerCol) {
            if (isset($headerCol["colspan"])) {
                $limit += intval($headerCol["colspan"]);
            } else {
                $limit += 1;
            }
        }
        $offset = 0;
        $fieldMappers = $this->mapFields($head, $colIndex, $offset, $limit);

        $rows = $body;
        $rowIndex = intval($rowNumber);
        foreach ($rows as $row) {
            foreach ($row as $field => $value) {
                $colIndex = $fieldMappers[$field];
                $range = \PHPExcel_Cell::stringFromColumnIndex($colIndex) . $rowIndex;
                $sheet->setCellValue($range, $value);
            }
            $rowIndex++;
        }
        return $rowIndex;
    }

    private function mapFields($header = [], &$colIndex, $offset, $limit)
    {
        $fieldMappers = [];

        $colOffset = 0;
        $rowNumber = count($header);
        $colNumber = 0;
        foreach ($header[0] as $headerCol) {
            $colspan = isset($headerCol["colspan"]) ? intval($headerCol["colspan"]) : 1;
            $colNumber += $colspan;
            if ($colNumber <= $offset) {
                continue;
            } else if ($colNumber <= $limit) {
                $rowspan = isset($headerCol["rowspan"]) ? intval($headerCol["rowspan"]) : 1;
                if ($rowspan === $rowNumber) {
                    $fieldMappers[$headerCol["field"]] = $colIndex;
                    $colIndex++;
                } else {
                    $colLimit = $colOffset + $colspan;
                    $fieldMappers = array_merge($fieldMappers, $this->mapFields(array_slice($header, 1), $colIndex, $colOffset, $colLimit));
                    $colOffset += $colspan;
                }
            } else {
                break;
            }
        }
        return $fieldMappers;
    }

    public function output($fileName)
    {
        $fileName = iconv("UTF-8", "GBK", $fileName);

        header("Content-Type: application/vnd.ms-excel");
        header("Content-Disposition: attachment;filename=" . $fileName . ".xls");
        header("Cache-Control: max-age=0");

        $writer = \PHPExcel_IOFactory::createWriter($this->_excel, 'Excel5');
        $writer->save('php://output');
    }

    public function saveAs($fileName)
    {
        $fileName = iconv("UTF-8", "GBK", $fileName);
        $writer = \PHPExcel_IOFactory::createWriter($this->_excel, 'Excel5');
        $writer->save($fileName);
    }
}