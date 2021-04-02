<?php


namespace LSpreadsheet\Spreadsheet\Merge;


use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class RowHandle
{
    /**
     * @var array
     * @author luffyzhao@vip.126.com
     */
    private $item;
    /**
     * @var MergeIterator
     * @author luffyzhao@vip.126.com
     */
    private $iterator;
    /**
     * @var Worksheet
     * @author luffyzhao@vip.126.com
     */
    private $worksheet;
    /**
     * @var int
     * @author luffyzhao@vip.126.com
     */
    private $startRow;
    /**
     * @var int
     * @author luffyzhao@vip.126.com
     */
    private $endRow;
    /**
     * @var string
     * @author luffyzhao@vip.126.com
     */
    private $only;

    /**
     * RowIterator constructor.
     * @param MergeIterator $iterator
     * @param Worksheet $worksheet
     * @param int $startRow
     * @param string $only
     * @throws Exception
     * @author luffyzhao@vip.126.com
     */
    public function __construct(MergeIterator $iterator, Worksheet $worksheet, int $startRow, $only = 'A')
    {
        $this->iterator = $iterator;
        $this->worksheet = $worksheet;
        $this->startRow = $startRow;
        $this->endRow = $startRow;
        $this->only = $only;
        $this->initHandle();
    }

    /**
     * @throws Exception
     * @author luffyzhao@vip.126.com
     */
    protected function initHandle()
    {
        $endRow = $this->startRow;
        $only = $this->worksheet->getCell($this->only . $endRow)->getValue();
        while (true) {
            ++$endRow;
            $endOnly = $this->worksheet->getCell($this->only . $endRow)->getValue();
            if ($endOnly !== $only) {
                break;
            }
            ++$this->endRow;
            $this->iterator->next();
        }
        $endOnly = $this->worksheet->getCell($this->only . $endRow)->getValue();
        if(empty($endOnly)){
            $this->iterator->endRow = --$endRow;
        }
    }

    /**
     * @return Row
     * @throws Exception
     * @author luffyzhao@vip.126.com
     */
    public function getRow(){
        $maxCol = $this->worksheet->getHighestColumn();
        $rowIterator = $this->worksheet->getRowIterator($this->startRow, $this->endRow);
        $rowArr = [];
        foreach ($rowIterator as $index => $item){
            foreach ($item->getCellIterator('A', $maxCol) as $cell){
                $rowArr[$index][ Coordinate::columnIndexFromString($cell->getColumn())] = $cell->getValue();
            }
        }
        return new Row($rowArr);
    }

}
