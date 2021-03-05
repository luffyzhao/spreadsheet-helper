<?php


namespace LSpreadsheet\Spreadsheet\Merge;


use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class WorkSheetIterator
{
    /**
     * @var Worksheet
     * @author luffyzhao@vip.126.com
     */
    private $worksheet;

    /**
     * WorkSheetIterator constructor.
     * @param Worksheet $worksheet
     * @author luffyzhao@vip.126.com
     */
    public function __construct(Worksheet $worksheet)
    {
        $this->worksheet = $worksheet;
    }

    /**
     * @param string $only
     * @param int $minRow
     * @return MergeIterator
     * @author luffyzhao@vip.126.com
     */
    public function getMergeIterator($only = 'A', $minRow = 1)
    {
        return new MergeIterator($this->worksheet, $only, $minRow);
    }
}
