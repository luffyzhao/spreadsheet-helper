<?php


namespace LSpreadsheet\Spreadsheet\Merge;


use Exception;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

class Row
{
    /**
     * @var array
     * @author luffyzhao@vip.126.com
     */
    private $item;

    /**
     * Row constructor.
     * @param array $item
     * @author luffyzhao@vip.126.com
     */
    public function __construct(array $item)
    {
        $this->item = $item;
    }

    /**
     * @param $startCol
     * @param $endCol
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws Exception
     * @author luffyzhao@vip.126.com
     */
    public function getHeader($startCol, $endCol)
    {
        $startColumnIndex = Coordinate::columnIndexFromString($startCol);
        $endColumnIndex = Coordinate::columnIndexFromString($endCol);

        $lastItem = [];
        foreach ($this->item as $key => $item) {
            for ($i = $startColumnIndex; $i <= $endColumnIndex; $i++) {
                if (!empty($lastItem) && isset($lastItem[$i]) && $lastItem[$i] !== $item[$i]) {
                    throw new Exception(sprintf("第%d行数据有问题:第%s列数据和前面数据对不上！", $key, Coordinate::stringFromColumnIndex($i)));
                }
                $lastItem[$i] = $item[$i];
            }
        }
        return $lastItem;
    }

    /**
     * @param $startCol
     * @param $endCol
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @author luffyzhao@vip.126.com
     */
    public function getDetails($startCol, $endCol)
    {
        $startColumnIndex = Coordinate::columnIndexFromString($startCol);
        $endColumnIndex = Coordinate::columnIndexFromString($endCol);

        $array = [];
        foreach ($this->item as $index => $item) {
            for ($i = $startColumnIndex; $i <= $endColumnIndex; $i++) {
                $array[$index][$i] = $item[$i];
            }
        }
        return $array;
    }

    /**
     * @return array
     * @author luffyzhao@vip.126.com
     */
    public function getRow(){
        return $this->item;
    }
}
