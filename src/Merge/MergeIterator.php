<?php


namespace LSpreadsheet\Spreadsheet\Merge;


use Iterator;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class MergeIterator implements Iterator
{
    private $only;
    /**
     * @var Worksheet
     * @author luffyzhao@vip.126.com
     */
    private $worksheet;
    /**
     * @var int
     * @author luffyzhao@vip.126.com
     */
    private $currentRow;

    /**
     * Start position.
     *
     * @var int
     */
    private $startRow = 1;

    /**
     * End position.
     *
     * @var int
     */
    public $endRow = 1;

    /**
     * MergeIterator constructor.
     * @param Worksheet $worksheet
     * @param string $only
     * @param int $startRow
     * @author luffyzhao@vip.126.com
     */
    public function __construct(Worksheet $worksheet, $only = 'A', $startRow = 1)
    {
        $this->only = $only;
        $this->worksheet = $worksheet;
        $this->currentRow = $startRow;
        $this->startRow = $startRow;

        $this->endRow = $worksheet->getHighestRow();
    }


    /**
     * Return the current element
     * @link https://php.net/manual/en/iterator.current.php
     * @return mixed Can return any type.
     * @since 5.0.0
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function current()
    {
        return new RowHandle($this, $this->worksheet, $this->currentRow, $this->only);
    }

    /**
     * Move forward to next element
     * @link https://php.net/manual/en/iterator.next.php
     * @return void Any returned value is ignored.
     * @since 5.0.0
     */
    public function next()
    {
        ++$this->currentRow;
    }

    /**
     * Return the key of the current element
     * @link https://php.net/manual/en/iterator.key.php
     * @return mixed scalar on success, or null on failure.
     * @since 5.0.0
     */
    public function key()
    {
        return $this->currentRow;
    }

    /**
     * Checks if current position is valid
     * @link https://php.net/manual/en/iterator.valid.php
     * @return boolean The return value will be casted to boolean and then evaluated.
     * Returns true on success or false on failure.
     * @since 5.0.0
     */
    public function valid()
    {
        return $this->currentRow <= $this->endRow && $this->currentRow >= $this->startRow;
    }

    /**
     * Rewind the Iterator to the first element
     * @link https://php.net/manual/en/iterator.rewind.php
     * @return void Any returned value is ignored.
     * @since 5.0.0
     */
    public function rewind()
    {
        $this->currentRow = $this->startRow;
    }
}
