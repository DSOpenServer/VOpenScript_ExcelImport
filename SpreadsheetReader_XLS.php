<?php

class SpreadsheetReader_XLS implements Iterator, Countable
{
    /**
     * @var array Options array
     */
    private array $Options = [];

    /**
     * @var Spreadsheet_Excel_Reader|false File handle
     */
    private mixed $Handle = false;

    private int  $Index   = 0;
    private bool $Error   = false;

    /**
     * @var array|false Sheet information
     */
    private array|false $Sheets = false;

    private array $SheetIndexes = [];

    /**
     * @var int Current sheet index
     */
    private int $CurrentSheet = 0;

    /**
     * @var array Content of the current row
     */
    private array $CurrentRow = [];

    /**
     * @var int Column count in the sheet
     */
    private int $ColumnCount = 0;

    /**
     * @var int Row count in the sheet
     */
    private int $RowCount = 0;

    /**
     * @var array Template to use for empty rows. Retrieved rows are merged
     *            with this so that empty cells are added, too.
     */
    private array $EmptyRow = [];

    /**
     * @param string     $Filepath Path to file
     * @param array|null $Options  Options
     *
     * @throws Exception
     */
    public function __construct(string $Filepath, ?array $Options = null)
    {
        if (!is_readable($Filepath)) {
            throw new Exception('SpreadsheetReader_XLS: File not readable (' . $Filepath . ')');
        }

        if (!class_exists('Spreadsheet_Excel_Reader')) {
            throw new Exception('SpreadsheetReader_XLS: Spreadsheet_Excel_Reader class not available');
        }

        $this->Handle = new Spreadsheet_Excel_Reader($Filepath, false, 'UTF-8');

        if (function_exists('mb_convert_encoding')) {
            $this->Handle->setUTFEncoder('mb');
        }

        if (empty($this->Handle->sheets)) {
            $this->Error = true;
            return;
        }

        $this->ChangeSheet(0);
    }

    /**
     * Destructor — releases the file handle.
     */
    public function __destruct()
    {
        unset($this->Handle);
    }

    /**
     * Retrieves an array with information about sheets in the current file.
     *
     * @return array List of sheets (key is sheet index, value is name)
     */
    public function Sheets(): array
    {
        if ($this->Sheets === false) {
            $this->Sheets       = [];
            $this->SheetIndexes = array_keys($this->Handle->sheets);

            foreach ($this->SheetIndexes as $SheetIndex) {
                $this->Sheets[] = $this->Handle->boundsheets[$SheetIndex]['name'];
            }
        }

        return $this->Sheets;
    }

    /**
     * Changes the current sheet in the file to another.
     *
     * @param int $Index Sheet index
     *
     * @return bool True if sheet was successfully changed, false otherwise.
     */
    public function ChangeSheet(int $Index): bool
    {
        $Sheets = $this->Sheets();

        if (isset($this->Sheets[$Index])) {
            $this->rewind();
            $this->CurrentSheet = $this->SheetIndexes[$Index];

            $this->ColumnCount = $this->Handle->sheets[$this->CurrentSheet]['numCols'];
            $this->RowCount    = $this->Handle->sheets[$this->CurrentSheet]['numRows'];

            // For the case when Spreadsheet_Excel_Reader doesn't have the row count set correctly.
            if (!$this->RowCount && count($this->Handle->sheets[$this->CurrentSheet]['cells'])) {
                end($this->Handle->sheets[$this->CurrentSheet]['cells']);
                $this->RowCount = (int)key($this->Handle->sheets[$this->CurrentSheet]['cells']);
            }

            $this->EmptyRow = $this->ColumnCount
                ? array_fill(1, $this->ColumnCount, '')
                : [];

            return true;
        }

        return false;
    }

    /**
     * Magic getter — exposes the $Error property publicly.
     *
     * @param string $Name Property name
     *
     * @return mixed
     */
    public function __get(string $Name): mixed
    {
        if ($Name === 'Error') {
            return $this->Error;
        }
        return null;
    }

    // -------------------------------------------------------------------------
    // Iterator interface methods
    // -------------------------------------------------------------------------

    /**
     * Rewind the Iterator to the first element.
     * Similar to the reset() function for arrays in PHP.
     */
    public function rewind(): void
    {
        $this->Index = 0;
    }

    /**
     * Return the current element.
     * Similar to the current() function for arrays in PHP.
     *
     * @return array Current element from the collection
     */
    public function current(): array
    {
        if ($this->Index === 0) {
            $this->next();
        }

        return $this->CurrentRow;
    }

    /**
     * Move forward to the next element.
     * Similar to the next() function for arrays in PHP.
     */
    public function next(): void
    {
        // Internal counter is advanced here instead of in an if statement
        // because apparently it's fully possible that an empty row will not be
        // present at all.
        $this->Index++;

        if ($this->Error) {
            $this->CurrentRow = [];
            return;
        }

        if (isset($this->Handle->sheets[$this->CurrentSheet]['cells'][$this->Index])) {
            $this->CurrentRow = $this->Handle->sheets[$this->CurrentSheet]['cells'][$this->Index];

            if (!$this->CurrentRow) {
                $this->CurrentRow = [];
                return;
            }

            $this->CurrentRow = $this->CurrentRow + $this->EmptyRow;
            ksort($this->CurrentRow);
            $this->CurrentRow = array_values($this->CurrentRow);
        } else {
            $this->CurrentRow = $this->EmptyRow;
        }
    }

    /**
     * Return the identifying key of the current element.
     * Similar to the key() function for arrays in PHP.
     *
     * @return int
     */
    public function key(): int
    {
        return $this->Index;
    }

    /**
     * Check if there is a current element after calls to rewind() or next().
     * Used to check if we've iterated to the end of the collection.
     *
     * @return bool FALSE if there's nothing more to iterate over
     */
    public function valid(): bool
    {
        if ($this->Error) {
            return false;
        }

        return $this->Index <= $this->RowCount;
    }

    // -------------------------------------------------------------------------
    // Countable interface method
    // -------------------------------------------------------------------------

    /**
     * Ostensibly should return the count of the contained items but this just returns
     * the number of rows read so far. Not perfectly correct but at least coherent.
     */
    public function count(): int
    {
        if ($this->Error) {
            return 0;
        }

        return $this->RowCount;
    }
}