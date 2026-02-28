<?php
class SpreadsheetReader_ODS implements Iterator, Countable
{
    private array $Options = [
        'TempDir'               => '',
        'ReturnDateTimeObjects' => false,
    ];

    private string $TempDir = '';

    /**
     * @var string Path to temporary content file
     */
    private string $ContentPath = '';

    /**
     * @var XMLReader|false XML reader object
     */
    private XMLReader|false $Content = false;

    /**
     * @var XMLReader|null Sheet reader object
     */
    private ?XMLReader $SheetReader = null;

    /**
     * @var array|false Data about separate sheets in the file
     */
    private array|false $Sheets = false;

    private ?array $CurrentRow = null;

    /**
     * @var int Number of the sheet we're currently reading
     */
    private int $CurrentSheet = 0;

    private int  $Index     = 0;
    private bool $TableOpen = false;
    private bool $RowOpen   = false;
    private bool $Valid     = false;

    /**
     * @param string     $Filepath Path to file
     * @param array|null $Options  Options:
     *                             TempDir               => string Temporary directory path
     *                             ReturnDateTimeObjects => bool   True => dates/times returned as DateTime objects, false => strings
     *
     * @throws Exception
     */
    public function __construct(string $Filepath, ?array $Options = null)
    {
        if (!is_readable($Filepath)) {
            throw new Exception('SpreadsheetReader_ODS: File not readable (' . $Filepath . ')');
        }

        $this->TempDir = isset($Options['TempDir']) && is_writable($Options['TempDir'])
            ? $Options['TempDir']
            : sys_get_temp_dir();

        $this->TempDir = rtrim($this->TempDir, DIRECTORY_SEPARATOR);
        $this->TempDir = $this->TempDir . DIRECTORY_SEPARATOR . uniqid() . DIRECTORY_SEPARATOR;

        if ($Options !== null) {
            $this->Options = array_merge($this->Options, $Options);
        }

        $Zip    = new ZipArchive();
        $Status = $Zip->open($Filepath);

        if ($Status !== true) {
            throw new Exception('SpreadsheetReader_ODS: File not readable (' . $Filepath . ') (Error ' . $Status . ')');
        }

        if ($Zip->locateName('content.xml') !== false) {
            $Zip->extractTo($this->TempDir, 'content.xml');
            $this->ContentPath = $this->TempDir . 'content.xml';
        }

        $Zip->close();

        if ($this->ContentPath && is_readable($this->ContentPath)) {
            $this->Content = new XMLReader();
            $this->Content->open($this->ContentPath);
            $this->Valid = true;
        }
    }

    /**
     * Destructor — closes and deletes temp files.
     */
    public function __destruct()
    {
        if ($this->Content instanceof XMLReader) {
            $this->Content->close();
            unset($this->Content);
        }
        if (isset($this->ContentPath) && file_exists($this->ContentPath)) {
            @unlink($this->ContentPath);
        }
    }

    /**
     * Retrieves an array with information about sheets in the current file.
     *
     * @return array List of sheets (key is sheet index, value is name)
     */
    public function Sheets(): array
    {
        if ($this->Sheets === false) {
            $this->Sheets = [];

            if ($this->Valid) {
                $this->SheetReader = new XMLReader();
                $this->SheetReader->open($this->ContentPath);

                while ($this->SheetReader->read()) {
                    if ($this->SheetReader->name === 'table:table') {
                        $this->Sheets[] = $this->SheetReader->getAttribute('table:name');
                        $this->SheetReader->next();
                    }
                }

                $this->SheetReader->close();
                $this->SheetReader = null;
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
        if (isset($Sheets[$Index])) {
            $this->CurrentSheet = $Index;
            $this->rewind();
            return true;
        }

        return false;
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
        if ($this->Index > 0) {
            // If the worksheet was already iterated, XML file is reopened.
            // Otherwise it should be at the beginning anyway.
            $this->Content->close();
            $this->Content->open($this->ContentPath);
            $this->Valid = true;

            $this->TableOpen = false;
            $this->RowOpen   = false;
            $this->CurrentRow = null;
        }

        $this->Index = 0;
    }

    /**
     * Return the current element.
     * Similar to the current() function for arrays in PHP.
     *
     * @return array|null Current element from the collection
     */
    public function current(): ?array
    {
        if ($this->Index === 0 && is_null($this->CurrentRow)) {
            $this->next();
            $this->Index--;
        }
        return $this->CurrentRow;
    }

    /**
     * Move forward to the next element.
     * Similar to the next() function for arrays in PHP.
     */
    public function next(): void
    {
        $this->Index++;
        $this->CurrentRow = [];

        if (!$this->TableOpen) {
            $TableCounter = 0;
            $SkipRead     = false;

            while ($this->Valid = ($SkipRead || $this->Content->read())) {
                if ($SkipRead) {
                    $SkipRead = false;
                }

                if ($this->Content->name === 'table:table' && $this->Content->nodeType !== XMLReader::END_ELEMENT) {
                    if ($TableCounter === $this->CurrentSheet) {
                        $this->TableOpen = true;
                        break;
                    }

                    $TableCounter++;
                    $this->Content->next();
                    $SkipRead = true;
                }
            }
        }

        if ($this->TableOpen && !$this->RowOpen) {
            while ($this->Valid = $this->Content->read()) {
                switch ($this->Content->name) {
                    case 'table:table':
                        $this->TableOpen = false;
                        $this->Content->next('office:document-content');
                        $this->Valid = false;
                        break 2;
                    case 'table:table-row':
                        if ($this->Content->nodeType !== XMLReader::END_ELEMENT) {
                            $this->RowOpen = true;
                            break 2;
                        }
                        break;
                }
            }
        }

        if ($this->RowOpen) {
            $LastCellContent = '';

            while ($this->Valid = $this->Content->read()) {
                switch ($this->Content->name) {
                    case 'table:table-cell':
                        if ($this->Content->nodeType === XMLReader::END_ELEMENT || $this->Content->isEmptyElement) {
                            if ($this->Content->nodeType === XMLReader::END_ELEMENT) {
                                $LastCellContent = $LastCellContent; // already set
                            } elseif ($this->Content->isEmptyElement) {
                                $LastCellContent = '';
                            }

                            $this->CurrentRow[] = $LastCellContent;

                            $RepeatedAttr = $this->Content->getAttribute('table:number-columns-repeated');
                            if ($RepeatedAttr !== null) {
                                $RepeatedColumnCount = (int)$RepeatedAttr;
                                if ($RepeatedColumnCount > 1) {
                                    $this->CurrentRow = array_pad(
                                        $this->CurrentRow,
                                        count($this->CurrentRow) + $RepeatedColumnCount - 1,
                                        $LastCellContent
                                    );
                                }
                            }
                        } else {
                            $LastCellContent = '';
                        }
                        // intentional fall-through to 'text:p'
                        // no break

                    case 'text:p':
                        if ($this->Content->nodeType !== XMLReader::END_ELEMENT) {
                            $LastCellContent = $this->Content->readString();
                        }
                        break;

                    case 'table:table-row':
                        $this->RowOpen = false;
                        break 2;
                }
            }
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
        return $this->Valid;
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
        return $this->Index + 1;
    }
}