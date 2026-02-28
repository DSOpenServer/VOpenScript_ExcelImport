<?php

class SpreadsheetReader implements SeekableIterator, Countable
{
    const TYPE_XLSX = 'XLSX';
    const TYPE_XLS  = 'XLS';
    const TYPE_CSV  = 'CSV';
    const TYPE_ODS  = 'ODS';

    private array $Options = [
        'Delimiter' => '',
        'Enclosure' => '"',
    ];

    /**
     * @var int Current row in the file
     */
    private int $Index = 0;

    /**
     * @var SpreadsheetReader_XLSX|SpreadsheetReader_XLS|SpreadsheetReader_CSV|SpreadsheetReader_ODS|array Handle for the reader object
     */
    private mixed $Handle = [];

    /**
     * @var string|false Type of the contained spreadsheet
     */
    private string|false $Type = false;

    /**
     * @param string      $Filepath         Path to file
     * @param string|false $OriginalFilename Original filename (in case of an uploaded file), used to determine file type, optional
     * @param string|false $MimeType         MIME type from an upload, used to determine file type, optional
     *
     * @throws Exception
     */
    public function __construct(string $Filepath, string|false $OriginalFilename = false, string|false $MimeType = false)
    {
        if (!is_readable($Filepath)) {
            throw new Exception('SpreadsheetReader: File (' . $Filepath . ') not readable');
        }

        // To avoid timezone warnings and exceptions for formatting dates retrieved from files
        $DefaultTZ = @date_default_timezone_get();
        if ($DefaultTZ) {
            date_default_timezone_set($DefaultTZ);
        }

        // Checking the other parameters for correctness
        if (!empty($OriginalFilename) && !is_scalar($OriginalFilename)) {
            throw new Exception('SpreadsheetReader: Original file (2nd parameter) path is not a string or a scalar value.');
        }
        if (!empty($MimeType) && !is_scalar($MimeType)) {
            throw new Exception('SpreadsheetReader: Mime type (3rd parameter) path is not a string or a scalar value.');
        }

        // 1. Determine type
        if (!$OriginalFilename) {
            $OriginalFilename = $Filepath;
        }

        $Extension = strtolower(pathinfo($OriginalFilename, PATHINFO_EXTENSION));

        switch ($MimeType) {
            case 'text/csv':
            case 'text/comma-separated-values':
            case 'text/plain':
                $this->Type = self::TYPE_CSV;
                break;
            case 'application/vnd.ms-excel':
            case 'application/msexcel':
            case 'application/x-msexcel':
            case 'application/x-ms-excel':
            case 'application/x-excel':
            case 'application/x-dos_ms_excel':
            case 'application/xls':
            case 'application/xlt':
            case 'application/x-xls':
                // Excel does weird stuff
                if (in_array($Extension, ['csv', 'tsv', 'txt'])) {
                    $this->Type = self::TYPE_CSV;
                } else {
                    $this->Type = self::TYPE_XLS;
                }
                break;
            case 'application/vnd.oasis.opendocument.spreadsheet':
            case 'application/vnd.oasis.opendocument.spreadsheet-template':
                $this->Type = self::TYPE_ODS;
                break;
            case 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
            case 'application/vnd.openxmlformats-officedocument.spreadsheetml.template':
            case 'application/xlsx':
            case 'application/xltx':
                $this->Type = self::TYPE_XLSX;
                break;
            case 'application/xml':
                // Excel 2004 xml format uses this
                break;
        }

        if (!$this->Type) {
            switch ($Extension) {
                case 'xlsx':
                case 'xltx': // XLSX template
                case 'xlsm': // Macro-enabled XLSX
                case 'xltm': // Macro-enabled XLSX template
                    $this->Type = self::TYPE_XLSX;
                    break;
                case 'xls':
                case 'xlt':
                    $this->Type = self::TYPE_XLS;
                    break;
                case 'ods':
                case 'odt':
                    $this->Type = self::TYPE_ODS;
                    break;
                default:
                    $this->Type = self::TYPE_CSV;
                    break;
            }
        }

        // Pre-checking XLS files, in case they are renamed CSV or XLSX files
        if ($this->Type === self::TYPE_XLS) {
            self::Load(self::TYPE_XLS);
            $this->Handle = new SpreadsheetReader_XLS($Filepath);
            if ($this->Handle->Error) {
                $this->Handle->__destruct();

                // PHP 8: zip_open() is deprecated/removed; use ZipArchive instead
                $zip = new ZipArchive();
                if ($zip->open($Filepath) === true) {
                    $zip->close();
                    $this->Type = self::TYPE_XLSX;
                } else {
                    $this->Type = self::TYPE_CSV;
                }
            }
        }

        // 2. Create handle
        switch ($this->Type) {
            case self::TYPE_XLSX:
                self::Load(self::TYPE_XLSX);
                $this->Handle = new SpreadsheetReader_XLSX($Filepath);
                break;
            case self::TYPE_CSV:
                self::Load(self::TYPE_CSV);
                $this->Handle = new SpreadsheetReader_CSV($Filepath, $this->Options);
                break;
            case self::TYPE_XLS:
                // Everything already happens above
                break;
            case self::TYPE_ODS:
                self::Load(self::TYPE_ODS);
                $this->Handle = new SpreadsheetReader_ODS($Filepath, $this->Options);
                break;
        }
    }

    /**
     * Gets information about separate sheets in the given file.
     *
     * @return array Associative array where key is sheet index and value is sheet name
     */
    public function Sheets(): array
    {
        return $this->Handle->Sheets();
    }

    /**
     * Changes the current sheet to another from the file.
     * Note that changing the sheet will rewind the file to the beginning, even if
     * the current sheet index is provided.
     *
     * @param int $Index Sheet index
     *
     * @return bool True if sheet could be changed to the specified one, false otherwise
     */
    public function ChangeSheet(int $Index): bool
    {
        return $this->Handle->ChangeSheet($Index);
    }

    /**
     * Autoloads the required class for the particular spreadsheet type.
     *
     * @param string $Type Spreadsheet type, one of the TYPE_* constants of this class
     *
     * @throws Exception
     */
    private static function Load(string $Type): void
    {
        if (!in_array($Type, [self::TYPE_XLSX, self::TYPE_XLS, self::TYPE_CSV, self::TYPE_ODS])) {
            throw new Exception('SpreadsheetReader: Invalid type (' . $Type . ')');
        }

        if (!class_exists('SpreadsheetReader_' . $Type, false)) {
            require(dirname(__FILE__) . DIRECTORY_SEPARATOR . 'SpreadsheetReader_' . $Type . '.php');
        }
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
        if ($this->Handle) {
            $this->Handle->rewind();
        }
    }

    /**
     * Return the current element.
     * Similar to the current() function for arrays in PHP.
     *
     * @return mixed Current element from the collection
     */
    public function current(): mixed
    {
        if ($this->Handle) {
            return $this->Handle->current();
        }
        return null;
    }

    /**
     * Move forward to the next element.
     * Similar to the next() function for arrays in PHP.
     */
    public function next(): void
    {
        if ($this->Handle) {
            $this->Index++;
            $this->Handle->next();
        }
    }

    /**
     * Return the identifying key of the current element.
     * Similar to the key() function for arrays in PHP.
     *
     * @return int|null
     */
    public function key(): int|null
    {
        if ($this->Handle) {
            return $this->Handle->key();
        }
        return null;
    }

    /**
     * Check if there is a current element after calls to rewind() or next().
     * Used to check if we've iterated to the end of the collection.
     *
     * @return bool FALSE if there's nothing more to iterate over
     */
    public function valid(): bool
    {
        if ($this->Handle) {
            return $this->Handle->valid();
        }
        return false;
    }

    // -------------------------------------------------------------------------
    // Countable interface method
    // -------------------------------------------------------------------------

    public function count(): int
    {
        if ($this->Handle) {
            return $this->Handle->count();
        }
        return 0;
    }

    // -------------------------------------------------------------------------
    // SeekableIterator interface method
    // -------------------------------------------------------------------------

    /**
     * Takes a position and traverses the file to that position.
     * The value can be retrieved with a current() call afterwards.
     *
     * @param int $Position Position in file
     *
     * @throws OutOfBoundsException
     */
    public function seek(int $Position): void
    {
        if (!$this->Handle) {
            throw new OutOfBoundsException('SpreadsheetReader: No file opened');
        }

        $CurrentIndex = $this->Handle->key();

        if ($CurrentIndex !== $Position) {
            if ($Position < $CurrentIndex || is_null($CurrentIndex) || $Position === 0) {
                $this->rewind();
            }

            while ($this->Handle->valid() && ($Position > $this->Handle->key())) {
                $this->Handle->next();
            }

            if (!$this->Handle->valid()) {
                throw new OutOfBoundsException('SpreadsheetError: Position ' . $Position . ' not found');
            }
        }
    }
}