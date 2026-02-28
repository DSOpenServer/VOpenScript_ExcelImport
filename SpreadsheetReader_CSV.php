<?php

class SpreadsheetReader_CSV implements Iterator, Countable
{
    /**
     * @var array Options array, pre-populated with the default values.
     */
    private array $Options = [
        'Delimiter' => ';',
        'Enclosure' => '"',
    ];

    private string $Encoding  = 'UTF-8';
    private int    $BOMLength = 0;

    /**
     * @var resource|false File handle
     */
    private mixed $Handle = false;

    private string $Filepath   = '';
    private int    $Index      = 0;
    private ?array $CurrentRow = null;

    /**
     * @param string     $Filepath Path to file
     * @param array|null $Options  Options:
     *                             Enclosure => string CSV enclosure
     *                             Delimiter => string CSV delimiter
     *
     * @throws Exception
     */
    public function __construct(string $Filepath, ?array $Options = null)
    {
        $this->Filepath = $Filepath;

        if (!is_readable($Filepath)) {
            throw new Exception('SpreadsheetReader_CSV: File not readable (' . $Filepath . ')');
        }

        if ($Options !== null) {
            $this->Options = array_merge($this->Options, $Options);
        }

        $this->Handle = fopen($Filepath, 'r');

        // Checking the file for byte-order mark to determine encoding
        $BOM16 = bin2hex(fread($this->Handle, 2));
        if ($BOM16 === 'fffe') {
            $this->Encoding  = 'UTF-16LE';
            $this->BOMLength = 2;
        } elseif ($BOM16 === 'feff') {
            $this->Encoding  = 'UTF-16BE';
            $this->BOMLength = 2;
        }

        if (!$this->BOMLength) {
            fseek($this->Handle, 0);
            $BOM32 = bin2hex(fread($this->Handle, 4));
            if ($BOM32 === '0000feff') {
                $this->Encoding  = 'UTF-32';
                $this->BOMLength = 4;
            } elseif ($BOM32 === 'fffe0000') {
                $this->Encoding  = 'UTF-32';
                $this->BOMLength = 4;
            }
        }

        fseek($this->Handle, 0);
        $BOM8 = bin2hex(fread($this->Handle, 3));
        if ($BOM8 === 'efbbbf') {
            $this->Encoding  = 'UTF-8';
            $this->BOMLength = 3;
        }

        // Seeking the place right after BOM as the start of the real content
        if ($this->BOMLength) {
            fseek($this->Handle, $this->BOMLength);
        }

        // Checking for the delimiter if it should be determined automatically
        if (!$this->Options['Delimiter']) {
            $Semicolon = ';';
            $Tab       = "\t";
            $Comma     = ',';

            // Reading the first row and checking if a specific separator character
            // has more columns than others (most columns = most likely delimiter).
            $SemicolonCount = count(fgetcsv($this->Handle, null, $Semicolon));
            fseek($this->Handle, $this->BOMLength);
            $TabCount = count(fgetcsv($this->Handle, null, $Tab));
            fseek($this->Handle, $this->BOMLength);
            $CommaCount = count(fgetcsv($this->Handle, null, $Comma));
            fseek($this->Handle, $this->BOMLength);

            $Delimiter = $Semicolon;
            if ($TabCount > $SemicolonCount || $CommaCount > $SemicolonCount) {
                $Delimiter = $CommaCount > $TabCount ? $Comma : $Tab;
            }

            $this->Options['Delimiter'] = $Delimiter;
        }
    }

    /**
     * Destructor — closes the file handle if still open.
     */
    public function __destruct()
    {
        if ($this->Handle !== false && is_resource($this->Handle)) {
            fclose($this->Handle);
        }
    }

    /**
     * Returns information about sheets in the file.
     * Because CSV doesn't have any, it's just a single entry.
     *
     * @return array Sheet data
     */
    public function Sheets(): array
    {
        return [0 => basename($this->Filepath)];
    }

    /**
     * Changes sheet to another. Because CSV doesn't have any sheets
     * it just rewinds the file so the behaviour is compatible with other
     * sheet readers. (If an invalid index is given, it doesn't do anything.)
     *
     * @param int $Index Sheet index
     *
     * @return bool Status
     */
    public function ChangeSheet(int $Index): bool
    {
        if ($Index === 0) {
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
        fseek($this->Handle, $this->BOMLength);
        $this->CurrentRow = null;
        $this->Index      = 0;
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
        $this->CurrentRow = [];

        // Finding the place the next line starts for UTF-16 encoded files.
        // Line breaks could be 0x0D 0x00 0x0A 0x00 and PHP could split lines on the
        // first or the second linebreak leaving unnecessary \0 characters that mess up the output.
        if ($this->Encoding === 'UTF-16LE' || $this->Encoding === 'UTF-16BE') {
            while (!feof($this->Handle)) {
                $Char = ord(fgetc($this->Handle));
                if (!$Char || $Char === 10 || $Char === 13) {
                    continue;
                } else {
                    // Step back to the last position before the significant byte
                    if ($this->Encoding === 'UTF-16LE') {
                        fseek($this->Handle, ftell($this->Handle) - 1);
                    } else {
                        fseek($this->Handle, ftell($this->Handle) - 2);
                    }
                    break;
                }
            }
        }

        $this->Index++;
        $Row = fgetcsv($this->Handle, null, $this->Options['Delimiter'], $this->Options['Enclosure']);
        $this->CurrentRow = $Row !== false ? $Row : [];

        if ($this->CurrentRow) {
            // Converting multi-byte unicode strings and trimming enclosure symbols
            // because those aren't recognized in the relevant encodings.
            if ($this->Encoding !== 'ASCII' && $this->Encoding !== 'UTF-8') {
                foreach ($this->CurrentRow as $Key => $Value) {
                    $this->CurrentRow[$Key] = trim(trim(
                        mb_convert_encoding($Value, 'UTF-8', $this->Encoding),
                        $this->Options['Enclosure']
                    ));
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
        return !empty($this->CurrentRow) || !feof($this->Handle);
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