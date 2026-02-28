<?php

define('NUM_BIG_BLOCK_DEPOT_BLOCKS_POS', 0x2c);
define('SMALL_BLOCK_DEPOT_BLOCK_POS', 0x3c);
define('ROOT_START_BLOCK_POS', 0x30);
define('BIG_BLOCK_SIZE', 0x200);
define('SMALL_BLOCK_SIZE', 0x40);
define('EXTENSION_BLOCK_POS', 0x44);
define('NUM_EXTENSION_BLOCK_POS', 0x48);
define('PROPERTY_STORAGE_BLOCK_SIZE', 0x80);
define('BIG_BLOCK_DEPOT_BLOCKS_POS', 0x4c);
define('SMALL_BLOCK_THRESHOLD', 0x1000);
define('SIZE_OF_NAME_POS', 0x40);
define('TYPE_POS', 0x42);
define('START_BLOCK_POS', 0x74);
define('SIZE_POS', 0x78);
define('IDENTIFIER_OLE', pack("CCCCCCCC", 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1));


function GetInt4d(string $data, int $pos): int
{
    $value = ord($data[$pos]) | (ord($data[$pos + 1]) << 8) | (ord($data[$pos + 2]) << 16) | (ord($data[$pos + 3]) << 24);
    if ($value >= 4294967294) {
        $value = -2;
    }
    return $value;
}

function gmgetdate(?int $ts = null): array
{
    $k = ['seconds', 'minutes', 'hours', 'mday', 'wday', 'mon', 'year', 'yday', 'weekday', 'month', 0];
    return array_combine($k, explode(":", gmdate('s:i:G:j:w:n:Y:z:l:F:U', $ts ?? time())));
}

function v(string $data, int $pos): int
{
    return ord($data[$pos]) | ord($data[$pos + 1]) << 8;
}


class OLERead
{
    public string $data = '';
    public int $error = 0;
    public int $numBigBlockDepotBlocks = 0;
    public int $sbdStartBlock = 0;
    public int $rootStartBlock = 0;
    public int $extensionBlock = 0;
    public int $numExtensionBlocks = 0;
    public array $bigBlockChain = [];
    public array $smallBlockChain = [];
    public string $entry = '';
    public array $props = [];
    public int $wrkbook = 0;
    public int $rootentry = 0;

    public function __construct() {}

    public function read(string $sFileName): bool
    {
        if (!is_readable($sFileName)) {
            $this->error = 1;
            return false;
        }

        $this->data = @file_get_contents($sFileName);
        if (!$this->data) {
            $this->error = 1;
            return false;
        }

        if (substr($this->data, 0, 8) !== IDENTIFIER_OLE) {
            $this->error = 1;
            return false;
        }

        $this->numBigBlockDepotBlocks = GetInt4d($this->data, NUM_BIG_BLOCK_DEPOT_BLOCKS_POS);
        $this->sbdStartBlock = GetInt4d($this->data, SMALL_BLOCK_DEPOT_BLOCK_POS);
        $this->rootStartBlock = GetInt4d($this->data, ROOT_START_BLOCK_POS);
        $this->extensionBlock = GetInt4d($this->data, EXTENSION_BLOCK_POS);
        $this->numExtensionBlocks = GetInt4d($this->data, NUM_EXTENSION_BLOCK_POS);

        $bigBlockDepotBlocks = [];
        $pos = BIG_BLOCK_DEPOT_BLOCKS_POS;
        $bbdBlocks = $this->numBigBlockDepotBlocks;

        if ($this->numExtensionBlocks != 0) {
            $bbdBlocks = (BIG_BLOCK_SIZE - BIG_BLOCK_DEPOT_BLOCKS_POS) / 4;
        }

        for ($i = 0; $i < $bbdBlocks; $i++) {
            $bigBlockDepotBlocks[$i] = GetInt4d($this->data, $pos);
            $pos += 4;
        }

        for ($j = 0; $j < $this->numExtensionBlocks; $j++) {
            $pos = ($this->extensionBlock + 1) * BIG_BLOCK_SIZE;
            $blocksToRead = min($this->numBigBlockDepotBlocks - $bbdBlocks, BIG_BLOCK_SIZE / 4 - 1);

            for ($i = $bbdBlocks; $i < $bbdBlocks + $blocksToRead; $i++) {
                $bigBlockDepotBlocks[$i] = GetInt4d($this->data, $pos);
                $pos += 4;
            }

            $bbdBlocks += $blocksToRead;
            if ($bbdBlocks < $this->numBigBlockDepotBlocks) {
                $this->extensionBlock = GetInt4d($this->data, $pos);
            }
        }

        // readBigBlockDepot
        $index = 0;
        $this->bigBlockChain = [];

        for ($i = 0; $i < $this->numBigBlockDepotBlocks; $i++) {
            $pos = ($bigBlockDepotBlocks[$i] + 1) * BIG_BLOCK_SIZE;
            for ($j = 0; $j < BIG_BLOCK_SIZE / 4; $j++) {
                $this->bigBlockChain[$index] = GetInt4d($this->data, $pos);
                $pos += 4;
                $index++;
            }
        }

        // readSmallBlockDepot
        $index = 0;
        $sbdBlock = $this->sbdStartBlock;
        $this->smallBlockChain = [];

        while ($sbdBlock != -2) {
            $pos = ($sbdBlock + 1) * BIG_BLOCK_SIZE;
            for ($j = 0; $j < BIG_BLOCK_SIZE / 4; $j++) {
                $this->smallBlockChain[$index] = GetInt4d($this->data, $pos);
                $pos += 4;
                $index++;
            }
            $sbdBlock = $this->bigBlockChain[$sbdBlock];
        }

        $this->entry = $this->_readData($this->rootStartBlock);
        $this->_readPropertySets();

        return true;
    }

    private function _readData(int $bl): string
    {
        $block = $bl;
        $data = '';
        while ($block != -2) {
            $pos = ($block + 1) * BIG_BLOCK_SIZE;
            $data .= substr($this->data, $pos, BIG_BLOCK_SIZE);
            $block = $this->bigBlockChain[$block];
        }
        return $data;
    }

    private function _readPropertySets(): void
    {
        $offset = 0;
        while ($offset < strlen($this->entry)) {
            $d = substr($this->entry, $offset, PROPERTY_STORAGE_BLOCK_SIZE);
            $nameSize = ord($d[SIZE_OF_NAME_POS]) | (ord($d[SIZE_OF_NAME_POS + 1]) << 8);
            $type = ord($d[TYPE_POS]);
            $startBlock = GetInt4d($d, START_BLOCK_POS);
            $size = GetInt4d($d, SIZE_POS);
            $name = '';
            for ($i = 0; $i < $nameSize; $i++) {
                $name .= $d[$i];
            }
            $name = str_replace("\x00", "", $name);
            $this->props[] = [
                'name'       => $name,
                'type'       => $type,
                'startBlock' => $startBlock,
                'size'       => $size,
            ];
            if (strtolower($name) === "workbook" || strtolower($name) === "book") {
                $this->wrkbook = count($this->props) - 1;
            }
            if ($name === "Root Entry") {
                $this->rootentry = count($this->props) - 1;
            }
            $offset += PROPERTY_STORAGE_BLOCK_SIZE;
        }
    }

    public function getWorkBook(): string
    {
        if ($this->props[$this->wrkbook]['size'] < SMALL_BLOCK_THRESHOLD) {
            $rootdata = $this->_readData($this->props[$this->rootentry]['startBlock']);
            $streamData = '';
            $block = $this->props[$this->wrkbook]['startBlock'];
            while ($block != -2) {
                $pos = $block * SMALL_BLOCK_SIZE;
                $streamData .= substr($rootdata, $pos, SMALL_BLOCK_SIZE);
                $block = $this->smallBlockChain[$block];
            }
            return $streamData;
        } else {
            $numBlocks = $this->props[$this->wrkbook]['size'] / BIG_BLOCK_SIZE;
            if ($this->props[$this->wrkbook]['size'] % BIG_BLOCK_SIZE != 0) {
                $numBlocks++;
            }
            if ($numBlocks == 0) return '';

            $streamData = '';
            $block = $this->props[$this->wrkbook]['startBlock'];
            while ($block != -2) {
                $pos = ($block + 1) * BIG_BLOCK_SIZE;
                $streamData .= substr($this->data, $pos, BIG_BLOCK_SIZE);
                $block = $this->bigBlockChain[$block];
            }
            return $streamData;
        }
    }
}


define('SPREADSHEET_EXCEL_READER_BIFF8', 0x600);
define('SPREADSHEET_EXCEL_READER_BIFF7', 0x500);
define('SPREADSHEET_EXCEL_READER_WORKBOOKGLOBALS', 0x5);
define('SPREADSHEET_EXCEL_READER_WORKSHEET', 0x10);
define('SPREADSHEET_EXCEL_READER_TYPE_BOF', 0x809);
define('SPREADSHEET_EXCEL_READER_TYPE_EOF', 0x0a);
define('SPREADSHEET_EXCEL_READER_TYPE_BOUNDSHEET', 0x85);
define('SPREADSHEET_EXCEL_READER_TYPE_DIMENSION', 0x200);
define('SPREADSHEET_EXCEL_READER_TYPE_ROW', 0x208);
define('SPREADSHEET_EXCEL_READER_TYPE_DBCELL', 0xd7);
define('SPREADSHEET_EXCEL_READER_TYPE_FILEPASS', 0x2f);
define('SPREADSHEET_EXCEL_READER_TYPE_NOTE', 0x1c);
define('SPREADSHEET_EXCEL_READER_TYPE_TXO', 0x1b6);
define('SPREADSHEET_EXCEL_READER_TYPE_RK', 0x7e);
define('SPREADSHEET_EXCEL_READER_TYPE_RK2', 0x27e);
define('SPREADSHEET_EXCEL_READER_TYPE_MULRK', 0xbd);
define('SPREADSHEET_EXCEL_READER_TYPE_MULBLANK', 0xbe);
define('SPREADSHEET_EXCEL_READER_TYPE_INDEX', 0x20b);
define('SPREADSHEET_EXCEL_READER_TYPE_SST', 0xfc);
define('SPREADSHEET_EXCEL_READER_TYPE_EXTSST', 0xff);
define('SPREADSHEET_EXCEL_READER_TYPE_CONTINUE', 0x3c);
define('SPREADSHEET_EXCEL_READER_TYPE_LABEL', 0x204);
define('SPREADSHEET_EXCEL_READER_TYPE_LABELSST', 0xfd);
define('SPREADSHEET_EXCEL_READER_TYPE_NUMBER', 0x203);
define('SPREADSHEET_EXCEL_READER_TYPE_NAME', 0x18);
define('SPREADSHEET_EXCEL_READER_TYPE_ARRAY', 0x221);
define('SPREADSHEET_EXCEL_READER_TYPE_STRING', 0x207);
define('SPREADSHEET_EXCEL_READER_TYPE_FORMULA', 0x406);
define('SPREADSHEET_EXCEL_READER_TYPE_FORMULA2', 0x6);
define('SPREADSHEET_EXCEL_READER_TYPE_FORMAT', 0x41e);
define('SPREADSHEET_EXCEL_READER_TYPE_XF', 0xe0);
define('SPREADSHEET_EXCEL_READER_TYPE_BOOLERR', 0x205);
define('SPREADSHEET_EXCEL_READER_TYPE_FONT', 0x0031);
define('SPREADSHEET_EXCEL_READER_TYPE_PALETTE', 0x0092);
define('SPREADSHEET_EXCEL_READER_TYPE_UNKNOWN', 0xffff);
define('SPREADSHEET_EXCEL_READER_TYPE_NINETEENFOUR', 0x22);
define('SPREADSHEET_EXCEL_READER_TYPE_MERGEDCELLS', 0xE5);
define('SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS', 25569);
define('SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS1904', 24107);
define('SPREADSHEET_EXCEL_READER_MSINADAY', 86400);
define('SPREADSHEET_EXCEL_READER_TYPE_HYPER', 0x01b8);
define('SPREADSHEET_EXCEL_READER_TYPE_COLINFO', 0x7d);
define('SPREADSHEET_EXCEL_READER_TYPE_DEFCOLWIDTH', 0x55);
define('SPREADSHEET_EXCEL_READER_TYPE_STANDARDWIDTH', 0x99);
define('SPREADSHEET_EXCEL_READER_DEF_NUM_FORMAT', "%s");


class Spreadsheet_Excel_Reader
{
    public array $colnames = [];
    public array $colindexes = [];
    public int $standardColWidth = 0;
    public int $defaultColWidth = 0;

    public array $boundsheets = [];
    public array $formatRecords = [];
    public array $fontRecords = [];
    public array $xfRecords = [];
    public array $colInfo = [];
    public array $rowInfo = [];
    public array $sst = [];
    public array $sheets = [];

    public string $data = '';
    public OLERead $_ole;
    public string $_defaultEncoding = "UTF-8";
    public string $_defaultFormat = SPREADSHEET_EXCEL_READER_DEF_NUM_FORMAT;
    public array $_columnsFormat = [];
    public int $_rowoffset = 1;
    public int $_coloffset = 1;
    public bool $store_extended_info = true;
    public bool $nineteenFour = false;
    public int $version = 0;
    public int $sn = 0;
    public string $_encoderFunction = '';

    public array $dateFormats = [
        0xe  => "m/d/Y",
        0xf  => "M-d-Y",
        0x10 => "d-M",
        0x11 => "M-Y",
        0x12 => "h:i a",
        0x13 => "h:i:s a",
        0x14 => "H:i",
        0x15 => "H:i:s",
        0x16 => "d/m/Y H:i",
        0x2d => "i:s",
        0x2e => "H:i:s",
        0x2f => "i:s.S",
    ];

    public array $numberFormats = [
        0x1  => "0",
        0x2  => "0.00",
        0x3  => "#,##0",
        0x4  => "#,##0.00",
        0x5  => "\$#,##0;(\$#,##0)",
        0x6  => "\$#,##0;[Red](\$#,##0)",
        0x7  => "\$#,##0.00;(\$#,##0.00)",
        0x8  => "\$#,##0.00;[Red](\$#,##0.00)",
        0x9  => "0%",
        0xa  => "0.00%",
        0xb  => "0.00E+00",
        0x25 => "#,##0;(#,##0)",
        0x26 => "#,##0;[Red](#,##0)",
        0x27 => "#,##0.00;(#,##0.00)",
        0x28 => "#,##0.00;[Red](#,##0.00)",
        0x29 => "#,##0;(#,##0)",
        0x2a => "\$#,##0;(\$#,##0)",
        0x2b => "#,##0.00;(#,##0.00)",
        0x2c => "\$#,##0.00;(\$#,##0.00)",
        0x30 => "##0.0E+0",
    ];

    public array $colors = [
        0x00 => "#000000",
        0x01 => "#FFFFFF",
        0x02 => "#FF0000",
        0x03 => "#00FF00",
        0x04 => "#0000FF",
        0x05 => "#FFFF00",
        0x06 => "#FF00FF",
        0x07 => "#00FFFF",
        0x08 => "#000000",
        0x09 => "#FFFFFF",
        0x0A => "#FF0000",
        0x0B => "#00FF00",
        0x0C => "#0000FF",
        0x0D => "#FFFF00",
        0x0E => "#FF00FF",
        0x0F => "#00FFFF",
        0x10 => "#800000",
        0x11 => "#008000",
        0x12 => "#000080",
        0x13 => "#808000",
        0x14 => "#800080",
        0x15 => "#008080",
        0x16 => "#C0C0C0",
        0x17 => "#808080",
        0x18 => "#9999FF",
        0x19 => "#993366",
        0x1A => "#FFFFCC",
        0x1B => "#CCFFFF",
        0x1C => "#660066",
        0x1D => "#FF8080",
        0x1E => "#0066CC",
        0x1F => "#CCCCFF",
        0x20 => "#000080",
        0x21 => "#FF00FF",
        0x22 => "#FFFF00",
        0x23 => "#00FFFF",
        0x24 => "#800080",
        0x25 => "#800000",
        0x26 => "#008080",
        0x27 => "#0000FF",
        0x28 => "#00CCFF",
        0x29 => "#CCFFFF",
        0x2A => "#CCFFCC",
        0x2B => "#FFFF99",
        0x2C => "#99CCFF",
        0x2D => "#FF99CC",
        0x2E => "#CC99FF",
        0x2F => "#FFCC99",
        0x30 => "#3366FF",
        0x31 => "#33CCCC",
        0x32 => "#99CC00",
        0x33 => "#FFCC00",
        0x34 => "#FF9900",
        0x35 => "#FF6600",
        0x36 => "#666699",
        0x37 => "#969696",
        0x38 => "#003366",
        0x39 => "#339966",
        0x3A => "#003300",
        0x3B => "#333300",
        0x3C => "#993300",
        0x3D => "#993366",
        0x3E => "#333399",
        0x3F => "#333333",
        0x40 => "#000000",
        0x41 => "#FFFFFF",
        0x43 => "#000000",
        0x4D => "#000000",
        0x4E => "#FFFFFF",
        0x4F => "#000000",
        0x50 => "#FFFFFF",
        0x51 => "#000000",
        0x7FFF => "#000000",
    ];

    public array $lineStyles = [
        0x00 => "",
        0x01 => "Thin",
        0x02 => "Medium",
        0x03 => "Dashed",
        0x04 => "Dotted",
        0x05 => "Thick",
        0x06 => "Double",
        0x07 => "Hair",
        0x08 => "Medium dashed",
        0x09 => "Thin dash-dotted",
        0x0A => "Medium dash-dotted",
        0x0B => "Thin dash-dot-dotted",
        0x0C => "Medium dash-dot-dotted",
        0x0D => "Slanted medium dash-dotted",
    ];

    public array $lineStylesCss = [
        "Thin"                      => "1px solid",
        "Medium"                    => "2px solid",
        "Dashed"                    => "1px dashed",
        "Dotted"                    => "1px dotted",
        "Thick"                     => "3px solid",
        "Double"                    => "double",
        "Hair"                      => "1px solid",
        "Medium dashed"             => "2px dashed",
        "Thin dash-dotted"          => "1px dashed",
        "Medium dash-dotted"        => "2px dashed",
        "Thin dash-dot-dotted"      => "1px dashed",
        "Medium dash-dot-dotted"    => "2px dashed",
        "Slanted medium dash-dotte" => "2px dashed",
    ];

    // -------------------------------------------------------------------------
    // Constructor
    // -------------------------------------------------------------------------

    public function __construct(string $file = '', bool $store_extended_info = true, string $outputEncoding = '')
    {
        $this->_ole = new OLERead();
        $this->setUTFEncoder('iconv');
        if ($outputEncoding !== '') {
            $this->setOutputEncoding($outputEncoding);
        }
        for ($i = 1; $i < 245; $i++) {
            $name = strtolower((($i - 1) / 26 >= 1 ? chr((int)(($i - 1) / 26) + 64) : '') . chr(($i - 1) % 26 + 65));
            $this->colnames[$name] = $i;
            $this->colindexes[$i] = $name;
        }
        $this->store_extended_info = $store_extended_info;
        if ($file !== "") {
            $this->read($file);
        }
    }

    // -------------------------------------------------------------------------
    // Public API
    // -------------------------------------------------------------------------

    public function val(int $row, int|string $col, int $sheet = 0): string
    {
        $col = $this->getCol($col);
        if (array_key_exists($row, $this->sheets[$sheet]['cells']) && array_key_exists($col, $this->sheets[$sheet]['cells'][$row])) {
            return $this->sheets[$sheet]['cells'][$row][$col];
        }
        return "";
    }

    public function value(int $row, int|string $col, int $sheet = 0): string
    {
        return $this->val($row, $col, $sheet);
    }

    public function info(int $row, int|string $col, string $type = '', int $sheet = 0): mixed
    {
        $col = $this->getCol($col);
        if (
            array_key_exists('cellsInfo', $this->sheets[$sheet])
            && array_key_exists($row, $this->sheets[$sheet]['cellsInfo'])
            && array_key_exists($col, $this->sheets[$sheet]['cellsInfo'][$row])
            && array_key_exists($type, $this->sheets[$sheet]['cellsInfo'][$row][$col])
        ) {
            return $this->sheets[$sheet]['cellsInfo'][$row][$col][$type];
        }
        return "";
    }

    public function type(int $row, int|string $col, int $sheet = 0): mixed
    {
        return $this->info($row, $col, 'type', $sheet);
    }

    public function raw(int $row, int|string $col, int $sheet = 0): mixed
    {
        return $this->info($row, $col, 'raw', $sheet);
    }

    public function rowspan(int $row, int|string $col, int $sheet = 0): int
    {
        $val = $this->info($row, $col, 'rowspan', $sheet);
        return ($val === "") ? 1 : (int)$val;
    }

    public function colspan(int $row, int|string $col, int $sheet = 0): int
    {
        $val = $this->info($row, $col, 'colspan', $sheet);
        return ($val === "") ? 1 : (int)$val;
    }

    public function hyperlink(int $row, int|string $col, int $sheet = 0): string
    {
        $link = $this->sheets[$sheet]['cellsInfo'][$row][$col]['hyperlink'] ?? null;
        if ($link) {
            return $link['link'];
        }
        return '';
    }

    public function rowcount(int $sheet = 0): int
    {
        return $this->sheets[$sheet]['numRows'];
    }

    public function colcount(int $sheet = 0): int
    {
        return $this->sheets[$sheet]['numCols'];
    }

    public function colwidth(int $col, int $sheet = 0): float
    {
        return $this->colInfo[$sheet][$col]['width'] / 9142 * 200;
    }

    public function colhidden(int $col, int $sheet = 0): bool
    {
        return (bool)($this->colInfo[$sheet][$col]['hidden'] ?? false);
    }

    public function rowheight(int $row, int $sheet = 0): mixed
    {
        return $this->rowInfo[$sheet][$row]['height'] ?? 0;
    }

    public function rowhidden(int $row, int $sheet = 0): bool
    {
        return (bool)($this->rowInfo[$sheet][$row]['hidden'] ?? false);
    }

    // -------------------------------------------------------------------------
    // CSS / Style helpers
    // -------------------------------------------------------------------------

    public function style(int $row, int|string $col, int $sheet = 0, string $properties = ''): string
    {
        $css = "";

        $font = $this->font($row, $col, $sheet);
        if ($font !== "") $css .= "font-family:$font;";

        $align = $this->align($row, $col, $sheet);
        if ($align !== "") $css .= "text-align:$align;";

        $height = $this->height($row, $col, $sheet);
        if ($height !== "" && $height !== false) $css .= "font-size:{$height}px;";

        $bgcolor = $this->bgColor($row, $col, $sheet);
        if ($bgcolor !== "") {
            $bgcolor = $this->colors[$bgcolor];
            $css .= "background-color:$bgcolor;";
        }

        $color = $this->color($row, $col, $sheet);
        if ($color !== "") $css .= "color:$color;";

        if ($this->bold($row, $col, $sheet)) $css .= "font-weight:bold;";
        if ($this->italic($row, $col, $sheet)) $css .= "font-style:italic;";
        if ($this->underline($row, $col, $sheet)) $css .= "text-decoration:underline;";

        $bLeft   = $this->borderLeft($row, $col, $sheet);
        $bRight  = $this->borderRight($row, $col, $sheet);
        $bTop    = $this->borderTop($row, $col, $sheet);
        $bBottom = $this->borderBottom($row, $col, $sheet);

        $bLeftCol   = $this->borderLeftColor($row, $col, $sheet);
        $bRightCol  = $this->borderRightColor($row, $col, $sheet);
        $bTopCol    = $this->borderTopColor($row, $col, $sheet);
        $bBottomCol = $this->borderBottomColor($row, $col, $sheet);

        if ($bLeft !== "" && $bLeft === $bRight && $bRight === $bTop && $bTop === $bBottom) {
            $css .= "border:" . $this->lineStylesCss[$bLeft] . ";";
        } else {
            if ($bLeft !== "")   $css .= "border-left:" . $this->lineStylesCss[$bLeft] . ";";
            if ($bRight !== "")  $css .= "border-right:" . $this->lineStylesCss[$bRight] . ";";
            if ($bTop !== "")    $css .= "border-top:" . $this->lineStylesCss[$bTop] . ";";
            if ($bBottom !== "") $css .= "border-bottom:" . $this->lineStylesCss[$bBottom] . ";";
        }

        if ($bLeft !== "" && $bLeftCol !== "")   $css .= "border-left-color:$bLeftCol;";
        if ($bRight !== "" && $bRightCol !== "")  $css .= "border-right-color:$bRightCol;";
        if ($bTop !== "" && $bTopCol !== "")      $css .= "border-top-color:$bTopCol;";
        if ($bBottom !== "" && $bBottomCol !== "") $css .= "border-bottom-color:$bBottomCol;";

        return $css;
    }

    // Format properties
    public function format(int $row, int|string $col, int $sheet = 0): mixed      { return $this->info($row, $col, 'format', $sheet); }
    public function formatIndex(int $row, int|string $col, int $sheet = 0): mixed { return $this->info($row, $col, 'formatIndex', $sheet); }
    public function formatColor(int $row, int|string $col, int $sheet = 0): mixed { return $this->info($row, $col, 'formatColor', $sheet); }

    // XF record helpers
    public function xfRecord(int $row, int|string $col, int $sheet = 0): ?array
    {
        $xfIndex = $this->info($row, $col, 'xfIndex', $sheet);
        if ($xfIndex !== "") {
            return $this->xfRecords[$xfIndex] ?? null;
        }
        return null;
    }

    public function xfProperty(int $row, int|string $col, int $sheet, string $prop): mixed
    {
        $xfRecord = $this->xfRecord($row, $col, $sheet);
        return $xfRecord[$prop] ?? "";
    }

    public function align(int $row, int|string $col, int $sheet = 0): mixed      { return $this->xfProperty($row, $col, $sheet, 'align'); }
    public function bgColor(int $row, int|string $col, int $sheet = 0): mixed    { return $this->xfProperty($row, $col, $sheet, 'bgColor'); }
    public function borderLeft(int $row, int|string $col, int $sheet = 0): mixed { return $this->xfProperty($row, $col, $sheet, 'borderLeft'); }
    public function borderRight(int $row, int|string $col, int $sheet = 0): mixed{ return $this->xfProperty($row, $col, $sheet, 'borderRight'); }
    public function borderTop(int $row, int|string $col, int $sheet = 0): mixed  { return $this->xfProperty($row, $col, $sheet, 'borderTop'); }
    public function borderBottom(int $row, int|string $col, int $sheet = 0): mixed{ return $this->xfProperty($row, $col, $sheet, 'borderBottom'); }

    public function borderLeftColor(int $row, int|string $col, int $sheet = 0): string
    {
        return $this->colors[$this->xfProperty($row, $col, $sheet, 'borderLeftColor')] ?? "";
    }

    public function borderRightColor(int $row, int|string $col, int $sheet = 0): string
    {
        return $this->colors[$this->xfProperty($row, $col, $sheet, 'borderRightColor')] ?? "";
    }

    public function borderTopColor(int $row, int|string $col, int $sheet = 0): string
    {
        return $this->colors[$this->xfProperty($row, $col, $sheet, 'borderTopColor')] ?? "";
    }

    public function borderBottomColor(int $row, int|string $col, int $sheet = 0): string
    {
        return $this->colors[$this->xfProperty($row, $col, $sheet, 'borderBottomColor')] ?? "";
    }

    // Font helpers
    public function fontRecord(int $row, int|string $col, int $sheet = 0): ?array
    {
        $xfRecord = $this->xfRecord($row, $col, $sheet);
        if ($xfRecord !== null) {
            $font = $xfRecord['fontIndex'] ?? null;
            if ($font !== null) {
                return $this->fontRecords[$font] ?? null;
            }
        }
        return null;
    }

    public function fontProperty(int $row, int|string $col, int $sheet = 0, string $prop = ''): mixed
    {
        $font = $this->fontRecord($row, $col, $sheet);
        return $font[$prop] ?? false;
    }

    public function fontIndex(int $row, int|string $col, int $sheet = 0): mixed { return $this->xfProperty($row, $col, $sheet, 'fontIndex'); }

    public function color(int $row, int|string $col, int $sheet = 0): string
    {
        $formatColor = $this->formatColor($row, $col, $sheet);
        if ($formatColor !== "") return $formatColor;
        $ci = $this->fontProperty($row, $col, $sheet, 'color');
        return $this->rawColor($ci);
    }

    public function rawColor(mixed $ci): string
    {
        if ($ci !== 0x7FFF && $ci !== '') {
            return $this->colors[$ci] ?? "";
        }
        return "";
    }

    public function bold(int $row, int|string $col, int $sheet = 0): mixed      { return $this->fontProperty($row, $col, $sheet, 'bold'); }
    public function italic(int $row, int|string $col, int $sheet = 0): mixed    { return $this->fontProperty($row, $col, $sheet, 'italic'); }
    public function underline(int $row, int|string $col, int $sheet = 0): mixed { return $this->fontProperty($row, $col, $sheet, 'under'); }
    public function height(int $row, int|string $col, int $sheet = 0): mixed    { return $this->fontProperty($row, $col, $sheet, 'height'); }
    public function font(int $row, int|string $col, int $sheet = 0): mixed      { return $this->fontProperty($row, $col, $sheet, 'font'); }

    // -------------------------------------------------------------------------
    // HTML dump
    // -------------------------------------------------------------------------

    public function dump(bool $row_numbers = false, bool $col_letters = false, int $sheet = 0, string $table_class = 'excel'): string
    {
        $out = "<table class=\"$table_class\" cellspacing=0>";

        if ($col_letters) {
            $out .= "<thead>\n\t<tr>";
            if ($row_numbers) {
                $out .= "\n\t\t<th>&nbsp</th>";
            }
            for ($i = 1; $i <= $this->colcount($sheet); $i++) {
                $style = "width:" . ($this->colwidth($i, $sheet) * 1) . "px;";
                if ($this->colhidden($i, $sheet)) $style .= "display:none;";
                $out .= "\n\t\t<th style=\"$style\">" . strtoupper($this->colindexes[$i]) . "</th>";
            }
            $out .= "</tr></thead>\n";
        }

        $out .= "<tbody>\n";
        for ($row = 1; $row <= $this->rowcount($sheet); $row++) {
            $rowheight = $this->rowheight($row, $sheet);
            $style = "height:" . ($rowheight * (4 / 3)) . "px;";
            if ($this->rowhidden($row, $sheet)) $style .= "display:none;";
            $out .= "\n\t<tr style=\"$style\">";
            if ($row_numbers) {
                $out .= "\n\t\t<th>$row</th>";
            }
            for ($col = 1; $col <= $this->colcount($sheet); $col++) {
                $rowspan = $this->rowspan($row, $col, $sheet);
                $colspan = $this->colspan($row, $col, $sheet);
                for ($i = 0; $i < $rowspan; $i++) {
                    for ($j = 0; $j < $colspan; $j++) {
                        if ($i > 0 || $j > 0) {
                            $this->sheets[$sheet]['cellsInfo'][$row + $i][$col + $j]['dontprint'] = 1;
                        }
                    }
                }
                if (!($this->sheets[$sheet]['cellsInfo'][$row][$col]['dontprint'] ?? false)) {
                    $style = $this->style($row, $col, $sheet);
                    if ($this->colhidden($col, $sheet)) $style .= "display:none;";
                    $out .= "\n\t\t<td style=\"$style\""
                        . ($colspan > 1 ? " colspan=$colspan" : "")
                        . ($rowspan > 1 ? " rowspan=$rowspan" : "") . ">";
                    $val = $this->val($row, $col, $sheet);
                    if ($val === '') {
                        $val = "&nbsp;";
                    } else {
                        $val = htmlentities($val);
                        $link = $this->hyperlink($row, $col, $sheet);
                        if ($link !== '') {
                            $val = "<a href=\"$link\">$val</a>";
                        }
                    }
                    $out .= "<nobr>" . nl2br($val) . "</nobr>";
                    $out .= "</td>";
                }
            }
            $out .= "</tr>\n";
        }
        $out .= "</tbody></table>";
        return $out;
    }

    // -------------------------------------------------------------------------
    // Settings
    // -------------------------------------------------------------------------

    public function setOutputEncoding(string $encoding): void
    {
        $this->_defaultEncoding = $encoding;
    }

    public function setUTFEncoder(string $encoder = 'iconv'): void
    {
        $this->_encoderFunction = '';
        if ($encoder === 'iconv') {
            $this->_encoderFunction = function_exists('iconv') ? 'iconv' : '';
        } elseif ($encoder === 'mb') {
            $this->_encoderFunction = function_exists('mb_convert_encoding') ? 'mb_convert_encoding' : '';
        }
    }

    public function setRowColOffset(int $iOffset): void
    {
        $this->_rowoffset = $iOffset;
        $this->_coloffset = $iOffset;
    }

    public function setDefaultFormat(string $sFormat): void
    {
        $this->_defaultFormat = $sFormat;
    }

    public function setColumnFormat(int $column, string $sFormat): void
    {
        $this->_columnsFormat[$column] = $sFormat;
    }

    // -------------------------------------------------------------------------
    // File reading
    // -------------------------------------------------------------------------

    public function read(string $sFileName): void
    {
        $res = $this->_ole->read($sFileName);
        if ($res === false) {
            if ($this->_ole->error == 1) {
                die('The filename ' . $sFileName . ' is not readable');
            }
        }
        $this->data = $this->_ole->getWorkBook();
        $this->_parse();
    }

    // -------------------------------------------------------------------------
    // Internal helpers
    // -------------------------------------------------------------------------

    private function myHex(int $d): string
    {
        return $d < 16 ? "0" . dechex($d) : dechex($d);
    }

    private function dumpHexData(string $data, int $pos, int $length): string
    {
        $info = "";
        for ($i = 0; $i <= $length; $i++) {
            $info .= ($i == 0 ? "" : " ") . $this->myHex(ord($data[$pos + $i])) . (ord($data[$pos + $i]) > 31 ? "[" . $data[$pos + $i] . "]" : '');
        }
        return $info;
    }

    private function getCol(int|string $col): int|string
    {
        if (is_string($col)) {
            $col = strtolower($col);
            if (array_key_exists($col, $this->colnames)) {
                $col = $this->colnames[$col];
            }
        }
        return $col;
    }

    private function _format_value(string $format, mixed $num, mixed $f): array
    {
        if ((!$f && $format === "%s") || ($f == 49) || ($format === "GENERAL")) {
            return ['string' => $num, 'formatColor' => null];
        }

        $parts = explode(";", $format);
        $pattern = $parts[0];

        if (count($parts) > 2 && $num == 0) {
            $pattern = $parts[2];
        }
        if (count($parts) > 1 && $num < 0) {
            $pattern = $parts[1];
            $num = abs($num);
        }

        $color = "";
        $matches = [];
        $color_regex = "/^\[(BLACK|BLUE|CYAN|GREEN|MAGENTA|RED|WHITE|YELLOW)\]/i";
        if (preg_match($color_regex, $pattern, $matches)) {
            $color = strtolower($matches[1]);
            $pattern = preg_replace($color_regex, "", $pattern);
        }

        $pattern = preg_replace("/_./", "", $pattern);
        $pattern = preg_replace("/\\\\/", "", $pattern);
        $pattern = preg_replace("/\"/", "", $pattern);
        $pattern = preg_replace("/\#/", "0", $pattern);

        $has_commas = (bool)preg_match("/,/", $pattern);
        if ($has_commas) {
            $pattern = preg_replace("/,/", "", $pattern);
        }

        if (preg_match("/\d(\%)([^\%]|$)/", $pattern, $matches)) {
            $num = $num * 100;
            $pattern = preg_replace("/(\d)(\%)([^\%]|$)/", "$1%$3", $pattern);
        }

        $number_regex = "/(\d+)(\.?)(\d*)/";
        if (preg_match($number_regex, $pattern, $matches)) {
            $left  = $matches[1];
            $dec   = $matches[2];
            $right = $matches[3];
            if ($has_commas) {
                $formatted = number_format((float)$num, strlen($right));
            } else {
                $sprintf_pattern = "%1." . strlen($right) . "f";
                $formatted = sprintf($sprintf_pattern, $num);
            }
            $pattern = preg_replace($number_regex, $formatted, $pattern);
        }

        return [
            'string'      => $pattern,
            'formatColor' => $color,
        ];
    }

    private function _parse(): bool
    {
        $pos  = 0;
        $data = $this->data;

        $code          = v($data, $pos);
        $length        = v($data, $pos + 2);
        $version       = v($data, $pos + 4);
        $substreamType = v($data, $pos + 6);

        $this->version = $version;

        if ($version !== SPREADSHEET_EXCEL_READER_BIFF8 && $version !== SPREADSHEET_EXCEL_READER_BIFF7) {
            return false;
        }
        if ($substreamType !== SPREADSHEET_EXCEL_READER_WORKBOOKGLOBALS) {
            return false;
        }

        $pos += $length + 4;

        $code   = v($data, $pos);
        $length = v($data, $pos + 2);

        while ($code !== SPREADSHEET_EXCEL_READER_TYPE_EOF) {
            switch ($code) {
                case SPREADSHEET_EXCEL_READER_TYPE_SST:
                    $spos          = $pos + 4;
                    $limitpos      = $spos + $length;
                    $uniqueStrings = $this->_GetInt4d($data, $spos + 4);
                    $spos += 8;

                    for ($i = 0; $i < $uniqueStrings; $i++) {
                        if ($spos == $limitpos) {
                            $opcode    = v($data, $spos);
                            $conlength = v($data, $spos + 2);
                            if ($opcode !== 0x3c) return false;
                            $spos += 4;
                            $limitpos = $spos + $conlength;
                        }

                        $numChars    = ord($data[$spos]) | (ord($data[$spos + 1]) << 8);
                        $spos += 2;
                        $optionFlags = ord($data[$spos]);
                        $spos++;

                        $asciiEncoding  = (($optionFlags & 0x01) == 0);
                        $extendedString = (($optionFlags & 0x04) != 0);
                        $richString     = (($optionFlags & 0x08) != 0);

                        $formattingRuns    = 0;
                        $extendedRunLength = 0;

                        if ($richString) {
                            $formattingRuns = v($data, $spos);
                            $spos += 2;
                        }
                        if ($extendedString) {
                            $extendedRunLength = $this->_GetInt4d($data, $spos);
                            $spos += 4;
                        }

                        $len = $asciiEncoding ? $numChars : $numChars * 2;

                        if ($spos + $len < $limitpos) {
                            $retstr = substr($data, $spos, $len);
                            $spos += $len;
                        } else {
                            $retstr    = substr($data, $spos, $limitpos - $spos);
                            $bytesRead = $limitpos - $spos;
                            $charsLeft = $numChars - ($asciiEncoding ? $bytesRead : $bytesRead / 2);
                            $spos      = $limitpos;

                            while ($charsLeft > 0) {
                                $opcode    = v($data, $spos);
                                $conlength = v($data, $spos + 2);
                                if ($opcode !== 0x3c) return false;
                                $spos += 4;
                                $limitpos = $spos + $conlength;
                                $option   = ord($data[$spos]);
                                $spos++;

                                if ($asciiEncoding && $option == 0) {
                                    $len = min($charsLeft, $limitpos - $spos);
                                    $retstr .= substr($data, $spos, $len);
                                    $charsLeft -= $len;
                                } elseif (!$asciiEncoding && $option != 0) {
                                    $len = min($charsLeft * 2, $limitpos - $spos);
                                    $retstr .= substr($data, $spos, $len);
                                    $charsLeft -= $len / 2;
                                } elseif (!$asciiEncoding && $option == 0) {
                                    $len = min($charsLeft, $limitpos - $spos);
                                    for ($j = 0; $j < $len; $j++) {
                                        $retstr .= $data[$spos + $j] . chr(0);
                                    }
                                    $charsLeft -= $len;
                                } else {
                                    $newstr = '';
                                    for ($j = 0; $j < strlen($retstr); $j++) {
                                        $newstr .= $retstr[$j] . chr(0);
                                    }
                                    $retstr = $newstr;
                                    $len = min($charsLeft * 2, $limitpos - $spos);
                                    $retstr .= substr($data, $spos, $len);
                                    $charsLeft -= $len / 2;
                                    $asciiEncoding = false;
                                }
                                $spos += $len;
                            }
                        }

                        $retstr = $asciiEncoding ? $retstr : $this->_encodeUTF16($retstr);

                        if ($richString) {
                            $spos += 4 * $formattingRuns;
                        }
                        if ($extendedString) {
                            $spos += $extendedRunLength;
                        }
                        $this->sst[] = $retstr;
                    }
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_FILEPASS:
                    return false;

                case SPREADSHEET_EXCEL_READER_TYPE_NAME:
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_FORMAT:
                    $indexCode = v($data, $pos + 4);
                    if ($version == SPREADSHEET_EXCEL_READER_BIFF8) {
                        $numchars = v($data, $pos + 6);
                        if (ord($data[$pos + 8]) == 0) {
                            $formatString = substr($data, $pos + 9, $numchars);
                        } else {
                            $formatString = substr($data, $pos + 9, $numchars * 2);
                        }
                    } else {
                        $numchars     = ord($data[$pos + 6]);
                        $formatString = substr($data, $pos + 7, $numchars * 2);
                    }
                    $this->formatRecords[$indexCode] = $formatString;
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_FONT:
                    $height = v($data, $pos + 4);
                    $option = v($data, $pos + 6);
                    $color  = v($data, $pos + 8);
                    $weight = v($data, $pos + 10);
                    $under  = ord($data[$pos + 14]);
                    $font   = "";
                    $numchars = ord($data[$pos + 18]);
                    if ((ord($data[$pos + 19]) & 1) == 0) {
                        $font = substr($data, $pos + 20, $numchars);
                    } else {
                        $font = $this->_encodeUTF16(substr($data, $pos + 20, $numchars * 2));
                    }
                    $this->fontRecords[] = [
                        'height' => $height / 20,
                        'italic' => (bool)($option & 2),
                        'color'  => $color,
                        'under'  => ($under != 0),
                        'bold'   => ($weight == 700),
                        'font'   => $font,
                        'raw'    => $this->dumpHexData($data, $pos + 3, $length),
                    ];
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_PALETTE:
                    $colors = ord($data[$pos + 4]) | ord($data[$pos + 5]) << 8;
                    for ($coli = 0; $coli < $colors; $coli++) {
                        $colOff = $pos + 2 + ($coli * 4);
                        $colr   = ord($data[$colOff]);
                        $colg   = ord($data[$colOff + 1]);
                        $colb   = ord($data[$colOff + 2]);
                        $this->colors[0x07 + $coli] = '#' . $this->myHex($colr) . $this->myHex($colg) . $this->myHex($colb);
                    }
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_XF:
                    $fontIndexCode = (ord($data[$pos + 4]) | ord($data[$pos + 5]) << 8) - 1;
                    $fontIndexCode = max(0, $fontIndexCode);
                    $indexCode     = ord($data[$pos + 6]) | ord($data[$pos + 7]) << 8;
                    $alignbit      = ord($data[$pos + 10]) & 3;
                    $bgi           = (ord($data[$pos + 22]) | ord($data[$pos + 23]) << 8) & 0x3FFF;
                    $bgcolor       = ($bgi & 0x7F);
                    $align         = "";

                    if ($alignbit == 3) $align = "right";
                    if ($alignbit == 2) $align = "center";

                    $fillPattern = (ord($data[$pos + 21]) & 0xFC) >> 2;
                    if ($fillPattern == 0) $bgcolor = "";

                    $xf                 = [];
                    $xf['formatIndex']  = $indexCode;
                    $xf['align']        = $align;
                    $xf['fontIndex']    = $fontIndexCode;
                    $xf['bgColor']      = $bgcolor;
                    $xf['fillPattern']  = $fillPattern;

                    $border = ord($data[$pos + 14]) | (ord($data[$pos + 15]) << 8) | (ord($data[$pos + 16]) << 16) | (ord($data[$pos + 17]) << 24);
                    $xf['borderLeft']   = $this->lineStyles[($border & 0xF)];
                    $xf['borderRight']  = $this->lineStyles[($border & 0xF0) >> 4];
                    $xf['borderTop']    = $this->lineStyles[($border & 0xF00) >> 8];
                    $xf['borderBottom'] = $this->lineStyles[($border & 0xF000) >> 12];

                    $xf['borderLeftColor']   = ($border & 0x7F0000) >> 16;
                    $xf['borderRightColor']  = ($border & 0x3F800000) >> 23;
                    $border = ord($data[$pos + 18]) | ord($data[$pos + 19]) << 8;
                    $xf['borderTopColor']    = ($border & 0x7F);
                    $xf['borderBottomColor'] = ($border & 0x3F80) >> 7;

                    if (array_key_exists($indexCode, $this->dateFormats)) {
                        $xf['type']   = 'date';
                        $xf['format'] = $this->dateFormats[$indexCode];
                        if ($align === '') $xf['align'] = 'right';
                    } elseif (array_key_exists($indexCode, $this->numberFormats)) {
                        $xf['type']   = 'number';
                        $xf['format'] = $this->numberFormats[$indexCode];
                        if ($align === '') $xf['align'] = 'right';
                    } else {
                        $isdate    = false;
                        $formatstr = '';
                        if ($indexCode > 0) {
                            if (isset($this->formatRecords[$indexCode])) {
                                $formatstr = $this->formatRecords[$indexCode];
                            }
                            if ($formatstr !== "") {
                                $tmp = preg_replace("/\;.*/", "", $formatstr);
                                $tmp = preg_replace("/^\[[^\]]*\]/", "", $tmp);
                                if (preg_match("/[^hmsday\/\-:\s\\\,AMP]/i", $tmp) == 0) {
                                    $isdate    = true;
                                    $formatstr = $tmp;
                                    $formatstr = str_replace(['AM/PM', 'mmmm', 'mmm'], ['a', 'F', 'M'], $formatstr);
                                    $formatstr = preg_replace("/(h:?)mm?/", "$1i", $formatstr);
                                    $formatstr = preg_replace("/mm?(:?s)/", "i$1", $formatstr);
                                    $formatstr = preg_replace("/(^|[^m])m([^m]|$)/", '$1n$2', $formatstr);
                                    $formatstr = preg_replace("/(^|[^m])m([^m]|$)/", '$1n$2', $formatstr);
                                    $formatstr = str_replace('mm', 'm', $formatstr);
                                    $formatstr = preg_replace("/(^|[^d])d([^d]|$)/", '$1j$2', $formatstr);
                                    $formatstr = str_replace(['dddd', 'ddd', 'dd', 'yyyy', 'yy', 'hh', 'h'], ['l', 'D', 'd', 'Y', 'y', 'H', 'g'], $formatstr);
                                    $formatstr = preg_replace("/ss?/", 's', $formatstr);
                                }
                            }
                        }
                        if ($isdate) {
                            $xf['type']   = 'date';
                            $xf['format'] = $formatstr;
                            if ($align === '') $xf['align'] = 'right';
                        } else {
                            if (preg_match("/[0#]/", $formatstr)) {
                                $xf['type'] = 'number';
                                if ($align === '') $xf['align'] = 'right';
                            } else {
                                $xf['type'] = 'other';
                            }
                            $xf['format'] = $formatstr;
                            $xf['code']   = $indexCode;
                        }
                    }
                    $this->xfRecords[] = $xf;
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_NINETEENFOUR:
                    $this->nineteenFour = (ord($data[$pos + 4]) == 1);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_BOUNDSHEET:
                    $rec_offset          = $this->_GetInt4d($data, $pos + 4);
                    $rec_typeFlag        = ord($data[$pos + 8]);
                    $rec_visibilityFlag  = ord($data[$pos + 9]);
                    $rec_length          = ord($data[$pos + 10]);

                    if ($version == SPREADSHEET_EXCEL_READER_BIFF8) {
                        $chartype = ord($data[$pos + 11]);
                        if ($chartype == 0) {
                            $rec_name = substr($data, $pos + 12, $rec_length);
                        } else {
                            $rec_name = $this->_encodeUTF16(substr($data, $pos + 12, $rec_length * 2));
                        }
                    } else {
                        $rec_name = substr($data, $pos + 11, $rec_length);
                    }
                    $this->boundsheets[] = ['name' => $rec_name, 'offset' => $rec_offset];
                    break;
            }

            $pos += $length + 4;
            $code   = ord($data[$pos]) | ord($data[$pos + 1]) << 8;
            $length = ord($data[$pos + 2]) | ord($data[$pos + 3]) << 8;
        }

        foreach ($this->boundsheets as $key => $val) {
            $this->sn = $key;
            $this->_parsesheet($val['offset']);
        }
        return true;
    }

    private function _parsesheet(int $spos): int
    {
        $data          = $this->data;
        $code          = ord($data[$spos]) | ord($data[$spos + 1]) << 8;
        $length        = ord($data[$spos + 2]) | ord($data[$spos + 3]) << 8;
        $version       = ord($data[$spos + 4]) | ord($data[$spos + 5]) << 8;
        $substreamType = ord($data[$spos + 6]) | ord($data[$spos + 7]) << 8;

        if ($version !== SPREADSHEET_EXCEL_READER_BIFF8 && $version !== SPREADSHEET_EXCEL_READER_BIFF7) {
            return -1;
        }
        if ($substreamType !== SPREADSHEET_EXCEL_READER_WORKSHEET) {
            return -2;
        }

        $spos += $length + 4;

        $previousRow = 0;
        $previousCol = 0;
        $cont        = true;

        while ($cont) {
            $lowcode = ord($data[$spos]);
            if ($lowcode == SPREADSHEET_EXCEL_READER_TYPE_EOF) break;

            $code   = $lowcode | ord($data[$spos + 1]) << 8;
            $length = ord($data[$spos + 2]) | ord($data[$spos + 3]) << 8;
            $spos  += 4;

            $this->sheets[$this->sn]['maxrow'] = $this->_rowoffset - 1;
            $this->sheets[$this->sn]['maxcol'] = $this->_coloffset - 1;

            switch ($code) {
                case SPREADSHEET_EXCEL_READER_TYPE_DIMENSION:
                    if (!isset($this->sheets[$this->sn]['numRows'])) {
                        if ($length == 10 || $version == SPREADSHEET_EXCEL_READER_BIFF7) {
                            $this->sheets[$this->sn]['numRows'] = ord($data[$spos + 2]) | ord($data[$spos + 3]) << 8;
                            $this->sheets[$this->sn]['numCols'] = ord($data[$spos + 6]) | ord($data[$spos + 7]) << 8;
                        } else {
                            $this->sheets[$this->sn]['numRows'] = ord($data[$spos + 4]) | ord($data[$spos + 5]) << 8;
                            $this->sheets[$this->sn]['numCols'] = ord($data[$spos + 10]) | ord($data[$spos + 11]) << 8;
                        }
                    }
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_MERGEDCELLS:
                    $cellRanges = ord($data[$spos]) | ord($data[$spos + 1]) << 8;
                    for ($i = 0; $i < $cellRanges; $i++) {
                        $fr = ord($data[$spos + 8 * $i + 2]) | ord($data[$spos + 8 * $i + 3]) << 8;
                        $lr = ord($data[$spos + 8 * $i + 4]) | ord($data[$spos + 8 * $i + 5]) << 8;
                        $fc = ord($data[$spos + 8 * $i + 6]) | ord($data[$spos + 8 * $i + 7]) << 8;
                        $lc = ord($data[$spos + 8 * $i + 8]) | ord($data[$spos + 8 * $i + 9]) << 8;
                        if ($lr - $fr > 0) {
                            $this->sheets[$this->sn]['cellsInfo'][$fr + 1][$fc + 1]['rowspan'] = $lr - $fr + 1;
                        }
                        if ($lc - $fc > 0) {
                            $this->sheets[$this->sn]['cellsInfo'][$fr + 1][$fc + 1]['colspan'] = $lc - $fc + 1;
                        }
                    }
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_RK:
                case SPREADSHEET_EXCEL_READER_TYPE_RK2:
                    $row      = ord($data[$spos]) | ord($data[$spos + 1]) << 8;
                    $column   = ord($data[$spos + 2]) | ord($data[$spos + 3]) << 8;
                    $rknum    = $this->_GetInt4d($data, $spos + 6);
                    $numValue = $this->_GetIEEE754($rknum);
                    $info     = $this->_getCellDetails($spos, $numValue, $column);
                    $this->addcell($row, $column, $info['string'], $info);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_LABELSST:
                    $row     = ord($data[$spos]) | ord($data[$spos + 1]) << 8;
                    $column  = ord($data[$spos + 2]) | ord($data[$spos + 3]) << 8;
                    $xfindex = ord($data[$spos + 4]) | ord($data[$spos + 5]) << 8;
                    $index   = $this->_GetInt4d($data, $spos + 6);
                    $this->addcell($row, $column, $this->sst[$index], ['xfIndex' => $xfindex]);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_MULRK:
                    $row      = ord($data[$spos]) | ord($data[$spos + 1]) << 8;
                    $colFirst = ord($data[$spos + 2]) | ord($data[$spos + 3]) << 8;
                    $colLast  = ord($data[$spos + $length - 2]) | ord($data[$spos + $length - 1]) << 8;
                    $columns  = $colLast - $colFirst + 1;
                    $tmppos   = $spos + 4;
                    for ($i = 0; $i < $columns; $i++) {
                        $numValue = $this->_GetIEEE754($this->_GetInt4d($data, $tmppos + 2));
                        $info     = $this->_getCellDetails($tmppos - 4, $numValue, $colFirst + $i + 1);
                        $tmppos  += 6;
                        $this->addcell($row, $colFirst + $i, $info['string'], $info);
                    }
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_NUMBER:
                    $row    = ord($data[$spos]) | ord($data[$spos + 1]) << 8;
                    $column = ord($data[$spos + 2]) | ord($data[$spos + 3]) << 8;
                    $tmp    = unpack("ddouble", substr($data, $spos + 6, 8));
                    $numValue = $this->isDate($spos) ? $tmp['double'] : $this->createNumber($spos);
                    $info   = $this->_getCellDetails($spos, $numValue, $column);
                    $this->addcell($row, $column, $info['string'], $info);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_FORMULA:
                case SPREADSHEET_EXCEL_READER_TYPE_FORMULA2:
                    $row    = ord($data[$spos]) | ord($data[$spos + 1]) << 8;
                    $column = ord($data[$spos + 2]) | ord($data[$spos + 3]) << 8;
                    if ((ord($data[$spos + 6]) == 0) && (ord($data[$spos + 12]) == 255) && (ord($data[$spos + 13]) == 255)) {
                        $previousRow = $row;
                        $previousCol = $column;
                    } elseif ((ord($data[$spos + 6]) == 1) && (ord($data[$spos + 12]) == 255) && (ord($data[$spos + 13]) == 255)) {
                        $this->addcell($row, $column, ord($this->data[$spos + 8]) == 1 ? "TRUE" : "FALSE");
                    } elseif ((ord($data[$spos + 6]) == 2) && (ord($data[$spos + 12]) == 255) && (ord($data[$spos + 13]) == 255)) {
                        // Error formula – ignored
                    } elseif ((ord($data[$spos + 6]) == 3) && (ord($data[$spos + 12]) == 255) && (ord($data[$spos + 13]) == 255)) {
                        $this->addcell($row, $column, '');
                    } else {
                        $tmp      = unpack("ddouble", substr($data, $spos + 6, 8));
                        $numValue = $this->isDate($spos) ? $tmp['double'] : $this->createNumber($spos);
                        $info     = $this->_getCellDetails($spos, $numValue, $column);
                        $this->addcell($row, $column, $info['string'], $info);
                    }
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_BOOLERR:
                    $row    = ord($data[$spos]) | ord($data[$spos + 1]) << 8;
                    $column = ord($data[$spos + 2]) | ord($data[$spos + 3]) << 8;
                    $this->addcell($row, $column, (string)ord($data[$spos + 6]));
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_STRING:
                    if ($version == SPREADSHEET_EXCEL_READER_BIFF8) {
                        $xpos          = $spos;
                        $numChars      = ord($data[$xpos]) | (ord($data[$xpos + 1]) << 8);
                        $xpos += 2;
                        $optionFlags   = ord($data[$xpos]);
                        $xpos++;
                        $asciiEncoding = (($optionFlags & 0x01) == 0);
                        $extendedString = (($optionFlags & 0x04) != 0);
                        $richString    = (($optionFlags & 0x08) != 0);

                        $formattingRuns    = 0;
                        $extendedRunLength = 0;

                        if ($richString) {
                            $formattingRuns = ord($data[$xpos]) | (ord($data[$xpos + 1]) << 8);
                            $xpos += 2;
                        }
                        if ($extendedString) {
                            $extendedRunLength = $this->_GetInt4d($this->data, $xpos);
                            $xpos += 4;
                        }
                        $len    = $asciiEncoding ? $numChars : $numChars * 2;
                        $retstr = substr($data, $xpos, $len);
                        $xpos  += $len;
                        $retstr = $asciiEncoding ? $retstr : $this->_encodeUTF16($retstr);
                    } else {
                        $xpos     = $spos;
                        $numChars = ord($data[$xpos]) | (ord($data[$xpos + 1]) << 8);
                        $xpos += 2;
                        $retstr   = substr($data, $xpos, $numChars);
                    }
                    $this->addcell($previousRow, $previousCol, $retstr);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_ROW:
                    $row     = ord($data[$spos]) | ord($data[$spos + 1]) << 8;
                    $rowInfo = ord($data[$spos + 6]) | ((ord($data[$spos + 7]) << 8) & 0x7FFF);
                    $rowHeight = ($rowInfo & 0x8000) > 0 ? -1 : ($rowInfo & 0x7FFF);
                    $rowHidden = (ord($data[$spos + 12]) & 0x20) >> 5;
                    $this->rowInfo[$this->sn][$row + 1] = ['height' => $rowHeight / 20, 'hidden' => $rowHidden];
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_DBCELL:
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_MULBLANK:
                    $row    = ord($data[$spos]) | ord($data[$spos + 1]) << 8;
                    $column = ord($data[$spos + 2]) | ord($data[$spos + 3]) << 8;
                    $cols   = ($length / 2) - 3;
                    for ($c = 0; $c < $cols; $c++) {
                        $xfindex = ord($data[$spos + 4 + ($c * 2)]) | ord($data[$spos + 5 + ($c * 2)]) << 8;
                        $this->addcell($row, $column + $c, "", ['xfIndex' => $xfindex]);
                    }
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_LABEL:
                    $row    = ord($data[$spos]) | ord($data[$spos + 1]) << 8;
                    $column = ord($data[$spos + 2]) | ord($data[$spos + 3]) << 8;
                    $this->addcell($row, $column, substr($data, $spos + 8, ord($data[$spos + 6]) | ord($data[$spos + 7]) << 8));
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_EOF:
                    $cont = false;
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_HYPER:
                    $row     = ord($this->data[$spos]) | ord($this->data[$spos + 1]) << 8;
                    $row2    = ord($this->data[$spos + 2]) | ord($this->data[$spos + 3]) << 8;
                    $column  = ord($this->data[$spos + 4]) | ord($this->data[$spos + 5]) << 8;
                    $column2 = ord($this->data[$spos + 6]) | ord($this->data[$spos + 7]) << 8;
                    $linkdata = [];
                    $flags   = ord($this->data[$spos + 28]);
                    $udesc   = "";
                    $ulink   = "";
                    $uloc    = 32;
                    $linkdata['flags'] = $flags;
                    if (($flags & 1) > 0) {
                        if (($flags & 0x14) == 0x14) {
                            $uloc   += 4;
                            $descLen = ord($this->data[$spos + 32]) | ord($this->data[$spos + 33]) << 8;
                            $udesc   = substr($this->data, $spos + $uloc, $descLen * 2);
                            $uloc   += 2 * $descLen;
                        }
                        $ulink = $this->read16bitstring($this->data, $spos + $uloc + 20);
                        if ($udesc === "") $udesc = $ulink;
                    }
                    $linkdata['desc'] = $udesc;
                    $linkdata['link'] = $this->_encodeUTF16($ulink);
                    for ($r = $row; $r <= $row2; $r++) {
                        for ($c = $column; $c <= $column2; $c++) {
                            $this->sheets[$this->sn]['cellsInfo'][$r + 1][$c + 1]['hyperlink'] = $linkdata;
                        }
                    }
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_DEFCOLWIDTH:
                    $this->defaultColWidth = ord($data[$spos + 4]) | ord($data[$spos + 5]) << 8;
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_STANDARDWIDTH:
                    $this->standardColWidth = ord($data[$spos + 4]) | ord($data[$spos + 5]) << 8;
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_COLINFO:
                    $colfrom = ord($data[$spos + 0]) | ord($data[$spos + 1]) << 8;
                    $colto   = ord($data[$spos + 2]) | ord($data[$spos + 3]) << 8;
                    $cw      = ord($data[$spos + 4]) | ord($data[$spos + 5]) << 8;
                    $cxf     = ord($data[$spos + 6]) | ord($data[$spos + 7]) << 8;
                    $co      = ord($data[$spos + 8]);
                    for ($coli = $colfrom; $coli <= $colto; $coli++) {
                        $this->colInfo[$this->sn][$coli + 1] = [
                            'width'     => $cw,
                            'xf'        => $cxf,
                            'hidden'    => ($co & 0x01),
                            'collapsed' => ($co & 0x1000) >> 12,
                        ];
                    }
                    break;

                default:
                    break;
            }
            $spos += $length;
        }

        if (!isset($this->sheets[$this->sn]['numRows'])) {
            $this->sheets[$this->sn]['numRows'] = $this->sheets[$this->sn]['maxrow'];
        }
        if (!isset($this->sheets[$this->sn]['numCols'])) {
            $this->sheets[$this->sn]['numCols'] = $this->sheets[$this->sn]['maxcol'];
        }

        return 0;
    }

    public function isDate(int $spos): bool
    {
        $xfindex = ord($this->data[$spos + 4]) | ord($this->data[$spos + 5]) << 8;
        return ($this->xfRecords[$xfindex]['type'] ?? '') === 'date';
    }

    private function _getCellDetails(int $spos, mixed $numValue, int $column): array
    {
        $xfindex  = ord($this->data[$spos + 4]) | ord($this->data[$spos + 5]) << 8;
        $xfrecord = $this->xfRecords[$xfindex];
        $type     = $xfrecord['type'];
        $format   = $xfrecord['format'];
        $formatIndex = $xfrecord['formatIndex'];
        $fontIndex   = $xfrecord['fontIndex'];
        $formatColor = "";
        $string      = '';
        $raw         = '';
        $rectype     = '';

        if (isset($this->_columnsFormat[$column + 1])) {
            $format = $this->_columnsFormat[$column + 1];
        }

        if ($type === 'date') {
            $rectype  = 'date';
            $utcDays  = floor($numValue - ($this->nineteenFour ? SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS1904 : SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS));
            $utcValue = $utcDays * SPREADSHEET_EXCEL_READER_MSINADAY;
            $dateinfo = gmgetdate((int)$utcValue);
            $raw      = $numValue;

            $fractionalDay = $numValue - floor($numValue) + .0000001;
            $totalseconds  = (int)floor(SPREADSHEET_EXCEL_READER_MSINADAY * $fractionalDay);
            $secs          = $totalseconds % 60;
            $totalseconds -= $secs;
            $hours         = (int)floor($totalseconds / (60 * 60));
            $mins          = (int)floor($totalseconds / 60) % 60;
            $string        = date($format, mktime($hours, $mins, $secs, $dateinfo["mon"], $dateinfo["mday"], $dateinfo["year"]));
        } elseif ($type === 'number') {
            $rectype     = 'number';
            $formatted   = $this->_format_value($format, $numValue, $formatIndex);
            $string      = $formatted['string'];
            $formatColor = $formatted['formatColor'];
            $raw         = $numValue;
        } else {
            if ($format === "") $format = $this->_defaultFormat;
            $rectype     = 'unknown';
            $formatted   = $this->_format_value($format, $numValue, $formatIndex);
            $string      = $formatted['string'];
            $formatColor = $formatted['formatColor'];
            $raw         = $numValue;
        }

        return [
            'string'      => $string,
            'raw'         => $raw,
            'rectype'     => $rectype,
            'format'      => $format,
            'formatIndex' => $formatIndex,
            'fontIndex'   => $fontIndex,
            'formatColor' => $formatColor,
            'xfIndex'     => $xfindex,
        ];
    }

    public function createNumber(int $spos): float
    {
        $rknumhigh   = $this->_GetInt4d($this->data, $spos + 10);
        $rknumlow    = $this->_GetInt4d($this->data, $spos + 6);
        $sign        = ($rknumhigh & 0x80000000) >> 31;
        $exp         = ($rknumhigh & 0x7ff00000) >> 20;
        $mantissa    = (0x100000 | ($rknumhigh & 0x000fffff));
        $mantissalow1 = ($rknumlow & 0x80000000) >> 31;
        $mantissalow2 = ($rknumlow & 0x7fffffff);
        $value        = $mantissa / pow(2, (20 - ($exp - 1023)));
        if ($mantissalow1 != 0) $value += 1 / pow(2, (21 - ($exp - 1023)));
        $value += $mantissalow2 / pow(2, (52 - ($exp - 1023)));
        if ($sign) $value = -1 * $value;
        return $value;
    }

    public function addcell(int $row, int $col, string $string, ?array $info = null): void
    {
        $this->sheets[$this->sn]['maxrow'] = max($this->sheets[$this->sn]['maxrow'] ?? 0, $row + $this->_rowoffset);
        $this->sheets[$this->sn]['maxcol'] = max($this->sheets[$this->sn]['maxcol'] ?? 0, $col + $this->_coloffset);
        $this->sheets[$this->sn]['cells'][$row + $this->_rowoffset][$col + $this->_coloffset] = $string;
        if ($this->store_extended_info && $info) {
            foreach ($info as $key => $val) {
                $this->sheets[$this->sn]['cellsInfo'][$row + $this->_rowoffset][$col + $this->_coloffset][$key] = $val;
            }
        }
    }

    private function _GetIEEE754(int $rknum): float
    {
        if (($rknum & 0x02) != 0) {
            $value = $rknum >> 2;
        } else {
            $sign     = ($rknum & 0x80000000) >> 31;
            $exp      = ($rknum & 0x7ff00000) >> 20;
            $mantissa = (0x100000 | ($rknum & 0x000ffffc));
            $value    = $mantissa / pow(2, (20 - ($exp - 1023)));
            if ($sign) $value = -1 * $value;
        }
        if (($rknum & 0x01) != 0) {
            $value /= 100;
        }
        return (float)$value;
    }

    private function _encodeUTF16(string $string): string
    {
        $result = $string;
        if ($this->_defaultEncoding) {
            switch ($this->_encoderFunction) {
                case 'iconv':
                    $result = iconv('UTF-16LE', $this->_defaultEncoding, $string);
                    break;
                case 'mb_convert_encoding':
                    $result = mb_convert_encoding($string, $this->_defaultEncoding, 'UTF-16LE');
                    break;
            }
        }
        return $result;
    }

    private function _GetInt4d(string $data, int $pos): int
    {
        $value = ord($data[$pos]) | (ord($data[$pos + 1]) << 8) | (ord($data[$pos + 2]) << 16) | (ord($data[$pos + 3]) << 24);
        if ($value >= 4294967294) {
            $value = -2;
        }
        return $value;
    }

    private function read16bitstring(string $data, int $start): string
    {
        $len = 0;
        while (ord($data[$start + $len]) + ord($data[$start + $len + 1]) > 0) {
            $len++;
        }
        return substr($data, $start, $len);
    }
}
