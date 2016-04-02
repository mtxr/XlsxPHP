namespace XlsxPhp;

use XlsxPhp\BufferedWriter;

class Writer
{
    //------------------------------------------------------------------
    //http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
    const EXCEL_2007_MAX_ROW = 1048576;
    const EXCEL_2007_MAX_COL = 16384;
    const XLSXPHP            = "XLSXPHP";
    //------------------------------------------------------------------
    protected author              = "XLSXPHP";
    protected sheets              = [];
    protected shared_strings      = [];//unique set
    protected shared_string_count = 0;      //count of non-unique references to the unique set
    protected shared_string_count_seq    = 0;
    protected temp_files          = [];
    protected cell_formats        = [];//contains excel format like YYYY-MM-DD HH:MM:SS
    protected cell_types          = [];//contains friendly format like datetime
    protected tmp_prefix          = "xlsxphp_";
    protected sheet_count         = 0;
    protected format_count        = 0;
    protected cell_position         = 0;
    protected bufferSize          = 8191;

    protected current_sheet = "";

    public function __construct(bufferSize = 8191)
    {
        if (!class_exists("ZipArchive")) {
            throw new \Exception("ZipArchive not found");
        }
        let this->bufferSize = bufferSize;
        this->addCellFormat("GENERAL");
    }

    public function setAuthor(string author = self::XLSXPHP)
    {
        let this->author = author;
    }

    public function __destruct()
    {
        var temp_file;
        if (!empty this->temp_files) {
            for temp_file in this->temp_files {
                unlink(temp_file);
            }
        }
    }

    protected function tempFilename() -> string
    {
        var filename;
        let filename = tempnam(sys_get_temp_dir(), this->tmp_prefix);
        let this->temp_files[] = filename;
        return filename;
    }

    public function writeToStdOut()
    {
        var temp_file;
        let temp_file = this->tempFilename();
        self::writeToFile(temp_file);
        readfile(temp_file);
    }

    public function writeToString()
    {
        var temp_file;
        let temp_file = this->tempFilename();
        self::writeToFile(temp_file);
        return file_get_contents(temp_file);
    }

    public function writeToFile(filename)
    {
        var sheet_name, sheet;
        for sheet_name, sheet in this->sheets {
            self::finalizeSheet(sheet_name);//making sure all footers have been written
        }

        if ( file_exists( filename ) ) {
            if ( is_writable( filename ) ) {
                unlink( filename ); //if the zip already exists, remove it
            } else {
                self::log( "Error in " . __CLASS__ . "::" . __FUNCTION__ . ", file is not writeable." );
                return;
            }
        }
        var zip;
        let zip = new \ZipArchive();
        if (empty this->sheets) {
            self::log("Error in ".__CLASS__."::".__FUNCTION__.", no worksheets defined.");
            return;
        }
        if (!zip->open(filename, \ZipArchive::CREATE)) {
            self::log("Error in ".__CLASS__."::".__FUNCTION__.", unable to create zip.");
            return;
        }

        zip->addEmptyDir("docProps/");
        zip->addFromString("docProps/app.xml" , self::buildAppXML() );
        zip->addFromString("docProps/core.xml", self::buildCoreXML());

        zip->addEmptyDir("_rels/");
        zip->addFromString("_rels/.rels", self::buildRelationshipsXML());

        zip->addEmptyDir("xl/worksheets/");
        for sheet in this->sheets {
            zip->addFile(sheet["filename"], "xl/worksheets/" . sheet["xmlname"]);
        }
        if (!empty this->shared_strings) {
            zip->addFile(this->writeSharedStringsXML(), "xl/sharedStrings.xml");  //zip->addFromString("xl/sharedStrings.xml",     self::buildSharedStringsXML() );
        }
        zip->addFromString("xl/workbook.xml", self::buildWorkbookXML());
        zip->addFile(this->writeStylesXML(), "xl/styles.xml");  //zip->addFromString("xl/styles.xml"           , self::buildStylesXML() );
        zip->addFromString("[Content_Types].xml", self::buildContentTypesXML());

        zip->addEmptyDir("xl/_rels/");
        zip->addFromString("xl/_rels/workbook.xml.rels", self::buildWorkbookRelsXML() );
        zip->close();
    }

    protected function initializeSheet(sheet_name)
    {
        //if already initialized
        if (this->current_sheet==sheet_name || isset this->sheets[sheet_name]) {
            return;
        }
        var sheet, sheet_filename, sheet_xmlname, tabselected, max_cell;
        let sheet_filename = this->tempFilename();
        let sheet_xmlname = "sheet" . (this->sheet_count + 1).".xml";
        let sheet = [];
        let sheet["filename"]           = sheet_filename;
        let sheet["sheetname"]          = sheet_name;
        let sheet["xmlname"]            = sheet_xmlname;
        let sheet["row_count"]          = 0;
        let sheet["file_writer"]        = new \XlsxPhp\BufferedWriter(sheet_filename, "w", false, this->bufferSize);
        let sheet["columns"]            = [];
        let sheet["columns_count"]      = 0;
        let sheet["merge_cells"]        = [];
        let sheet["max_cell_tag_start"] = 0;
        let sheet["max_cell_tag_end"]   = 0;
        let sheet["finalized"]          = false;

        let tabselected = this->sheet_count == 1 ? "true" : "false";//only first sheet is selected
        let max_cell    = Writer::xlsCell(self::EXCEL_2007_MAX_ROW, self::EXCEL_2007_MAX_COL);//XFE1048577
        sheet["file_writer"]->write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
            ->write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">")
            ->write(  "<sheetPr filterMode=\"false\">")
            ->write(    "<pageSetUpPr fitToPage=\"false\"/>")
            ->write(  "</sheetPr>");
        let sheet["max_cell_tag_start"] = sheet["file_writer"]->ftell();
        sheet["file_writer"]->write("<dimension ref=\"A1:" . max_cell . "\"/>");
        let sheet["max_cell_tag_end"] = sheet["file_writer"]->ftell();
        sheet["file_writer"]->write(  "<sheetViews>")
            ->write(    "<sheetView colorId=\"64\" defaultGridColor=\"true\" rightToLeft=\"false\" showFormulas=\"false\" showGridLines=\"true\" showOutlineSymbols=\"true\" showRowColHeaders=\"true\" showZeros=\"true\" tabSelected=\"" . tabselected . "\" topLeftCell=\"A1\" view=\"normal\" windowProtection=\"false\" workbookViewId=\"0\" zoomScale=\"100\" zoomScaleNormal=\"100\" zoomScalePageLayoutView=\"100\">")
            ->write(      "<selection activeCell=\"A1\" activeCellId=\"0\" pane=\"topLeft\" sqref=\"A1\"/>")
            ->write(    "</sheetView>")
            ->write(  "</sheetViews>")
            ->write(  "<cols>")
            ->write(    "<col collapsed=\"false\" hidden=\"false\" max=\"1025\" min=\"1\" style=\"0\" width=\"11.5\"/>")
            ->write(  "</cols>")
            ->write(  "<sheetData>");

        let this->sheets[sheet_name] = sheet;
        let this->sheet_count = this->sheet_count + 1;
    }

    private function determineCellType(cell_format)
    {
        let cell_format = str_replace("[RED]", "", cell_format);
        if (cell_format=="GENERAL") {
            return "string";
        }
        if (cell_format=="0") {
            return "numeric";
        }
        if (preg_match("/[H]{1,2}:[M]{1,2}/", cell_format)) {
            return "datetime";
        }
        if (preg_match("/[M]{1,2}:[S]{1,2}/", cell_format)) {
            return "datetime";
        }
        if (preg_match("/[YY]{2,4}/", cell_format)) {
            return "date";
        }
        if (preg_match("/[D]{1,2}/", cell_format)) {
            return "date";
        }
        if (preg_match("/[M]{1,2}/", cell_format)) {
            return "date";
        }
        if (preg_match("/$/", cell_format)) {
            return "currency";
        }
        if (preg_match("/%/", cell_format)) {
            return "percent";
        }
        if (preg_match("/0/", cell_format)) {
            return "numeric";
        }
        return "string";
    }

    private function escapeCellFormat(cell_format)
    {
        var ignore_until, escaped, i, c;
        let ignore_until = "";
        let escaped = "";
        for i in range(0, strlen(cell_format))
        {
            let c = substr(cell_format, i, 1);
            if (ignore_until=="" && c=='[') {
                let ignore_until=']';
            } elseif (ignore_until=="" && c=='"') {
                let ignore_until='"';
            } elseif (ignore_until==c) {
                let ignore_until="";
            }
            if (ignore_until=="" && (c==' ' || c=='-'  || c=='('  || c==')') && (i==0 || cell_format[-1+i]!='_')) {
                let escaped.= "".c;
            } else {
                let escaped.= c;
            }
        }
        return escaped;
        //return str_replace( array(" ","-", "(", ")"), array("\ ","\-", "\(", "\)"), cell_format);//TODO, needs more escaping
    }

    private function addCellFormat(cell_format)
    {
        //for backwards compatibility, to handle older versions
        if (cell_format=="string") {
            let cell_format = "GENERAL";
        }
        elseif (cell_format=="integer") {
            let cell_format = "0";
        }
        elseif (cell_format=="date") {
            let cell_format = "YYYY-MM-DD";
        }
        elseif (cell_format=="datetime") {
            let cell_format = "YYYY-MM-DD HH:MM:SS";
        }
        elseif (cell_format=="dollar") {
            let cell_format = "[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00";
        }
        elseif (cell_format=="money") {
            let cell_format = "[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00";
        }
        elseif (cell_format=="euro") {
            let cell_format = "#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]";
        }
        elseif (cell_format=="NN") {
            let cell_format = "DDD";
        }
        elseif (cell_format=="NNN") {
            let cell_format = "DDDD";
        }
        elseif (cell_format=="NNNN") {
            let cell_format = "DDDD\", \"";
        }

        let cell_format = strtoupper(cell_format);
        var position;
        let position = isset this->cell_formats[this->escapeCellFormat(cell_format)] ? this->cell_position[this->escapeCellFormat(cell_format)] : false;
        if (position===false)
        {

            let position = this->format_count;
            let this->format_count = this->format_count + 1;
            let this->cell_position[this->escapeCellFormat(cell_format)] = position;
            let this->cell_formats[] = this->escapeCellFormat(cell_format);
            let this->cell_types[] = this->determineCellType(cell_format);
        }

        return position;
    }

    public function writeSheetHeader(sheet_name, array header_types, suppress_row = false)
    {
        if (empty sheet_name) {
            throw new \Exception("sheet_name empty");
        }
        if (empty header_types) {
            throw new \Exception("header_types empty");
        }
        if (isset this->sheets[sheet_name]) {
            throw new \Exception("already set sheet_name");
        }

        self::initializeSheet(sheet_name);
        var v, header_row, k;
        let this->sheets[sheet_name]["columns"] = [];
        for v in header_types {
            let this->sheets[sheet_name]["columns"][] = this->addCellFormat(v);
            let this->sheets[sheet_name]["columns_count"] = this->sheets[sheet_name]["columns_count"] + 1;
        }
        if (!suppress_row)
        {
            let header_row = array_keys(header_types);

            this->sheets[sheet_name]["file_writer"]->write("<row collapsed=\"false\" customFormat=\"false\" customHeight=\"false\" hidden=\"false\" ht=\"12.1\" outlineLevel=\"0\" r=\"" . (1) . "\">");
            for k, v in header_row {
                let this->sheets[sheet_name]["file_writer"] = this->writeCell(this->sheets[sheet_name]["file_writer"], 0, k, v, "0");//'0'=>"string"
            }
            this->sheets[sheet_name]["file_writer"]->write("</row>");
            let this->sheets[sheet_name]["row_count"] = this->sheets[sheet_name]["row_count"] + 1;
        }
        let this->current_sheet = sheet_name;
    }

    public function writeSheetRow(sheet_name, array row)
    {
        if (empty sheet_name || empty row) {
            return;
        }

        self::initializeSheet(sheet_name);
        var v, column_count;
        if (empty this->sheets[sheet_name]["columns"])
        {
            let this->sheets[sheet_name]["columns"] = array_fill(0, count(row), "0");//'0'=>"string"
        }

        this->sheets[sheet_name]["file_writer"]->write("<row collapsed=\"false\" customFormat=\"false\" customHeight=\"false\" hidden=\"false\" ht=\"12.1\" outlineLevel=\"0\" r=\"" . (this->sheets[sheet_name]["row_count"] + 1) . "\">");
        let column_count = 0;
        for v in row {
            let this->sheets[sheet_name]["file_writer"] = this->writeCell(this->sheets[sheet_name]["file_writer"], this->sheets[sheet_name]["row_count"], column_count, v, this->sheets[sheet_name]["columns"][column_count]);
            let column_count++;
        }
        this->sheets[sheet_name]["file_writer"]->write("</row>");
        let this->sheets[sheet_name]["row_count"] = this->sheets[sheet_name]["row_count"] + 1;
        let this->current_sheet = sheet_name;
    }

    protected function finalizeSheet(sheet_name)
    {
        if (empty sheet_name || this->sheets[sheet_name]["finalized"]) {
            return;
        }

        var max_cell, max_cell_tag, padding_length, range;

        this->sheets[sheet_name]["file_writer"]->write(    "</sheetData>");

        if (!empty this->sheets[sheet_name]["merge_cells"]) {
            this->sheets[sheet_name]["file_writer"]->write(    "<mergeCells>");
            for range in this->sheets[sheet_name]["merge_cells"] {
                this->sheets[sheet_name]["file_writer"]->write(        "<mergeCell ref=\"" . range . "\"/>");
            }
            this->sheets[sheet_name]["file_writer"]->write(    "</mergeCells>");
        }

        this->sheets[sheet_name]["file_writer"]->write(    "<printOptions headings=\"false\" gridLines=\"false\" gridLinesSet=\"true\" horizontalCentered=\"false\" verticalCentered=\"false\"/>")
            ->write(    "<pageMargins left=\"0.5\" right=\"0.5\" top=\"1.0\" bottom=\"1.0\" header=\"0.5\" footer=\"0.5\"/>")
            ->write(    "<pageSetup blackAndWhite=\"false\" cellComments=\"none\" copies=\"1\" draft=\"false\" firstPageNumber=\"1\" fitToHeight=\"1\" fitToWidth=\"1\" horizontalDpi=\"300\" orientation=\"portrait\" pageOrder=\"downThenOver\" paperSize=\"1\" scale=\"100\" useFirstPageNumber=\"true\" usePrinterDefaults=\"false\" verticalDpi=\"300\"/>")
            ->write(    "<headerFooter differentFirst=\"false\" differentOddEven=\"false\">")
            ->write(        "<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>")
            ->write(        "<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>")
            ->write(    "</headerFooter>")
            ->write("</worksheet>");

        let max_cell = self::xlsCell(this->sheets[sheet_name]["row_count"] - 1, this->sheets[sheet_name]["columns_count"] - 1);
        let max_cell_tag = "<dimension ref=\"A1:" . max_cell . "\"/>";
        let padding_length = this->sheets[sheet_name]["max_cell_tag_end"] - this->sheets[sheet_name]["max_cell_tag_start"] - strlen(max_cell_tag);
        this->sheets[sheet_name]["file_writer"]->fseek(this->sheets[sheet_name]["max_cell_tag_start"]);
        this->sheets[sheet_name]["file_writer"]->write(max_cell_tag.str_repeat(" ", padding_length));
        this->sheets[sheet_name]["file_writer"]->close();
        let this->sheets[sheet_name]["finalized"]=true;
    }

    public function markMergedCell(sheet_name, start_cell_row, start_cell_column, end_cell_row, end_cell_column)
    {
        if (empty sheet_name || this->sheets[sheet_name]["finalized"]) {
            return;
        }

        self::initializeSheet(sheet_name);
        var startCell, endCell;
        let startCell = self::xlsCell(start_cell_row, start_cell_column);
        let endCell = self::xlsCell(end_cell_row, end_cell_column);
        let this->sheets[sheet_name]["merge_cells"][] =  startCell . ":" . endCell;
    }

    public function writeSheet(array data, sheet_name="", array header_types=[])
    {
        let sheet_name = empty sheet_name ? "Sheet1" : sheet_name;
        let data = empty data ? [[""]] : data;
        if (!empty header_types)
        {
            this->writeSheetHeader(sheet_name, header_types);
        }
        var row;
        for row in data {
            this->writeSheetRow(sheet_name, row);
        }
        this->finalizeSheet(sheet_name);
    }

    // protected function writeCell(<BuffererWriter> file, row_number, column_number, value, cell_format_index)
    protected function writeCell(file, row_number, column_number, value, cell_format_index)
    {
        var cell_type, cell_name;
        let cell_type = this->cell_types[cell_format_index];
        let cell_name = self::xlsCell(row_number, column_number);

        if (!is_scalar(value) || value==="") { //objects, array, empty
            file->write("<c r=\"".cell_name."\" s=\"".cell_format_index."\"/>");
        } elseif (is_string(value) && substr(value, 0, 1)=="="){
            file->write("<c r=\"".cell_name."\" s=\"".cell_format_index."\" t=\"s\"><f>".self::xmlspecialchars(value)."</f></c>");
        } elseif (cell_type=="date") {
            file->write("<c r=\"".cell_name."\" s=\"".cell_format_index."\" t=\"n\"><v>".intval(self::convert_date_time(value))."</v></c>");
        } elseif (cell_type=="datetime") {
            file->write("<c r=\"".cell_name."\" s=\"".cell_format_index."\" t=\"n\"><v>".self::convert_date_time(value)."</v></c>");
        } elseif (cell_type=="currency" || cell_type=="percent" || cell_type=="numeric") {
            file->write("<c r=\"".cell_name."\" s=\"".cell_format_index."\" t=\"n\"><v>".self::xmlspecialchars(value)."</v></c>");//int,float,currency
        } elseif (!is_string(value)){
            file->write("<c r=\"".cell_name."\" s=\"".cell_format_index."\" t=\"n\"><v>".(value*1)."</v></c>");
        } elseif (substr(value, 0, 1)!="0" && substr(value, 0, 1)!="+" && filter_var(value, FILTER_VALIDATE_INT, ["options":["max_range":2147483647]])){
            file->write("<c r=\"".cell_name."\" s=\"".cell_format_index."\" t=\"n\"><v>".(value*1)."</v></c>");
        } else { //implied: (cell_format=="string")
            file->write("<c r=\"".cell_name."\" s=\"".cell_format_index."\" t=\"s\"><v>".self::xmlspecialchars(this->setSharedString(value))."</v></c>");
        }
        return file;
    }//

    protected function writeStylesXML()
    {
        var temporary_filename, file, i, v, countCellFormats, alignmentData;
        let temporary_filename = this->tempFilename();
        let file = new BufferedWriter(temporary_filename, "w", false, this->bufferSize);
        let countCellFormats = this->format_count;
        file->write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
            ->write("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"."\n")
            ->write("<numFmts count=\"".countCellFormats."\">"."\n");
        let alignmentData = "";
        for i, v in this->cell_formats {
            file->write("<numFmt numFmtId=\"".(164+i)."\" formatCode=\"".self::xmlspecialchars(v)."\" />"."\n");
            let alignmentData = alignmentData . "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"false\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"".(164+i)."\" xfId=\"0\"/>"."\n";
        }
        //file->write(     "<numFmt formatCode=\"GENERAL\" numFmtId=\"164\"/>"."\n");
        //file->write(     "<numFmt formatCode=\"[$$-1009]#,##0.00;[RED]\-[$$-1009]#,##0.00\" numFmtId=\"165\"/>"."\n");
        //file->write(     "<numFmt formatCode=\"YYYY-MM-DD\ HH:MM:SS\" numFmtId=\"166\"/>"."\n");
        //file->write(     "<numFmt formatCode=\"YYYY-MM-DD\" numFmtId=\"167\"/>"."\n");
        file->write("</numFmts>"."\n")
            ->write("<fonts count=\"4\">"."\n")
            ->write(       "<font><name val=\"Arial\"/><charset val=\"1\"/><family val=\"2\"/><sz val=\"10\"/></font>"."\n")
            ->write(       "<font><name val=\"Arial\"/><family val=\"0\"/><sz val=\"10\"/></font>"."\n")
            ->write(       "<font><name val=\"Arial\"/><family val=\"0\"/><sz val=\"10\"/></font>"."\n")
            ->write(       "<font><name val=\"Arial\"/><family val=\"0\"/><sz val=\"10\"/></font>"."\n")
            ->write("</fonts>"."\n")
            ->write("<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"."\n")
            ->write("<borders count=\"1\"><border diagonalDown=\"false\" diagonalUp=\"false\"><left/><right/><top/><bottom/><diagonal/></border></borders>"."\n")
            ->write(   "<cellStyleXfs count=\"20\">"."\n")
            ->write(       "<xf applyAlignment=\"true\" applyBorder=\"true\" applyFont=\"true\" applyProtection=\"true\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"164\">"."\n")
            ->write(       "<alignment horizontal=\"general\" indent=\"0\" shrinkToFit=\"false\" textRotation=\"0\" vertical=\"bottom\" wrapText=\"false\"/>"."\n")
            ->write(       "<protection hidden=\"false\" locked=\"true\"/>"."\n")
            ->write(       "</xf>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"2\" numFmtId=\"0\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"2\" numFmtId=\"0\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"43\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"41\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"44\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"42\"/>"."\n")
            ->write(       "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"9\"/>"."\n")
            ->write(   "</cellStyleXfs>"."\n");

        file->write(   "<cellXfs count=\"".countCellFormats."\">"."\n");
        file->write(alignmentData);
        // for i,v in this->cell_formats {
        //     let v = v;
        //     file->write("<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"false\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"".(164+i)."\" xfId=\"0\"/>"."\n");
        // }
        file->write(   "</cellXfs>"."\n");
        //file->write( "<cellXfs count=\"4\">"."\n");
        //file->write(     "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"false\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"164\" xfId=\"0\"/>"."\n");
        //file->write(     "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"false\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"165\" xfId=\"0\"/>"."\n");
        //file->write(     "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"false\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"166\" xfId=\"0\"/>"."\n");
        //file->write(     "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"false\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"167\" xfId=\"0\"/>"."\n");
        //file->write( "</cellXfs>"."\n");
        file->write(   "<cellStyles count=\"6\">"."\n")
            ->write(       "<cellStyle builtinId=\"0\" customBuiltin=\"false\" name=\"Normal\" xfId=\"0\"/>"."\n")
            ->write(       "<cellStyle builtinId=\"3\" customBuiltin=\"false\" name=\"Comma\" xfId=\"15\"/>"."\n")
            ->write(       "<cellStyle builtinId=\"6\" customBuiltin=\"false\" name=\"Comma [0]\" xfId=\"16\"/>"."\n")
            ->write(       "<cellStyle builtinId=\"4\" customBuiltin=\"false\" name=\"Currency\" xfId=\"17\"/>"."\n")
            ->write(       "<cellStyle builtinId=\"7\" customBuiltin=\"false\" name=\"Currency [0]\" xfId=\"18\"/>"."\n")
            ->write(       "<cellStyle builtinId=\"5\" customBuiltin=\"false\" name=\"Percent\" xfId=\"19\"/>"."\n")
            ->write(   "</cellStyles>"."\n")
            ->write("</styleSheet>");
        file->close();
        return temporary_filename;
    }

    protected function setSharedString(v)
    {
        var string_value;
        if (isset this->shared_strings[v])
        {
            let string_value = this->shared_strings[v];
        }
        else
        {
            let string_value = this->shared_string_count_seq;
            let this->shared_strings[v] = this->shared_string_count_seq;
            let this->shared_string_count_seq = this->shared_string_count_seq + 1;
        }
        let this->shared_string_count++;//non-unique count
        return string_value;
    }

    protected function writeSharedStringsXML()
    {
        var temporary_filename, file, s, c;
        let temporary_filename = this->tempFilename();
        let file = new BufferedWriter(temporary_filename, "w", true, this->bufferSize);
        file->write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
            ->write("<sst count=\"".(this->shared_string_count)."\" uniqueCount=\"".this->shared_string_count_seq."\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        for s, c in this->shared_strings {
            let c = c;
            file->write("<si><t>".self::xmlspecialchars(s)."</t></si>");
        }
        file->write("</sst>");
        file->close();

        return temporary_filename;
    }

    protected function buildAppXML()
    {
        var app_xml;
        let app_xml="";
        let app_xml.="<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"."\n";
        let app_xml.="<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><TotalTime>0</TotalTime></Properties>";
        return app_xml;
    }

    protected function buildCoreXML()
    {
        var core_xml;
        let core_xml="";
        let core_xml.="<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"."\n";
        let core_xml.="<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"."\n";
        let core_xml.="<dcterms:created xsi:type=\"dcterms:W3CDTF\">".date("Y-m-d\TH:i:s.00\Z")."</dcterms:created>"."\n";
        let core_xml.="<dc:creator>".self::xmlspecialchars(this->author)."</dc:creator>"."\n";
        let core_xml.="<cp:revision>0</cp:revision>"."\n";
        let core_xml.="</cp:coreProperties>";
        return core_xml;
    }

    protected function buildRelationshipsXML()
    {
        var rels_xml;
        let rels_xml="";
        let rels_xml.="<?xml version=\"1.0\" encoding=\"UTF-8\"?>"."\n";
        let rels_xml.="<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
        let rels_xml.="<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>";
        let rels_xml.="<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>";
        let rels_xml.="<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>";
        let rels_xml.="\n";
        let rels_xml.="</Relationships>";
        return rels_xml;
    }

    protected function buildWorkbookXML()
    {
        var i, workbook_xml;
        let i=0;
        let workbook_xml="";
        let workbook_xml.="<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"."\n";
        let workbook_xml.="<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">";
        let workbook_xml.="<fileVersion appName=\"Calc\"/><workbookPr backupFile=\"false\" showObjects=\"all\" date1904=\"false\"/><workbookProtection/>";
        let workbook_xml.="<bookViews><workbookView activeTab=\"0\" firstSheet=\"0\" showHorizontalScroll=\"true\" showSheetTabs=\"true\" showVerticalScroll=\"true\" tabRatio=\"212\" windowHeight=\"8192\" windowWidth=\"16384\" xWindow=\"0\" yWindow=\"0\"/></bookViews>";
        let workbook_xml.="<sheets>";
        var sheet;
        for sheet in this->sheets {
            let workbook_xml.="<sheet name=\"".self::xmlspecialchars(sheet["sheetname"])."\" sheetId=\"".(i+1)."\" state=\"visible\" r:id=\"rId".(i+2)."\"/>";
            let i++;
        }
        let workbook_xml.="</sheets>";
        let workbook_xml.="<calcPr iterateCount=\"100\" refMode=\"A1\" iterate=\"false\" iterateDelta=\"0.001\"/></workbook>";
        return workbook_xml;
    }

    protected function buildWorkbookRelsXML()
    {
        var i, wkbkrels_xml, sheet;
        let i=0;
        let wkbkrels_xml="";
        let wkbkrels_xml.="<?xml version=\"1.0\" encoding=\"UTF-8\"?>"."\n";
        let wkbkrels_xml.="<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
        let wkbkrels_xml.="<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>";
        for sheet in this->sheets {
            let wkbkrels_xml.="<Relationship Id=\"rId".(i+2)."\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/".(sheet["xmlname"])."\"/>";
            let i++;
        }
        if (!empty this->shared_strings) {
            let wkbkrels_xml.="<Relationship Id=\"rId".(this->sheet_count+2)."\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>";
        }
        let wkbkrels_xml.="\n";
        let wkbkrels_xml.="</Relationships>";
        return wkbkrels_xml;
    }

    protected function buildContentTypesXML()
    {
        var content_types_xml, sheet;
        let content_types_xml="";
        let content_types_xml.="<?xml version=\"1.0\" encoding=\"UTF-8\"?>"."\n";
        let content_types_xml.="<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">";
        let content_types_xml.="<Override PartName=\"/_rels/.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>";
        let content_types_xml.="<Override PartName=\"/xl/_rels/workbook.xml.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>";
        for sheet in this->sheets  {
            let content_types_xml.="<Override PartName=\"/xl/worksheets/".(sheet["xmlname"])."\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
        }
        if (!empty this->shared_strings) {
            let content_types_xml.="<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>";
        }
        let content_types_xml.="<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>";
        let content_types_xml.="<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>";
        let content_types_xml.="<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>";
        let content_types_xml.="<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>";
        let content_types_xml.="\n";
        let content_types_xml.="</Types>";
        return content_types_xml;
    }

    //------------------------------------------------------------------
    /*
     * @param row_number int, zero based
     * @param column_number int, zero based
     * @return Cell label/coordinates, ex: A1, C3, AA42
     * */
    public static function xlsCell(row_number, column_number)
    {
        var n, r;
        let n = column_number;
        let r = "";
        while n>=0 {
            let r = chr(n%26 + 0x41) . r;
            let n = intval(n / 26) - 1;
        }

        return r . (row_number+1);
    }
    //------------------------------------------------------------------
    public static function log(message) {
        file_put_contents("php://stderr", date("Y-m-d H:i:s:").rtrim(is_array(message) ? json_encode(message) : message)."\n");
    }
    //------------------------------------------------------------------
    public static function sanitize_filename(filename) //http://msdn.microsoft.com/en-us/library/aa365247%28VS.85%29.aspx
    {
        var nonprinting, invalid_chars, all_invalids;
        let nonprinting = array_map("chr", range(0,31));
        let invalid_chars = ["<", ">", "?", "", ":", "|", "\\", "/", "*", "&"];
        let all_invalids = array_merge(nonprinting, invalid_chars);
        return str_replace(all_invalids, "", filename);
    }
    //------------------------------------------------------------------
    public static function xmlspecialchars(val)
    {
        return str_replace("", "&#39;", htmlspecialchars(val));
    }
    //------------------------------------------------------------------
    public static function array_first_key(array arr)
    {
        var first_key;
        reset(arr);
        let first_key = key(arr);
        return first_key;
    }
    //------------------------------------------------------------------
    public static function convert_date_time(date_input) //thanks to Excel::Writer::XLSX::Worksheet.pm (perl)
    {
        var days, seconds, year, month, day, hour, min, sec, date_time, epoch, offset, norm, range, leap, mdays, matches;
        let days    = 0;    // Number of days since epoch
        let seconds = 0;    // Time expressed as fraction of 24h hours in seconds
        let year    = 0;
        let month   = 0;
        let day     = 0;
        let hour    = 0;
        let min     = 0;
        let sec     = 0;

        let date_time = date_input;
        if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", date_time, matches))
        {
            let year  = matches[1];
            let month = matches[2];
            let day   = matches[3];
        }
        if (preg_match("/(\d{2}):(\d{2}):(\d{2})/", date_time, matches))
        {
            let hour  = matches[1];
            let min   = matches[2];
            let sec   = matches[3];
            let seconds = ( hour * 60 * 60 + min * 60 + sec ) / ( 24 * 60 * 60 );
        }

        //using 1900 as epoch, not 1904, ignoring 1904 special case

        // Special cases for Excel.
        if ("year-month-day"=="1899-12-31") {
            return seconds;
        } // Excel 1900 epoch
        if ("year-month-day"=="1900-01-00") {
            return seconds;
        } // Excel 1900 epoch
        if ("year-month-day"=="1900-02-29") {
            return 60 + seconds;
        } // Excel false leapday

        // We calculate the date by calculating the number of days since the epoch
        // and adjust for the number of leap days. We calculate the number of leap
        // days by normalising the year in relation to the epoch. Thus the year 2000
        // becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
        let epoch  = 1900;
        let offset = 0;
        let norm   = 300;
        let range  = year - epoch;

        // Set month days and check for leap year.
        let leap = ((year % 400 == 0) || ((year % 4 == 0) && (year % 100)) ) ? 1 : 0;
        let mdays = [31, (leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 ];

        // Some boundary checks
        if(year < epoch || year > 9999) {
            return 0;
        }
        if(month < 1 || month > 12)  {
            return 0;
        }
        if(day < 1 || day > mdays[ month - 1 ]) {
            return 0;
        }

        // Accumulate the number of days since the epoch.
        let days = day;    // Add days for current month
        let days += array_sum( array_slice(mdays, 0, 1-month ) );    // Add days for past months
        let days += range * 365;                      // Add days for past years
        let days += intval( ( range ) / 4 );             // Add leapdays
        let days -= intval( ( range + offset ) / 100 ); // Subtract 100 year leapdays
        let days += intval( ( range + offset + norm ) / 400 );  // Add 400 year leapdays
        let days -= leap;                                      // Already counted above

        // Adjust for Excel erroneously treating 1900 as a leap year.
        if (days > 59) {
            let days++;
        }

        return days + seconds;
    }
    //------------------------------------------------------------------
}

// vim: set filetype=php expandtab tabstop=4 shiftwidth=4 autoindent smartindent:
