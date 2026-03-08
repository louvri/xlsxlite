package xlsxlite

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"sort"
	"strconv"
	"time"
)

// Writer creates XLSX files using streaming writes.
// Rows are flushed to the underlying zip entry immediately,
// so memory usage stays proportional to a single row regardless of file size.
type Writer struct {
	zw         *zip.Writer
	styles     *StyleSheet
	sheets     []sheetMeta
	sharedStrs *sharedStringWriter
}

type sheetMeta struct {
	config SheetConfig
	rID    string
}

// NewWriter creates a new streaming XLSX writer that writes to w.
// Typical workflow: register styles via StyleSheet(), create sheets via NewSheet(),
// write rows via SheetWriter.WriteRow(), then call Close() to finalize.
func NewWriter(w io.Writer) *Writer {
	return &Writer{
		zw:         zip.NewWriter(w),
		styles:     NewStyleSheet(),
		sharedStrs: newSharedStringWriter(),
	}
}

// StyleSheet returns the style sheet for registering styles.
// Styles must be registered before writing rows that reference them.
func (w *Writer) StyleSheet() *StyleSheet {
	return w.styles
}

// SheetWriter writes rows to a single worksheet in streaming fashion.
// Rows are flushed immediately on WriteRow, so memory usage is O(1) per row.
type SheetWriter struct {
	parent     *Writer
	zEntry     io.Writer
	config     SheetConfig
	rowNum     int
	sheetIndex int
	buf        bytes.Buffer // reusable buffer for building row XML
}

// NewSheet begins a new worksheet. You must call SheetWriter.Close()
// when done writing rows to this sheet before starting another.
func (w *Writer) NewSheet(config SheetConfig) (*SheetWriter, error) {
	idx := len(w.sheets) + 1
	rID := "rId" + strconv.Itoa(idx)
	w.sheets = append(w.sheets, sheetMeta{config: config, rID: rID})

	path := "xl/worksheets/sheet" + strconv.Itoa(idx) + ".xml"
	entry, err := w.zw.Create(path)
	if err != nil {
		return nil, fmt.Errorf("create sheet entry: %w", err)
	}

	sw := &SheetWriter{
		parent:     w,
		zEntry:     entry,
		config:     config,
		sheetIndex: idx,
	}

	// Write sheet header
	if err := sw.writeHeader(); err != nil {
		return nil, err
	}

	return sw, nil
}

func (sw *SheetWriter) writeHeader() error {
	b := &sw.buf
	b.Reset()

	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	b.WriteByte('\n')
	b.WriteString(`<worksheet xmlns="`)
	b.WriteString(nsSpreadsheetML)
	b.WriteString(`" xmlns:r="`)
	b.WriteString(nsOfficeDocRels)
	b.WriteString(`">`)
	b.WriteByte('\n')

	// Column widths
	if len(sw.config.ColWidths) > 0 {
		b.WriteString(`<cols>`)
		cols := make([]int, 0, len(sw.config.ColWidths))
		for col := range sw.config.ColWidths {
			cols = append(cols, col)
		}
		sort.Ints(cols)
		for _, col := range cols {
			colNum := col + 1 // 1-based
			b.WriteString(`<col min="`)
			b.Write(strconv.AppendInt(nil, int64(colNum), 10))
			b.WriteString(`" max="`)
			b.Write(strconv.AppendInt(nil, int64(colNum), 10))
			b.WriteString(`" width="`)
			b.Write(strconv.AppendFloat(nil, sw.config.ColWidths[col], 'g', -1, 64))
			b.WriteString(`" customWidth="1"/>`)
		}
		b.WriteString("</cols>\n")
	}

	// Freeze panes
	if sw.config.FreezeRow > 0 || sw.config.FreezeCol > 0 {
		topLeft := CellRef(sw.config.FreezeCol, sw.config.FreezeRow+1)
		b.WriteString(`<sheetViews><sheetView tabSelected="1" workbookViewId="0">`)
		b.WriteString(`<pane`)
		if sw.config.FreezeRow > 0 {
			b.WriteString(` ySplit="`)
			b.Write(strconv.AppendInt(nil, int64(sw.config.FreezeRow), 10))
			b.WriteByte('"')
		}
		if sw.config.FreezeCol > 0 {
			b.WriteString(` xSplit="`)
			b.Write(strconv.AppendInt(nil, int64(sw.config.FreezeCol), 10))
			b.WriteByte('"')
		}
		b.WriteString(` topLeftCell="`)
		b.WriteString(topLeft)
		b.WriteString(`" state="frozen"/>`)
		b.WriteString("</sheetView></sheetViews>\n")
	}

	b.WriteString("<sheetData>\n")

	_, err := sw.zEntry.Write(b.Bytes())
	return err
}

// WriteRow writes a single row and immediately flushes it to the zip entry.
// Rows must be written in sequential order. If Row.RowIndex is set (> 0),
// it overrides the auto-incremented row number, allowing gaps between rows.
func (sw *SheetWriter) WriteRow(row Row) error {
	sw.rowNum++
	rowIdx := sw.rowNum
	if row.RowIndex > 0 {
		rowIdx = row.RowIndex
		sw.rowNum = rowIdx
	}

	b := &sw.buf
	b.Reset()

	b.WriteString(`<row r="`)
	b.Write(strconv.AppendInt(b.AvailableBuffer(), int64(rowIdx), 10))
	b.WriteByte('"')
	if row.Height > 0 {
		b.WriteString(` ht="`)
		b.Write(strconv.AppendFloat(b.AvailableBuffer(), row.Height, 'g', -1, 64))
		b.WriteString(`" customHeight="1"`)
	}
	b.WriteByte('>')

	for colIdx, cell := range row.Cells {
		sw.appendCell(b, colIdx, rowIdx, cell)
	}

	b.WriteString("</row>\n")

	_, err := sw.zEntry.Write(b.Bytes())
	return err
}

func (sw *SheetWriter) appendCell(b *bytes.Buffer, colIdx, rowIdx int, cell Cell) {
	if cell.Type == CellTypeEmpty && cell.Value == nil {
		if cell.StyleID > 0 {
			b.WriteString(`<c r="`)
			appendCellRef(b, colIdx, rowIdx)
			b.WriteString(`" s="`)
			b.Write(strconv.AppendInt(b.AvailableBuffer(), int64(cell.StyleID), 10))
			b.WriteString(`"/>`)
		}
		return
	}

	b.WriteString(`<c r="`)
	appendCellRef(b, colIdx, rowIdx)
	b.WriteByte('"')
	if cell.StyleID > 0 {
		b.WriteString(` s="`)
		b.Write(strconv.AppendInt(b.AvailableBuffer(), int64(cell.StyleID), 10))
		b.WriteByte('"')
	}

	switch v := cell.Value.(type) {
	case string:
		idx := sw.parent.sharedStrs.Add(v)
		b.WriteString(` t="s"><v>`)
		b.Write(strconv.AppendInt(b.AvailableBuffer(), int64(idx), 10))
		b.WriteString(`</v></c>`)

	case float64:
		b.WriteString(`><v>`)
		b.Write(strconv.AppendFloat(b.AvailableBuffer(), v, 'f', -1, 64))
		b.WriteString(`</v></c>`)

	case float32:
		b.WriteString(`><v>`)
		b.Write(strconv.AppendFloat(b.AvailableBuffer(), float64(v), 'f', -1, 32))
		b.WriteString(`</v></c>`)

	case int:
		b.WriteString(`><v>`)
		b.Write(strconv.AppendInt(b.AvailableBuffer(), int64(v), 10))
		b.WriteString(`</v></c>`)

	case int64:
		b.WriteString(`><v>`)
		b.Write(strconv.AppendInt(b.AvailableBuffer(), v, 10))
		b.WriteString(`</v></c>`)

	case bool:
		if v {
			b.WriteString(` t="b"><v>1</v></c>`)
		} else {
			b.WriteString(` t="b"><v>0</v></c>`)
		}

	case time.Time:
		serial := TimeToExcelSerial(v)
		b.WriteString(`><v>`)
		b.Write(strconv.AppendFloat(b.AvailableBuffer(), serial, 'f', -1, 64))
		b.WriteString(`</v></c>`)

	case nil:
		if cell.StyleID > 0 {
			b.WriteString(`/>`)
		}

	default:
		s := fmt.Sprintf("%v", v)
		idx := sw.parent.sharedStrs.Add(s)
		b.WriteString(` t="s"><v>`)
		b.Write(strconv.AppendInt(b.AvailableBuffer(), int64(idx), 10))
		b.WriteString(`</v></c>`)
	}
}

// appendCellRef writes a cell reference like "A1" directly into the buffer.
func appendCellRef(b *bytes.Buffer, col, row int) {
	// Column letters
	c := col
	var colBuf [4]byte // max 4 letters for Excel columns
	i := len(colBuf)
	for c >= 0 {
		i--
		colBuf[i] = byte('A' + c%26)
		c = c/26 - 1
	}
	b.Write(colBuf[i:])
	// Row number
	b.Write(strconv.AppendInt(b.AvailableBuffer(), int64(row), 10))
}

// Close finalizes the sheet XML (closes sheetData, writes merge cells, etc.).
// Must be called before creating another sheet or closing the Writer.
func (sw *SheetWriter) Close() error {
	b := &sw.buf
	b.Reset()

	b.WriteString("</sheetData>\n")

	// Merge cells
	if len(sw.config.MergeCells) > 0 {
		b.WriteString(`<mergeCells count="`)
		b.Write(strconv.AppendInt(b.AvailableBuffer(), int64(len(sw.config.MergeCells)), 10))
		b.WriteString(`">`)
		for _, mc := range sw.config.MergeCells {
			b.WriteString(`<mergeCell ref="`)
			b.WriteString(RangeRef(mc.StartCol, mc.StartRow, mc.EndCol, mc.EndRow))
			b.WriteString(`"/>`)
		}
		b.WriteString("</mergeCells>\n")
	}

	b.WriteString(`</worksheet>`)

	_, err := sw.zEntry.Write(b.Bytes())
	return err
}

// Close finalizes the entire XLSX package by writing styles, shared strings,
// workbook, content types, and relationships, then closes the underlying zip writer.
// All SheetWriters must be closed before calling this method.
func (w *Writer) Close() error {
	// 1. Write shared strings
	if err := w.writeSharedStrings(); err != nil {
		return fmt.Errorf("write shared strings: %w", err)
	}

	// 2. Write styles
	if err := w.writeStyles(); err != nil {
		return fmt.Errorf("write styles: %w", err)
	}

	// 3. Write workbook
	if err := w.writeWorkbook(); err != nil {
		return fmt.Errorf("write workbook: %w", err)
	}

	// 4. Write workbook relationships
	if err := w.writeWorkbookRels(); err != nil {
		return fmt.Errorf("write workbook rels: %w", err)
	}

	// 5. Write root relationships
	if err := w.writeRootRels(); err != nil {
		return fmt.Errorf("write root rels: %w", err)
	}

	// 6. Write content types
	if err := w.writeContentTypes(); err != nil {
		return fmt.Errorf("write content types: %w", err)
	}

	return w.zw.Close()
}

func (w *Writer) writeSharedStrings() error {
	entry, err := w.zw.Create("xl/sharedStrings.xml")
	if err != nil {
		return err
	}
	return w.sharedStrs.writeXML(entry)
}

func (w *Writer) writeStyles() error {
	entry, err := w.zw.Create("xl/styles.xml")
	if err != nil {
		return err
	}
	return w.styles.writeXML(entry)
}

func (w *Writer) writeWorkbook() error {
	entry, err := w.zw.Create("xl/workbook.xml")
	if err != nil {
		return err
	}
	write := func(s string) error {
		_, err := io.WriteString(entry, s)
		return err
	}

	if err := write(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"); err != nil {
		return err
	}
	if err := write(`<workbook xmlns="` + nsSpreadsheetML + `" xmlns:r="` + nsOfficeDocRels + `">` + "\n"); err != nil {
		return err
	}
	if err := write(`<sheets>` + "\n"); err != nil {
		return err
	}
	for i, sheet := range w.sheets {
		name := escapeXMLAttr(sheet.config.Name)
		if name == "" {
			name = "Sheet" + strconv.Itoa(i+1)
		}
		if err := write(`<sheet name="` + name + `" sheetId="` + strconv.Itoa(i+1) + `" r:id="` + sheet.rID + `"/>`); err != nil {
			return err
		}
	}
	if err := write("\n" + `</sheets>` + "\n"); err != nil {
		return err
	}
	return write(`</workbook>`)
}

func (w *Writer) writeWorkbookRels() error {
	entry, err := w.zw.Create("xl/_rels/workbook.xml.rels")
	if err != nil {
		return err
	}
	write := func(s string) error {
		_, err := io.WriteString(entry, s)
		return err
	}

	if err := write(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"); err != nil {
		return err
	}
	if err := write(`<Relationships xmlns="` + nsRelationships + `">` + "\n"); err != nil {
		return err
	}

	// Sheet relationships
	for i, sheet := range w.sheets {
		if err := write(`<Relationship Id="` + sheet.rID + `" Type="` + relTypeWorksheet + `" Target="worksheets/sheet` + strconv.Itoa(i+1) + `.xml"/>`); err != nil {
			return err
		}
	}

	// Styles relationship
	stylesRID := "rId" + strconv.Itoa(len(w.sheets)+1)
	if err := write(`<Relationship Id="` + stylesRID + `" Type="` + relTypeStyles + `" Target="styles.xml"/>`); err != nil {
		return err
	}

	// Shared strings relationship
	ssRID := "rId" + strconv.Itoa(len(w.sheets)+2)
	if err := write(`<Relationship Id="` + ssRID + `" Type="` + relTypeSharedStrings + `" Target="sharedStrings.xml"/>`); err != nil {
		return err
	}

	return write("\n" + `</Relationships>`)
}

func (w *Writer) writeRootRels() error {
	entry, err := w.zw.Create("_rels/.rels")
	if err != nil {
		return err
	}
	write := func(s string) error {
		_, err := io.WriteString(entry, s)
		return err
	}

	if err := write(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"); err != nil {
		return err
	}
	if err := write(`<Relationships xmlns="` + nsRelationships + `">` + "\n"); err != nil {
		return err
	}
	if err := write(`<Relationship Id="rId1" Type="` + nsOfficeDocRels + `/officeDocument" Target="xl/workbook.xml"/>`); err != nil {
		return err
	}
	return write("\n" + `</Relationships>`)
}

func (w *Writer) writeContentTypes() error {
	entry, err := w.zw.Create("[Content_Types].xml")
	if err != nil {
		return err
	}
	write := func(s string) error {
		_, err := io.WriteString(entry, s)
		return err
	}

	if err := write(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"); err != nil {
		return err
	}
	if err := write(`<Types xmlns="` + nsContentTypes + `">` + "\n"); err != nil {
		return err
	}
	if err := write(`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` + "\n"); err != nil {
		return err
	}
	if err := write(`<Default Extension="xml" ContentType="application/xml"/>` + "\n"); err != nil {
		return err
	}

	// Workbook
	if err := write(`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` + "\n"); err != nil {
		return err
	}

	// Sheets
	ct := "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
	for i := range w.sheets {
		if err := write(`<Override PartName="/xl/worksheets/sheet` + strconv.Itoa(i+1) + `.xml" ContentType="` + ct + `"/>` + "\n"); err != nil {
			return err
		}
	}

	// Styles
	if err := write(`<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` + "\n"); err != nil {
		return err
	}

	// Shared strings
	if err := write(`<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` + "\n"); err != nil {
		return err
	}

	return write(`</Types>`)
}

// ──────────────────────────────────────────────
// Shared string writer (write-side only, minimal memory)
// ──────────────────────────────────────────────

// sharedStringWriter collects unique strings during writing.
// It uses a map for dedup but only stores each unique string once.
type sharedStringWriter struct {
	strings []string
	index   map[string]int
}

func newSharedStringWriter() *sharedStringWriter {
	return &sharedStringWriter{
		index: make(map[string]int),
	}
}

func (ss *sharedStringWriter) Add(s string) int {
	if idx, ok := ss.index[s]; ok {
		return idx
	}
	idx := len(ss.strings)
	ss.strings = append(ss.strings, s)
	ss.index[s] = idx
	return idx
}

func (ss *sharedStringWriter) writeXML(w io.Writer) error {
	var b bytes.Buffer
	// Guard against integer overflow on 32-bit systems
	if len(ss.strings) < 1<<20 {
		b.Grow(64 * len(ss.strings))
	}

	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	b.WriteByte('\n')
	b.WriteString(`<sst xmlns="`)
	b.WriteString(nsSpreadsheetML)
	b.WriteString(`" count="`)
	b.Write(strconv.AppendInt(b.AvailableBuffer(), int64(len(ss.strings)), 10))
	b.WriteString(`" uniqueCount="`)
	b.Write(strconv.AppendInt(b.AvailableBuffer(), int64(len(ss.strings)), 10))
	b.WriteString(`">`)

	// Flush in chunks to avoid building one giant buffer
	const flushThreshold = 32 * 1024
	for _, s := range ss.strings {
		b.WriteString(`<si><t>`)
		appendEscapedXML(&b, s)
		b.WriteString(`</t></si>`)

		if b.Len() >= flushThreshold {
			if _, err := w.Write(b.Bytes()); err != nil {
				return err
			}
			b.Reset()
		}
	}

	b.WriteString(`</sst>`)
	_, err := w.Write(b.Bytes())
	return err
}

func appendEscapedXML(b *bytes.Buffer, s string) {
	last := 0
	for i := 0; i < len(s); i++ {
		var esc string
		switch s[i] {
		case '&':
			esc = "&amp;"
		case '<':
			esc = "&lt;"
		case '>':
			esc = "&gt;"
		default:
			continue
		}
		b.WriteString(s[last:i])
		b.WriteString(esc)
		last = i + 1
	}
	b.WriteString(s[last:])
}

