package xlsxlite

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// Reader reads XLSX files using streaming XML parsing.
// The shared string table is loaded into memory (unavoidable for random access),
// but worksheet rows are read one-at-a-time via a streaming XML decoder.
type Reader struct {
	zr            *zip.Reader
	closer        io.Closer
	sharedStrings []string
	sheets        []readerSheet
	styles        *readerStyles
}

type readerSheet struct {
	name string
	path string // e.g. "xl/worksheets/sheet1.xml"
}

// readerStyles holds parsed style info needed to detect date cells.
type readerStyles struct {
	numFmts map[int]string // numFmtId → formatCode
	xfs     []readerXf     // cellXfs entries
}

type readerXf struct {
	numFmtID int
}

// OpenReader opens an XLSX file for streaming reading from an io.ReaderAt
// (e.g. *os.File, *bytes.Reader). It parses the workbook structure, loads the
// shared string table, and reads style definitions. Use OpenFile for a simpler
// file-based API.
func OpenReader(r io.ReaderAt, size int64) (*Reader, error) {
	zr, err := zip.NewReader(r, size)
	if err != nil {
		return nil, fmt.Errorf("open zip: %w", err)
	}

	reader := &Reader{zr: zr}

	// Parse workbook relationships to find sheet paths
	sheetPaths, err := reader.parseWorkbookRels()
	if err != nil {
		return nil, fmt.Errorf("parse rels: %w", err)
	}

	// Parse workbook to get sheet names + order
	if err := reader.parseWorkbook(sheetPaths); err != nil {
		return nil, fmt.Errorf("parse workbook: %w", err)
	}

	// Load shared strings (streamed parsing, but stored for random access)
	if err := reader.loadSharedStrings(); err != nil {
		return nil, fmt.Errorf("load shared strings: %w", err)
	}

	// Load styles (small file, needed to detect dates)
	if err := reader.loadStyles(); err != nil {
		// Non-fatal: some XLSX files don't have styles
		reader.styles = &readerStyles{}
	}

	return reader, nil
}

// Close releases any resources held by the Reader.
// If the Reader was created via OpenFile, this closes the underlying file.
func (r *Reader) Close() error {
	if r.closer != nil {
		return r.closer.Close()
	}
	return nil
}

// SheetNames returns the names of all sheets in workbook order.
func (r *Reader) SheetNames() []string {
	names := make([]string, len(r.sheets))
	for i, s := range r.sheets {
		names[i] = s.name
	}
	return names
}

// SheetCount returns the number of sheets.
func (r *Reader) SheetCount() int {
	return len(r.sheets)
}

// ──────────────────────────────────────────────
// Row iterator (streaming, O(1) memory per row)
// ──────────────────────────────────────────────

// RowIterator streams rows from a worksheet one at a time.
// Use Next to advance, Row to get the current row, and Err to check for errors.
// The caller must call Close when done to release the underlying zip entry reader.
type RowIterator struct {
	reader   *Reader
	decoder  *xml.Decoder
	closer   io.ReadCloser
	current  *Row
	err      error
	done     bool
	valBuf   strings.Builder // reusable buffer for cell value parsing
	cellsBuf []Cell          // reusable buffer for collecting cells per row
}

// OpenSheet returns a RowIterator for the sheet with the given name.
// The caller must call RowIterator.Close when done reading.
func (r *Reader) OpenSheet(name string) (*RowIterator, error) {
	for _, s := range r.sheets {
		if s.name == name {
			return r.openSheetByPath(s.path)
		}
	}
	return nil, fmt.Errorf("sheet %q not found", name)
}

// OpenSheetByIndex returns a RowIterator for the sheet at the given 0-based index.
// The caller must call RowIterator.Close when done reading.
func (r *Reader) OpenSheetByIndex(index int) (*RowIterator, error) {
	if index < 0 || index >= len(r.sheets) {
		return nil, fmt.Errorf("sheet index %d out of range (0-%d)", index, len(r.sheets)-1)
	}
	return r.openSheetByPath(r.sheets[index].path)
}

func (r *Reader) openSheetByPath(path string) (*RowIterator, error) {
	f := r.findFile(path)
	if f == nil {
		return nil, fmt.Errorf("sheet file %q not found in archive", path)
	}

	rc, err := f.Open()
	if err != nil {
		return nil, fmt.Errorf("open sheet file: %w", err)
	}

	lr := io.LimitReader(rc, MaxDecompressedSize)
	return &RowIterator{
		reader:  r,
		decoder: xml.NewDecoder(lr),
		closer:  rc,
	}, nil
}

// Next advances to the next row. Returns false when there are no more rows
// or an error occurred. After Next returns false, call Err to check for errors.
func (it *RowIterator) Next() bool {
	if it.done {
		return false
	}

	for {
		tok, err := it.decoder.Token()
		if err != nil {
			if err != io.EOF {
				it.err = err
			}
			it.done = true
			return false
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "row" {
				row, err := it.parseRow(t)
				if err != nil {
					it.err = err
					it.done = true
					return false
				}
				it.current = row
				return true
			}
		}
	}
}

// Row returns the current row. Only valid after Next() returns true.
func (it *RowIterator) Row() *Row {
	return it.current
}

// Err returns the first error encountered during iteration, or nil if
// iteration completed successfully.
func (it *RowIterator) Err() error {
	return it.err
}

// Close releases resources held by the iterator. It is safe to call Close
// multiple times. After Close, Next will always return false.
func (it *RowIterator) Close() error {
	it.done = true
	if it.closer != nil {
		return it.closer.Close()
	}
	return nil
}

func (it *RowIterator) parseRow(start xml.StartElement) (*Row, error) {
	row := &Row{}

	// Get row number
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "r":
			if n, err := strconv.Atoi(attr.Value); err == nil {
				row.RowIndex = n
			}
		case "ht":
			if h, err := strconv.ParseFloat(attr.Value, 64); err == nil {
				row.Height = h
			}
		}
	}

	// Reuse cells buffer, clear from previous row
	cells := it.cellsBuf[:0]

	for {
		tok, err := it.decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "c" {
				col, cell, err := it.parseCell(t)
				if err != nil {
					return nil, err
				}
				// Grow slice if needed
				if col >= MaxColumns {
					continue // skip columns beyond safety limit
				}
				if col >= len(cells) {
					for len(cells) <= col {
						cells = append(cells, Cell{})
					}
				}
				cells[col] = cell
			}
		case xml.EndElement:
			if t.Name.Local == "row" {
				// Copy to a new slice so the row owns its data
				row.Cells = make([]Cell, len(cells))
				copy(row.Cells, cells)
				// Save buffer for reuse (keep capacity)
				it.cellsBuf = cells
				return row, nil
			}
		}
	}
}

func (it *RowIterator) parseCell(start xml.StartElement) (int, Cell, error) {
	var (
		cellRef   string
		cellType  string
		cellStyle int
	)

	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "r":
			cellRef = attr.Value
		case "t":
			cellType = attr.Value
		case "s":
			if n, err := strconv.Atoi(attr.Value); err == nil {
				cellStyle = n
			}
		}
	}

	col := 0
	if cellRef != "" {
		c, _, err := ParseCellRef(cellRef)
		if err == nil {
			col = c
		}
	}

	// Read the value element using a reusable builder
	it.valBuf.Reset()
	var inValue bool

	for {
		tok, err := it.decoder.Token()
		if err != nil {
			return col, Cell{}, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "v" || t.Name.Local == "t" {
				inValue = true
			}
		case xml.CharData:
			if inValue {
				it.valBuf.Write(t)
			}
		case xml.EndElement:
			if t.Name.Local == "v" || t.Name.Local == "t" {
				inValue = false
			}
			if t.Name.Local == "c" {
				cell := it.resolveCell(it.valBuf.String(), cellType, cellStyle)
				return col, cell, nil
			}
		}
	}
}

func (it *RowIterator) resolveCell(rawValue, cellType string, styleID int) Cell {
	cell := Cell{StyleID: styleID}

	if rawValue == "" {
		cell.Type = CellTypeEmpty
		return cell
	}

	switch cellType {
	case "s": // shared string
		idx, err := strconv.Atoi(rawValue)
		if err == nil && idx >= 0 && idx < len(it.reader.sharedStrings) {
			cell.Value = it.reader.sharedStrings[idx]
			cell.Type = CellTypeString
		}
	case "str", "inlineStr": // inline string
		cell.Value = rawValue
		cell.Type = CellTypeString
	case "b": // boolean
		cell.Value = rawValue == "1"
		cell.Type = CellTypeBool
	default: // number or date
		f, err := strconv.ParseFloat(rawValue, 64)
		if err != nil {
			cell.Value = rawValue
			cell.Type = CellTypeString
		} else {
			// Check if this is a date format via styles
			if it.reader.isDateStyle(styleID) {
				cell.Value = ExcelSerialToTime(f)
				cell.Type = CellTypeDate
			} else {
				cell.Value = f
				cell.Type = CellTypeNumber
			}
		}
	}

	return cell
}

// ──────────────────────────────────────────────
// Internal parsing helpers
// ──────────────────────────────────────────────

func (r *Reader) findFile(path string) *zip.File {
	for _, f := range r.zr.File {
		if f.Name == path {
			return f
		}
	}
	return nil
}

func (r *Reader) parseWorkbookRels() (map[string]string, error) {
	f := r.findFile("xl/_rels/workbook.xml.rels")
	if f == nil {
		return nil, fmt.Errorf("workbook.xml.rels not found")
	}

	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	type Rel struct {
		ID     string `xml:"Id,attr"`
		Type   string `xml:"Type,attr"`
		Target string `xml:"Target,attr"`
	}
	type Rels struct {
		Relationships []Rel `xml:"Relationship"`
	}

	lr := io.LimitReader(rc, MaxDecompressedSize)
	var rels Rels
	if err := xml.NewDecoder(lr).Decode(&rels); err != nil {
		return nil, err
	}

	result := make(map[string]string)
	for _, rel := range rels.Relationships {
		if strings.Contains(rel.Type, "worksheet") {
			target := rel.Target
			// Sanitize: reject absolute paths and path traversal
			if strings.HasPrefix(target, "/") || strings.Contains(target, "..") {
				continue
			}
			if !strings.HasPrefix(target, "xl/") {
				target = "xl/" + target
			}
			result[rel.ID] = target
		}
	}
	return result, nil
}

func (r *Reader) parseWorkbook(sheetPaths map[string]string) error {
	f := r.findFile("xl/workbook.xml")
	if f == nil {
		return fmt.Errorf("workbook.xml not found")
	}

	rc, err := f.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	lr := io.LimitReader(rc, MaxDecompressedSize)
	decoder := xml.NewDecoder(lr)
	for {
		tok, err := decoder.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "sheet" {
				var name, rid string
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "name":
						name = attr.Value
					case "id":
						rid = attr.Value
					}
				}
				if path, ok := sheetPaths[rid]; ok {
					r.sheets = append(r.sheets, readerSheet{name: name, path: path})
				}
			}
		}
	}
	return nil
}

// loadSharedStrings streams the shared string table into a string slice.
// This is the one part we can't fully stream (cells reference by index),
// but we parse it with a streaming XML decoder to avoid building a DOM.
func (r *Reader) loadSharedStrings() error {
	f := r.findFile("xl/sharedStrings.xml")
	if f == nil {
		// No shared strings is valid (e.g., numbers-only sheet)
		return nil
	}

	rc, err := f.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	lr := io.LimitReader(rc, MaxDecompressedSize)
	decoder := xml.NewDecoder(lr)
	var (
		inSI    bool
		inT     bool
		current strings.Builder
	)

	for {
		tok, err := decoder.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "si":
				inSI = true
				current.Reset()
			case "t":
				if inSI {
					inT = true
				}
			}
		case xml.CharData:
			if inT {
				current.Write(t)
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "t":
				inT = false
			case "si":
				inSI = false
				if len(r.sharedStrings) >= MaxSharedStrings {
					return fmt.Errorf("shared string table exceeds limit of %d entries", MaxSharedStrings)
				}
				r.sharedStrings = append(r.sharedStrings, current.String())
			}
		}
	}

	return nil
}

func (r *Reader) loadStyles() error {
	f := r.findFile("xl/styles.xml")
	if f == nil {
		return fmt.Errorf("styles.xml not found")
	}

	rc, err := f.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	r.styles = &readerStyles{
		numFmts: make(map[int]string),
	}

	lr := io.LimitReader(rc, MaxDecompressedSize)
	decoder := xml.NewDecoder(lr)
	var inCellXfs bool

	for {
		tok, err := decoder.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "numFmt":
				var id int
				var code string
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "numFmtId":
						id, _ = strconv.Atoi(attr.Value)
					case "formatCode":
						code = attr.Value
					}
				}
				r.styles.numFmts[id] = code

			case "cellXfs":
				inCellXfs = true

			case "xf":
				if inCellXfs {
					var numFmtID int
					for _, attr := range t.Attr {
						if attr.Name.Local == "numFmtId" {
							numFmtID, _ = strconv.Atoi(attr.Value)
						}
					}
					r.styles.xfs = append(r.styles.xfs, readerXf{numFmtID: numFmtID})
				}
			}
		case xml.EndElement:
			if t.Name.Local == "cellXfs" {
				inCellXfs = false
			}
		}
	}

	return nil
}

// isDateStyle checks if a style index corresponds to a date format.
func (r *Reader) isDateStyle(styleID int) bool {
	if r.styles == nil || styleID < 0 || styleID >= len(r.styles.xfs) {
		return false
	}
	numFmtID := r.styles.xfs[styleID].numFmtID
	return isDateNumFmtID(numFmtID, r.styles.numFmts)
}

// isDateNumFmtID checks if a number format ID represents a date/time format.
// Built-in date format IDs: 14-22, 27-36, 45-47, 50-58
func isDateNumFmtID(id int, customFmts map[int]string) bool {
	// Built-in date formats
	switch {
	case id >= 14 && id <= 22:
		return true
	case id >= 27 && id <= 36:
		return true
	case id >= 45 && id <= 47:
		return true
	case id >= 50 && id <= 58:
		return true
	}

	// Check custom format codes for date patterns
	if code, ok := customFmts[id]; ok {
		lower := strings.ToLower(code)
		// Common date/time tokens
		for _, token := range []string{"yy", "mm", "dd", "hh", "ss", "am/pm"} {
			if strings.Contains(lower, token) {
				return true
			}
		}
	}

	return false
}
