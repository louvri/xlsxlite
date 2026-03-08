// Package xlsxlite is a lightweight, memory-efficient XLSX read/write library.
//
// Unlike DOM-based libraries that load entire worksheet XML into memory,
// xlsxlite uses streaming XML parsing (encoding/xml.Decoder) for reading
// and streaming XML writing for output. This keeps memory usage proportional
// to a single row rather than the entire file, enabling processing of
// arbitrarily large spreadsheets.
//
// Design principles:
//   - Zero external dependencies (stdlib only)
//   - Streaming read via row iterator (cursor pattern)
//   - Streaming write via row-by-row flush
//   - Shared string table with chunked/indexed access for reads
//   - Style support: fonts, fills, borders, alignment, number formats
//   - Merge cells, column widths, row heights
//   - No formula engine, no charts, no images (keeps it lean)
package xlsxlite

import (
	"encoding/xml"
	"time"
)

// ──────────────────────────────────────────────
// Cell value types
// ──────────────────────────────────────────────

// CellType represents the data type of a cell.
type CellType int

const (
	CellTypeEmpty        CellType = iota // empty or nil cell
	CellTypeString                       // string value (stored as shared string)
	CellTypeNumber                       // numeric value (float64, int, int64, or float32)
	CellTypeBool                         // boolean value
	CellTypeDate                         // date/time value (stored as Excel serial number)
	CellTypeInlineString                 // inline string (not shared)
)

// Cell represents a single cell value with optional style.
// The Value field holds the Go value (string, float64, int, int64, float32, bool,
// time.Time, or nil). The Type field must match the value's kind. Use the cell
// constructors (StringCell, NumberCell, etc.) or MakeRow for convenience.
type Cell struct {
	Value   any // string, float64, bool, time.Time, or nil
	Type    CellType
	StyleID int // index into the styles table; 0 = default
}

// Row represents a single row of cells.
// Set RowIndex to override the auto-incremented row number (1-based).
// Set Height to apply a custom row height in points.
type Row struct {
	Cells    []Cell
	Height   float64 // custom row height; 0 = default
	RowIndex int     // 1-based row number
}

// ──────────────────────────────────────────────
// Style types
// ──────────────────────────────────────────────

// Font describes a font style. All fields are optional; zero values are omitted
// from the output. Color is an ARGB hex string (e.g. "FF000000" for black).
type Font struct {
	Name      string
	Size      float64
	Bold      bool
	Italic    bool
	Underline bool
	Color     string // ARGB hex, e.g. "FF000000"
}

// Fill describes a cell fill/background. Set Type to "pattern" and Pattern to
// "solid" for a solid color fill. FgColor and BgColor are ARGB hex strings.
type Fill struct {
	Type    string // "pattern" or "none"
	Pattern string // "solid", "gray125", etc.
	FgColor string // ARGB hex
	BgColor string // ARGB hex
}

// BorderEdge describes one edge of a cell border.
type BorderEdge struct {
	Style string // "thin", "medium", "thick", "dashed", etc.
	Color string // ARGB hex
}

// Border describes all four borders of a cell.
type Border struct {
	Left   BorderEdge
	Right  BorderEdge
	Top    BorderEdge
	Bottom BorderEdge
}

// Alignment describes cell alignment.
type Alignment struct {
	Horizontal string // "left", "center", "right", "fill", "justify"
	Vertical   string // "top", "center", "bottom"
	WrapText   bool
}

// Style combines font, fill, border, alignment and number format.
// All fields are optional (nil pointers and empty strings are omitted).
// Register styles via StyleSheet.AddStyle before writing rows.
type Style struct {
	Font         *Font
	Fill         *Fill
	Border       *Border
	Alignment    *Alignment
	NumberFormat string // custom format string, e.g. "yyyy-mm-dd"
}

// ──────────────────────────────────────────────
// Merge cell
// ──────────────────────────────────────────────

// MergeCell represents a merged cell range.
// Columns are 0-based, rows are 1-based (matching Excel conventions).
type MergeCell struct {
	StartCol int // 0-based column
	StartRow int // 1-based row
	EndCol   int // 0-based column
	EndRow   int // 1-based row
}

// ──────────────────────────────────────────────
// Sheet metadata
// ──────────────────────────────────────────────

// SheetConfig holds configuration for a worksheet.
// Pass this to Writer.NewSheet to create a new sheet with the given settings.
type SheetConfig struct {
	Name       string
	ColWidths  map[int]float64  // 0-based col index → width
	MergeCells []MergeCell
	FreezeRow  int // freeze panes: first N rows
	FreezeCol  int // freeze panes: first N cols
}

// ──────────────────────────────────────────────
// Safety limits
// ──────────────────────────────────────────────

const (
	// MaxColumns is the maximum number of columns supported (Excel's limit: XFD = 16384).
	MaxColumns = 16384
	// MaxSharedStrings is the maximum number of shared strings allowed when reading.
	MaxSharedStrings = 10_000_000
	// MaxDecompressedSize is the maximum decompressed size of any single zip entry (256 MB).
	MaxDecompressedSize = 256 << 20
	// MaxGCSDownloadSize is the maximum size for GCS file downloads (512 MB).
	MaxGCSDownloadSize = 512 << 20
)

// ──────────────────────────────────────────────
// XML namespace constants
// ──────────────────────────────────────────────

const (
	nsSpreadsheetML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	nsRelationships = "http://schemas.openxmlformats.org/package/2006/relationships"
	nsOfficeDocRels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	nsContentTypes  = "http://schemas.openxmlformats.org/package/2006/content-types"

	relTypeWorksheet     = nsOfficeDocRels + "/worksheet"
	relTypeStyles        = nsOfficeDocRels + "/styles"
	relTypeSharedStrings = nsOfficeDocRels + "/sharedStrings"
)

// ──────────────────────────────────────────────
// Internal XML structs (minimal, for streaming)
// ──────────────────────────────────────────────

// xmlWorkbook is the minimal workbook.xml structure.
type xmlWorkbook struct {
	XMLName xml.Name       `xml:"workbook"`
	Xmlns   string         `xml:"xmlns,attr"`
	XmlnsR  string         `xml:"xmlns:r,attr"`
	Sheets  xmlWorkbookSheets `xml:"sheets"`
}

type xmlWorkbookSheets struct {
	Sheets []xmlSheet `xml:"sheet"`
}

type xmlSheet struct {
	Name    string `xml:"name,attr"`
	SheetID int    `xml:"sheetId,attr"`
	RID     string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

// ──────────────────────────────────────────────
// Helper: Excel date epoch
// ──────────────────────────────────────────────

// excelEpoch is December 30, 1899 (Excel's date system).
var excelEpoch = time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)

// TimeToExcelSerial converts a time.Time to an Excel serial date number.
// The serial number represents days since December 30, 1899 (Excel's epoch).
func TimeToExcelSerial(t time.Time) float64 {
	duration := t.Sub(excelEpoch)
	return duration.Hours() / 24.0
}

// ExcelSerialToTime converts an Excel serial date number back to time.Time.
// Used internally by the reader to convert numeric cells with date styles.
func ExcelSerialToTime(serial float64) time.Time {
	days := time.Duration(serial * 24 * float64(time.Hour))
	return excelEpoch.Add(days)
}
