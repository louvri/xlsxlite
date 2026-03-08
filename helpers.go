package xlsxlite

import (
	"fmt"
	"os"
)

// OpenFile opens an XLSX file from disk for streaming reading.
// The caller must call Reader.Close when done to release the underlying file handle.
func OpenFile(path string) (*Reader, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("open file: %w", err)
	}

	info, err := f.Stat()
	if err != nil {
		f.Close()
		return nil, fmt.Errorf("stat file: %w", err)
	}

	r, err := OpenReader(f, info.Size())
	if err != nil {
		f.Close()
		return nil, err
	}
	r.closer = f
	return r, nil
}

// CreateFile creates a new XLSX file on disk for streaming writing.
// Returns the Writer and the underlying os.File. Call Writer.Close() first
// to finalize the XLSX package, then close the file.
func CreateFile(path string) (*Writer, *os.File, error) {
	f, err := os.Create(path)
	if err != nil {
		return nil, nil, fmt.Errorf("create file: %w", err)
	}

	w := NewWriter(f)
	return w, f, nil
}

// ──────────────────────────────────────────────
// Quick cell constructors
// ──────────────────────────────────────────────

// StringCell creates a string cell. The value is stored as a shared string
// in the XLSX file. No type coercion is performed, so leading zeros and
// numeric-looking strings are preserved as-is.
func StringCell(v string) Cell {
	return Cell{Value: v, Type: CellTypeString}
}

// NumberCell creates a numeric cell from a float64 value.
func NumberCell(v float64) Cell {
	return Cell{Value: v, Type: CellTypeNumber}
}

// IntCell creates a numeric cell from an int value.
func IntCell(v int) Cell {
	return Cell{Value: v, Type: CellTypeNumber}
}

// BoolCell creates a boolean cell.
func BoolCell(v bool) Cell {
	return Cell{Value: v, Type: CellTypeBool}
}

// DateCell creates a date cell with the given style. The styleID should reference
// a style with a date NumberFormat (e.g. "yyyy-mm-dd") so Excel displays it
// as a date instead of a raw serial number.
func DateCell(v any, styleID int) Cell {
	return Cell{Value: v, Type: CellTypeDate, StyleID: styleID}
}

// StyledCell creates a cell with a specific style. The cell type is auto-detected
// from the Go value type: string → CellTypeString, float64/float32/int/int64 →
// CellTypeNumber, bool → CellTypeBool, anything else → CellTypeEmpty.
func StyledCell(v any, styleID int) Cell {
	t := CellTypeEmpty
	switch v.(type) {
	case string:
		t = CellTypeString
	case float64, float32, int, int64:
		t = CellTypeNumber
	case bool:
		t = CellTypeBool
	}
	return Cell{Value: v, Type: t, StyleID: styleID}
}

// EmptyCell returns an empty cell (useful for padding in rows).
func EmptyCell() Cell {
	return Cell{Type: CellTypeEmpty}
}

// MakeRow is a convenience to create a Row from a variadic list of values.
// Accepted types: string, float64, float32, int, int64, bool, Cell, and nil.
// Strings are stored as shared strings with no type coercion (leading zeros are
// preserved). Any other type is converted to a string via fmt.Sprintf.
func MakeRow(values ...any) Row {
	cells := make([]Cell, len(values))
	for i, v := range values {
		switch val := v.(type) {
		case Cell:
			cells[i] = val
		case string:
			cells[i] = StringCell(val)
		case float64:
			cells[i] = NumberCell(val)
		case float32:
			cells[i] = NumberCell(float64(val))
		case int:
			cells[i] = IntCell(val)
		case int64:
			cells[i] = Cell{Value: val, Type: CellTypeNumber}
		case bool:
			cells[i] = BoolCell(val)
		case nil:
			cells[i] = EmptyCell()
		default:
			cells[i] = StringCell(fmt.Sprintf("%v", val))
		}
	}
	return Row{Cells: cells}
}
