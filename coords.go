package xlsxlite

import (
	"fmt"
	"strconv"
	"strings"
)

// ColIndexToLetter converts a 0-based column index to Excel column letters.
// 0 → "A", 25 → "Z", 26 → "AA", etc.
func ColIndexToLetter(col int) string {
	var result []byte
	for col >= 0 {
		result = append([]byte{byte('A' + col%26)}, result...)
		col = col/26 - 1
	}
	return string(result)
}

// LetterToColIndex converts Excel column letters to a 0-based column index.
// "A" → 0, "Z" → 25, "AA" → 26, etc.
func LetterToColIndex(letters string) int {
	letters = strings.ToUpper(letters)
	result := 0
	for _, c := range letters {
		result = result*26 + int(c-'A') + 1
	}
	return result - 1
}

// CellRef creates a cell reference like "A1" from 0-based col and 1-based row.
func CellRef(col, row int) string {
	return ColIndexToLetter(col) + strconv.Itoa(row)
}

// ParseCellRef parses a cell reference like "A1" into a 0-based column index
// and 1-based row number. Returns an error for invalid references.
func ParseCellRef(ref string) (col, row int, err error) {
	i := 0
	for i < len(ref) && ref[i] >= 'A' && ref[i] <= 'Z' {
		i++
	}
	if i == 0 {
		// Try lowercase
		for i < len(ref) && ref[i] >= 'a' && ref[i] <= 'z' {
			i++
		}
	}
	if i == 0 || i == len(ref) {
		return 0, 0, fmt.Errorf("invalid cell reference: %q", ref)
	}
	col = LetterToColIndex(ref[:i])
	row, err = strconv.Atoi(ref[i:])
	if err != nil {
		return 0, 0, fmt.Errorf("invalid cell reference: %q", ref)
	}
	return col, row, nil
}

// RangeRef creates a range reference like "A1:D5" from 0-based column indices
// and 1-based row numbers.
func RangeRef(startCol, startRow, endCol, endRow int) string {
	return CellRef(startCol, startRow) + ":" + CellRef(endCol, endRow)
}
