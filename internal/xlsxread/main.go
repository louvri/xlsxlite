package main

import (
	"fmt"
	"os"
	"strings"
	"time"

	"github.com/louvri/xlsxlite"
)

const maxSampleRows = 5
const maxDisplayCols = 10
const maxCellWidth = 40

func main() {
	if len(os.Args) < 2 {
		fmt.Fprintf(os.Stderr, "Usage: %s <file.xlsx>\n", os.Args[0])
		os.Exit(1)
	}
	path := os.Args[1]

	start := time.Now()
	reader, err := xlsxlite.OpenFile(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening file: %v\n", err)
		os.Exit(1)
	}
	defer reader.Close()
	openElapsed := time.Since(start)

	fmt.Println()
	printKV("File", path)
	printKV("Open time", openElapsed.String())
	printKV("Sheets", fmt.Sprintf("%d", reader.SheetCount()))
	fmt.Println()

	for i, name := range reader.SheetNames() {
		fmt.Printf("═══ Sheet %d: %q ═══\n\n", i, name)

		sheetStart := time.Now()
		iter, err := reader.OpenSheet(name)
		if err != nil {
			fmt.Fprintf(os.Stderr, "  Error opening sheet: %v\n", err)
			continue
		}

		totalRows := 0
		maxCols := 0
		nonEmptyTotal := 0

		// Collect sample rows for table display
		var samples []sampleRow

		for iter.Next() {
			row := iter.Row()
			totalRows++
			cols := len(row.Cells)
			if cols > maxCols {
				maxCols = cols
			}
			for _, cell := range row.Cells {
				if cell.Type != xlsxlite.CellTypeEmpty {
					nonEmptyTotal++
				}
			}
			if totalRows <= maxSampleRows {
				displayCols := cols
				if displayCols > maxDisplayCols {
					displayCols = maxDisplayCols
				}
				sr := sampleRow{index: row.RowIndex}
				for j := 0; j < displayCols; j++ {
					sr.cells = append(sr.cells, cellDisplay(row.Cells[j]))
					sr.types = append(sr.types, cellTypeName(row.Cells[j].Type))
				}
				samples = append(samples, sr)
			}
		}
		iter.Close()

		if err := iter.Err(); err != nil {
			fmt.Fprintf(os.Stderr, "  Iteration error: %v\n", err)
		}

		sheetElapsed := time.Since(sheetStart)

		// Print sample rows as a table
		if len(samples) > 0 {
			printTable(samples, maxCols)
			fmt.Println()
		}

		// Print summary
		printKV("Rows", fmt.Sprintf("%d", totalRows))
		printKV("Max columns", fmt.Sprintf("%d", maxCols))
		printKV("Non-empty cells", fmt.Sprintf("%d", nonEmptyTotal))
		printKV("Read time", sheetElapsed.String())
		fmt.Println()
	}

	printKV("Total elapsed", time.Since(start).String())
	fmt.Println()
}

func printKV(key, value string) {
	fmt.Printf("  %-18s %s\n", key+":", value)
}

func printTable(samples []sampleRow, totalCols int) {
	if len(samples) == 0 {
		return
	}

	numCols := len(samples[0].cells)
	truncated := totalCols > maxDisplayCols

	// Calculate column widths (including the row index column)
	colWidths := make([]int, numCols)
	for _, sr := range samples {
		for j, cell := range sr.cells {
			if len(cell) > colWidths[j] {
				colWidths[j] = len(cell)
			}
		}
	}
	// Cap column widths
	for j := range colWidths {
		if colWidths[j] > maxCellWidth {
			colWidths[j] = maxCellWidth
		}
		if colWidths[j] < 3 {
			colWidths[j] = 3
		}
	}

	// Row index column width
	rowIdxWidth := 5 // "Row #"
	for _, sr := range samples {
		w := len(fmt.Sprintf("%d", sr.index))
		if w > rowIdxWidth {
			rowIdxWidth = w
		}
	}

	// Print type header — use last sample row's types (more likely to be data, not headers)
	typeRow := samples[len(samples)-1]
	fmt.Printf("  %-*s", rowIdxWidth, "")
	for j := 0; j < numCols; j++ {
		fmt.Printf("  %-*s", colWidths[j], fmt.Sprintf("[%s]", typeRow.types[j]))
	}
	if truncated {
		fmt.Printf("  ...")
	}
	fmt.Println()

	// Print separator
	fmt.Printf("  %s", strings.Repeat("─", rowIdxWidth))
	for j := 0; j < numCols; j++ {
		fmt.Printf("──%s", strings.Repeat("─", colWidths[j]))
	}
	if truncated {
		fmt.Printf("─────")
	}
	fmt.Println()

	// Print data rows
	for _, sr := range samples {
		fmt.Printf("  %-*d", rowIdxWidth, sr.index)
		for j, cell := range sr.cells {
			display := cell
			if len(display) > maxCellWidth {
				display = display[:maxCellWidth-1] + "…"
			}
			fmt.Printf("  %-*s", colWidths[j], display)
		}
		if truncated {
			fmt.Printf("  ...")
		}
		fmt.Println()
	}
}

type sampleRow struct {
	index int
	cells []string
	types []string
}

func cellDisplay(c xlsxlite.Cell) string {
	if c.Type == xlsxlite.CellTypeEmpty || c.Value == nil {
		return ""
	}
	return fmt.Sprintf("%v", c.Value)
}

func cellTypeName(t xlsxlite.CellType) string {
	switch t {
	case xlsxlite.CellTypeEmpty:
		return "empty"
	case xlsxlite.CellTypeString:
		return "str"
	case xlsxlite.CellTypeNumber:
		return "num"
	case xlsxlite.CellTypeBool:
		return "bool"
	case xlsxlite.CellTypeDate:
		return "date"
	case xlsxlite.CellTypeInlineString:
		return "inline"
	default:
		return "?"
	}
}
