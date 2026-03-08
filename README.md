# xlsxlite

[![CI](https://github.com/louvri/xlsxlite/actions/workflows/ci.yml/badge.svg)](https://github.com/louvri/xlsxlite/actions/workflows/ci.yml)

A lightweight, memory-efficient XLSX read/write library for Go. The core library has **zero external dependencies** — stdlib only. An optional `gcs` subpackage adds Google Cloud Storage support.

## Why?

Most Go XLSX libraries load entire worksheet XML DOMs into memory. A 7MB XLSX can consume ~1GB of RAM; 1M rows can exceed 3.6GB allocations. That's fine for small files, but breaks down for large-scale ETL, report generation, or memory-constrained environments (like your GKE pods).

**xlsxlite** takes a streaming approach:

| | DOM-based libraries | xlsxlite |
|---|---|---|
| **Architecture** | Full DOM in memory | Streaming XML read/write |
| **Dependencies** | Multiple | 0 for core (stdlib only); `gcs` subpackage adds `cloud.google.com/go/storage` |
| **Features** | Full (charts, formulas, images, pivot tables) | Essential (cells, styles, merges, col widths, freeze panes) |
| **Write pattern** | Random access (any cell, any order) | Sequential (row-by-row, streaming) |
| **Read pattern** | Full DOM or streaming option | Streaming iterator only |

### Memory model

```
DOM-based:  Load entire XML → Build struct tree → Hold in memory → Write all at once
xlsxlite:   Stream XML token-by-token → Process one row → Discard → Next row
```

Write-side memory is O(unique strings) rather than O(total cells). Read-side memory is O(shared string table + 1 row).

## Install

```bash
go get github.com/louvri/xlsxlite
```

For optional Google Cloud Storage support (separate module, won't pollute your dependency tree unless you import it):

```bash
go get github.com/louvri/xlsxlite/gcs
```

## Quick Start

### Writing

```go
package main

import (
    "os"
    "github.com/louvri/xlsxlite"
)

func main() {
    f, _ := os.Create("report.xlsx")
    defer f.Close()

    w := xlsxlite.NewWriter(f)

    // Register styles before writing rows
    ss := w.StyleSheet()
    headerStyle := ss.AddStyle(xlsxlite.Style{
        Font: &xlsxlite.Font{Bold: true, Size: 12, Color: "FFFFFFFF"},
        Fill: &xlsxlite.Fill{Type: "pattern", Pattern: "solid", FgColor: "FF4472C4"},
        Alignment: &xlsxlite.Alignment{Horizontal: "center"},
    })

    // Create sheet with configuration
    sw, _ := w.NewSheet(xlsxlite.SheetConfig{
        Name:      "Sales Report",
        ColWidths: map[int]float64{0: 25, 1: 15, 2: 15},
        FreezeRow: 1,
        MergeCells: []xlsxlite.MergeCell{
            {StartCol: 0, StartRow: 1, EndCol: 2, EndRow: 1},
        },
    })

    // Write header row with styles
    sw.WriteRow(xlsxlite.Row{
        Cells: []xlsxlite.Cell{
            xlsxlite.StyledCell("Product", headerStyle),
            xlsxlite.StyledCell("Revenue", headerStyle),
            xlsxlite.StyledCell("Units", headerStyle),
        },
    })

    // Stream data rows — memory stays flat regardless of row count
    for i := 0; i < 1_000_000; i++ {
        sw.WriteRow(xlsxlite.MakeRow("Widget", 29.99, 150))
    }

    sw.Close() // finalize sheet
    w.Close()  // finalize XLSX package
}
```

### Reading

```go
package main

import (
    "fmt"
    "github.com/louvri/xlsxlite"
)

func main() {
    reader, _ := xlsxlite.OpenFile("report.xlsx")
    defer reader.Close()

    fmt.Println("Sheets:", reader.SheetNames())
    fmt.Println("Count:", reader.SheetCount())

    iter, _ := reader.OpenSheet("Sales Report")
    defer iter.Close()

    for iter.Next() {
        row := iter.Row()
        for _, cell := range row.Cells {
            fmt.Printf("%v\t", cell.Value)
        }
        fmt.Println()
    }

    if err := iter.Err(); err != nil {
        fmt.Println("Error:", err)
    }
}
```

You can also open sheets by index:

```go
iter, err := reader.OpenSheetByIndex(0) // 0-based
```

Or read from any `io.ReaderAt` (e.g. `*bytes.Reader`):

```go
reader, err := xlsxlite.OpenReader(readerAt, size)
```

### Google Cloud Storage

The `gcs` subpackage provides helpers for reading/writing XLSX files directly from/to GCS:

```go
import "github.com/louvri/xlsxlite/gcs"

// Reading — downloads object into memory, then streams rows
reader, err := gcs.OpenFile(ctx, client, "my-bucket", "reports/data.xlsx")

// Writing — streams directly to GCS object
w, closeFn, err := gcs.CreateFile(ctx, client, "my-bucket", "reports/output.xlsx")
// ... write sheets and rows ...
w.Close()      // finalize XLSX
closeFn()      // complete GCS upload
```

## API Reference

### Cell constructors

```go
xlsxlite.StringCell("hello")              // string (shared string, no type coercion)
xlsxlite.NumberCell(42.5)                  // float64
xlsxlite.IntCell(100)                      // int
xlsxlite.BoolCell(true)                    // bool
xlsxlite.DateCell(time.Now(), dateStyleID) // date with style
xlsxlite.StyledCell("value", styleID)      // any value with style (type auto-detected)
xlsxlite.EmptyCell()                       // empty cell (padding)
xlsxlite.MakeRow("a", 1, true, nil)       // mixed types in one row
```

`MakeRow` accepts: `string`, `float64`, `float32`, `int`, `int64`, `bool`, `Cell`, `nil`. Any other type is converted to a string via `fmt.Sprintf`.

### Cell values on read

When reading cells, values are returned as Go types based on the XLSX cell type:

| XLSX type | `Cell.Type` | `Cell.Value` Go type |
|---|---|---|
| Shared/inline string | `CellTypeString` | `string` |
| Number | `CellTypeNumber` | `float64` |
| Boolean | `CellTypeBool` | `bool` |
| Date (detected via style) | `CellTypeDate` | `time.Time` |
| Empty | `CellTypeEmpty` | `nil` |

### Leading zeros (phone numbers, zip codes, etc.)

Many XLSX libraries auto-detect numeric-looking strings and store them as numbers, which strips leading zeros. xlsxlite does **no auto-type coercion** — the Go type determines the cell type directly:

```go
// Safe: stored as a shared string, leading zeros preserved
xlsxlite.MakeRow("0812345678")
xlsxlite.StringCell("00501")

// Unsafe: stored as a number, leading zero is gone
xlsxlite.MakeRow(812345678)
xlsxlite.IntCell(501)
```

If your value is a string in Go, it stays a string in the XLSX. No surprises.

### Styles

Register styles before writing rows. Identical styles are deduplicated automatically.

```go
ss := w.StyleSheet()
id := ss.AddStyle(xlsxlite.Style{
    Font:         &xlsxlite.Font{Name: "Arial", Size: 11, Bold: true, Color: "FF000000"},
    Fill:         &xlsxlite.Fill{Type: "pattern", Pattern: "solid", FgColor: "FFFFFF00"},
    Border:       &xlsxlite.Border{Bottom: xlsxlite.BorderEdge{Style: "thin", Color: "FF000000"}},
    Alignment:    &xlsxlite.Alignment{Horizontal: "center", Vertical: "center", WrapText: true},
    NumberFormat: "yyyy-mm-dd",
})
```

All style fields are optional — nil pointers and empty strings are omitted. Colors use ARGB hex format (e.g. `"FF000000"` for black, `"FFFFFFFF"` for white).

### Sheet configuration

```go
xlsxlite.SheetConfig{
    Name:       "My Sheet",
    ColWidths:  map[int]float64{0: 20, 1: 15},  // 0-based col index → width
    MergeCells: []xlsxlite.MergeCell{
        {StartCol: 0, StartRow: 1, EndCol: 3, EndRow: 1}, // cols are 0-based, rows are 1-based
    },
    FreezeRow:  1,  // freeze first N rows
    FreezeCol:  2,  // freeze first N columns
}
```

### Row options

```go
// Custom row height
sw.WriteRow(xlsxlite.Row{
    Cells:  []xlsxlite.Cell{xlsxlite.StringCell("tall row")},
    Height: 30.0, // height in points
})

// Skip row numbers (e.g. for sparse data)
sw.WriteRow(xlsxlite.Row{
    Cells:    []xlsxlite.Cell{xlsxlite.StringCell("row 10")},
    RowIndex: 10, // 1-based, overrides auto-increment
})
```

### Coordinate helpers

```go
xlsxlite.ColIndexToLetter(0)          // "A"
xlsxlite.ColIndexToLetter(26)         // "AA"
xlsxlite.LetterToColIndex("AA")       // 26
xlsxlite.CellRef(0, 1)               // "A1" (0-based col, 1-based row)
xlsxlite.ParseCellRef("B3")          // col=1, row=3, err=nil
xlsxlite.RangeRef(0, 1, 3, 5)        // "A1:D5"
```

### Date helpers

```go
serial := xlsxlite.TimeToExcelSerial(time.Now()) // time.Time → Excel serial number
t := xlsxlite.ExcelSerialToTime(45823)           // Excel serial number → time.Time
```

### File helpers

```go
// Write to disk
w, f, err := xlsxlite.CreateFile("output.xlsx")
defer f.Close()
// ... write sheets ...
w.Close()

// Read from disk
reader, err := xlsxlite.OpenFile("input.xlsx")
defer reader.Close()
```

## Security

xlsxlite includes protections against common attack vectors in XLSX processing:

- **Zip bomb / decompression bomb**: All zip entry reads are capped at 256 MB (`MaxDecompressedSize`)
- **Shared string table DoS**: Limited to 10M entries (`MaxSharedStrings`)
- **Column expansion DoS**: Capped at 16,384 columns (`MaxColumns`, Excel's own limit)
- **XML injection**: All user-provided values written to XML attributes are escaped
- **Path traversal**: Relationship targets with `..` or absolute paths are rejected
- **GCS download size**: Capped at 512 MB (`MaxGCSDownloadSize`)

## Benchmarks

Measured on Apple M4 Pro, Go 1.25.0, `go test -bench=. -benchmem -count=3 -benchtime=5s`:

```
goos: darwin
goarch: arm64
cpu: Apple M4 Pro

BenchmarkWrite100kRows-12     51    116ms/op     52MB/op     800k allocs/op
BenchmarkRead100kRows-12      19    305ms/op    319MB/op     7.7M allocs/op
BenchmarkWrite1MRows-12        5   1134ms/op    499MB/op     7.0M allocs/op
```

- **Write 100k rows**: 116ms, 52 MB allocations (3 cells per row × 100k rows)
- **Read 100k rows**: 305ms, 319 MB allocations (dominated by shared string table)
- **Write 1M rows**: 1.13s, 499 MB allocations — scales linearly
- **100k rows file size**: 1.7 MB (compressed ZIP)

## What's NOT included (by design)

- Formula calculation engine
- Charts, sparklines, pivot tables
- Images / drawings
- Encryption / decryption
- Conditional formatting
- Data validation
- Comments / rich text

If you need these features, use a full-featured XLSX library. If you need to stream millions of rows with styling and minimal memory, use xlsxlite.

## CLI Tool

A bash script is included for quick inspection of any XLSX file:

```bash
./scripts/xlsxread.sh path/to/file.xlsx
```

Example output:

```
  File:              data.xlsx
  Open time:         4.7ms
  Sheets:            1

═══ Sheet 0: "ExcelJS sheet" ═══

         [str]      [str]          [str]                                     [num]             [num]
  ─────────────────────────────────────────────────────────────────────────────────────────────────────────
  1      branch_no  branch_name    detail_address                            longitude         latitude
  2      TH_BTH01   TH NONGSA      Ruko kopkarlak no 6, Batam Kota           1.10205262111752  104.075602236687
  3      TH_BTH02   TH BATU AJI    Ruko Cemara Asri Blok BB8 No. 10 Tembesi  1.04129828830344  103.989648220723
  4      TH_BTH03   TH LUBUK BAJA  Jl. Pembangunan Citra Mas Blok A No 8, …  1.13631267303236  104.007846409705
  5      TH_BTH04   TH SEKUPANG    RUKO KHARISMA BUSINESS CENTER BLOK C NO…  1.11091211704821  103.949332371163

  Rows:              537
  Max columns:       5
  Non-empty cells:   2685
  Read time:         5.2ms

  Total elapsed:     10.2ms
```

It displays a table with column type headers (`[str]`, `[num]`, `[bool]`, `[date]`), the first 5 sample rows (up to 10 columns, values truncated at 40 chars), followed by summary stats and timing.

## Project structure

```
xlsxlite/
├── types.go              # Core types: Cell, Row, Style, Font, Fill, Border, etc.
├── coords.go             # Coordinate helpers: CellRef, ParseCellRef, ColIndexToLetter
├── helpers.go            # File helpers and cell constructors: OpenFile, MakeRow, etc.
├── reader.go             # Streaming XLSX reader with row iterator
├── writer.go             # Streaming XLSX writer with shared string dedup
├── styles.go             # Style management with deduplication
├── gcs/
│   └── gcs.go            # Google Cloud Storage read/write helpers
├── scripts/
│   └── xlsxread.sh       # Bash script for inspecting XLSX files
├── internal/
│   └── xlsxread/main.go  # Reader tool implementation
├── .github/
│   └── workflows/
│       └── ci.yml        # Lint, test, benchmark, and auto-release on merge
├── xlsxlite_test.go      # Unit tests (55 tests)
└── bench_test.go         # Benchmarks
```

## License

BSD-3-Clause
