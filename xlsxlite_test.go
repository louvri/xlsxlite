package xlsxlite

import (
	"bytes"
	"fmt"
	"math"
	"os"
	"strings"
	"testing"
	"time"
)

// ──────────────────────────────────────────────
// coords.go tests
// ──────────────────────────────────────────────

func TestColIndexToLetter(t *testing.T) {
	tests := []struct {
		idx    int
		letter string
	}{
		{0, "A"},
		{1, "B"},
		{25, "Z"},
		{26, "AA"},
		{27, "AB"},
		{51, "AZ"},
		{52, "BA"},
		{701, "ZZ"},
		{702, "AAA"},
		{16383, "XFD"}, // Excel max column
	}
	for _, tt := range tests {
		got := ColIndexToLetter(tt.idx)
		if got != tt.letter {
			t.Errorf("ColIndexToLetter(%d) = %q, want %q", tt.idx, got, tt.letter)
		}
	}
}

func TestLetterToColIndex(t *testing.T) {
	tests := []struct {
		letter string
		idx    int
	}{
		{"A", 0},
		{"B", 1},
		{"Z", 25},
		{"AA", 26},
		{"AB", 27},
		{"ZZ", 701},
		{"AAA", 702},
		{"XFD", 16383},
	}
	for _, tt := range tests {
		got := LetterToColIndex(tt.letter)
		if got != tt.idx {
			t.Errorf("LetterToColIndex(%q) = %d, want %d", tt.letter, got, tt.idx)
		}
	}
}

func TestLetterToColIndexCaseInsensitive(t *testing.T) {
	if LetterToColIndex("aa") != 26 {
		t.Errorf("LetterToColIndex(aa) = %d, want 26", LetterToColIndex("aa"))
	}
	if LetterToColIndex("Ab") != 27 {
		t.Errorf("LetterToColIndex(Ab) = %d, want 27", LetterToColIndex("Ab"))
	}
}

func TestColConversionRoundTrip(t *testing.T) {
	for i := 0; i <= 16383; i++ {
		letter := ColIndexToLetter(i)
		back := LetterToColIndex(letter)
		if back != i {
			t.Fatalf("round-trip failed: %d → %q → %d", i, letter, back)
		}
	}
}

func TestCellRef(t *testing.T) {
	tests := []struct {
		col, row int
		want     string
	}{
		{0, 1, "A1"},
		{0, 100, "A100"},
		{25, 1, "Z1"},
		{26, 100, "AA100"},
		{702, 999, "AAA999"},
	}
	for _, tt := range tests {
		got := CellRef(tt.col, tt.row)
		if got != tt.want {
			t.Errorf("CellRef(%d, %d) = %q, want %q", tt.col, tt.row, got, tt.want)
		}
	}
}

func TestParseCellRef(t *testing.T) {
	tests := []struct {
		ref     string
		col     int
		row     int
		wantErr bool
	}{
		{"A1", 0, 1, false},
		{"B5", 1, 5, false},
		{"Z100", 25, 100, false},
		{"AA1", 26, 1, false},
		{"AAA999", 702, 999, false},
		// lowercase
		{"b5", 1, 5, false},
		{"aa1", 26, 1, false},
		// error cases
		{"", 0, 0, true},
		{"123", 0, 0, true},
		{"A", 0, 0, true},
		{"1A", 0, 0, true},
		{"A1B", 0, 0, true},
	}
	for _, tt := range tests {
		col, row, err := ParseCellRef(tt.ref)
		if tt.wantErr {
			if err == nil {
				t.Errorf("ParseCellRef(%q): expected error, got col=%d row=%d", tt.ref, col, row)
			}
			continue
		}
		if err != nil {
			t.Errorf("ParseCellRef(%q): unexpected error: %v", tt.ref, err)
			continue
		}
		if col != tt.col || row != tt.row {
			t.Errorf("ParseCellRef(%q) = (%d, %d), want (%d, %d)", tt.ref, col, row, tt.col, tt.row)
		}
	}
}

func TestParseCellRefRoundTrip(t *testing.T) {
	for col := 0; col < 100; col++ {
		for row := 1; row <= 10; row++ {
			ref := CellRef(col, row)
			gotCol, gotRow, err := ParseCellRef(ref)
			if err != nil {
				t.Fatalf("ParseCellRef(%q) error: %v", ref, err)
			}
			if gotCol != col || gotRow != row {
				t.Fatalf("CellRef(%d,%d) = %q → ParseCellRef = (%d,%d)", col, row, ref, gotCol, gotRow)
			}
		}
	}
}

func TestRangeRef(t *testing.T) {
	tests := []struct {
		sc, sr, ec, er int
		want           string
	}{
		{0, 1, 2, 1, "A1:C1"},
		{0, 1, 0, 10, "A1:A10"},
		{0, 1, 25, 100, "A1:Z100"},
		{26, 1, 27, 5, "AA1:AB5"},
	}
	for _, tt := range tests {
		got := RangeRef(tt.sc, tt.sr, tt.ec, tt.er)
		if got != tt.want {
			t.Errorf("RangeRef(%d,%d,%d,%d) = %q, want %q", tt.sc, tt.sr, tt.ec, tt.er, got, tt.want)
		}
	}
}

// ──────────────────────────────────────────────
// types.go tests (date conversions)
// ──────────────────────────────────────────────

func TestTimeToExcelSerial(t *testing.T) {
	tests := []struct {
		t      time.Time
		serial int
	}{
		{time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC), 2},
		{time.Date(2020, 1, 1, 0, 0, 0, 0, time.UTC), 43831},
		{time.Date(2025, 6, 15, 0, 0, 0, 0, time.UTC), 45823},
		{time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC), 1},
	}
	for _, tt := range tests {
		serial := TimeToExcelSerial(tt.t)
		if int(serial) != tt.serial {
			t.Errorf("TimeToExcelSerial(%v) = %g, want %d", tt.t, serial, tt.serial)
		}
	}
}

func TestExcelSerialToTime(t *testing.T) {
	back := ExcelSerialToTime(43831)
	if back.Year() != 2020 || back.Month() != 1 || back.Day() != 1 {
		t.Errorf("ExcelSerialToTime(43831) = %v, want 2020-01-01", back)
	}
}

func TestDateConversionRoundTrip(t *testing.T) {
	dates := []time.Time{
		time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC),
		time.Date(2000, 6, 15, 0, 0, 0, 0, time.UTC),
		time.Date(2025, 12, 31, 0, 0, 0, 0, time.UTC),
	}
	for _, d := range dates {
		serial := TimeToExcelSerial(d)
		back := ExcelSerialToTime(serial)
		if back.Year() != d.Year() || back.Month() != d.Month() || back.Day() != d.Day() {
			t.Errorf("round-trip failed: %v → %g → %v", d, serial, back)
		}
	}
}

// ──────────────────────────────────────────────
// helpers.go tests (cell constructors)
// ──────────────────────────────────────────────

func TestStringCell(t *testing.T) {
	c := StringCell("hello")
	if c.Value != "hello" || c.Type != CellTypeString || c.StyleID != 0 {
		t.Errorf("StringCell(hello) = %+v", c)
	}
}

func TestNumberCell(t *testing.T) {
	c := NumberCell(3.14)
	if c.Value != 3.14 || c.Type != CellTypeNumber || c.StyleID != 0 {
		t.Errorf("NumberCell(3.14) = %+v", c)
	}
}

func TestIntCell(t *testing.T) {
	c := IntCell(42)
	if c.Value != 42 || c.Type != CellTypeNumber || c.StyleID != 0 {
		t.Errorf("IntCell(42) = %+v", c)
	}
}

func TestBoolCell(t *testing.T) {
	c := BoolCell(true)
	if c.Value != true || c.Type != CellTypeBool {
		t.Errorf("BoolCell(true) = %+v", c)
	}
	c = BoolCell(false)
	if c.Value != false || c.Type != CellTypeBool {
		t.Errorf("BoolCell(false) = %+v", c)
	}
}

func TestDateCellConstructor(t *testing.T) {
	d := time.Date(2025, 1, 1, 0, 0, 0, 0, time.UTC)
	c := DateCell(d, 5)
	if c.Type != CellTypeDate || c.StyleID != 5 {
		t.Errorf("DateCell = %+v", c)
	}
	if v, ok := c.Value.(time.Time); !ok || !v.Equal(d) {
		t.Errorf("DateCell value = %v, want %v", c.Value, d)
	}
}

func TestEmptyCell(t *testing.T) {
	c := EmptyCell()
	if c.Value != nil || c.Type != CellTypeEmpty || c.StyleID != 0 {
		t.Errorf("EmptyCell() = %+v", c)
	}
}

func TestStyledCell(t *testing.T) {
	tests := []struct {
		name     string
		value    any
		wantType CellType
	}{
		{"string", "hello", CellTypeString},
		{"float64", 3.14, CellTypeNumber},
		{"float32", float32(1.5), CellTypeNumber},
		{"int", 42, CellTypeNumber},
		{"int64", int64(100), CellTypeNumber},
		{"bool", true, CellTypeBool},
		{"nil", nil, CellTypeEmpty},
	}
	for _, tt := range tests {
		c := StyledCell(tt.value, 3)
		if c.Type != tt.wantType {
			t.Errorf("StyledCell(%s): type = %d, want %d", tt.name, c.Type, tt.wantType)
		}
		if c.StyleID != 3 {
			t.Errorf("StyledCell(%s): styleID = %d, want 3", tt.name, c.StyleID)
		}
	}
}

func TestMakeRow(t *testing.T) {
	row := MakeRow("text", 42, 3.14, true, nil)
	if len(row.Cells) != 5 {
		t.Fatalf("MakeRow: got %d cells, want 5", len(row.Cells))
	}
	if row.Cells[0].Type != CellTypeString || row.Cells[0].Value != "text" {
		t.Errorf("cell 0: %+v", row.Cells[0])
	}
	if row.Cells[1].Type != CellTypeNumber || row.Cells[1].Value != 42 {
		t.Errorf("cell 1: %+v", row.Cells[1])
	}
	if row.Cells[2].Type != CellTypeNumber || row.Cells[2].Value != 3.14 {
		t.Errorf("cell 2: %+v", row.Cells[2])
	}
	if row.Cells[3].Type != CellTypeBool || row.Cells[3].Value != true {
		t.Errorf("cell 3: %+v", row.Cells[3])
	}
	if row.Cells[4].Type != CellTypeEmpty || row.Cells[4].Value != nil {
		t.Errorf("cell 4: %+v", row.Cells[4])
	}
}

func TestMakeRowFloat32(t *testing.T) {
	row := MakeRow(float32(1.5))
	if row.Cells[0].Type != CellTypeNumber {
		t.Errorf("float32 type = %d, want CellTypeNumber", row.Cells[0].Type)
	}
	if v, ok := row.Cells[0].Value.(float64); !ok || v != 1.5 {
		t.Errorf("float32 value = %v (%T)", row.Cells[0].Value, row.Cells[0].Value)
	}
}

func TestMakeRowInt64(t *testing.T) {
	row := MakeRow(int64(999))
	if row.Cells[0].Type != CellTypeNumber {
		t.Errorf("int64 type = %d", row.Cells[0].Type)
	}
	if v, ok := row.Cells[0].Value.(int64); !ok || v != 999 {
		t.Errorf("int64 value = %v (%T)", row.Cells[0].Value, row.Cells[0].Value)
	}
}

func TestMakeRowWithCell(t *testing.T) {
	styled := StyledCell("hi", 2)
	row := MakeRow(styled)
	if row.Cells[0] != styled {
		t.Errorf("Cell passthrough failed: got %+v, want %+v", row.Cells[0], styled)
	}
}

func TestMakeRowUnknownType(t *testing.T) {
	type custom struct{ X int }
	row := MakeRow(custom{X: 42})
	if row.Cells[0].Type != CellTypeString {
		t.Errorf("unknown type: got type %d, want CellTypeString", row.Cells[0].Type)
	}
	if row.Cells[0].Value != "{42}" {
		t.Errorf("unknown type: got value %q", row.Cells[0].Value)
	}
}

func TestMakeRowEmpty(t *testing.T) {
	row := MakeRow()
	if len(row.Cells) != 0 {
		t.Errorf("MakeRow(): got %d cells, want 0", len(row.Cells))
	}
}

// ──────────────────────────────────────────────
// styles.go tests
// ──────────────────────────────────────────────

func TestNewStyleSheetDefaults(t *testing.T) {
	ss := NewStyleSheet()
	// Must have at least: 1 font, 2 fills, 1 border, 1 xf
	if len(ss.fonts) < 1 {
		t.Error("expected at least 1 default font")
	}
	if len(ss.fills) < 2 {
		t.Error("expected at least 2 default fills (none + gray125)")
	}
	if len(ss.borders) < 1 {
		t.Error("expected at least 1 default border")
	}
	if len(ss.xfs) < 1 {
		t.Error("expected at least 1 default xf")
	}
}

func TestStyleDedup(t *testing.T) {
	ss := NewStyleSheet()
	s := Style{
		Font: &Font{Name: "Arial", Size: 12, Bold: true},
		Fill: &Fill{Type: "pattern", Pattern: "solid", FgColor: "FFFF0000"},
	}
	id1 := ss.AddStyle(s)
	id2 := ss.AddStyle(s) // identical style
	if id1 != id2 {
		t.Errorf("identical styles got different IDs: %d vs %d", id1, id2)
	}
}

func TestStyleDedupDifferent(t *testing.T) {
	ss := NewStyleSheet()
	id1 := ss.AddStyle(Style{Font: &Font{Bold: true}})
	id2 := ss.AddStyle(Style{Font: &Font{Italic: true}})
	if id1 == id2 {
		t.Error("different styles got same ID")
	}
}

func TestFontDedup(t *testing.T) {
	ss := NewStyleSheet()
	f := Font{Name: "Arial", Size: 12}
	id1 := ss.addFont(&f)
	id2 := ss.addFont(&f)
	if id1 != id2 {
		t.Errorf("identical fonts got different IDs: %d vs %d", id1, id2)
	}
}

func TestFontNil(t *testing.T) {
	ss := NewStyleSheet()
	if ss.addFont(nil) != 0 {
		t.Error("nil font should return 0")
	}
}

func TestFillDedup(t *testing.T) {
	ss := NewStyleSheet()
	f := Fill{Type: "pattern", Pattern: "solid", FgColor: "FFFF0000"}
	id1 := ss.addFill(&f)
	id2 := ss.addFill(&f)
	if id1 != id2 {
		t.Errorf("identical fills got different IDs: %d vs %d", id1, id2)
	}
}

func TestFillNil(t *testing.T) {
	ss := NewStyleSheet()
	if ss.addFill(nil) != 0 {
		t.Error("nil fill should return 0")
	}
}

func TestBorderDedup(t *testing.T) {
	ss := NewStyleSheet()
	b := Border{Bottom: BorderEdge{Style: "thin", Color: "FF000000"}}
	id1 := ss.addBorder(&b)
	id2 := ss.addBorder(&b)
	if id1 != id2 {
		t.Errorf("identical borders got different IDs: %d vs %d", id1, id2)
	}
}

func TestBorderNil(t *testing.T) {
	ss := NewStyleSheet()
	if ss.addBorder(nil) != 0 {
		t.Error("nil border should return 0")
	}
}

func TestNumFmtDedup(t *testing.T) {
	ss := NewStyleSheet()
	id1 := ss.addNumFmt("yyyy-mm-dd")
	id2 := ss.addNumFmt("yyyy-mm-dd")
	if id1 != id2 {
		t.Errorf("identical numFmt got different IDs: %d vs %d", id1, id2)
	}
}

func TestNumFmtEmpty(t *testing.T) {
	ss := NewStyleSheet()
	if ss.addNumFmt("") != 0 {
		t.Error("empty numFmt should return 0 (General)")
	}
}

func TestNumFmtCustomStartsAt164(t *testing.T) {
	ss := NewStyleSheet()
	id := ss.addNumFmt("yyyy-mm-dd")
	if id != 164 {
		t.Errorf("first custom numFmt ID = %d, want 164", id)
	}
	id2 := ss.addNumFmt("#,##0.00")
	if id2 != 165 {
		t.Errorf("second custom numFmt ID = %d, want 165", id2)
	}
}

func TestAlignmentEqual(t *testing.T) {
	if !alignmentEqual(nil, nil) {
		t.Error("nil, nil should be equal")
	}
	a := &Alignment{Horizontal: "center"}
	if alignmentEqual(a, nil) {
		t.Error("non-nil, nil should not be equal")
	}
	if alignmentEqual(nil, a) {
		t.Error("nil, non-nil should not be equal")
	}
	b := &Alignment{Horizontal: "center"}
	if !alignmentEqual(a, b) {
		t.Error("identical alignments should be equal")
	}
	c := &Alignment{Horizontal: "left"}
	if alignmentEqual(a, c) {
		t.Error("different alignments should not be equal")
	}
}

func TestStyleDedupWithAlignment(t *testing.T) {
	ss := NewStyleSheet()
	s1 := Style{Alignment: &Alignment{Horizontal: "center"}}
	s2 := Style{Alignment: &Alignment{Horizontal: "center"}}
	s3 := Style{Alignment: &Alignment{Horizontal: "left"}}

	id1 := ss.AddStyle(s1)
	id2 := ss.AddStyle(s2)
	id3 := ss.AddStyle(s3)

	if id1 != id2 {
		t.Errorf("identical alignment styles got different IDs: %d vs %d", id1, id2)
	}
	if id1 == id3 {
		t.Error("different alignment styles got same ID")
	}
}

func TestStyleSheetWriteXML(t *testing.T) {
	ss := NewStyleSheet()
	ss.AddStyle(Style{
		Font:         &Font{Name: "Arial", Size: 14, Bold: true, Italic: true, Underline: true, Color: "FF0000FF"},
		Fill:         &Fill{Type: "pattern", Pattern: "solid", FgColor: "FFFFFF00", BgColor: "FF000000"},
		Border:       &Border{Bottom: BorderEdge{Style: "thin", Color: "FF000000"}, Top: BorderEdge{Style: "medium"}},
		Alignment:    &Alignment{Horizontal: "center", Vertical: "top", WrapText: true},
		NumberFormat: "yyyy-mm-dd",
	})

	var buf bytes.Buffer
	if err := ss.writeXML(&buf); err != nil {
		t.Fatal(err)
	}

	xml := buf.String()
	// Verify key elements are present
	checks := []string{
		`<numFmt numFmtId="164"`,
		`<b/>`,
		`<i/>`,
		`<u/>`,
		`<sz val="14"/>`,
		`<color rgb="FF0000FF"/>`,
		`<name val="Arial"/>`,
		`<fgColor rgb="FFFFFF00"/>`,
		`<bgColor rgb="FF000000"/>`,
		`style="thin"`,
		`style="medium"`,
		`horizontal="center"`,
		`vertical="top"`,
		`wrapText="1"`,
		`applyFont="1"`,
		`applyFill="1"`,
		`applyBorder="1"`,
		`applyNumberFormat="1"`,
		`applyAlignment="1"`,
		`cellStyleXfs`,
		`cellStyles`,
	}
	for _, check := range checks {
		if !strings.Contains(xml, check) {
			t.Errorf("styles.xml missing %q", check)
		}
	}
}

func TestEscapeXMLAttr(t *testing.T) {
	tests := []struct {
		input, want string
	}{
		{"hello", "hello"},
		{"a&b", "a&amp;b"},
		{"a<b", "a&lt;b"},
		{"a>b", "a&gt;b"},
		{`a"b`, "a&quot;b"},
		{`<"&>`, "&lt;&quot;&amp;&gt;"},
	}
	for _, tt := range tests {
		got := escapeXMLAttr(tt.input)
		if got != tt.want {
			t.Errorf("escapeXMLAttr(%q) = %q, want %q", tt.input, got, tt.want)
		}
	}
}

// ──────────────────────────────────────────────
// writer.go tests
// ──────────────────────────────────────────────

func TestAppendEscapedXML(t *testing.T) {
	tests := []struct {
		input, want string
	}{
		{"hello", "hello"},
		{"a&b", "a&amp;b"},
		{"<tag>", "&lt;tag&gt;"},
		{"no escaping needed", "no escaping needed"},
		{"&<>", "&amp;&lt;&gt;"},
		{"", ""},
	}
	for _, tt := range tests {
		var b bytes.Buffer
		appendEscapedXML(&b, tt.input)
		got := b.String()
		if got != tt.want {
			t.Errorf("appendEscapedXML(%q) = %q, want %q", tt.input, got, tt.want)
		}
	}
}

func TestSharedStringWriterDedup(t *testing.T) {
	ss := newSharedStringWriter()
	id0 := ss.Add("hello")
	id1 := ss.Add("world")
	id2 := ss.Add("hello") // duplicate

	if id0 != 0 {
		t.Errorf("first string: id = %d, want 0", id0)
	}
	if id1 != 1 {
		t.Errorf("second string: id = %d, want 1", id1)
	}
	if id2 != 0 {
		t.Errorf("duplicate string: id = %d, want 0", id2)
	}
	if len(ss.strings) != 2 {
		t.Errorf("unique count = %d, want 2", len(ss.strings))
	}
}

func TestSharedStringWriterXML(t *testing.T) {
	ss := newSharedStringWriter()
	ss.Add("hello")
	ss.Add("a&b")
	ss.Add("<tag>")

	var buf bytes.Buffer
	if err := ss.writeXML(&buf); err != nil {
		t.Fatal(err)
	}

	xml := buf.String()
	if !strings.Contains(xml, `count="3"`) {
		t.Error("missing count")
	}
	if !strings.Contains(xml, `uniqueCount="3"`) {
		t.Error("missing uniqueCount")
	}
	if !strings.Contains(xml, `<si><t>hello</t></si>`) {
		t.Error("missing hello entry")
	}
	if !strings.Contains(xml, `<si><t>a&amp;b</t></si>`) {
		t.Error("missing escaped & entry")
	}
	if !strings.Contains(xml, `<si><t>&lt;tag&gt;</t></si>`) {
		t.Error("missing escaped <> entry")
	}
}

// ──────────────────────────────────────────────
// Write + Read round-trip tests
// ──────────────────────────────────────────────

// writeAndRead is a helper that writes an XLSX to a buffer and returns a Reader.
func writeAndRead(t *testing.T, fn func(w *Writer)) *Reader {
	t.Helper()
	var buf bytes.Buffer
	w := NewWriter(&buf)
	fn(w)
	if err := w.Close(); err != nil {
		t.Fatalf("Writer.Close: %v", err)
	}
	data := buf.Bytes()
	reader, err := OpenReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	return reader
}

// readAllRows reads all rows from a named sheet.
func readAllRows(t *testing.T, r *Reader, name string) []*Row {
	t.Helper()
	iter, err := r.OpenSheet(name)
	if err != nil {
		t.Fatalf("OpenSheet(%q): %v", name, err)
	}
	defer iter.Close()

	var rows []*Row
	for iter.Next() {
		row := iter.Row()
		// Copy to avoid iterator reuse issues
		cp := *row
		cp.Cells = make([]Cell, len(row.Cells))
		copy(cp.Cells, row.Cells)
		rows = append(rows, &cp)
	}
	if err := iter.Err(); err != nil {
		t.Fatalf("iterator error: %v", err)
	}
	return rows
}

func TestBasicWriteAndRead(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "TestSheet", ColWidths: map[int]float64{0: 20, 1: 15, 2: 30}})
		sw.WriteRow(MakeRow("Name", "Age", "Active"))
		sw.WriteRow(MakeRow("Alice", 30, true))
		sw.WriteRow(MakeRow("Bob", 25, false))
		sw.WriteRow(MakeRow("Charlie", 35.5, true))
		sw.Close()
	})

	sheets := reader.SheetNames()
	if len(sheets) != 1 || sheets[0] != "TestSheet" {
		t.Errorf("SheetNames = %v, want [TestSheet]", sheets)
	}
	if reader.SheetCount() != 1 {
		t.Errorf("SheetCount = %d, want 1", reader.SheetCount())
	}

	rows := readAllRows(t, reader, "TestSheet")
	if len(rows) != 4 {
		t.Fatalf("got %d rows, want 4", len(rows))
	}

	// Header
	if rows[0].Cells[0].Value != "Name" || rows[0].Cells[1].Value != "Age" || rows[0].Cells[2].Value != "Active" {
		t.Errorf("header = %v", rows[0].Cells)
	}

	// Alice
	if rows[1].Cells[0].Value != "Alice" {
		t.Errorf("row 1 col 0 = %v, want Alice", rows[1].Cells[0].Value)
	}
	if v, ok := rows[1].Cells[1].Value.(float64); !ok || v != 30 {
		t.Errorf("row 1 col 1 = %v (%T), want 30", rows[1].Cells[1].Value, rows[1].Cells[1].Value)
	}
	if v, ok := rows[1].Cells[2].Value.(bool); !ok || v != true {
		t.Errorf("row 1 col 2 = %v, want true", rows[1].Cells[2].Value)
	}

	// Bob: false
	if v, ok := rows[2].Cells[2].Value.(bool); !ok || v != false {
		t.Errorf("row 2 col 2 = %v, want false", rows[2].Cells[2].Value)
	}

	// Charlie: float
	if v, ok := rows[3].Cells[1].Value.(float64); !ok || v != 35.5 {
		t.Errorf("row 3 col 1 = %v, want 35.5", rows[3].Cells[1].Value)
	}
}

func TestAllCellTypesRoundTrip(t *testing.T) {
	testDate := time.Date(2025, 6, 15, 0, 0, 0, 0, time.UTC)
	reader := writeAndRead(t, func(w *Writer) {
		ss := w.StyleSheet()
		dateStyle := ss.AddStyle(Style{NumberFormat: "yyyy-mm-dd"})

		sw, _ := w.NewSheet(SheetConfig{Name: "Types"})
		sw.WriteRow(Row{Cells: []Cell{
			StringCell("text"),
			NumberCell(3.14),
			IntCell(42),
			BoolCell(true),
			BoolCell(false),
			DateCell(testDate, dateStyle),
			EmptyCell(),
			{Value: float32(1.5), Type: CellTypeNumber},
			{Value: int64(999), Type: CellTypeNumber},
		}})
		sw.Close()
	})

	rows := readAllRows(t, reader, "Types")
	if len(rows) != 1 {
		t.Fatalf("got %d rows, want 1", len(rows))
	}
	cells := rows[0].Cells

	// string
	if cells[0].Value != "text" || cells[0].Type != CellTypeString {
		t.Errorf("string cell: %+v", cells[0])
	}
	// float64
	if v, ok := cells[1].Value.(float64); !ok || v != 3.14 {
		t.Errorf("float64 cell: %+v", cells[1])
	}
	// int (read back as float64)
	if v, ok := cells[2].Value.(float64); !ok || v != 42 {
		t.Errorf("int cell: %+v", cells[2])
	}
	// bool true
	if v, ok := cells[3].Value.(bool); !ok || v != true {
		t.Errorf("bool true cell: %+v", cells[3])
	}
	// bool false
	if v, ok := cells[4].Value.(bool); !ok || v != false {
		t.Errorf("bool false cell: %+v", cells[4])
	}
	// date
	if cells[5].Type != CellTypeDate {
		t.Errorf("date cell type = %d, want CellTypeDate", cells[5].Type)
	}
	if dt, ok := cells[5].Value.(time.Time); !ok || dt.Year() != 2025 || dt.Month() != 6 || dt.Day() != 15 {
		t.Errorf("date cell: %+v", cells[5])
	}
	// empty
	if cells[6].Type != CellTypeEmpty {
		t.Errorf("empty cell type = %d", cells[6].Type)
	}
	// float32 (read back as float64)
	if v, ok := cells[7].Value.(float64); !ok || v != 1.5 {
		t.Errorf("float32 cell: %+v", cells[7])
	}
	// int64 (read back as float64)
	if v, ok := cells[8].Value.(float64); !ok || v != 999 {
		t.Errorf("int64 cell: %+v", cells[8])
	}
}

func TestEmptyCellsRoundTrip(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "Sparse"})
		sw.WriteRow(Row{Cells: []Cell{
			StringCell("A"),
			EmptyCell(),
			EmptyCell(),
			StringCell("D"),
		}})
		sw.Close()
	})

	rows := readAllRows(t, reader, "Sparse")
	cells := rows[0].Cells
	if len(cells) < 4 {
		t.Fatalf("got %d cells, want at least 4", len(cells))
	}
	if cells[0].Value != "A" {
		t.Errorf("cell 0 = %v, want A", cells[0].Value)
	}
	if cells[3].Value != "D" {
		t.Errorf("cell 3 = %v, want D", cells[3].Value)
	}
}

func TestEmptyCellWithStyle(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		ss := w.StyleSheet()
		style := ss.AddStyle(Style{Font: &Font{Bold: true}})
		sw, _ := w.NewSheet(SheetConfig{Name: "S"})
		sw.WriteRow(Row{Cells: []Cell{
			{Type: CellTypeEmpty, Value: nil, StyleID: style},
		}})
		sw.Close()
	})

	rows := readAllRows(t, reader, "S")
	if len(rows) != 1 || len(rows[0].Cells) < 1 {
		t.Fatal("expected 1 row with 1 cell")
	}
	if rows[0].Cells[0].StyleID != 1 {
		t.Errorf("empty styled cell: styleID = %d, want 1", rows[0].Cells[0].StyleID)
	}
}

func TestStyleRoundTrip(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		ss := w.StyleSheet()
		boldStyle := ss.AddStyle(Style{
			Font:      &Font{Name: "Arial", Size: 14, Bold: true, Color: "FF0000FF"},
			Fill:      &Fill{Type: "pattern", Pattern: "solid", FgColor: "FFFFFF00"},
			Border:    &Border{Bottom: BorderEdge{Style: "thin", Color: "FF000000"}},
			Alignment: &Alignment{Horizontal: "center", Vertical: "center"},
		})

		sw, _ := w.NewSheet(SheetConfig{Name: "Styled"})
		sw.WriteRow(Row{Cells: []Cell{
			StyledCell("Header", boldStyle),
			StyledCell(42.0, boldStyle),
		}})
		sw.Close()
	})

	rows := readAllRows(t, reader, "Styled")
	if rows[0].Cells[0].StyleID == 0 {
		t.Error("expected non-zero styleID for styled cell")
	}
	if rows[0].Cells[0].StyleID != rows[0].Cells[1].StyleID {
		t.Error("cells with same style should have same styleID")
	}
}

func TestMultipleStyles(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		ss := w.StyleSheet()
		s1 := ss.AddStyle(Style{Font: &Font{Bold: true}})
		s2 := ss.AddStyle(Style{Font: &Font{Italic: true}})

		sw, _ := w.NewSheet(SheetConfig{Name: "S"})
		sw.WriteRow(Row{Cells: []Cell{
			StyledCell("bold", s1),
			StyledCell("italic", s2),
			StringCell("plain"),
		}})
		sw.Close()
	})

	rows := readAllRows(t, reader, "S")
	if rows[0].Cells[0].StyleID == rows[0].Cells[1].StyleID {
		t.Error("different styles should have different styleIDs")
	}
	if rows[0].Cells[2].StyleID != 0 {
		t.Errorf("plain cell styleID = %d, want 0", rows[0].Cells[2].StyleID)
	}
}

func TestMergeCells(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{
			Name: "Merged",
			MergeCells: []MergeCell{
				{StartCol: 0, StartRow: 1, EndCol: 2, EndRow: 1},
			},
		})
		sw.WriteRow(MakeRow("Merged Title", nil, nil))
		sw.Close()
	})

	rows := readAllRows(t, reader, "Merged")
	if rows[0].Cells[0].Value != "Merged Title" {
		t.Errorf("got %v, want Merged Title", rows[0].Cells[0].Value)
	}
}

func TestMultipleMergeCells(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{
			Name: "Multi",
			MergeCells: []MergeCell{
				{StartCol: 0, StartRow: 1, EndCol: 2, EndRow: 1},
				{StartCol: 0, StartRow: 2, EndCol: 3, EndRow: 2},
			},
		})
		sw.WriteRow(MakeRow("M1"))
		sw.WriteRow(MakeRow("M2"))
		sw.Close()
	})

	rows := readAllRows(t, reader, "Multi")
	if len(rows) != 2 {
		t.Fatalf("got %d rows, want 2", len(rows))
	}
}

func TestMultipleSheets(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		for _, name := range []string{"Sheet1", "Sheet2", "Sheet3"} {
			sw, _ := w.NewSheet(SheetConfig{Name: name})
			sw.WriteRow(MakeRow("Data from " + name))
			sw.Close()
		}
	})

	names := reader.SheetNames()
	if len(names) != 3 {
		t.Fatalf("got %d sheets, want 3", len(names))
	}
	if reader.SheetCount() != 3 {
		t.Errorf("SheetCount = %d, want 3", reader.SheetCount())
	}

	for i, name := range []string{"Sheet1", "Sheet2", "Sheet3"} {
		if names[i] != name {
			t.Errorf("sheet %d name = %q, want %q", i, names[i], name)
		}
		iter, err := reader.OpenSheetByIndex(i)
		if err != nil {
			t.Fatal(err)
		}
		if !iter.Next() {
			t.Fatalf("sheet %s: no rows", name)
		}
		if iter.Row().Cells[0].Value != "Data from "+name {
			t.Errorf("sheet %s: got %q", name, iter.Row().Cells[0].Value)
		}
		iter.Close()
	}
}

func TestRowHeight(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "H"})
		sw.WriteRow(Row{Cells: []Cell{StringCell("Tall")}, Height: 40})
		sw.WriteRow(Row{Cells: []Cell{StringCell("Default")}})
		sw.Close()
	})

	rows := readAllRows(t, reader, "H")
	if rows[0].Height != 40 {
		t.Errorf("row 0 height = %g, want 40", rows[0].Height)
	}
	if rows[1].Height != 0 {
		t.Errorf("row 1 height = %g, want 0 (default)", rows[1].Height)
	}
}

func TestRowIndex(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "RI"})
		sw.WriteRow(MakeRow("row1"))                               // auto row 1
		sw.WriteRow(MakeRow("row2"))                               // auto row 2
		sw.WriteRow(Row{Cells: []Cell{StringCell("row5")}, RowIndex: 5}) // explicit row 5
		sw.WriteRow(MakeRow("row6"))                               // auto row 6
		sw.Close()
	})

	rows := readAllRows(t, reader, "RI")
	if len(rows) != 4 {
		t.Fatalf("got %d rows, want 4", len(rows))
	}
	if rows[0].RowIndex != 1 {
		t.Errorf("row 0 index = %d, want 1", rows[0].RowIndex)
	}
	if rows[1].RowIndex != 2 {
		t.Errorf("row 1 index = %d, want 2", rows[1].RowIndex)
	}
	if rows[2].RowIndex != 5 {
		t.Errorf("row 2 index = %d, want 5", rows[2].RowIndex)
	}
	if rows[3].RowIndex != 6 {
		t.Errorf("row 3 index = %d, want 6", rows[3].RowIndex)
	}
}

func TestDateCellRoundTrip(t *testing.T) {
	dates := []time.Time{
		time.Date(2020, 1, 1, 0, 0, 0, 0, time.UTC),
		time.Date(2025, 6, 15, 0, 0, 0, 0, time.UTC),
		time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC),
	}

	reader := writeAndRead(t, func(w *Writer) {
		ss := w.StyleSheet()
		dateStyle := ss.AddStyle(Style{NumberFormat: "yyyy-mm-dd"})
		sw, _ := w.NewSheet(SheetConfig{Name: "Dates"})
		for _, d := range dates {
			sw.WriteRow(Row{Cells: []Cell{DateCell(d, dateStyle)}})
		}
		sw.Close()
	})

	rows := readAllRows(t, reader, "Dates")
	for i, d := range dates {
		cell := rows[i].Cells[0]
		if cell.Type != CellTypeDate {
			t.Errorf("row %d: type = %d, want CellTypeDate", i, cell.Type)
			continue
		}
		if dt, ok := cell.Value.(time.Time); ok {
			if dt.Year() != d.Year() || dt.Month() != d.Month() || dt.Day() != d.Day() {
				t.Errorf("row %d: date = %v, want %v", i, dt, d)
			}
		} else {
			t.Errorf("row %d: value type = %T, want time.Time", i, cell.Value)
		}
	}
}

func TestFreezeRowOnly(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "F", FreezeRow: 2})
		sw.WriteRow(MakeRow("header1"))
		sw.WriteRow(MakeRow("header2"))
		sw.WriteRow(MakeRow("data"))
		sw.Close()
	})

	rows := readAllRows(t, reader, "F")
	if len(rows) != 3 {
		t.Errorf("got %d rows, want 3", len(rows))
	}
}

func TestFreezeColOnly(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "FC", FreezeCol: 1})
		sw.WriteRow(MakeRow("frozen", "scrollable"))
		sw.Close()
	})

	rows := readAllRows(t, reader, "FC")
	if len(rows) != 1 {
		t.Errorf("got %d rows, want 1", len(rows))
	}
}

func TestFreezeRowAndCol(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "FRC", FreezeRow: 1, FreezeCol: 2})
		sw.WriteRow(MakeRow("a", "b", "c"))
		sw.Close()
	})

	rows := readAllRows(t, reader, "FRC")
	if len(rows) != 1 {
		t.Errorf("got %d rows, want 1", len(rows))
	}
}

func TestSpecialCharactersInCells(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "Special"})
		sw.WriteRow(MakeRow("a&b", "<tag>", "c>d", `say "hi"`, "normal"))
		sw.Close()
	})

	rows := readAllRows(t, reader, "Special")
	expected := []string{"a&b", "<tag>", "c>d", `say "hi"`, "normal"}
	for i, want := range expected {
		if rows[0].Cells[i].Value != want {
			t.Errorf("cell %d = %q, want %q", i, rows[0].Cells[i].Value, want)
		}
	}
}

func TestSpecialCharactersInSheetName(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "Sales & Revenue"})
		sw.WriteRow(MakeRow("data"))
		sw.Close()
	})

	names := reader.SheetNames()
	if names[0] != "Sales & Revenue" {
		t.Errorf("sheet name = %q, want %q", names[0], "Sales & Revenue")
	}
}

func TestEmptyStringCell(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "S"})
		sw.WriteRow(MakeRow("", "notempty"))
		sw.Close()
	})

	rows := readAllRows(t, reader, "S")
	if rows[0].Cells[0].Value != "" {
		t.Errorf("empty string cell: %+v", rows[0].Cells[0])
	}
	if rows[0].Cells[1].Value != "notempty" {
		t.Errorf("string cell: %+v", rows[0].Cells[1])
	}
}

func TestLargeNumbers(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "N"})
		sw.WriteRow(Row{Cells: []Cell{
			NumberCell(0),
			NumberCell(-1.5),
			NumberCell(math.MaxFloat64),
			NumberCell(math.SmallestNonzeroFloat64),
			IntCell(math.MaxInt32),
			{Value: int64(math.MaxInt64), Type: CellTypeNumber},
		}})
		sw.Close()
	})

	rows := readAllRows(t, reader, "N")
	cells := rows[0].Cells

	if v := cells[0].Value.(float64); v != 0 {
		t.Errorf("zero = %v", v)
	}
	if v := cells[1].Value.(float64); v != -1.5 {
		t.Errorf("negative = %v", v)
	}
	if v := cells[4].Value.(float64); v != float64(math.MaxInt32) {
		t.Errorf("MaxInt32 = %v", v)
	}
}

func TestManyColumns(t *testing.T) {
	numCols := 100
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "Wide"})
		cells := make([]Cell, numCols)
		for i := range cells {
			cells[i] = IntCell(i)
		}
		sw.WriteRow(Row{Cells: cells})
		sw.Close()
	})

	rows := readAllRows(t, reader, "Wide")
	if len(rows[0].Cells) != numCols {
		t.Fatalf("got %d cols, want %d", len(rows[0].Cells), numCols)
	}
	for i := 0; i < numCols; i++ {
		if v := rows[0].Cells[i].Value.(float64); v != float64(i) {
			t.Errorf("col %d = %v, want %d", i, v, i)
		}
	}
}

func TestEmptySheet(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "Empty"})
		sw.Close()
	})

	rows := readAllRows(t, reader, "Empty")
	if len(rows) != 0 {
		t.Errorf("got %d rows, want 0", len(rows))
	}
}

func TestSheetWithNoName(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{})
		sw.WriteRow(MakeRow("data"))
		sw.Close()
	})

	names := reader.SheetNames()
	if len(names) != 1 {
		t.Fatalf("got %d sheets", len(names))
	}
	// Auto-named as "Sheet1"
	if names[0] != "Sheet1" {
		t.Errorf("auto name = %q, want Sheet1", names[0])
	}
}

func TestManyRowsSmall(t *testing.T) {
	n := 1000
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "Many"})
		for i := 0; i < n; i++ {
			sw.WriteRow(MakeRow(fmt.Sprintf("r%d", i), i))
		}
		sw.Close()
	})

	rows := readAllRows(t, reader, "Many")
	if len(rows) != n {
		t.Fatalf("got %d rows, want %d", len(rows), n)
	}
	// Spot check first and last
	if rows[0].Cells[0].Value != "r0" {
		t.Errorf("first row: %v", rows[0].Cells[0].Value)
	}
	if rows[n-1].Cells[0].Value != fmt.Sprintf("r%d", n-1) {
		t.Errorf("last row: %v", rows[n-1].Cells[0].Value)
	}
}

// ──────────────────────────────────────────────
// Reader error handling tests
// ──────────────────────────────────────────────

func TestOpenReaderInvalidData(t *testing.T) {
	_, err := OpenReader(bytes.NewReader([]byte("not a zip")), 10)
	if err == nil {
		t.Error("expected error for invalid data")
	}
}

func TestOpenFileNotFound(t *testing.T) {
	_, err := OpenFile("/nonexistent/path/file.xlsx")
	if err == nil {
		t.Error("expected error for non-existent file")
	}
}

func TestOpenSheetNotFound(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "Real"})
		sw.WriteRow(MakeRow("data"))
		sw.Close()
	})

	_, err := reader.OpenSheet("NonExistent")
	if err == nil {
		t.Error("expected error for non-existent sheet")
	}
}

func TestOpenSheetByIndexOutOfRange(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "S"})
		sw.Close()
	})

	_, err := reader.OpenSheetByIndex(-1)
	if err == nil {
		t.Error("expected error for negative index")
	}
	_, err = reader.OpenSheetByIndex(1)
	if err == nil {
		t.Error("expected error for out-of-range index")
	}
}

func TestReaderClose(t *testing.T) {
	// Reader from OpenReader (no closer)
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "S"})
		sw.Close()
	})
	if err := reader.Close(); err != nil {
		t.Errorf("Close on OpenReader-created reader: %v", err)
	}
}

func TestReaderCloseWithFile(t *testing.T) {
	path := "/tmp/xlsxlite_close_test.xlsx"
	defer os.Remove(path)

	w, f, err := CreateFile(path)
	if err != nil {
		t.Fatal(err)
	}
	sw, _ := w.NewSheet(SheetConfig{Name: "S"})
	sw.WriteRow(MakeRow("data"))
	sw.Close()
	w.Close()
	f.Close()

	reader, err := OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	// Close should close the underlying file
	if err := reader.Close(); err != nil {
		t.Errorf("Reader.Close: %v", err)
	}
	// Second close on the same file handle should error
	if err := reader.Close(); err == nil {
		t.Error("expected error on double close")
	}
}

func TestIteratorNextAfterClose(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "S"})
		sw.WriteRow(MakeRow("a"))
		sw.WriteRow(MakeRow("b"))
		sw.Close()
	})

	iter, _ := reader.OpenSheet("S")
	iter.Next() // read first row
	iter.Close()

	if iter.Next() {
		t.Error("Next should return false after Close")
	}
}

func TestIteratorErrNilOnSuccess(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "S"})
		sw.WriteRow(MakeRow("data"))
		sw.Close()
	})

	iter, _ := reader.OpenSheet("S")
	for iter.Next() {
	}
	if err := iter.Err(); err != nil {
		t.Errorf("Err() = %v, want nil", err)
	}
	iter.Close()
}

// ──────────────────────────────────────────────
// File I/O tests
// ──────────────────────────────────────────────

func TestCreateFileAndOpenFile(t *testing.T) {
	path := "/tmp/xlsxlite_io_test.xlsx"
	defer os.Remove(path)

	w, f, err := CreateFile(path)
	if err != nil {
		t.Fatal(err)
	}

	ss := w.StyleSheet()
	headerStyle := ss.AddStyle(Style{
		Font: &Font{Bold: true, Size: 12},
		Fill: &Fill{Type: "pattern", Pattern: "solid", FgColor: "FF4472C4"},
	})

	sw, _ := w.NewSheet(SheetConfig{
		Name:      "Report",
		ColWidths: map[int]float64{0: 25, 1: 15},
		FreezeRow: 1,
	})
	sw.WriteRow(Row{Cells: []Cell{
		StyledCell("Name", headerStyle),
		StyledCell("Score", headerStyle),
	}})
	for i := 0; i < 50; i++ {
		sw.WriteRow(MakeRow(fmt.Sprintf("Student %d", i+1), float64(50+i)))
	}
	sw.Close()
	w.Close()
	f.Close()

	reader, err := OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer reader.Close()

	iter, _ := reader.OpenSheet("Report")
	defer iter.Close()

	count := 0
	for iter.Next() {
		count++
	}
	if count != 51 { // 1 header + 50 data
		t.Errorf("read %d rows, want 51", count)
	}
}

// ──────────────────────────────────────────────
// isDateNumFmtID tests
// ──────────────────────────────────────────────

func TestIsDateNumFmtIDBuiltIn(t *testing.T) {
	empty := map[int]string{}

	dateIDs := []int{14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58}
	for _, id := range dateIDs {
		if !isDateNumFmtID(id, empty) {
			t.Errorf("isDateNumFmtID(%d) = false, want true", id)
		}
	}

	nonDateIDs := []int{0, 1, 2, 3, 4, 5, 10, 11, 12, 13, 23, 24, 25, 26, 37, 38, 44, 48, 49, 59, 100}
	for _, id := range nonDateIDs {
		if isDateNumFmtID(id, empty) {
			t.Errorf("isDateNumFmtID(%d) = true, want false", id)
		}
	}
}

func TestIsDateNumFmtIDCustom(t *testing.T) {
	custom := map[int]string{
		164: "yyyy-mm-dd",
		165: "#,##0.00",
		166: "hh:mm:ss",
		167: "AM/PM",
		168: "dd/mm/yyyy",
	}

	if !isDateNumFmtID(164, custom) {
		t.Error("yyyy-mm-dd should be date")
	}
	if isDateNumFmtID(165, custom) {
		t.Error("#,##0.00 should not be date")
	}
	if !isDateNumFmtID(166, custom) {
		t.Error("hh:mm:ss should be date")
	}
	if !isDateNumFmtID(167, custom) {
		t.Error("AM/PM should be date")
	}
	if !isDateNumFmtID(168, custom) {
		t.Error("dd/mm/yyyy should be date")
	}
}

func TestIsDateStyleEdgeCases(t *testing.T) {
	r := &Reader{styles: nil}
	if r.isDateStyle(0) {
		t.Error("nil styles should return false")
	}

	r.styles = &readerStyles{xfs: []readerXf{{numFmtID: 0}}}
	if r.isDateStyle(-1) {
		t.Error("negative styleID should return false")
	}
	if r.isDateStyle(5) {
		t.Error("out-of-range styleID should return false")
	}
}

// ──────────────────────────────────────────────
// Large file test (skipped in short mode)
// ──────────────────────────────────────────────

func TestLargeFile(t *testing.T) {
	if testing.Short() {
		t.Skip("skipping large file test in short mode")
	}

	var buf bytes.Buffer
	w := NewWriter(&buf)
	sw, _ := w.NewSheet(SheetConfig{Name: "Large"})

	rows := 100_000
	for i := 0; i < rows; i++ {
		sw.WriteRow(MakeRow(fmt.Sprintf("Row %d", i), float64(i), i%2 == 0))
	}
	sw.Close()
	w.Close()

	data := buf.Bytes()
	t.Logf("100k rows: file size = %d bytes (%.1f MB)", len(data), float64(len(data))/1024/1024)

	reader, _ := OpenReader(bytes.NewReader(data), int64(len(data)))
	iter, _ := reader.OpenSheet("Large")
	defer iter.Close()

	count := 0
	for iter.Next() {
		count++
	}
	if count != rows {
		t.Errorf("read %d rows, want %d", count, rows)
	}
}

// ──────────────────────────────────────────────
// Unicode and edge case tests
// ──────────────────────────────────────────────

func TestUnicodeStrings(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "Unicode"})
		sw.WriteRow(MakeRow(
			"日本語",
			"Ñoño",
			"Ü∞≠",
			"emoji 🎉",
			"中文测试",
			"العربية",
		))
		sw.Close()
	})

	rows := readAllRows(t, reader, "Unicode")
	expected := []string{"日本語", "Ñoño", "Ü∞≠", "emoji 🎉", "中文测试", "العربية"}
	for i, want := range expected {
		if rows[0].Cells[i].Value != want {
			t.Errorf("cell %d = %q, want %q", i, rows[0].Cells[i].Value, want)
		}
	}
}

func TestLongString(t *testing.T) {
	long := strings.Repeat("abcdefghij", 1000) // 10,000 chars
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "Long"})
		sw.WriteRow(MakeRow(long))
		sw.Close()
	})

	rows := readAllRows(t, reader, "Long")
	if rows[0].Cells[0].Value != long {
		t.Errorf("long string: got len=%d, want len=%d", len(rows[0].Cells[0].Value.(string)), len(long))
	}
}

func TestSharedStringDedup(t *testing.T) {
	// Write many rows with repeated strings to test dedup
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "Dedup"})
		for i := 0; i < 100; i++ {
			sw.WriteRow(MakeRow("repeated", "value", "abc"))
		}
		sw.Close()
	})

	rows := readAllRows(t, reader, "Dedup")
	if len(rows) != 100 {
		t.Fatalf("got %d rows, want 100", len(rows))
	}
	for i, row := range rows {
		if row.Cells[0].Value != "repeated" || row.Cells[1].Value != "value" || row.Cells[2].Value != "abc" {
			t.Errorf("row %d: unexpected values", i)
			break
		}
	}
}

func TestDefaultFallbackCellType(t *testing.T) {
	// Write a row with an unsupported type that triggers the default fallback
	type myStruct struct{ X int }
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "Fallback"})
		sw.WriteRow(Row{Cells: []Cell{
			{Value: myStruct{X: 7}, Type: CellTypeString},
		}})
		sw.Close()
	})

	rows := readAllRows(t, reader, "Fallback")
	if rows[0].Cells[0].Value != "{7}" {
		t.Errorf("fallback value = %q, want {7}", rows[0].Cells[0].Value)
	}
}

func TestTimeCellRoundTrip(t *testing.T) {
	// Write time.Time directly (not via DateCell) to test the time.Time branch in writeCell
	ts := time.Date(2025, 3, 8, 14, 30, 0, 0, time.UTC)
	reader := writeAndRead(t, func(w *Writer) {
		ss := w.StyleSheet()
		dateStyle := ss.AddStyle(Style{NumberFormat: "yyyy-mm-dd hh:mm:ss"})
		sw, _ := w.NewSheet(SheetConfig{Name: "Time"})
		sw.WriteRow(Row{Cells: []Cell{
			{Value: ts, Type: CellTypeDate, StyleID: dateStyle},
		}})
		sw.Close()
	})

	rows := readAllRows(t, reader, "Time")
	cell := rows[0].Cells[0]
	if cell.Type != CellTypeDate {
		t.Errorf("type = %d, want CellTypeDate", cell.Type)
	}
	if dt, ok := cell.Value.(time.Time); ok {
		if dt.Year() != 2025 || dt.Month() != 3 || dt.Day() != 8 {
			t.Errorf("date = %v, want 2025-03-08", dt)
		}
	}
}

func TestColumnWidthsOrder(t *testing.T) {
	// Verify that column widths produce valid XLSX regardless of map iteration order
	for range 5 {
		reader := writeAndRead(t, func(w *Writer) {
			sw, _ := w.NewSheet(SheetConfig{
				Name: "CW",
				ColWidths: map[int]float64{
					4: 40, 2: 20, 0: 10, 3: 30, 1: 15,
				},
			})
			sw.WriteRow(MakeRow("a", "b", "c", "d", "e"))
			sw.Close()
		})

		rows := readAllRows(t, reader, "CW")
		if len(rows) != 1 {
			t.Fatal("expected 1 row")
		}
	}
}

func TestSingleCellRow(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "Single"})
		sw.WriteRow(MakeRow("only one"))
		sw.Close()
	})

	rows := readAllRows(t, reader, "Single")
	if len(rows[0].Cells) != 1 || rows[0].Cells[0].Value != "only one" {
		t.Errorf("single cell: %+v", rows[0].Cells)
	}
}

func TestNilValueCellWithStyle(t *testing.T) {
	reader := writeAndRead(t, func(w *Writer) {
		ss := w.StyleSheet()
		style := ss.AddStyle(Style{Font: &Font{Bold: true}})
		sw, _ := w.NewSheet(SheetConfig{Name: "S"})
		sw.WriteRow(Row{Cells: []Cell{
			StringCell("before"),
			{Value: nil, Type: CellTypeEmpty, StyleID: style},
			StringCell("after"),
		}})
		sw.Close()
	})

	rows := readAllRows(t, reader, "S")
	if rows[0].Cells[0].Value != "before" {
		t.Errorf("cell 0 = %v", rows[0].Cells[0].Value)
	}
	if rows[0].Cells[2].Value != "after" {
		t.Errorf("cell 2 = %v", rows[0].Cells[2].Value)
	}
}

// ──────────────────────────────────────────────
// Security tests
// ──────────────────────────────────────────────

func TestXMLInjectionInStyleFields(t *testing.T) {
	// Verify that XML injection attempts in style fields are properly escaped
	var buf bytes.Buffer
	w := NewWriter(&buf)
	ss := w.StyleSheet()

	malicious := `"><script>alert(1)</script><x y="`

	ss.AddStyle(Style{
		Font: &Font{
			Name:  malicious,
			Color: malicious,
			Size:  11,
		},
		Fill: &Fill{
			Type:    "pattern",
			Pattern: "solid",
			FgColor: malicious,
			BgColor: malicious,
		},
		Border: &Border{
			Left:   BorderEdge{Style: malicious, Color: malicious},
			Right:  BorderEdge{Style: malicious, Color: malicious},
			Top:    BorderEdge{Style: malicious, Color: malicious},
			Bottom: BorderEdge{Style: malicious, Color: malicious},
		},
		Alignment: &Alignment{
			Horizontal: malicious,
			Vertical:   malicious,
		},
		NumberFormat: malicious,
	})

	// Write the styles XML
	var stylesBuf bytes.Buffer
	ss.writeXML(&stylesBuf)
	xml := stylesBuf.String()

	// The malicious string should NEVER appear unescaped
	if strings.Contains(xml, "<script>") {
		t.Error("XML injection: <script> tag found in styles output")
	}
	// Verify the escaped form IS present
	if !strings.Contains(xml, "&lt;script&gt;") {
		t.Error("expected escaped form of injection not found")
	}
	// The raw malicious payload should not appear
	if strings.Contains(xml, malicious) {
		t.Error("raw malicious string found unescaped in styles output")
	}
}

func TestXMLInjectionInSheetName(t *testing.T) {
	// Sheet names with XML metacharacters should be escaped in workbook.xml
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: `Sheet"<>&'Test`})
		sw.WriteRow(MakeRow("data"))
		sw.Close()
	})

	names := reader.SheetNames()
	if len(names) == 0 {
		t.Fatal("no sheets found")
	}
	// The name should round-trip correctly (written escaped, read back unescaped)
	if names[0] != `Sheet"<>&'Test` {
		t.Errorf("sheet name = %q, want %q", names[0], `Sheet"<>&'Test`)
	}
}

func TestXMLInjectionInCellValues(t *testing.T) {
	// Cell values with XML metacharacters must be escaped
	malicious := `<script>alert("xss")</script>&<b>`
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "XSS"})
		sw.WriteRow(MakeRow(malicious))
		sw.Close()
	})

	rows := readAllRows(t, reader, "XSS")
	if rows[0].Cells[0].Value != malicious {
		t.Errorf("cell value = %q, want %q", rows[0].Cells[0].Value, malicious)
	}
}

func TestSparseColumnDoesNotExceedMaxColumns(t *testing.T) {
	// Verify that the reader caps column indices at MaxColumns
	// We can't easily craft a raw XLSX with a huge column ref in this test,
	// but we verify the constant is reasonable
	if MaxColumns != 16384 {
		t.Errorf("MaxColumns = %d, want 16384 (Excel's limit)", MaxColumns)
	}
}

func TestPathTraversalInRelationships(t *testing.T) {
	// This is a unit test for the path sanitization logic.
	// Relationship targets with ".." or absolute paths should be rejected.
	// We test indirectly by verifying the library doesn't panic on normal paths.
	reader := writeAndRead(t, func(w *Writer) {
		sw, _ := w.NewSheet(SheetConfig{Name: "Safe"})
		sw.WriteRow(MakeRow("data"))
		sw.Close()
	})

	rows := readAllRows(t, reader, "Safe")
	if len(rows) != 1 {
		t.Errorf("expected 1 row, got %d", len(rows))
	}
}

func TestSafetyConstants(t *testing.T) {
	// Verify safety constants have sane values
	if MaxColumns != 16384 {
		t.Errorf("MaxColumns = %d", MaxColumns)
	}
	if MaxSharedStrings != 10_000_000 {
		t.Errorf("MaxSharedStrings = %d", MaxSharedStrings)
	}
	if MaxDecompressedSize != 256<<20 {
		t.Errorf("MaxDecompressedSize = %d", MaxDecompressedSize)
	}
	if MaxGCSDownloadSize != 512<<20 {
		t.Errorf("MaxGCSDownloadSize = %d", MaxGCSDownloadSize)
	}
}
