package xlsxlite

import (
	"fmt"
	"io"
	"strings"
)

// StyleSheet manages Excel styles and generates styles.xml content.
// Styles in XLSX are composed of 4 independent tables (fonts, fills, borders, numFmts)
// and a cross-reference table (cellXfs) that combines indices from each.
type StyleSheet struct {
	fonts    []Font
	fills    []Fill
	borders  []Border
	numFmts  []numFmt
	xfs      []cellXf
	styles   []Style // original style objects, for dedup
}

type numFmt struct {
	id     int
	format string
}

type cellXf struct {
	fontID    int
	fillID    int
	borderID  int
	numFmtID  int
	alignment *Alignment
}

// NewStyleSheet creates a StyleSheet with the defaults required by Excel:
// one font (Calibri 11), two fills (none + gray125), one border (empty),
// and one default cell format. This is called automatically by NewWriter.
func NewStyleSheet() *StyleSheet {
	ss := &StyleSheet{}

	// Default font (index 0)
	ss.fonts = append(ss.fonts, Font{Name: "Calibri", Size: 11})

	// Excel requires these two fills at index 0 and 1
	ss.fills = append(ss.fills, Fill{Type: "none"})
	ss.fills = append(ss.fills, Fill{Type: "pattern", Pattern: "gray125"})

	// Default border (index 0) — no borders
	ss.borders = append(ss.borders, Border{})

	// Default xf (index 0)
	ss.xfs = append(ss.xfs, cellXf{})
	ss.styles = append(ss.styles, Style{})

	return ss
}

// AddStyle registers a Style and returns its 0-based index for use in Cell.StyleID.
// Identical styles are deduplicated — calling AddStyle with the same parameters
// multiple times returns the same index.
func (ss *StyleSheet) AddStyle(s Style) int {
	fontID := ss.addFont(s.Font)
	fillID := ss.addFill(s.Fill)
	borderID := ss.addBorder(s.Border)
	numFmtID := ss.addNumFmt(s.NumberFormat)

	xf := cellXf{
		fontID:    fontID,
		fillID:    fillID,
		borderID:  borderID,
		numFmtID:  numFmtID,
		alignment: s.Alignment,
	}

	// Check for existing identical xf
	for i, existing := range ss.xfs {
		if existing.fontID == xf.fontID &&
			existing.fillID == xf.fillID &&
			existing.borderID == xf.borderID &&
			existing.numFmtID == xf.numFmtID &&
			alignmentEqual(existing.alignment, xf.alignment) {
			return i
		}
	}

	ss.xfs = append(ss.xfs, xf)
	ss.styles = append(ss.styles, s)
	return len(ss.xfs) - 1
}

func alignmentEqual(a, b *Alignment) bool {
	if a == nil && b == nil {
		return true
	}
	if a == nil || b == nil {
		return false
	}
	return *a == *b
}

func (ss *StyleSheet) addFont(f *Font) int {
	if f == nil {
		return 0
	}
	for i, existing := range ss.fonts {
		if existing == *f {
			return i
		}
	}
	ss.fonts = append(ss.fonts, *f)
	return len(ss.fonts) - 1
}

func (ss *StyleSheet) addFill(f *Fill) int {
	if f == nil {
		return 0
	}
	for i, existing := range ss.fills {
		if existing == *f {
			return i
		}
	}
	ss.fills = append(ss.fills, *f)
	return len(ss.fills) - 1
}

func (ss *StyleSheet) addBorder(b *Border) int {
	if b == nil {
		return 0
	}
	for i, existing := range ss.borders {
		if existing == *b {
			return i
		}
	}
	ss.borders = append(ss.borders, *b)
	return len(ss.borders) - 1
}

func (ss *StyleSheet) addNumFmt(format string) int {
	if format == "" {
		return 0 // General
	}
	for _, nf := range ss.numFmts {
		if nf.format == format {
			return nf.id
		}
	}
	// Custom number formats start at ID 164
	id := 164 + len(ss.numFmts)
	ss.numFmts = append(ss.numFmts, numFmt{id: id, format: format})
	return id
}

// writeXML writes the complete styles.xml content.
func (ss *StyleSheet) writeXML(w io.Writer) error {
	write := func(s string) error {
		_, err := io.WriteString(w, s)
		return err
	}

	if err := write(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"); err != nil {
		return err
	}
	if err := write(`<styleSheet xmlns="` + nsSpreadsheetML + `">` + "\n"); err != nil {
		return err
	}

	// Number formats
	if len(ss.numFmts) > 0 {
		if err := write(fmt.Sprintf(`<numFmts count="%d">`, len(ss.numFmts))); err != nil {
			return err
		}
		for _, nf := range ss.numFmts {
			if err := write(fmt.Sprintf(`<numFmt numFmtId="%d" formatCode="%s"/>`,
				nf.id, escapeXMLAttr(nf.format))); err != nil {
				return err
			}
		}
		if err := write(`</numFmts>` + "\n"); err != nil {
			return err
		}
	}

	// Fonts
	if err := write(fmt.Sprintf(`<fonts count="%d">`, len(ss.fonts))); err != nil {
		return err
	}
	for _, f := range ss.fonts {
		if err := write(`<font>`); err != nil {
			return err
		}
		if f.Bold {
			if err := write(`<b/>`); err != nil {
				return err
			}
		}
		if f.Italic {
			if err := write(`<i/>`); err != nil {
				return err
			}
		}
		if f.Underline {
			if err := write(`<u/>`); err != nil {
				return err
			}
		}
		if f.Size > 0 {
			if err := write(fmt.Sprintf(`<sz val="%g"/>`, f.Size)); err != nil {
				return err
			}
		}
		if f.Color != "" {
			if err := write(fmt.Sprintf(`<color rgb="%s"/>`, escapeXMLAttr(f.Color))); err != nil {
				return err
			}
		}
		if f.Name != "" {
			if err := write(fmt.Sprintf(`<name val="%s"/>`, escapeXMLAttr(f.Name))); err != nil {
				return err
			}
		}
		if err := write(`</font>`); err != nil {
			return err
		}
	}
	if err := write(`</fonts>` + "\n"); err != nil {
		return err
	}

	// Fills
	if err := write(fmt.Sprintf(`<fills count="%d">`, len(ss.fills))); err != nil {
		return err
	}
	for _, f := range ss.fills {
		if err := write(`<fill>`); err != nil {
			return err
		}
		pattern := f.Pattern
		if f.Type == "none" {
			pattern = "none"
		}
		if pattern == "" {
			pattern = "none"
		}
		if err := write(fmt.Sprintf(`<patternFill patternType="%s"`, escapeXMLAttr(pattern))); err != nil {
			return err
		}
		if f.FgColor != "" || f.BgColor != "" {
			if err := write(`>`); err != nil {
				return err
			}
			if f.FgColor != "" {
				if err := write(fmt.Sprintf(`<fgColor rgb="%s"/>`, escapeXMLAttr(f.FgColor))); err != nil {
					return err
				}
			}
			if f.BgColor != "" {
				if err := write(fmt.Sprintf(`<bgColor rgb="%s"/>`, escapeXMLAttr(f.BgColor))); err != nil {
					return err
				}
			}
			if err := write(`</patternFill>`); err != nil {
				return err
			}
		} else {
			if err := write(`/>`); err != nil {
				return err
			}
		}
		if err := write(`</fill>`); err != nil {
			return err
		}
	}
	if err := write(`</fills>` + "\n"); err != nil {
		return err
	}

	// Borders
	if err := write(fmt.Sprintf(`<borders count="%d">`, len(ss.borders))); err != nil {
		return err
	}
	for _, b := range ss.borders {
		if err := write(`<border>`); err != nil {
			return err
		}
		writeBorderEdge := func(tag string, e BorderEdge) error {
			if e.Style == "" {
				return write(fmt.Sprintf(`<%s/>`, tag))
			}
			s := fmt.Sprintf(`<%s style="%s"`, tag, escapeXMLAttr(e.Style))
			if e.Color != "" {
				s += fmt.Sprintf(`><color rgb="%s"/></%s>`, escapeXMLAttr(e.Color), tag)
			} else {
				s += `/>`
			}
			return write(s)
		}
		if err := writeBorderEdge("left", b.Left); err != nil {
			return err
		}
		if err := writeBorderEdge("right", b.Right); err != nil {
			return err
		}
		if err := writeBorderEdge("top", b.Top); err != nil {
			return err
		}
		if err := writeBorderEdge("bottom", b.Bottom); err != nil {
			return err
		}
		if err := write(`<diagonal/>`); err != nil {
			return err
		}
		if err := write(`</border>`); err != nil {
			return err
		}
	}
	if err := write(`</borders>` + "\n"); err != nil {
		return err
	}

	// Cell style xfs (required, at least 1)
	if err := write(`<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` + "\n"); err != nil {
		return err
	}

	// Cell xfs
	if err := write(fmt.Sprintf(`<cellXfs count="%d">`, len(ss.xfs))); err != nil {
		return err
	}
	for _, xf := range ss.xfs {
		attrs := fmt.Sprintf(`numFmtId="%d" fontId="%d" fillId="%d" borderId="%d"`,
			xf.numFmtID, xf.fontID, xf.fillID, xf.borderID)
		if xf.fontID > 0 {
			attrs += ` applyFont="1"`
		}
		if xf.fillID > 0 {
			attrs += ` applyFill="1"`
		}
		if xf.borderID > 0 {
			attrs += ` applyBorder="1"`
		}
		if xf.numFmtID > 0 {
			attrs += ` applyNumberFormat="1"`
		}
		if xf.alignment != nil {
			attrs += ` applyAlignment="1"`
			if err := write(fmt.Sprintf(`<xf %s>`, attrs)); err != nil {
				return err
			}
			alignAttrs := ""
			if xf.alignment.Horizontal != "" {
				alignAttrs += fmt.Sprintf(` horizontal="%s"`, escapeXMLAttr(xf.alignment.Horizontal))
			}
			if xf.alignment.Vertical != "" {
				alignAttrs += fmt.Sprintf(` vertical="%s"`, escapeXMLAttr(xf.alignment.Vertical))
			}
			if xf.alignment.WrapText {
				alignAttrs += ` wrapText="1"`
			}
			if err := write(fmt.Sprintf(`<alignment%s/>`, alignAttrs)); err != nil {
				return err
			}
			if err := write(`</xf>`); err != nil {
				return err
			}
		} else {
			if err := write(fmt.Sprintf(`<xf %s/>`, attrs)); err != nil {
				return err
			}
		}
	}
	if err := write(`</cellXfs>` + "\n"); err != nil {
		return err
	}

	// Cell styles (required)
	if err := write(`<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` + "\n"); err != nil {
		return err
	}

	return write(`</styleSheet>`)
}

func escapeXMLAttr(s string) string {
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, `"`, "&quot;")
	return s
}
