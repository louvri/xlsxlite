package xlsxlite

import (
	"bytes"
	"fmt"
	"testing"
)

func BenchmarkWrite100kRows(b *testing.B) {
	b.ReportAllocs()
	for n := 0; n < b.N; n++ {
		var buf bytes.Buffer
		w := NewWriter(&buf)

		sw, err := w.NewSheet(SheetConfig{Name: "Bench"})
		if err != nil {
			b.Fatal(err)
		}

		for i := 0; i < 100_000; i++ {
			err := sw.WriteRow(MakeRow(
				fmt.Sprintf("Row %d", i),
				float64(i),
				i%2 == 0,
				"constant string value",
			))
			if err != nil {
				b.Fatal(err)
			}
		}

		if err := sw.Close(); err != nil {
			b.Fatal(err)
		}
		if err := w.Close(); err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkRead100kRows(b *testing.B) {
	// Prepare a 100k row file
	var buf bytes.Buffer
	w := NewWriter(&buf)
	sw, _ := w.NewSheet(SheetConfig{Name: "Bench"})
	for i := 0; i < 100_000; i++ {
		_ = sw.WriteRow(MakeRow(
			fmt.Sprintf("Row %d", i),
			float64(i),
			i%2 == 0,
			"constant string value",
		))
	}
	_ = sw.Close()
	_ = w.Close()
	data := buf.Bytes()

	b.ResetTimer()
	b.ReportAllocs()

	for n := 0; n < b.N; n++ {
		reader, err := OpenReader(bytes.NewReader(data), int64(len(data)))
		if err != nil {
			b.Fatal(err)
		}

		iter, err := reader.OpenSheet("Bench")
		if err != nil {
			b.Fatal(err)
		}

		count := 0
		for iter.Next() {
			_ = iter.Row()
			count++
		}
		iter.Close()

		if count != 100_000 {
			b.Fatalf("read %d rows", count)
		}
	}
}

func BenchmarkWrite1MRows(b *testing.B) {
	if testing.Short() {
		b.Skip("skipping 1M row benchmark")
	}
	b.ReportAllocs()
	for n := 0; n < b.N; n++ {
		var buf bytes.Buffer
		w := NewWriter(&buf)
		sw, _ := w.NewSheet(SheetConfig{Name: "Bench"})

		for i := 0; i < 1_000_000; i++ {
			_ = sw.WriteRow(MakeRow(
				fmt.Sprintf("Row %d", i),
				float64(i),
				i%2 == 0,
			))
		}

		_ = sw.Close()
		_ = w.Close()
	}
}
