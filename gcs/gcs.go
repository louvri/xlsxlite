// Package gcs provides Google Cloud Storage helpers for xlsxlite.
//
// Reading downloads the object into memory (required because zip.Reader
// needs io.ReaderAt, which GCS objects don't implement).
//
// Writing streams directly to a GCS object writer.
package gcs

import (
	"bytes"
	"context"
	"fmt"
	"io"

	"cloud.google.com/go/storage"
	"github.com/louvri/xlsxlite"
)

// OpenFile opens an XLSX file from GCS for streaming reading.
// The entire object is downloaded into memory (up to MaxGCSDownloadSize) since
// zip.Reader requires io.ReaderAt. The returned Reader does not hold any GCS
// resources — it reads from the in-memory copy.
func OpenFile(ctx context.Context, client *storage.Client, bucket, object string) (*xlsxlite.Reader, error) {
	rc, err := client.Bucket(bucket).Object(object).NewReader(ctx)
	if err != nil {
		return nil, fmt.Errorf("gcs open %s/%s: %w", bucket, object, err)
	}
	defer rc.Close()

	data, err := io.ReadAll(io.LimitReader(rc, xlsxlite.MaxGCSDownloadSize))
	if err != nil {
		return nil, fmt.Errorf("gcs read %s/%s: %w", bucket, object, err)
	}

	return xlsxlite.OpenReader(bytes.NewReader(data), int64(len(data)))
}

// CreateFile creates a new XLSX file on GCS for streaming writing.
// The returned Writer streams rows directly into the GCS object writer.
// Call Writer.Close() first to finalize the XLSX package, then call the
// returned closer function to complete the GCS upload.
//
// Example:
//
//	w, closeFn, err := gcs.CreateFile(ctx, client, "bucket", "path/to/file.xlsx")
//	if err != nil { ... }
//	// write sheets and rows to w ...
//	if err := w.Close(); err != nil { ... }
//	if err := closeFn(); err != nil { ... }
func CreateFile(ctx context.Context, client *storage.Client, bucket, object string) (*xlsxlite.Writer, func() error, error) {
	obj := client.Bucket(bucket).Object(object)
	gw := obj.NewWriter(ctx)
	gw.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

	w := xlsxlite.NewWriter(gw)
	closer := func() error {
		return gw.Close()
	}
	return w, closer, nil
}
