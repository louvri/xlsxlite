// Package gcs provides Google Cloud Storage helpers for xlsxlite.
//
// Reading: OpenFile downloads entirely into memory; OpenFileLowMem
// spills to a temp file to keep heap usage low for large files.
//
// Writing streams directly to a GCS object writer.
package gcs

import (
	"bytes"
	"context"
	"fmt"
	"io"
	"os"

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

// OpenFileLowMem opens an XLSX file from GCS using a temporary file on disk
// instead of holding the entire file in memory. This is ideal for large files
// (e.g. 500k+ rows) where the file content would otherwise stay in heap.
//
// The returned Reader uses an os.File-backed io.ReaderAt, so zip decompression
// reads from disk. The temp file is automatically deleted when Reader.Close is called.
func OpenFileLowMem(ctx context.Context, client *storage.Client, bucket, object string) (*xlsxlite.Reader, error) {
	rc, err := client.Bucket(bucket).Object(object).NewReader(ctx)
	if err != nil {
		return nil, fmt.Errorf("gcs open %s/%s: %w", bucket, object, err)
	}
	defer rc.Close()

	tmp, err := os.CreateTemp("", "xlsxlite-*.xlsx")
	if err != nil {
		return nil, fmt.Errorf("create temp file: %w", err)
	}

	// On any error below, clean up the temp file.
	cleanup := func() {
		tmp.Close()
		os.Remove(tmp.Name())
	}

	written, err := io.Copy(tmp, io.LimitReader(rc, xlsxlite.MaxGCSDownloadSize))
	if err != nil {
		cleanup()
		return nil, fmt.Errorf("gcs download %s/%s: %w", bucket, object, err)
	}

	if _, err := tmp.Seek(0, io.SeekStart); err != nil {
		cleanup()
		return nil, fmt.Errorf("seek temp file: %w", err)
	}

	reader, err := xlsxlite.OpenReader(tmp, written)
	if err != nil {
		cleanup()
		return nil, err
	}

	// Attach cleanup so Reader.Close() removes the temp file.
	tmpName := tmp.Name()
	reader.SetCloser(fileCloser{file: tmp, path: tmpName})

	return reader, nil
}

// fileCloser closes the file handle and removes the file from disk.
type fileCloser struct {
	file *os.File
	path string
}

func (fc fileCloser) Close() error {
	fc.file.Close()
	return os.Remove(fc.path)
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
