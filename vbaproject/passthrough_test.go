package vbaproject

import (
	"bytes"
	"testing"

	"github.com/kay-ws/ovba-writer/cfb"
)

// PROJECTwm is a root-level stream that lives outside the VBA storage and the
// PROJECT stream. The minimal-profile writer used to drop every stream it did
// not regenerate; the structural-boundary pass-through must instead carry any
// stream whose first path segment is neither "VBA" nor "PROJECT" through a
// Read -> Write round-trip byte-for-byte. PROJECTwm is the witness on a
// non-form fixture (the same mechanism later carries UserForm designer storages).
func TestWritePreservesRootStreamsOutsideVBA(t *testing.T) {
	in := loadBin(t, "p2_refs.bin")

	orig, err := cfb.Open(in)
	if err != nil {
		t.Fatal(err)
	}
	want, ok := orig.Stream("PROJECTwm")
	if !ok {
		t.Skip("fixture has no PROJECTwm; nothing to assert")
	}

	p, err := Read(in)
	if err != nil {
		t.Fatal(err)
	}
	out, err := Write(p)
	if err != nil {
		t.Fatalf("write: %v", err)
	}

	got, err := cfb.Open(out)
	if err != nil {
		t.Fatal(err)
	}
	have, ok := got.Stream("PROJECTwm")
	if !ok {
		t.Fatal("PROJECTwm was dropped: root streams outside VBA/ must pass through")
	}
	if !bytes.Equal(have, want) {
		t.Errorf("PROJECTwm not preserved verbatim: %d bytes -> %d bytes", len(want), len(have))
	}
}
