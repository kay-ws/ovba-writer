# ovba-writer

Pure-Go reader and writer for `vbaProject.bin` — the binary VBA container inside
Excel `.xlsm`/`.xltm` workbooks. No Excel, no COM, no Windows. Runs anywhere Go
runs, so VBA source can be read and written from Linux, CI, and agent sandboxes.

## Why

Everything in a macro-enabled workbook is plain OOXML XML — except one file,
`xl/vbaProject.bin`, which is an OLE2/[MS-CFB] compound file holding
[MS-OVBA]-compressed VBA source (and a per-version compiled cache). That single
file is the only reason round-tripping VBA has needed a Windows/Excel host.

`ovba-writer` reads and writes that file directly in Go, so the macro source of a
workbook can be extracted, edited, and written back without launching Excel.

## Status

Experimental (`v0`). The writer is **source-only by design** (see
[Scope](#scope--limitations)) and is validated three ways:

- **Byte-exact** compressed module streams against golden fixtures (standard and
  document modules); the CFB envelope is validated for semantic equivalence by
  reading it back with an independent reader.
- **Independent extraction** parity with `olevba` (oletools), including class
  modules.
- **Real Excel**: workbooks written by this library open in Excel, recompile
  from source, and run the edited macros.

## Install

```sh
go get github.com/kay-ws/ovba-writer
```

## Quick start

`vbaproject` is the high-level read-modify-write API. Feed it the bytes of
`xl/vbaProject.bin` (extract it from the `.xlsm` zip yourself, e.g. with
`archive/zip`).

```go
import "github.com/kay-ws/ovba-writer/vbaproject"

proj, err := vbaproject.Read(bin) // bin = contents of xl/vbaProject.bin
if err != nil {
    return err
}

for i := range proj.Modules {
    if proj.Modules[i].Name == "Module1" {
        proj.Modules[i].Source =
            "Attribute VB_Name = \"Module1\"\r\n" +
            "Sub Hello()\r\n" +
            "    Debug.Print \"edited from Go\"\r\n" +
            "End Sub\r\n"
    }
}

out, err := vbaproject.Write(proj) // out = new xl/vbaProject.bin
```

`Read` → edit `Module.Source` → `Write` is the whole contract. References,
project properties, and protection fields (`CMG`/`DPB`/`GC`) are preserved
verbatim; `Write` reconstructs only the module list and module streams.

If your source comes from on-disk `.bas`/`.cls`/document files (e.g. a VBA
project exported to a source tree), normalize it to the in-bin representation
first:

```go
src, err := vbaproject.NormalizeModuleSource(vbaproject.ModuleClass, diskText, existing)
```

## How it works

Three layers, each independently testable:

| Package      | Layer                | Spec                        |
| ------------ | -------------------- | --------------------------- |
| `cfb`        | OLE2 envelope        | [MS-CFB] (reader + writer)  |
| `ovba`       | VBA content + RLE compression | [MS-OVBA]          |
| `vbaproject` | high-level model + read-modify-write | wires the two layers |

See [`docs/DESIGN.md`](docs/DESIGN.md) for the format walkthrough, the
source-only rationale, and the spec section map.

## Scope & limitations

- **Source-only by design.** The per-version compiled p-code (`PerformanceCache`)
  is never emitted. [MS-OVBA] specifies that this cache is implementation-
  specific and "must be ignored on read," and that the interoperable (source)
  representation is used across version boundaries — Excel recompiles from source
  on a version mismatch. This library writes only that documented, interoperable
  representation; it never touches the undocumented p-code bytecode.
- **Standard, class, and document modules** are editable. UserForms (`.frm`/
  `.frx`) are detected but not authored (a host fallback is expected for those).
- **Protected projects** are read with `CMG`/`DPB`/`GC` preserved verbatim; the
  password protection is not decrypted.

## References

Primary sources (Microsoft Open Specifications):

- [MS-OVBA]: Office VBA File Format Structure —
  <https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc>
- [MS-CFB]: Compound File Binary File Format —
  <https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b>
- [MS-WINERRATA] errata for [MS-CFB] —
  <https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-winerrata/c14df6f6-ae1d-45bf-b62f-7b2ed199ba44>

## License

[MIT](LICENSE)

[MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc
[MS-CFB]: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b
