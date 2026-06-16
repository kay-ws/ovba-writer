# Design

This document is the technical companion to the [README](../README.md). It explains
the on-disk formats `ovba-writer` works with, how the three packages map to them,
and the design decisions that make a source-only writer correct and safe.

## Overview

A macro-enabled Excel workbook (`.xlsm`/`.xltm`) is an OOXML package: a ZIP of XML
parts. Every part is human-readable XML except one — `xl/vbaProject.bin`. That file
is an OLE2 / [MS-CFB] *compound file* (a small file system inside a file) whose
streams hold [MS-OVBA]-compressed VBA source plus a per-version compiled cache.
Because that single opaque blob has historically required a Windows/Excel host to
round-trip, editing VBA from other platforms has been awkward.

`ovba-writer` reads and writes `vbaProject.bin` directly in pure Go: extract the
part from the workbook ZIP, edit module source, and write a new part back, all
without Excel, COM, or Windows.

## Architecture: three layers

The code is three packages, each independently testable and ignorant of the layers
above it.

| Package      | Layer                                   | Spec                       |
| ------------ | --------------------------------------- | -------------------------- |
| `cfb`        | OLE2 envelope (reader + writer)         | [MS-CFB]                   |
| `ovba`       | VBA content: RLE compression + `dir`/PROJECT assembly | [MS-OVBA]        |
| `vbaproject` | high-level `Project` model, read-modify-write | wires the two layers |

- **`cfb`** parses and emits the compound-file envelope. It deals only in named
  streams and storages; it knows nothing about VBA. The reader is general; the
  writer is a deliberately *minimal profile* (see below).
- **`ovba`** implements the content layer: the [MS-OVBA] compression codec and the
  assembly/parsing of the `dir` stream records and the `PROJECT` text stream. It
  takes and returns plain byte slices and knows nothing about CFB.
- **`vbaproject`** is the public, high-level API. It opens a `cfb.Container`, decodes
  the content with `ovba`, exposes a `*Project` model for editing, and reassembles a
  new bin. This is the read-modify-write center.

## MS-CFB envelope essentials

The compound file begins with a 512-byte header: the signature
`D0 CF 11 E0 A1 B1 1A E1`, a major version, and the sector geometry. The reader
accepts **major version 3 only**, which fixes sectors at **512 bytes**
(`sectorShift = 9`) and mini-sectors at **64 bytes** (`miniSectorShift = 6`). The
mini-stream cutoff is **4096 bytes**: a stream smaller than the cutoff is stored in
the mini-FAT chain (packed into the Root Entry's mini-stream container), while a
larger stream occupies whole 512-byte sectors via the regular FAT.

Sector allocation is tracked by the **FAT**, addressed through the **DIFAT** (109
entries inline in the header, then optional DIFAT-sector chains). Small streams use
the parallel **mini-FAT**. The reserved sector values are `FREESECT` (`0xFFFFFFFF`),
`ENDOFCHAIN` (`0xFFFFFFFE`), `FATSECT` (`0xFFFFFFFD`), and `DIFSECT` (`0xFFFFFFFC`).

Streams and storages are described by 128-byte **directory entries**, four per
sector. The name field is 64 bytes of UTF-16LE plus a NUL terminator, and the
stored length counts the terminator. Sibling entries at one level form a tree keyed
by name; a storage's `child` field points at the root of its children's tree.

### Name ordering

[MS-CFB] §2.6.4 orders names by **UTF-16 code-unit length first**, and only then by
case-insensitive (upper-cased) code-unit comparison. So `"dir"` sorts before
`"PROJECT"` because it is shorter, not because of its letters. Both the reader's
in-order walk and the writer's tree construction follow this rule.

### Red-black coloring is optional

The directory tree is specified as a red-black tree, but balancing via node color
is a SHOULD: major implementations ignore the color flag and treat the structure as
an ordered binary search tree (see [MS-WINERRATA]). The writer therefore builds a
balanced, name-ordered BST (median-as-root) and marks every node **black**; it never
performs red-black rotations. Readers that honor the errata reconstruct the same
ordering regardless.

### Minimal-profile writer

The writer targets the small, regular shape that `vbaProject.bin` actually takes. It
lays sectors out in a fixed order (large streams, mini-stream container, mini-FAT,
directory, FAT) and fails loudly if the FAT would need more than the 109 inline
DIFAT entries — i.e. it does not emit DIFAT sectors. It does not aim to reproduce
Excel's envelope byte-for-byte; correctness is defined as *semantic equivalence* —
reading the output back yields the same tree and the same stream contents. (Byte
exactness is asserted one layer down, on the compressed streams; see
[Verification](#verification).)

## MS-OVBA compression essentials

Compressed data lives in a **CompressedContainer**: a single signature byte `0x01`
followed by a sequence of **CompressedChunks**, each covering up to 4096 bytes of
decompressed output. A chunk has a 2-byte little-endian header: bits 0–11 hold
`size − 3` (the total chunk byte count including the header), bits 12–14 are the
signature `0b011`, and bit 15 is the **compressed flag**.

The flag is decided by a single test: **does the compressed chunk fit in 4096
bytes?** It is *not* "did compression shrink the input." A chunk whose token stream
grows past 4096 bytes is emitted raw — flag `0`, body zero-padded to exactly 4096
bytes — but a chunk that ends up *larger than its input yet still under 4096 bytes*
stays compressed. ([MS-OVBA] §2.4.1.3.7.)

Within a compressed chunk, a **FlagByte** precedes each group of up to eight tokens;
each bit selects a literal byte (`0`) or a 2-byte **CopyToken** (`1`). A CopyToken
encodes an (offset, length) back-reference. The bit split depends on the current
position `pos` within the chunk's decompressed output:

```
bitCount       = max(ceil(log2(pos)), 4)
lengthBitCount = 16 - bitCount
token          = ((offset - 1) << lengthBitCount) | (length - 3)
maxLength      = (0xFFFF >> bitCount) + 3
```

So early in a chunk fewer bits go to offset and more to length, and the split widens
as output grows ([MS-OVBA] §2.4.1.3.19.2–.3). Matches shorter than 3 bytes are not
worth a token and are emitted as literals.

**Overlapping copy (RLE) is allowed**: a CopyToken may reference output that the copy
itself is still producing (e.g. offset 1, length 3 over a run of identical bytes), so
copying proceeds one byte at a time rather than as a block move ([MS-OVBA]
§2.4.1.3.19.4). The encoder's greedy longest-match search and the decoder's copy loop
both honor this.

## The `dir` and module streams

`VBA/dir` is itself a CompressedContainer. Decompressed, it is a flat sequence of
`id(2) + size(4) + payload` records (one exception: **PROJECTVERSION**, id `0x0009`,
has a fixed 12-byte layout with no size field — [MS-OVBA] §2.3.4.2.1.10). The parser
walks these records to recover project attributes (SYSKIND, LCID, **CODEPAGE**), the
reference section, and the per-module entries.

A module's `dir` entry carries its `MODULENAME`, `MODULESTREAMNAME`, `MODULEOFFSET`,
and **MODULETYPE**. The type record id is the type: `0x0021` = procedural / standard
module, `0x0022` = non-procedural (class or document) module ([MS-OVBA]
§2.3.4.2.3.2.8).

The source for each module lives in the stream named by `MODULESTREAMNAME` under
`VBA/`. That stream is laid out as:

```
[ PerformanceCache ] [ CompressedSourceCode ]
                     ^
                     MODULEOFFSET
```

`MODULEOFFSET` is the byte boundary between the compiled cache and the compressed
source. Reading a module is therefore: skip `MODULEOFFSET` bytes → `Decompress` the
remainder → decode the bytes per the project **CODEPAGE**.

## Source-only design and the p-code rationale

**Source-only by design.** The per-version compiled p-code (the `PerformanceCache`)
is never emitted. [MS-OVBA] (Module Stream) specifies this cache is
implementation-specific and version-dependent, is `MODULEOFFSET` bytes in size, and
*must be ignored on read*; across version boundaries the interoperable (source)
representation is used instead, so Excel recompiles from source on a version
mismatch. The library writes only that documented interoperable representation and
never touches the undocumented p-code bytecode (whose internals are known only from
reverse-engineering work such as `pcodedmp`). The practical guarantee that a given
Excel build recompiles and runs a source-only bin was confirmed empirically against
real Excel.

Concretely, the writer sets `MODULEOFFSET = 0` for every module (so the whole stream
is `CompressedSourceCode`) and emits a fixed 7-byte `_VBA_PROJECT` stub
(`CC 61 FF FF 00 03 00`) in place of any compiled project cache. Because the
PerformanceCache is gone, Excel sees a version mismatch and recompiles from source on
open.

## Module source normalization

`NormalizeModuleSource(mt, disk, existing)` converts on-disk source text (a `.bas`,
`.cls`, or document export) into the exact in-bin representation a module stream
expects. It first normalizes line endings to CRLF, then dispatches by module type.
There are three kinds:

- **Standard** (`ModuleStd`) — **identity**. The in-bin form already matches the disk
  form (an `Attribute VB_Name` line followed by code). The only guard is that
  `Attribute VB_Name` must be present; its absence signals corrupt input and is a
  hard error rather than a silent pass.

- **Class** (`ModuleClass`) — strip the leading `VERSION ... BEGIN ... END` header
  block that disk `.cls` files carry, then inject the attributes Excel keeps inline.
  Specifically, after the `Attribute VB_Name` line it inserts:

  ```
  Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
  ```

  and after the `Attribute VB_Exposed` line it inserts:

  ```
  Attribute VB_TemplateDerived = False
  Attribute VB_Customizable = False
  ```

  Both `VB_Name` and `VB_Exposed` must appear (they are the injection anchors); if
  either is missing, normalization fails. The `VB_Base` GUID above is the generic,
  host-independent class base — distinct from the host GUIDs carried by document
  modules.

- **Document** (`ModuleDocument`) — a document export (e.g. `ThisWorkbook`, a sheet)
  is attribute-less pure code, so the attributes cannot be reconstructed from disk.
  Normalization requires the `existing` in-bin module and **prepends its leading
  contiguous `Attribute ` header block** to the disk code. The document's host GUID,
  `VB_PredeclaredId`, and so on therefore come from the existing module and are never
  hardcoded.

## Encoding boundary

There are two encodings, and the boundary between them is explicit:

- **On disk**, source is UTF-8 (no BOM) with CRLF line endings.
- **In the bin**, a module stream is encoded per the project **CODEPAGE** (e.g. 932 /
  Shift-JIS, 1252 / Windows-1252).

`Read` decodes each module stream from its CODEPAGE into a UTF-8 Go string, so
`Module.Source` is always UTF-8. `NormalizeModuleSource` is a pure string→string
transform and is encoding-agnostic. The MBCS encode happens only in the write path
(`Write` → `encodeMBCS`), and it fails loudly if a character cannot be represented in
the target code page rather than emitting mojibake. Code pages outside the supported
set (including UTF-16 / 1200, which would silently corrupt if treated as bytes) are
rejected.

## Read-modify-write contract

The intended workflow is `Read` → edit `Module.Source` → `Write`. The contract that
makes this safe is **minimal reconstruction**: `Write` rebuilds only the module list
and the module streams; everything else is preserved byte-for-byte.

The `Project` model carries verbatim state captured at read time:

- `ProjectInfoRaw` — the `PROJECTINFORMATION` span of `dir`.
- `ReferencesRaw` — the `PROJECTREFERENCES` span of `dir`.
- `ProjectStreamRaw` — the entire `PROJECT` text stream (including `CMG`/`DPB`/`GC`
  protection fields and the Host Extender Info).
- `RawStreams` — every stream the writer does not own, by full path (see
  [Stream pass-through](#stream-pass-through)).

On write, the new `dir` is `ProjectInfoRaw ++ ReferencesRaw ++ BuildProjectModules(...)`:
the project information and references are spliced back unchanged, and only
`PROJECTMODULES` is rebuilt from the model. The `PROJECT` stream is written verbatim, and
the streams in `RawStreams` are re-emitted unchanged. This avoids re-deriving structures
whose exact bytes must round-trip — references, LibIDs, protection state, designer
storages — and confines authored changes to the module streams, which are the only thing
the editor actually touched.

## Stream pass-through

`Read` and `Write` own two namespaces and treat everything else as opaque:

- the `VBA` storage — `dir`, `_VBA_PROJECT` (a stub), and the module streams are
  regenerated from the model, and its performance caches (`__SRP_*`) are dropped; and
- the `PROJECT` stream — re-emitted verbatim from `ProjectStreamRaw`.

Every other stream is carried through unchanged. `Read` records each stream whose first
path segment is neither `VBA` nor `PROJECT` in `Project.RawStreams` (keyed by full
`/`-separated path), and `Write` re-adds them before the regenerated streams.

The boundary is structural rather than a list of known names: a stream is carried because
the writer does not own its namespace, not because it was recognized. This captures a
UserForm's root-level designer storage (`UserForm1/f`, `o`, `\x01CompObj`, `\x03VBFrame`)
as a whole subtree — without parsing or classifying it — along with `PROJECTwm` and any
other opaque payload. Caches under `VBA/` stay dropped because they fall inside the owned
namespace.

Directory-entry metadata (CLSID, state bits, timestamps, color) is not preserved; the
minimal-profile writer synthesizes it. This is sound for the streams in scope: in
Excel-produced bins the designer storage carries no non-zero CLSID or state — a form's
identity lives in the contents of `\x03VBFrame`/`\x01CompObj`, not the directory entry —
and the storage timestamps and tree coloring are already normalized for the `VBA` storage.

## Write guards

Because the writer trusts its input — `Read`/`Write` is designed for bins produced by
real Excel, not adversarial CFB files — its guards exist to catch *silent
self-corruption*, not to harden against malicious input. There are two:

- **Module-set invariant.** The `PROJECT` stream (preserved verbatim) enumerates the
  module set via its `Module=`/`Class=`/`BaseClass=`/`Document=` lines, while the new
  `PROJECTMODULES` is rebuilt from `Project.Modules`. If those two sets disagree, the
  output bin would be internally inconsistent. `Write` checks that they match exactly
  and rejects adding, removing, or renaming modules. (Editing source within the
  existing set is fully supported; changing the set is not.)

- **Non-ASCII name rejection.** Module and stream names must be ASCII. A `dir` entry
  writes each name twice — once in the MBCS field (raw CODEPAGE bytes) and once in the
  UTF-16LE field — and only ASCII names round-trip losslessly through both. Non-ASCII
  names are rejected rather than written into one field correctly and the other
  corrupted.

`Write` also refuses protected projects, which fall outside the source-only scope.

## Scope and non-goals

- **Editable:** standard, class, document, and UserForm modules. Their source is read,
  exposed, and rewritten. For a UserForm only the code-behind module is editable; its
  designer layout is preserved, not authored.
- **Preserved, not modified:** the UserForm designer storage (`UserForm1/...`) and any
  other non-`VBA` payload, carried verbatim by the [stream pass-through](#stream-pass-through).
  Editing form *layout* — controls, positions, properties inside `\x03VBFrame`/`o` — is
  out of scope.
- **Preserved, not modified:** protected projects. `CMG`/`DPB`/`GC` are read and kept
  verbatim, but the password protection is **not decrypted**, and `Write` refuses to
  emit a protected project. Protection is detected by a heuristic on the `DPB` length
  rather than by decrypting the protection state, since decryption is not performed.

## Verification

Correctness is checked three independent ways:

1. **Byte-exact golden streams.** The `ovba` compressor is compared byte-for-byte
   against compressed module streams captured from an Excel-accepted bin. This pins
   the compression codec exactly, independent of the CFB envelope.
2. **Independent extraction parity.** Source decoded by `Read` is compared against the
   same modules extracted by `olevba` (from the oletools project), an external oracle
   with no shared code. A full read→write→read round-trip over a corpus (standard,
   class, document, references, and MBCS workbooks) confirms the verbatim spans and
   every module's source and type survive unchanged.
3. **Real Excel.** Workbooks written by the library open in Excel, recompile from
   source, and run the edited macros.

## References

Primary sources are the Microsoft Open Specifications. The sections below map each
part of the format to where the code implements or relies on it.

### [MS-OVBA] — Office VBA File Format Structure

Root: <https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc>

| Section          | Topic                                  | Code |
| ---------------- | -------------------------------------- | ---- |
| §2.4.1           | CompressedContainer                    | `ovba/compress.go`, `ovba/decompress.go` |
| §2.4.1.3.2       | Decompressing a chunk                  | `ovba/decompress.go` (`decompressChunk`) |
| §2.4.1.3.7       | Chunk flag ("fits in 4096")            | `ovba/compress.go` (`compressChunk`) |
| §2.4.1.3.19.2    | CopyToken packing                      | `ovba/compress.go` (`packCopyToken`) |
| §2.4.1.3.19.3    | CopyToken bit counts                   | `ovba/compress.go` (`copyTokenHelp`) |
| §2.4.1.3.19.4    | Overlapping copy (RLE)                 | `ovba/compress.go` (`longestMatch`), `ovba/decompress.go` |
| §2.3.4.2.1.10    | PROJECTVERSION record                  | `ovba/parse.go`, `ovba/project.go` (`dirVersionRecord`) |
| §2.3.4.2.2 / .2.5 | REFERENCE / REFERENCEREGISTERED       | `ovba/project.go` (`refRegistered`, `refLibid`) |
| §2.3.4.2.3.2.8   | MODULETYPE (`0x0021` / `0x0022`)       | `ovba/parse.go`, `vbaproject/write.go` |
| §2.4.3           | Project protection / DPB               | `vbaproject/read.go` (preserved, not decrypted) |

### [MS-CFB] — Compound File Binary File Format

Root: <https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b>

| Section | Topic                          | Code |
| ------- | ------------------------------ | ---- |
| §2.2    | FAT / DIFAT sector values      | `cfb/doc.go`, `cfb/reader.go`, `cfb/writer.go` |
| §2.6.1  | DirectoryEntry                 | `cfb/reader.go` (`parseDirEntry`), `cfb/writer.go` (`writeDirEntry`) |
| §2.6.4  | Name ordering                  | `cfb/writer.go` (`nameLess`) |

### [MS-WINERRATA] — errata for [MS-CFB]

<https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-winerrata/c14df6f6-ae1d-45bf-b62f-7b2ed199ba44>

Documents that directory-tree red-black coloring is advisory; the writer builds an
ordered tree and marks all nodes black accordingly.

[MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc
[MS-CFB]: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b
[MS-WINERRATA]: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-winerrata/c14df6f6-ae1d-45bf-b62f-7b2ed199ba44
