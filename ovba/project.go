package ovba

import "encoding/binary"

// VBAProjectStub returns the fixed 7-byte stub for the _VBA_PROJECT stream.
// It marks the project as source-only (no p-code). The value is verified
// against a known-good binary: cc61ffff000300.
func VBAProjectStub() []byte {
	return []byte{0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x03, 0x00}
}

// ProjectStream returns the plaintext PROJECT stream with a fixed set of
// project constants. Line endings are CRLF, and the stream ends with a single
// trailing CRLF to match a known-good binary.
func ProjectStream() []byte {
	const s = "ID=\"{9E394C0B-697E-4AEE-9FA6-446F51FB30DC}\"\r\n" +
		"Document=Sheet1/&H00000000\r\n" +
		"Document=ThisWorkbook/&H00000000\r\n" +
		"Module=Spike\r\n" +
		"Name=\"VBAProject\"\r\n" +
		"HelpContextID=\"0\"\r\n" +
		"CMG=\"6D6F7625A5A1A9A1A9A1A9A1A9\"\r\n" +
		"DPB=\"3D3F26A4400941094109\"\r\n" +
		"GC=\"7E7C652AB2EA03EB03EBFC\"\r\n" +
		"\r\n" +
		"[Host Extender Info]\r\n" +
		"&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n"
	return []byte(s)
}

// recU32 / recU16 build an "id + size + fixed-width integer (LE)" record.
func recU32(id uint16, size uint32, v uint32) []byte {
	p := make([]byte, 4)
	binary.LittleEndian.PutUint32(p, v)
	return sizedRecord(id, p[:size])
}

func recU16(id uint16, size uint32, v uint16) []byte {
	p := make([]byte, 2)
	binary.LittleEndian.PutUint16(p, v)
	return sizedRecord(id, p[:size])
}

// dirVersionRecord is the PROJECTVERSION record. [MS-OVBA] §2.3.4.2.1.10:
// Id(2)=0x0009, Reserved(4)=0x00000004, VersionMajor(4), VersionMinor(2).
// It has a special layout with no size field. The values come from the dir.plain of an accepted bin.
func dirVersionRecord() []byte {
	b := make([]byte, 12)
	binary.LittleEndian.PutUint16(b[0:], 0x0009)
	binary.LittleEndian.PutUint32(b[2:], 0x00000004)
	binary.LittleEndian.PutUint32(b[6:], 0x65BE0257)
	binary.LittleEndian.PutUint16(b[10:], 0x0011)
	return b
}

// refLibid builds the payload of a REFERENCEREGISTERED. [MS-OVBA] §2.3.4.2.2.5:
// SizeOfLibid(4) + Libid(ASCII) + Reserved1(4)=0 + Reserved2(2)=0.
func refLibid(libid string) []byte {
	out := make([]byte, 4+len(libid)+6)
	binary.LittleEndian.PutUint32(out[0:], uint32(len(libid)))
	copy(out[4:], libid)
	// the trailing 6 bytes remain zero for Reserved1/Reserved2
	return out
}

// refRegistered builds a REFERENCENAME + REFERENCEREGISTERED pair. [MS-OVBA] §2.3.4.2.2.
func refRegistered(name, libid string) []byte {
	var b []byte
	b = append(b, sizedRecord(0x0016, []byte(name))...)  // REFERENCENAME
	b = append(b, sizedRecord(0x003E, utf16le(name))...) // ...Unicode
	b = append(b, sizedRecord(0x000D, refLibid(libid))...)
	return b
}

// modRecord builds one module entry within dir (MODULENAME..MODULE terminator).
// streamName is the module stream name (it can differ from Name in real bins).
// modType is 0x0021 (procedural/std) or 0x0022 (document/class). [MS-OVBA] §2.3.4.2.3.2.8.
func modRecord(name, streamName string, modType uint16) []byte {
	nu := utf16le(name)
	su := utf16le(streamName)
	var b []byte
	b = append(b, sizedRecord(0x0019, []byte(name))...)       // MODULENAME
	b = append(b, sizedRecord(0x0047, nu)...)                 // MODULENAMEUNICODE
	b = append(b, sizedRecord(0x001A, []byte(streamName))...) // MODULESTREAMNAME
	b = append(b, sizedRecord(0x0032, su)...)                 // ...Unicode
	b = append(b, sizedRecord(0x001C, nil)...)                // MODULEDOCSTRING
	b = append(b, sizedRecord(0x0048, nil)...)                // ...Unicode
	b = append(b, recU32(0x0031, 4, 0x00000000)...)           // MODULEOFFSET = 0 (source-only)
	b = append(b, recU32(0x001E, 4, 0x00000000)...)           // MODULEHELPCONTEXT
	b = append(b, recU16(0x002C, 2, 0xFFFF)...)               // MODULECOOKIE
	b = append(b, sizedRecord(modType, nil)...)               // MODULETYPE
	b = append(b, sizedRecord(0x002B, nil)...)                // MODULE terminator
	return b
}

// ModuleSpec specifies one module used to build the PROJECTMODULES section.
type ModuleSpec struct {
	Name       string
	StreamName string
	TypeID     uint16 // 0x0021=std / 0x0022=class or document
}

// BuildProjectModules builds the PROJECTMODULES section of the dir stream:
// the MODULES count, PROJECTCOOKIE, one record per module (MODULEOFFSET=0),
// and the terminator.
func BuildProjectModules(specs []ModuleSpec) []byte {
	b := recU16(0x000F, 2, uint16(len(specs)))  // MODULES count
	b = append(b, recU16(0x0013, 2, 0xFFFF)...) // PROJECTCOOKIE
	for _, m := range specs {
		b = append(b, modRecord(m.Name, m.StreamName, m.TypeID)...)
	}
	b = append(b, sizedRecord(0x0010, nil)...) // terminator
	return b
}

// DirStreamPlain builds the uncompressed bytes of the dir stream with a fixed
// layout: the VBAProject information, the stdole/Office references, and the
// Sheet1/ThisWorkbook/Spike modules.
func DirStreamPlain() []byte {
	var b []byte
	// --- PROJECTINFORMATION ---
	b = append(b, recU32(0x0001, 4, 0x00000003)...)             // SYSKIND = 3 (Win64)
	b = append(b, recU32(0x0002, 4, 0x00000409)...)             // LCID
	b = append(b, recU32(0x0014, 4, 0x00000409)...)             // LCIDINVOKE
	b = append(b, recU16(0x0003, 2, 0x04E4)...)                 // CODEPAGE = 1252
	b = append(b, sizedRecord(0x0004, []byte("VBAProject"))...) // NAME
	b = append(b, sizedRecord(0x0005, nil)...)                  // DOCSTRING
	b = append(b, sizedRecord(0x0040, nil)...)                  // ...Unicode
	b = append(b, sizedRecord(0x0006, nil)...)                  // HELPFILE1
	b = append(b, sizedRecord(0x003D, nil)...)                  // HELPFILE2
	b = append(b, recU32(0x0007, 4, 0x00000000)...)             // HELPCONTEXT
	b = append(b, recU32(0x0008, 4, 0x00000000)...)             // LIBFLAGS
	b = append(b, dirVersionRecord()...)                        // PROJECTVERSION
	b = append(b, sizedRecord(0x000C, nil)...)                  // CONSTANTS
	b = append(b, sizedRecord(0x003C, nil)...)                  // ...Unicode
	// --- PROJECTREFERENCES ---
	b = append(b, refRegistered("stdole",
		"*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation")...)
	b = append(b, refRegistered("Office",
		"*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL#Microsoft Office 16.0 Object Library")...)
	// --- PROJECTMODULES ---
	b = append(b, BuildProjectModules([]ModuleSpec{
		{Name: "Sheet1", StreamName: "Sheet1", TypeID: 0x0022},
		{Name: "ThisWorkbook", StreamName: "ThisWorkbook", TypeID: 0x0022},
		{Name: "Spike", StreamName: "Spike", TypeID: 0x0021},
	})...)
	return b
}
