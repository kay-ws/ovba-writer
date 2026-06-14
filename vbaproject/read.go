package vbaproject

import (
	"fmt"

	"github.com/kay-ws/ovba-writer/cfb"
	"github.com/kay-ws/ovba-writer/ovba"
)

// Read parses vbaProject.bin and decomposes it into a *Project model.
func Read(data []byte) (*Project, error) {
	c, err := cfb.Open(data)
	if err != nil {
		return nil, fmt.Errorf("vbaproject: cfb open: %w", err)
	}
	projRaw, ok := c.Stream("PROJECT")
	if !ok {
		return nil, fmt.Errorf("vbaproject: missing PROJECT stream")
	}
	pt := ovba.ParseProjectText(projRaw)

	dirComp, ok := c.Stream("VBA/dir")
	if !ok {
		return nil, fmt.Errorf("vbaproject: missing VBA/dir stream")
	}
	dirPlain, err := ovba.Decompress(dirComp)
	if err != nil {
		return nil, fmt.Errorf("vbaproject: dir decompress: %w", err)
	}
	di := ovba.ParseDir(dirPlain)

	p := &Project{
		Props: ProjectProps{
			ProjectID: pt.ID, Name: pt.Name,
			SysKind: di.SysKind, LCID: di.LCID, CodePage: di.CodePage,
		},
		Protection: Protection{
			CMG: pt.CMG, DPB: pt.DPB, GC: pt.GC,
			IsProtected: isProtected(pt.DPB),
		},
	}
	for _, name := range di.RefNames {
		p.References = append(p.References, Reference{Name: name})
	}
	p.ReferencesRaw = di.RefsRaw
	p.ProjectInfoRaw = di.ProjectInfoRaw
	p.ProjectStreamRaw = projRaw
	for _, dm := range di.Modules {
		stream := dm.StreamName
		if stream == "" {
			stream = dm.Name
		}
		raw, ok := c.Stream("VBA/" + stream)
		if !ok {
			return nil, fmt.Errorf("vbaproject: missing module stream VBA/%s", stream)
		}
		if int(dm.Offset) > len(raw) {
			return nil, fmt.Errorf("vbaproject: %s offset %d > stream %d", dm.Name, dm.Offset, len(raw))
		}
		plain, err := ovba.Decompress(raw[dm.Offset:])
		if err != nil {
			return nil, fmt.Errorf("vbaproject: %s decompress: %w", dm.Name, err)
		}
		src, err := decodeMBCS(plain, di.CodePage)
		if err != nil {
			return nil, fmt.Errorf("vbaproject: %s decode: %w", dm.Name, err)
		}
		p.Modules = append(p.Modules, Module{
			Name:       dm.Name,
			StreamName: stream,
			Type:       classify(dm.Name, pt.Kinds[dm.Name], c),
			Source:     src,
		})
	}
	return p, nil
}

// classify decides the module type primarily from the PROJECT key, secondarily from the presence of a CFB sub-storage.
func classify(name, projectKey string, c *cfb.Container) ModuleType {
	switch projectKey {
	case "Module":
		return ModuleStd
	case "Class":
		return ModuleClass
	case "Document":
		return ModuleDocument
	case "BaseClass":
		return ModuleForm
	}
	// Fallback when absent from PROJECT: if a same-named sub-storage exists, it is a form.
	if _, ok := c.Stream(name + "/\x03VBFrame"); ok {
		return ModuleForm
	}
	return ModuleClass
}

// isProtected determines whether a project is protected. The spec-compliant check decrypts
// CMG/ProjectProtectionState ([MS-OVBA] §2.4.3), but this library does not decrypt the DPB, so it uses a
// heuristic: an unprotected DPB is a short obfuscated value (~16-20 chars measured in the corpus), while a
// protected one contains a key hash and is longer (over 60 chars in the protected sample). The threshold 28
// is the valley between them. Protected projects are treated as an edge case in v1 (see DESIGN.md).
func isProtected(dpb string) bool { return len(dpb) > 28 }
