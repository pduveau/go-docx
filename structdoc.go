/*
   Copyright (c) 2020 gingfrederik
   Copyright (c) 2021 Gonzalo Fernandez-Victorio
   Copyright (c) 2021 Basement Crowd Ltd (https://www.basementcrowd.com)
   Copyright (c) 2023 Fumiama Minamoto (源文雨)

   This program is free software: you can redistribute it and/or modify
   it under the terms of the GNU Affero General Public License as published
   by the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   This program is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU Affero General Public License for more details.

   You should have received a copy of the GNU Affero General Public License
   along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

package docx

import (
	"encoding/xml"
	"io"
	"reflect"
	"regexp"
	"strings"
)

//nolint:revive,stylecheck
const (
	XMLNS_WPC      = `http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas`
	XMLNS_CX       = "http://schemas.microsoft.com/office/drawing/2014/chartex"
	XMLNS_CX1      = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
	XMLNS_CX2      = "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
	XMLNS_CX3      = "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"
	XMLNS_CX4      = "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"
	XMLNS_CX5      = "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"
	XMLNS_CX6      = "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"
	XMLNS_CX7      = "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"
	XMLNS_CX8      = "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"
	XMLNS_MC       = `http://schemas.openxmlformats.org/markup-compatibility/2006`
	XMLNS_AINK     = "http://schemas.microsoft.com/office/drawing/2016/ink"
	XMLNS_AM3D     = "http://schemas.microsoft.com/office/drawing/2017/model3d"
	XMLNS_O        = `urn:schemas-microsoft-com:office:office`
	XMLNS_OEL      = "http://schemas.microsoft.com/office/2019/extlst"
	XMLNS_R        = `http://schemas.openxmlformats.org/officeDocument/2006/relationships`
	XMLNS_M        = "http://schemas.openxmlformats.org/officeDocument/2006/math"
	XMLNS_V        = `urn:schemas-microsoft-com:vml`
	XMLNS_WP14     = `http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing`
	XMLNS_WP       = `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing`
	XMLNS_W10      = "urn:schemas-microsoft-com:office:word"
	XMLNS_W        = `http://schemas.openxmlformats.org/wordprocessingml/2006/main`
	XMLNS_W14      = "http://schemas.microsoft.com/office/word/2010/wordml"
	XMLNS_W15      = "http://schemas.microsoft.com/office/word/2012/wordml"
	XMLNS_W16CEX   = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
	XMLNS_W16CID   = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
	XMLNS_W16      = "http://schemas.microsoft.com/office/word/2018/wordml"
	XMLNS_W16DU    = "http://schemas.microsoft.com/office/word/2023/wordml/word16du"
	XMLNS_W16SDTDH = "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
	XMLNS_W16SDTFL = "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock"
	XMLNS_W16SE    = "http://schemas.microsoft.com/office/word/2015/wordml/symex"
	XMLNS_WPG      = `http://schemas.microsoft.com/office/word/2010/wordprocessingGroup`
	XMLNS_WPI      = "http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
	XMLNS_WNE      = "http://schemas.microsoft.com/office/word/2006/wordml"
	XMLNS_WPS      = `http://schemas.microsoft.com/office/word/2010/wordprocessingShape`
	MC_IGNORABLE   = "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14"

	XMLNS_PICTURE = `http://schemas.openxmlformats.org/drawingml/2006/picture`
)

// Document <w:document>
type Document struct {
	XMLName     xml.Name `xml:"w:document"`
	XMLwpc      string   `xml:"xmlns:wpc,attr,omitempty"`
	XMLcx       string   `xml:"xmlns:cx,attr,omitempty"`
	XMLcx1      string   `xml:"xmlns:cx1,attr,omitempty"`
	XMLcx2      string   `xml:"xmlns:cx2,attr,omitempty"`
	XMLcx3      string   `xml:"xmlns:cx3,attr,omitempty"`
	XMLcx4      string   `xml:"xmlns:cx4,attr,omitempty"`
	XMLcx5      string   `xml:"xmlns:cx5,attr,omitempty"`
	XMLcx6      string   `xml:"xmlns:cx6,attr,omitempty"`
	XMLcx7      string   `xml:"xmlns:cx7,attr,omitempty"`
	XMLcx8      string   `xml:"xmlns:cx8,attr,omitempty"`
	XMLmc       string   `xml:"xmlns:mc,attr,omitempty"`
	XMLaink     string   `xml:"xmlns:aink,attr,omitempty"`
	XMLam3d     string   `xml:"xmlns:am3d,attr,omitempty"`
	XMLo        string   `xml:"xmlns:o,attr,omitempty"`
	XMLoel      string   `xml:"xmlns:oel,attr,omitempty"`
	XMLr        string   `xml:"xmlns:r,attr,omitempty"`
	XMLm        string   `xml:"xmlns:m,attr,omitempty"`
	XMLv        string   `xml:"xmlns:v,attr,omitempty"`
	XMLwp14     string   `xml:"xmlns:wp14,attr,omitempty"`
	XMLwp       string   `xml:"xmlns:wp,attr,omitempty"`
	XMLw10      string   `xml:"xmlns:w10,attr,omitempty"`
	XMLw        string   `xml:"xmlns:w,attr,omitempty"`
	XMLw14      string   `xml:"xmlns:w14,attr,omitempty"`
	XMLw15      string   `xml:"xmlns:w15,attr,omitempty"`
	XMLw16cex   string   `xml:"xmlns:w16cex,attr,omitempty"`
	XMLw16cid   string   `xml:"xmlns:w16cid,attr,omitempty"`
	XMLw16      string   `xml:"xmlns:w16,attr,omitempty"`
	XMLw16du    string   `xml:"xmlns:w16du,attr,omitempty"`
	XMLw16sdtdh string   `xml:"xmlns:w16sdtdh,attr,omitempty"`
	XMLw16sdtfl string   `xml:"xmlns:w16sdtfl,attr,omitempty"`
	XMLw16se    string   `xml:"xmlns:w16se,attr,omitempty"`
	XMLwpg      string   `xml:"xmlns:wpg,attr,omitempty"`
	XMLwpi      string   `xml:"xmlns:wpi,attr,omitempty"`
	XMLwne      string   `xml:"xmlns:wne,attr,omitempty"`
	XMLwps      string   `xml:"xmlns:wps,attr,omitempty"`
	MCIgnorable string   `xml:"mc:Ignorable,attr,omitempty"`

	Body Body `xml:"w:body"`
}

// UnmarshalXML ...
func (doc *Document) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "wpc":
			doc.XMLwpc = attr.Value
		case "cx":
			doc.XMLcx = attr.Value
		case "cx1":
			doc.XMLcx1 = attr.Value
		case "cx2":
			doc.XMLcx2 = attr.Value
		case "cx3":
			doc.XMLcx3 = attr.Value
		case "cx4":
			doc.XMLcx4 = attr.Value
		case "cx5":
			doc.XMLcx5 = attr.Value
		case "cx6":
			doc.XMLcx6 = attr.Value
		case "cx7":
			doc.XMLcx7 = attr.Value
		case "cx8":
			doc.XMLcx8 = attr.Value
		case "mc":
			doc.XMLmc = attr.Value
		case "aink":
			doc.XMLaink = attr.Value
		case "am3d":
			doc.XMLam3d = attr.Value
		case "o":
			doc.XMLo = attr.Value
		case "oel":
			doc.XMLoel = attr.Value
		case "r":
			doc.XMLr = attr.Value
		case "m":
			doc.XMLm = attr.Value
		case "v":
			doc.XMLv = attr.Value
		case "wp14":
			doc.XMLwp14 = attr.Value
		case "wp":
			doc.XMLwp = attr.Value
		case "w10":
			doc.XMLw10 = attr.Value
		case "w":
			doc.XMLw = attr.Value
		case "w14":
			doc.XMLw14 = attr.Value
		case "w15":
			doc.XMLw15 = attr.Value
		case "w16cex":
			doc.XMLw16cex = attr.Value
		case "w16cid":
			doc.XMLw16cid = attr.Value
		case "w16":
			doc.XMLw16 = attr.Value
		case "w16du":
			doc.XMLw16du = attr.Value
		case "w16sdtdh":
			doc.XMLw16sdtdh = attr.Value
		case "w16sdtfl":
			doc.XMLw16sdtfl = attr.Value
		case "w16se":
			doc.XMLw16se = attr.Value
		case "wpg":
			doc.XMLwpg = attr.Value
		case "wpi":
			doc.XMLwpi = attr.Value
		case "wne":
			doc.XMLwne = attr.Value
		case "wps":
			doc.XMLwps = attr.Value
		case "Ignorable":
			doc.MCIgnorable = attr.Value
		default:
		}
	}
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			if tt.Name.Local == "body" {
				err = d.DecodeElement(&doc.Body, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				continue
			}
			err = d.Skip() // skip unsupported tags
			if err != nil {
				return err
			}
		}
	}
	return nil
}

func getAtt(atts []xml.Attr, name string) string {
	for _, at := range atts {
		if at.Name.Local == name {
			return at.Value
		}
	}
	return ""
}

// Body <w:body>
type Body struct {
	Items []interface{}

	file *Docx
}

// UnmarshalXML ...
func (b *Body) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "p":
				var value Paragraph
				value.file = b.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				b.Items = append(b.Items, &value)
			case "tbl":
				var value Table
				value.file = b.file
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				b.Items = append(b.Items, &value)
			case "sdt":
				var value Sdt
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				b.Items = append(b.Items, &value)
			case "sectPr":
				var value SectionProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				b.Items = append(b.Items, &value)
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
			}
		}
	}
	return nil
}

// KeepElements keep named elems amd removes others
//
// names: *docx.Paragraph *docx.Table
func (b *Body) KeepElements(name ...string) {
	items := make([]interface{}, 0, len(b.Items))
	namemap := make(map[string]struct{}, len(name)*2)
	for _, n := range name {
		namemap[n] = struct{}{}
	}
	for _, item := range b.Items {
		_, ok := namemap[reflect.ValueOf(item).Type().String()]
		if ok {
			items = append(items, item)
		}
	}
	b.Items = items
}

// DropDrawingOf drops all matched drawing in body
// name: Canvas, Shape, Group, ShapeAndCanvas, ShapeAndCanvasAndGroup, NilPicture
func (b *Body) DropDrawingOf(name string) {
	for _, item := range b.Items {
		switch o := item.(type) {
		case *Paragraph:
			f := reflect.ValueOf(o).MethodByName("Drop" + name)
			if !f.IsValid() {
				continue
			}
			_ = f.Call(nil)
		case *Table:
			for _, tr := range o.Rows {
				for _, tc := range tr.Cells {
					for _, p := range tc.Paragraphs {
						f := reflect.ValueOf(p).MethodByName("Drop" + name)
						if !f.IsValid() {
							continue
						}
						_ = f.Call(nil)
					}
				}
			}
		}
	}
}

type _justification string

const ( //	w:jc possible values：
	JUSTIFICATION_RIGHT      _justification = "start"
	JUSTIFICATION_CENTER     _justification = "center"
	JUSTIFICATION_LEFT       _justification = "end"
	JUSTIFICATION_JUSTIFIED  _justification = "both"       // justify
	JUSTIFICATION_DISTRIBUTE _justification = "distribute" // disperse Alignment
)

// ParagraphSplitRule check whether the paragraph is a separator or not
type ParagraphSplitRule func(*Paragraph) bool

// SplitDocxByPlainTextRegex matches p.String()
func SplitDocxByPlainTextRegex(re *regexp.Regexp) ParagraphSplitRule {
	return func(p *Paragraph) bool {
		return re.MatchString(p.String())
	}
}

// SplitByParagraph splits a doc to many docs by using a matched paragraph
// as the separator.
//
// The separator will be placed to the first doc item
func (f *Docx) SplitByParagraph(separator ParagraphSplitRule) (docs []*Docx) {
	items := f.Document.Body.Items
newdoclop:
	for len(items) > 0 {
		ndoc := new(Docx)

		// migrate base data
		ndoc.mediaNameIdx = make(map[string]int, 64)
		ndoc.slowIDs = make(map[string]uintptr, 64)
		ndoc.template = f.template
		ndoc.tmplfs = f.tmplfs
		ndoc.tmpfslst = f.tmpfslst

		ndoc.Document.XMLw = XMLNS_W
		ndoc.Document.XMLr = XMLNS_R
		ndoc.Document.XMLwp = XMLNS_WP
		// ndoc.Document.XMLMC = XMLNS_MC
		// ndoc.Document.XMLO = XMLNS_O
		// ndoc.Document.XMLV = XMLNS_V
		ndoc.Document.XMLwps = XMLNS_WPS
		ndoc.Document.XMLwpc = XMLNS_WPC
		ndoc.Document.XMLwpg = XMLNS_WPG
		// ndoc.Document.XMLWP14 = XMLNS_WP14
		ndoc.Document.XMLName.Space = XMLNS_W
		ndoc.Document.XMLName.Local = "document"
		ndoc.Document.Body.file = ndoc

		ndoc.docRelation = Relationships{
			Xmlns: XMLNS_REL,
			Relationship: []Relationship{
				{
					ID:     "rId1",
					Type:   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`,
					Target: "styles.xml",
				},
				{
					ID:     "rId2",
					Type:   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme`,
					Target: "theme/theme1.xml",
				},
				{
					ID:     "rId3",
					Type:   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable`,
					Target: "fontTable.xml",
				},
			},
		}

		ndoc.rID = 3

		for i, item := range items {
			switch o := item.(type) {
			case *Paragraph:
				if separator(o) && len(ndoc.Document.Body.Items) > 0 {
					items = items[i:]
					docs = append(docs, ndoc)
					continue newdoclop
				}
				np := o.copymedia(ndoc)
				ndoc.Document.Body.Items = append(ndoc.Document.Body.Items, &np)
			case *Table:
				nt := o.copymedia(ndoc)
				ndoc.Document.Body.Items = append(ndoc.Document.Body.Items, &nt)
			default:
				ndoc.Document.Body.Items = append(ndoc.Document.Body.Items, o)
			}
		}

		if len(ndoc.Document.Body.Items) > 0 {
			docs = append(docs, ndoc)
		}
		break
	}
	return
}

func (r *Run) copymedia(to *Docx) *Run {
	nr := *r
	nr.Children = make([]interface{}, 0, len(r.Children))
	nr.file = to
	for _, rc := range r.Children {
		if d, ok := rc.(*Drawing); ok {
			nr.Children = append(nr.Children, d.copymedia(to))
			continue
		}
		nr.Children = append(nr.Children, rc)
	}
	return &nr
}

func (p *Paragraph) copymedia(to *Docx) (np Paragraph) {
	np = *p
	np.Children = make([]interface{}, 0, len(p.Children))
	np.file = to
	for _, pc := range p.Children {
		if r, ok := pc.(*Run); ok {
			np.Children = append(np.Children, r.copymedia(to))
			continue
		}
		if h, ok := pc.(*Hyperlink); ok {
			tgt, err := p.file.ReferTarget(h.ID)
			if err != nil {
				continue
			}
			rid := to.addLinkRelation(tgt)
			nh := &Hyperlink{
				ID:   rid,
				Run:  make([]*Run, 0),
				file: to,
			}
			for _, ru := range h.Run {
				nh.Run = append(nh.Run, ru.copymedia(to))
			}
			np.Children = append(np.Children, nh)
			continue
		}
		np.Children = append(np.Children, pc)
	}
	return
}

func (t *Table) copymedia(to *Docx) (nt Table) {
	nt = *t
	nt.Rows = make([]*WTableRow, 0, len(t.Rows))
	nt.file = to
	for _, tr := range t.Rows {
		ntr := *tr
		ntr.Cells = make([]*WTableCell, 0, len(tr.Cells))
		ntr.file = to
		for _, tc := range tr.Cells {
			ntc := *tc
			ntc.Paragraphs = make([]*Paragraph, 0, len(tc.Paragraphs))
			ntc.file = to
			for _, p := range tc.Paragraphs {
				np := p.copymedia(to)
				ntc.Paragraphs = append(ntc.Paragraphs, &np)
			}
			ntr.Cells = append(ntr.Cells, &ntc)
		}
		nt.Rows = append(nt.Rows, &ntr)
	}
	return
}

// AppendFile appends all contents in af to f
func (f *Docx) AppendFile(af *Docx) {
	for _, item := range af.Document.Body.Items {
		switch o := item.(type) {
		case *Paragraph:
			np := o.copymedia(f)
			f.Document.Body.Items = append(f.Document.Body.Items, &np)
		case *Table:
			nt := o.copymedia(f)
			f.Document.Body.Items = append(f.Document.Body.Items, &nt)
		default:
			f.Document.Body.Items = append(f.Document.Body.Items, o)
		}
	}
}
