package docx

import (
	"encoding/xml"
	"io"
	"strings"
)

type SdtID struct {
	XMLName xml.Name `xml:"w:id,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

type SdtDocPartGallery struct {
	XMLName xml.Name `xml:"w:docPartGallery,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

type SdtDocPartUnique struct {
	XMLName xml.Name `xml:"w:docPartUnique,omitempty"`
}

type SdtDocPartObj struct {
	XMLName        xml.Name `xml:"w:docPartObj,omitempty"`
	DocPartGallery *SdtDocPartGallery
	DocPartUnique  *SdtDocPartUnique
}

// UnmarshalXML ...
func (p *SdtDocPartObj) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "docPartGallery":
				p.DocPartGallery = &SdtDocPartGallery{Val: getAtt(tt.Attr, "val")}
			case "docPartUnique":
				p.DocPartUnique = &SdtDocPartUnique{}
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

type SdtProperties struct {
	XMLName    xml.Name `xml:"w:sdtPr,omitempty"`
	RPr        *RunProperties
	ID         *SdtID
	DocPartObj *SdtDocPartObj
}

func (p *SdtProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "id":
				p.ID = &SdtID{Val: getAtt(tt.Attr, "val")}
			case "docPartObj":
				var value SdtDocPartObj
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.DocPartObj = &value
			case "rPr":
				var value RunProperties
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.RPr = &value
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
				continue
			}
		}
	}
	return nil
}

type SdtEndProperties struct {
	XMLName xml.Name `xml:"w:sdtEndPr,omitempty"`
	RPr     RunProperties
}

// UnmarshalXML ...
func (p *SdtEndProperties) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if tt, ok := t.(xml.StartElement); ok {
			if tt.Name.Local == "rPr" {
				err = d.DecodeElement(&p.RPr, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			} else {
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
			}
		}
	}
	return nil
}

type SdtContent struct {
	XMLName xml.Name `xml:"w:sdtContent,omitempty"`
	P       []*Paragraph
}

// UnmarshalXML ...
func (c *SdtContent) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var value *Paragraph

	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if tt, ok := t.(xml.StartElement); ok {
			if tt.Name.Local == "p" && value == nil {
				value = &Paragraph{}
				err = d.DecodeElement(value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
			} else {
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
			}
		}
	}
	c.P = []*Paragraph{
		value,
		&Paragraph{
			Children: []interface{}{
				&Run{
					FldChar: &RunFldChar{FldCharType: "begin"},
				},
				&Run{
					InstrText: &RunInstrText{Space: "preserve", Text: ` TOC \o "1-3" \h \z \u `},
				},
				&Run{
					FldChar: &RunFldChar{FldCharType: "separate"},
				},
				&Run{
					RunProperties: &RunProperties{
						Bold: &Bold{},
						BCs:  &struct{}{},
					},
					Children: []interface{}{
						&Text{
							Text: "To be updated",
						},
					},
				},
				&Run{
					RunProperties: &RunProperties{
						Bold: &Bold{},
						BCs:  &struct{}{},
					},
					FldChar: &RunFldChar{FldCharType: "end"},
				},
			},
		},
	}
	return nil
}

type Sdt struct {
	XMLName    xml.Name `xml:"w:sdt,omitempty"`
	SdtPr      *SdtProperties
	SdtEndPr   *SdtEndProperties
	SdtContent *SdtContent
}

// UnmarshalXML ...
func (p *Sdt) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {

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
			case "sdtPr":
				var pr SdtProperties
				err = d.DecodeElement(&pr, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.SdtPr = &pr
			case "sdtEndPr":
				var pr SdtEndProperties
				err = d.DecodeElement(&pr, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.SdtEndPr = &pr
			case "sdtContent":
				var content SdtContent
				err = d.DecodeElement(&content, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				p.SdtContent = &content
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

/* type SdtFldSimple struct {
	XMLName xml.Name `xml:"w:fldSimple,omitempty"`
	Instr   string   `xml:"instr,attr,omitempty"`
	R       Run
}

// UnmarshalXML ...
func (p *SdtFldSimple) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "instr" {
			p.Instr = attr.Value
		}
	}
	return nil
}
*/
