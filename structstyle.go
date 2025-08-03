/*
   Copyright (c) 2025 Philippe Duveau

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
	"strings"
)

func getVal(atts []xml.Attr) *StyleVal {
	for _, at := range atts {
		if at.Name.Local == "val" {
			return &StyleVal{Val: at.Value}
		}
	}
	return &StyleVal{}
}

func getW14Val(atts []xml.Attr) *StyleW14Val {
	for _, at := range atts {
		if at.Name.Local == "val" {
			return &StyleW14Val{Val: at.Value}
		}
	}
	return &StyleW14Val{}
}

func trapExpectedError(err error) bool {
	return err != nil && !strings.HasPrefix(err.Error(), "expected")
}

var emptyStruct = &StyleVal{}

// types that has attr Val only and require getAtt
type StyleVal struct {
	Val string `xml:"w:val,attr,omitempty"`
}

type StyleW14Val struct {
	Val string `xml:"w14:val,attr,omitempty"`
}

// composed high level structs

// types that require UnmarshalXML methods
type StyleLsdException struct {
	XMLName        xml.Name `xml:"w:lsdException,omitempty"`
	Name           string   `xml:"w:name,attr,omitempty"`
	SemiHidden     string   `xml:"w:semiHidden,attr,omitempty"`
	UiPriority     string   `xml:"w:uiPriority,attr,omitempty"`
	UnhideWhenUsed string   `xml:"w:unhideWhenUsed,attr,omitempty"`
	QFormat        string   `xml:"w:qFormat,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleLsdException) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var err error
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "name":
			s.Name = attr.Value
		case "uiPriority":
			s.UiPriority = attr.Value
		case "qFormat":
			s.QFormat = attr.Value
		case "semiHidden":
			s.SemiHidden = attr.Value
		case "unhideWhenUsed":
			s.UnhideWhenUsed = attr.Value
		default:
		}
	}
	_, err = d.Token()
	return err
}

type StyleLatentStyles struct {
	XMLName           xml.Name `xml:"w:latentStyles,omitempty"`
	DefLockedState    string   `xml:"w:defLockedState,attr,omitempty"`
	DefUIPriority     string   `xml:"w:defUIPriority,attr,omitempty"`
	DefSemiHidden     string   `xml:"w:defSemiHidden,attr,omitempty"`
	DefUnhideWhenUsed string   `xml:"w:defUnhideWhenUsed,attr,omitempty"`
	DefQFormat        string   `xml:"w:defQFormat,attr,omitempty"`
	Count             string   `xml:"w:count,attr,omitempty"`
	LsdException      []*StyleLsdException
}

// UnmarshalXML...
func (s *StyleLatentStyles) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "defLockedState":
			s.DefLockedState = attr.Value
		case "defUIPriority":
			s.DefUIPriority = attr.Value
		case "defQFormat":
			s.DefQFormat = attr.Value
		case "defSemiHidden":
			s.DefSemiHidden = attr.Value
		case "defUnhideWhenUsed":
			s.DefUnhideWhenUsed = attr.Value
		case "count":
			s.Count = attr.Value
		default:
		}
	}
	for {
		token, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if tt, ok := token.(xml.StartElement); ok {
			if tt.Name.Local == "lsdException" {
				var value StyleLsdException
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.LsdException = append(s.LsdException, &value)
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

type StyleColor struct {
	XMLName    xml.Name `xml:"w:color,omitempty"`
	Val        string   `xml:"w:val,attr,omitempty"`
	Color      string   `xml:"w:color,attr,omitempty"`
	ThemeColor string   `xml:"w:themeColor,attr,omitempty"`
	ThemeShade string   `xml:"w:themeShade,attr,omitempty"`
	ThemeTint  string   `xml:"w:themeTint,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleColor) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "val":
			s.Val = attr.Value
		case "color":
			s.Color = attr.Value
		case "themeColor":
			s.ThemeColor = attr.Value
		case "themeShade":
			s.ThemeShade = attr.Value
		case "themeTint":
			s.ThemeTint = attr.Value
		default:
		}
	}
	_, err := d.Token()
	return err
}

type StyleBorder struct {
	Val        string `xml:"w:val,attr,omitempty"`
	Sz         string `xml:"w:sz,attr,omitempty"`
	Space      string `xml:"w:space,attr,omitempty"`
	Color      string `xml:"w:color,attr,omitempty"`
	ThemeColor string `xml:"w:themeColor,attr,omitempty"`
	ThemeShade string `xml:"w:themeShade,attr,omitempty"`
	ThemeTint  string `xml:"w:themeTint,attr,omitempty"`
	W          string `xml:"w:w,attr,omitempty"`
	Type       string `xml:"w:type,attr,omitempty"`
	Shadow     string `xml:"w:shadow,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleBorder) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "val":
			s.Val = attr.Value
		case "sz":
			s.Sz = attr.Value
		case "space":
			s.Space = attr.Value
		case "w":
			s.W = attr.Value
		case "type":
			s.Type = attr.Value
		case "shadow":
			s.Shadow = attr.Value
		case "color":
			s.Color = attr.Value
		case "themeColor":
			s.ThemeColor = attr.Value
		case "themeShade":
			s.ThemeShade = attr.Value
		case "themeTint":
			s.ThemeTint = attr.Value
		default:
		}
	}
	_, err := d.Token()
	return err
}

type StyleCellMar struct {
	W    string `xml:"w:w,attr,omitempty"`
	Type string `xml:"w:type,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleCellMar) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "w":
			s.W = attr.Value
		case "type":
			s.Type = attr.Value
		default:
		}
	}
	_, err := d.Token()
	return err
}

type StyleInd struct {
	XMLName   xml.Name `xml:"w:ind,omitempty"`
	Left      string   `xml:"w:left,attr,omitempty"`
	Right     string   `xml:"w:right,attr,omitempty"`
	Hanging   string   `xml:"w:hanging,attr,omitempty"`
	FirstLine string   `xml:"w:firstLine,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleInd) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "left":
			s.Left = attr.Value
		case "right":
			s.Right = attr.Value
		case "hanging":
			s.Hanging = attr.Value
		case "firstLine":
			s.FirstLine = attr.Value
		default:
		}
	}
	_, err = d.Token()
	return
}

type StyleLang struct {
	XMLName  xml.Name `xml:"w:lang,omitempty"`
	Val      string   `xml:"w:val,attr,omitempty"`
	EastAsia string   `xml:"w:eastAsia,attr,omitempty"`
	Bidi     string   `xml:"w:bidi,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleLang) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "val":
			s.Val = attr.Value
		case "eastAsia":
			s.EastAsia = attr.Value
		case "bidi":
			s.Bidi = attr.Value
		default:
		}
	}
	_, err := d.Token()
	return err
}

type StyleNumPr struct {
	XMLName xml.Name  `xml:"w:numPr,omitempty"`
	Ilvl    *StyleVal `xml:"w:ilvl,omitempty"`
	NumId   *StyleVal `xml:"w:numId,omitempty"`
}

// UnmarshalXML
func (s *StyleNumPr) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "ilvl":
				s.Ilvl = getVal(tt.Attr)
			case "numId":
				s.NumId = getVal(tt.Attr)
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

type StyleRFonts struct {
	XMLName       xml.Name `xml:"w:rFonts,omitempty"`
	Ascii         string   `xml:"w:ascii,attr,omitempty"`
	AsciiTheme    string   `xml:"w:asciiTheme,attr,omitempty"`
	EastAsia      string   `xml:"w:eastAsia,attr,omitempty"`
	EastAsiaTheme string   `xml:"w:eastAsiaTheme,attr,omitempty"`
	HAnsi         string   `xml:"w:hAnsi,attr,omitempty"`
	HAnsiTheme    string   `xml:"w:hAnsiTheme,attr,omitempty"`
	Cs            string   `xml:"w:cs,attr,omitempty"`
	Cstheme       string   `xml:"w:cstheme,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleRFonts) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "asciiTheme":
			s.AsciiTheme = attr.Value
		case "eastAsiaTheme":
			s.EastAsiaTheme = attr.Value
		case "hAnsiTheme":
			s.HAnsiTheme = attr.Value
		case "cstheme":
			s.Cstheme = attr.Value
		case "ascii":
			s.Ascii = attr.Value
		case "eastAsia":
			s.EastAsia = attr.Value
		case "hAnsi":
			s.HAnsi = attr.Value
		case "cs":
			s.Cs = attr.Value
		default:
		}
	}
	_, err := d.Token()
	return err
}

type StyleShd struct {
	XMLName        xml.Name `xml:"w:shd,omitempty"`
	Val            string   `xml:"w:val,attr,omitempty"`
	Color          string   `xml:"w:color,attr,omitempty"`
	Fill           string   `xml:"w:fill,attr,omitempty"`
	ThemeFill      string   `xml:"w:themeFill,attr,omitempty"`
	ThemeFillShade string   `xml:"w:themeFillShade,attr,omitempty"`
	ThemeFillTint  string   `xml:"w:themeFillTint,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleShd) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "val":
			s.Val = attr.Value
		case "color":
			s.Color = attr.Value
		case "fill":
			s.Fill = attr.Value
		case "themeFill":
			s.ThemeFill = attr.Value
		case "themeFillShade":
			s.ThemeFillShade = attr.Value
		case "themeFillTint":
			s.ThemeFillTint = attr.Value
		default:
		}
	}
	_, err := d.Token()
	return err
}

type StyleSpacing struct {
	XMLName           xml.Name `xml:"w:spacing,omitempty"`
	Val               string   `xml:"w:val,attr,omitempty"`
	Before            string   `xml:"w:before,attr,omitempty"`
	BeforeAutoSpacing string   `xml:"w:beforeAutospacing,attr,omitempty"`
	After             string   `xml:"w:after,attr,omitempty"`
	AfterAutoSpacing  string   `xml:"w:afterAutospacing,attr,omitempty"`
	Line              string   `xml:"w:line,attr,omitempty"`
	LineRule          string   `xml:"w:lineRule,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleSpacing) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "val":
			s.Val = attr.Value
		case "after":
			s.After = attr.Value
		case "afterAutospacing":
			s.AfterAutoSpacing = attr.Value
		case "before":
			s.Before = attr.Value
		case "beforeAutospacing":
			s.BeforeAutoSpacing = attr.Value
		case "line":
			s.Line = attr.Value
		case "lineRule":
			s.LineRule = attr.Value
		default:
		}
	}
	_, err = d.Token()
	return
}

type StyleTblInd struct {
	XMLName xml.Name `xml:"w:tblInd,omitempty"`
	W       string   `xml:"w:w,attr,omitempty"`
	Type    string   `xml:"w:type,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleTblInd) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "type":
			s.Type = attr.Value
		case "w":
			s.W = attr.Value
		default:
		}
	}
	_, err = d.Token()
	return
}

type StyleTab struct {
	XMLName xml.Name `xml:"w:tab,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
	Leader  string   `xml:"w:leader,attr,omitempty"`
	Pos     string   `xml:"w:pos,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleTab) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "val":
			s.Val = attr.Value
		case "leader":
			s.Leader = attr.Value
		case "pos":
			s.Pos = attr.Value
		}
	}
	_, err = d.Token()
	return
}

type StyleTabs struct {
	XMLName xml.Name `xml:"w:tabs,omitempty"`
	Tabs    []*StyleTab
}

// UnmarshalXML
func (s *StyleTabs) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	s.Tabs = make([]*StyleTab, 0)
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
			case "tab":
				var v StyleTab
				err = d.DecodeElement(&v, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Tabs = append(s.Tabs, &v)
			default:
				err = d.Skip() // skip unsupported tags
			}
			if err != nil {
				return err
			}
		}
	}
	return nil
}

type StyleBdr struct {
	XMLName xml.Name `xml:"w:bdr,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
	Sz      string   `xml:"w:sz,attr,omitempty"`
	Space   string   `xml:"w:space,attr,omitempty"`
	Color   string   `xml:"w:color,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleBdr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "val":
			s.Val = attr.Value
		case "color":
			s.Color = attr.Value
		case "sz":
			s.Sz = attr.Value
		case "space":
			s.Space = attr.Value
		default:
		}
	}
	_, err = d.Token()
	return
}

type StylePBdr struct {
	XMLName xml.Name     `xml:"w:pBdr,omitempty"`
	Top     *StyleBorder `xml:"w:top,omitempty"`
	Left    *StyleBorder `xml:"w:left,omitempty"`
	Bottom  *StyleBorder `xml:"w:bottom,omitempty"`
	Right   *StyleBorder `xml:"w:right,omitempty"`
}

// UnmarshalXML
func (s *StylePBdr) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "top":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Top = &value
			case "bottom":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Bottom = &value
			case "left":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Left = &value
			case "right":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Right = &value
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

type StyleFramePr struct {
	XMLName xml.Name `xml:"w:framePr,omitempty"`
	W       string   `xml:"w:w,attr,omitempty"`
	H       string   `xml:"w:h,attr,omitempty"`
	HRule   string   `xml:"w:hRule,attr,omitempty"`
	HSpace  string   `xml:"w:hSpace,attr,omitempty"`
	Wrap    string   `xml:"w:wrap,attr,omitempty"`
	VAnchor string   `xml:"w:vAnchor,attr,omitempty"`
	HAnchor string   `xml:"w:hAnchor,attr,omitempty"`
	XAlign  string   `xml:"w:xAlign,attr,omitempty"`
	Y       string   `xml:"w:y,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleFramePr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "hRule":
			s.HRule = attr.Value
		case "wrap":
			s.Wrap = attr.Value
		case "vAnchor":
			s.VAnchor = attr.Value
		case "hAnchor":
			s.HAnchor = attr.Value
		case "xAlign":
			s.XAlign = attr.Value
		case "hSpace":
			s.HSpace = attr.Value
		case "w":
			s.W = attr.Value
		case "h":
			s.H = attr.Value
		case "y":
			s.Y = attr.Value
		default:
		}
	}
	_, err = d.Token()
	return
}

type StylePPr struct {
	XMLName           xml.Name `xml:"w:pPr,omitempty"`
	FramePr           *StyleFramePr
	KeepNext          *StyleVal `xml:"w:keepNext,omitempty"`
	KeepLines         *StyleVal `xml:"w:keepLines,omitempty"`
	WidowControl      *StyleVal `xml:"w:widowControl,omitempty"`
	NumPr             *StyleNumPr
	PBdr              *StylePBdr
	Shd               *StyleShd
	Tabs              *StyleTabs
	Spacing           *StyleSpacing
	Ind               *StyleInd
	Jc                *StyleVal `xml:"w:jc,omitempty"`
	ContextualSpacing *StyleVal `xml:"w:contextualSpacing,omitempty"`
	OutlineLvl        *StyleVal `xml:"w:outlineLvl,omitempty"`
	AutoSpaceDE       *StyleVal `xml:"w:autoSpaceDE,omitempty"`
	AutoSpaceDN       *StyleVal `xml:"w:autoSpaceDN,omitempty"`
	AdjustRightInd    *StyleVal `xml:"w:adjustRightInd,omitempty"`
}

// UnmarshalXML
func (s *StylePPr) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "adjustRightInd":
				s.AdjustRightInd = getVal(tt.Attr)
			case "autoSpaceDE":
				s.AutoSpaceDE = getVal(tt.Attr)
			case "autoSpaceDN":
				s.AutoSpaceDN = getVal(tt.Attr)
			case "contextualSpacing":
				s.ContextualSpacing = emptyStruct
			case "keepNext":
				s.KeepNext = getVal(tt.Attr)
			case "keepLines":
				s.KeepLines = emptyStruct
			case "framePr":
				var value StyleFramePr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.FramePr = &value
			case "ind":
				var value StyleInd
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Ind = &value
			case "shd":
				var value StyleShd
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Shd = &value
			case "jc":
				s.Jc = getVal(tt.Attr)
			case "numPr":
				var value StyleNumPr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.NumPr = &value
			case "widowControl":
				s.WidowControl = getVal(tt.Attr)
			case "outlineLvl":
				s.OutlineLvl = getVal(tt.Attr)
			case "pBdr":
				var value StylePBdr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.PBdr = &value
			case "spacing":
				var value StyleSpacing
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Spacing = &value
			case "tabs":
				var value StyleTabs
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Tabs = &value
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

type StyleRPr struct {
	XMLName    xml.Name `xml:"w:rPr,omitempty"`
	RFonts     *StyleRFonts
	B          *StyleVal `xml:"w:b,omitempty"`
	BCs        *StyleVal `xml:"w:bCs,omitempty"`
	I          *StyleVal `xml:"w:i,omitempty"`
	ICs        *StyleVal `xml:"w:iCs,omitempty"`
	Caps       *StyleVal `xml:"w:caps,omitempty"`
	SmallCaps  *StyleVal `xml:"w:smallCaps,omitempty"`
	Vanish     *StyleVal `xml:"w:vanish,omitempty"`
	Color      *StyleColor
	Spacing    *StyleSpacing
	SnapToGrid *StyleVal `xml:"w:snapToGrid,omitempty"`
	U          *StyleVal `xml:"w:u,omitempty"`
	UCs        *StyleVal `xml:"w:uCs,omitempty"`
	Position   *StyleVal `xml:"w:position,omitempty"`
	NoProof    *StyleVal `xml:"w:noProof,omitempty"`
	Kern       *StyleVal `xml:"w:kern,omitempty"`
	Sz         *StyleVal `xml:"w:sz,omitempty"`
	SzCs       *StyleVal `xml:"w:szCs,omitempty"`
	Shd        *StyleShd
	Lang       *StyleLang
	Ligatures  *StyleW14Val `xml:"w14:ligatures,omitempty"`
	Bdr        *StyleBdr
	VertAlign  *StyleVal `xml:"w:vertAlign,omitempty"`
}

// UnmarshalXML
func (s *StyleRPr) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "b":
				s.B = getVal(tt.Attr)
			case "bCs":
				s.BCs = getVal(tt.Attr)
			case "i":
				s.I = getVal(tt.Attr)
			case "iCs":
				s.ICs = getVal(tt.Attr)
			case "u":
				s.U = getVal(tt.Attr)
			case "uCs":
				s.UCs = getVal(tt.Attr)
			case "caps":
				s.Caps = emptyStruct
			case "smallCaps":
				s.SmallCaps = emptyStruct
			case "noProof":
				s.NoProof = emptyStruct
			case "vanish":
				s.Vanish = emptyStruct
			case "color":
				var value StyleColor
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Color = &value
				continue
			case "kern":
				s.Kern = getVal(tt.Attr)
			case "snapToGrid":
				s.SnapToGrid = getVal(tt.Attr)
			case "spacing":
				var value StyleSpacing
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Spacing = &value
				continue
			case "sz":
				s.Sz = getVal(tt.Attr)
			case "szCs":
				s.SzCs = getVal(tt.Attr)
			case "position":
				s.Position = getVal(tt.Attr)
			case "lang":
				var value StyleLang
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Lang = &value
				continue
			case "ligatures":
				s.Ligatures = getW14Val(tt.Attr)
			case "vertAlign":
				s.VertAlign = getVal(tt.Attr)
			case "rFonts":
				var value StyleRFonts
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.RFonts = &value
				continue
			case "bdr":
				var value StyleBdr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Bdr = &value
				continue
			case "shd":
				var value StyleShd
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Shd = &value
				continue
			default:
				err = d.Skip() // skip unsupported tags
			}
			if err != nil {
				return err
			}
		}
	}
	return nil
}

type StyleTcBorders struct {
	XMLName xml.Name     `xml:"w:tcBorders,omitempty"`
	Top     *StyleBorder `xml:"w:top,omitempty"`
	Left    *StyleBorder `xml:"w:left,omitempty"`
	Bottom  *StyleBorder `xml:"w:bottom,omitempty"`
	Right   *StyleBorder `xml:"w:right,omitempty"`
	InsideH *StyleBorder `xml:"w:insideH,omitempty"`
	InsideV *StyleBorder `xml:"w:insideV,omitempty"`
}

// UnmarshalXML
func (s *StyleTcBorders) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "top":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Top = &value
			case "bottom":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Bottom = &value
			case "left":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Left = &value
			case "right":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Right = &value
			case "insideH":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.InsideH = &value
			case "insideV":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.InsideV = &value
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

type StyleTblCellMar struct {
	XMLName xml.Name      `xml:"w:tblCellMar,omitempty"`
	Top     *StyleCellMar `xml:"w:top,omitempty"`
	Left    *StyleCellMar `xml:"w:left,omitempty"`
	Bottom  *StyleCellMar `xml:"w:bottom,omitempty"`
	Right   *StyleCellMar `xml:"w:right,omitempty"`
}

// UnmarshalXML
func (s *StyleTblCellMar) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "top":
				var value StyleCellMar
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Top = &value
			case "left":
				var value StyleCellMar
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Left = &value
			case "bottom":
				var value StyleCellMar
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Bottom = &value
			case "right":
				var value StyleCellMar
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Right = &value
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

type StyleTblBorders struct {
	XMLName xml.Name     `xml:"w:tblBorders,omitempty"`
	Top     *StyleBorder `xml:"w:top,omitempty"`
	Left    *StyleBorder `xml:"w:left,omitempty"`
	Bottom  *StyleBorder `xml:"w:bottom,omitempty"`
	Right   *StyleBorder `xml:"w:right,omitempty"`
	InsideH *StyleBorder `xml:"w:insideH,omitempty"`
	InsideV *StyleBorder `xml:"w:insideV,omitempty"`
}

// UnmarshalXML
func (s *StyleTblBorders) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "top":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Top = &value
			case "bottom":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Bottom = &value
			case "left":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Left = &value
			case "right":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Right = &value
			case "insideH":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.InsideH = &value
			case "insideV":
				var value StyleBorder
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.InsideV = &value
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

type StyleTcPr struct {
	XMLName   xml.Name `xml:"w:tcPr,omitempty"`
	TcBorders *StyleTcBorders
	VAlign    *StyleVal `xml:"w:vAlign,omitempty"`
	Shd       *StyleShd
}

// UnmarshalXML
func (s *StyleTcPr) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "tcBorders":
				var value StyleTcBorders
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.TcBorders = &value
			case "shd":
				var value StyleShd
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Shd = &value
			case "vAlign":
				s.VAlign = getVal(tt.Attr)
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

type StyleTblPr struct {
	XMLName             xml.Name `xml:"w:tblPr,omitempty"`
	TblInd              *StyleTblInd
	TblCellMar          *StyleTblCellMar
	TblStyleRowBandSize *StyleVal `xml:"w:tblStyleRowBandSize,omitempty"`
	TblStyleColBandSize *StyleVal `xml:"w:tblStyleColBandSize,omitempty"`
	TblBorders          *StyleTblBorders
}

// UnmarshalXML
func (s *StyleTblPr) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "tblInd":
				var value StyleTblInd
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.TblInd = &value
			case "tblCellMar":
				var value StyleTblCellMar
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.TblCellMar = &value
			case "tblBorders":
				var value StyleTblBorders
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.TblBorders = &value
			case "tblStyleRowBandSize":
				s.TblStyleRowBandSize = getVal(tt.Attr)
			case "tblStyleColBandSize":
				s.TblStyleColBandSize = getVal(tt.Attr)
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

type StyleTblStylePr struct {
	XMLName xml.Name `xml:"w:tblStylePr,omitempty"`
	Type    string   `xml:"w:type,attr,omitempty"`
	PPr     *StylePPr
	RPr     *StyleRPr
	TblPr   *StyleTblPr
	TcPr    *StyleTcPr
}

// UnmarshalXML
func (s *StyleTblStylePr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Value != "" && attr.Name.Local == "type" {
			s.Type = attr.Value
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
			switch tt.Name.Local {
			case "type":
				s.Type = getAtt(tt.Attr, "type")
			case "tblPr":
				var value StyleTblPr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.TblPr = &value
			case "pPr":
				var value StylePPr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.PPr = &value
			case "rPr":
				var value StyleRPr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.RPr = &value
			case "tcPr":
				var value StyleTcPr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.TcPr = &value
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

type StyleRPrDefault struct {
	XMLName xml.Name `xml:"w:rPrDefault,omitempty"`
	RPr     *StyleRPr
}

// UnmarshalXML
func (s *StyleRPrDefault) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
				var value StyleRPr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.RPr = &value
			} else {
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

type StylePPrDefault struct {
	XMLName xml.Name `xml:"w:pPrDefault,omitempty"`
	PPr     *StylePPr
}

// UnmarshalXML
func (s *StylePPrDefault) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			if tt.Name.Local == "pPr" {
				var value StylePPr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.PPr = &value
			} else {
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

type StyleDocDefaults struct {
	XMLName    xml.Name `xml:"w:docDefaults,omitempty"`
	RPrDefault *StyleRPrDefault
	PPrDefault *StylePPrDefault
}

// UnmarshalXML
func (s *StyleDocDefaults) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
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
			case "rPrDefault":
				var value StyleRPrDefault
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.RPrDefault = &value
			case "pPrDefault":
				var value StylePPrDefault
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.PPrDefault = &value
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

type StyleStyle struct {
	XMLName        xml.Name  `xml:"w:style,omitempty"`
	Type           string    `xml:"w:type,attr,omitempty"`
	Default        string    `xml:"w:default,attr,omitempty"`
	CustomStyle    string    `xml:"w:customStyle,attr,omitempty"`
	StyleId        string    `xml:"w:styleId,attr,omitempty"`
	Name           *StyleVal `xml:"w:name,omitempty"`
	Aliases        *StyleVal `xml:"w:aliases,omitempty"`
	BasedOn        *StyleVal `xml:"w:basedOn,omitempty"`
	Next           *StyleVal `xml:"w:next,omitempty"`
	Link           *StyleVal `xml:"w:link,omitempty"`
	AutoRedifine   *StyleVal `xml:"w:autoRedefine,omitempty"`
	Hidden         *StyleVal `xml:"w:hidden,omitempty"`
	UiPriority     *StyleVal `xml:"w:uiPriority,omitempty"`
	SemiHidden     *StyleVal `xml:"w:semiHidden,omitempty"`
	Locked         *StyleVal `xml:"w:locked,omitempty"`
	UnhideWhenUsed *StyleVal `xml:"w:unhideWhenUsed,omitempty"`
	QFormat        *StyleVal `xml:"w:qFormat,omitempty"`
	Rsid           *StyleVal `xml:"w:rsid,omitempty"`
	PPr            *StylePPr
	RPr            *StyleRPr
	TblPr          *StyleTblPr
	TcPr           *StyleTcPr
	TblStylePr     []*StyleTblStylePr
}

// UnmarshalXML...
func (s *StyleStyle) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "type":
			s.Type = attr.Value
		case "default":
			s.Default = attr.Value
		case "styleId":
			s.StyleId = attr.Value
		case "customStyle":
			s.CustomStyle = attr.Value
		default:
		}
	}
	for {
		token, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if tt, ok := token.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "qFormat":
				s.QFormat = emptyStruct
			case "hidden":
				s.Hidden = emptyStruct
			case "unhideWhenUsed":
				s.UnhideWhenUsed = emptyStruct
			case "semiHidden":
				s.SemiHidden = emptyStruct
			case "autoRedefine":
				s.AutoRedifine = emptyStruct
			case "locked":
				s.Locked = emptyStruct
			case "name":
				s.Name = getVal(tt.Attr)
			case "aliases":
				s.Aliases = getVal(tt.Attr)
			case "basedOn":
				s.BasedOn = getVal(tt.Attr)
			case "next":
				s.Next = getVal(tt.Attr)
			case "link":
				s.Link = getVal(tt.Attr)
			case "uiPriority":
				s.UiPriority = getVal(tt.Attr)
			case "rsid":
				s.Rsid = getVal(tt.Attr)
			case "pPr":
				var value StylePPr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.PPr = &value
			case "rPr":
				var value StyleRPr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.RPr = &value
			case "tblPr":
				var value StyleTblPr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.TblPr = &value
			case "tcPr":
				var value StyleTcPr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.TcPr = &value
			case "tblStylePr":
				var value StyleTblStylePr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.TblStylePr = append(s.TblStylePr, &value)
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

type Styles struct {
	XMLName      xml.Name `xml:"w:styles,omitempty"`
	Mc           string   `xml:"xmlns:mc,attr,omitempty"`
	R            string   `xml:"xmlns:r,attr,omitempty"`
	W            string   `xml:"xmlns:w,attr,omitempty"`
	W14          string   `xml:"xmlns:w14,attr,omitempty"`
	W15          string   `xml:"xmlns:w15,attr,omitempty"`
	W16cex       string   `xml:"xmlns:w16cex,attr,omitempty"`
	W16cid       string   `xml:"xmlns:w16cid,attr,omitempty"`
	W16          string   `xml:"xmlns:w16,attr,omitempty"`
	W16du        string   `xml:"xmlns:w16du,attr,omitempty"`
	W16sdtdh     string   `xml:"xmlns:w16sdtdh,attr,omitempty"`
	W16sdtfl     string   `xml:"xmlns:w16sdtfl,attr,omitempty"`
	W16se        string   `xml:"xmlns:w16se,attr,omitempty"`
	Ignorable    string   `xml:"mc:Ignorable,attr,omitempty"`
	DocDefaults  *StyleDocDefaults
	LatentStyles *StyleLatentStyles
	Styles       []*StyleStyle
}

// UnmarshalXML...
func (s *Styles) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "mc":
			s.Mc = attr.Value
		case "r":
			s.R = attr.Value
		case "w":
			s.W = attr.Value
		case "w14":
			s.W14 = attr.Value
		case "w15":
			s.W15 = attr.Value
		case "w16cex":
			s.W16cex = attr.Value
		case "w16cid":
			s.W16cid = attr.Value
		case "w16":
			s.W16 = attr.Value
		case "w16du":
			s.W16du = attr.Value
		case "w16sdtdh":
			s.W16sdtdh = attr.Value
		case "w16sdtfl":
			s.W16sdtfl = attr.Value
		case "w16se":
			s.W16se = attr.Value
		case "Ignorable":
			s.Ignorable = attr.Value
		default:
		}
	}
	for {
		token, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if tt, ok := token.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "docDefaults":
				var value StyleDocDefaults
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.DocDefaults = &value
			case "latentStyles":
				var value StyleLatentStyles
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.LatentStyles = &value
			case "style":
				var value StyleStyle
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Styles = append(s.Styles, &value)
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
