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

func getVal(atts []xml.Attr) string {
	for _, at := range atts {
		if at.Name.Local == "val" {
			return at.Value
		}
	}
	return ""
}

func trapExpectedError(err error) bool {
	return err != nil && !strings.HasPrefix(err.Error(), "expected")
}

// types that do not require UnmarshalXML methods
type StyleAutoRedefine struct {
	XMLName xml.Name `xml:"w:autoRedefine,omitempty"`
}
type StyleContextualSpacing struct {
	XMLName xml.Name `xml:"w:contextualSpacing,omitempty"`
}
type StyleHidden struct {
	XMLName xml.Name `xml:"w:hidden,omitempty"`
}
type StyleKeepLines struct {
	XMLName xml.Name `xml:"w:keepLines,omitempty"`
}
type StyleLocked struct {
	XMLName xml.Name `xml:"w:locked,omitempty"`
}
type StyleNoProof struct {
	XMLName xml.Name `xml:"w:noProof,omitempty"`
}
type StyleQFormat struct {
	XMLName xml.Name `xml:"w:qFormat,omitempty"`
}
type StyleSemiHidden struct {
	XMLName xml.Name `xml:"w:semiHidden,omitempty"`
}
type StyleSmallCaps struct {
	XMLName xml.Name `xml:"w:smallCaps,omitempty"`
}
type StyleCaps struct {
	XMLName xml.Name `xml:"w:caps,omitempty"`
}
type StyleUnhideWhenUsed struct {
	XMLName xml.Name `xml:"w:unhideWhenUsed,omitempty"`
}
type StyleVanish struct {
	XMLName xml.Name `xml:"w:vanish,omitempty"`
}

// types that has attr Val only and require getAtt
type StyleAliases struct {
	XMLName xml.Name `xml:"w:aliases,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleAutoSpaceDE struct {
	XMLName xml.Name `xml:"w:autoSpaceDE,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleAutoSpaceDN struct {
	XMLName xml.Name `xml:"w:autoSpaceDN,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleAdjustRightInd struct {
	XMLName xml.Name `xml:"w:adjustRightInd,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleB struct {
	XMLName xml.Name `xml:"w:b,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleBCs struct {
	XMLName xml.Name `xml:"w:bCs,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleI struct {
	XMLName xml.Name `xml:"w:i,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleICs struct {
	XMLName xml.Name `xml:"w:iCs,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleU struct {
	XMLName xml.Name `xml:"w:u,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleUCs struct {
	XMLName xml.Name `xml:"w:uCs,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleBasedOn struct {
	XMLName xml.Name `xml:"w:basedOn,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleIlvl struct {
	XMLName xml.Name `xml:"w:ilvl,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleJc struct {
	XMLName xml.Name `xml:"w:jc,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleKeepNext struct {
	XMLName xml.Name `xml:"w:keepNext,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleKern struct {
	XMLName xml.Name `xml:"w:kern,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleLigatures struct {
	XMLName xml.Name `xml:"w14:ligatures,omitempty"`
	Val     string   `xml:"w14:val,attr,omitempty"`
}
type StyleLink struct {
	XMLName xml.Name `xml:"w:link,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleName struct {
	XMLName xml.Name `xml:"w:name,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleNext struct {
	XMLName xml.Name `xml:"w:next,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleNumID struct {
	XMLName xml.Name `xml:"w:numId,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleOutlineLvl struct {
	XMLName xml.Name `xml:"w:outlineLvl,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleRsid struct {
	XMLName xml.Name `xml:"w:rsid,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StylePosition struct {
	XMLName xml.Name `xml:"w:position,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleSnapToGrid struct {
	XMLName xml.Name `xml:"w:snapToGrid,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleSz struct {
	XMLName xml.Name `xml:"w:sz,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleSzCs struct {
	XMLName xml.Name `xml:"w:szCs,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleTblStyleRowBandSize struct {
	XMLName xml.Name `xml:"w:tblStyleRowBandSize,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleTblStyleColBandSize struct {
	XMLName xml.Name `xml:"w:tblStyleColBandSize,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleUiPriority struct {
	XMLName xml.Name `xml:"w:uiPriority,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleVertAlign struct {
	XMLName xml.Name `xml:"w:vertAlign,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleVAlign struct {
	XMLName xml.Name `xml:"w:vAlign,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
type StyleWidowControl struct {
	XMLName xml.Name `xml:"w:widowControl,omitempty"`
	Val     string   `xml:"w:val,attr,omitempty"`
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
			if err != nil {
				return err
			}
		case "qFormat":
			s.QFormat = attr.Value
			if err != nil {
				return err
			}
		case "semiHidden":
			s.SemiHidden = attr.Value
			if err != nil {
				return err
			}
		case "unhideWhenUsed":
			s.UnhideWhenUsed = attr.Value
			if err != nil {
				return err
			}
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
	var err error
	for _, attr := range start.Attr {
		if attr.Value == "" {
			continue
		}
		switch attr.Name.Local {
		case "defLockedState":
			s.DefLockedState = attr.Value
			if err != nil {
				return err
			}
		case "defUIPriority":
			s.DefUIPriority = attr.Value
			if err != nil {
				return err
			}
		case "defQFormat":
			s.DefQFormat = attr.Value
			if err != nil {
				return err
			}
		case "defSemiHidden":
			s.DefSemiHidden = attr.Value
			if err != nil {
				return err
			}
		case "defUnhideWhenUsed":
			s.DefUnhideWhenUsed = attr.Value
			if err != nil {
				return err
			}
		case "count":
			s.Count = attr.Value
			if err != nil {
				return err
			}
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

type StyleBorderTop struct {
	XMLName    xml.Name `xml:"w:top,omitempty"`
	Val        string   `xml:"w:val,attr,omitempty"`
	Sz         string   `xml:"w:sz,attr,omitempty"`
	Space      string   `xml:"w:space,attr,omitempty"`
	Color      string   `xml:"w:color,attr,omitempty"`
	ThemeColor string   `xml:"w:themeColor,attr,omitempty"`
	ThemeShade string   `xml:"w:themeShade,attr,omitempty"`
	ThemeTint  string   `xml:"w:themeTint,attr,omitempty"`
	W          string   `xml:"w:w,attr,omitempty"`
	Type       string   `xml:"w:type,attr,omitempty"`
	Shadow     string   `xml:"w:shadow,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleBorderTop) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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

type StyleBorderBottom struct {
	XMLName    xml.Name `xml:"w:bottom,omitempty"`
	Val        string   `xml:"w:val,attr,omitempty"`
	Sz         string   `xml:"w:sz,attr,omitempty"`
	Space      string   `xml:"w:space,attr,omitempty"`
	Color      string   `xml:"w:color,attr,omitempty"`
	ThemeColor string   `xml:"w:themeColor,attr,omitempty"`
	ThemeShade string   `xml:"w:themeShade,attr,omitempty"`
	ThemeTint  string   `xml:"w:themeTint,attr,omitempty"`
	W          string   `xml:"w:w,attr,omitempty"`
	Type       string   `xml:"w:type,attr,omitempty"`
	Shadow     string   `xml:"w:shadow,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleBorderBottom) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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

type StyleBorderLeft struct {
	XMLName    xml.Name `xml:"w:left,omitempty"`
	Val        string   `xml:"w:val,attr,omitempty"`
	Sz         string   `xml:"w:sz,attr,omitempty"`
	Space      string   `xml:"w:space,attr,omitempty"`
	Color      string   `xml:"w:color,attr,omitempty"`
	ThemeColor string   `xml:"w:themeColor,attr,omitempty"`
	ThemeShade string   `xml:"w:themeShade,attr,omitempty"`
	ThemeTint  string   `xml:"w:themeTint,attr,omitempty"`
	W          string   `xml:"w:w,attr,omitempty"`
	Type       string   `xml:"w:type,attr,omitempty"`
	Shadow     string   `xml:"w:shadow,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleBorderLeft) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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

type StyleBorderRight struct {
	XMLName    xml.Name `xml:"w:right,omitempty"`
	Val        string   `xml:"w:val,attr,omitempty"`
	Sz         string   `xml:"w:sz,attr,omitempty"`
	Space      string   `xml:"w:space,attr,omitempty"`
	Color      string   `xml:"w:color,attr,omitempty"`
	ThemeColor string   `xml:"w:themeColor,attr,omitempty"`
	ThemeShade string   `xml:"w:themeShade,attr,omitempty"`
	ThemeTint  string   `xml:"w:themeTint,attr,omitempty"`
	W          string   `xml:"w:w,attr,omitempty"`
	Type       string   `xml:"w:type,attr,omitempty"`
	Shadow     string   `xml:"w:shadow,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleBorderRight) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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

type StyleInsideH struct {
	XMLName    xml.Name `xml:"w:insideH,omitempty"`
	Val        string   `xml:"w:val,attr,omitempty"`
	Sz         string   `xml:"w:sz,attr,omitempty"`
	Space      string   `xml:"w:space,attr,omitempty"`
	Color      string   `xml:"w:color,attr,omitempty"`
	ThemeColor string   `xml:"w:themeColor,attr,omitempty"`
	ThemeShade string   `xml:"w:themeShade,attr,omitempty"`
	ThemeTint  string   `xml:"w:themeTint,attr,omitempty"`
	W          string   `xml:"w:w,attr,omitempty"`
	Type       string   `xml:"w:type,attr,omitempty"`
	Shadow     string   `xml:"w:shadow,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleInsideH) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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

type StyleInsideV struct {
	XMLName    xml.Name `xml:"w:insideV,omitempty"`
	Val        string   `xml:"w:val,attr,omitempty"`
	Sz         string   `xml:"w:sz,attr,omitempty"`
	Space      string   `xml:"w:space,attr,omitempty"`
	Color      string   `xml:"w:color,attr,omitempty"`
	ThemeColor string   `xml:"w:themeColor,attr,omitempty"`
	ThemeShade string   `xml:"w:themeShade,attr,omitempty"`
	ThemeTint  string   `xml:"w:themeTint,attr,omitempty"`
	W          string   `xml:"w:w,attr,omitempty"`
	Type       string   `xml:"w:type,attr,omitempty"`
	Shadow     string   `xml:"w:shadow,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleInsideV) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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

type StyleMarginTop struct {
	XMLName xml.Name `xml:"w:top,omitempty"`
	W       string   `xml:"w:w,attr,omitempty"`
	Type    string   `xml:"w:type,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleMarginTop) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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

type StyleMarginBottom struct {
	XMLName xml.Name `xml:"w:bottom,omitempty"`
	W       string   `xml:"w:w,attr,omitempty"`
	Type    string   `xml:"w:type,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleMarginBottom) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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

type StyleMarginLeft struct {
	XMLName xml.Name `xml:"w:left,omitempty"`
	W       string   `xml:"w:w,attr,omitempty"`
	Type    string   `xml:"w:type,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleMarginLeft) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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

type StyleMarginRight struct {
	XMLName xml.Name `xml:"w:right,omitempty"`
	W       string   `xml:"w:w,attr,omitempty"`
	Type    string   `xml:"w:type,attr,omitempty"`
}

// UnmarshalXML...
func (s *StyleMarginRight) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
	XMLName xml.Name `xml:"w:numPr,omitempty"`
	Ilvl    *StyleIlvl
	NumId   *StyleNumID
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
				s.Ilvl = &StyleIlvl{Val: getAtt(tt.Attr, "val")}
			case "numId":
				s.NumId = &StyleNumID{Val: getAtt(tt.Attr, "val")}
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
		if err != nil {
			return
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
			if err != nil {
				return
			}
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
	XMLName xml.Name `xml:"w:pBdr,omitempty"`
	Top     *StyleBorderTop
	Left    *StyleBorderLeft
	Bottom  *StyleBorderBottom
	Right   *StyleBorderRight
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
				var value StyleBorderTop
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Top = &value
			case "bottom":
				var value StyleBorderBottom
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Bottom = &value
			case "left":
				var value StyleBorderLeft
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Left = &value
			case "right":
				var value StyleBorderRight
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
		if err != nil {
			return
		}
	}
	_, err = d.Token()
	return
}

type StylePPr struct {
	XMLName           xml.Name `xml:"w:pPr,omitempty"`
	FramePr           *StyleFramePr
	KeepNext          *StyleKeepNext
	KeepLines         *StyleKeepLines
	WidowControl      *StyleWidowControl
	NumPr             *StyleNumPr
	PBdr              *StylePBdr
	Shd               *StyleShd
	Tabs              *StyleTabs
	Spacing           *StyleSpacing
	Ind               *StyleInd
	Jc                *StyleJc
	ContextualSpacing *StyleContextualSpacing
	OutlineLvl        *StyleOutlineLvl
	AutoSpaceDE       *StyleAutoSpaceDE
	AutoSpaceDN       *StyleAutoSpaceDN
	AdjustRightInd    *StyleAdjustRightInd
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
				s.AdjustRightInd = &StyleAdjustRightInd{Val: getAtt(tt.Attr, "val")}
			case "autoSpaceDE":
				s.AutoSpaceDE = &StyleAutoSpaceDE{Val: getAtt(tt.Attr, "val")}
			case "autoSpaceDN":
				s.AutoSpaceDN = &StyleAutoSpaceDN{Val: getAtt(tt.Attr, "val")}
			case "contextualSpacing":
				s.ContextualSpacing = &StyleContextualSpacing{}
			case "keepNext":
				s.KeepNext = &StyleKeepNext{Val: getAtt(tt.Attr, "val")}
			case "keepLines":
				s.KeepLines = &StyleKeepLines{}
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
				s.Jc = &StyleJc{Val: getVal(tt.Attr)}
			case "numPr":
				var value StyleNumPr
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.NumPr = &value
			case "widowControl":
				s.WidowControl = &StyleWidowControl{Val: getAtt(tt.Attr, "val")}
			case "outlineLvl":
				s.OutlineLvl = &StyleOutlineLvl{Val: getAtt(tt.Attr, "val")}
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
	B          *StyleB
	BCs        *StyleBCs
	I          *StyleI
	ICs        *StyleICs
	Caps       *StyleCaps
	SmallCaps  *StyleSmallCaps
	Vanish     *StyleVanish
	Color      *StyleColor
	Spacing    *StyleSpacing
	SnapToGrid *StyleSnapToGrid
	U          *StyleU
	UCs        *StyleUCs
	Position   *StylePosition
	NoProof    *StyleNoProof
	Kern       *StyleKern
	Sz         *StyleSz
	SzCs       *StyleSzCs
	Shd        *StyleShd
	Lang       *StyleLang
	Ligatures  *StyleLigatures
	Bdr        *StyleBdr
	VertAlign  *StyleVertAlign
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
				s.B = &StyleB{Val: getAtt(tt.Attr, "val")}
			case "bCs":
				s.BCs = &StyleBCs{Val: getAtt(tt.Attr, "val")}
			case "i":
				s.I = &StyleI{Val: getAtt(tt.Attr, "val")}
			case "iCs":
				s.ICs = &StyleICs{Val: getAtt(tt.Attr, "val")}
			case "u":
				s.U = &StyleU{Val: getVal(tt.Attr)}
			case "uCs":
				s.UCs = &StyleUCs{Val: getVal(tt.Attr)}
			case "caps":
				s.Caps = &StyleCaps{}
			case "smallCaps":
				s.SmallCaps = &StyleSmallCaps{}
			case "noProof":
				s.NoProof = &StyleNoProof{}
			case "vanish":
				s.Vanish = &StyleVanish{}
			case "color":
				var value StyleColor
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Color = &value
				continue
			case "kern":
				s.Kern = &StyleKern{Val: getAtt(tt.Attr, "val")}
			case "snapToGrid":
				s.SnapToGrid = &StyleSnapToGrid{Val: getAtt(tt.Attr, "val")}
			case "spacing":
				var value StyleSpacing
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Spacing = &value
				continue
			case "sz":
				s.Sz = &StyleSz{Val: getAtt(tt.Attr, "val")}
			case "szCs":
				s.SzCs = &StyleSzCs{Val: getAtt(tt.Attr, "val")}
			case "position":
				s.Position = &StylePosition{Val: getAtt(tt.Attr, "val")}
			case "lang":
				var value StyleLang
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Lang = &value
				continue
			case "ligatures":
				s.Ligatures = &StyleLigatures{Val: getVal(tt.Attr)}
			case "vertAlign":
				s.VertAlign = &StyleVertAlign{Val: getVal(tt.Attr)}
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
	XMLName xml.Name `xml:"w:tcBorders,omitempty"`
	Top     *StyleBorderTop
	Left    *StyleBorderLeft
	Bottom  *StyleBorderBottom
	Right   *StyleBorderRight
	InsideH *StyleInsideH
	InsideV *StyleInsideV
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
				var value StyleBorderTop
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Top = &value
			case "bottom":
				var value StyleBorderBottom
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Bottom = &value
			case "left":
				var value StyleBorderLeft
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Left = &value
			case "right":
				var value StyleBorderRight
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Right = &value
			case "insideH":
				var value StyleInsideH
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.InsideH = &value
			case "insideV":
				var value StyleInsideV
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
	XMLName xml.Name `xml:"w:tblCellMar,omitempty"`
	Top     *StyleMarginTop
	Left    *StyleMarginLeft
	Bottom  *StyleMarginBottom
	Right   *StyleMarginRight
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
				var value StyleMarginTop
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Top = &value
			case "left":
				var value StyleMarginLeft
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Left = &value
			case "bottom":
				var value StyleMarginBottom
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Bottom = &value
			case "right":
				var value StyleMarginRight
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
	XMLName xml.Name `xml:"w:tblBorders,omitempty"`
	Top     *StyleBorderTop
	Left    *StyleBorderLeft
	Bottom  *StyleBorderBottom
	Right   *StyleBorderRight
	InsideH *StyleInsideH
	InsideV *StyleInsideV
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
				var value StyleBorderTop
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Top = &value
			case "bottom":
				var value StyleBorderBottom
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Bottom = &value
			case "left":
				var value StyleBorderLeft
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Left = &value
			case "right":
				var value StyleBorderRight
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.Right = &value
			case "insideH":
				var value StyleInsideH
				err = d.DecodeElement(&value, &tt)
				if trapExpectedError(err) {
					return err
				}
				s.InsideH = &value
			case "insideV":
				var value StyleInsideV
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
	VAlign    *StyleVAlign
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
				s.VAlign = &StyleVAlign{Val: getAtt(tt.Attr, "val")}
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
	TblStyleRowBandSize *StyleTblStyleRowBandSize
	TblStyleColBandSize *StyleTblStyleColBandSize
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
				s.TblStyleRowBandSize = &StyleTblStyleRowBandSize{Val: getAtt(tt.Attr, "val")}
			case "tblStyleColBandSize":
				s.TblStyleColBandSize = &StyleTblStyleColBandSize{Val: getAtt(tt.Attr, "val")}
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
	XMLName        xml.Name `xml:"w:style,omitempty"`
	Type           string   `xml:"w:type,attr,omitempty"`
	Default        string   `xml:"w:default,attr,omitempty"`
	CustomStyle    string   `xml:"w:customStyle,attr,omitempty"`
	StyleId        string   `xml:"w:styleId,attr,omitempty"`
	Name           *StyleName
	Aliases        *StyleAliases
	BasedOn        *StyleBasedOn
	Next           *StyleNext
	Link           *StyleLink
	AutoRedifine   *StyleAutoRedefine
	Hidden         *StyleHidden
	UiPriority     *StyleUiPriority
	SemiHidden     *StyleSemiHidden
	Locked         *StyleLocked
	UnhideWhenUsed *StyleUnhideWhenUsed
	QFormat        *StyleQFormat
	Rsid           *StyleRsid
	PPr            *StylePPr
	RPr            *StyleRPr
	TblPr          *StyleTblPr
	TcPr           *StyleTcPr
	TblStylePr     []*StyleTblStylePr
}

// UnmarshalXML...
func (s *StyleStyle) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var err error
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
		if err != nil {
			return nil
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
				s.QFormat = &StyleQFormat{}
			case "hidden":
				s.Hidden = &StyleHidden{}
			case "unhideWhenUsed":
				s.UnhideWhenUsed = &StyleUnhideWhenUsed{}
			case "semiHidden":
				s.SemiHidden = &StyleSemiHidden{}
			case "autoRedefine":
				s.AutoRedifine = &StyleAutoRedefine{}
			case "locked":
				s.Locked = &StyleLocked{}
			case "name":
				s.Name = &StyleName{Val: getVal(tt.Attr)}
			case "aliases":
				s.Aliases = &StyleAliases{Val: getVal(tt.Attr)}
			case "basedOn":
				s.BasedOn = &StyleBasedOn{Val: getVal(tt.Attr)}
			case "next":
				s.Next = &StyleNext{Val: getVal(tt.Attr)}
			case "link":
				s.Link = &StyleLink{Val: getVal(tt.Attr)}
			case "uiPriority":
				s.UiPriority = &StyleUiPriority{Val: getAtt(tt.Attr, "val")}
			case "rsid":
				s.Rsid = &StyleRsid{Val: getVal(tt.Attr)}
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
