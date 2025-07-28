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
	"strings"
)

// Hyperlink element contains links
type Hyperlink struct {
	XMLName xml.Name `xml:"w:hyperlink,omitempty"`
	ID      string   `xml:"r:id,attr"`
	Anchor  string   `xml:"w:anchor,attr,omitempty"`
	History int      `xml:"w:history,attr,omitempty"`
	Run     []*Run
	file    *Docx
}

func (o *Hyperlink) String() string {
	var s string
	for _, r := range o.Run {
		if r.InstrText.Text != "" {
			s += "[" + (string)(r.InstrText.Text) + "]"
		}

	}
	link, err := o.file.ReferTarget(o.ID)
	if err != nil {
		s += "(" + o.ID + ")"
	} else {
		s += "(" + link + ")"
	}
	return s
}

// UnmarshalXML ...
func (h *Hyperlink) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	h.Run = make([]*Run, 0)
	for {
		var t xml.Token
		t, err = d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return
		}

		if tt, ok := t.(xml.StartElement); ok {
			if tt.Name.Local == "r" {
				var v Run
				err = d.DecodeElement(&v, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				h.Run = append(h.Run, &v)
				continue
			}
			err = d.Skip() // skip unsupported tags
			if err != nil {
				return
			}
		}
	}
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "id":
			h.ID = a.Value
		case "history":
			h.History, err = GetInt(a.Value)
		case "anchor":
			h.Anchor = a.Value
		}
	}

	if h.ID == "" {
		h.ID = h.Anchor
	}

	return
}
