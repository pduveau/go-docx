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
)

func newEmptyFile() *Docx {
	docx := &Docx{
		Document: Document{
			XMLName: xml.Name{
				Space: "w",
			},
			XMLwpc:      XMLNS_WPC,
			XMLcx:       XMLNS_CX,
			XMLcx1:      XMLNS_CX1,
			XMLcx2:      XMLNS_CX2,
			XMLcx3:      XMLNS_CX3,
			XMLcx4:      XMLNS_CX4,
			XMLcx5:      XMLNS_CX5,
			XMLcx6:      XMLNS_CX6,
			XMLcx7:      XMLNS_CX7,
			XMLcx8:      XMLNS_CX8,
			XMLmc:       XMLNS_MC,
			XMLaink:     XMLNS_AINK,
			XMLam3d:     XMLNS_AM3D,
			XMLo:        XMLNS_O,
			XMLoel:      XMLNS_OEL,
			XMLr:        XMLNS_R,
			XMLm:        XMLNS_M,
			XMLv:        XMLNS_V,
			XMLwp14:     XMLNS_WP14,
			XMLwp:       XMLNS_WP,
			XMLw10:      XMLNS_W10,
			XMLw:        XMLNS_W,
			XMLw14:      XMLNS_W14,
			XMLw15:      XMLNS_W15,
			XMLw16cex:   XMLNS_W16CEX,
			XMLw16cid:   XMLNS_W16CID,
			XMLw16:      XMLNS_W16,
			XMLw16du:    XMLNS_W16DU,
			XMLw16sdtdh: XMLNS_W16SDTDH,
			XMLw16sdtfl: XMLNS_W16SDTFL,
			XMLw16se:    XMLNS_W16SE,
			XMLwpg:      XMLNS_WPG,
			XMLwpi:      XMLNS_WPI,
			XMLwne:      XMLNS_WNE,
			XMLwps:      XMLNS_WPS,
			MCIgnorable: MC_IGNORABLE,
			Body: Body{
				Items: make([]interface{}, 0, 64),
			},
		},
		docRelation: Relationships{
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
		},
		media:        make([]Media, 0, 64),
		mediaNameIdx: make(map[string]int, 64),
		rID:          3,
		picturesId:   0,
	}
	docx.Document.Body.file = docx
	return docx
}
