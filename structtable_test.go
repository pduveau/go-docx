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
	"hash/crc64"
	"io"
	"os"
	"testing"
)

func TestTableStructure(t *testing.T) {

	w := New().WithDefaultTheme()
	// add new paragraph
	para1 := w.AddParagraph()
	// add text
	para1.AddText("table")
	tab1 := w.AddTable(4, 3, 1000).Justification("center")
	tab1.Borders(TABLE_BORDER_EXTERN|TABLE_BORDER_INSIDEH, "single", "#ff0000", 4, 0)
	tab1.Properties.Position = &WTablePositioningProperties{LeftFromText: 2333}
	para2 := tab1.Rows[3].Justification("center").Cells[2].Shade("clear", "auto", "E7E6E6").AddParagraph()
	r, err := para2.AddAnchorDrawingFrom("testdata/fumiama.JPG")
	if err != nil {
		t.Fatal(err)
	}
	tab1.Rows[0].Cells[0].Properties.VMerge = &WvMerge{Val: "restart"}
	tab1.Rows[1].Cells[0].Properties.VMerge = &WvMerge{}
	tab1.Rows[2].Cells[0].Properties.VMerge = &WvMerge{}
	r.Children[0].(*Drawing).Anchor.Graphic.GraphicData.Pic.BlipFill.Blip.AlphaModFix = &AAlphaModFix{Amount: 50000}
	r.Children[0].(*Drawing).Anchor.Graphic.GraphicData.Pic.NonVisualPicProperties.CNvPicPr.Locks = &APicLocks{NoChangeAspect: 1}
	r.Children[0].(*Drawing).Anchor.Graphic.GraphicData.Pic.SpPr.Xfrm.Rot = 50000
	para3 := tab1.Rows[0].Cells[0].AddParagraph()
	para3.AddText("first cell")

	f, err := os.Create("TestMarshalTableStructure.xml")
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	_, err = marshaller{data: &w.Document}.WriteTo(f)
	if err != nil {
		t.Fatal(err)
	}
	_, err = f.Seek(0, io.SeekStart)
	if err != nil {
		t.Fatal(err)
	}
	w = New().WithDefaultTheme()
	err = xml.NewDecoder(f).Decode(&w.Document)
	if err != nil {
		t.Fatal(err)
	}
	f1, err := os.Create("TestUnmarshalTableStructure.xml")
	if err != nil {
		t.Fatal(err)
	}
	defer f1.Close()
	_, err = marshaller{data: &w.Document}.WriteTo(f1)
	if err != nil {
		t.Fatal(err)
	}
	_, err = f.Seek(0, io.SeekStart)
	if err != nil {
		t.Fatal(err)
	}
	_, err = f1.Seek(0, io.SeekStart)
	if err != nil {
		t.Fatal(err)
	}
	h := crc64.New(crc64.MakeTable(crc64.ECMA))
	_, err = io.Copy(h, f)
	if err != nil {
		t.Fatal(err)
	}
	crc1 := h.Sum64()
	h.Reset()
	_, err = io.Copy(h, f1)
	if err != nil {
		t.Fatal(err)
	}
	crc2 := h.Sum64()
	if crc1 != crc2 {
		t.Fail()
	}
}
