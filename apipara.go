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

// AddParagraph adds a new paragraph
func (f *Docx) AddParagraph() *Paragraph {
	p := &Paragraph{
		Children: make([]interface{}, 0, 64),
		file:     f,
	}
	f.Document.Body.Items = append(f.Document.Body.Items, p)
	return p
}

// AddParagraph adds a new paragraph
func (c *WTableCell) AddParagraph() *Paragraph {
	c.Paragraphs = append(c.Paragraphs, &Paragraph{
		Children: make([]interface{}, 0, 64),
		file:     c.file,
	})

	return c.Paragraphs[len(c.Paragraphs)-1]
}

// Justification allows to set para's horizonal alignment
func (p *Paragraph) Justification(val _justification) *Paragraph {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}
	p.Properties.Justification = &Justification{Val: (string)(val)}
	return p
}

// AddPageBreaks adds a pagebreaks
func (p *Paragraph) AddPageBreaks() *Run {
	c := make([]interface{}, 1, 64)
	c[0] = &BarterRabbet{
		Type: "page",
	}
	run := &Run{
		RunProperties: &RunProperties{},
		Children:      c,
	}
	p.Children = append(p.Children, run)
	return run
}

// Style name
func (p *Paragraph) Style(val string) *Paragraph {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}
	p.Properties.Style = &Style{Val: val}
	return p
}

// NumPr number properties
func (p *Paragraph) NumPr(numID, ilvl string) *Paragraph {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}
	// Initialize run properties if not exist
	if p.Properties.RunProperties == nil {
		p.Properties.RunProperties = &RunProperties{}
	}
	p.Properties.NumProperties = &NumProperties{
		NumID: &NumID{
			Val: numID,
		},
		Ilvl: &Ilevel{
			Val: ilvl,
		},
	}
	return p
}

// NumFont sets the font for numbering
func (p *Paragraph) NumFont(ascii, eastAsia, hansi, hint string) *Paragraph {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}
	if p.Properties.RunProperties == nil {
		p.Properties.RunProperties = &RunProperties{}
	}
	p.Properties.RunProperties.Fonts = &RunFonts{
		ASCII:    ascii,
		EastAsia: eastAsia,
		HAnsi:    hansi,
		Hint:     hint,
	}
	return p
}

// NumSize sets the size for numbering
func (p *Paragraph) NumSize(size string) *Paragraph {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}
	if p.Properties.RunProperties == nil {
		p.Properties.RunProperties = &RunProperties{}
	}
	p.Properties.RunProperties.Size = &Size{Val: size}
	return p
}

// LangCheck set the language parameter
// if a parameter is a string it will be used as the language
// if a parameter is a boolean it will be used to set the check (true) or nocheck (false)
// the two parameter can be provide together
func (p *Paragraph) LangCheck(check ...any) *Paragraph {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}
	if p.Properties.RunProperties == nil {
		p.Properties.RunProperties = &RunProperties{}
	}

	proof := true
	lang := ""
	for _, c := range check {
		switch v := c.(type) {
		case bool:
			proof = v
		case string:
			lang = v
		}
	}
	if lang != "" {
		p.Properties.RunProperties.Lang = &Lang{Val: lang}
	} else {
		p.Properties.RunProperties.Lang = nil
	}
	if proof {
		p.Properties.RunProperties.NoProof = nil
	} else {
		p.Properties.RunProperties.NoProof = &NoProof{}
	}

	for _, child := range p.Children {
		switch r := child.(type) {
		case *Run:
			r.LangCheck(check...)
		}
	}
	return p
}

func (p *Paragraph) KeepLines(val ...bool) *Paragraph {
	if len(val) == 0 || val[0] {
		if p.Properties == nil {
			p.Properties = &ParagraphProperties{}
		}
		p.Properties.KeepLines = &KeepLines{}
	} else {
		if p.Properties != nil {
			p.Properties.KeepLines = nil
		}
	}
	return p
}

func (p *Paragraph) KeepNext(val ...bool) *Paragraph {
	if len(val) == 0 || val[0] {
		if p.Properties == nil {
			p.Properties = &ParagraphProperties{}
		}
		p.Properties.KeepNext = &KeepNext{}
	} else {
		if p.Properties != nil {
			p.Properties.KeepNext = nil
		}
	}
	return p
}

func (p *Paragraph) PageBreakBefore(val ...bool) *Paragraph {
	if len(val) == 0 || val[0] {
		if p.Properties == nil {
			p.Properties = &ParagraphProperties{}
		}
		p.Properties.PageBreakBefore = &PageBreakBefore{}
	} else {
		if p.Properties != nil {
			p.Properties.PageBreakBefore = nil
		}
	}
	return p
}

func (p *Paragraph) SuppressAutoHyphens(val ...bool) *Paragraph {
	if len(val) == 0 || val[0] {
		if p.Properties == nil {
			p.Properties = &ParagraphProperties{}
		}
		p.Properties.SuppressAutoHyphens = &SuppressAutoHyphens{}
	} else {
		if p.Properties != nil {
			p.Properties.SuppressAutoHyphens = nil
		}
	}
	return p
}
