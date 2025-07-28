/*
   Copyright (c) 2020 gingfrederik
   Copyright (c) 2021 Gonzalo Fernandez-Victorio
   Copyright (c) 2021 Basement Crowd Ltd (https://www.basementcrowd.com)
   Copyright (c) 2024 Fumiama Minamoto (源文雨)

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
	"io/fs"
)

// UseTemplate will replace template files
func (f *Docx) UseTemplate(template string, tmpfslst []string, tmplfs fs.FS) *Docx {
	f.template = template
	f.tmplfs = tmplfs
	f.tmpfslst = tmpfslst
	return f
}

// WithDefaultTheme use default theme embedded
func (f *Docx) WithDefaultTheme() *Docx {
	return f.UseTemplate("default", DefaultTemplateFilesList, TemplateXMLFS)
}

func A4_PORTRAIT() *PageSize {
	return &PageSize{W: 11906, H: 16838}
}

func A4_LANDSCAPE() *PageSize {
	return &PageSize{W: 16838, H: 11906, Orientation: "landscape"}
}

func A3_PORTRAIT() *PageSize {
	return &PageSize{W: 16838, H: 23811}
}

func A3_LANDSCAPE() *PageSize {
	return &PageSize{W: 23811, H: 16838, Orientation: "landscape"}
}

func DEFAULT_MARGIN() *PageMargin {
	return &PageMargin{Top: 1417, Left: 1417, Bottom: 1417, Right: 1417, Header: 708, Footer: 708}
}

const DEFAULT_COLS_SPACE = 708
const DEFAULT_LINEPITCH = 360

// WithA3Page use A3 PageSize
func (f *Docx) WithA3Page() *Docx {
	sectpr := &SectionProperties{
		PageSize: A3_PORTRAIT(),
	}
	f.Document.Body.Items = append(f.Document.Body.Items, sectpr)
	return f
}

// WithA4Page use A4 PageSize
func (f *Docx) WithA4Page() *Docx {
	sectpr := &SectionProperties{
		PageSize: A4_PORTRAIT(),
	}
	f.Document.Body.Items = append(f.Document.Body.Items, sectpr)
	return f
}
