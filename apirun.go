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

// Color allows to set run color
func (r *Run) Color(color string) *Run {
	r.RunProperties.Color = &Color{
		Val: color,
	}
	return r
}

// Size allows to set run size
func (r *Run) Size(size string) *Run {
	r.RunProperties.Size = &Size{
		Val: size,
	}
	return r
}

// SizeCs allows to set run sizecs
func (r *Run) SizeCs(size string) *Run {
	r.RunProperties.SizeCs = &SizeCs{
		Val: size,
	}
	return r
}

// Shade allows to set run shade
func (r *Run) Shade(val, color, fill string) *Run {
	r.RunProperties.Shade = &Shade{
		Val:   val,
		Color: color,
		Fill:  fill,
	}
	return r
}

// Spacing allows to set run spacing
func (r *Run) Spacing(line int) *Run {
	r.RunProperties.Spacing = &Spacing{
		Line: line,
	}
	return r
}

// Bold ...
func (r *Run) Bold(val ...bool) *Run {
	if len(val) == 0 || val[0] {
		r.RunProperties.Bold = &Bold{}
	} else {
		r.RunProperties.Bold = nil
	}
	return r
}

// Italic ...
func (r *Run) Italic(val ...bool) *Run {
	if len(val) == 0 || val[0] {
		r.RunProperties.Italic = &Italic{}
	} else {
		r.RunProperties.Italic = nil
	}
	return r
}

type _underline string

const (
	UNDERLINE_NONE       _underline = "none"       // Specifies that no underline should be applied.
	UNDERLINE_SINGLE     _underline = "single"     // Specifies a single underline.
	UNDERLINE_WORDS      _underline = "words"      // Specifies that only words within the text should be underlined.
	UNDERLINE_DOUBLE     _underline = "double"     // Specifies a double underline.
	UNDERLINE_THICK      _underline = "thick"      // Specifies a thick underline.
	UNDERLINE_DOTTED     _underline = "dotted"     // Specifies a dotted underline.
	UNDERLINE_DASH       _underline = "dash"       // Specifies a dash underline.
	UNDERLINE_DOTDASH    _underline = "dotDash"    // Specifies an alternating dot-dash underline.
	UNDERLINE_DOTDOTDASH _underline = "dotDotDash" // Specifies an alternating dot-dot-dash underline.
	UNDERLINE_WAVE       _underline = "wave"       // Specifies a wavy underline.
	UNDERLINE_DASHLONG   _underline = "dashLong"   // Specifies a long dash underline.
	UNDERLINE_WAVYDOUBLE _underline = "wavyDouble" // Specifies a double wavy underline.

)

// Underline has several possible values including
func (r *Run) Underline(val _underline) *Run {
	r.RunProperties.Underline = &Underline{Val: (string)(val)}
	return r
}

func (r *Run) UnderlineSingle(val ...bool) *Run {
	if len(val) == 0 || val[0] {
		r.RunProperties.Underline = &Underline{Val: "single"}
	} else {
		r.RunProperties.Underline = nil
	}
	return r
}

// Highlight ...
func (r *Run) Highlight(val string) *Run {
	r.RunProperties.Highlight = &Highlight{Val: val}
	return r
}

// Strike ...
func (r *Run) Strike(val ...bool) *Run {
	trueFalseStr := "false"
	if len(val) == 0 || val[0] {
		trueFalseStr = "true"
	}
	r.RunProperties.Strike = &Strike{Val: trueFalseStr}
	return r
}

// AddTab add a tab in front of the run
func (r *Run) AddTab() *Run {
	r.Children = append(r.Children, &Tab{})
	return r
}

// Font sets the font of the run
func (r *Run) Font(ascii, eastAsia, hansi, hint string) *Run {
	r.RunProperties.Fonts = &RunFonts{
		ASCII:    ascii,
		EastAsia: eastAsia,
		HAnsi:    hansi,
		Hint:     hint,
	}
	return r
}

// LangCheck set the language parameter
// if a parameter is a string it will be used as the language
// if a parameter is a boolean it will be used to set the check (true) or nocheck (false)
// the two parameter can be provide together
func (r *Run) LangCheck(check ...any) {
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
		r.RunProperties.Lang = &Lang{Val: lang}
	} else {
		r.RunProperties.Lang = nil
	}
	if proof {
		r.RunProperties.NoProof = nil
	} else {
		r.RunProperties.NoProof = &NoProof{}
	}
}
