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

import "strings"

// AddTab adds tab to para
func (p *Paragraph) AddTab() *Run {
	c := make([]interface{}, 1, 64)
	c[0] = &Tab{}

	run := &Run{
		RunProperties: &RunProperties{},
		Children:      c,
	}

	p.Children = append(p.Children, run)

	return run
}

// AddText adds text to paragraph with \n
func (p *Paragraph) AddText(text string) *Run {
	if text == "\t" {
		return p.AddTab()
	}
	c := make([]interface{}, 0, 64)
	for i, s := range strings.Split(text, "\n") {
		if i > 0 {
			c = append(c, &BarterRabbet{})
		}
		for tabIndex, k := range strings.Split(s, "\t") {
			if tabIndex > 0 {
				c = append(c, &Tab{})
			}
			if k != "" {
				c = append(c, &Text{
					Text: k,
				})
			}
		}
	}
	run := &Run{
		RunProperties: &RunProperties{},
		Children:      c,
	}
	if p.Properties != nil && p.Properties.RunProperties != nil {
		if p.Properties.RunProperties.Lang != nil {
			tmp := *(p.Properties.RunProperties.Lang)
			run.RunProperties.Lang = &tmp
		}
		if p.Properties.RunProperties.NoProof != nil {
			tmp := *(p.Properties.RunProperties.NoProof)
			run.RunProperties.NoProof = &tmp
		}
	}
	p.Children = append(p.Children, run)
	return run
}
