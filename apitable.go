/*
   Copyright (c) 2020 gingfrederik
   Copyright (c) 2021 Gonzalo Fernandez-Victorio
   Copyright (c) 2021 Basement Crowd Ltd (https://www.basementcrowd.com)
   Copyright (c) 2023 Fumiama Minamoto (源文雨)
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
	"fmt"
	"reflect"
	"slices"
)

// AddTable add a new table to body by col*row
//
// unit: twips (1/20 point)
func (f *Docx) AddTable(
	row int,
	col int,
	tableWidth int64,
	borderColors *APITableBorderColors,
) *Table {
	trs := make([]*WTableRow, row)
	for i := 0; i < row; i++ {
		cells := make([]*WTableCell, col)
		for i := range cells {
			cells[i] = &WTableCell{
				Properties: &WTableCellProperties{
					Width: &WTableCellWidth{Type: "auto"},
				},
				file: f,
			}
		}
		trs[i] = &WTableRow{
			Properties: &WTableRowProperties{},
			Cells:      cells,
		}
	}

	if borderColors == nil {
		borderColors = new(APITableBorderColors)
	}
	borderColors.applyDefault()

	tbl := &Table{
		Properties: &WTableProperties{
			Look: &WTableLook{
				Val: "0000",
			},
		},
		Grid: &WTableGrid{},
		Rows: trs,
	}

	tbl.Width(tableWidth)

	tbl.Style("TableGrid", 0)

	f.Document.Body.Items = append(f.Document.Body.Items, tbl)
	return tbl
}

func (f *Docx) AddTableEmpty() *Table {
	tbl := &Table{
		Properties: &WTableProperties{
			Look: &WTableLook{
				Val: "0000",
			},
		},
		Grid: &WTableGrid{},
		Rows: make([]*WTableRow, 0),
	}

	tbl.Style("TableGrid", 0)

	f.Document.Body.Items = append(f.Document.Body.Items, tbl)
	return tbl
}

// AddTableTwips add a new table to body by height and width
//
// unit: twips (1/20 point)
func (f *Docx) AddTableTwips(
	rowHeights []int64,
	colWidths []int64,
	tableWidth int64,
	borderColors *APITableBorderColors,
) *Table {
	grids := make([]*WGridCol, len(colWidths))
	trs := make([]*WTableRow, len(rowHeights))
	var total int64 = 0
	for i, w := range colWidths {
		if w > 0 {
			total += w
			grids[i] = &WGridCol{
				W: w,
			}
		}
	}
	if tableWidth <= 0 {
		tableWidth = total
	}

	for i, h := range rowHeights {
		cells := make([]*WTableCell, len(colWidths))
		for i, w := range colWidths {
			cells[i] = &WTableCell{
				Properties: &WTableCellProperties{
					Width: &WTableCellWidth{W: w, Type: "dxa"},
				},
				file: f,
			}
		}
		trs[i] = &WTableRow{
			Properties: &WTableRowProperties{},
			Cells:      cells,
		}
		if h > 0 {
			trs[i].Properties.Height = &WTableRowHeight{
				Val: h,
			}
		}
	}

	if borderColors == nil {
		borderColors = new(APITableBorderColors)
	}
	borderColors.applyDefault()

	tbl := &Table{
		Properties: &WTableProperties{
			Look: &WTableLook{
				Val: "0000",
			},
		},
		Grid: &WTableGrid{
			GridCols: grids,
		},
		Rows: trs,
	}

	tbl.Style("TableGrid", 0)

	tbl.Width(tableWidth)

	f.Document.Body.Items = append(f.Document.Body.Items, tbl)
	return tbl
}

type t_TABLE_BORDER int

const (
	TABLE_BORDER_TOP     t_TABLE_BORDER = 1  // Top border of the table
	TABLE_BORDER_LEFT    t_TABLE_BORDER = 2  // Left border of the table
	TABLE_BORDER_BOTTOM  t_TABLE_BORDER = 4  // Bottom border of the table
	TABLE_BORDER_RIGHT   t_TABLE_BORDER = 8  // Right border of the table
	TABLE_BORDER_INSIDEH t_TABLE_BORDER = 16 // Inside horizontal cell border
	TABLE_BORDER_INSIDEV t_TABLE_BORDER = 32 // Inside vertical cell border
	TABLE_BORDER_EXTERN  t_TABLE_BORDER = 15 // All external borders of the table
	TABLE_BORDER_INSIDE  t_TABLE_BORDER = 48 // Horizontal and vertical inside cell borders
	TABLE_BORDER_ALL     t_TABLE_BORDER = 63 // Inside and external borders
)

var v_border_list []t_TABLE_BORDER = []t_TABLE_BORDER{
	TABLE_BORDER_TOP, TABLE_BORDER_LEFT, TABLE_BORDER_BOTTOM,
	TABLE_BORDER_RIGHT, TABLE_BORDER_INSIDEH, TABLE_BORDER_INSIDEV,
}

// Width allows to set width of the table
func (t *Table) Borders(which t_TABLE_BORDER, border, color string, size, space int) *Table {
	if t.Properties.Borders == nil {
		t.Properties.Borders = &WTableBorders{}
	}
	g := t.Properties.Borders
	b := &WTableBorder{Val: border, Size: size, Space: space, Color: color}
	for _, f := range v_border_list {
		switch which & f {
		case TABLE_BORDER_TOP:
			g.Top = b
		case TABLE_BORDER_LEFT:
			g.Left = b
		case TABLE_BORDER_BOTTOM:
			g.Bottom = b
		case TABLE_BORDER_RIGHT:
			g.Right = b
		case TABLE_BORDER_INSIDEH:
			g.InsideH = b
		case TABLE_BORDER_INSIDEV:
			g.InsideV = b
		}
	}
	return t
}

// Width allows to set width of the table
func (t *Table) Width(w int64) *Table {
	typ := "auto"
	if w > 0 {
		typ = "dxa"
	}
	t.Properties.Width = &WTableWidth{
		W:    w,
		Type: typ,
	}
	return t
}

type t_TABLE_STYLE_OPTION int

const (
	TABLE_STYLE_OPTION_FIRST_ROW       t_TABLE_STYLE_OPTION = 0x0020
	TABLE_STYLE_OPTION_LAST_ROW        t_TABLE_STYLE_OPTION = 0x0040
	TABLE_STYLE_OPTION_FIRST_COLUMN    t_TABLE_STYLE_OPTION = 0x0080
	TABLE_STYLE_OPTION_LAST_COLUMN     t_TABLE_STYLE_OPTION = 0x0100
	TABLE_STYLE_OPTION_HORIZONTAL_BAND t_TABLE_STYLE_OPTION = 0x0200
	TABLE_STYLE_OPTION_VERTICAL_BAND   t_TABLE_STYLE_OPTION = 0x0400
)

var v_TABLE_STYLE_OPT []t_TABLE_STYLE_OPTION = []t_TABLE_STYLE_OPTION{
	TABLE_STYLE_OPTION_FIRST_ROW, TABLE_STYLE_OPTION_LAST_ROW, TABLE_STYLE_OPTION_FIRST_COLUMN,
	TABLE_STYLE_OPTION_LAST_COLUMN, TABLE_STYLE_OPTION_HORIZONTAL_BAND, TABLE_STYLE_OPTION_VERTICAL_BAND,
}

func (t *Table) setConfStyle() {
	g := t.Properties.Look
	if g.FirstRow == 1 {
		t.confStyle.firstRow = &WTableConfStyle{
			Val:      "100000000000",
			FirstRow: 1,
		}
	}
	if g.LastRow == 1 {
		t.confStyle.lastRow = &WTableConfStyle{
			Val:     "010000000000",
			LastRow: 1,
		}
	}
	if g.FirstCol == 1 {
		t.confStyle.firstCol = &WTableConfStyle{
			Val:      "001000000000",
			FirstCol: 1,
		}
	}
	if g.LastCol == 1 {
		t.confStyle.lastCol = &WTableConfStyle{
			Val:     "000100000000",
			LastCol: 1,
		}
	}
	if g.NoHBand == 0 {
		t.confStyle.oddHBand = &WTableConfStyle{
			Val:      "000000100000",
			OddHBand: 1,
		}
	}
	if g.NoVBand == 0 {
		t.confStyle.oddVBand = &WTableConfStyle{
			Val:      "000010000000",
			OddVBand: 1,
		}
	}
	t.confStyle.none = &WTableConfStyle{
		Val: "000000000000",
	}
}

// Style allows to set table table style
func (t *Table) Style(style string, option t_TABLE_STYLE_OPTION) *Table {
	t.Properties.Style = &WTableStyle{
		Val: style,
	}
	b := option ^ (TABLE_STYLE_OPTION_HORIZONTAL_BAND | TABLE_STYLE_OPTION_VERTICAL_BAND)
	t.Properties.Look = &WTableLook{
		Val:     fmt.Sprintf("%04X", b),
		NoHBand: 1,
		NoVBand: 1,
	}
	g := t.Properties.Look
	for _, opt := range v_TABLE_STYLE_OPT {
		switch option & opt {
		case TABLE_STYLE_OPTION_FIRST_ROW:
			g.FirstRow = 1
		case TABLE_STYLE_OPTION_FIRST_COLUMN:
			g.FirstCol = 1
		case TABLE_STYLE_OPTION_LAST_ROW:
			g.LastRow = 1
		case TABLE_STYLE_OPTION_LAST_COLUMN:
			g.LastCol = 1
		case TABLE_STYLE_OPTION_HORIZONTAL_BAND:
			g.NoHBand = 0
		case TABLE_STYLE_OPTION_VERTICAL_BAND:
			g.NoVBand = 0
		}
	}

	t.setConfStyle()
	return t
}

// ColGrid allows to set cols width
// length of w must be equals to the highest number of columns in the rows of the table
// Any Merge must be applied before in order to be taken into account
func (t *Table) ColGrid(w []int64) *Table {
	var g []*WGridCol = make([]*WGridCol, len(w))
	var total int64 = 0

	for i, v := range w {
		total += v
		g[i] = &WGridCol{
			W: v,
		}
	}
	t.Grid = &WTableGrid{
		GridCols: g,
	}
	t.Width(total)

	lw := len(w)

	for j := range t.Rows {
		i := 0
		for _, c := range t.Rows[j].Cells {
			if i < lw {
				v := w[i]
				if c.Properties != nil && c.Properties.GridSpan != nil {
					s := i + c.Properties.GridSpan.Val
					for v = 0; i < lw && i < s; i++ {
						v += w[i]
					}
					i--
				}
				c.Width(v)
				i++
			}
		}
	}

	return t
}

// Justification allows to set table's horizonal alignment
//
//	w:jc possible values：
//		start
//		center
//		end
//		both： justify
//		distribute： disperse Alignment
func (t *Table) Justification(val string) *Table {
	if t.Properties.Justification == nil {
		t.Properties.Justification = &Justification{Val: val}
		return t
	}
	t.Properties.Justification.Val = val
	return t
}

// Append/Insert a row in a table
// append if position is < 0
func (t *Table) AddRow(position int) *WTableRow {
	v := &WTableRow{
		Properties: &WTableRowProperties{},
	}
	if t.Rows != nil && position >= 0 && position < len(t.Rows) {
		t.Rows = slices.Insert(t.Rows, position, v)
	} else {
		t.Rows = append(t.Rows, v)
	}
	return v
}

// Justification allows to set table's horizonal alignment
//
//	w:jc possible values：
//		start
//		center
//		end
//		both： justify
//		distribute： disperse Alignment
func (w *WTableRow) Justification(val string) *WTableRow {
	if w.Properties.Justification == nil {
		w.Properties.Justification = &Justification{Val: val}
		return w
	}
	w.Properties.Justification.Val = val
	return w
}

// Append/Insert a cell in a table row
// append if position is < 0
func (r *WTableRow) AddCell(position int) *WTableCell {
	v := &WTableCell{
		Properties: &WTableCellProperties{},
	}
	if r.Cells != nil && position >= 0 && position < len(r.Cells) {
		r.Cells = slices.Insert(r.Cells, position, v)
	} else {
		r.Cells = append(r.Cells, v)
	}
	return v
}

// Merge allows to merge a cells rectangle in th table
// Do not apply Merge and table style with vertical band together. The result is unknown
// Do not insert row or cell in a merged aera. The result is unknown
func (t *Table) Merge(firstRow, firstCol, lastRow, lastCol int) *Table {
	if t.Rows != nil && 0 <= firstRow && firstRow < lastRow && lastRow < len(t.Rows) && 0 <= firstCol && firstCol < lastCol {
		starts := make([]int, 0)
		ends := make([]int, 0)
		for i := firstRow; i <= lastRow; i++ {
			cols := 0
			start := -1
			end := -1
			for k, a := range t.Rows[i].Cells {
				if a.Properties != nil && a.Properties.GridSpan != nil {
					cols += a.Properties.GridSpan.Val
				} else {
					cols++
				}
				if cols == firstCol {
					start = k
				}
				if cols == lastCol {
					end = k + 1
				}
			}
			if end == -1 || start == -1 {
				// an existing span is intersecting with this one or not enough cells
				// then ignore
				return t
			}
			starts = append(starts, start)
			ends = append(ends, end)
		}
		// everything is fine then apply
		for i := firstRow; i <= lastRow; i++ {
			r := t.Rows[i]
			p := r.Cells[starts[i-firstRow]].Properties
			if starts[i-firstRow] < ends[i-firstRow] {
				// merge horizontally
				if ends[i-firstRow] == len(r.Cells) {
					r.Cells = r.Cells[:starts[i-firstRow]+1]
				} else {
					r.Cells = append(r.Cells[:starts[i-firstRow]+1], r.Cells[:ends[i-firstRow]]...)
				}
				p.GridSpan = &WGridSpan{
					Val: lastCol - firstCol + 1,
				}
			}
			if firstRow < lastRow {
				// merge vertically
				p.VMerge = &WvMerge{}
				if i == firstRow {
					p.VMerge.Val = "restart"
				}
			}
		}
	}
	return t
}

// Shade allows to set cell's shade
func (c *WTableCell) Shade(val, color, fill string) *WTableCell {
	c.Properties.Shade = &Shade{
		Val:   val,
		Color: color,
		Fill:  fill,
	}
	return c
}

// Width allows to set width of the cell
func (c *WTableCell) Width(w int64) *WTableCell {
	typ := "auto"
	if w > 0 {
		typ = "dxa"
	}
	c.Properties.Width = &WTableCellWidth{
		W:    w,
		Type: typ,
	}
	return c
}

// APITableBorderColors customizable param
type APITableBorderColors struct {
	Top     string
	Left    string
	Bottom  string
	Right   string
	InsideH string
	InsideV string
}

func (tbc *APITableBorderColors) applyDefault() {
	tbcR := reflect.ValueOf(tbc).Elem()

	for i := 0; i < tbcR.NumField(); i++ {
		if tbcR.Field(i).IsZero() {
			tbcR.Field(i).SetString("#000000")
		}
	}
}
