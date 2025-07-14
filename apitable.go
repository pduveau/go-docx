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
	"slices"
)

// AddTable add a new table to body by col*row
//
// unit: twips (1/20 point)
func (f *Docx) AddTable(
	row int,
	col int,
	tableWidth int,
) *Table {
	tbl := &Table{
		Properties: &WTableProperties{
			Look: &WTableLook{
				Val: "0000",
			},
		},
		Grid: &WTableGrid{},
		Rows: make([]*WTableRow, row),
	}

	for i := range tbl.Rows {
		tbl.Rows[i] = &WTableRow{
			Properties: &WTableRowProperties{},
			Cells:      make([]*WTableCell, col),
			table:      tbl,
		}
		for j := range tbl.Rows[i].Cells {
			tbl.Rows[i].Cells[j] = &WTableCell{
				Properties: &WTableCellProperties{
					Width: &WTableCellWidth{Type: "auto"},
				},
				file: f,
				row:  tbl.Rows[i],
			}
		}
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
	rowHeights []int,
	colWidths []int,
	tableWidth int,
) *Table {
	tbl := &Table{
		Properties: &WTableProperties{
			Look: &WTableLook{
				Val: "0000",
			},
		},
		Grid: &WTableGrid{
			GridCols: make([]*WGridCol, len(colWidths)),
		},
		Rows: make([]*WTableRow, len(rowHeights)),
	}

	var total int = 0
	for i, w := range colWidths {
		if w > 0 {
			total += w
			tbl.Grid.GridCols[i] = &WGridCol{
				W: w,
			}
		}
	}
	if tableWidth <= 0 {
		tableWidth = total
	}

	for i, h := range rowHeights {
		tbl.Rows[i] = &WTableRow{
			Properties: &WTableRowProperties{},
			Cells:      make([]*WTableCell, len(colWidths)),
		}
		for j, w := range colWidths {
			tbl.Rows[i].Cells[j] = &WTableCell{
				Properties: &WTableCellProperties{
					Width: &WTableCellWidth{W: w, Type: "dxa"},
				},
				row:  tbl.Rows[i],
				file: f,
			}
		}
		if h > 0 {
			tbl.Rows[i].Properties.Height = &WTableRowHeight{
				Val: h,
			}
		}
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
func (t *Table) Width(w int) *Table {
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
// length of w should be equals to the highest number of columns in the rows of the table
// Warning:
// - To avoid Microsoft Word potential warning, apply this when the table is fully defined (and merged)
func (t *Table) ColGrid(w []int) *Table {
	var g []*WGridCol = make([]*WGridCol, len(w))
	var total int = 0

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
				if c.Properties.GridSpan != nil {
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

// Merge allows to merge a cells rectangle in the table
// All parameters are indexes in the table Arrays
//
// Warnings:
// - only apply merge when rows and cells are existing and definitive
// - merging a rectangle that overlap another merge will not be applied
func (t *Table) Merge(firstRow, lastRow, firstCol, lastCol int) *Table {
	if t.Rows != nil && 0 <= firstRow && lastRow < len(t.Rows) && 0 <= firstCol &&
		((firstRow <= lastRow && firstCol < lastCol) || (firstRow < lastRow && firstCol <= lastCol)) {
		starts := make([]int, firstRow)
		ends := make([]int, firstRow)
		for i := firstRow; i <= lastRow; i++ {
			cols := 0
			start := -1
			end := -1
			for k, a := range t.Rows[i].Cells {
				if cols == firstCol {
					start = k
				}
				if cols == lastCol {
					end = k + 1
				}
				if a.Properties.GridSpan != nil {
					cols += a.Properties.GridSpan.Val
				} else {
					cols++
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
			p := r.Cells[starts[i]].Properties
			if starts[i] < ends[i] {
				w := 0
				for _, c := range r.Cells[starts[i]:ends[i]] {
					if c.Properties.Width != nil {
						switch c.Properties.Width.Type {
						case "auto":
							w = -1
							break
						case "dxa":
							w += c.Properties.Width.W
						}
					}
				}
				r.Cells[starts[i]].Width(w)
				// merge horizontally
				if ends[i] == len(r.Cells) {
					r.Cells = r.Cells[:starts[i]+1]
				} else {
					r.Cells = append(r.Cells[:starts[i]+1], r.Cells[ends[i]:]...)
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

// Justification allows to set table's horizonal alignment
func (t *Table) Justification(val _justification) *Table {
	if t.Properties.Justification == nil {
		t.Properties.Justification = &Justification{}
	}
	t.Properties.Justification.Val = (string)(val)
	return t
}

// Append/Insert a row in a table
// parameter position is used to insert the row or append if not provided or position is < 0
func (t *Table) AddRow(position ...int) *WTableRow {
	v := &WTableRow{
		Properties: &WTableRowProperties{},
		Cells:      make([]*WTableCell, 0),
	}
	if len(position) > 0 && position[0] >= 0 && position[0] < len(t.Rows) {
		t.Rows = slices.Insert(t.Rows, position[0], v)
	} else {
		t.Rows = append(t.Rows, v)
	}
	return v
}

// Justification allows to set table's horizonal alignment
func (w *WTableRow) Justification(val _justification) *WTableRow {
	if w.Properties.Justification == nil {
		w.Properties.Justification = &Justification{}
	}
	w.Properties.Justification.Val = (string)(val)
	return w
}

// Append/Insert a cell in a table row
// Takes until two parameters
//   - first: insert at in the row cells array or append if not provided, >= len of cells array or < 0
//   - second: span over the provided number of columns (1 = no span)
func (r *WTableRow) AddCell(position_span ...int) *WTableCell {
	c := &WTableCell{
		Properties: &WTableCellProperties{
			Width: &WTableCellWidth{
				Type: "auto",
			},
		},
		row: r,
	}

	switch len(position_span) {
	default:
		if position_span[1] > 1 {
			c.Properties.GridSpan = &WGridSpan{
				Val: position_span[1],
			}
		}
		fallthrough
	case 1:
		if position_span[0] >= 0 && position_span[0] < len(r.Cells) {
			r.Cells = slices.Insert(r.Cells, position_span[0], c)
			return c
		}
		fallthrough
	case 0:
		r.Cells = append(r.Cells, c)
	}
	return c
}

// Width allows to set width of the table
func (c *WTableCell) Borders(which t_TABLE_BORDER, border, color string, size, space int) *WTableCell {
	if c.Properties.Borders == nil {
		c.Properties.Borders = &WTableCellBorders{}
	}
	g := c.Properties.Borders
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
		}
	}
	return c
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
func (c *WTableCell) Width(w int) *WTableCell {
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

type _valign string

const (
	TABLE_VALIGN_TOP    _valign = "top"
	TABLE_VALIGN_CENTER _valign = "center"
	TABLE_VALIGN_BOTTOM _valign = "bottom"
)

func (c *WTableCell) VAlign(val _valign) *WTableCell {
	c.Properties.VAlign = &WVerticalAlignment{
		Val: (string)(val),
	}
	return c
}
