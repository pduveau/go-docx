# Docx library

One of the most functional libraries to read and write .docx (a.k.a. Microsoft Word documents or ECMA-376 Office Open XML) files in Go.

This is a variant optimized and expanded by pduveau. The original repo is [gonfva/docxlib](https://github.com/gonfva/docxlib).

## Introduction

> As part of my work for [Basement Crowd](https://www.basementcrowd.com) and [FromCounsel](https://www.fromcounsel.com), we were in need of a basic library to manipulate (both read and write) Microsoft Word documents.
> 
> The difference with other projects is the following:
> - [UniOffice](https://github.com/unidoc/unioffice) is probably the most complete but it is also commercial (you need to pay). It also very complete, but too much for my needs.
> - [gingfrederik/docx](https://github.com/gingfrederik/docx) only allows to write.
> 
> There are also a couple of other projects [kingzbauer/docx](https://github.com/kingzbauer/docx) and [nguyenthenguyen/docx](https://github.com/nguyenthenguyen/docx)
> 
> [gingfrederik/docx](https://github.com/gingfrederik/docx) was a heavy influence (the original structures and the main method come from that project).
> 
> However, those original structures didn't handle reading and extending them was particularly difficult due to Go xml parser being a bit limited including a [6 year old bug](https://github.com/golang/go/issues/9519).
> 
> Additionally, my requirements go beyond the original structure and a hard fork seemed more sensible.
> 
> The plan is to evolve the library, so the API is likely to change according to my company's needs. But please do feel free to send patches, reports and PRs (or fork).
> 
> In the mean time, shared as an example in case somebody finds it useful.

The Introduction above is copied from the original repo. I had evolved that repo again to fit my needs. Here are the supported functions now.

- [x] Parse and save document
- [x] Edit text (color, size, alignment, link, ...)
- [x] Edit picture
- [x] Edit table
- [x] Edit shape
- [x] Edit canvas
- [x] Edit group

## Quick Start
```bash
go run cmd/main/main.go -u
```
And you will see two files generated under `pwd` with the same contents as below.

<table>
	<tr>
		<td align="center"><img src="https://user-images.githubusercontent.com/41315874/223348099-4a6099d2-0fec-4e13-92a7-152c00bc6f6b.png"></td>
		<td align="center"><img src="https://user-images.githubusercontent.com/41315874/223349486-e78ac0f1-c879-4888-9110-ea4db2590241.png"></td>
	</tr>
	<tr>
		<td align="center">p1</td>
		<td align="center">p2</td>
	</tr>
</table>

## Use Package in your Project
```bash
go get -d github.com/pduveau/go-docx@latest
```
### Generate Document
```go
package main

import (
	"os"
	"strings"

	"github.com/pduveau/go-docx"
)

func main() {
	w := docx.New().WithDefaultTheme().WithA4Page()
	// add new paragraph
	para1 := w.AddParagraph()
	// add text
	para1.AddText("test").AddTab()
	para1.AddText("size").Size("44").AddTab()

	// save to file
	f, err := os.Create("generated.docx")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	_, err = w.WriteTo(f)
	if err != nil {
		panic(err)
	}
}
```
### Parse Document
```go
package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/pduveau/go-docx"
)

func main() {
	doc, err := docx.ReadDocument("file2parse.docx")
	if err != nil {
		panic(err)
	}
	fmt.Println("Plain text:")
	for _, it := range doc.Document.Body.Items {
		switch it.(type) {
		case *docx.Paragraph, *docx.Table: // printable
			fmt.Println(it)
		}
	}
}
```
### Generate Document with styles and tables
```go
package main

import (
	"os"
	"strings"

	"github.com/pduveau/go-docx"
)

func main() {
	w := docx.New().WithDefaultTheme().WithA4Page()

	// ada a paragraph style
	myStylePara := &docx.StyleStyle{
		Type:        "paragraph",
		CustomStyle: "1",
		StyleId:     "myStylePara",
		Name:        &docx.StyleVal{Val: "My Style Para"},
		QFormat:     &docx.StyleVal{},
		PPr: &docx.StylePPr{
			Spacing: &docx.StyleSpacing{
				After:    "0",
				Line:     "240",
				LineRule: "auto",
			},
		},
		RPr: &docx.StyleRPr{
			RFonts: &docx.StyleRFonts{
				Ascii: "Calibri",
				HAnsi: "Calibri",
			},
			Sz: &docx.StyleVal{Val: "22"},
		},
	}

	w.AddOrReplaceStyle(myStylePara)

	myTableBorder := &docx.StyleBorder{
		Val:   "single",
		Sz:    "4",
		Space: "0",
		Color: "auto",
	}

	myTableCellMargin := &docx.StyleCellMar{
		W:    "57",
		Type: "dxa",
	}

	myStyleTable := &docx.StyleStyle{
		Type:        "table",
		CustomStyle: "1",
		StyleId:     "myStyleTable",
		Name:        &docx.StyleVal{Val: "My Style Table"},
		BasedOn:     &docx.StyleVal{Val: "TableauNormal"},
		UiPriority:  &docx.StyleVal{Val: "99"},
		PPr: &docx.StylePPr{
			Spacing: &docx.StyleSpacing{
				After:    "0",
				Line:     "240",
				LineRule: "auto",
			},
		},
		RPr: &docx.StyleRPr{
			RFonts: &docx.StyleRFonts{
				Ascii: "Calibri",
				HAnsi: "Calibri",
			},
			Sz: &docx.StyleVal{Val: "22"},
		},
		TblPr: &docx.StyleTblPr{
			TblStyleRowBandSize: &docx.StyleVal{Val: "1"},
			TblStyleColBandSize: &docx.StyleVal{Val: "1"},
			TblBorders: &docx.StyleTblBorders{
				Top:     myTableBorder,
				Bottom:  myTableBorder,
				Left:    myTableBorder,
				Right:   myTableBorder,
				InsideH: myTableBorder,
				InsideV: myTableBorder,
			},
			TblCellMar: &docx.StyleTblCellMar{
				Left:  myTableCellMargin,
				Right: myTableCellMargin,
			},
		},
		TcPr: &docx.StyleTcPr{
			Shd: &docx.StyleShd{
				Val:   "clear",
				Color: "auto",
				Fill:  "auto",
			},
			VAlign: &docx.StyleVal{Val: "center"},
		},
		TblStylePr: []*docx.StyleTblStylePr{
			&docx.StyleTblStylePr{
				Type: "firstCol",
				RPr: &docx.StyleRPr{
					B: &docx.StyleVal{},
				},
			},
			&docx.StyleTblStylePr{
				Type: "band1Vert",
				PPr: &docx.StylePPr{
					Jc: &docx.StyleVal{Val: "left"},
				},
			},
			&docx.StyleTblStylePr{
				Type: "band1Horz",
				TcPr: &docx.StyleTcPr{
					Shd: &docx.StyleShd{
						Val:            "clear",
						Color:          "auto",
						Fill:           "D9D9D9",
						ThemeFill:      "background1",
						ThemeFillShade: "D9",
					},
				},
			},
		},
	}

	w.AddOrReplaceStyle(myTableBorder)

	tab := w.AddTableEmpty()
	tab.Style("myStyleTable", docx.TABLE_STYLE_OPTION_HORIZONTAL_BAND|docx.TABLE_STYLE_OPTION_FIRST_COLUMN)

	row := tab.AddRow()

	// KeepNext enforce no page break between the two rows
	row.AddCell().AddParagraph().Style("myStylePara").KeepNext().LangCheck(false).AddText("fr-FR")

	row.AddCell().AddParagraph().Style("myStylePara").KeepNext().LangCheck("fr-FR").AddText("Il faut beau.").Italic(true)

	row = tab.AddRow()

	row.AddCell().AddParagraph().Style("myStylePara").LangCheck(false).AddText("en-GB")

	// KeepLine enforce no page break in the middle of the cell as the \n change line without changing paragraph (maj+enter)
	para = row.AddCell().AddParagraph().Style("myStylePara").LangCheck("en-GB").KeepLines()
	para.AddText("Weather is nice\n").Bold(true)
	para.AddText("but not for a long time.").Underline("single")

	// save to file
	err := w.WriteDocument("newdoc.docx")
	if err != nil {
		panic(err)
	}
}
```

## License

AGPL-3.0. See [LICENSE](LICENSE)
