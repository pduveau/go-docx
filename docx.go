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

// Package docx is one of the most functional libraries to read and write .docx
// (a.k.a. Microsoft Word documents or ECMA-376 Office Open XML) files in Go.
package docx

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"io/fs"
	"os"
	"regexp"
	"strconv"
	"sync/atomic"
)

// Docx is the structure that allow to access the internal represntation
// in memory of the doc (either read or about to be written)
type Docx struct {
	Document Document // Document is word/document.xml

	docRelation Relationships // docRelation is word/_rels/document.xml.rels

	styles *Styles

	media        []Media
	mediaNameIdx map[string]int

	rID     uintptr
	imageID uint64
	docID   uint64

	picturesPrefix string
	picturesId     uint64

	shapesPrefix string
	shapesId     uint64

	template string
	tmplfs   fs.FS
	tmpfslst []string

	io.Reader
	io.WriterTo
}

// New generates a new empty docx file that we can manipulate and
// later on, save
func New() *Docx {
	return newEmptyFile()
}

// Set the picture and shpaes prefixes in order to name them for a table of pictures
// must be set before adding drawings
func (f *Docx) SetPPrefixex(pictures, shapes string) {
	f.picturesPrefix = pictures + " "
	f.shapesPrefix = shapes + " "
}

// Parse generates a new docx file in memory from a reader
// You can it invoke from a file
//
//	readFile, err := os.Open(FILE_PATH)
//	if err != nil {
//		panic(err)
//	}
//	fileinfo, err := readFile.Stat()
//	if err != nil {
//		panic(err)
//	}
//	size := fileinfo.Size()
//	doc, err := docxlib.Parse(readFile, int64(size))
//
// but also you can invoke from a webform (BEWARE of trusting users data!!!)
//
//	func uploadFile(w http.ResponseWriter, r *http.Request) {
//		r.ParseMultipartForm(10 << 20)
//
//		file, handler, err := r.FormFile("file")
//		if err != nil {
//			fmt.Println("Error Retrieving the File")
//			fmt.Println(err)
//			http.Error(w, err.Error(), http.StatusBadRequest)
//			return
//		}
//		defer file.Close()
//		docxlib.Parse(file, handler.Size)
//	}
func Parse(reader io.ReaderAt, size int64) (doc *Docx, err error) {
	zipReader, err := zip.NewReader(reader, size)
	if err != nil {
		return nil, err
	}
	doc, err = unpack(zipReader)
	return
}

// LoadBodyItems will load body and media to a new Docx struct.
// You should call UseTemplate to set a template later.
func LoadBodyItems(items []interface{}, media []Media) *Docx {
	doc := &Docx{
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
			Body:        Body{Items: items},
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
		media:          media,
		mediaNameIdx:   make(map[string]int, 64),
		rID:            3,
		picturesId:     0,
		picturesPrefix: "Picture ",
		shapesId:       0,
		shapesPrefix:   "Shape ",
		imageID:        0,
	}
	doc.Document.Body.file = doc
	for i, m := range media {
		doc.mediaNameIdx[m.Name] = i
	}
	atomic.StoreUint64(&doc.picturesId, uint64(len(media)+1))
	return doc
}

// addImage add image to docx and return its rId
func (f *Docx) addImage(format string, data []byte) string {
	m := Media{Name: "image" + strconv.FormatUint(atomic.AddUint64(&f.imageID, 1), 10) + "." + format, Data: data}
	f.addMedia(m)
	return f.addImageRelation(m)
}

// increaseDocID
func (f *Docx) increaseDocID() (n uint64) {
	n = atomic.AddUint64(&f.docID, 1)
	return
}

// IncreasePicturesID
func (f *Docx) increasePictureID() (id uint64, name string) {
	id = atomic.AddUint64(&f.picturesId, 1)
	name = f.picturesPrefix + strconv.FormatUint(id, 10)
	return
}

// increaseShapesID
func (f *Docx) increaseShapesID() (name string) {
	id := atomic.AddUint64(&f.shapesId, 1)
	name = f.shapesPrefix + strconv.FormatUint(id, 10)
	return
}

// WriteTo allows to save a docx to a writer
func (f *Docx) WriteTo(writer io.Writer) (_ int64, err error) {
	zipWriter := zip.NewWriter(writer)
	defer zipWriter.Close()

	err = f.pack(zipWriter)
	return
}

// ReadDocument allow to load a document as a template to append
func ReadDocument(path string) (doc *Docx, err error) {
	var f *os.File
	var fi fs.FileInfo

	f, err = os.Open(path)
	if err != nil {
		return
	}

	fi, err = f.Stat()
	if err != nil {
		return
	}
	doc, err = Parse(f, int64(fi.Size()))
	return
}

func (d *Docx) WriteDocument(path string) (err error) {
	var f *os.File

	f, err = os.Create(path)
	if err == nil {
		defer f.Close()
		_, err = d.WriteTo(f)
	}
	return
}

// ClearDoc empty the document body
func (d *Docx) ClearDoc() {
	d.Document.Body.Items = make([]interface{}, 0)
}

func (f *Docx) FindItemIndex(filter *regexp.Regexp) (int, []string) {
	for i, item := range f.Document.Body.Items {
		switch p := item.(type) {
		case *Paragraph:
			s := p.Text()
			if filter.MatchString(s) {
				return i, filter.FindStringSubmatch(s)
			}
		}
	}
	return -1, []string{}
}

type Insertable interface {
	LinkToDoc(f *Docx)
}

// insert indepent item(s) at position.
// if position is < 0 or larger then the document Items array then the item(s) are appended at the end
// return the position to the next item after insertion or -1 if nothing inserted
// WARNING : once inserted it cannot be inserted in another document
func (f *Docx) InsertAt(position int, items ...interface{}) int {
	if len(items) == 0 {
		return -1
	}
	for _, item := range items {
		if v, ok := item.(Insertable); ok {
			v.LinkToDoc(f)
		}
	}
	if position < 0 || position >= len(f.Document.Body.Items) {
		f.Document.Body.Items = append(f.Document.Body.Items, items...)
		return len(f.Document.Body.Items)
	}
	var start []interface{} = make([]interface{}, position, len(f.Document.Body.Items)+len(items))

	if position > 0 {
		copy(start, f.Document.Body.Items[:position])
	}
	start = append(start, items...)
	l := len(start)
	f.Document.Body.Items = append(start, f.Document.Body.Items[position:]...)
	return l
}

// replace at position and length with indepent item(s).
// if position is < 0 or larger then the document Items array then the item(s) are appended at the end
// return the position to the next item after insertion or -1 if nothing inserted
func (f *Docx) ReplaceAt(position, length int, items ...interface{}) int {
	if len(items) == 0 {
		return -1
	}
	for _, item := range items {
		if v, ok := item.(Insertable); ok {
			v.LinkToDoc(f)
		}
	}
	if position < 0 || position >= len(f.Document.Body.Items) {
		f.Document.Body.Items = append(f.Document.Body.Items, items...)
		return len(f.Document.Body.Items)
	}

	var start []interface{} = make([]interface{}, position, len(f.Document.Body.Items)+len(items))

	if position > 0 {
		copy(start, f.Document.Body.Items[:position])
	}
	start = append(start, items...)
	l := len(start)

	if position+length < len(f.Document.Body.Items) {
		f.Document.Body.Items = append(start, f.Document.Body.Items[(position+length):]...)
	} else {
		f.Document.Body.Items = start
	}
	return l
}

// Add or replace a style in the table identified by the styleId attribut
// only one control (styleId != "") is done so be careful with this
func (d *Docx) AddOrReplaceStyle(in *StyleStyle) {
	if in.StyleId == "" {
		return
	}
	for i, t := range d.styles.Styles {
		if t.StyleId == in.StyleId {
			d.styles.Styles[i] = in
			return
		}
	}
	d.styles.Styles = append(d.styles.Styles, in)
}
