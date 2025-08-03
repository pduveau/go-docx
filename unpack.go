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
	"archive/zip"
	"encoding/xml"
	"errors"
	"io"
	"regexp"
	"strconv"
	"strings"
)

var mediaFileRe = regexp.MustCompile(`^image(.*)\.(.*)$`)
var docsFileRe = regexp.MustCompile(`^/word/[a-z]([0-9]+)\.xml$`)

// unpack receives a zip file (word documents are a zip with multiple xml inside)
// and parses the files that are relevant for us:
//
//  1. Document
//  2. Relationships
//  3. Media
//
// Then it stores all other files into tmpfslist for packing.
func unpack(zipReader *zip.Reader) (docx *Docx, err error) {
	var n uint64
	docx = new(Docx)
	docx.mediaNameIdx = make(map[string]int, 64)
	docx.picturesId = 0
	docx.tmplfs = zipReader
	docx.tmpfslst = make([]string, 0, 64)
	docx.rID = 3
	docx.picturesPrefix = "Picture "
	docx.shapesId = 0
	docx.shapesPrefix = "Shape "
	docx.imageID = 0

	for _, f := range zipReader.File {
		if f.Name == "word/_rels/document.xml.rels" {
			err = docx.parseDocRelation(f)
			if err != nil {
				return
			}
			continue
		}
		if f.Name == "word/document.xml" {
			err = docx.parseDocument(f)
			if err != nil {
				return
			}
			continue
		}
		if f.Name == "word/styles.xml" {
			err = docx.parseStyles(f)
			if err != nil {
				return
			}
			continue
		}
		if strings.HasPrefix(f.Name, MEDIA_FOLDER) {
			n, err = docx.parseMedia(f)
			if err != nil {
				return
			}
			if n > docx.imageID {
				docx.imageID = n
			}
			continue
		}

		if docsFileRe.MatchString(f.Name) {
			n, _ = strconv.ParseUint(string(mediaFileRe.FindStringSubmatch(f.Name)[1]), 10, 64)
			if n > docx.docID {
				docx.docID = n
			}
		}
		// fill remaining files into tmpfslst
		docx.tmpfslst = append(docx.tmpfslst, f.Name)
	}
	return
}

// parseDocument processes one of the relevant files, the one with the actual document
func (f *Docx) parseDocument(file *zip.File) error {
	zf, err := file.Open()
	if err != nil {
		return err
	}
	defer zf.Close()

	f.Document.XMLw = XMLNS_W
	f.Document.XMLr = XMLNS_R
	f.Document.XMLwp = XMLNS_WP
	f.Document.XMLmc = XMLNS_MC
	f.Document.XMLo = XMLNS_O
	f.Document.XMLv = XMLNS_V
	f.Document.XMLwps = XMLNS_WPS
	f.Document.XMLwpc = XMLNS_WPC
	f.Document.XMLwpg = XMLNS_WPG
	f.Document.XMLwp14 = XMLNS_WP14
	f.Document.XMLName.Space = XMLNS_W
	f.Document.XMLName.Local = "document"

	f.Document.Body.file = f
	err = xml.NewDecoder(zf).Decode(&f.Document)
	return err
}

// parseDocument processes one of the relevant files, the one with the actual document
func (f *Docx) parseStyles(file *zip.File) error {
	zf, err := file.Open()
	if err != nil {
		return err
	}
	defer zf.Close()

	err = xml.NewDecoder(zf).Decode(&f.styles)
	return err
}

// parseDocRelation processes one of the relevant files, the one with the relationships
func (f *Docx) parseDocRelation(file *zip.File) error {
	zf, err := file.Open()
	if err != nil {
		return err
	}
	defer zf.Close()

	f.docRelation.Xmlns = XMLNS_R
	err = xml.NewDecoder(zf).Decode(&f.docRelation)
	if err != nil {
		return err
	}
	for _, r := range f.docRelation.Relationship {
		if !strings.HasPrefix(r.ID, "rId") {
			return errors.New("invalid rel ID: " + r.ID)
		}
		id, err := strconv.ParseUint(r.ID[3:], 10, 64)
		if err != nil {
			return err
		}
		if f.rID < uintptr(id) {
			f.rID = uintptr(id)
		}
	}
	return nil
}

// parseMedia add the media into Docx struct
func (f *Docx) parseMedia(file *zip.File) (id uint64, err error) {
	var zf io.ReadCloser
	name := file.Name[len(MEDIA_FOLDER):]
	zf, err = file.Open()
	if err != nil {
		return
	}
	data, err := io.ReadAll(zf)
	if err != nil {
		return
	}
	f.mediaNameIdx[name] = len(f.media)
	if mediaFileRe.MatchString(name) {
		id, _ = strconv.ParseUint(string(mediaFileRe.FindStringSubmatch(name)[1]), 10, 64)
	}
	f.media = append(f.media, Media{Name: name, Data: data})
	err = zf.Close()
	return
}
