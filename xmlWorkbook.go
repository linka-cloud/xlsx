package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
)

const (
	// sheet state values as defined by
	// http://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.sheetstatevalues.aspx
	sheetStateVisible    = "visible"
	sheetStateHidden     = "hidden"
	sheetStateVeryHidden = "veryHidden"
)

// xmlxWorkbookRels contains xmlxWorkbookRelations
// which maps sheet id and sheet XML
type XLSXWorkbookRels struct {
	XMLName       xml.Name               `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []XLSXWorkbookRelation `xml:"Relationship"`
}

// xmlxWorkbookRelation maps sheet id and xl/worksheets/sheet%d.xml
type XLSXWorkbookRelation struct {
	Id     string `xml:",attr"`
	Target string `xml:",attr"`
	Type   string `xml:",attr"`
}

// XLSXWorkbook directly maps the workbook element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXWorkbook struct {
	XMLName            xml.Name               `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main workbook"`
	FileVersion        XLSXFileVersion        `xml:"fileVersion"`
	WorkbookPr         XLSXWorkbookPr         `xml:"workbookPr"`
	WorkbookProtection XLSXWorkbookProtection `xml:"workbookProtection"`
	BookViews          XLSXBookViews          `xml:"bookViews"`
	Sheets             XLSXSheets             `xml:"sheets"`
	DefinedNames       XLSXDefinedNames       `xml:"definedNames"`
	CalcPr             XLSXCalcPr             `xml:"calcPr"`
}

// XLSXWorkbookProtection directly maps the workbookProtection element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXWorkbookProtection struct {
	// We don't need this, yet.
}

// XLSXFileVersion directly maps the fileVersion element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXFileVersion struct {
	AppName      string `xml:"appName,attr,omitempty"`
	LastEdited   string `xml:"lastEdited,attr,omitempty"`
	LowestEdited string `xml:"lowestEdited,attr,omitempty"`
	RupBuild     string `xml:"rupBuild,attr,omitempty"`
}

// XLSXWorkbookPr directly maps the workbookPr element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXWorkbookPr struct {
	DefaultThemeVersion string `xml:"defaultThemeVersion,attr,omitempty"`
	BackupFile          bool   `xml:"backupFile,attr,omitempty"`
	ShowObjects         string `xml:"showObjects,attr,omitempty"`
	Date1904            bool   `xml:"date1904,attr"`
}

// XLSXBookViews directly maps the bookViews element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXBookViews struct {
	WorkBookView []XLSXWorkBookView `xml:"workbookView"`
}

// XLSXWorkBookView directly maps the workbookView element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXWorkBookView struct {
	ActiveTab            int    `xml:"activeTab,attr,omitempty"`
	FirstSheet           int    `xml:"firstSheet,attr,omitempty"`
	ShowHorizontalScroll bool   `xml:"showHorizontalScroll,attr,omitempty"`
	ShowVerticalScroll   bool   `xml:"showVerticalScroll,attr,omitempty"`
	ShowSheetTabs        bool   `xml:"showSheetTabs,attr,omitempty"`
	TabRatio             int    `xml:"tabRatio,attr,omitempty"`
	WindowHeight         int    `xml:"windowHeight,attr,omitempty"`
	WindowWidth          int    `xml:"windowWidth,attr,omitempty"`
	XWindow              string `xml:"xWindow,attr,omitempty"`
	YWindow              string `xml:"yWindow,attr,omitempty"`
}

// XLSXSheets directly maps the sheets element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXSheets struct {
	Sheet []XLSXSheet `xml:"sheet"`
}

// XLSXSheet directly maps the sheet element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXSheet struct {
	Name    string `xml:"name,attr,omitempty"`
	SheetId string `xml:"sheetId,attr,omitempty"`
	Id      string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
	State   string `xml:"state,attr,omitempty"`
}

// XLSXDefinedNames directly maps the definedNames element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXDefinedNames struct {
	DefinedName []XLSXDefinedName `xml:"definedName"`
}

// XLSXDefinedName directly maps the definedName element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
// for a descriptions of the attributes see
// https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.definedname.aspx
type XLSXDefinedName struct {
	Data              string `xml:",chardata"`
	Name              string `xml:"name,attr"`
	Comment           string `xml:"comment,attr,omitempty"`
	CustomMenu        string `xml:"customMenu,attr,omitempty"`
	Description       string `xml:"description,attr,omitempty"`
	Help              string `xml:"help,attr,omitempty"`
	ShortcutKey       string `xml:"shortcutKey,attr,omitempty"`
	StatusBar         string `xml:"statusBar,attr,omitempty"`
	LocalSheetID      int    `xml:"localSheetId,attr,omitempty"`
	FunctionGroupID   int    `xml:"functionGroupId,attr,omitempty"`
	Function          bool   `xml:"function,attr,omitempty"`
	Hidden            bool   `xml:"hidden,attr,omitempty"`
	VbProcedure       bool   `xml:"vbProcedure,attr,omitempty"`
	PublishToServer   bool   `xml:"publishToServer,attr,omitempty"`
	WorkbookParameter bool   `xml:"workbookParameter,attr,omitempty"`
	Xlm               bool   `xml:"xml,attr,omitempty"`
}

// XLSXCalcPr directly maps the calcPr element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXCalcPr struct {
	CalcId       string  `xml:"calcId,attr,omitempty"`
	IterateCount int     `xml:"iterateCount,attr,omitempty"`
	RefMode      string  `xml:"refMode,attr,omitempty"`
	Iterate      bool    `xml:"iterate,attr,omitempty"`
	IterateDelta float64 `xml:"iterateDelta,attr,omitempty"`
}

// Helper function to lookup the file corresponding to a XLSXSheet object in the worksheets map
func worksheetFileForSheet(sheet XLSXSheet, worksheets map[string]*zip.File, sheetXMLMap map[string]string) *zip.File {
	sheetName, ok := sheetXMLMap[sheet.Id]
	if !ok {
		if sheet.SheetId != "" {
			sheetName = fmt.Sprintf("sheet%s", sheet.SheetId)
		} else {
			sheetName = fmt.Sprintf("sheet%s", sheet.Id)
		}
	}
	return worksheets[sheetName]
}

// getWorksheetFromSheet() is an internal helper function to open a
// sheetN.xml file, referred to by an xlsx.XLSXSheet struct, from the XLSX
// file and unmarshal it an xlsx.XLSXWorksheet struct
func getWorksheetFromSheet(sheet XLSXSheet, worksheets map[string]*zip.File, sheetXMLMap map[string]string, rowLimit int, valueOnly bool) (*XLSXWorksheet, error) {
	var r io.Reader
	var decoder *xml.Decoder
	var worksheet *XLSXWorksheet
	var err error

	wrap := func(err error) (*XLSXWorksheet, error) {
		return nil, fmt.Errorf("getWorksheetFromSheet: %w", err)
	}

	worksheet = new(XLSXWorksheet)

	f := worksheetFileForSheet(sheet, worksheets, sheetXMLMap)
	if f == nil {
		return wrap(fmt.Errorf("Unable to find sheet '%s'", sheet))
	}
	if rc, err := f.Open(); err != nil {
		return wrap(fmt.Errorf("file.Open: %w", err))
	} else {
		defer rc.Close()
		r = rc
	}

	if rowLimit != NoRowLimit {
		r, err = truncateSheetXML(r, rowLimit)
		if err != nil {
			return wrap(err)
		}
	}

	if valueOnly {
		r, err = truncateSheetXMLValueOnly(r)
		if err != nil {
			return wrap(err)
		}
	}

	decoder = xml.NewDecoder(r)
	err = decoder.Decode(worksheet)
	if err != nil {
		return wrap(fmt.Errorf("xml.Decoder.Decode: %w", err))
	}

	worksheet.mapMergeCells()

	return worksheet, nil
}
