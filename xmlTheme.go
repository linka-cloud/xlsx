package xlsx

import "encoding/xml"

// XLSXTheme directly maps the theme element in the namespace
// http://schemas.openxmlformats.org/drawingml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXTheme struct {
	ThemeElements XLSXThemeElements `xml:"themeElements"`
}

// XLSXThemeElements directly maps the themeElements element in the namespace
// http://schemas.openxmlformats.org/drawingml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXThemeElements struct {
	ClrScheme XLSXClrScheme `xml:"clrScheme"`
}

// XLSXClrScheme directly maps the clrScheme element in the namespace
// http://schemas.openxmlformats.org/drawingml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXClrScheme struct {
	Name     string            `xml:"name,attr"`
	Children []XLSXClrSchemeEl `xml:",any"`
}

// XLSXClrScheme maps to children of the clrScheme element in the namespace
// http://schemas.openxmlformats.org/drawingml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXClrSchemeEl struct {
	XMLName xml.Name
	SysClr  *XLSXSysClr  `xml:"sysClr"`
	SrgbClr *XLSXSrgbClr `xml:"srgbClr"`
}

// XLSXSysClr directly maps the sysClr element in the namespace
// http://schemas.openxmlformats.org/drawingml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXSysClr struct {
	Val     string `xml:"val,attr"`
	LastClr string `xml:"lastClr,attr"`
}

// XLSXSrgbClr directly maps the srgbClr element in the namespace
// http://schemas.openxmlformats.org/drawingml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXSrgbClr struct {
	Val string `xml:"val,attr"`
}
