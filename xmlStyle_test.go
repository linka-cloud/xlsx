package xlsx

import (
	"strconv"
	"testing"

	qt "github.com/frankban/quicktest"
)

func TestIndexedColor(t *testing.T) {
	c := qt.New(t)

	colors := XLSXColors{}
	c.Run("Unitialised", func(c *qt.C) {
		c.Assert(colors.indexedColor(1), qt.Equals, "FF000000")
	})

	c.Run("Initialised", func(c *qt.C) {
		colors.IndexedColors = []XLSXRgbColor{{Rgb: "00FF00FF"}}
		c.Assert(colors.indexedColor(1), qt.Equals, "00FF00FF")
	})
}

func TestXMLStyle(t *testing.T) {
	c := qt.New(t)

	// Test we produce valid output for an empty style file.
	c.Run("MarshalEmptyXLSXStyleSheet", func(c *qt.C) {
		styles := newXLSXStyleSheet(nil)
		result, err := styles.Marshal()
		c.Assert(err, qt.IsNil)
		c.Assert(string(result), qt.Equals, `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></styleSheet>`)
	})

	// Test we produce valid output for a style file with one font definition.
	c.Run("MarshalXLSXStyleSheetWithAFont", func(c *qt.C) {
		styles := newXLSXStyleSheet(nil)
		styles.Fonts = XLSXFonts{}
		styles.Fonts.Count = 1
		styles.Fonts.Font = make([]XLSXFont, 1)
		font := XLSXFont{}
		font.Sz.Val = "10"
		font.Name.Val = "Andale Mono"
		font.B = &XLSXVal{}
		font.I = &XLSXVal{}
		font.U = &XLSXVal{}
		font.Strike = &XLSXVal{}
		styles.Fonts.Font[0] = font

		expected := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="10"/><name val="Andale Mono"/><b/><i/><u/><strike/></font></fonts></styleSheet>`
		result, err := styles.Marshal()
		c.Assert(err, qt.IsNil)
		c.Assert(string(result), qt.Equals, expected)
	})

	// Test we produce valid output for a style file with one fill definition.
	c.Run("MarshalXLSXStyleSheetWithAFill", func(c *qt.C) {
		styles := newXLSXStyleSheet(nil)
		styles.Fills = XLSXFills{}
		styles.Fills.Count = 1
		styles.Fills.Fill = make([]XLSXFill, 1)
		fill := XLSXFill{}
		patternFill := XLSXPatternFill{
			PatternType: "solid",
			FgColor:     XLSXColor{RGB: "#FFFFFF"},
			BgColor:     XLSXColor{RGB: "#000000"}}
		fill.PatternFill = patternFill
		styles.Fills.Fill[0] = fill

		expected := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fills count="1"><fill><patternFill patternType="solid"><fgColor rgb="#FFFFFF"/><bgColor rgb="#000000"/></patternFill></fill></fills></styleSheet>`
		result, err := styles.Marshal()
		c.Assert(err, qt.IsNil)
		c.Assert(string(result), qt.Equals, expected)
	})

	// Test we produce valid output for a style file with one border definition.
	// Empty elements are required to accommodate for Excel quirks.
	c.Run("MarshalXLSXStyleSheetWithABorder", func(c *qt.C) {
		styles := newXLSXStyleSheet(nil)
		styles.Borders = XLSXBorders{}
		styles.Borders.Count = 1
		styles.Borders.Border = make([]XLSXBorder, 1)
		border := XLSXBorder{}
		border.Left.Style = "solid"
		border.Top.Style = ""
		styles.Borders.Border[0] = border
		expected := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><borders count="1"><border><left style="solid"></left><right/><top/><bottom/></border></borders></styleSheet>`

		result, err := styles.Marshal()
		c.Assert(err, qt.IsNil)
		c.Assert(string(result), qt.Equals, expected)
	})

	// Test we produce valid output for a style file with one cellStyleXf definition.
	c.Run("MarshalXLSXStyleSheetWithACellStyleXf", func(c *qt.C) {
		styles := newXLSXStyleSheet(nil)
		styles.CellStyleXfs = &XLSXCellStyleXfs{}
		styles.CellStyleXfs.Count = 1
		styles.CellStyleXfs.Xf = make([]XLSXXf, 1)
		xf := XLSXXf{}
		xf.ApplyAlignment = true
		xf.ApplyBorder = true
		xf.ApplyFont = true
		xf.ApplyFill = true
		xf.ApplyProtection = true
		xf.BorderId = 0
		xf.FillId = 0
		xf.FontId = 0
		xf.NumFmtId = 0
		xf.Alignment = XLSXAlignment{
			Horizontal:   "left",
			Indent:       1,
			ShrinkToFit:  true,
			TextRotation: 0,
			Vertical:     "middle",
			WrapText:     false}
		styles.CellStyleXfs.Xf[0] = xf

		expected := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellStyleXfs count="1"><xf applyAlignment="1" applyBorder="1" applyFont="1" applyFill="1" applyNumberFormat="0" applyProtection="1" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="left" indent="1" shrinkToFit="1" textRotation="0" vertical="middle" wrapText="0"/></xf></cellStyleXfs></styleSheet>`
		result, err := styles.Marshal()
		c.Assert(err, qt.IsNil)
		c.Assert(string(result), qt.Equals, expected)
	})

	// Test we produce valid output for a style file with one cellStyle definition.
	c.Run("MarshalXLSXStyleSheetWithACellStyle", func(c *qt.C) {
		var builtInId int
		styles := newXLSXStyleSheet(nil)
		styles.CellStyles = &XLSXCellStyles{Count: 2}
		styles.CellStyles.CellStyle = make([]XLSXCellStyle, 2)

		builtInId = 31
		styles.CellStyles.CellStyle[0] = XLSXCellStyle{
			Name:      "Bob",
			BuiltInId: &builtInId, // XXX Todo - work out built-ins!
			XfId:      0,
		}
		styles.CellStyles.CellStyle[1] = XLSXCellStyle{
			Name: "Unknown",
			XfId: 1,
		}
		styles.CellStyleXfs = &XLSXCellStyleXfs{
			Count: 1,
			Xf: []XLSXXf{{}},
		}
		expected := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellStyleXfs count="1"><xf applyAlignment="0" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf></cellStyleXfs><cellStyles count="1"><cellStyle builtInId="31" name="Bob" xfId="0"></cellStyle></cellStyles></styleSheet>`
		result, err := styles.Marshal()
		c.Assert(err, qt.IsNil)
		c.Assert(string(result), qt.Equals, expected)
	})

	// Test we produce valid output for a style file with one cellXf
	// definition.
	c.Run("MarshalXLSXStyleSheetWithACellXf", func(c *qt.C) {
		styles := newXLSXStyleSheet(nil)
		styles.CellXfs = XLSXCellXfs{}
		styles.CellXfs.Count = 1
		styles.CellXfs.Xf = make([]XLSXXf, 1)
		xf := XLSXXf{}
		xf.ApplyAlignment = true
		xf.ApplyBorder = true
		xf.ApplyFont = true
		xf.ApplyFill = true
		xf.ApplyNumberFormat = true
		xf.ApplyProtection = true
		xf.BorderId = 0
		xf.FillId = 0
		xf.FontId = 0
		xf.NumFmtId = 0
		xf.Alignment = XLSXAlignment{
			Horizontal:   "left",
			Indent:       1,
			ShrinkToFit:  true,
			TextRotation: 0,
			Vertical:     "middle",
			WrapText:     false}
		styles.CellXfs.Xf[0] = xf

		expected := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="1"><xf applyAlignment="1" applyBorder="1" applyFont="1" applyFill="1" applyNumberFormat="1" applyProtection="1" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="left" indent="1" shrinkToFit="1" textRotation="0" vertical="middle" wrapText="0"/></xf></cellXfs></styleSheet>`
		result, err := styles.Marshal()
		c.Assert(err, qt.IsNil)
		c.Assert(string(result), qt.Equals, expected)
	})

	// Test we produce valid output for a style file with one NumFmt
	// definition.
	c.Run("MarshalXLSXStyleSheetWithANumFmt", func(c *qt.C) {
		styles := &XLSXStyleSheet{}
		styles.NumFmts = &XLSXNumFmts{}
		styles.NumFmts.NumFmt = make([]XLSXNumFmt, 0)
		numFmt := XLSXNumFmt{NumFmtId: 164, FormatCode: "GENERAL"}
		styles.addNumFmt(numFmt)

		expected := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="GENERAL"/></numFmts></styleSheet>`
		result, err := styles.Marshal()
		c.Assert(err, qt.IsNil)
		c.Assert(string(result), qt.Equals, expected)
	})

	c.Run("Fontqt.Equals", func(c *qt.C) {
		fontA := XLSXFont{Sz: XLSXVal{Val: "11"},
			Color:  XLSXColor{RGB: "FFFF0000"},
			Name:   XLSXVal{Val: "Calibri"},
			Family: XLSXVal{Val: "2"},
			B:      &XLSXVal{},
			I:      &XLSXVal{},
			U:      &XLSXVal{}}
		fontB := XLSXFont{Sz: XLSXVal{Val: "11"},
			Color:  XLSXColor{RGB: "FFFF0000"},
			Name:   XLSXVal{Val: "Calibri"},
			Family: XLSXVal{Val: "2"},
			B:      &XLSXVal{},
			I:      &XLSXVal{},
			U:      &XLSXVal{}}

		c.Assert(fontA.Equals(fontB), qt.Equals, true)
		fontB.Sz.Val = "12"
		c.Assert(fontA.Equals(fontB), qt.Equals, false)
		fontB.Sz.Val = "11"
		fontB.Color.RGB = "12345678"
		c.Assert(fontA.Equals(fontB), qt.Equals, false)
		fontB.Color.RGB = "FFFF0000"
		fontB.Name.Val = "Arial"
		c.Assert(fontA.Equals(fontB), qt.Equals, false)
		fontB.Name.Val = "Calibri"
		fontB.Family.Val = "1"
		c.Assert(fontA.Equals(fontB), qt.Equals, false)
		fontB.Family.Val = "2"
		fontB.B = nil
		c.Assert(fontA.Equals(fontB), qt.Equals, false)
		fontB.B = &XLSXVal{}
		fontB.I = nil
		c.Assert(fontA.Equals(fontB), qt.Equals, false)
		fontB.I = &XLSXVal{}
		fontB.U = nil
		c.Assert(fontA.Equals(fontB), qt.Equals, false)
		fontB.U = &XLSXVal{}
		// For sanity
		c.Assert(fontA.Equals(fontB), qt.Equals, true)
	})

	c.Run("FillEquals", func(c *qt.C) {
		fillA := XLSXFill{PatternFill: XLSXPatternFill{
			PatternType: "solid",
			FgColor:     XLSXColor{RGB: "FFFF0000"},
			BgColor:     XLSXColor{RGB: "0000FFFF"}}}
		fillB := XLSXFill{PatternFill: XLSXPatternFill{
			PatternType: "solid",
			FgColor:     XLSXColor{RGB: "FFFF0000"},
			BgColor:     XLSXColor{RGB: "0000FFFF"}}}
		c.Assert(fillA.Equals(fillB), qt.Equals, true)
		fillB.PatternFill.PatternType = "gray125"
		c.Assert(fillA.Equals(fillB), qt.Equals, false)
		fillB.PatternFill.PatternType = "solid"
		fillB.PatternFill.FgColor.RGB = "00FF00FF"
		c.Assert(fillA.Equals(fillB), qt.Equals, false)
		fillB.PatternFill.FgColor.RGB = "FFFF0000"
		fillB.PatternFill.BgColor.RGB = "12456789"
		c.Assert(fillA.Equals(fillB), qt.Equals, false)
		fillB.PatternFill.BgColor.RGB = "0000FFFF"
		// For sanity
		c.Assert(fillA.Equals(fillB), qt.Equals, true)
	})

	c.Run("BorderEquals", func(c *qt.C) {
		borderA := XLSXBorder{Left: XLSXLine{Style: "none"},
			Right:  XLSXLine{Style: "none"},
			Top:    XLSXLine{Style: "none"},
			Bottom: XLSXLine{Style: "none"}}
		borderB := XLSXBorder{Left: XLSXLine{Style: "none"},
			Right:  XLSXLine{Style: "none"},
			Top:    XLSXLine{Style: "none"},
			Bottom: XLSXLine{Style: "none"}}
		c.Assert(borderA.Equals(borderB), qt.Equals, true)
		borderB.Left.Style = "thin"
		c.Assert(borderA.Equals(borderB), qt.Equals, false)
		borderB.Left.Style = "none"
		borderB.Right.Style = "thin"
		c.Assert(borderA.Equals(borderB), qt.Equals, false)
		borderB.Right.Style = "none"
		borderB.Top.Style = "thin"
		c.Assert(borderA.Equals(borderB), qt.Equals, false)
		borderB.Top.Style = "none"
		borderB.Bottom.Style = "thin"
		c.Assert(borderA.Equals(borderB), qt.Equals, false)
		borderB.Bottom.Style = "none"
		// for sanity
		c.Assert(borderA.Equals(borderB), qt.Equals, true)
	})

	c.Run("XfEquals", func(c *qt.C) {
		xfA := XLSXXf{
			ApplyAlignment:  true,
			ApplyBorder:     true,
			ApplyFont:       true,
			ApplyFill:       true,
			ApplyProtection: true,
			BorderId:        0,
			FillId:          0,
			FontId:          0,
			NumFmtId:        0}
		xfB := XLSXXf{
			ApplyAlignment:  true,
			ApplyBorder:     true,
			ApplyFont:       true,
			ApplyFill:       true,
			ApplyProtection: true,
			BorderId:        0,
			FillId:          0,
			FontId:          0,
			NumFmtId:        0}
		c.Assert(xfA.Equals(xfB), qt.Equals, true)
		xfB.ApplyAlignment = false
		c.Assert(xfA.Equals(xfB), qt.Equals, false)
		xfB.ApplyAlignment = true
		xfB.ApplyBorder = false
		c.Assert(xfA.Equals(xfB), qt.Equals, false)
		xfB.ApplyBorder = true
		xfB.ApplyFont = false
		c.Assert(xfA.Equals(xfB), qt.Equals, false)
		xfB.ApplyFont = true
		xfB.ApplyFill = false
		c.Assert(xfA.Equals(xfB), qt.Equals, false)
		xfB.ApplyFill = true
		xfB.ApplyProtection = false
		c.Assert(xfA.Equals(xfB), qt.Equals, false)
		xfB.ApplyProtection = true
		xfB.BorderId = 1
		c.Assert(xfA.Equals(xfB), qt.Equals, false)
		xfB.BorderId = 0
		xfB.FillId = 1
		c.Assert(xfA.Equals(xfB), qt.Equals, false)
		xfB.FillId = 0
		xfB.FontId = 1
		c.Assert(xfA.Equals(xfB), qt.Equals, false)
		xfB.FontId = 0
		xfB.NumFmtId = 1
		c.Assert(xfA.Equals(xfB), qt.Equals, false)
		xfB.NumFmtId = 0
		// for sanity
		c.Assert(xfA.Equals(xfB), qt.Equals, true)

		var i1 int = 1

		xfA.XfId = &i1
		c.Assert(xfA.Equals(xfB), qt.Equals, false)

		xfB.XfId = &i1
		c.Assert(xfA.Equals(xfB), qt.Equals, true)

		var i2 int = 1
		xfB.XfId = &i2
		c.Assert(xfA.Equals(xfB), qt.Equals, true)

		i2 = 2
		c.Assert(xfA.Equals(xfB), qt.Equals, false)
	})

}

func TestStyleSheet(t *testing.T) {
	c := qt.New(t)

	c.Run("NewNumFmt", func(c *qt.C) {
		styles := newXLSXStyleSheet(nil)
		styles.NumFmts = &XLSXNumFmts{}
		styles.NumFmts.NumFmt = make([]XLSXNumFmt, 0)

		c.Assert(styles.newNumFmt("0"), qt.DeepEquals, XLSXNumFmt{1, "0"})
		c.Assert(styles.newNumFmt("0.00e+00"), qt.DeepEquals, XLSXNumFmt{11, "0.00e+00"})
		c.Assert(styles.newNumFmt("mm-dd-yy"), qt.DeepEquals, XLSXNumFmt{14, "mm-dd-yy"})
		c.Assert(styles.newNumFmt("hh:mm:ss"), qt.DeepEquals, XLSXNumFmt{164, "hh:mm:ss"})
		c.Assert(len(styles.NumFmts.NumFmt), qt.Equals, 1)
	})

	c.Run("AddNumFmt", func(c *qt.C) {
		styles := &XLSXStyleSheet{}
		styles.NumFmts = &XLSXNumFmts{}
		styles.NumFmts.NumFmt = make([]XLSXNumFmt, 0)

		styles.addNumFmt(XLSXNumFmt{1, "0"})
		c.Assert(styles.NumFmts.Count, qt.Equals, 0)
		styles.addNumFmt(XLSXNumFmt{14, "mm-dd-yy"})
		c.Assert(styles.NumFmts.Count, qt.Equals, 0)
		styles.addNumFmt(XLSXNumFmt{164, "hh:mm:ss"})
		c.Assert(styles.NumFmts.Count, qt.Equals, 1)
		styles.addNumFmt(XLSXNumFmt{165, "yyyy/mm/dd"})
		c.Assert(styles.NumFmts.Count, qt.Equals, 2)
		styles.addNumFmt(XLSXNumFmt{165, "yyyy/mm/dd"})
		c.Assert(styles.NumFmts.Count, qt.Equals, 2)
	})

	c.Run("GetStyle", func(c *qt.C) {
		c.Run("NoNamedStyleIndex", func(c *qt.C) {
			styles := newXLSXStyleSheet(nil)
			s0 := styles.getStyle(0)
			c.Assert(s0.NamedStyleIndex, qt.Equals, (*int)(nil))
		})
		c.Run("NamedStyleIndex", func(c *qt.C) {
			styles := newXLSXStyleSheet(nil)
			namedStyleId := 20
			csXfs := XLSXCellStyleXfs{}
			csXfs.addXf(XLSXXf{XfId: &namedStyleId})
			styles.CellStyleXfs = &csXfs
			cellStyleId := 0
			styles.CellXfs.addXf(XLSXXf{XfId: &cellStyleId})
			s0 := styles.getStyle(0)
			c.Assert(s0.NamedStyleIndex, qt.Equals, &cellStyleId)
		})

		c.Run("NamedStyleWins", func(c *qt.C) {
			styles := newXLSXStyleSheet(nil)
			namedStyleId := 20
			csXfs := XLSXCellStyleXfs{}
			csXfs.addXf(XLSXXf{XfId: &namedStyleId,
				ApplyBorder: true,
				ApplyFont:   false,
			})
			styles.CellStyleXfs = &csXfs
			cellStyleId := 0
			styles.CellXfs.addXf(
				XLSXXf{
					XfId:        &cellStyleId,
					ApplyBorder: false,
					ApplyFont:   true,
				})
			s0 := styles.getStyle(0)
			c.Assert(s0.NamedStyleIndex, qt.Equals, &cellStyleId)
			c.Assert(s0.ApplyBorder, qt.Equals, true)
			c.Assert(s0.ApplyFont, qt.Equals, true)
		})

	})

	c.Run("PopulateStyleFromXf", func(c *qt.C) {
		c.Run("ApplyBorder", func(c *qt.C) {
			styles := newXLSXStyleSheet(nil)
			style := &Style{}
			xf := XLSXXf{
				ApplyBorder: true,
			}
			styles.populateStyleFromXf(style, xf)
			c.Assert(style.ApplyBorder, qt.Equals, true)

			xf = XLSXXf{
				ApplyBorder: false,
			}
			styles.populateStyleFromXf(style, xf)
			c.Assert(style.ApplyBorder, qt.Equals, false)
		})

		c.Run("ApplyFill", func(c *qt.C) {
			styles := newXLSXStyleSheet(nil)
			style := &Style{}
			xf := XLSXXf{
				ApplyFill: true,
			}
			styles.populateStyleFromXf(style, xf)
			c.Assert(style.ApplyFill, qt.Equals, true)

			xf = XLSXXf{
				ApplyFill: false,
			}
			styles.populateStyleFromXf(style, xf)
			c.Assert(style.ApplyFill, qt.Equals, false)
		})
		c.Run("ApplyFont", func(c *qt.C) {
			styles := newXLSXStyleSheet(nil)
			style := &Style{}
			xf := XLSXXf{
				ApplyFont: true,
			}
			styles.populateStyleFromXf(style, xf)
			c.Assert(style.ApplyFont, qt.Equals, true)

			xf = XLSXXf{
				ApplyFont: false,
			}
			styles.populateStyleFromXf(style, xf)
			c.Assert(style.ApplyFont, qt.Equals, false)
		})
		c.Run("ApplyAlignment", func(c *qt.C) {
			styles := newXLSXStyleSheet(nil)
			style := &Style{}
			xf := XLSXXf{
				ApplyAlignment: true,
			}
			styles.populateStyleFromXf(style, xf)
			c.Assert(style.ApplyAlignment, qt.Equals, true)

			xf = XLSXXf{
				ApplyAlignment: false,
			}
			styles.populateStyleFromXf(style, xf)
			c.Assert(style.ApplyAlignment, qt.Equals, false)
		})
		c.Run("Border", func(c *qt.C) {
			styles := newXLSXStyleSheet(nil)
			line := XLSXLine{Style: "fake", Color: XLSXColor{RGB: "00aaff"}}

			borders := XLSXBorders{}
			border := XLSXBorder{
				Left:   line,
				Right:  line,
				Top:    line,
				Bottom: line,
			}
			borders.addBorder(border)

			styles.Borders = borders
			style := &Style{}
			xf := XLSXXf{
				ApplyBorder: true,
				BorderId:    0,
			}
			styles.populateStyleFromXf(style, xf)

			c.Assert(style.Border.Left, qt.Equals, border.Left.Style)
			c.Assert(style.Border.LeftColor, qt.Equals, border.Left.Color.RGB)
			c.Assert(style.Border.Right, qt.Equals, border.Right.Style)
			c.Assert(style.Border.RightColor, qt.Equals, border.Right.Color.RGB)
			c.Assert(style.Border.Top, qt.Equals, border.Top.Style)
			c.Assert(style.Border.TopColor, qt.Equals, border.Top.Color.RGB)
			c.Assert(style.Border.Bottom, qt.Equals, border.Bottom.Style)
			c.Assert(style.Border.BottomColor, qt.Equals, border.Bottom.Color.RGB)

		})

		c.Run("Fill", func(c *qt.C) {
			styles := newXLSXStyleSheet(nil)

			fills := XLSXFills{}
			pattern := XLSXPatternFill{
				PatternType: "fake",
				FgColor:     XLSXColor{RGB: "00aaff"},
				BgColor:     XLSXColor{RGB: "ffaa00"},
			}
			fill := XLSXFill{
				PatternFill: pattern,
			}
			fills.addFill(fill)

			styles.Fills = fills
			style := &Style{}
			xf := XLSXXf{
				ApplyFill: true,
				FillId:    0,
			}
			styles.populateStyleFromXf(style, xf)
			c.Assert(style.Fill.PatternType, qt.Equals, pattern.PatternType)
			c.Assert(style.Fill.FgColor, qt.Equals, styles.argbValue(pattern.FgColor))
			c.Assert(style.Fill.BgColor, qt.Equals, styles.argbValue(pattern.BgColor))

		})
		c.Run("Font", func(c *qt.C) {
			styles := newXLSXStyleSheet(nil)

			fonts := XLSXFonts{}

			sz := 10.0
			szVal := strconv.FormatFloat(sz, 'f', -1, 64)
			name := 0
			nameVal := strconv.Itoa(name)
			family := 2
			familyVal := strconv.Itoa(family)
			charset := 10
			charsetVal := strconv.Itoa(charset)

			font := XLSXFont{
				Sz:      XLSXVal{szVal},
				Name:    XLSXVal{nameVal},
				Family:  XLSXVal{familyVal},
				Charset: XLSXVal{charsetVal},
				Color:   XLSXColor{RGB: "00aaff"},
				B:       &XLSXVal{"1"},
				I:       &XLSXVal{"1"},
				U:       &XLSXVal{"1"},
				Strike:  &XLSXVal{"1"},
			}

			fonts.addFont(font)

			styles.Fonts = fonts
			style := &Style{}
			xf := XLSXXf{
				ApplyFont: true,
				FontId:    0,
			}
			styles.populateStyleFromXf(style, xf)

			c.Assert(style.Font.Size, qt.Equals, sz)
			c.Assert(style.Font.Name, qt.Equals, nameVal)
			c.Assert(style.Font.Family, qt.Equals, family)
			c.Assert(style.Font.Charset, qt.Equals, charset)
			c.Assert(style.Font.Color, qt.Equals, font.Color.RGB)
			c.Assert(style.Font.Bold, qt.Equals, true)
			c.Assert(style.Font.Italic, qt.Equals, true)
			c.Assert(style.Font.Underline, qt.Equals, true)
			c.Assert(style.Font.Strike, qt.Equals, true)
		})

		c.Run("Alignment", func(c *qt.C) {
			styles := newXLSXStyleSheet(nil)
			style := &Style{}

			alignment := XLSXAlignment{
				Horizontal:   "left",
				Indent:       10,
				ShrinkToFit:  true,
				TextRotation: 80,
				Vertical:     "top",
				WrapText:     true,
			}
			xf := XLSXXf{
				ApplyAlignment: true,
				Alignment:      alignment,
			}
			styles.populateStyleFromXf(style, xf)
			c.Assert(style.Alignment.Horizontal, qt.Equals, alignment.Horizontal)
			c.Assert(style.Alignment.Indent, qt.Equals, alignment.Indent)
			c.Assert(style.Alignment.ShrinkToFit, qt.Equals, alignment.ShrinkToFit)
			c.Assert(style.Alignment.TextRotation, qt.Equals, alignment.TextRotation)
			c.Assert(style.Alignment.Vertical, qt.Equals, alignment.Vertical)
			c.Assert(style.Alignment.WrapText, qt.Equals, alignment.WrapText)
		})

	})
}
