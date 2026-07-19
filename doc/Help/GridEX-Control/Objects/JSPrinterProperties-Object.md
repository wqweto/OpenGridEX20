# JSPrinterProperties Object

## JSPrinterProperties

**![](../../images/JSPrinterProperties.jpg)  
 PrinterProperties** object enables you to specify the printer and print job properties to be used when a **GridEX** control is going to be printed.

### Syntax

 *gridex*.**PrinterProperties**  
 The *gridex* placeholder represents an object expression that evaluates to a **GridEX** control.

### Remarks

 With a **JSPrinterProperties** object, you can modify document attributes such as page layout, margins, page headers and page footers.

### Data Type

 **JSPrinterProperties**

**See Also:** [PrinterProperties Property](../Properties.md#printerproperties-property-gridex-control)

## BottomMargin Property (JSPrinterProperties Object)

Returns or sets the height, in twips, of the bottom margin.

### Syntax

 *object*.**BottomMargin** [ = *number* ]  
 The **BottomMargin** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *number* | A Long expression specifying the height, in twips, of the bottom margin. |

### Data Type

 Long

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [FooterDistance Property](#footerdistance-property-jsprinterproperties-object), [HeaderDistance Property](#headerdistance-property-jsprinterproperties-object), [LeftMargin Property](#leftmargin-property-jsprinterproperties-object), [RightMargin Property](#rightmargin-property-jsprinterproperties-object), [TopMargin Property](#topmargin-property-jsprinterproperties-object)  
**Example:** [PrinterPropertiesMargins Example](../../Examples.md#printerpropertiesmargins-example)

## CardColumnsPerPage Property (JSPrinterProperties Object)

Returns or sets the number of card columns to be printed in a page.

### Syntax

 *object*.**CardColumnsPerPage** [ = *number* ]  
 The **CardColumnsPerPage** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *number* | A Long expression specifying the number of card columns to be printed in a page. |

### Remarks

 **CardColumnsPerPage** must be greather than 0 and less than or equal to 6.  
 The number of column cards printed is independent of the number of column cards in a **GridEX** control. Therefore, **CardWidth** property does not apply to the card width of a printed card.

### Data Type

 Long

**Applies To:** [GridEX Control](../../GridEX-Control.md#gridex-control)

## ClientHeight Property (JSPrinterProperties Object)

Returns the height, in twips, of the printable area in a page.

### Syntax

 *object*.**ClientHeight** [ = *number* ]  
 The **ClientHeight** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *number* | A Long expression specifying the height, in twips, of the printable area in a page. |

### Data Type

 Long

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [ClientWidth Property](#clientwidth-property-jsprinterproperties-object), [PaperHeight Property](#paperheight-property-jsprinterproperties-object), [PaperWidth Property](#paperwidth-property-jsprinterproperties-object)

## ClientWidth Property (JSPrinterProperties Object)

Returns the width, in twips, of the printable area in a page.

### Syntax

 *object*.**ClientWidth** [ = *number* ]  
 The **ClientWidth** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *number* | A Long expression specifying the width, in twips, of the printable area in a page. |

### Data Type

 Long

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [PaperHeight Property](#paperheight-property-jsprinterproperties-object), [PaperWidth Property](#paperwidth-property-jsprinterproperties-object), [ClientHeight Property](#clientheight-property-jsprinterproperties-object)

## Collate Property (JSPrinterProperties Object)

Returns or sets a value that determines whether collation should be used when printing multiple copies.

### Syntax

 *object*.**Collate** [ = *value* ]  
 The **Collate** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether collation should be used when printing multiple copies, as described in settings. |

### Settings

 The settings for *value* are:

| Setting | Description |
| --- | --- |
| **True** | Collate when printing multiple copies. |
| **False** | Do not collate when printing multiple copies. |

### Data Type

 Boolean

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [Copies Property](#copies-property-jsprinterproperties-object)

## ColorMode Property (JSPrinterProperties Object)

Returns or sets a value that determines whether a document should be printed in color or monochrome.

### Syntax

 *object*.**ColorMode** [ = *value* ]  
 The **ColorMode** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines whether a document should be printed in color, as described in settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
| jgexPPCMMonochrome | 1 | Document is printed in monochrome. |
| jgexPPCMColor | 2 | Document is printed in color. |

### Remarks

 This property has no effect with monochrome printers.

### Data Type

 **jgexPPColorModeConstants**

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [TransparentBackground Property](#transparentbackground-property-jsprinterproperties-object), [TranslateColors Property](#translatecolors-property-jsprinterproperties-object), [PrintQuality Property](#printquality-property-jsprinterproperties-object)

## Copies Property (JSPrinterProperties Object)

Returns or sets a value that determines the number of copies to be printed.

### Syntax

 *object*.**Copies** [ = *number* ]  
 The **Copies** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *number* | A Long expression specifying the number of copies to be printed. |

### Data Type

 Long

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [Collate Property](#collate-property-jsprinterproperties-object)

## DeviceName Property (JSPrinterProperties Object)

Returns the name of the device a driver supports.

### Syntax

 *object*.**DeviceName**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 **DeviceName** property is a read-only property.  
 To change the **DeviceName** property use the **ChangePrinter** method.

### Data Type

 String

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [DriverName Property](#drivername-property-jsprinterproperties-object), [ChangePrinter Method](#changeprinter-method-jsprinterproperties-object)  
**Example:** [PrintingProgress Example](../../Examples.md#printingprogress-example)

## DocumentName Property (JSPrinterProperties Object)

Returns or sets a string that specifies the name of the document to be printed.

### Syntax

 *object*.**DocumentName** [ = *value*]  
 The **DocumentName** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A string expression that specifies the name of the document to be printed. |

### Remarks

 Use this property to specify the name of the document that is shown in the print queue window when the **GridEX** control prints its contents.

### Data Type

 String

**Applies To:** [JSPrinterProperties](#jsprinterproperties)

## DriverName Property (JSPrinterProperties Object)

Returns the name of the printer's driver.

### Syntax

 *object*.**DriverName**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 **DriverName** property is a read-only property.  
 To change the **DriverName** property use the **ChangePrinter** method.

### Data Type

 String

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [DeviceName Property](#devicename-property-jsprinterproperties-object), [ChangePrinter Method](#changeprinter-method-jsprinterproperties-object)

## FitColumns Property (JSPrinterProperties Object)

Returns or sets a value indicating whether the **GridEX** control should scale printed output to fit all visible columns in a page.

### Syntax

 *object*.**FitColumns** [ = *value*]  
 The **FitColumns** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether the **GridEX** control should scale printed output to fit all visible columns in a page, as described in Settings. |

### Settings

 The settings for *value* are:

| Setting | Description |
| --- | --- |
| **True** | Printed output should be scaled so that all visible columns fit in one page. |
| **False** | Default. Printed output is not scaled. |

### Remarks

 This property is only used when the **GridEX** control is in **Table** view.

### Data Type

 Boolean

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [Orientation Property](#orientation-property-jsprinterproperties-object), [View Property](../Properties.md#view-property-gridex-control)  
**Example:** [PrinterProperties Example](../../Examples.md#printerproperties-example)

## FooterDistance Property (JSPrinterProperties Object)

Returns or sets the distance, in twips, from the bottom edge of the paper to the top edge of the footer.

### Syntax

 *object*.**FooterDistance** [ = *number* ]  
 The **FooterDistance** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *number* | A Long expression specifying the distance, in twips, from the bottom edge of the paper to the top edge of the footer. |

### Remarks

 **FooterDistance** should be less than **BottomMargin** property.  
 If the **FooterDistance** lies in the non-printable area of the page no footers will be printed.

### Data Type

 Long

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [HeaderDistance Property](#headerdistance-property-jsprinterproperties-object), [LeftMargin Property](#leftmargin-property-jsprinterproperties-object), [RightMargin Property](#rightmargin-property-jsprinterproperties-object), [TopMargin Property](#topmargin-property-jsprinterproperties-object), [BottomMargin Property](#bottommargin-property-jsprinterproperties-object)  
**Example:** [PrinterPropertiesMargins Example](../../Examples.md#printerpropertiesmargins-example)

## FooterString Property (JSPrinterProperties Object)

Returns or sets the string displayed in the page footer.

### Syntax

 *object*.**FooterString** (*position*) [ = *value* ]  
 The **FooterString** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *position* | A value or constant that specifies the position of the footer string, as described in settings. |
| *value* | A String expression specifying the footer string. |

### Settings

 The settings for *position* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexHFCenter** | 1 | Index of the footer string to be printed horizontally centered. |
|  **jgexHFLeft** | 2 | Index of the footer string to be printed at the left. |
|  **jgexHFRight** | 3 | Index of the footer string to be printed at the right. |

### Remarks

 Use this property to specify the text to be printed in the page footer.  
 You can specify the three footer strings in the same document.

### Data Type

 String

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [HeaderString Property](#headerstring-property-jsprinterproperties-object), [PageFooterFont Property](#pagefooterfont-property-jsprinterproperties-object), [PageHeaderFont Property](#pageheaderfont-property-jsprinterproperties-object)  
**Example:** [PrinterPropertiesMargins Example](../../Examples.md#printerpropertiesmargins-example)

## HeaderDistance Property (JSPrinterProperties Object)

Returns or sets the distance, in twips, from the top edge of the paper to the top edge of the header.

### Syntax

 *object*.**HeaderDistance** [ = *number* ]  
 The **HeaderDistance** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *number* | A Long expression specifying the distance, in twips, from the top edge of the paper to the top edge of the header. |

### Remarks

 **HeaderDistance** should be less than **TopMargin** property.  
 If the **HeaderDistance** lies in the non-printable area of the page no headers will be printed.

### Data Type

 Long

**Applies To:** [GridEX Control](../../GridEX-Control.md#gridex-control)  
**See Also:** [FooterDistance Property](#footerdistance-property-jsprinterproperties-object), [LeftMargin Property](#leftmargin-property-jsprinterproperties-object), [RightMargin Property](#rightmargin-property-jsprinterproperties-object), [TopMargin Property](#topmargin-property-jsprinterproperties-object), [BottomMargin Property](#bottommargin-property-jsprinterproperties-object)  
**Example:** [PrinterPropertiesMargins Example](../../Examples.md#printerpropertiesmargins-example)

## HeaderString Property (JSPrinterProperties Object)

Returns or sets the string displayed in the page header.

### Syntax

 *object*.**HeaderString** (*position*) [ = *value* ]  
 The **HeaderString** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *position* | A value or constant that specifies the position of the header string, as described in settings. |
| *value* | A String expression specifying the header string. |

### Settings

 The settings for *position* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexHFCenter** | 1 | Index of the header string to be printed horizontally centered. |
|  **jgexHFLeft** | 2 | Index of the header string to be printed at the left. |
|  **jgexHFRight** | 3 | Index of the header string to be printed at the right. |

### Remarks

 Use this property to specify the text to be printed in the page header.  
 You can specify the three header strings in the same document.

### Data Type

 String

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [FooterString Property](#footerstring-property-jsprinterproperties-object), [PageFooterFont Property](#pagefooterfont-property-jsprinterproperties-object), [PageHeaderFont Property](#pageheaderfont-property-jsprinterproperties-object)  
**Example:** [PrinterPropertiesMargins Example](../../Examples.md#printerpropertiesmargins-example)

## LeftMargin Property (JSPrinterProperties Object)

Returns or sets the width, in twips, of the left margin.

### Syntax

 *object*.**LeftMargin** [ = *number* ]  
 The **LeftMargin** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *number* | A Long expression specifying the width, in twips, of the left margin. |

### Data Type

 Long

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [FooterDistance Property](#footerdistance-property-jsprinterproperties-object), [HeaderDistance Property](#headerdistance-property-jsprinterproperties-object), [RightMargin Property](#rightmargin-property-jsprinterproperties-object), [TopMargin Property](#topmargin-property-jsprinterproperties-object), [BottomMargin Property](#bottommargin-property-jsprinterproperties-object)  
**Example:** [PrinterPropertiesMargins Example](../../Examples.md#printerpropertiesmargins-example)

## MeasurementUnits Property (JSPrinterProperties Object)

Returns or sets a value that determines whether the Page Setup dialog should use inches or millimeters as unit of measurement for margins.

### Syntax

 *object*.**MeasurementUnits**= [*value*]  
 The **MeasurementUnits** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines whether the Page Setup dialog should use inches or millimeters, as described in settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexMUInches** | 0 | Default. Inches are the unit of measurement for margins in the Page Setup dialog. |
|  **jgexMUMilimeters** | 1 | Millimeters are the unit of measurement for margins in the Page Setup dialog. |

### Data Type

 **jgexMeasurementUnitsConstants**

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [PageSetup Method](#pagesetup-method-jsprinterproperties-object)

## Orientation Property (JSPrinterProperties Object)

Returns or sets a value indicating whether documents are printed in portrait or landscape mode.

### Syntax

 *object*.**Orientation** [ = *value* ]  
 The **Orientation** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines whether documents are printed in portrait or landscape mode, as described in Settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexPPPortrait** | 1 | Documents are printed with the top at the narrow side of the paper. |
|  **jgexPPLandscape** | 2 | Documents are printed with the top at the wide side of the paper. |

### Remarks

 This property has no effect with some printer drivers that don't support landscape printing.

### Data Type

 **jgexPPOrientationConstants**

**Applies To:** [JSPrinterProperties](#jsprinterproperties)

## PageFooterFont Property (JSPrinterProperties Object)

Returns or sets a Font object used to draw page footers in a document.

### Syntax

 *object*.**PageFooterFont** [ = *value* ]  
 The **PageFooterFont** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | object expression that evaluates to an StdFont object. |

### Data Type

 **StdFont**

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [FooterString Property](#footerstring-property-jsprinterproperties-object), [HeaderString Property](#headerstring-property-jsprinterproperties-object), [PageHeaderFont Property](#pageheaderfont-property-jsprinterproperties-object)

## PageHeaderFont Property (JSPrinterProperties Object)

Returns or sets a **Font** object used to draw page headers in a document.

### Syntax

 *object*.**PageHeaderFont** [ = *value* ]  
 The **PageHeaderFont** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | object expression that evaluates to an StdFont object. |

### Data Type

 **StdFont**

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [FooterString Property](#footerstring-property-jsprinterproperties-object), [HeaderString Property](#headerstring-property-jsprinterproperties-object), [PageFooterFont Property](#pagefooterfont-property-jsprinterproperties-object)

## PaperBin Property (JSPrinterProperties Object)

Returns or sets a value indicating the default paper bin on the printer from which paper is fed when printing.

### Syntax

 *object*.**PaperBin** [ = *value* ]  
 The **PaperBin** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant indicating the default paper bin on the printer from which paper is fed when printing, as described in settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexPPUpper** | 1 | Use paper from the upper bin. |
|  **jgexPPLower** | 2 | Use paper from the lower bin. |
|  **jgexPPMiddle** | 3 | Use paper from the middle bin. |
|  **jgexPPManual** | 4 | Wait for manual insertion of each sheet of paper. |
|  **jgexPPEnvelope** | 5 | Use envelopes from the envelope feeder. |
|  **jgexPPEnvManual** | 6 | Use envelopes from the envelope feeder, but wait for manual insertion. |
|  **jgexPPAuto** | 7 | (Default) Use paper from the current default bin. |
|  **jgexPPTractor** | 8 | Use paper fed from the tractor feeder. |
|  **jgexPPSmallFmt** | 9 | Use paper from the small paper feeder. |
|  **jgexPPLargeFmt** | 10 | Use paper from the large paper bin. |
|  **jgexPPLargeCapacity** | 11 | Use paper from the large capacity feeder. |
|  **jgexPPCassete** | 14 | Use paper from the attached cassette cartridge. |

### Remarks

 Availability of bin options varies depending on the printer. Check the printer documentation.

### Data Type

 **jgexPPPaperBinConstants**

**Applies To:** [JSPrinterProperties](#jsprinterproperties)

## PaperHeight Property (JSPrinterProperties Object)

Returns or sets the physical height, in twips, of the paper set up for the printing device.

### Syntax

 *object*.**PaperHeight** [ = *number* ]  
 The **PaperHeight** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *number* | A Long expression specifying the height, in twips, of the paper set up for the printing device. |

### Remarks

 Use this property to use a custom paper size for printing.  
 If a predefined paper size is to be used, use **PaperSize** property instead.

### Data Type

 Long

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [ClientWidth Property](#clientwidth-property-jsprinterproperties-object), [PaperWidth Property](#paperwidth-property-jsprinterproperties-object), [ClientHeight Property](#clientheight-property-jsprinterproperties-object)

## PageSetupCanceled Property (JSPrinterProperties Object)

Returns whether the user has canceled the Page Setup dialog. Read only.

### Syntax

 *object*.**PageSetupCanceled**= [*value*]  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Settings

 The settings for *value* are:

| Setting | Description |
| --- | --- |
| **True** | Page Setup dialog was accepted. |
| **False** | Page Setup dialog was canceled. |

### Remarks

 This property is used only after the **PageSetup** method is called. By examining this property you can know whether to repaginate a **GEXPreview** control or not.  
 An example of this is as follows:

```vb
frmMain.GridEX1.PrinterProperties.PageSetup
If Not GridEX1.PrinterProperties.PageSetupCanceled Then
frmPrintPreview.GEXPreview1.Repaginate
End If
```

### Data Type

 Boolean

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [PageSetup Method](#pagesetup-method-jsprinterproperties-object)

## PaperSize Property (JSPrinterProperties Object)

Returns or sets a value indicating the paper size for the current printer.

### Syntax

 *object*.**PaperSize** [ = *value* ]  
 The **PaperSize** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| value | A value or constant indicating the paper size for the current printer, as described in settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexPSLetter** | 1 | Letter, 8 1/2 x 11 in. |
|  **jgexPSLetterSmall** | 2 | Letter Small, 8 1/2 x 11 in. |
|  **jgexPSTabloid** | 3 | Tabloid, 11 x 17 in. |
|  **jgexPSLedger** | 4 | Ledger, 17 x 11 in. |
|  **jgexPSLegal** | 5 | Legal, 8 1/2 x 14 in. |
|  **jgexPSStatement** | 6 | Statement, 5 1/2 x 8 1/2 in. |
|  **jgexPSExecutive** | 7 | Executive, 7 1/2 x 10 1/2 in |
|  **jgexPSA3** | 8 | A3, 297 x 420 mm |
|  **jgexPSA4** | 9 | A4, 210 x 297 mm |
|  **jgexPSA4Small** | 10 | A4 Small, 210 x 297 mm |
|  **jgexPSA5** | 11 | A5, 148 x 210 mm |
|  **jgexPSB4** | 12 | B4, 250 x 354 mm |
|  **jgexPSB5** | 13 | B5, 182 x 257 mm |
|  **jgexPSFolio** | 14 | Folio, 8 1/2 x 13 in. |
|  **jgexPSQuarto** | 15 | Quarto, 215 x 275 mm |
|  **jgexPS10x14** | 16 | 10 x 14 in |
|  **jgexPS11x17** | 17 | 11 x 17 in. |
|  **jgexPSNote** | 18 | Note, 8 1/2 x 11 in. |
|  **jgexPSEnv9** | 19 | Envelope #9, 3 7/8 x 8 7/8 in. |
|  **jgexPSEnv10** | 20 | Envelope #10, 4 1/8 x 9 1/2 in. |
|  **jgexPSEnv11** | 21 | Envelope #11, 4 1/2 x 10 3/8 in. |
|  **jgexPSEnv12** | 22 | Envelope #12, 4 1/2 x 11 in. |
|  **jgexPSEnv14** | 23 | Envelope #14, 5 x 11 1/2 in |
|  **jgexPSCSheet** | 24 | C size sheet |
|  **jgexPSDSheet** | 25 | D size sheet |
|  **jgexPSESheet** | 26 | E size sheet |
|  **jgexPSEnvDL** | 27 | Envelope DL, 110 x 220 mm |
|  **jgexPSEnvC3** | 29 | Envelope C3, 324 x 458 mm |
|  **jgexPSEnvC4** | 30 | Envelope C4, 229 x 324 mm |
|  **jgexPSEnvC5** | 28 | Envelope C5, 162 x 229 mm |
|  **jgexPSEnvC6** | 31 | Envelope C6, 114 x 162 mm |
|  **jgexPSEnvC65** | 32 | Envelope C65, 114 x 229 mm |
|  **jgexPSEnvB4** | 33 | Envelope B4, 250 x 353 mm |
|  **jgexPSEnvB5** | 34 | Envelope B5, 176 x 250 mm |
|  **jgexPSEnvB6** | 35 | Envelope B6, 176 x 125 mm |
|  **jgexPSEnvItaly** | 36 | Envelope, 110 x 230 mm |
|  **jgexPSEnvMonarch** | 37 | Envelope Monarch, 3 7/8 x 7 1/2 in. |
|  **jgexPSEnvPersonal** | 38 | Envelope, 3 5/8 x 6 1/2 in. |
|  **jgexPSFanfoldUS** | 39 | U.S. Standard Fanfold, 14 7/8 x 11 in. |
|  **jgexPSFanfoldStdGerman** | 40 | German Standard Fanfold, 8 1/2 x 12 in. |
|  **jgexPSFanfoldLglGerman** | 41 | German Legal Fanfold, 8 1/2 x 13 in. |
|  **jgexPSUser** | 256 | User-defined |

### Remarks

 Setting **PaperHeight** or **PaperWidth**, changes **PaperSizeProperty** to **jgexPSUser**

### Data Type

 **jgexPPPaperBinConstants**

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [PaperWidth Property](#paperwidth-property-jsprinterproperties-object), [PaperHeight Property](#paperheight-property-jsprinterproperties-object), [PageSetup Method](#pagesetup-method-jsprinterproperties-object)

## PaperWidth Property (JSPrinterProperties Object)

Returns or sets the physical width, in twips, of the paper set up for the printing device.

### Syntax

 *object*.**PaperWidth** [ = *number* ]  
 The **PaperWidth** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| number | A Long expression specifying the width, in twips, of the paper set up for the printing device. |

### Remarks

 Use this property to use a custom paper size for printing.  
 If a predefined paper size is to be used, use **PaperSize** property instead.

### Data Type

 Long

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [ClientWidth Property](#clientwidth-property-jsprinterproperties-object), [PaperHeight Property](#paperheight-property-jsprinterproperties-object), [ClientHeight Property](#clientheight-property-jsprinterproperties-object)

## PrintPreviewRows Property (JSPrinterProperties Object)

Returns or sets a value indicating whether the **GridEX** control prints preview rows or not.

### Syntax

 *object*.**PrintPreviewRows** = [*value*]  
 The **PrintPreviewRows** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether whether the **GridEX** control will print preview rows, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | The **GridEX** control prints preview rows. |
| **False** | The **GridEX** control doesn't print preview rows. |

### Remarks

 This property has no effect if the **GridEX** control is in **CardView** or if the **GirdEX** control doesn't show preview rows.

### Data Type

 Boolean

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [PreviewColumn Property](../Properties.md#previewcolumn-property-gridex-control), [PreviewRowIndent Property](../Properties.md#previewrowindent-property-gridex-control), [PreviewRowLines Property](../Properties.md#previewrowlines-property-gridex-control), [PreviewRowVisible Property](JSRowData-Object.md#previewrowvisible-property-jsrowdata-object)  
**Example:** [PrinterProperties Example](../../Examples.md#printerproperties-example)

## PrintProgressDialog Property (JSPrinterProperties Object)

Returns or sets a value indicating whether print progress dialog should be displayed when **GridEX** control is printing a document.

### Syntax

 *object*.**PrintProgressDialog** = [*value*]  
 The **PrintProgressDialog** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether print progress dialog should be displayed when **GridEX** control is printing a document, as described in Settings. |

### Settings

 The settings for *value* are:

| Setting | Description |
| --- | --- |
| **True** | Default. **GridEX** control should display the Print Progress dialog when printing. |
| **False** | **GridEX** control doesn't display the Print Progress dialog when printing. |

### Remarks

 The Print Progress dialog gives users information about the print operation progress and let them cancel the print operation at any time.

### Data Type

 Boolean

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [PrintingProgress Event](../Events.md#printingprogress-event-gridex-control)  
**Example:** [PrintingProgress Example](../../Examples.md#printingprogress-example)

## PrintQuality Property (JSPrinterProperties Object)

Returns or sets a value indicating the printer resolution.

### Syntax

 *object*.**PrintQuality** [ = *value* ]  
 The **PrintQuality** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant indicating the printer resolution as described in Settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexPPDraft** | -1 | Draft resolution. |
|  **jgexPPLow** | -2 | Low resolution. |
|  **jgexPPMedium** | -3 | Medium resolution. |
|  **jgexPPHigh** | -4 | High resolution. |

In addition to the predefined print quality values, you can also set value to a positive value representing the dots per inch (dpi) resolution.

### Remarks

 The effect of this property depends on the printer driver.

### Data Type

 Integer

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [TransparentBackground Property](#transparentbackground-property-jsprinterproperties-object), [TranslateColors Property](#translatecolors-property-jsprinterproperties-object), [ColorMode Property](#colormode-property-jsprinterproperties-object)

## RepeatFrozenCols Property (JSPrinterProperties Object)

Returns or sets a value indicating whether frozen columns should appear on each page.

### Syntax

 *object*.**RepeatFrozenCols** = [*value*]  
 The **RepeatFrozenCols** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether frozen columns should appear on each page, as described in Settings. |

### Settings

 The settings for *value* are:

| Setting | Description |
| --- | --- |
| **True** | Frozen columns should appear on each page. |
| **False** | Frozen columns should appear only on the first page. |

### Remarks

 This property has no effect if the **GridEX** control is in **CardView** or if the control's **FrozenColumns** property is equal to 0.

### Data Type

 Boolean

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [RepeatHeaders Property](#repeatheaders-property-jsprinterproperties-object)  
**Example:** [PrinterProperties Example](../../Examples.md#printerproperties-example)

## RepeatHeaders Property (JSPrinterProperties Object)

Returns or sets a value indicating whether column headers should appear on each page.

### Syntax

 *object*.**RepeatHeaders** = [*value*]  
 The **RepeatHeaders** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether column headers should appear on each page, as described in Settings. |

### Settings

 The settings for *value* are:

| Setting | Description |
| --- | --- |
| **True** | Column headers should appear on each page. |
| **False** | Column headers should appear only on the first page. |

### Remarks

 This property has no effect if the **GridEX** control is in CardView.

### Data Type

 Boolean

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [RepeatFrozenCols Property](#repeatfrozencols-property-jsprinterproperties-object)  
**Example:** [PrinterProperties Example](../../Examples.md#printerproperties-example)

## RightMargin Property (JSPrinterProperties Object)

Returns or sets the width, in twips, of the right margin.

### Syntax

 *object*.**RightMargin** [ = *number* ]  
 The **RightMargin** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *number* | A Long expression specifying the width, in twips, of the right margin. |

### Data Type

 Long

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [FooterDistance Property](#footerdistance-property-jsprinterproperties-object), [HeaderDistance Property](#headerdistance-property-jsprinterproperties-object), [LeftMargin Property](#leftmargin-property-jsprinterproperties-object), [TopMargin Property](#topmargin-property-jsprinterproperties-object), [BottomMargin Property](#bottommargin-property-jsprinterproperties-object)  
**Example:** [PrinterPropertiesMargins Example](../../Examples.md#printerpropertiesmargins-example)

## TopMargin Property (JSPrinterProperties Object)

Returns or sets the height, in twips, of the top margin.

### Syntax

 *object*.**TopMargin** [ = *number* ]  
 The **TopMargin** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| number | A Long expression specifying the height, in twips, of the top margin. |

### Data Type

 Long

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [FooterDistance Property](#footerdistance-property-jsprinterproperties-object), [HeaderDistance Property](#headerdistance-property-jsprinterproperties-object), [LeftMargin Property](#leftmargin-property-jsprinterproperties-object), [RightMargin Property](#rightmargin-property-jsprinterproperties-object), [BottomMargin Property](#bottommargin-property-jsprinterproperties-object)  
**Example:** [PrinterPropertiesMargins Example](../../Examples.md#printerpropertiesmargins-example)

## TranslateColors Property (JSPrinterProperties Object)

Returns or sets a value indicating whether system colors in a **GridEX** control should be translated to black and white colors.

### Syntax

 *object*.**TranslateColors** = [*value*]  
 The **TranslateColors** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether system colors in a **GridEX** control should be translated to black and white colors, as described in Settings. |

### Settings

 The settings for *value* are:

| Setting | Description |
| --- | --- |
| **True** | Systems colors are translated to Black and White equivalents. |
| **False** | Default. Systems colors are printed as they are displayed in the screen. |

### Remarks

 Setting this property to **True**, the **GridEX** control will use a black and white equivalent instead of the actual system color in the user's machine.  
 System colors are translated according to the following table:

| System Color | Description | Equivalent color used |
| --- | --- | --- |
| &H80000000 | Scroll bar color. | Light gray. |
| &H80000001 | Desktop color. | Dark gray. |
| &H80000002 | Color of the title bar for the active window. | Black. |
| &H80000003 | Color of the title bar for the inactive window. | Dark gray. |
| &H80000004 | Menu background color. | Light gray. |
| &H80000005 | Window background color. | White. |
| &H80000006 | Window frame color. | Black. |
| &H80000007 | Color of text on menus. | Black. |
| &H80000008 | Color of text in windows. | Black. |
| &H80000009 | Color of text in caption, size box, and scroll arrow. | White. |
| &H8000000A | Border color of active window. | Light gray. |
| &H8000000B | Border color of inactive window. | Light gray. |
| &H8000000C | Background color of multiple-document interface (MDI) applications. | Dark gray. |
| &H8000000D | Background color of items selected in a control. | Black. |
| &H8000000E | Text color of items selected in a control. | White. |
| &H8000000F | Color of shading on the face of command buttons. | Light gray. |
| &H80000010 | Color of shading on the edge of command buttons. | Dark gray. |
| &H80000011 | Grayed (disabled) text. | Light gray. |
| &H80000012 | Text color on push buttons. | Black. |
| &H80000013 | Color of text in an inactive caption. | Light gray. |
| &H80000014 | Highlight color for 3D display elements. | White. |
| &H80000015 | Darkest shadow color for 3D display elements. | Black. |
| &H80000016 | Second lightest of the 3D colors after vb3Dhighlight. | Light gray. |
| &H80000017 | Color of text in ToolTips. | Black. |
| &H80000018 | Background color of ToolTips. | White. |

### Data Type

 Boolean

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [TransparentBackground Property](#transparentbackground-property-jsprinterproperties-object), [PrintQuality Property](#printquality-property-jsprinterproperties-object), [ColorMode Property](#colormode-property-jsprinterproperties-object)  
**Example:** [PrinterProperties Example](../../Examples.md#printerproperties-example)

## TransparentBackground Property (JSPrinterProperties Object)

Returns or sets a value indicating whether background colors should be printed.

### Syntax

 *object*.**TransparentBackground** = [*value*]  
 The **TransparentBackground** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether background colors should be printed. |

### Settings

 The settings for *value* are:

| Setting | Description |
| --- | --- |
| **True** | Headers and cells are printed without using their background color. |
| **False** | (Default). Headers and cells are printed using their background color. |

### Data Type

 Boolean

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [TranslateColors Property](#translatecolors-property-jsprinterproperties-object), [PrintQuality Property](#printquality-property-jsprinterproperties-object), [ColorMode Property](#colormode-property-jsprinterproperties-object)

## ChangePrinter Method (JSPrinterProperties Object)

Changes printer's device name and driver.

### Syntax

 *object*.**ChangePrinter** *devicename, drivername*  
 The **ChangePrinter** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *devicename* | A string that represents the name of a printer that could be set to an empty string to change to the default printer. |
| *drivername* | A string that represents the name of a printer driver that could be set to an empty string to change to the default printer driver. |

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [DriverName Property](#drivername-property-jsprinterproperties-object), [DeviceName Property](#devicename-property-jsprinterproperties-object)

## PageSetup Method (JSPrinterProperties Object)

Displays the page setup dialog.

### Syntax

 *object*.**PageSetup** *hwndowner*  
 The **PageSetup** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *hwndowner* | A Long expression that represents the window handle of the parent window for the dialog. |

**Applies To:** [JSPrinterProperties](#jsprinterproperties)  
**See Also:** [PageSetup Method](#pagesetup-method-jsprinterproperties-object), [PaperSize Property](#papersize-property-jsprinterproperties-object), [Orientation Property](#orientation-property-jsprinterproperties-object), [LeftMargin Property](#leftmargin-property-jsprinterproperties-object), [RightMargin Property](#rightmargin-property-jsprinterproperties-object), [TopMargin Property](#topmargin-property-jsprinterproperties-object), [BottomMargin Property](#bottommargin-property-jsprinterproperties-object)
