# GEXPreview Control — Properties

## BackColor Property (GEXPreview Control)

Returns or sets the background color of the **GEXPreview** control.

### Syntax

 *object*.**BackColor** [ =*color* ]  
 The **BackColor** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A numeric expression that specifies the color. |

### Data Type

 OLE_COLOR

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [BackColorBkg Property](../GridEX-Control/Properties.md#backcolorbkg-property-gridex-control), [BackColorGBBox Property](../GridEX-Control/Properties.md#backcolorgbbox-property-gridex-control), [BackColorHeader Property](../GridEX-Control/Properties.md#backcolorheader-property-gridex-control), [BackColorInfoText Property](../GridEX-Control/Properties.md#backcolorinfotext-property-gridex-control), [BackColor Property](../GridEX-Control/Properties.md#backcolor-property-gridex-control), [BackColor Property](../GridEX-Control/Objects/JSFormatStyle-Object.md#backcolor-property-jsformatstyle-object), [BackColorRowGroup Property](../GridEX-Control/Properties.md#backcolorrowgroup-property-gridex-control), [ForeColorHeader Property](../GridEX-Control/Properties.md#forecolorheader-property-gridex-control), [ForeColorInfoText Property](../GridEX-Control/Properties.md#forecolorinfotext-property-gridex-control), [ForeColor Property](../GridEX-Control/Objects/JSFormatStyle-Object.md#forecolor-property-jsformatstyle-object), [ForeColorRowGroup Property](../GridEX-Control/Properties.md#forecolorrowgroup-property-gridex-control), [ForeColor Property](../GridEX-Control/Properties.md#forecolor-property-gridex-control), [MaskColor Property](../GridEX-Control/Properties.md#maskcolor-property-gridex-control), [RowColorEven Property](../GridEX-Control/Properties.md#rowcoloreven-property-gridex-control), [RowColorOdd Property](../GridEX-Control/Properties.md#rowcolorodd-property-gridex-control), [GridLinesColor Property](../GridEX-Control/Properties.md#gridlinescolor-property-gridex-control)

## CloseButtonText Property (GEXPreview Control)

Returns or sets the text displayed in the "Close" button.

### Syntax

 *object*.**CloseButtonText** [ = *value* ]  
 The **CloseButtonText** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A string that represents the text displayed in the "Close" button. |

### Remarks

 You can change this property to any string for customization.

### Data Type

 String

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [PageSetupText Property](#pagesetuptext-property-gexpreview-control), [PrintText Property](#printtext-property-gexpreview-control), [ToolbarVisible Property](#toolbarvisible-property-gexpreview-control)

## CurrentPage Property (GEXPreview Control)

Returns or sets the page that is displayed in a **GEXPreview** control.

### Syntax

 *object*.**CurrentPage** [ = *value*]  
 The **CurrentPage** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Long expression that represents the page that is displayed in a **GEXPreview** control. |

### Data Type

 Long

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [NextPage Method](Methods.md#nextpage-method-gexpreview-control), [PreviousPage Method](Methods.md#previouspage-method-gexpreview-control), [Repaginate Method](Methods.md#repaginate-method-gexpreview-control), [TotalPages Property](#totalpages-property-gexpreview-control)

## hWnd Property (GEXPreview Control)

Returns the handle of a **GEXPreview** control.

### Syntax

 *object*.**hWnd**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The Microsoft Windows operating environment identifies each control in an application by assigning it a handle, or hWnd.  
 The **hWnd** property is used with Windows API calls. Many Windows operating environment functions require the hWnd of the active window as an argument.

### Data Type

 Long

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [hWnd Property](../GridEX-Control/Properties.md#hwnd-property-gridex-control)

## PageSetupText Property (GEXPreview Control)

Returns or sets the text displayed in the "Page Setup" button.

### Syntax

 *object*.**PageSetupText** [ = *value* ]  
 The **PageSetupText** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A string that represents the text displayed in the "Page Setup" button. |

### Remarks

 You can change this property to any string for customization.

### Data Type

 String

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [PrintText Property](#printtext-property-gexpreview-control), [ToolbarVisible Property](#toolbarvisible-property-gexpreview-control), [CloseButtonText Property](#closebuttontext-property-gexpreview-control)

## PrintText Property (GEXPreview Control)

Returns or sets the text displayed in the "Print" button.  
 Syntax  
 *object*.**PrintText** [ = *value* ]  
 The **PrintText** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A string that represents the text displayed in the "Print" button. |

### Remarks

 You can change this property to any string for customization.

### Data Type

 String

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [PageSetupText Property](#pagesetuptext-property-gexpreview-control), [ToolbarVisible Property](#toolbarvisible-property-gexpreview-control), [CloseButtonText Property](#closebuttontext-property-gexpreview-control)

## ToolbarFont Property (GEXPreview Control)

Returns or sets a Font object used in **GEXPreview** control's toolbar.

### Syntax

 *object*.**ToolbarFont**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Data Type

 **StdFont**

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [ToolbarVisible Property](#toolbarvisible-property-gexpreview-control)

## ToolbarVisible Property (GEXPreview Control)

Returns or sets a value indicating whether the tool bar is displayed in a **GEXPreview** control.

### Syntax

 *object*.**ToolbarVisible** [ = *value* ]  
 The **ToolbarVisible** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether the tool bar is displayed in a **GEXPreview** control, as described in Settings. |

### Settings

 The settings for *value* are:

| Setting | Description |
| --- | --- |
| **True** | **GEXPreview** control's tool bar is displayed. |
| **False** | **GEXPreview** control's tool bar isn't displayed. |

### Data Type

 Boolean

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [PageSetupText Property](#pagesetuptext-property-gexpreview-control), [PrintText Property](#printtext-property-gexpreview-control), [CloseButtonText Property](#closebuttontext-property-gexpreview-control)

## TotalPages Property (GEXPreview Control)

Returns the number of pages in a **GEXPreview** control. Read Only.

### Syntax

 object.**TotalPages**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 Use this property if you want to display in a status bar or another control the number of pages in a document being previewed using a **GEXPreview** control.

### Data Type

 Long

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [NextPage Method](Methods.md#nextpage-method-gexpreview-control), [PreviousPage Method](Methods.md#previouspage-method-gexpreview-control), [CurrentPage Property](#currentpage-property-gexpreview-control), [Repaginate Method](Methods.md#repaginate-method-gexpreview-control)

## Zoom Property (GEXPreview Control)

Returns or sets a value that indicates how **GEXPreview** control should display pages.

### Syntax

 *object*.**Zoom** [= *value*]  
 The **Zoom** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines how **GEXPreview** control should display pages, as described in Settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexZoomCurrentSize** | 0 | Default. **GEXPreview** control displays the page at its actual size. |
|  **jgexZoomOnePage** | 1 | **GEXPreview** control scales the page so it can be seen entirely. |
|  **jgexZoomTwoPages** | 2 | **GEXPreview** control displays two pages. |

### Data Type

 **jgexZoomConstants**

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [PreviousPage Method](Methods.md#previouspage-method-gexpreview-control), [NextPage Method](Methods.md#nextpage-method-gexpreview-control), [ZoomChanged Event](Events.md#zoomchanged-event-gexpreview-control)
