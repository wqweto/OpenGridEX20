# JSFormatStyle Object

## JSFormatStyle Object

![](../../images/JSFormatStyle.jpg)  
 Contains format settings.

### Syntax

 *object*.**FormatStyle**  
 The object placeholder represents an object expression that evaluates to a **JSFmtCondition** object.  
 *gridex*.**FormatStyles** (*index*)  
 The **JSFormatStyle** object syntax has these parts:

| Part | Description |
| --- | --- |
| *gridex* | An object expression that evaluates to a **GridEX** control. |
| *index* | Either a Long or string that uniquely identifies a member of an object collection. An integer would be the value of the **Index** property; a string would be the value of the **Name** property. |

### Remarks

 A **JSFormatStyle** object can either be linked to a **JSFmtCondition** or a member of the **JSFormatStyles** collection.  
 A **GridEX** control has some predefined format styles that can not be removed but their properties can be changed. The predefined styles are: Default, OddRow, EvenRow, RowGroup and PreviewRow.  
 In adition, with a **JSFormatStyle** object, you can specify the format settings for rows that meet the criteria specified in its **JSFmtCondition** object parent.  
 To use a **JSFormatStyle** object, you can use the FormatStyle property of a **JSFmtCondition** object or you can use the **FormatStyles** property of the **GridEX** control to obtain them.

**Note** A **JSFormatStyle** obtained from a **JSFmtCondition** object is not present in the **JSFormatStyles** collection.

**See Also:** [FormatStyles Property](../Properties.md#formatstyles-property-gridex-control), [FormatStyle Property](JSFmtCondition-Object.md#formatstyle-property-jsfmtcondition-object), [JSFormatStyles Collection](JSFormatStyles-Collection.md#jsformatstyles-collection), [Item Property](JSFormatStyles-Collection.md#item-property-jsformatstyles-collection)

## BackColor Property (JSFormatStyle Object)

Returns or sets the background color of a **JSFormatStyle** object.

### Syntax

 *object*.**BackColor** [ = *color* ]  
 The **BackColor** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A numeric expression that specifies the color. |

### Remarks

 The **BackColor** property refers to the color used to display the cells, rows or columns where a **FormatStyle** is selected. To clear the **BackColor** property of a **JSFormatStyle**, the **Reset** method must be used.  
 If the **BackColor** property is not set in a **JSFormatStyle** object, the **BackColor** property of the **GridEX** control is used instead.

### Data Type

 OLE_COLOR

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [BackColorBkg Property](../Properties.md#backcolorbkg-property-gridex-control), [BackColorGBBox Property](../Properties.md#backcolorgbbox-property-gridex-control), [BackColorHeader Property](../Properties.md#backcolorheader-property-gridex-control), [BackColorInfoText Property](../Properties.md#backcolorinfotext-property-gridex-control), [BackColor Property](../Properties.md#backcolor-property-gridex-control), [BackColorRowGroup Property](../Properties.md#backcolorrowgroup-property-gridex-control), [ForeColorHeader Property](../Properties.md#forecolorheader-property-gridex-control), [ForeColorInfoText Property](../Properties.md#forecolorinfotext-property-gridex-control), [ForeColor Property](#forecolor-property-jsformatstyle-object), [ForeColorRowGroup Property](../Properties.md#forecolorrowgroup-property-gridex-control), [MaskColor Property](../Properties.md#maskcolor-property-gridex-control), [RowColorEven Property](../Properties.md#rowcoloreven-property-gridex-control), [RowColorOdd Property](../Properties.md#rowcolorodd-property-gridex-control), [GridLinesColor Property](../Properties.md#gridlinescolor-property-gridex-control)  
**Example:** [FormatStyles Example](../../Examples.md#formatstyles-example)

## BackGroundPicture Property (JSFormatStyle Object)

Returns or sets the background picture of a **JSFormatStyle** object.

### Syntax

 *object*.**BackGroundPicture** [ = *picture* ]  
 The **BackGroundPicture** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *picture* | A Picture object that represents the picture to be used as the background picture of a **JSFormatStyle** object. |

### Remarks

 The **BackGroundPicture** property refers to the picture used to draw the background of cells, rows or columns where a **FormatStyle** is selected.  
 Use the**JSFormatStyle**'s **Reset** method to clear the **BackGroundPicture** property.  
 This property must be used in conjunction with **DrawModeBackGroundPicture** property to tell **GridEX** control how a background picture will be draw in a cell.

### Data Type

 **StdPicture**

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)

## DrawModeBackGroundPicture Property (JSFormatStyle Object)

Returns or sets a value or constant that determines the way a background picture is drawn.

### Syntax

 *object*.**DrawModeBackGroundPicture** [ = *value* ]  
 The **DrawModeBackGroundPicture** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines the draw mode for background pictures, as described in Settings. |

### Settings

 The settings for value are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexDMCenter** | 0 | Default. Background picture is centered in the cell boundaries. |
|  **jgexDMTile** | 1 | Background is tiled in the cell boundaries. |
|  **jgexDMStretch** | 2 | Background picture is resized to fit the cell boundaries. |

### Remarks

 Use the **DrawModeBackGroundPicture** property to tell **GridEX** control how to draw the background picture in a cell.  
 This property must be used in conjunction with **BackGroundPicture** property, if **BackGroundPicture** is not set or undefined, this property has no effect.

### Data Type

 **jgexDrawModePictureBackgroundConstants**

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [Picture Property](#picture-property-jsformatstyle-object), [PictureDrawMode Property](#picturedrawmode-property-jsformatstyle-object), [PictureHorzAlignment Property](#picturehorzalignment-property-jsformatstyle-object), [PictureVertAlignment Property](#picturevertalignment-property-jsformatstyle-object), [BackGroundPicture Property](#backgroundpicture-property-jsformatstyle-object)

## FontBold Property (JSFormatStyle Object)

Returns or sets font bold style used in a **JSFormatStyle** Object.

### Syntax

 object.**FontBold** [ = *boolean* ]  
 The **FontBold**property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *boolean* | A Boolean expression specifying the font style as described in Settings. |

### Settings

 The settings for *boolean* are:

| Setting | Description |
| --- | --- |
| **True** | Turns on the formatting in that style. |
| **False** | (Default) Turns off the formatting in that style. |

### Remarks

 Use this font property to format text in rows that meet the criteria specified in **JSFmtCondition** object settings or in cells where the **JSFormatStyle** will be applied.

### Data Type

 Boolean

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [Font Property](../Properties.md#font-property-gridex-control), [FontCharset Property](#fontcharset-property-jsformatstyle-object), [FontItalic Property](#fontitalic-property-jsformatstyle-object), [FontName Property](#fontname-property-jsformatstyle-object), [FontSize Porperty](#fontsize-porperty-jsformatstyle-object), [FontStrikeThru Property](#fontstrikethru-property-jsformatstyle-object), [FontUnderline Property](#fontunderline-property-jsformatstyle-object)  
**Example:** [CellStyle Example](../../Examples.md#cellstyle-example)

## FontCharset Property (JSFormatStyle Object)

Sets or returns the character set of the font used in a **JSFormatStyle** Object.  
 Syntax  
 *object*.**FontCharset** [ = *value* ]  
 The **FontCharset** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A integer expression specifying the character set of the font used in a **JSFormatStyle** Object. |

### Data Type

 Integer

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [Font Property](../Properties.md#font-property-gridex-control), [FontItalic Property](#fontitalic-property-jsformatstyle-object), [FontName Property](#fontname-property-jsformatstyle-object), [FontSize Porperty](#fontsize-porperty-jsformatstyle-object), [FontStrikeThru Property](#fontstrikethru-property-jsformatstyle-object), [FontUnderline Property](#fontunderline-property-jsformatstyle-object), [FontBold Property](#fontbold-property-jsformatstyle-object)

## FontItalic Property (JSFormatStyle Object)

Returns or sets font italic style used in a **JSFormatStyle** Object.

### Syntax

 *object*.**FontItalic** [ = *boolean* ]  
 The **FontItalic** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *boolean* | A Boolean expression specifying the font style as described in Settings. |

### Settings

 The settings for *boolean* are:

| Setting | Description |
| --- | --- |
| **True** | Turns on the formatting in that style. |
| **False** | (Default) Turns off the formatting in that style. |

### Remarks

 Use this font property to format text in rows that meet the criteria specified in **JSFmtCondition** object settings or in cells where the **JSFormatStyle** will be applied.

### Data Type

 Boolean

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [Font Property](../Properties.md#font-property-gridex-control), [FontCharset Property](#fontcharset-property-jsformatstyle-object), [FontName Property](#fontname-property-jsformatstyle-object), [FontSize Porperty](#fontsize-porperty-jsformatstyle-object), [FontStrikeThru Property](#fontstrikethru-property-jsformatstyle-object), [FontUnderline Property](#fontunderline-property-jsformatstyle-object), [FontBold Property](#fontbold-property-jsformatstyle-object)  
**Example:** [CellStyle Example](../../Examples.md#cellstyle-example)

## FontName Property (JSFormatStyle Object)

Sets or returns the typeface name of the font used in a **JSFormatStyle** Object.

### Syntax

 *object*.**FontName** [ = *value* ]  
 The **FontName** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A String expression specifying the typeface name of the font used in a **JSFormatStyle** Object. |

### Data Type

 String

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [Font Property](../Properties.md#font-property-gridex-control), [FontCharset Property](#fontcharset-property-jsformatstyle-object), [FontItalic Property](#fontitalic-property-jsformatstyle-object), [FontSize Porperty](#fontsize-porperty-jsformatstyle-object), [FontStrikeThru Property](#fontstrikethru-property-jsformatstyle-object), [FontUnderline Property](#fontunderline-property-jsformatstyle-object), [FontBold Property](#fontbold-property-jsformatstyle-object)

## FontSize Porperty (JSFormatStyle Object)

Returns or sets the size of the font used in a **JSFormatStyle** Object.

### Syntax

 *object*.**FontSize** [ = *value* ]  
 The **FontSize** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A numeric expression specifying the size of the font in points. |

### Data Type

 Currency

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [Font Property](../Properties.md#font-property-gridex-control), [FontCharset Property](#fontcharset-property-jsformatstyle-object), [FontItalic Property](#fontitalic-property-jsformatstyle-object), [FontName Property](#fontname-property-jsformatstyle-object), [FontStrikeThru Property](#fontstrikethru-property-jsformatstyle-object), [FontUnderline Property](#fontunderline-property-jsformatstyle-object), [FontBold Property](#fontbold-property-jsformatstyle-object)

## FontStrikeThru Property (JSFormatStyle Object)

Returns or sets font strikethrough style used in a **JSFormatStyle** Object.

### Syntax

 *object*.**FontStrikeThru** [ = *boolean* ]  
 The **FontStrikeThru** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *boolean* | A Boolean expression specifying the font style as described in Settings. |

### Settings

 The settings for *boolean* are:

| Setting | Description |
| --- | --- |
| **True** | Turns on the formatting in that style. |
| **False** | (Default) Turns off the formatting in that style. |

### Remarks

 Use this font property to format text in rows that meet the criteria specified in **JSFmtCondition** object settings or in cells where the **JSFormatStyle** will be applied.

### Data Type

 Boolean

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [Font Property](../Properties.md#font-property-gridex-control), [FontCharset Property](#fontcharset-property-jsformatstyle-object), [FontItalic Property](#fontitalic-property-jsformatstyle-object), [FontName Property](#fontname-property-jsformatstyle-object), [FontSize Porperty](#fontsize-porperty-jsformatstyle-object), [FontUnderline Property](#fontunderline-property-jsformatstyle-object), [FontBold Property](#fontbold-property-jsformatstyle-object)

## FontUnderline Property (JSFormatStyle Object)

Returns or sets font underline style used in a **JSFormatStyle** Object.

### Syntax

 *object*.**FontUnderline** [ = *boolean* ]  
 The **FontUnderline** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *boolean* | A Boolean expression specifying the font style as described in Settings. |

### Settings

 The settings for *boolean* are:

| Setting | Description |
| --- | --- |
| **True** | Turns on the formatting in that style. |
| **False** | (Default) Turns off the formatting in that style. |

### Remarks

 Use this font property to format text in rows that meet the criteria specified in **JSFmtCondition** object settings or in cells where the **JSFormatStyle** will be applied.

### Data Type

 Boolean

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [Font Property](../Properties.md#font-property-gridex-control), [FontCharset Property](#fontcharset-property-jsformatstyle-object), [FontItalic Property](#fontitalic-property-jsformatstyle-object), [FontName Property](#fontname-property-jsformatstyle-object), [FontSize Porperty](#fontsize-porperty-jsformatstyle-object), [FontStrikeThru Property](#fontstrikethru-property-jsformatstyle-object), [FontBold Property](#fontbold-property-jsformatstyle-object)  
**Example:** [CellStyle Example](../../Examples.md#cellstyle-example)

## ForeColor Property (JSFormatStyle Object)

Returns or sets the foreground color of the **JSFormatStyle** object.

### Syntax

 *object*.**ForeColor** [ = *color* ]  
 The **ForeColor** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A numeric expression that specifies the color. |

### Remarks

 The ForeColor property refers to the text color used to display the rows that meet the criteria specified in the **JSFmtCondition** settings or in cells where the **JSFormatStyle** will be applied. .

### Data Type

 OLE_COLOR

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [BackColorBkg Property](../Properties.md#backcolorbkg-property-gridex-control), [BackColorGBBox Property](../Properties.md#backcolorgbbox-property-gridex-control), [BackColorHeader Property](../Properties.md#backcolorheader-property-gridex-control), [BackColorInfoText Property](../Properties.md#backcolorinfotext-property-gridex-control), [BackColor Property](../Properties.md#backcolor-property-gridex-control), [BackColor Property](#backcolor-property-jsformatstyle-object), [BackColorRowGroup Property](../Properties.md#backcolorrowgroup-property-gridex-control), [ForeColorHeader Property](../Properties.md#forecolorheader-property-gridex-control), [ForeColorInfoText Property](../Properties.md#forecolorinfotext-property-gridex-control), [ForeColor Property](../Properties.md#forecolor-property-gridex-control), [ForeColorRowGroup Property](../Properties.md#forecolorrowgroup-property-gridex-control), [MaskColor Property](../Properties.md#maskcolor-property-gridex-control), [RowColorEven Property](../Properties.md#rowcoloreven-property-gridex-control), [RowColorOdd Property](../Properties.md#rowcolorodd-property-gridex-control), [GridLinesColor Property](../Properties.md#gridlinescolor-property-gridex-control)  
**Example:** [FormatStyles Example](../../Examples.md#formatstyles-example)

## Name Property (JSFormatStyle Object)

Returns a string that uniquely identifies a member in a **JSFormatStyles** collection.

### Syntax

 *object*.**Name**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 If the string is not unique within the **JSFormatStyles** collection, an error will occur when adding a **JSFormatStyle** object to the **JSFormatStyles** collection.  
 You can set the **Name** property when you add the object to the collection using the **JSFormatStyles**' **Add** method.  
 Once a **JSFormatStyle** is added, the **Name** property is Read-Only

### Data Type

 String

Notes The name of a predefined style cannot be used as a name for a new **JSFormatStyle**.The predefined styles are: Default, OddRow, EvenRow, RowGroup and PreviewRow

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [Add Method](JSFormatStyles-Collection.md#add-method-jsformatstyles-collection), [Item Property](JSFormatStyles-Collection.md#item-property-jsformatstyles-collection), [HeaderStyle Property](JSColumn-Object.md#headerstyle-property-jscolumn-object), [CellStyle Property](JSColumn-Object.md#cellstyle-property-jscolumn-object), [CellStyle Property](JSRowData-Object.md#cellstyle-property-jsrowdata-object), [RowStyle Property](JSRowData-Object.md#rowstyle-property-jsrowdata-object)

## Picture Property (JSFormatStyle Object)

Returns or sets a picture of a **JSFormatStyle** object.

### Syntax

 *object*.**Picture** [ = *picture* ]  
 The **Picture** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *picture* | A Picture object that represents the picture to be used as the foreground picture of a **JSFormatStyle** object. |

### Remarks

 The **Picture** property refers to the picture to be draw in cells, rows or columns where a **FormatStyle** is selected.  
 Use the **JSFormatStyle**'s **Reset** method to clear the **Picture** property.  
 This property must be used in conjunction with **PictureHorzAlignment** and **PictureVertAlignment** properties to tell the **GridEX** control where the picture will be drawn.

### Data Type

 **StdPicture**

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [DrawModeBackGroundPicture Property](#drawmodebackgroundpicture-property-jsformatstyle-object), [PictureDrawMode Property](#picturedrawmode-property-jsformatstyle-object), [PictureHorzAlignment Property](#picturehorzalignment-property-jsformatstyle-object), [PictureVertAlignment Property](#picturevertalignment-property-jsformatstyle-object), [BackGroundPicture Property](#backgroundpicture-property-jsformatstyle-object)

## PictureDrawMode Property (JSFormatStyle Object)

Returns or sets a value or constant that determines how **GridEX** control draws cells with pictures.

### Syntax

 *object*.**PictureDrawMode** [ = *value* ]  
 The **PictureDrawMode** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines how **GridEX** control draws cells with pictures, as described in Settings. |

### Settings

 The settings for value are:

| Constant | Value | Description |
| --- | --- | --- |
| jgexTextOnly | 0 | Default. Only text is draw in a cell even if the format style picture or CellPicture in for the row are set. |
| jgexPictureOnly | 1 | Only the picture is drawn in the cell. |
| jgexPictureAndText | 2 | Picture and text are drawn in the cell. |

### Remarks

 This property must be used in conjunction with **Picture** property of a **JSFormatStyle** or the **CellPicture** property in the **JSRowData** object.

### Data Type

 **jgexPictureDrawModeConstants**

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [DrawModeBackGroundPicture Property](#drawmodebackgroundpicture-property-jsformatstyle-object), [Picture Property](#picture-property-jsformatstyle-object), [PictureHorzAlignment Property](#picturehorzalignment-property-jsformatstyle-object), [PictureVertAlignment Property](#picturevertalignment-property-jsformatstyle-object), [BackGroundPicture Property](#backgroundpicture-property-jsformatstyle-object)

## PictureHorzAlignment Property (JSFormatStyle Object)

Returns or sets a value or constant that determines the horizontal alignment of a picture in a **JSFormatStyle** object.

### Syntax

 *object*.**PictureHorzAlignment** [ = *value* ]  
 The **PictureHorzAlignment** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines the horizontal alignment of a picture, as described in Settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexHPALeft** | 0 | Default. Picture is draw at the left of the cell. |
|  **jgexHPACenter** | 1 | Picture is drawn horizontally centered in a cell. |
|  **jgexHPARight** | 2 | Picture is draw at the right of a cell. |
|  **jgexHPALeftOfText** | 3 | Picture is draw at the left of a cell and text at the the right of the picture. |
|  **jgexHPARightOfText** | 4 | Picture is draw at the right of the text. |

### Remarks

 Use the **PictureHorzAlignment** and **PictureVertALignment** properties to tell **GridEX** control how to draw the picture in a cell.  
 This property must be used in conjunction with **Picture** and **PictureVertAlignment** properties, if **Picture** is not set or undefined, this property has no effect unless **CellPicture** property is set in the **RowFormat** event for a row.

### Data Type

 **jgexHorzPictureAlignmentConstants**

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [DrawModeBackGroundPicture Property](#drawmodebackgroundpicture-property-jsformatstyle-object), [Picture Property](#picture-property-jsformatstyle-object), [PictureDrawMode Property](#picturedrawmode-property-jsformatstyle-object), [PictureVertAlignment Property](#picturevertalignment-property-jsformatstyle-object), [BackGroundPicture Property](#backgroundpicture-property-jsformatstyle-object)

## PictureVertAlignment Property (JSFormatStyle Object)

Returns or sets a value or constant that determines the vertical alignment of a picture in a **JSFormatStyle** object.

### Syntax

 *object*.**PictureVertAlignment** [ = *value* ]  
 The **PictureVertAlignment** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines the vertical alignment of a picture, as described in Settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexVPATop** | 0 | Default. Picture is draw at the top of the cell. |
|  **jgexVPACenter** | 1 | Picture is drawn vertically centered in a cell. |
|  **jgexVPABottom** | 2 | Picture is draw at the bottom of a cell. |
|  **jgexVPATopOfText** | 3 | Picture is draw at the top of a cell and text at the the bottom of the picture. |
|  **jgexVPABottomOfText** | 4 | Picture is draw at the bottom of the text. |

### Remarks

 Use the **PictureHorzAlignment** and **PictureVertALignment** properties to tell **GridEX** control how to draw the picture in a cell.  
 This property must be used in conjunction with **Picture** and **PictureHorzAlignment** properties, if **Picture** is not set or undefined, this property has no effect unless **CellPicture** property is set in the **RowFormat** event for a row.

### Data Type

 **jgexVertPictureAlignmentConstants**

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [DrawModeBackGroundPicture Property](#drawmodebackgroundpicture-property-jsformatstyle-object), [Picture Property](#picture-property-jsformatstyle-object), [PictureDrawMode Property](#picturedrawmode-property-jsformatstyle-object), [PictureHorzAlignment Property](#picturehorzalignment-property-jsformatstyle-object), [BackGroundPicture Property](#backgroundpicture-property-jsformatstyle-object)

## TextAlignment Property (JSFormatStyle Object)

Returns or sets a value that determines the alignment of text in a **JSFormatStyle** object.

### Syntax

 *object*.**TextAlignment** [ = *value* ]  
 The **TextAlignment** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| value | A value or constant that specifies the type of alignment of a **JSFormatStyle**, as described in Settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexAlignLeft** | 0 | Text is left aligned. |
|  **jgexAlignCenter** | 1 | Text is centered. |
|  **jgexAlignRight** | 2 | Text is right aligned. |

### Data Type

 **jgexAlignmentConstants**

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [TextAlignment Property](JSColumn-Object.md#textalignment-property-jscolumn-object), [HeaderAlignment Property](JSColumn-Object.md#headeralignment-property-jscolumn-object)

## Reset Method (JSFormatStyle Object)

Clears all properties in a **JSFormatStyle** object.

### Syntax

 *object*.**Reset**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 When calling this method from the default format style, **BackColor**, **ForeColor** and **Font** properties won't be cleared.

**Applies To:** [JSFormatStyle Object](#jsformatstyle-object)  
**See Also:** [Remove Method](JSFormatStyles-Collection.md#remove-method-jsformatstyles-collection)
