# GridEX Control — Properties

## ActAsDropDown Property (GridEX Control)

Returns or sets a value indicating whether the **GridEX** control will behave as a drop down list for a column in another **GridEX** control.

### Syntax

 *object*.**ActAsDropDown** = [*value*]  
 The **ActAsDropDown** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether a control will be a drop down list for a column in another grid, as described in Settings. |

### Settings

 The settings for value are:

| Value | Description |
| --- | --- |
| **False** | Default. The control acts as a grid |
| **True** | The control acts as a DropDown List for a column in a different **GridEX** control. |

### Remarks

 If the **ActAsDropDown** property is **True**, the **GridEX** control can only be used in conjunction with another **GridEX** control.  
 The "parent" control must have a column with **EditType = jgexCombo** and the **DropDownControl** property must be set to the **ActAsDropDown** GridEX control.  
 When **ActAsDropDown** property is set to True, some properties and methods are limited.  
 The limited features in an **ActAsDropDown** control are:

| **AllowAddNew:** | Must be equal to **False** |
| --- | --- |
| **AllowDelete:** | Must be equal to **False** |
| **AllowEdit:** | Must be equal to **False** |
| **ContinuousScroll:** | Must be equal to **True** |
| **DetectRowDrag:** | Must be equal to **False** |
| **GroupByBoxVisible:** | Must be equal to **False** |
| **HideSelection:** | Must be equal to **jgexHighLightNormal** |
| **MultiSelect:** | Must be equal to **False** |
| **View:** | Must be equal to **jgexTable** |

Also, the control can not be grouped while is acting as a dropdown list.  
 When you change the **ActAsDropDown** property to **True**, **GridEX** automatically changes the above properties to their defaults and any further attempts to change those properties will result in a trappable error.  
 In an **ActAsDropDown** control, you need to specify the **BoundColumnIndex** and the **ReplaceColumnIndex** properties to indicate the control how it will interact with the data from the "parent column".  
 **BoundColumnIndex** represents the **Index** or key of the column, that has the data to be written to the parent column, when the user selects an entry from the list.  
 **ReplaceColumnIndex** represents the **Index** or key of the column with the text to be displayed instead of the value in the parent column.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BoundColumnIndex Property](#boundcolumnindex-property-gridex-control), [ReplaceColumnIndex Property](#replacecolumnindex-property-gridex-control), [EditType Property](Objects/JSColumn-Object.md#edittype-property-jscolumn-object), [ReplaceValues Property](Objects/JSColumn-Object.md#replacevalues-property-jscolumn-object)  
**Example:** [ActAsDropDown Example](../Examples.md#actasdropdown-example)

## ADORecordset Property (GridEX Control)

Returns or sets an **ADOR.Recordset** object defined by **GridEX** control's properties or by an existing **ADOR.Recordset** object.

### Syntax

 *object*.**ADORecordset** [ = *value* ]  
 The **ADORecordset** property syntax has these parts:

| **Part** | **Description** |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | An object variable containing an **ADOR.Recordset** or an **ADODB.Recordset** object |

### Remarks

 The **GridEX** control is automatically initialized when the control first appears. If the **DataMode** property setting is **jgexADO** and the **DatabaseName**, **LockType**, **Options**, **CursorLocation** **RecordSource** and **RecordsetType** properties are valid, or if you set these properties at run time and use the **Rebind** method, the control attempts to create a new Recordset object based on those properties. This Recordset is accessible through the **GridEX** control's **ADORecordset** property. However, if one or more of these properties are set incorrectly at design time, an error event is sent to the application when the control attempts to use the properties to create the **Recordset** object.  
 You can also request the type of **Recordset** to be created by setting the **GridEX** control's **RecordsetType** property. If you don't request a specific type, a Keyset-type cursor is created. Using the **RecordsetType** property, you can request to create either a Static-, or Keyset-type cursor. However, if the control cannot create the type requested, an error event occurs.  
 If you create a **Recordset** object using code, you can set the **ADORecordset** property of the control to this new **Recordset**. Dynamic or Forward-only cursors can not be used within a **GridEX** control, trying to assign a **Recordset** of these types will generate an error.  
 When a new **Recordset** is assigned to the **ADORecordset** property, the control releases any existing **Recordset**.  
 If you set this property at runtime, the control’s **DataMode** property changes to **jgexADO** no matter the previous **DataMode** property of the control. However if you try to read the **ADORecordset** property when the **DataMode** property of the control is other than **jgexADO**, a trappable error occurs.  
 **Note** You must avoid the use of **Close** method or the **Filter** property in any variable set to **ADORecordset** object because it will make invalid this object for any further operations until the **Rebind** method is used. If need to filter the recordset, you can clone the recordset and apply the filter to the clone.

### Data Type

 **Object**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [CursorLocation Property](#cursorlocation-property-gridex-control), [DatabaseName Property](#databasename-property-gridex-control), [DataMode Property](#datamode-property-gridex-control), [LockType Property](#locktype-property-gridex-control), [Rebind Method](Methods.md#rebind-method-gridex-control), [Recordset Property](#recordset-property-gridex-control), [RecordsetType Property](#recordsettype-property-gridex-control), [RecordSource Property](#recordsource-property-gridex-control)  
**Example:** [ADORecordset Example](../Examples.md#adorecordset-example)

## AllowAddNew Property (GridEX Control)

Returns or sets a value indicating whether the user can add new records to the **Recordset**object underlying a **GridEX** control.

### Syntax

 *object*.**AllowAddNew** [ = *value* ]  
 The **AllowAddNew** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether a user can add new records, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | Users can add records to the underlying **Recordset** object of the **GridEX** control. |
| **False** | Users can't add records to the underlying **Recordset** object of the **GridEX** control. |

### Remarks

 If the **AllowAddNew** property is **True**, the **GridEX** control allows the users to add new records, either using the top row or the last row in the grid, according to the **NewRowPos** property. If the **AllowAddNew** property is **False**, the user cannot add new records, and neither a blank line at the bottom nor a top new row is displayed.  
 The underlying **Recordset** may not accept insertions even if the **AllowAddNew** property is **True**. In this case, the control will act as if the **AllowAddNew** property was set to **False**.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AllowEdit Property](#allowedit-property-gridex-control), [NewRowPos Property](#newrowpos-property-gridex-control)

## AllowCardSizing Property (GridEX Control)

Returns or sets a value indicating whether the user can resize cards in a **GridEX** control.

### Syntax

 *object*.**AllowCardSizing** [ = *value* ]  
 The **AllowCardSizing** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether a user can size cards, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | Users can resize cards interactively in a **GridEX** control while in card view. |
| **False** | Users cannot resize cards in a **GridEX** control while in card view. |

### Remarks

 If the **AllowCardSizing** property is **True**, the **GridEX** control allows the user to change the width of the cards, dragging the bar drawn between each card while the control is in card view.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [CardWidth Property](#cardwidth-property-gridex-control), [CardResize Event](Events.md#cardresize-event-gridex-control)  
**Example:** [View Example](../Examples.md#view-example)

## AllowColumnDrag Property (GridEX Control)

Returns or sets a value indicating whether the user can drag columns or groups in a **GridEX** control.

### Syntax

 *object*.**AllowColumnDrag** [ = *value* ]  
 The **AllowColumnDrag** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether a user can drag columns or groups while a **GridEX** control is in table view, as described in Settings. |

### Settings

 The settings for *value* are:

| Setting | Description |
| --- | --- |
| **True** | Users can drag columns while a **GridEX** control is in table view. |
| **False** | Users cannot drag columns in the **GridEX** control. |

### Remarks

 If the **AllowColumnDrag** property setting is **True**, the **GridEX** control allows the users to drag a header to change the column position and/or the group settings.  
 The **AllowColumnDrag** property has no effect if the **View** property setting is set to **jgexCard**.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BeforeGroupDrag Event](Events.md#beforegroupdrag-event-gridex-control), [BeforeColumnDrag Event](Events.md#beforecolumndrag-event-gridex-control)

## AllowDelete Property (GridEX Control)

Returns or sets a value indicating whether the user can delete records from the underlying **Recordset** object in a **GridEX** control.

### Syntax

 *object*.**AllowDelete** [ = *value* ]  
 The **AllowDelete** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether a user can delete records, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | Users can delete records from the underlying **Recordset** object in the **GridEX** control. |
| **False** | Users cannot delete records from the underlying **Recordset** object in the **GridEX** control. |

### Remarks

 You can use the **AllowDelete** property to prevent the user from deleting records from the **Recordset** through interaction with the **GridEX** control.  
 The underlying **Recordset** may not enable deletions even if the **AllowDelete** property is **True** for the **GridEX** control. In this case, the control will act as if the **AllowDelete** property was set to False. If an error occurs when the control attempts to physically delete the record, the **Error** event is triggered.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BeforeDelete Event](Events.md#beforedelete-event-gridex-control), [BeforeDeleteEx Event](Events.md#beforedeleteex-event-gridex-control), [AfterDelete Event](Events.md#afterdelete-event-gridex-control), [Error Event](Events.md#error-event-gridex-control), [Delete Method](Methods.md#delete-method-gridex-control)

## AllowEdit Property (GridEX Control)

Returns or sets a value indicating whether a user can modify the data contained in the **GridEX**  
 control.

### Syntax

 *object*.**AllowEdit** [ = *value* ]  
 The **AllowEdit** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether the user can change data, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | The user can modify data in the **GridEX** control |
| **False** | The user cannot modify data in the **GridEX** control |

### Remarks

 When the **AllowEdit** property setting is **False**, the user can still scroll through the **GridEX** control and select data, but cannot change any of the values; any attempt to change the data in the control is ignored.  
 You can also use the **JSColumn** object properties to make individual columns of the **GridEX** control read-only, but the AllowEdit property setting takes precedence over the **JSColumn** settings (without changing the **JSColumn** settings).  
 **Note:** The **Recordset** object may not enable edits even if **AllowEdit** is True for the **GridEX** control; in this case, the control acts as if the **AllowEdit** property was set to **False**. If the Recordset is marked as updateable and for some reason the edited record could not be updated an **Error** event is fired.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AfterColEdit Event](Events.md#aftercoledit-event-gridex-control), [AllowAddNew Property](#allowaddnew-property-gridex-control), [BeforeColEdit Event](Events.md#beforecoledit-event-gridex-control), [EditType Property](Objects/JSColumn-Object.md#edittype-property-jscolumn-object)

## AllowRowSizing Property (GridEX Control)

Returns or sets a value indicating whether a user can resize rows in a GridEX control.

### Syntax

 *object*.**AllowRowSizing** [ = *value*]  
 The **AllowRowSizing** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether a user can resize rows, as described in Settings. |

### Settings

 The settings for *value* are:

| Setting | Description |
| --- | --- |
| **True** | Users can resize rows interactively in a **GridEX** control. |
| **False** | Users cannot resize rows in a **GridEX** control. |

### Remarks

 If the **AllowRowSizing** property is **True**, the **GridEX** control allows the user to change the height of the rows, dragging the bar drawn between each row header.

**Note**: In order to allow a user to resize rows, the **RowHeaders** property must be also **True**.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RowHeight Property](#rowheight-property-gridex-control)

## AutomaticArrange Property (GridEX Control)

Returns or sets a value indicating whether the **GridEX** control will reposition sorted or grouped records after the user has updated the data in a record.

### Syntax

 *object*.**AutomaticArrange** [ = *value* ]  
 The **AutomaticArrange** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether the updated record will be repositioned, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | The **GridEX** control will change the position and group of the updated record. |
| **False** | The **GridEX** control will not change the position and group of the updated record. |

### Remarks

 If the **AutomaticArrange** property setting is **True**, the **GridEX** control will change the position and/or grouping of a record if necessary. If the **AutomaticArrange** property setting is **False** the edited record will remain in its position until the program requests a new sorting or grouping operation.  
 Even if the **AutomaticArrange** property setting is **True**, the control will not reposition records when changes were made anywhere except through user interaction with the **GridEX** control. If you change a record in the **Recordset** clone, and you want to update the position and grouping for the record, you must make the modified record the current row by moving to it, and then call the **GridEX** control’s **Update** method.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Groups Property](#groups-property-gridex-control), [SortKeys Property](#sortkeys-property-gridex-control), [RefreshSort Method](Methods.md#refreshsort-method-gridex-control), [RefreshGroups Method](Methods.md#refreshgroups-method-gridex-control)

## AutomaticSort Property (GridEX Control)

Returns or sets a value that determines whether the **GridEX** control should sort values automatically when the user clicks on a column header.

### Syntax

 *object*.**AutomaticSort** = [ *value* ]  
 The **AutomaticSort** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether the **GridEX** control should sort values automatically when the user clicks on a column header as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | The **GridEX** control sorts records as the result of a column header click. |
| **False** | Default. The **GridEX** control doesn't sort records when the user clicks in a column. |

### Remarks

 In previous versions of the **GridEX** control, you had to write code in the **ColumnHeaderClick** and **GroupByBoxHeaderClick** events to sort records. Setting this property to **True**, **GridEX** control handles the sorting when the user clicks in a column header or in a group by box header.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Groups Property](#groups-property-gridex-control), [SortKeys Property](#sortkeys-property-gridex-control), [ColumnHeaderClick Event](Events.md#columnheaderclick-event-gridex-control), [GroupByBoxHeaderClick Event](Events.md#groupbyboxheaderclick-event-gridex-control)

## BackColor Property (GridEX Control)

Returns or sets the background color for all the cells in a **GridEX** control.

### Syntax

 *object*.**BackColor** [ =*color* ]  
 The **BackColor** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A numeric expression that specifies the color. |

### Remarks

 The **BackColor** property, in table view, refers to the background color of all the cells.  
 The **BackColor** property, in card view, refers to the background color of the card body.

### Data Type

 OLE_COLOR

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BackColorBkg Property](#backcolorbkg-property-gridex-control), [BackColorGBBox Property](#backcolorgbbox-property-gridex-control), [BackColorHeader Property](#backcolorheader-property-gridex-control), [BackColorInfoText Property](#backcolorinfotext-property-gridex-control), [BackColor Property](Objects/JSFormatStyle-Object.md#backcolor-property-jsformatstyle-object), [BackColorRowGroup Property](#backcolorrowgroup-property-gridex-control), [ForeColorHeader Property](#forecolorheader-property-gridex-control), [ForeColorInfoText Property](#forecolorinfotext-property-gridex-control), [ForeColor Property](Objects/JSFormatStyle-Object.md#forecolor-property-jsformatstyle-object), [ForeColorRowGroup Property](#forecolorrowgroup-property-gridex-control), [ForeColor Property](#forecolor-property-gridex-control), [MaskColor Property](#maskcolor-property-gridex-control), [RowColorEven Property](#rowcoloreven-property-gridex-control), [RowColorOdd Property](#rowcolorodd-property-gridex-control), [GridLinesColor Property](#gridlinescolor-property-gridex-control)

## BackColorBkg Property (GridEX Control)

Returns or sets the background color of any space without cells in a **GridEX** control.

### Syntax

 *object*.**BackColorBkg** [ =*color* ]  
 The **BackColorBkg**, property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A numeric expression that specifies the color. |

### Remarks

 The **BackColorBkg** property, in table view, refers to the background color of the whitespace in the grid, that is, any space without cells. In card view it represents the background color of the whitespace around the cards.

### Data Type

 OLE_COLOR

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BackColorGBBox Property](#backcolorgbbox-property-gridex-control), [BackColorHeader Property](#backcolorheader-property-gridex-control), [BackColorInfoText Property](#backcolorinfotext-property-gridex-control), [BackColor Property](#backcolor-property-gridex-control), [BackColor Property](Objects/JSFormatStyle-Object.md#backcolor-property-jsformatstyle-object), [BackColorRowGroup Property](#backcolorrowgroup-property-gridex-control), [ForeColorHeader Property](#forecolorheader-property-gridex-control), [ForeColorInfoText Property](#forecolorinfotext-property-gridex-control), [ForeColor Property](Objects/JSFormatStyle-Object.md#forecolor-property-jsformatstyle-object), [ForeColorRowGroup Property](#forecolorrowgroup-property-gridex-control), [MaskColor Property](#maskcolor-property-gridex-control), [RowColorEven Property](#rowcoloreven-property-gridex-control), [RowColorOdd Property](#rowcolorodd-property-gridex-control), [GridLinesColor Property](#gridlinescolor-property-gridex-control)

## BackColorGBBox Property (GridEX Control)

Returns or sets the background color of the “group by” box area in a **GridEX** control.

### Syntax

 *object*.**BackColorGBBox** [ =*color*]  
 The **BackColorGBBox** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A numeric expression that specifies the color. |

### Remarks

 The **BackColorGBBox** property refers to the background color of the “group by” box. The **BackColorGBBox** is not used in card view.

### Data Type

 OLE_COLOR

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BackColorBkg Property](#backcolorbkg-property-gridex-control), [BackColorHeader Property](#backcolorheader-property-gridex-control), [BackColorInfoText Property](#backcolorinfotext-property-gridex-control), [BackColor Property](#backcolor-property-gridex-control), [BackColor Property](Objects/JSFormatStyle-Object.md#backcolor-property-jsformatstyle-object), [BackColorRowGroup Property](#backcolorrowgroup-property-gridex-control), [ForeColorHeader Property](#forecolorheader-property-gridex-control), [ForeColorInfoText Property](#forecolorinfotext-property-gridex-control), [ForeColor Property](Objects/JSFormatStyle-Object.md#forecolor-property-jsformatstyle-object), [ForeColorRowGroup Property](#forecolorrowgroup-property-gridex-control), [ForeColor Property](#forecolor-property-gridex-control), [MaskColor Property](#maskcolor-property-gridex-control), [RowColorEven Property](#rowcoloreven-property-gridex-control), [RowColorOdd Property](#rowcolorodd-property-gridex-control), [GridLinesColor Property](#gridlinescolor-property-gridex-control)

## BackColorHeader Property (GridEX Control)

Returns or sets the background color of the column and row headers in a **GridEX** control.

### Syntax

 *object*.**BackColorHeader** [ =*color* ]  
 The **BackColorHeader** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A numeric expression that specifies the color. |

### Remarks

 The **BackColorHeader** property in table view refers to the background color of the column and row headers.  
 The **BackColorHeader** property in card view represents the background color of each card title bar.

### Data Type

 OLE_COLOR

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BackColorBkg Property](#backcolorbkg-property-gridex-control), [BackColorGBBox Property](#backcolorgbbox-property-gridex-control), [BackColorInfoText Property](#backcolorinfotext-property-gridex-control), [BackColor Property](#backcolor-property-gridex-control), [BackColor Property](Objects/JSFormatStyle-Object.md#backcolor-property-jsformatstyle-object), [BackColorRowGroup Property](#backcolorrowgroup-property-gridex-control), [ForeColorHeader Property](#forecolorheader-property-gridex-control), [ForeColorInfoText Property](#forecolorinfotext-property-gridex-control), [ForeColor Property](Objects/JSFormatStyle-Object.md#forecolor-property-jsformatstyle-object), [ForeColorRowGroup Property](#forecolorrowgroup-property-gridex-control), [MaskColor Property](#maskcolor-property-gridex-control), [RowColorEven Property](#rowcoloreven-property-gridex-control), [RowColorOdd Property](#rowcolorodd-property-gridex-control), [GridLinesColor Property](#gridlinescolor-property-gridex-control)

## BackColorInfoText Property (GridEX Control)

Returns or sets the background color of the rectangle surrounding the information text displayed in the “group by” box.

### Syntax

 *object*.**BackColorInfoText** [ =*color*]  
 The **BackColorInfoText** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A numeric expression that specifies the color. |

### Remarks

 The **BackColorInfoText** property in table view refers to the background color of the rectangle surrounding the information text displayed in the “group by” box when the **GridEX** control has no groups. If the **GroupByBoxInfoText** property is a zero length string, the **GridEX** control will not display any rectangle with this color as its background color.

### Data Type

 OLE_COLOR

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BackColorBkg Property](#backcolorbkg-property-gridex-control), [BackColorGBBox Property](#backcolorgbbox-property-gridex-control), [BackColorHeader Property](#backcolorheader-property-gridex-control), [BackColor Property](#backcolor-property-gridex-control), [BackColor Property](Objects/JSFormatStyle-Object.md#backcolor-property-jsformatstyle-object), [BackColorRowGroup Property](#backcolorrowgroup-property-gridex-control), [ForeColorHeader Property](#forecolorheader-property-gridex-control), [ForeColorInfoText Property](#forecolorinfotext-property-gridex-control), [ForeColor Property](Objects/JSFormatStyle-Object.md#forecolor-property-jsformatstyle-object), [ForeColorRowGroup Property](#forecolorrowgroup-property-gridex-control), [MaskColor Property](#maskcolor-property-gridex-control), [RowColorEven Property](#rowcoloreven-property-gridex-control), [RowColorOdd Property](#rowcolorodd-property-gridex-control), [GridLinesColor Property](#gridlinescolor-property-gridex-control)

## BackColorRowGroup Property (GridEX Control)

Returns or sets the background color of the "group rows" in a **GridEX** control.

### Syntax

 *object*.**BackColorRowGroup** [ =*color* ]  
 The**BackColorRowGroup** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A numeric expression that specifies the color. |

### Remarks

 The **BackColorRowGroup** property refers to the background color of the "group rows" in table view. The **BackColorRowGroup** is not used in card view.

### Data Type

 OLE_COLOR

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BackColorBkg Property](#backcolorbkg-property-gridex-control), [BackColorGBBox Property](#backcolorgbbox-property-gridex-control), [BackColorHeader Property](#backcolorheader-property-gridex-control), [BackColorInfoText Property](#backcolorinfotext-property-gridex-control), [BackColor Property](#backcolor-property-gridex-control), [BackColor Property](Objects/JSFormatStyle-Object.md#backcolor-property-jsformatstyle-object), [ForeColorHeader Property](#forecolorheader-property-gridex-control), [ForeColorInfoText Property](#forecolorinfotext-property-gridex-control), [ForeColor Property](Objects/JSFormatStyle-Object.md#forecolor-property-jsformatstyle-object), [ForeColorRowGroup Property](#forecolorrowgroup-property-gridex-control), [MaskColor Property](#maskcolor-property-gridex-control), [RowColorEven Property](#rowcoloreven-property-gridex-control), [RowColorOdd Property](#rowcolorodd-property-gridex-control), [GridLinesColor Property](#gridlinescolor-property-gridex-control)

## BorderStyle Property (GridEX Control)

Returns or sets the border style for an object.

### Syntax

 *object*.**BorderStyle** [= *value*]  
 The **BorderStyle** property syntax has these parts:

| Part | Description |
| --- | --- |
| object | An object expression that evaluates to an object in the Applies To list. |
| value | A value or constant that determines the border style, as described in Settings. |

### Settings

 The **BorderStyle** property settings are:

| Constant | Setting | Description |
| --- | --- | --- |
| **jgexNone** | 0 | None (no border). |
| **jgexFixed** | 1 | Fixed. The control appears with a 3D border. |
| **jgexSingle3D** | 2 | Single3D. The control appears with a ligth 3D border. |
| **jgexFlat** | 3 | Flat. The control appears with a solid line as a border. |

### Remarks

 If the **BorderStyle** property is set to **jgexFixed**, the control appears with a 3D sunken border.

### Data Type

 Integer

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [CardBorders Property](#cardborders-property-gridex-control), [HeaderStyle Property](#headerstyle-property-gridex-control)

## BoundColumnIndex Property (GridEX Control)

Returns or sets the index or key of the column used to supply a data value to another **GridEX** control.

### Syntax

*object*.**BoundColumnIndex** [ = *value* ]

The **BoundColumnIndex** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A variant that represents the index or key of the column used to supply a data value to another grid control. |

### Remarks

This property is only used when **ActAsDropDown** property is True.  
 When the user selects a row in a DropDown grid, the value in the column, identified by the **BoundColumnIndex** property, is the one passed to a column with **EditType = jgexEditCombo** into another **GridEX** control.

### Data Type

 Variant

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ReplaceColumnIndex Property](#replacecolumnindex-property-gridex-control), [ActAsDropDown Property](#actasdropdown-property-gridex-control), [EditType Property](Objects/JSColumn-Object.md#edittype-property-jscolumn-object), [ReplaceValues Property](Objects/JSColumn-Object.md#replacevalues-property-jscolumn-object)  
**Example:** [ActAsDropDown Example](../Examples.md#actasdropdown-example)

## CalendarNoneText Property (GridEX Control)

Returns or sets the text displayed in the calendar drop-down ‘None’ button.

### Syntax

 *object*.**CalendarNoneText** [ = *value* ]  
 The **CalendarNoneText** property syntax has these parts:

| Part | Description |
| --- | --- |
| *Object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A string that represents the text that will be displayed in the “None” button of a drop-down calendar. |

### Remarks

 The **CalendarNoneText** is displayed in the “None” button of a drop-down calendar. You can change this property to any string for customization.  
 Setting this property to an empty string will hide the “None” button in the calendar.  
 This property is used only when the **GridEX** control has one or more columns with its **EditType** property set to **jgexEditCalendarDropDown**.

### Data Type

 String

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [CalendarTodayText Property](#calendartodaytext-property-gridex-control)

## CalendarTodayText Property (GridEX Control)

Returns or sets the text displayed in the calendar drop-down ‘Today’ button.

### Syntax

 *object*.**CalendarTodayText** [ = *value* ]  
 The **CalendarTodayText** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A string that represents the text that will be displayed in the “Today” button of a drop-down calendar. |

### Remarks

 The **CalendarTodayText** is displayed in the “Today” button of a drop-down calendar. You can change this property to any string for customization.  
 Setting this property to an empty string will hide the today button in the calendar.  
 This property is used only when the **GridEX** control has one or more columns with its **EditType** property set to **jgexEditCalendarDropDown**.

### Data Type

 String

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [CalendarNoneText Property](#calendarnonetext-property-gridex-control)

## CardBorders Property (GridEX Control)

Returns or sets whether the **GridEX** will display borders around the individual cards while the control is in card view.

### Syntax

 *object*.**CardBorders** [ = *value*]  
 The **CardBorders** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether the **GridEX** control will show borders in cards, as described in Settings. |

### Settings

 The **CardBorders** property settings are:

| Setting | Description |
| --- | --- |
| **True** | The cards in **GridEX** control will appear with borders. |
| **False** | The cards in **GridEX** control will appear with no borders. |

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BorderStyle Property](#borderstyle-property-gridex-control)  
**Example:** [View Example](../Examples.md#view-example)

## CardCaptionPrefix Property (GridEX Control)

Returns or sets the text that will appear prefixed to the caption in every card.

### Syntax

 *object*.**CardCaptionPrefix** [ = *value*]  
 The **CardCaptionPrefix** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A string expression that specifies the text that will appear prefixed to the caption in every card. |

### Data Type

 String

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Caption Property](Objects/JSColumn-Object.md#caption-property-jscolumn-object), [CardIcon Property](Objects/JSColumn-Object.md#cardicon-property-jscolumn-object), [CardCaption Property](Objects/JSColumn-Object.md#cardcaption-property-jscolumn-object)  
**Example:** [View Example](../Examples.md#view-example)

## CardSpacing Property (GridEX Control)

Returns or sets the horizontal and vertical spacing between cards.

### Syntax

 *object*.**CardSpacing** [ = *value* ]  
 The **CardSpacing** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Long expression that specifies the space, in twips, between cards while the GridEX is in card view. |

**Remarks**:  
 The **CardSpacing** property is given in twips.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [CardWidth Property](#cardwidth-property-gridex-control), [View Property](#view-property-gridex-control)  
**Example:** [View Example](../Examples.md#view-example)

## CardWidth Property (GridEX Control)

Returns or sets the width of all cards when the **GridEX** is in card view.

### Syntax

 *object*.**CardWidth** [ = *value*]  
 The **CardWidth** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Long expression that specifies the width, in twips, of all cards while the **GridEX** control is in card view. |

**Remarks**:  
 The **CardWidth** property is given in twips.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [CardSpacing Property](#cardspacing-property-gridex-control), [View Property](#view-property-gridex-control), [CardResize Event](Events.md#cardresize-event-gridex-control), [AllowCardSizing Property](#allowcardsizing-property-gridex-control)  
**Example:** [View Example](../Examples.md#view-example)

## Col Property (GridEX Control)

Returns or sets the active column in a **GridEX** control. Not available at design-time.

### Syntax

 *object*.**Col** [ = *value* ]  
 The **Col** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | An integer that represents the position of the column containing the active cell. |

### Remarks

 Use this property to specify a cell in a **GridEX** control or to find out which column contains the active cell. In the **GridEX** control, columns and rows are numbered starting from 1, beginning at the left for columns.  
 This property represents the column position and not its **Index** property. The index and the position of a column are not always the same because the position of the columns could change in response to a user action. If you want to find out the index of the current column, you may do something like this:

```vb
Dim Col as JSColumn
Dim ColPosition as Integer
Dim ColIndex as Integer

'Retrieve the actual column position

ColPosition = GridEX1.Col
Set Col = GridEX1.Columns.ItemByPosition(ColPosition)
ColIndex = Col.Index
Debug.Print ColIndex

'A value of 0 in the Col property means that the entire row is selected.
```

If you set the **Col** property with a value greater than the visible columns count, the **GridEX** control will select the last visible column.  
 If you set the **Col** property to a non-selectable column the **GridEX** control will select the next available column that can be selected, instead of the one you specified.  
 This property can be used in both, table and card, views. The only difference between views is that, in table view, the **Col** property represents a column in the current row while in card view represents a row field in the current card.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ItemByPosition Method](Objects/JSColumns-Collection.md#itembyposition-method-jscolumns-collection), [ColPosition Property](Objects/JSColumn-Object.md#colposition-property-jscolumn-object), [Columns Property](#columns-property-gridex-control), [JSColumn Object](Objects/JSColumn-Object.md#jscolumn-object), [Row Property](#row-property-gridex-control), [RowColChange Event](Events.md#rowcolchange-event-gridex-control)  
**Example:** [Col Example](../Examples.md#col-example)

## ColumnAutoResize Property (GridEX Control)

Returns or sets a value indicating whether the **GridEX** will automatically size its visible columns to fit on the control's client width.

### Syntax

 *object*.**ColumnAutoResize** [ = *value*]  
 The **ColumnAutoResize** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether the GridEX control will size its columns to fit on screen, as described in Settings. |

### Settings

 The **ColumnAutoResize** property settings are:

| Setting | Description |
| --- | --- |
| **False** | Columns won't be sized when the internal width of the control changes. |
| **True** | Visible columns in **GridEX** control will be sized to fit **GridEX'**s internal width. |

### Remarks

 This property is only used when the **GridEX** control is in table view.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ColResize Event](Events.md#colresize-event-gridex-control), [Autosize Method](Objects/JSColumn-Object.md#autosize-method-jscolumn-object), [AllowSizing Property](Objects/JSColumn-Object.md#allowsizing-property-jscolumn-object)

## ColumnHeaderFont Property (GridEX Control)

Returns or sets a value indicating the font used in column headers in a **GridEX** control.

### Syntax

 *object*.**ColumnHeaderFont** [ = *font*]  
 The **ColumnHeaderFont** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *font* | An object expression that evaluates to a **Font** object. |

### Remarks

 The **ColumnHeaderFont** is used in column headers as well as in the “group by” box information text, if the control is in table view. In card view, the **ColumnHeaderFont** property represents the font used in the caption bar of all cards.  
 Changing the **ColumnHeaderFont** property may resize the headers to accommodate the new font.

### Data Type

 **StdFont**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ColumnHeaderHeight Property](#columnheaderheight-property-gridex-control), [Font Property](#font-property-gridex-control), [ColumnHeaders Property](#columnheaders-property-gridex-control), [HeaderStyle Property](Objects/JSColumn-Object.md#headerstyle-property-jscolumn-object)

## ColumnHeaderHeight Property (GridEX Control)

Returns or sets the height of the column header row in a **GridEX** control.

### Syntax

 *object*.**ColumnHeaderHeight** [ = *value*]  
 The **ColumnHeaderHeight** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Long expression that specifies the height of the column header row in a **GridEX** control. |

### Remarks

 The **ColumnHeaderHeight** property is calculated every time the **ColumnHeaderFont** changes to accommodate the new **Font**. However, if you want another height than the default height you can set this property to any valid dimension.  
 The **ColumnHeaderHeight** property is given in twips.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Font Property](#font-property-gridex-control), [ColumnHeaders Property](#columnheaders-property-gridex-control), [ColumnHeaderFont Property](#columnheaderfont-property-gridex-control)

## ColumnHeaders Property (GridEX Control)

Returns or sets a value indicating whether the column headers are displayed in a **GridEX** control.

### Syntax

 *object*.**ColumnHeaders** [ = *value*]  
 The **ColumnHeaders** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether column headers are displayed, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | **GridEX** control's column headers are displayed |
| **False** | **GridEX** control's column headers are not displayed |

**Remarks**:  
 This property has no effect when the **GridEX** control is in card view.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ColumnHeaderHeight Property](#columnheaderheight-property-gridex-control), [RowHeaders Property](#rowheaders-property-gridex-control), [ColumnHeaderFont Property](#columnheaderfont-property-gridex-control)

## Columns Property (GridEX Control)

Returns the **JSColumns** collection of a **GridEX** control.

### Syntax

 *object*.**Columns**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The Columns property returns a **JSColumns** object.  
 The **JSColumns** object is a collection of **JSColumn** objects in a **GridEX** control.  
 You can get a specific column through the **Item** property or through the **ItemByPosition** method.  
 You can manipulate most of the **GridEX** control's attributes by changing the properties of **JSColumn** objects.

### Data Type

 **JSColumns**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [JSColumn Object](Objects/JSColumn-Object.md#jscolumn-object), [JSColumns Collection](Objects/JSColumns-Collection.md#jscolumns-collection)  
**Example:** [Columns Example](../Examples.md#columns-example)

## Connect Property (GridEX Control)

Returns or sets a value that indicates to a **GridEX** control the connection parameter used to open the source database when the **DataMode** Property is set to **jgexDAO**.

### Syntax

 *object*.**Connect** [ = *value*]  
 The **Connect** property syntax has these parts.

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A String that specifies the **Connect** property of the database for use when its **DataMode** property is set to **jgexDAO**. |

### Remarks

 This property has effect only when the **DataMode** property setting is **jgexDAO**.  
 To open a password-protected Access database, set the **Connect** property to the following string:  
 ;pwd=MyPassword  
 Where “MyPassword” is the actual database password.

### Data Type

 String

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [DatabaseName Property](#databasename-property-gridex-control), [DataMode Property](#datamode-property-gridex-control)

## ContinuousScroll Property (GridEX Control)

Returns or sets a value that determines whether a **GridEX** control should scroll its contents while the user drags the scroll thumb along the vertical scroll bar.

### Syntax

 *object*.**ContinuousScroll** [ = *boolean* ]  
 The **ContinuousScroll** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *boolean* | A Boolean expression that determines if scrolling is effected while the user drags the scroll thumb. This property is **False** by default. |

### Remarks

 This property should normally be set to **False** to avoid excessive flickering. Set it to **True** only if you require that rows are viewed as they are being scrolled.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ScrollToolTips Property](#scrolltooltips-property-gridex-control), [ScrollToolTip Column Property](#scrolltooltip-column-property-gridex-control)

## CursorLocation Property (GridEX Control)

Returns or sets the location of the cursor engine when **DataMode** property setting is **jgexADO**.

### Syntax

 *object*.**CursorLocation** [ = *value* ]  
 The **CursorLocation** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines if server-side or client-side cursors are used, as described in Settings. |

### Settings

 The settings for value are:

| Constant | Value | Description |
| --- | --- | --- |
| **jgexUseServer** | 2 | Default. Uses data-provider– or driver-supplied cursors. |
| **jgexUseClient** | 1 | Uses client-side cursors supplied by a local cursor library. |

**Remarks**:  
 This property allows you to choose between using a client-side cursor library or one that is located on the server.  
 This property setting affects only when a recordset is opened in a **GridEX** control. Changing the **CursorLocation** property has no effect on the current **Recordset** object until a **Rebind** call is made.

### Data Type

 **jgexCursorLocationConstants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ADORecordset Property](#adorecordset-property-gridex-control), [DataMode Property](#datamode-property-gridex-control)  
**Example:** [RebindSample Example](../Examples.md#rebindsample-example)

## DatabaseName Property (GridEX Control)

Returns or sets the name and location of the source of data for a **GridEX** control.  
 Syntax  
 *object*.**DatabaseName** [ = *name* ]  
 The **DatabaseName** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *name* | A string expression that indicates the location of the database file(s), in DAO mode, or the Data Source Name in ADO mode. |

### Remarks

 If you change the **DatabaseName** property after the control's **Database** object is open, you must use the **Rebind** method to re-open the new database.  
 When a **GridEX** control’s **DataMode** property is set to **jgexDAO**, the **DatabaseName** represents the name parameter of the **OpenDatabase** method. In ADO mode, it represents the connectionstring of an **ADOR.Recordset**

### Data Type

 String

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [DataMode Property](#datamode-property-gridex-control), [Connect Property](#connect-property-gridex-control), [Rebind Method](Methods.md#rebind-method-gridex-control)  
**Example:** [RebindSample Example](../Examples.md#rebindsample-example)

## DataChanged Property (GridEX Control)

Returns or sets a value indicating that the data in a **GridEX** control has been changed by some process other than by retrieving data from the current record. Not available at design time.

### Syntax

 *object*.**DataChanged** [ = *value*]  
 The DataChanged property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that indicates whether data has changed, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | The data currently in the control is not the same as in the current record. |
| **False** | (Default) The data currently in the control, if any, is the same as the data in the current record. |

### Remarks

 When the user has made changes to current record in a **GridEX** control, the **DataChanged** property is set to **True** and before the user moves to another record the control will try to write those changes in the database.  
 If you do not wish to save changes to the database, you can set the **DataChanged** property to False.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Value Property](#value-property-gridex-control), [DataChanged Property](Objects/JSColumn-Object.md#datachanged-property-jscolumn-object), [Update Method](Methods.md#update-method-gridex-control)  
**Example:** [DataChanged Example](../Examples.md#datachanged-example)

## DataMode Property (GridEX Control)

Returns or sets a value representing the **GridEX** data retrieval mode. Read-only at run time.

### Syntax

 *object*.**DataMode** [ = *value* ]  
 The **DataMode** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines the data retrieval mode in a **GridEX** control, as described in Settings. |

### Settings

 The settings for value are:

| Constant | Value | Description |
| --- | --- | --- |
| **jgexDAO** | 0 | The **GridEX** uses DAO 3.6 to create the Recordset. |
| **jgexADO** | 1 | The **GridEX** uses ADO to create the Recordset. |
| **jgexUnbound** | 99 | The **GridEX** uses the unbound events to retrieve and update displayed data. |

**Remarks**:  
 The **DataMode** property controls how a **GridEX** control connects to and manipulates its source data.  
 In DAO and ADO mode, the data is retrieved and updated automatically using the **Recordset** object created by the control.  
 In unbound mode, you are responsible for maintaining data and supplying the **GridEX** control with the appropriate data when requested through the unbound events.  
 Normally, the unbound mode of the **GridEX** control is used when displaying data that is not stored in a database accessible by DAO or ADO, or to display and manipulate custom data, such as arrays. You can use the unbound mode for whatever type of data you have available. For example, you can use the unbound mode of the **GridEX** control to display data from a proprietary database format or use it to manage data that you keep track of in a text file.

**Note** In unbound mode there is a performance penalty in sorting and grouping compared with the other modes.

### Data Type

 **jgexDataMode Constants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ADORecordset Property](#adorecordset-property-gridex-control), [Recordset Property](#recordset-property-gridex-control), [DatabaseName Property](#databasename-property-gridex-control), [UnboundAddNew Event](Events.md#unboundaddnew-event-gridex-control), [UnboundReadData Event](Events.md#unboundreaddata-event-gridex-control), [UnboundDelete Event](Events.md#unbounddelete-event-gridex-control), [UnboundUpdate Event](Events.md#unboundupdate-event-gridex-control), [RecordSource Property](#recordsource-property-gridex-control)  
**Example:** [RebindSample Example](../Examples.md#rebindsample-example)

## DefaultColumnWidth Property (GridEX Control)

Returns or sets a value indicating the default column width used for new columns added to a **GridEX** control.

### Syntax

 *object*.**DefaultColumnWidth** [ = *value*]  
 The **DefaultColumnWidth** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Long expression that identifies the default width, in twips, for new columns. |

### Remarks

 The **DefaultColumnWidth** property is given in twips.  
 If you set this property, the control will use it to set the initial width of all columns added to the **JSColumns** collection as a result of a binding operation or for manually added columns.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Width Property](Objects/JSColumn-Object.md#width-property-jscolumn-object)

## DefaultGroupMode Property (GridEX Control)

Returns or sets a value that determines how the control groups records.

### Syntax

 *object*.**DefaultGroupMode** [= *value*]  
 The **DefaultGroupMode** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines the default group mode, as described in Settings. |

### Settings

 The **DefaultGroupMode** property settings are:

| Constant | Value | Description |
| --- | --- | --- |
| **jgexDGMExpanded** | 0 | Default. The control groups records expanding all the groups. |
| **jgexDGMCollapsed** | 1 | The control groups records collapsing all the groups. |

### Remarks

 This property is used to specify how groups will be shown the first time groups are calculated by the control in response to a change in the **Groups** collection.  
 To collapse or expand all groups once the groups are calculated use the **CollapseAll** or **ExpandAll** methods respectively.

### Data Type

 **jgexDefaultGroupModeConstants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ExpandAll Method](Methods.md#expandall-method-gridex-control), [CollapseAll Method](Methods.md#collapseall-method-gridex-control), [RefreshGroups Method](Methods.md#refreshgroups-method-gridex-control)

## DetectRowDrag Property (GridEX Control)

Returns or sets a value indicating whether a **GridEX** control will detect and signal when a user is about to drag the selected rows.  
 Syntax  
 *object*.**DetectRowDrag** [ = *value* ]  
 The **DetectRowDrag** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether a **RowDrag** event be triggered when a user attempts to drag the selected row(s), as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | A **RowDrag** event will be triggered when the control detects the user is about to drag a row. |
| **False** | Default. No row drag detect is made and the **RowDrag** event is not raised. |

### Remarks

 This property is useful if you are implementing OLE Drag in a **GridEX** control.  
 In order to make a more efficient OLE Drag implementation, the **OLEDrag** method must be called when a drag is detected (in the **RowDrag** event) instead of calling it in the **MouseDown** event.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RowDrag Event](Events.md#rowdrag-event-gridex-control), [OLEDrag Method](Methods.md#oledrag-method-gridex-control)  
**Example:** [RowDrag Example](../Examples.md#rowdrag-example)

## EditMode Property (GridEX Control)

Returns or sets whether the cell editor is active. Not available at design time.

### Syntax

 *object*.**EditMode** [ = *value* ]  
 The **EditMode** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines if the cell editor is active or inactive, as described in Settings. |

### Settings

 The settings for value are:

| Constant | Value | Description |
| --- | --- | --- |
| **jgexEditModeOff** | 0 | The cell editor is inactive. No editing is in progress. |
| **jgexEditModeOn** | 1 | The cell editor is active. The current cell is being edited by the user. |

### Remarks

 If the current cell is not already in edit mode, setting **EditMode** to **jgexEditModeOn** will show the cell editor.  
 If the grid is already in edit mode, setting **EditMode** to **jgexEditModeOff** will exit edit mode.  
 Note that this property does not determine if the user has changed data in the cell; it only returns whether or not the cell editor is active.  
 If you need to determine if the data has changed, use the **DataChanged** property.

### Data Type

 **jgexEditModeConstants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [hWndEdit Property](#hwndedit-property-gridex-control), [DataChanged Property](#datachanged-property-gridex-control), [BeforeColEdit Event](Events.md#beforecoledit-event-gridex-control), [AfterColEdit Event](Events.md#aftercoledit-event-gridex-control)  
**Example:** [EditMode Example](../Examples.md#editmode-example)

## EmptyRows Property (GridEX Control)

Returns or sets a value indicating whether the **GridEX** control will display empty rows below the last row.

### Syntax

 *object*.**EmptyRows** [ = *value* ]  
 The **EmptyRows** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether empty rows are displayed below the last row, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | The **GridEX** control will show empty rows below the last row. |
| **False** | Default. The **GridEX** control will paint the empty space below the last row using the **BackColorBkg** property. |

### Remarks

 This property is only used when the **GridEX** control is in table view.  
 Empty rows are only decorative and can not be selected by the user.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [GridLinesColor Property](#gridlinescolor-property-gridex-control), [GridLines Property](#gridlines-property-gridex-control), [GridLineStyle Property](#gridlinestyle-property-gridex-control)  
**Example:** [View Example](../Examples.md#view-example)

## Enabled Property (GridEX Control)

Returns or sets a value that determines whether a **GridEX** control can respond to user-generated events.

### Syntax

 *object*.**Enabled** [ = *value* ]  
 The **Enabled** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that specifies whether a **GridEX** control can respond to user-generated events, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | (Default) The control will respond to user generated events. |
| **False** | Prevents control from responding to events |

### Remarks

 The Enabled property allows you to enabled or disable a **GridEX** control at run time. For example, you can disable a **GridEX** control when no data is loaded on it.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AllowEdit Property](#allowedit-property-gridex-control)

## ErrorText Property (GridEX Control)

Returns or sets the error message String from the underlying data source. Not available at design time.

### Syntax

 *object*.**ErrorText** [ = *value*]  
 The **ErrorText** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A string that represents the text that will be displayed if an error occurs. |

### Remarks

 When a database error occurs as a result of user interaction with the control, such as when the user enters text into a numeric field and then attempts to update the current record, the control’s **Error** event is fired. However, the error code passed to the event handler in the *errnumber* parameter may not identify the specific error that occurred, or may even differ across operating environments. For these reasons, the **ErrorText** property is provided so that your application can parse the actual error message and determine the nature of the error.  
 If you want to customize the error description displayed by the control you may set the **ErrorText** property to any string or cancel the default response in the **Error** event and display your custom message.

**Note** The **ErrorText** property is only valid within a **GridEX** control's **Error** event handler.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Error Event](Events.md#error-event-gridex-control)  
**Example:** [Error Example](../Examples.md#error-example)

## Exclusive Property (GridEX Control)

Returns or sets a value that indicates whether the underlying database for a **GridEX** control is opened for single- or multi-user access when DAO mode is specified.

### Syntax

 *object*.**Exclusive** [ = *value*]  
 The **Exclusive** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines user access, as described in Settings. |

### Settings

 The settings for value are:

| Part | Description |
| --- | --- |
| **True** | The database is open for single-user access. No one else can open the database until it is closed. |
| **False** | (Default) The database is open for multi-user access. Other users can open the database and have access to the data while it is open. |

### Remarks

 This property only affects the way that **GridEX** open a database in DAO mode.  
 The value of this property, along with the **DatabaseName**, **ReadOnly**, and **Connect** properties, is used when opening a database.  
 This property corresponds to the *Exclusive* argument in the **OpenDatabase** method.  
 The **Exclusive** property is used only when opening the **Database**. If you change the value of this property at run time, you must use the **Rebind** method for the change to take effect. If someone else already has the database open, you will not be able to open it for exclusive use and an **Error** event is fired.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Connect Property](#connect-property-gridex-control), [DatabaseName Property](#databasename-property-gridex-control), [ReadOnly Property](#readonly-property-gridex-control), [Rebind Method](Methods.md#rebind-method-gridex-control)  
**Example:** [RebindSample Example](../Examples.md#rebindsample-example)

## FirstItem Property (GridEX Control)

Returns or sets the row position of the first visible row or card in a **GridEX** control. Not available at design time.

### Syntax

 *object*.**FirstItem** [ = *value*]  
 The **FirstItem** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A long expression corresponding to the record position of the first visible row in the **GridEX** control. |

### Remarks

 For a **GridEX** control, setting the **FirstItem** property causes the control to scroll, so that the specified row or card becomes the topmost row (if in table view) or the upper-left card (if in card view).  
 When the **GridEX** control is in card view, setting the **FirstItem** property ensures visibility of a card, but depending on the extent of the visible area of the grid and the actual positional index of the card, the specified card may not be positioned exactly in the one in upper-left corner.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [FirstItemChange Event](Events.md#firstitemchange-event-gridex-control)

## FmtConditions Property (GridEX Control)

Returns the **JSFmtConditions** object in a **GridEX** control.

### Syntax

 *object*.**FmtConditions**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The **FmtConditions** property returns a **JSFmtConditions** collection. You can get a specific **JSFmtCondition** through the **Item** property in the **JSFmtConditions** collection, you can also get a special **JSFmtCondition** object for expanding it through the group rows.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [JSFmtCondition Object](Objects/JSFmtCondition-Object.md#jsfmtcondition-object), [JSFmtConditions Collection](Objects/JSFmtConditions-Collection.md#jsfmtconditions-collection)  
**Example:** [FmtConditions Example](../Examples.md#fmtconditions-example)

## Font Property (GridEX Control)

Returns or sets a **Font** object used in a **GridEX** control.

### Syntax

 *object*.**Font**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 Setting the **Font** property of a **GridEx** control affects the font used to render the cells and/or cards.  
 Use the **Font** property of an object to obtain a specific **Font** object whose properties you want to use.  
 For example, the following code changes the **Bold** property setting of a **Font** object identified by the **Font** property of a **GridEX** control:

```vb
GridEX1.Font.Bold = True
```

### Data Type

 **StdFont**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ColumnHeaderFont Property](#columnheaderfont-property-gridex-control), [JSFormatStyle Object](Objects/JSFormatStyle-Object.md#jsformatstyle-object), [ForeColor Property](#forecolor-property-gridex-control)

## ForeColor Property (GridEX Control)

Returns or sets the default foreground color, for the text in **GridEX**'s cells.

### Syntax

 *object*.**ForeColor** [ =*color* ]  
 The **ForeColor** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A numeric expression that specifies the color. |

### Remarks

 The **ForeColor** property refers to the default text color of all the cells in table and card view.

### Data Type

 OLE_COLOR

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BackColorBkg Property](#backcolorbkg-property-gridex-control), [BackColorGBBox Property](#backcolorgbbox-property-gridex-control), [BackColorHeader Property](#backcolorheader-property-gridex-control), [BackColorInfoText Property](#backcolorinfotext-property-gridex-control), [BackColor Property](#backcolor-property-gridex-control), [BackColor Property](Objects/JSFormatStyle-Object.md#backcolor-property-jsformatstyle-object), [BackColorRowGroup Property](#backcolorrowgroup-property-gridex-control), [ForeColorHeader Property](#forecolorheader-property-gridex-control), [ForeColorInfoText Property](#forecolorinfotext-property-gridex-control), [ForeColor Property](Objects/JSFormatStyle-Object.md#forecolor-property-jsformatstyle-object), [ForeColorRowGroup Property](#forecolorrowgroup-property-gridex-control), [MaskColor Property](#maskcolor-property-gridex-control), [RowColorEven Property](#rowcoloreven-property-gridex-control), [RowColorOdd Property](#rowcolorodd-property-gridex-control), [GridLinesColor Property](#gridlinescolor-property-gridex-control)

## ForeColorHeader Property (GridEX Control)

Returns or sets the foreground color, for the text in **GridEX**'s Column Headers.

### Syntax

 *object*.**ForeColorHeader** [ =*color* ]  
 The **ForeColorHeader** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A numeric expression that specifies the color. |

### Remarks

 The **ForeColorHeader** property in table view refers to the text color of the column and row headers while in card view it refers to the text color of the card caption.

### Data Type

 OLE_COLOR

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BackColorBkg Property](#backcolorbkg-property-gridex-control), [BackColorGBBox Property](#backcolorgbbox-property-gridex-control), [BackColorHeader Property](#backcolorheader-property-gridex-control), [BackColorInfoText Property](#backcolorinfotext-property-gridex-control), [BackColor Property](#backcolor-property-gridex-control), [BackColor Property](Objects/JSFormatStyle-Object.md#backcolor-property-jsformatstyle-object), [BackColorRowGroup Property](#backcolorrowgroup-property-gridex-control), [ForeColorInfoText Property](#forecolorinfotext-property-gridex-control), [ForeColor Property](Objects/JSFormatStyle-Object.md#forecolor-property-jsformatstyle-object), [ForeColor Property](#forecolor-property-gridex-control), [ForeColorRowGroup Property](#forecolorrowgroup-property-gridex-control), [MaskColor Property](#maskcolor-property-gridex-control), [RowColorEven Property](#rowcoloreven-property-gridex-control), [RowColorOdd Property](#rowcolorodd-property-gridex-control), [GridLinesColor Property](#gridlinescolor-property-gridex-control)

## ForeColorInfoText Property (GridEX Control)

Returns or sets the foreground color, for the text in **GridEX**'s **GroupByBoxInfoText**.

### Syntax

 *object*.**ForeColorInfoText** [ =*color*]  
 The **ForeColorInfoText** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A numeric expression that specifies the color. |

### Remarks

 In Table view, the **ForeColorInfoText** property refers to the text color of the information text displayed in the “group by” box when there are no active groups. The **ForeColorInfoText** is not used in card view.

### Data Type

 OLE_COLOR

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BackColorBkg Property](#backcolorbkg-property-gridex-control), [BackColorGBBox Property](#backcolorgbbox-property-gridex-control), [BackColorHeader Property](#backcolorheader-property-gridex-control), [BackColorInfoText Property](#backcolorinfotext-property-gridex-control), [BackColor Property](#backcolor-property-gridex-control), [BackColor Property](Objects/JSFormatStyle-Object.md#backcolor-property-jsformatstyle-object), [BackColorRowGroup Property](#backcolorrowgroup-property-gridex-control), [ForeColorHeader Property](#forecolorheader-property-gridex-control), [ForeColor Property](Objects/JSFormatStyle-Object.md#forecolor-property-jsformatstyle-object), [ForeColor Property](#forecolor-property-gridex-control), [ForeColorRowGroup Property](#forecolorrowgroup-property-gridex-control), [MaskColor Property](#maskcolor-property-gridex-control), [RowColorEven Property](#rowcoloreven-property-gridex-control), [RowColorOdd Property](#rowcolorodd-property-gridex-control), [GridLinesColor Property](#gridlinescolor-property-gridex-control)

## ForeColorRowGroup Property (GridEX Control)

Returns or sets the default foreground color, for the text in **GridEX**'s group rows.

### Syntax

 *object*.**ForeColorRowGroup** [=*color* ]  
 The **ForeColorRowGroup** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A numeric expression that specifies the color. |

### Remarks

 The **ForeColorRowGroup** property refers to the text color of the group rows in table view. The **ForeColorRowGroup** is not used in card view.

### Data Type

 OLE_COLOR

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BackColorBkg Property](#backcolorbkg-property-gridex-control), [BackColorGBBox Property](#backcolorgbbox-property-gridex-control), [BackColorHeader Property](#backcolorheader-property-gridex-control), [BackColorInfoText Property](#backcolorinfotext-property-gridex-control), [BackColor Property](#backcolor-property-gridex-control), [BackColor Property](Objects/JSFormatStyle-Object.md#backcolor-property-jsformatstyle-object), [BackColorRowGroup Property](#backcolorrowgroup-property-gridex-control), [ForeColorHeader Property](#forecolorheader-property-gridex-control), [ForeColorInfoText Property](#forecolorinfotext-property-gridex-control), [ForeColor Property](Objects/JSFormatStyle-Object.md#forecolor-property-jsformatstyle-object), [ForeColor Property](#forecolor-property-gridex-control), [MaskColor Property](#maskcolor-property-gridex-control), [RowColorEven Property](#rowcoloreven-property-gridex-control), [RowColorOdd Property](#rowcolorodd-property-gridex-control), [GridLinesColor Property](#gridlinescolor-property-gridex-control)

## FormatStyles Property (GridEX Control)

Returns the **JSFormatStyles** collection of a **GridEX** control.

### Syntax

 *object*.**FormatStyles**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The **FormatStyles** property returns a **JSFormatStyles** object. The **JSFormatStyles** object is a collection of **JSFormatStyles** objects in a **GridEX** control. You can get a specific **JSFormatStyle** object through the **Item** property.  
 **JSFormatStyle** objects returned by the **FormatStyle** property in a **JSFmtCondition** object are not present in this collection.

### Data Type

 **JSFormatStyles**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [JSFormatStyle Object](Objects/JSFormatStyle-Object.md#jsformatstyle-object), [JSFormatStyles Collection](Objects/JSFormatStyles-Collection.md#jsformatstyles-collection)  
**Example:** [FormatStyles Example](../Examples.md#formatstyles-example)

## FrozenColumns Property (GridEX Control)

Returns or sets the number of fixed columns at the left of the control.

### Syntax

 *object*.**FrozenColumns** [ = *value* ]  
 The **FrozenColumns** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | An Integer expression specifying the number of columns to be fixed. |

### Remarks

 You can use this property to specify the number of non-scrollable columns at the left of the control.  
 If you want to "freeze" a column that is not the leftmost column of the grid, change its **ColPosition** to 1 first, then you can set the **FrozenColumns** property to 1 in order to freeze it.  
 To "unfreeze" previously frozen columns set the **FrozenColumns** property to 0.

**Note** This property is not used when the **GridEX** control is in card view.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ColPosition Property](Objects/JSColumn-Object.md#colposition-property-jscolumn-object), [View Property](#view-property-gridex-control)  
**Example:** [FrozenColumns Example](../Examples.md#frozencolumns-example)

## FullyLoaded Property (GridEX Control)

Returns whether the control has loaded the entire **Recordset**. Read only. Not available at design time.

### Syntax

 *object*.**FullyLoaded**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Settings

 The settings for **FullyLoaded** property are:

| Setting | Description |
| --- | --- |
| False | The **GridEX** control has not reached the end of the **Recordset**. |
| True | The **GridEX** control has reached the end of the **Recordset**. |

### Remarks

 When the **GridEX** control is in unbound mode the **FullyLoaded** property has no meaning and always returns **True**.  
 The **FullyLoaded** property is read-only. However, if you want to set this property to **True**, you can call the **LoadEntireRecordset** method.  
 An example of this is as follows:

```vb
If Not GridEX1.FullyLoaded then
GridEX1.LoadEntireRecordset
End if
Debug.Print GridEX1.FullyLoaded 'Prints "True"
```

If you are adding records to the **GridEX** control’s **Recordset** clone in code when the **FullyLoaded** property is **True**, the control will not show those records unless you use the **SearchNewRecords** method.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [LoadEntireRecordset Method](Methods.md#loadentirerecordset-method-gridex-control), [SearchNewRecords Method](Methods.md#searchnewrecords-method-gridex-control)  
**Example:** [RowSelected Example](../Examples.md#rowselected-example)

## GridImages Property (GridEX Control)

Returns the **JSGridImages** collection of a **GridEX** control.

### Syntax

 *object*.**GridImages**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The **GridImages** property returns a **JSGridImages** collection. You can access a specific **JSGridImage** object through the **Item** property of this collection.

### Data Type

 **JSGridImages**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [JSGridImage Object](Objects/JSGridImage-Object.md#jsgridimage-object), [JSGridImages Collection](Objects/JSGridImages-Collection.md#jsgridimages-collection)  
**Example:** [GridImages Example](../Examples.md#gridimages-example)

## GridLineStyle Property (GridEX Control)

Returns or sets a value that determines the grid line style.

### Syntax

 *object*.**GridLineStyle** [= *value*]  
 The **GridLineStyle** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines how grid lines are painted, as described in Settings. |

### Settings

 The **GridLineStyle** property settings are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexGLSSolid** | 0 | Default. Solid line. |
|  **jgexGLSDashes** | 1 | Dashed line. |
|  **jgexGLSSmallDots** | 2 | Small dotted line. |

### Remarks

 This property is used in conjunction with **GridLines** property.  
 If **GridLines** property is equal to **jgexGLNone** this property has no effect.

### Data Type

 **jgexGridLineStyleConstants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [EmptyRows Property](#emptyrows-property-gridex-control), [GridLinesColor Property](#gridlinescolor-property-gridex-control), [GridLines Property](#gridlines-property-gridex-control)  
**Example:** [View Example](../Examples.md#view-example)

## GridLines Property (GridEX Control)

Returns or sets a value that determines how the **GridEX** control will draw lines between cells.

### Syntax

 *object*.**GridLines** [ = *value* ]  
 The **GridLines** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant specifying how the **GridEX** control should draw lines, as described in Settings. |

### Settings

 The settings for value are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexGLBoth** | -1 | Default. Vertical and horizontal lines are drawn in-between cells. |
|  **jgexGLNone** | 0 | No Lines. No lines are drawn in-between cells. |
|  **jgexGLVertical** | 1 | Only vertical lines are drawn in-between cells. |
|  **jgexGLHorizontal** | 2 | Only horizontal lines are drawn in-between cells. |

### Remarks

 When the **GridLines** property setting is other than **jgexGLNone**, the color of the lines is determined by the **GridLinesColor** property and the style by the **GridLineStyle** property.  
 This property has no effect when a **GridEX** control’s **View** property is set to **jgexCard**.

### Data Type

 **jgexGridLinesConstants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [EmptyRows Property](#emptyrows-property-gridex-control), [GridLinesColor Property](#gridlinescolor-property-gridex-control), [GridLineStyle Property](#gridlinestyle-property-gridex-control)  
**Example:** [View Example](../Examples.md#view-example)

## GridLinesColor Property (GridEX Control)

Returns or sets the color used to draw the lines between the **GridEX** control's cells.

### Syntax

 *object*.**GridLinesColor** [ = *color* ]  
 The **GridLinesColor** syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A value or constant that determines the color used to paint the grid lines in a GridEX control. |

### Remarks

 The **GridLinesColor** property is only used when the **GridLines** property setting is **True**.

### Data Type

 OLE_COLOR

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [EmptyRows Property](#emptyrows-property-gridex-control), [GridLines Property](#gridlines-property-gridex-control), [GridLineStyle Property](#gridlinestyle-property-gridex-control)

## GroupByBoxInfoText Property (GridEX Control)

Returns or sets the text displayed in the “group by” box when it is empty; that is, no groups are applied.

### Syntax

 *object*.**GroupByBoxInfoText** [ = *value* ]  
 The **GroupByBoxInfoText** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A string that represents the text that will be displayed in the “group by” box when no groups are applied. |

### Remarks

 The **GroupByBoxInfoText** is displayed in the “group by” box in order to give some information to the users.  
 You can change this property to any string for customization. The default value is “Drag a column header here to group by that column.”  
 This property is used only when the **GridEX** control is in table view.

### Data Type

 String

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [GroupByBoxVisible Property](#groupbyboxvisible-property-gridex-control), [ForeColorInfoText Property](#forecolorinfotext-property-gridex-control), [BackColorGBBox Property](#backcolorgbbox-property-gridex-control), [BackColorInfoText Property](#backcolorinfotext-property-gridex-control)

## GroupByBoxVisible Property (GridEX Control)

Returns or sets a value indicating whether the “group by” box is displayed in a **GridEX** control.

### Syntax

 *object*.**GroupByBoxVisible** [ = *value* ]  
 The **GroupByBoxVisible** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether “group by” box is displayed, as described in Settings. |

### Settings

 The settings for value are:

| Value | Description |
| --- | --- |
| **False** | GridEX control's “group by” box is not displayed |
| **True** | GridEX control's “group by” box is displayed |

### Remarks

 When the **GridEX** control is in card view this property has no effect.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [GroupByBoxInfoText Property](#groupbyboxinfotext-property-gridex-control), [ForeColorInfoText Property](#forecolorinfotext-property-gridex-control), [BackColorGBBox Property](#backcolorgbbox-property-gridex-control), [BackColorInfoText Property](#backcolorinfotext-property-gridex-control), [Groups Property](#groups-property-gridex-control)  
**Example:** [View Example](../Examples.md#view-example)

## GroupFooterStyle Property (GridEX Control)

Returns or sets a value that determines the style of group footers.

### Syntax

 *object*.**GroupFooterStyle** [= *value*]  
 The **GroupFooterStyle** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines the style of group footers, as described in Settings. |

### Settings

 The **GroupFooterStyle** property settings are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexNoGroupFooter** | 0 | Default. No group footers at the bottom of a group. |
|  **jgexCaptionGroupFooter** | 1 | A group footer, identical to the group header, appears at the bottom of each group. |
|  **jgexTotalsGroupFooter** | 2 | A group footer, displaying column totals, appears at the bottom of each group. |

### Remarks

 When **GroupFooterStyle** is equal to **jgexTotalsGroupFooter**, the value that appear in each column is calculated based on the **AggregateFunction**, **TotalRowPrefix** and **TotalRowFormat** properties of the column.

### Data Type

 **jgexGroupFooterStyleConstants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AggregateFunction Property](Objects/JSColumn-Object.md#aggregatefunction-property-jscolumn-object), [TotalRowFormat Property](Objects/JSColumn-Object.md#totalrowformat-property-jscolumn-object), [TotalRowPrefix Property](Objects/JSColumn-Object.md#totalrowprefix-property-jscolumn-object), [RowType Property](Objects/JSRowData-Object.md#rowtype-property-jsrowdata-object), [SubTotal Method](Objects/JSRowData-Object.md#subtotal-method-jsrowdata-object)  
**Example:** [ColumnTotalRowFormat Example](../Examples.md#columntotalrowformat-example)

## Groups Property (GridEX Control)

Returns the **JSGroups** collection in a **GridEX** control.

### Syntax

 *object*.**Groups**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The **Groups** property returns a **JSGroups** collection. You can access specific **JSGroup** objects through the **Item** property in the **JSGroups** collection. You can also add, delete and enumerate group objects using the collection.

### Data Type

 **JSGroups**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [JSGroup Object](Objects/JSGroup-Object.md#jsgroup-object), [JSGroups Collection](Objects/JSGroups-Collection.md#jsgroups-collection)  
**Example:** [Groups Example](../Examples.md#groups-example)

## HeaderStyle Property (GridEX Control)

Returns or sets the display style for the border of column headers in a **GridEX** control.

### Syntax

 *object*.**HeaderStyle** [= *value*]  
 The **HeaderStyle** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines the header style, as described in Settings. |

### Settings

 The **HeaderStyle** property settings are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexHSDouble3D** | 0 | (Default) Headers appear with a double 3D border. |
|  **jgexHSNoBorder** | 1 | Headers appear with no border. |
|  **jgexHSSingleFlat** | 2 | Headers appear with a single flat border. |
|  **jgexHSSingle3D** | 3 | Headers appear with a single 3D border. |

### Data Type

 **jgexHeaderStyleConstants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BorderStyle Property](#borderstyle-property-gridex-control)

## HideSelection Property (GridEX Control)

Returns or sets how selected items will be displayed when the **GridEX** control loses focus

### Syntax

 *object*.**HideSelection** [ = *value* ]  
 The **HideSelection** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines the way selected items will display when **GridEX** loses focus, as described in Settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexHideSelection** | 0 | (Default). The selected row(s) will be displayed as not selected, except for a frame around the row. |
|  **jgexHighLightInactive** | 1 | The selected row(s) will be displayed selected using the button face system color. |
|  **jgexHighLightNormal** | 2 | The selected row(s) will be displayed as normally selected using the same color as when the control has the focus. |

### Remarks

 You can use this property to indicate how selected rows are highlighted while another control or form or a dialog box has the focus.

### Data Type

 **jgexHideSelectionConstants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [SelectionStyle Property](#selectionstyle-property-gridex-control)

## HoldSortSettings Property (GridEX Control)

Returns or sets a value indicating whether the group and sort settings will be held when rebinding the control.

### Syntax

 *object*.**HoldSortSettings** [ = *value*]  
 The **HoldSortSettings** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether whether groups and sort keys are to be preserved when Rebind method is called, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | Default. Groups and sort keys are preserved when Rebind method is called or when setting the Recordset/ADORecordset properties. |
| **False** | Groups and sort keys are cleared when rebinding the control. |

### Remarks

 If **HoldSortSettings** property is True, even when **HoldFields** is not called before a **Rebind**, the control will keep the column layout in order to preserve the groups and sort settings.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [HoldFields Method](Methods.md#holdfields-method-gridex-control), [Rebind Method](Methods.md#rebind-method-gridex-control), [ADORecordset Property](#adorecordset-property-gridex-control), [Recordset Property](#recordset-property-gridex-control), [Refetch Method](Methods.md#refetch-method-gridex-control)  
**Example:** [HoldFields Example](../Examples.md#holdfields-example)

## hWnd Property (GridEX Control)

Returns the handle of a **GridEX** control.

### Syntax

 *object*.**hWnd**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The Microsoft Windows operating environment identifies each control in an application by assigning it a handle, or hWnd.  
 The hWnd property is used with Windows API calls. Many Windows operating environment functions require the hWnd of the active window as an argument.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [hWndEdit Property](#hwndedit-property-gridex-control), [hWnd Property](../GEXPreview-Control/Properties.md#hwnd-property-gexpreview-control)

## hWndEdit Property (GridEX Control)

Returns the handle of the cell editor in a **GridEX** control.

### Syntax

 *object*.**hWndEdit**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 You must not save this property in a variable for later use because the cell editor is created and destroyed as needed.  
 The Microsoft Windows operating environment identifies each control in an application by assigning it a handle, or hWnd.  
 The **hWndEdit** property is used with Windows API calls. Many Windows operating environment functions require the hWnd of the active window as an argument.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [hWnd Property](#hwnd-property-gridex-control)

## ImageHeight Property (GridEX Control)

The **ImageHeight** property returns or sets the height of **JSGridImage** objects in a **GridEX** control.

### Syntax

 *object*.**ImageHeight** [ = *value* ]  
 The **ImageHeight** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Long value that specifies the dimensions, in pixels, of **JSGridImage** objects in a **GridEX** control. |

### Remarks

 Both height and width are measured in pixels.  
 All **JSGridImage** objects in a **JSGridImages** collection have the same height and width.  
 You can change the values of both the **ImageHeight** and **ImageWidth** properties only when the **JSGridImages** collection contains no **JSGridImage** objects.

### Data Type

 Long

**Note** If the dimensions of the images added to the **JSGridImages** collection do not match those specified in the **GridEX** control, they will be stretched accordingly to accommodate the dimensions specified in **ImageHeight** and **ImageWidth** properties.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ImageWidth Property](#imagewidth-property-gridex-control), [GridImages Property](#gridimages-property-gridex-control)  
**Example:** [GridImages Example](../Examples.md#gridimages-example)

## ImageWidth Property (GridEX Control)

The **ImageWidth** property returns or sets the width of **JSGridImage** objects in a **GridEX** control.

### Syntax

 *object*.**ImageWidth** [ = *value* ]  
 The **ImageWidth** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Long value that specifies the dimensions, in pixels, of **JSGridImage** objects in a **GridEX** control. |

### Remarks

 Both height and width are measured in pixels.  
 All **JSGridImage** objects in a **JSGridImages** collection have the same height and width.  
 You can change the values of both the **ImageHeight** and **ImageWidth** properties only when the **JSGridImages** collection contains no **JSGridImage** objects.

### Data Type

 Long

**Note** If the dimensions of the images added to the **JSGridImages** collection do not match those specified in the **GridEX** control, they will be stretched accordingly to accommodate the dimensions specified in **ImageHeight** and **ImageWidth** properties.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ImageHeight Property](#imageheight-property-gridex-control), [GridImages Property](#gridimages-property-gridex-control)  
**Example:** [GridImages Example](../Examples.md#gridimages-example)

## ItemCount Property (GridEX Control)

Returns or sets the number of items (rows) in a **GridEX** control.  
 For DAO and ADO modes, this property is read only. For unbound mode, this property is read-write.

### Syntax

 *object*.**ItemCount** [ = *value* ]  
 The **ItemCount** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A long expression that determines the number of items displayed in a GridEX control. |

### Remarks

 Because the **GridEX** control loads records until is needed, the **ItemCount** property in DAO and ADO modes could be less than the real count, as will be the case if the **GridEX** control has not loaded the entire Recordset. To find out if the entire Recordset has been loaded you must examine the value of the **FullyLoaded** property.  
 If you want to force the load of the **Recordset**, use the **LoadEntireRecordset** method.  
 In unbound mode, the control will fetch data only for the number of rows specified in the **ItemCount** property.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [FullyLoaded Property](#fullyloaded-property-gridex-control), [LoadEntireRecordset Method](Methods.md#loadentirerecordset-method-gridex-control), [RowCount Property](#rowcount-property-gridex-control)  
**Example:** [ItemCount Example](../Examples.md#itemcount-example)

## LeftCol Property (GridEX Control)

Returns or sets the left-most visible column in a **GridEX** control. Not available at design time.

### Syntax

 *object*.**LeftCol** [ = *value* ]  
 The **LeftCol** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | An Integer expression specifying the left-most column. |

### Remarks

 You can use this property in code to scroll the **GridEX** control programmatically.  
 Use the **FirstItem** property to determine the topmost visible row in the **GridEX** control.

**Note** This property is not used when the **GridEX** control is in card view.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [FirstItem Property](#firstitem-property-gridex-control), [FirstItemChange Event](Events.md#firstitemchange-event-gridex-control), [LeftColChange Event](Events.md#leftcolchange-event-gridex-control)

## LockType Property (GridEX Control)

Returns or sets the type of locking for the **Recordset** object of a **GridEX** control.

### Syntax

 *object*.**LockType** [ = *value*]  
 The **LockType** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines the lock type of the Recordset created in a **GridEX** control, as described in Settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexLockReadOnly** | 1 | Prevents user from making changes to the **Recordset**. |
|  **jgexLockPessimistic** | 2 | Uses pessimistic locking to determine how changes are made to the **Recordset** in a multi-user environment. |
|  **jgexLockOptimistic** | 3 | Uses optimistic locking to determine how changes are made to the **Recordset** in a multi-user environment. |
|  **jgexLockBatchOptimistic** | 4 | Enables batch optimistic updates (ADO mode only) |

### Remarks

 This property is used only in DAO and ADO modes to create the **Recordset**.  
 The locking behavior of the **Recordset** is identical for locking values applied to **Recordset** objects created in code.  
 For further information about locking in recordsets, please refer to the corresponding DAO or ADO documentation.

### Data Type

 **jgexLockTypeConstants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ADORecordset Property](#adorecordset-property-gridex-control), [Recordset Property](#recordset-property-gridex-control), [DataMode Property](#datamode-property-gridex-control), [AllowAddNew Property](#allowaddnew-property-gridex-control), [AllowEdit Property](#allowedit-property-gridex-control), [AllowDelete Property](#allowdelete-property-gridex-control)  
**Example:** [RebindSample Example](../Examples.md#rebindsample-example)

## MaskColor Property (GridEX Control)

Returns or sets the color used to create masks for images in a **GridEX** control.

### Syntax

 *object*.**MaskColor** [ = *color*]  
 The **MaskColor** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A value or constant that determines the color used to create masks. |

### Remarks

 Every image in a **JSGridImages** collection has a corresponding mask associated with it. T  
 he mask is a monochrome image derived from the image itself, automatically generated using the **MaskColor** property as the specific color of the mask. This mask is not used directly, but is applied to the original bitmap when the **GridEX** control draws the image.  
 The **MaskColor** property determines which color of an image will be transparent.  
 It is good programming practice to use non-system colors for image's **MaskColor**.  
 When using icons in **JSGridImages**, the **MaskColor** property is not used, since icon images contain their own masks.

### Data Type

 OLE_COLOR

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BackColorBkg Property](#backcolorbkg-property-gridex-control), [BackColorGBBox Property](#backcolorgbbox-property-gridex-control), [BackColorHeader Property](#backcolorheader-property-gridex-control), [BackColorInfoText Property](#backcolorinfotext-property-gridex-control), [BackColor Property](#backcolor-property-gridex-control), [BackColor Property](Objects/JSFormatStyle-Object.md#backcolor-property-jsformatstyle-object), [BackColorRowGroup Property](#backcolorrowgroup-property-gridex-control), [ForeColorHeader Property](#forecolorheader-property-gridex-control), [ForeColorInfoText Property](#forecolorinfotext-property-gridex-control), [ForeColor Property](Objects/JSFormatStyle-Object.md#forecolor-property-jsformatstyle-object), [ForeColor Property](#forecolor-property-gridex-control), [ForeColorRowGroup Property](#forecolorrowgroup-property-gridex-control), [RowColorEven Property](#rowcoloreven-property-gridex-control), [RowColorOdd Property](#rowcolorodd-property-gridex-control), [GridLinesColor Property](#gridlinescolor-property-gridex-control)  
**Example:** [GridImages Example](../Examples.md#gridimages-example)

## MultiSelect Property (GridEX Control)

Returns/sets a value indicating whether a user can make multiple selections in a **GridEX** control.

### Syntax

 *object*.**MultiSelect** [ = *value* ]  
 The **MultiSelect** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether the user can make multiple selections, as described in Settings. |

### Settings

 The settings for *value* are:

| Setting | Description |
| --- | --- |
| **True** | The user can select multiple rows at a time. |
| **False** | (Default) Selecting multiple rows is not allowed. |

### Remarks

 Pressing SHIFT and clicking the mouse or pressing SHIFT and one of the arrow keys (UP ARROW, DOWN ARROW, LEFT ARROW, and RIGHT ARROW) extends the selection from the previously selected row to the current row.  
 Pressing CTRL and clicking the mouse selects or deselects individual rows in the list.  
 Also, pressing the mouse in leftmost area of a row, when the mouse pointer is an inverted arrow, and dragging the cursor up or down while the button is still pressed, will allow the user to select a range of rows.

**Note** The **SelectedItems** property in a **GridEX** control is only available when the **MultiSelect** property is set to **True**.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [SelectedItems Property](#selecteditems-property-gridex-control), [SelectionChange Event](Events.md#selectionchange-event-gridex-control)  
**Example:** [SelectedItems Example](../Examples.md#selecteditems-example)

## NewRowPos Property (GridEX Control)

Returns or sets the positioning of the row the user can use to add new records.

### Syntax

 *object*.**NewRowPos** [ = *value* ]  
 The **NewRowPos** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines the position of the row for new records, as described in Settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexTop** | 0 | The row for adding new records is positioned below the column headers row, separated from the other rows by a horizontal divider. |
|  **jgexBottom** | 1 | The row for adding new records is the last row in the grid. |

### Remarks

 If the **AllowAddNew** property is set to **False** or the underlying **Recordset** is not updateable, the **GridEX** control will hide the “new records row” regardless of the value of this property.  
 If the **NewRowPos** property setting is **jgexBottom**, the **GridEX** control will show the row for new records at the end of the grid, except when the records in the control are grouped, in which case the “new row” will not be displayed.  
 This property has no effect if the control is in card view because in that view users cannot add new records.

### Data Type

 **jgexNewRowPosConstants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AllowAddNew Property](#allowaddnew-property-gridex-control)  
**Example:** [View Example](../Examples.md#view-example)

## OLEDropMode Property (GridEX Control)

Returns or sets how a **GridEX** control handles drop operations.

### Syntax

 *object*.**OLEDropMode** [ = *value* ]  
 The **OLEDropMode** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| value | A value or constant that determines if a **GridEX** control handles or not OLE drop operations, as described in Settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexOLEDropNone** | 0 | (Default) None. The **GridEX** control does not accept OLE drops and displays the No Drop cursor. |
|  **jgexOLEDropManual** | 1 | Manual. The **GridEX** control triggers the OLE drop events, allowing the programmer to handle the OLE drop operation in code. |

Remarks  
 The **GridEX** control does not support automatic **OLEDropMode**. However, if you set the **OLEDropMode** property to **jgexOLEDropManual** you will receive OLE Drag/Drop events in order to let you manage the operation in code.

### Data Type

 **jgexOLEDropModeConstants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [OLECompleteDrag Event](Events.md#olecompletedrag-event-gridex-control), [OLEDragDrop Event](Events.md#oledragdrop-event-gridex-control), [OLEDrag Method](Methods.md#oledrag-method-gridex-control), [OLEDragOver Event](Events.md#oledragover-event-gridex-control), [OLEGiveFeedback Event](Events.md#olegivefeedback-event-gridex-control), [OLESetData Event](Events.md#olesetdata-event-gridex-control), [OLEStartDrag Event](Events.md#olestartdrag-event-gridex-control)

## Options Property (GridEX Control)

Returns or sets a value that specifies one or more characteristics of the **Recordset** object in the **GridEX** control.

### Syntax

 *object*.**Options** [ = *value*]  
 The **Options** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A long value that specifies the characteristics of the **Recordset**. |

### Remarks

 For more information about the possible values that can be used with this property, please refer to the appropriate DAO and/or ADO documentation.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ADORecordset Property](#adorecordset-property-gridex-control), [Recordset Property](#recordset-property-gridex-control)

## PreviewColumn Property (GridEX Control)

Returns or sets the index or key of the column to be displayed as the preview row.

### Syntax

 *object*.**PreviewColumn** [ = *value* ]  
 The **PreviewColumn** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A variant that represents the index or key of the preview row column. |

### Remarks

 This property is only used in table view.  
 If **PreviewRowLines** property is equal to 0, this property has no effect.

### Data Type

 Variant

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [PreviewRowIndent Property](#previewrowindent-property-gridex-control), [PreviewRowLines Property](#previewrowlines-property-gridex-control), [PreviewRowVisible Property](Objects/JSRowData-Object.md#previewrowvisible-property-jsrowdata-object)  
**Example:** [PreviewColumn Example](../Examples.md#previewcolumn-example)

## PreviewRowIndent Property (GridEX Control)

Returns or sets the left indent, in twips, for text in the preview row.

### Syntax

 *object*.**PreviewRowIndent** [ = *value* ]  
 The **PreviewRowIndent** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Long expression specifying the left indent of text in preview rows. |

### Remarks

 This property is only used when GridEX is in table view.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [PreviewColumn Property](#previewcolumn-property-gridex-control), [PreviewRowLines Property](#previewrowlines-property-gridex-control), [PreviewRowVisible Property](Objects/JSRowData-Object.md#previewrowvisible-property-jsrowdata-object)  
**Example:** [PreviewColumn Example](../Examples.md#previewcolumn-example)

## PreviewRowLines Property (GridEX Control)

Returns or sets the number of lines to be displayed in preview rows.

### Syntax

 *object*.**PreviewRowLines** [ = *value* ]  
 The **PreviewRowLines** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | An integer that represents the maximum number of lines in preview rows. |

### Remarks

 This property is only used in table view.  
 If **PreviewColumn** property is Null, this property has no effect.

### Data Type

 Variant

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [PreviewColumn Property](#previewcolumn-property-gridex-control), [PreviewRowIndent Property](#previewrowindent-property-gridex-control), [PreviewRowVisible Property](Objects/JSRowData-Object.md#previewrowvisible-property-jsrowdata-object)  
**Example:** [PreviewColumn Example](../Examples.md#previewcolumn-example)

## PrinterProperties Property (GridEX Control)

Returns the **JSPrinterProperties** object of a **GridEX** control.

### Syntax

 *object*.**PrinterProperties**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The **PrinterProperties** property returns a **JSPrinterProperties** object.  
 **JSPrinterProperties** object is used to specify printer specific properties like color mode, orientation, paper size and the page layout for the document to be printed.

### Data Type

 **JSPrinterProperties**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [PrintGrid Method](Methods.md#printgrid-method-gridex-control), [PrintPreview Method](Methods.md#printpreview-method-gridex-control)  
**Example:** [PrinterProperties Example](../Examples.md#printerproperties-example)

## ReadOnly Property (GridEX Control)

Returns or sets a value indicating whether the **Database** object in a **GridEX** control must be opened as read only.

### Syntax

 *object*.**ReadOnly** [ = *value* ]  
 The **ReadOnly** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether the underlying Database object is opened as read only, as described in Settings. |

### Settings

 The settings for *value* are:

| Setting | Description |
| --- | --- |
| **True** | The **Database** object underlying the **GridEX** control is opened as read only. |
| **False** | (Default) The **Database** object underlying the **GridEX** control is opened as read write. |

### Remarks

 This property is used in DAO mode only. The **ReadOnly** property is used to open the **Database** object that holds the **Recordset**.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [DataMode Property](#datamode-property-gridex-control), [Recordset Property](#recordset-property-gridex-control)  
**Example:** [RebindSample Example](../Examples.md#rebindsample-example)

## RecordNavigator Property (GridEX Control)

Returns or sets a value indicating whether the record navigator is visible or hidden.

### Syntax

 *object*.**RecordNavigator** [ = *value* ]  
 The **RecordNavigator** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether the record navigator is visible or hidden, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | Record Navigator is visible. |
| **False** | Default. Record Navigator is hidden. |

### Remarks

 When the **RecordNavigator** property is set to **True**, the **GridEX** control shows buttons at the bottom of the control to navigate through the records (MoveFirst, MovePrevious, MoveNext, MoveLast and AddNew). In addition the Record Navigator has a specific record box where the user can move to a record by entering its position in the grid.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RecordNavigatorString Property](#recordnavigatorstring-property-gridex-control)

## RecordNavigatorString Property (GridEX Control)

Returns or sets the text displayed in the record navigator.

### Syntax

 *object*.**RecordNavigatorString** [ = *value* ]  
 The **RecordNavigatorString** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A string that represents the text that will be displayed in the record navigator. |

### Remarks

 The **RecordNavigatorString** is displayed in the record navigator in order to give some information to the users.  
 The string is formed by two parts separated by a "|" character.  
 You can change this property to any string for customization. The default value is “Record:|of”.  
 For instance, you can change this value to "Order:|of" if you are presenting orders in the **GridEX** control or you can change this string for localization.

### Data Type

 String

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RecordNavigator Property](#recordnavigator-property-gridex-control)

## RecordSource Property (GridEX Control)

Returns or sets the underlying table name, SQL statement, or QueryDef object for the **Recordset** in a **GridEX** control.

### Syntax

 *object*.**RecordSource** [ = *value* ]  
 The **RecordSource** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A string expression specifying the name of a table, QueryDef or a SQL statement. |

### Remarks

 The **RecordSource** property specifies the source of the **Recordset** in DAO and ADO modes. This property has no use in unbound mode.  
 After changes at run time for the value of the **RecordSource** property, the **Rebind** method must be used to rebuild the **Recordset** object and reflect the new settings.

### Data Type

 String

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [DataMode Property](#datamode-property-gridex-control), [Recordset Property](#recordset-property-gridex-control), [ADORecordset Property](#adorecordset-property-gridex-control), [DatabaseName Property](#databasename-property-gridex-control), [Rebind Method](Methods.md#rebind-method-gridex-control)  
**Example:** [RebindSample Example](../Examples.md#rebindsample-example)

## Recordset Property (GridEX Control)

Returns or sets a **DAO.Recordset** object defined by a **GridEX** control's properties or by an existing **DAO.Recordset** object.

### Syntax

 *object*.**Recordset** [ = *value* ]  
 The **Recordset** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | An object variable containing a **DAO.Recordset** object. |

### Remarks

 The **GridEX** control is initialized when the control first appears. If the **DataMode** property setting is **jgexDAO** and the **DatabaseName**, **RecordSource**, **Exclusive**, **ReadOnly**, **LockType**, **Connect** and **RecordsetType** properties are valid, the control attempts to create a new **Recordset** object based on those properties. This also happens when you set these properties at runtime and use the **Rebind** Method. A clone of this Recordset is accessible through the **GridEX** control's **Recordset** property.  
 However, if one or more of these properties are set incorrectly, an **Error** event is sent to the application when the control attempts to use the properties to create the **Recordset** object.  
 You can also request the type of **Recordset** to be created by setting the **GridEX** control's **RecordsetType** property. If you do not request a specific type, a Dynaset-type **Recordset** is created by default. Using the **RecordsetType** property, you can request to create either a Snapshot-, or Table-type **Recordset**. However, if the control cannot create the type requested, an **Error** event occurs.  
 If you create an independent **Recordset** object in code, you can set the **Recordset** property of the control to this new **Recordset**. The existing Recordset in the control (if any) is released and a clone of the **Recordset** assigned to it replaces it.  
 If you set this property at run time the **DataMode** property of the control is changed to **jgexDAO**, no matter the previous setting for this property. However, if you try to read the **Recordset** property when the control’s mode is other than DAO, a trappable error occurs.

**Note** You must avoid using the **Close** method of the variable set to the **Recordset** property because it will invalidate this object for any further operations until the **Rebind** method is used.

### Data Type

 **Object**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Connect Property](#connect-property-gridex-control), [DatabaseName Property](#databasename-property-gridex-control), [DataMode Property](#datamode-property-gridex-control), [LockType Property](#locktype-property-gridex-control), [Rebind Method](Methods.md#rebind-method-gridex-control), [ADORecordset Property](#adorecordset-property-gridex-control), [RecordsetType Property](#recordsettype-property-gridex-control), [RecordSource Property](#recordsource-property-gridex-control), [Exclusive Property](#exclusive-property-gridex-control), [Options Property](#options-property-gridex-control), [ReadOnly Property](#readonly-property-gridex-control)  
**Example:** [Recordset Example](../Examples.md#recordset-example)

## RecordsetType Property (GridEX Control)

Returns or sets a value indicating the type of **Recordset** object you want to create through the **GridEX** control.

### Syntax

 *object*.**RecordsetType** [ = *value* ]  
 The **RecordsetType** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A constant or value that specifies a type of **Recordset**, as described in Settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexRSDAOTable** | 1 | A table-type Recordset (DAO mode only). |
|  **jgexRSDAODynaset** | 2 | (Default for DAO mode) A dynaset-type Recordset |
|  **jgexRSDAOSnapshot** | 4 | A snapshot-type Recordset (DAO mode only). |
|  **jgexRSADOKeyset** | 1 | (Default for ADO mode) A Keyset-cursor type. |
|  **jgexRSADOStatic** | 3 | A Static-cursor type (ADO mode only). |

### Remarks

 If you try to set a **Recordset** type that does not match with the **GridEX** control’s **DataMode** an error occurs.  
 For example, if you set the **RecordsetType** property to Static-cursor type and the **DataMode** property is set to DAO an error occurs. However, if you set **RecordsetType** property to Keyset-cursor type while the control’s **DataMode** is DAO no error occurs but the control assumes that the **RecordsetType** is table-type.

**Note** This property is ignored when **GridEX** is in unbound mode.

### Data Type

 **jgexRecordsetTypeConstants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [DataMode Property](#datamode-property-gridex-control), [Recordset Property](#recordset-property-gridex-control), [ADORecordset Property](#adorecordset-property-gridex-control), [Rebind Method](Methods.md#rebind-method-gridex-control)  
**Example:** [RebindSample Example](../Examples.md#rebindsample-example)

## Redraw Property (GridEX Control)

Returns or sets a value that determines whether drawing is enabled or not in a **GridEX** control.

### Syntax

 *object*.**Redraw** [ = *value* ]  
 The **Redraw** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether drawing is enabled or not, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | (Default) The **GridEX** control redraws after a change. |
| **False** | The **GridEX** control doesn't redraw ater a change. |

### Remarks

 By setting **Redraw** property to **False**, you can reduce the flickering produced when several **GridEX** settings are changed. It is very important to set the **Redraw** property back to **True** in order to enable the **GridEX** control to draw itself correctly.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Refresh Method](Methods.md#refresh-method-gridex-control)  
**Example:** [Redraw Example](../Examples.md#redraw-example)

## ReplaceColumnIndex Property (GridEX Control)

Returns or sets the index or key of the column that replaces Id values in a DropDown **GridEX** Control.

### Syntax

 *object*.**ReplaceColumnIndex** [ = *value* ]  
 The **ReplaceColumnIndex** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A variant that represents the index or key of the replace values column. |

### Remarks

 This property is only used when **ActAsDropDown** property is **True**.  
 By setting this property, a column with **EditType = jgexCombo** in another GridEX control will show the value of the replaced column instead of the Id value retrieved from the recordset.

### Data Type

 Variant

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BoundColumnIndex Property](#boundcolumnindex-property-gridex-control), [ActAsDropDown Property](#actasdropdown-property-gridex-control), [EditType Property](Objects/JSColumn-Object.md#edittype-property-jscolumn-object), [ReplaceValues Property](Objects/JSColumn-Object.md#replacevalues-property-jscolumn-object)  
**Example:** [ActAsDropDown Example](../Examples.md#actasdropdown-example)

## Row Property (GridEX Control)

Returns or sets the current row/card position in a **GridEX** control. Not available at design time.

### Syntax

 *object*.**Row** [ = *value*]  
 The **Row** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A long that represents the row/card position of the current row in a **GridEX** control. |

### Remarks

 The Row property in **GridEX** represents the position of a record in a **GridEX** control and not its original position in the **Recordset**.  
 If you need to obtain the original position of a given row, you must use the **RowIndex** method as follows:

```vb
Dim RowPosition as Long
Dim RowIndex as Long

RowPosition = GridEX1.Row
RowIndex = GridEX1.RowIndex(RowPosition)
Debug.print RowPosition,RowIndex
```

As result of sorting and/or grouping operations within the **GridEX**, the original position of the record may change and therefore be different to its index.  
 In some cases, as when a **GridEX** control is grouped, the **Row** property can return a value greater than the **ItemCount** setting. This is due to the fact that the **ItemCount** property does not take into account the number of group rows. If you want to find out the number of rows displayed use the **RowCount** property instead.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RowIndex Method](Methods.md#rowindex-method-gridex-control), [Col Property](#col-property-gridex-control), [ItemCount Property](#itemcount-property-gridex-control), [RowCount Property](#rowcount-property-gridex-control), [GetRowData Method](Methods.md#getrowdata-method-gridex-control)  
**Example:** [SortKeys Example](../Examples.md#sortkeys-example)

## RowBookmark Property (GridEX Control)

Returns or sets a value containing a bookmark for a row in the **GridEX** control. Not available at design time

### Syntax

 *object*.**RowBookmark**(*rowindex*) [ = *value* ]  
 The **RowBookmark** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *rowindex* | Required. A long in the range of 1 to the value of the **ItemCount** property. |
| *value* | A variant expression that represents the bookmark of a row. |

### Remarks

 The **RowBookmark** property is read-only for DAO and ADO modes and is read-write for unbound mode.  
 In unbound mode, you can use this value as a tag property to identify individual rows.  
 In bound mode, you can get the **Bookmark** of any row to position the **Recordset** clone of a **GridEX** control as follows:

```vb
Dim varBookmark as String
Dim RowIndex as Long
Dim rstCustomers as Recordset

Set rstCustomers = GridEX1.Recordset
RowIndex = GridEX1.RowIndex(GridEX1.Row)

varBookmark = GridEX1.RowBookmark(RowIndex)
rstCustomers.Bookmark = varBookmark
debug.print rstCustomers![CustomerId]
```

### Data Type

 Variant

**Note** The rowindex parameter in the **RowBookmark** property represents the original position of a record. As result of sorting and/or grouping operations within the **GridEX** control, the original position of the record may change, so the rowindex may not always match its positional index.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RefreshRowBookmark Method](Methods.md#refreshrowbookmark-method-gridex-control), [RefreshRowIndex Method](Methods.md#refreshrowindex-method-gridex-control), [Refresh Method](Methods.md#refresh-method-gridex-control), [MoveToBookmark Method](Methods.md#movetobookmark-method-gridex-control), [Row Property](#row-property-gridex-control), [RowIndex Method](Methods.md#rowindex-method-gridex-control)  
**Example:** [RowIndex Example](../Examples.md#rowindex-example)

## RowColorEven Property (GridEX Control)

Returns or sets the background color for even and odd rows when **UseEvenOddColor** setting is **True**.

### Syntax

 *object*.**RowColorEven** [ =*color* ]  
 The **RowColorEven** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A numeric expression that specifies the color. |

### Remarks

 Use these properties to specify the background color for alternating rows.  
 These properties are used only when the **UseEvenOddColor** property setting is **True**.

### Data Type

 OLE_COLOR

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BackColorBkg Property](#backcolorbkg-property-gridex-control), [BackColorGBBox Property](#backcolorgbbox-property-gridex-control), [BackColorHeader Property](#backcolorheader-property-gridex-control), [BackColorInfoText Property](#backcolorinfotext-property-gridex-control), [BackColor Property](#backcolor-property-gridex-control), [BackColor Property](Objects/JSFormatStyle-Object.md#backcolor-property-jsformatstyle-object), [BackColorRowGroup Property](#backcolorrowgroup-property-gridex-control), [ForeColorHeader Property](#forecolorheader-property-gridex-control), [ForeColorInfoText Property](#forecolorinfotext-property-gridex-control), [ForeColor Property](Objects/JSFormatStyle-Object.md#forecolor-property-jsformatstyle-object), [ForeColor Property](#forecolor-property-gridex-control), [ForeColorRowGroup Property](#forecolorrowgroup-property-gridex-control), [MaskColor Property](#maskcolor-property-gridex-control), [RowColorOdd Property](#rowcolorodd-property-gridex-control), [GridLinesColor Property](#gridlinescolor-property-gridex-control)

## RowColorOdd Property (GridEX Control)

Returns or sets the background color for even and odd rows when **UseEvenOddColor** setting is **True**.

### Syntax

 *object*.**RowColorOdd** [ =*color* ]  
 The **RowColorEven** and **RowColorOdd** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *color* | A numeric expression that specifies the color. |

### Remarks

 Use these properties to specify the background color for alternating rows.  
 These properties are used only when the **UseEvenOddColor** property setting is **True**.

### Data Type

 OLE_COLOR

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BackColorBkg Property](#backcolorbkg-property-gridex-control), [BackColorGBBox Property](#backcolorgbbox-property-gridex-control), [BackColorHeader Property](#backcolorheader-property-gridex-control), [BackColorInfoText Property](#backcolorinfotext-property-gridex-control), [BackColor Property](#backcolor-property-gridex-control), [BackColor Property](Objects/JSFormatStyle-Object.md#backcolor-property-jsformatstyle-object), [BackColorRowGroup Property](#backcolorrowgroup-property-gridex-control), [ForeColorHeader Property](#forecolorheader-property-gridex-control), [ForeColorInfoText Property](#forecolorinfotext-property-gridex-control), [ForeColor Property](Objects/JSFormatStyle-Object.md#forecolor-property-jsformatstyle-object), [ForeColor Property](#forecolor-property-gridex-control), [ForeColorRowGroup Property](#forecolorrowgroup-property-gridex-control), [MaskColor Property](#maskcolor-property-gridex-control), [RowColorEven Property](#rowcoloreven-property-gridex-control), [GridLinesColor Property](#gridlinescolor-property-gridex-control)

## RowCount Property (GridEX Control)

Returns the number of rows in a **GridEX** control. Read Only

### Syntax

 *object*.**RowCount**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 When the records in a **GridEX** control are grouped, the number of rows in a **GridEX** is different to the number of records.  
 For example, if you have 100 records in a grid, and there are four groups and all groups are collapsed, the **RowCount** property will return 4, even when the **ItemCount** property value is 100. When these four groups are expanded the **RowCount** will be **ItemCount** + Number of Groups, or 100 + 4 = 104.

**Note** The **RowCount** property is always equal to the **ItemCount** property if the **GridEX** control is not grouped.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ItemCount Property](#itemcount-property-gridex-control), [LoadEntireRecordset Method](Methods.md#loadentirerecordset-method-gridex-control), [FullyLoaded Property](#fullyloaded-property-gridex-control)  
**Example:** [RowExpanded Example](../Examples.md#rowexpanded-example)

## RowExpanded Property (GridEX Control)

Returns or sets whether a group row in a **GridEX** control is expanded or collapsed. Not available at design time

### Syntax

 *object*.**RowExpanded** ( *rowposition* ) [ = *value* ]  
 The **RowExpanded** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *rowposition* | Required. A long in the range of 1 to the setting of the **RowCount** property. |
| *value* | A Boolean expression that determines if a group row is expanded or collapsed. |

### Remarks

 If the rowposition parameter evaluates to a regular row this property allways returns **True** and setting this property to **False** has no effect.  
 If the group row is not already expanded, setting **RowExpanded** to **True** will expand the group row.  
 If the group row is already expanded, setting **RowExpanded** to False will close it.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [IsGroupItem Method](Methods.md#isgroupitem-method-gridex-control), [ExpandAll Method](Methods.md#expandall-method-gridex-control), [GroupRowLevel Method](Methods.md#grouprowlevel-method-gridex-control), [CollapseAll Method](Methods.md#collapseall-method-gridex-control)  
**Example:** [RowExpanded Example](../Examples.md#rowexpanded-example)

## RowHeaders Property (GridEX Control)

Returns or sets a value indicating whether row headers are displayed in a **GridEX** control.

### Syntax

 *object*.**RowHeaders** [ = *value*]  
 The **RowHeaders** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether row headers are displayed, as described in Settings. |

### Settings

 The settings for *value* are:

| Setting | Description |
| --- | --- |
| **True** | **GridEX** control's row headers are displayed |
| **False** | **GridEX** control's row headers aren't displayed |

### Remarks

 This property has no effect while the **GridEX** control is in card view.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ColumnHeaders Property](#columnheaders-property-gridex-control)  
**Example:** [View Example](../Examples.md#view-example)

## RowHeight Property (GridEX Control)

Returns or sets the height, in twips, of rows in a **GridEX** control.

### Syntax

 *object*.**RowHeight** [ = *value*]  
 The **RowHeight** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Long expression that determined the height, in twips, of rows in a **GridEX** control. |

### Remarks

 By setting this property to -1, the **GridEX** control calculates the height of each row based on its row style font settings and **ImageHeight**.  
 Although the height of rows is automatically calculated by the **GridEX** control to accommodate the font extent and image heights, you can use this property to set a different height than the default one.  
 This property affects all the rows in the control.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RowHeight Property](Objects/JSRowData-Object.md#rowheight-property-jsrowdata-object), [ImageHeight Property](#imageheight-property-gridex-control), [AllowRowSizing Property](#allowrowsizing-property-gridex-control)

## RowSelected Property (GridEX Control)

Returns or sets whether a row in a **GridEX** control is selected. Not available at design time

### Syntax

 *object*.**RowSelected** (*rowposition*)[ = *value* ]  
 The **RowSelected** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| rowposition | Required. A long in the range of 1 to the setting of the RowCount property. |
| *value* | A Boolean expression that determines if a row is selected. |

### Remarks

 When using this property in a non-multiselect **GridEX** control has the same effect as setting the **Row** property to the *rowposition* parameter.  
 If you want to iterate through the selected rows in a **GridEX** control, use the **JSSelectedItems** collection instead of testing the selected state of each row in the control.  
 Setting this property has the same effect as using the **Add** method in the **JSSelectedItems** collection.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Row Property](#row-property-gridex-control), [SelectedItems Property](#selecteditems-property-gridex-control), [Add Method](Objects/JSSelectedItems-Collection.md#add-method-jsselecteditems-collection), [MultiSelect Property](#multiselect-property-gridex-control)  
**Example:** [RowSelected Example](../Examples.md#rowselected-example)

## ScrollToolTip Column Property (GridEX Control)

Returns or sets the index or key of the column used to get values for scroll tool tips.  
 Syntax  
 *object*.**ScrollToolTipColumn** [ = *value* ]  
 The **ScrollToolTipColumn** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A variant that represents the index or key of the column used to to get values for scroll tool tips. |

### Remarks

 This property must be used in conjunction with **ScrollToolTips** property.  
 The tool tip text displayed when **ScrollToolTips** is set to **True** and **ScrollToolTipColumn** is set to a valid column index is formed by the caption of the column and the value of that column in an specific row.

### Data Type

 Variant

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ContinuousScroll Property](#continuousscroll-property-gridex-control), [ScrollToolTips Property](#scrolltooltips-property-gridex-control), [ShowToolTips Property](#showtooltips-property-gridex-control)  
**Example:** [ScrollTollTips Example](../Examples.md#scrolltolltips-example)

## ScrollToolTips Property (GridEX Control)

Returns or sets a value indicating whether the **GridEX** control should show tool tips while the user scrolls vertically.

### Syntax

 *object*.**ScrollToolTips** [ = *value* ]  
 The **ScrollToolTips** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether the **GridEX** control should display tool tips while the user scrolls vertically, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | The **GridEX** control shows tool tips while the user moves the scroll thumb. |
| **False** | Default. The **GridEX** control doesn't show tool tips while the user moves the scroll thumb. |

### Remarks

 To see Scroll tool tips, in adition to set this property to **True**, the **ScrollToolTipColumn** must be set to indicate the column where values are retrieved to get the text displayed in the tool tip.  
 In table view, **ScrollToolTips** are shown over the vertical scroll bar.  
 In card view, **ScrollToolTips** are shown over the horizontal scroll bar.

**Note** **ScrollToolTips** won't be displayed if **ContinousScroll** property is set to **True**.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ContinuousScroll Property](#continuousscroll-property-gridex-control), [ScrollToolTip Column Property](#scrolltooltip-column-property-gridex-control), [ShowToolTips Property](#showtooltips-property-gridex-control)  
**Example:** [ScrollTollTips Example](../Examples.md#scrolltolltips-example)

## SelectedItems Property (GridEX Control)

Returns the **JSSelectedItems** collection in a **GridEX** control.  
 This property is not available in a **GridEX** control whose **MultiSelect** property is set to **False**.

### Syntax

 *object*.**SelectedItems**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The **SelectedItems** property is only available when the **MultiSelect** property in a **GridEX** control is set to **True**.  
 The **SelectedItems** property returns a **JSSelectedItems** object.  
 The **JSSelectedItems** object is a collection of **JSSelectedItem** objects in a **GridEX** control.  
 You can access specific objects in this collection (which represent selected rows) through the **Item** property.  
 Using the **JSSelectedItems** collection you can select a row using its current position in the control, the original record index or using its bookmark.

### Data Type

 **JSSelectedItems**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [MultiSelect Property](#multiselect-property-gridex-control), [JSSelectedItems Collection](Objects/JSSelectedItems-Collection.md#jsselecteditems-collection), [RowSelected Property](#rowselected-property-gridex-control), [SelectionChange Event](Events.md#selectionchange-event-gridex-control)  
**Example:** [SelectedItems Example](../Examples.md#selecteditems-example)

## SelectionStyle Property (GridEX Control)

Returns or sets the style used by a **GridEX** control to highlight a selected cell.

### Syntax

 *object*.**SelectionStyle** [ = *value*]  
 The **SelectionStyle** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A constant or value that specifies the way a **GridEX** control will highlight a selected cell, as described in Settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexEntireRow** | 0 | (Default) When a cell is selected, the entire row is highlighted. |
|  **jgexSingleCell** | 1 | When a cell is selected, only that cell is highlighted. |

### Data Type

 Integer

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [HideSelection Property](#hideselection-property-gridex-control), [Row Property](#row-property-gridex-control), [Col Property](#col-property-gridex-control)

## SelLength Property (GridEX Control)

Returns or sets the number of characters selected. Not available at design time.

### Syntax

 *object*.**SelLength** [ = *value* ]  
 The **SelLength** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Long expression specifying the number of characters selected. <br> <br> The valid range of settings is 0 to the total number of characters in the cell. |

### Remarks

 This property has no effect when a **GridEX** control’s **EditMode** is set to **jgexEditModeOff**.  
 Use this property in conjunction with the **SelStart** property for tasks such as setting the insertion point, establishing an insertion range or selecting substrings in a cell editor.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [EditMode Property](#editmode-property-gridex-control), [EditType Property](Objects/JSColumn-Object.md#edittype-property-jscolumn-object), [SelStart Property](#selstart-property-gridex-control), [SelText Property](#seltext-property-gridex-control)  
**Example:** [EditMode Example](../Examples.md#editmode-example)

## SelStart Property (GridEX Control)

Returns or sets the starting point of text selected in a cell editor, or indicates the position of the insertion point if no text is selected. Not available at design time.

### Syntax

 *object*.**SelStart** [ = *value* ]  
 The **SelStart** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Long expression specifying the start character of a selection or the insertion point in a cell editor. <br> <br> The valid range of settings is 0 to the total number of characters in the cell. |

### Remarks

 This property has no effect when a **GridEX** control’s **EditMode** is set to **jgexEditModeOff**.  
 Use this property in conjunction with the **SelLength** property for tasks such as setting the insertion point, establishing an insertion range or selecting substrings in a cell editor.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [EditMode Property](#editmode-property-gridex-control), [EditType Property](Objects/JSColumn-Object.md#edittype-property-jscolumn-object), [SelText Property](#seltext-property-gridex-control), [SelLength Property](#sellength-property-gridex-control)  
**Example:** [EditMode Example](../Examples.md#editmode-example)

## SelText Property (GridEX Control)

Returns or sets the string containing the currently selected text in the cell editor. Not available at design time.

### Syntax

 *object*.**SelText** [ = *value* ]  
 The **SelText** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A String expression containing the selected text. |

### Remarks

 This property has no effect when a **GridEX** control’s **EditMode** is set to **jgexEditModeOff**.  
 Setting **SelText** to a new value sets **SelLength** to 0 and replaces the selected text with the new string.  
 This property is useful for copy, cut, and paste operations.

### Data Type

 String

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [EditMode Property](#editmode-property-gridex-control), [EditType Property](Objects/JSColumn-Object.md#edittype-property-jscolumn-object), [SelStart Property](#selstart-property-gridex-control), [SelLength Property](#sellength-property-gridex-control)

## ShowEmptyFields Property (GridEX Control)

Returns or sets a value indicating whether the **GridEX** control will display empty fields in cards.

### Syntax

 *object*.**ShowEmptyFields** [ = *value* ]  
 The **ShowEmptyFields** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether empty fields are shown in a card, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | Default. The GridEX control shows all fields in a card. |
| **False** | The GridEX control does not show fields that are empty in a card. |

### Remarks

 This property is only used when the **GridEX** control is in card view.  
 If **ShowEmptyFields** is set to **False**, cells with values equal to an empty string or Null, are not shown in a card.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [View Property](#view-property-gridex-control)  
**Example:** [View Example](../Examples.md#view-example)

## ShowToolTips Property (GridEX Control)

Returns or sets a value indicating whether the **GridEX** control should show tool tips over cells where text is not entirely displayed.

### Syntax

 *object*.**ShowToolTips** [ = *value* ]  
 The **ShowToolTips** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether the **GridEX** control should show tool tips over cells where text is not entirely displayed, as described in Settings. |

### Settings

 The settings for value are:

| Setting | Description |
| --- | --- |
| **True** | The GridEX control shows a tool tip when mouse pointer is over a cell with truncated text. |
| **False** | Default. The GridEX control doesn't show tool tips over cells. |

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [HeaderToolTip Property](Objects/JSColumn-Object.md#headertooltip-property-jscolumn-object), [ScrollToolTips Property](#scrolltooltips-property-gridex-control), [ScrollToolTip Column Property](#scrolltooltip-column-property-gridex-control)

## SortKeys Property (GridEX Control)

Returns the **JSSortKeys** object in a **GridEX** control.

### Syntax

 *object*.**SortKeys**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The **SortKeys** property returns a **JSSortKeys** collection.  
 You can get a specific **JSSortKey** object through the **Item** property in **JSSortKeys** object.  
 You can also add and delete **JSSortKey** objects.

### Data Type

 **JSSortKeys**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [JSSortKeys Collection](Objects/JSSortKeys-Collection.md#jssortkeys-collection), [JSSortKey Object](Objects/JSSortKey-Object.md#jssortkey-object)  
**Example:** [SortKeys Example](../Examples.md#sortkeys-example)

## TabKeyBehavior Property (GridEX Control)

Sets or returns a value that determines the behavior of the tab key.

### Syntax

 *object*.**TabKeyBehavior** [ = *value* ]  
 The **TabKeyBehavior** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines the behavior of the tab key, as described in Settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexColumnNavigation** | 0 | (Default) The tab key moves the current cell to the next or previous column. |
|  **jgexControlNavigation** | 1 | The tab key moves to the next control on the form. |

### Data Type

 **jgexTabKeyBehaviorConstants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Selectable Property](Objects/JSColumn-Object.md#selectable-property-jscolumn-object)

## UseEvenOddColor Property (GridEX Control)

Returns or sets a value indicating whether alternating background colors will be used in a **GridEX** control.

### Syntax

 *object*.**UseEvenOddColor** [ = *value* ]  
 The **UseEvenOddColor** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Boolean expression that determines whether rows will be displayed with alternating background colors, as described in Settings. |

### Settings

 The settings for *value* are:

| Setting | Description |
| --- | --- |
| **True** | The control alternates background row colors. |
| **False** | he control will use **GridEX**'s **BackColor** for all rows. |

### Remarks

 When the **GridEX** control is in card view this property has no effect.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RowColorEven Property](#rowcoloreven-property-gridex-control), [RowColorOdd Property](#rowcolorodd-property-gridex-control)

## Value Property (GridEX Control)

Returns or sets the value of a column in the current row of a **GridEX** control.

### Syntax

 *object*.**Value** (*colindex*) [ = *value*]  
 The **Value** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *colindex* | An integer that represent the index of the column. |
| *value* | A variant expression specifying the value in the column of the current row. |

### Data Type

 Variant

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Index Property](Objects/JSColumn-Object.md#index-property-jscolumn-object), [Value Property](Objects/JSRowData-Object.md#value-property-jsrowdata-object)  
**Example:** [BeforeUpdate Example](../Examples.md#beforeupdate-example)

## View Property (GridEX Control)

Returns or sets the way records are displayed in a **GridEX** control.  
 Syntax  
 *object*.**View** [ = *value*]  
 The **View** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that determines the way records are displayed in a **GridEX** control, as described in Settings. |

### Settings

 The settings for value are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexTable** | 0 | (Default) (Table view) The records are displayed as a table where each row represents a record. |
|  **jgexCard** | 1 | (Card view) The records are displayed as cards where each card represents a record and each row in the card represents a column. |

### Remarks

 If a **GridEX** control is in card view, regardless of the setting of the **AllowAddNew** property, the user will not be able to add new records. Also, the user will not be able to group records (as in table view). Any attempt to apply any grouping criteria while the **GridEX** control is in **Card** view will cause an error.  
 There are some properties that applies only to a particular view. For example, **CardWidth** only applies to card view, and **GroupByBoxVisible** applies to table view.  
 To find out if a view supports a particular property or method, search the documentation of the method.

**Note** In card view any reference to a row must be understood as a card.

### Data Type

 **jgexViewConstants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**Example:** [View Example](../Examples.md#view-example)
