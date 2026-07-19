# GridEX Control — Methods

## AboutBox Method (GridEX Control)

Displays the About box for the control.

### Syntax

 *object*.**AboutBox**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 This is the same as double-clicking the (About) entry in the Properties window.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)

## CellHeight Method (GridEX Control)

Returns the height, in twips, of a given cell.

### Syntax

 *object*.**CellHeight** ( *rowposition*, *colposition* )  
 The **CellHeight** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *rowposition* | (Optional) A long expression that identifies the row position, if omitted; the row property value is used. |
| *colposition* | (Optional) An integer expression that identifies the column position, if omitted; the **Col** property value is used. |

### Remarks

 Use the **CellHeight** method to obtain the height of a particular cell. This method returns 0 if the cell specified is not visible.  
 To know if a cell is visible use the **CellVisible** method.  
 If you want to obtain the height of the current cell use this method without parameters.

**Note:** This method returns only the visible height of a cell if a cell is partially hidden.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [CellLeft Method](#cellleft-method-gridex-control), [CellTop Method](#celltop-method-gridex-control), [CellVisible Method](#cellvisible-method-gridex-control), [CellWidth Method](#cellwidth-method-gridex-control)

## CellLeft Method (GridEX Control)

Returns the left, in twips, of a given cell.

### Syntax

 *object*.**CellLeft** ( *rowposition*, *colposition* )  
 The **CellLeft** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *rowposition* | (Optional) A long expression that identifies the row position, if omitted; the row property value is used. |
| *colposition* | (Optional) An integer expression that identifies the column position, if omitted; the **Col** property value is used. |

### Remarks

 Use the **CellLeft** method to get the leftmost point of a particular cell. This method returns 0 if the cell specified is not visible. To know if a cell is visible use the **CellVisible** method.  
 This method returns the left position of a cell based on the left border of the **GridEX** control. If you want to place a control in a given cell you must add the value of the **Left** property of the control to the value returned by this property.  
 If you want to get the left of the current cell use this method without parameters.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [CellHeight Method](#cellheight-method-gridex-control), [CellTop Method](#celltop-method-gridex-control), [CellVisible Method](#cellvisible-method-gridex-control), [CellWidth Method](#cellwidth-method-gridex-control)

## CellTop Method (GridEX Control)

Returns the top, in twips, of a given cell.

### Syntax

 *object*.**CellTop** ( *rowposition*, *colposition* )  
 The **CellTop** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *rowposition* | (Optional) A long expression that identifies the row position, if omitted; the row property value is used. |
| *colposition* | (Optional) An integer expression that identifies the column position, if omitted; the **Col** property value is used. |

### Remarks

 Use the **CellTop** method to get the top of a particular cell. This method returns 0 if the cell specified is not visible. To know if a cell is visible use the **CellVisible** method.  
 This method returns the top of a cell based on the top border of the **GridEX** control. If you want to place a control in a given cell you must add the **Top** of the control to the value returned by this method.  
 If you want to get the left of the current cell use this method without parameters.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [CellHeight Method](#cellheight-method-gridex-control), [CellLeft Method](#cellleft-method-gridex-control), [CellVisible Method](#cellvisible-method-gridex-control), [CellWidth Method](#cellwidth-method-gridex-control)

## CellVisible Method (GridEX Control)

Returns **True** if the specified cell is visible or **False** otherwise.

### Syntax

 *object*.**CellVisible** ( *rowposition, colposition* )  
 The **CellVisible** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *rowposition* | (Optional) A long expression that identifies the row position, if omitted; the row property value is used. |
| *colposition* | (Optional) An integer expression that identifies the column position, if omitted; the **Col** property value is used. |

**Settings**:  
 The settings for **CellVisible** are:

| Setting | Description |
| --- | --- |
| **False** | The cell specified with the rowposition and colposition parameters is not visible. |
| **True** | The cell specified with the rowposition and colposition parameters is visible or partially visible. |

### Remarks

 Use the **CellVisible** method to know if a cell is visible or hidden at any given moment.  
 If you want to know if the current cell is visible, use this method without parameters.  
 If a cell is hidden and you want to make it visible, use the **EnsureVisible** method.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [CellHeight Method](#cellheight-method-gridex-control), [CellLeft Method](#cellleft-method-gridex-control), [CellTop Method](#celltop-method-gridex-control), [CellWidth Method](#cellwidth-method-gridex-control)

## CellWidth Method (GridEX Control)

Returns the width, in twips, of a given cell.

### Syntax

 *object*.**CellWidth** ( *rowposition, colposition* )  
 The **CellWidth** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *rowposition* | (Optional) A long expression that identifies the row position, if omitted; the row property value is used. |
| *colposition* | (Optional) An integer expression that identifies the column position, if omitted; the **Col** property value is used. |

### Remarks

 Use the **CellWidth** method to get the width of a particular cell. This method returns 0 if the cell specified is not visible. To know if a cell is visible use the **CellVisible** method.  
 If you want to get the width of the current cell use this method without parameters.  
 This method only returns the visible width of a cell if a cell is partially hidden.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [CellHeight Method](#cellheight-method-gridex-control), [CellLeft Method](#cellleft-method-gridex-control), [CellTop Method](#celltop-method-gridex-control), [CellVisible Method](#cellvisible-method-gridex-control)

## ClearFields Method (GridEX Control)

Restores the default layout for the **GridEX** control columns.

### Syntax

 *object*.**ClearFields**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The **ClearFields** method restores the default column layout (displaying two blank columns), so that subsequent **ReBind** operations will automatically derive new column bindings from the data source, which may have possibly changed. You can cancel the **GridEX** control’s automatic layout behavior by invoking the **HoldFields** method.

**Note:** If you want to clear all the columns in a **GridEX** control, use the **Clear** method of the **JSColumns** collection.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [HoldFields Method](#holdfields-method-gridex-control), [Rebind Method](#rebind-method-gridex-control)  
**Example:** [RebindSample Example](../Examples.md#rebindsample-example)

## ColFromPoint Method (GridEX Control)

Returns the **JSColumn** object whose position matches the specified coordinates.

### Syntax

 *object*.**ColFromPoint** *x, y*  
 The **ColFromPoint** method syntax has these parts:

| Part | Description |
| --- | --- |
| object | An objects expression that evaluates to an object in the Applies To list. |
| x,y | Coordinates of a target column. |

### Remarks

 If the point specified in the x and y parameters is not contained in any column the return value is Nothing.

### Data Type

 **JSColumn**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [HitTest Method](#hittest-method-gridex-control), [RowFromPoint Method](#rowfrompoint-method-gridex-control), [GroupFromPoint Method](#groupfrompoint-method-gridex-control)  
**Example:** [ColFromPoint Example](../Examples.md#colfrompoint-example)

## CollapseAll Method (GridEX Control)

Collapses all group rows in a **GridEX** control.

### Syntax

 *object*.**CollapseAll**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 This method is ignored if the **GridEX** control is not grouped.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ExpandSelection Method](#expandselection-method-gridex-control), [ExpandAll Method](#expandall-method-gridex-control), [DefaultGroupMode Property](Properties.md#defaultgroupmode-property-gridex-control), [RowExpanded Property](Properties.md#rowexpanded-property-gridex-control)  
**Example:** [CollapseAll Example](../Examples.md#collapseall-example)

## Delete Method (GridEX Control)

Deletes selected rows in a **GridEX** control.

### Syntax

 *object*.**Delete**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 Using the **Delete** method is the same as pressing the DEL key while the entire row is selected.

**Note** To select the entire row if the **AllowEdit** property setting is set to **True**, move the cursor to the extreme left of the row and the cursor image will change to an inverted arrow, allowing the selection of the entire row.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AllowDelete Property](Properties.md#allowdelete-property-gridex-control), [BeforeDeleteEx Event](Events.md#beforedeleteex-event-gridex-control), [BeforeDelete Event](Events.md#beforedelete-event-gridex-control), [Update Method](#update-method-gridex-control), [AfterDelete Event](Events.md#afterdelete-event-gridex-control)

## EnsureVisible Method (GridEX Control)

Ensures visibility of a particular cell. If necessary, this method scrolls **GridEX** control.

### Syntax

 *object*.**EnsureVisible**( *row, col* )  
 The **EnsureVisible** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *row* | (Optional) A long expression that identifies the row, if omitted; the current row is used. |
| *col* | (Optional) An integer expression that identifies the column, if omitted; the current column is used. |

### Remarks

 Use the **EnsureVisible** method when you want a particular cell, which might be hidden, to be visible.  
 Sometimes the current cell could be hidden because the user has scrolled the control through the scrollbar. If you want to ensure the visibility of the current cell, use this method without parameters.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [CellVisible Method](#cellvisible-method-gridex-control), [FirstItem Property](Properties.md#firstitem-property-gridex-control), [LeftCol Property](Properties.md#leftcol-property-gridex-control)

## ExpandAll Method (GridEX Control)

Expands all group rows in **GridEX** control.

### Syntax

 *object*.**ExpandAll**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 This method is ignored if the **GridEX** control is not grouped.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ExpandSelection Method](#expandselection-method-gridex-control), [RowExpanded Property](Properties.md#rowexpanded-property-gridex-control), [CollapseAll Method](#collapseall-method-gridex-control)  
**Example:** [CollapseAll Example](../Examples.md#collapseall-example)

## ExpandSelection Method (GridEX Control)

Expands and selects all the children records of the currently selected group rows in **GridEX** control.

### Syntax

 *object*.**ExpandSelection**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 In order to use this method, the **MultiSelect** property must be set to **True**. Otherwise an error will occur.  
 This method is useful when you are about to begin an OLE Drag operation and want to select all the child rows of the selected group rows.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ExpandAll Method](#expandall-method-gridex-control), [CollapseAll Method](#collapseall-method-gridex-control)  
**Example:** [RowDrag Example](../Examples.md#rowdrag-example)

## Find Method (GridEX Control)

Sets the current row in a **GridEX** control to the one that matches the specified criteria.

### Syntax

 *object*.**Find** ( *colindex, operator, value1, value2, start, searchdirection* )  
 The **Find** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *colindex* | Required.An integer that represents the index of the column. |
| *operator* | Required. A value or constant that specifies the operator used in the comparison operation. |
| *value1* | Required. A variant that represents the value to be compared with the column values. |
| *value2* | Optional. A variant that represents the value that is to be compared with the column values. |
| *start* | Optional. A long that represents the row position to use as the starting position for the search. |
| *searchdirection* | Optional. A value or constant that specifies wich direction to search, as described in settings. |

### Settings

 the Settings for **Find** method are:

| Setting | Description |
| --- | --- |
| **True** | A record matching the specified criteria was found. |
| **False** | No record was found matching the specified criteria. |

The settings for *operator* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexEqual** | 0 | Finds a record where the value of the column is equal to the *Value1* parameter. |
|  **jgexNotEqual** | 1 | Finds a record where the value of the column is different to the *Value1* parameter. |
|  **jgexGreaterThan** | 2 | Finds a record where the value of the column is greater than the *Value1* parameter. |
|  **jgexLessThan** | 3 | Finds a record where the value of the column is less than the *Value1* parameter. |
|  **jgexGreaterThanOrEqualTo** | 4 | Finds a record where the value of the column is greater than or equal to the *Value1* parameter. |
|  **jgexLessThanOrEqualTo** | 5 | Finds a record where the value of the column is less than or equal to the *Value1* parameter. |
|  **jgexBetween** | 6 | Finds a record where the value of the column is greater than or equal to the *Value1* parameterand less than or equal to *Value2* parameter. |
|  **jgexNotBetween** | 7 | Finds a record where the value of the column is less than the *Value1* parameter and greater than *Value2* parameter. |
|  **jgexContains** | 8 | Finds a record where the value of the column contains *Value1* parameter. |
|  **jgexNotContains** | 9 | Finds a record where the value of the column does not contain *Value1* parameter. |

The settings for searchdirection are:

| Setting | Description |
| --- | --- |
| **1** | Default. Forward. The GridEX control searchs from the start row position to the last row in the control. |
| **-1** | Backwards. The GridEX control searchs from the start row position to the first row in the control. |

### Remarks

 If *start* parameter is ommited, the search begins from the first row in the control.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [MoveToBookmark Method](#movetobookmark-method-gridex-control), [MoveToRowIndex Method](#movetorowindex-method-gridex-control)  
**Example:** [Find Example](../Examples.md#find-example)

## GetClipString Method (GridEX Control)

Returns the contents of the selected rows in a **GridEX** control.

### Syntax

 *object*.**GetClipString** (*includeheaders* **As Boolean**, *columndelimeter* **As String**, *rowdelimeter* **As String**)  
 The **GetClipString** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *includeheaders* | Optional. A Boolean expression that determines if the string returned will include the captions of the column headers. |
| *columndelimeter* | Optional. Delimiter used between columns if specified, otherwise the TAB character. |
| *rowdelimeter* | Optional. Delimiter used between rows if specified, otherwise the CARRIAGE RETURN character. |

### Remarks

 This method can be used to return a string with the values of the selected rows. In an OLE Drag operation you can provide the string returned by this method to the **JSDataObject**.

### Data Type

 String

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ExpandSelection Method](#expandselection-method-gridex-control), [SetData Method](Objects/JSDataObject-Object.md#setdata-method-jsdataobject-object)  
**Example:** [RowDrag Example](../Examples.md#rowdrag-example)

## GetRowData Method (GridEX Control)

Returns a **JSRowData** object representing a row in a **GridEX** control.

### Syntax

 *object*.**GetRowData** *rowposition*  
 The **GetRowData** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *rowposition* | A long expresion representing the position of a row in a GridEX control. |

### Remarks

 Use this method to retrieve a **JSRowData** object representing a row to change row and cell formats or to get values from a specific row.  
 A **JSRowData** object may become outdated as the result of scrolling or refreshing the **GridEX** contents. When you try to change some of the properties of an outdated JSRowData an error occurs. To avoid working with outdated **JSRowData** objects try not to keep them between procedures.

### Data Type

 **JSRowData**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RowFormat Event](Events.md#rowformat-event-gridex-control), [JSRowData Object](Objects/JSRowData-Object.md#jsrowdata-object)  
**Example:** [GetRowData Example](../Examples.md#getrowdata-example)

## GroupFromPoint Method (GridEX Control)

Returns the **JSGroup** object whose position matches the specified coordinates.

### Syntax

 *object*.**GroupFromPoint** *x, y*  
 The **GroupFromPoint** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *x,y* | Coordinates of a target group. |

### Remarks

 If the point specified in the *x* and *y* parameters is not contained in any group the return value is Nothing.

### Data Type

 **JSGroup**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [HitTest Method](#hittest-method-gridex-control), [RowFromPoint Method](#rowfrompoint-method-gridex-control), [ColFromPoint Method](#colfrompoint-method-gridex-control)  
**Example:** [ColFromPoint Example](../Examples.md#colfrompoint-example)

## GroupRowLevel Method (GridEX Control)

Returns the group level of a row if it is a group row or 0 otherwise.

### Syntax

 *object*.**GroupRowLevel** (*rowposition* **As Long**)  
 The **GroupRowLevel** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *rowposition*  | A long expression that represents the position of the row. |

### Remarks

 When a **GridEX** control is grouped, some rows are group rows and do not represent a record.  
 You may use this method to find out the level of any row. If the row is a group this method returns a number ranging from 1 to 4. If the row is not a group, this methods returns zero.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [IsGroupItem Method](#isgroupitem-method-gridex-control), [RowExpanded Property](Properties.md#rowexpanded-property-gridex-control), [GroupLevel Property](Objects/JSRowData-Object.md#grouplevel-property-jsrowdata-object)  
**Example:** [RowExpanded Example](../Examples.md#rowexpanded-example)

## HitTest Method (GridEX Control)

Returns a value that represents the part of a **GridEX** control whose position matches the specified coordinates.

### Syntax

 *object*.**HitTest** *x, y*  
 The **HitTest** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *x,y* | Coordinates to search. |

The possible return values are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexHTNoWhere** | 0 | The point is outside the client area of the **GridEX** control. |
|  **jgexHTGroupByBox** | 1 | The point is in the “group by” box. |
|  **jgexHTColumnHeader** | 2 | The point is in the column header. |
|  **jgexHTRowHeader** | 3 | The point is in the row header. |
|  **jgexHTNewRow** | 4 | The point is in the top new row. |
|  **jgexHTCell** | 5 | The point is in a cell. |
|  **jgexHTCard** | 6 | The point is in card. |
|  **jgexHTBackGround** | 7 | The point is in the background. |

### Remarks

 The **HitTest** method provides a way to determine the location of a point inside the **GridEX** control.  
 You can also use the **ColFromPoint**, **GroupFromPoint** and **RowFromPoint** methods when you need more detailed information about objects in a given position.

### Data Type

 **jgexHitTestConstants**

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RowFromPoint Method](#rowfrompoint-method-gridex-control), [GroupFromPoint Method](#groupfrompoint-method-gridex-control), [ColFromPoint Method](#colfrompoint-method-gridex-control)  
**Example:** [ColFromPoint Example](../Examples.md#colfrompoint-example)

## HoldFields Method (GridEX Control)

Prevents recalculation of the column layout when the **Recordset** is created, refreshed or re-created.

### Syntax

 *object*.**HoldFields**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The **HoldFields** method sets the current column/field layout as the customized layout so that subsequent **Rebind** operations will use the current layout for display.  
 You can resume the **GridEX** control's automatic layout behavior by invoking the **ClearFields** method.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [HoldSortSettings Property](Properties.md#holdsortsettings-property-gridex-control), [ClearFields Method](#clearfields-method-gridex-control)  
**Example:** [HoldFields Example](../Examples.md#holdfields-example)

## IsGroupItem Method (GridEX Control)

Returns **True** if the specified row is a group row in a **GridEX** control.

### Syntax

 *object*.**IsGroupItem**(*row*)  
 The **IsGroupItem** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *row* | A long expression that represents the position of the row being tested. |

### Settings

 The settings for **IsGroupItem** are:

| Setting | Description |
| --- | --- |
| **True** | The row specified in the row parameter is a group row. |
| **False** | The row specified in the row parameter is not a group row. |

### Remarks

 When a **GridEX** control is grouped, some rows are group rows and do not represent a record.  
 You may use this method to find out whether the current row is a group or a record row.

**Note** If the control is not grouped, this method always returns **False**.

### Data Type

 Boolean

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RowExpanded Property](Properties.md#rowexpanded-property-gridex-control), [GroupRowLevel Method](#grouprowlevel-method-gridex-control), [RowType Property](Objects/JSRowData-Object.md#rowtype-property-jsrowdata-object), [RowType Property](Objects/JSSelectedItem-Object.md#rowtype-property-jsselecteditem-object)  
**Example:** [RowExpanded Example](../Examples.md#rowexpanded-example)

## LayoutString Method (GridEX Control)

Returns a string that contains the **GridEX** control layout settings.

### Syntax

 *object*.**LayoutString** ( *unicodecompression* **As Boolean** )  
 The **LayoutString** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *unicodecompression* | Optional. A Boolean expression that determines whether the string returned is expanded to prevent errors with Unicode compression, as described in Settings. |

The settings for *unicodecompression* are:

| Setting | Description |
| --- | --- |
| **True** | The result is a string with binary data that should not be compressed when saved to disk, in order to work with the **LoadLayoutString** method. |
| **False** | Default. The result is a string with binary data that can be compressed when saved to disk. |

### Remarks

 You can use this method to get a string with the layout settings in a **GridEX** control to be saved in a database or an ini file.  
 Once you have a layout string retrieved by this method, you can use the **LoadLayoutString** method to change the layout at runtime.

**Note**: If you are going to use **LayoutString** and **LoadLayoutString** methods in languages that require DBCS (double-byte character sets) unicodecompression must be **True**.

### Data Type

 String

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [SaveLayout Method](#savelayout-method-gridex-control), [LoadLayoutString Method](#loadlayoutstring-method-gridex-control), [LoadLayout Method](#loadlayout-method-gridex-control), [LoadLayoutFromURL Method](#loadlayoutfromurl-method-gridex-control)  
**Example:** [LayoutString Example](../Examples.md#layoutstring-example)

## LoadEntireRecordset Method (GridEX Control)

Loads the bookmarks of all records in the underlying **Recordset** of a **GridEX** control.

### Syntax

 *object*.**LoadEntireRecordset**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The **LoadEntireRecordset** method loads all the records in the underlying **Recordset**.  
 This method is only used for DAO and ADO modes, and is ignored when the control is in unbound mode. However, invoking this method while the **GridEX** control is in unbound mode will not generate an error.  
 The **GridEX** control usually loads the **Recordset** as needed, so this method ensures that all the records has been retrieved.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [FullyLoaded Property](Properties.md#fullyloaded-property-gridex-control), [SearchNewRecords Method](#searchnewrecords-method-gridex-control)  
**Example:** [RowSelected Example](../Examples.md#rowselected-example)

## LoadLayout Method (GridEX Control)

Loads a previously saved **GridEX** layout from a file.

### Syntax

 *object*.**LoadLayout** *filename*, *rebind* , *databasename*  
 The **LoadLayout** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *filename* | The path of the file containing **GridEX** layout binary data. |
| *rebind* | Optional.A Boolean expressions that determines whether the **GridEX** control rebinds itself using the data source properties saved in the layout. |
| *databasename* | Optional. A string expression that represents the value of the **DatabaseName** property to be used when rebinding instead of the value saved with the layout. |

### Settings

 The settings for *rebind* are

| Setting | Description |
| --- | --- |
| **True** | Default. the underlying recordset will be re-opened based on the data mode settings saved with the layout. |
| **False** | The current recordset bounded to the control is used. Data mode settings are ignored. |

### Remarks

 You can use this method to easily change previously saved layouts at runtime.  
 The file specified in the *filename* parameter must be previously saved using the **SaveLayout** method.  
 When rebinding, you can set the appropriate **DatabaseName** property value to be used instead of the one saved with the layout.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [SaveLayout Method](#savelayout-method-gridex-control), [LoadLayoutString Method](#loadlayoutstring-method-gridex-control), [LoadLayoutFromURL Method](#loadlayoutfromurl-method-gridex-control), [LayoutString Method](#layoutstring-method-gridex-control)  
**Example:** [SaveLayout Example](../Examples.md#savelayout-example)

## LoadLayoutFromURL Method (GridEX Control)

Loads a previously saved **GridEX** layout from an URL.

### Syntax

 *object*.**LoadLayoutFromURL** *url, rebind , databasename*  
 The **LoadLayoutFromURL** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *url* | Absolute URL specifying the path of the layout file to be loaded |
| *rebind* | Optional. A Boolean expressions that determines whether the **GridEX** control rebinds itself using the data source properties saved in the layout. |
| *databasename* | Optional. A string expression that represents the value of the **DatabaseName** property to be used when rebinding instead of the value saved with the layout. |

### Settings

 The settings for *rebind* are:

| Setting | Description |
| --- | --- |
| **True** | Default. the underlying recordset will be re-opened based on the data mode settings saved with the layout. |
| **False** | The current recordset bounded to the control is used. Data mode settings are ignored. |

### Remarks

 You can use this method to easily change previously saved layouts at runtime.  
 The file specified in the *url* parameter must be previously saved using the **SaveLayout**method.  
 When rebinding, you can set the appropriate **DatabaseName** property value to be used instead of the one saved with the layout.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [SaveLayout Method](#savelayout-method-gridex-control), [LoadLayoutString Method](#loadlayoutstring-method-gridex-control), [LoadLayout Method](#loadlayout-method-gridex-control), [LayoutString Method](#layoutstring-method-gridex-control)

## LoadLayoutString Method (GridEX Control)

Loads a previously saved **GridEX** layout from a layout string.

### Syntax

 *object*.**LoadLayoutString** *layoutstring, rebind , databasename*  
 The **LoadLayoutString** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *layoutstring* | A string that contains layout data retrieved from a previous call to the **LayoutString** method. |
| *rebind* | Optional. A Boolean expressions that determines whether the **GridEX** control rebinds itself using the data source properties saved in the layout. |
| *databasename* | Optional. A string expression that represents the value of the **DatabaseName** property to be used when rebinding instead of the value saved with the layout. |

### Settings

 The settings for *rebind* are:

| Setting | Description |
| --- | --- |
| **True** | Default. the underlying recordset will be re-opened based on the data mode settings saved with the layout. |
| **False** | The current recordset bounded to the control is used. Data mode settings are ignored. |

### Remarks

 You can use this method to easily change previously saved layouts at runtime.  
 The *layoutstring* parameter must be a previously saved layout string retrieved from the **LayoutString** method.  
 When rebinding, you can set the appropriate **DatabaseName** property value to be used instead of the one saved with the layout.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [SaveLayout Method](#savelayout-method-gridex-control), [LoadLayout Method](#loadlayout-method-gridex-control), [LoadLayoutFromURL Method](#loadlayoutfromurl-method-gridex-control), [LayoutString Method](#layoutstring-method-gridex-control)  
**Example:** [LayoutString Example](../Examples.md#layoutstring-example)

## MoveFirst Method (GridEX Control)

Sets the first accessible record in a **GridEX** control as the current row.

### Syntax

 *object*.**MoveFirst**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 If a **GridEX** control is grouped, only the open (visible) rows that represent records in the **Recordset** are considered.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [MoveLast Method](#movelast-method-gridex-control), [MovePrevious Method](#moveprevious-method-gridex-control), [MoveRelative Method](#moverelative-method-gridex-control), [MoveToBookmark Method](#movetobookmark-method-gridex-control), [MoveToRowIndex Method](#movetorowindex-method-gridex-control)

## MoveLast Method (GridEX Control)

Sets the last accessible record in a **GridEX** control as the current row.

### Syntax

 *object*.**MoveLast**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 If a **GridEX** control is grouped, only the open (visible) rows that represent records in the **Recordset** are considered.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [MoveFirst Method](#movefirst-method-gridex-control), [MovePrevious Method](#moveprevious-method-gridex-control), [MoveRelative Method](#moverelative-method-gridex-control), [MoveToBookmark Method](#movetobookmark-method-gridex-control), [MoveToRowIndex Method](#movetorowindex-method-gridex-control)

## MoveNext Method (GridEX Control)

Moves to the next accessible record in a **GridEX** control.

### Syntax

 *object*.**MoveNext**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 If a **GridEX** control is grouped, only the open (visible) rows that represent a record in the Recordset are considered.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [MoveFirst Method](#movefirst-method-gridex-control), [MoveLast Method](#movelast-method-gridex-control), [MovePrevious Method](#moveprevious-method-gridex-control), [MoveRelative Method](#moverelative-method-gridex-control), [MoveToBookmark Method](#movetobookmark-method-gridex-control), [MoveToRowIndex Method](#movetorowindex-method-gridex-control)

## MovePrevious Method (GridEX Control)

Moves to the previous accessible record in a **GridEX** control.

### Syntax

 *object*.**MovePrevious**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 If a **GridEX** control is grouped, only the open (visible) rows that represent a record in the **Recordset** are considered.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [MoveFirst Method](#movefirst-method-gridex-control), [MoveLast Method](#movelast-method-gridex-control), [MoveRelative Method](#moverelative-method-gridex-control), [MoveToBookmark Method](#movetobookmark-method-gridex-control), [MoveToRowIndex Method](#movetorowindex-method-gridex-control)

## MoveRelative Method (GridEX Control)

Moves the current row to the specified position.

### Syntax

 *object*.**MoveRelative** *nRows*  
 The **MoveRelative** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *nRows* | A long expression specifying the number of rows the position will move. <br> <br> If *nRows* is greater than 0, the position is moved forward (toward the last row).<br> <br> If *nRows* is less than 0, the position is moved backward (toward the first row). |

### Remarks

 If a **GridEX** control is grouped, only the open (visible) rows that represent a record in the **Recordset** are considered.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [MoveFirst Method](#movefirst-method-gridex-control), [MoveLast Method](#movelast-method-gridex-control), [MovePrevious Method](#moveprevious-method-gridex-control), [MoveToBookmark Method](#movetobookmark-method-gridex-control), [MoveToRowIndex Method](#movetorowindex-method-gridex-control)

## MoveToBookmark Method (GridEX Control)

Sets the current row in a **GridEX** control to the one that matches the specified bookmark.

### Syntax

 *object*.**MoveToBookmark** *vbookmark*  
 The **MoveToBookmark** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *vbookmark* | A Variant value identifying a bookmark. |

### Remarks

 If the row with the bookmark specified is hidden because its group is closed, the **GridEX** control will open the necessary group rows in order to ensure the visibility of the row.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [MoveFirst Method](#movefirst-method-gridex-control), [MoveLast Method](#movelast-method-gridex-control), [MovePrevious Method](#moveprevious-method-gridex-control), [MoveRelative Method](#moverelative-method-gridex-control), [Find Method](#find-method-gridex-control), [MoveToRowIndex Method](#movetorowindex-method-gridex-control)  
**Example:** [MoveToBookmark Example](../Examples.md#movetobookmark-example)

## MoveToRowIndex Method (GridEX Control)

Sets the current row in a **GridEX** control to the one that matches the index.

### Syntax

 *object*.**MoveToRowIndex** *rowindex*  
 The **MoveToRowIndex** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *rowindex* | A Long expression identifying a row index. |

### Remarks

 If the specified row is hidden because its group is closed, the **GridEX** control will expand the parent group in order to ensure the visibility of this row.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [MoveFirst Method](#movefirst-method-gridex-control), [MoveLast Method](#movelast-method-gridex-control), [MovePrevious Method](#moveprevious-method-gridex-control), [MoveRelative Method](#moverelative-method-gridex-control), [MoveToBookmark Method](#movetobookmark-method-gridex-control), [Find Method](#find-method-gridex-control)

## OLEDrag Method (GridEX Control)

Initiates an OLE drag/drop operation.

### Syntax

 *object*.**OLEDrag**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 When the **OLEDrag** method is called, the **GridEX**’s **OLEStartDrag** event occurs, allowing it to supply data to a target component.

**Note** It is recommended to use this method in the **RowDrag** event.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [OLECompleteDrag Event](Events.md#olecompletedrag-event-gridex-control), [OLEDragDrop Event](Events.md#oledragdrop-event-gridex-control), [OLEDragOver Event](Events.md#oledragover-event-gridex-control), [OLEDropMode Property](Properties.md#oledropmode-property-gridex-control), [OLEGiveFeedback Event](Events.md#olegivefeedback-event-gridex-control), [OLESetData Event](Events.md#olesetdata-event-gridex-control), [OLEStartDrag Event](Events.md#olestartdrag-event-gridex-control)  
**Example:** [RowDrag Example](../Examples.md#rowdrag-example)

## PrintGrid Method (GridEX Control)

Prints the contents of a **GridEX** control.

### Syntax

 *object*.**PrintGrid** *useprintsetupdlg, printselecteditems*  
 The **PrintGrid** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *useprintsetupdlg* | A Boolean expression that determines whether a print setup dialog is used to enable the user to specify the attributes of a printed page. |
| *printselecteditems* | A Boolean expressions that determines whether only the selected records are printed or all the records in the control. |

### Settings

 The settings for *useprintsetupdlg* are:

| Setting | Description |
| --- | --- |
| **True** | A page setup dialog is displayed before the **GridEX** control's contents is printed. |
| **False** | Default. **GridEX** control's contents is printed using the settings in the **JSPrinterProperties** object. |

The settings for *printselecteditems* are:

| Setting | Description |
| --- | --- |
| **True** | Only selected items in a **GridEX** control are printed |
| **False** | Default. All Items in the **GridEX** control are printed |

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BeforePrintPage Event](Events.md#beforeprintpage-event-gridex-control), [PrintingProgress Event](Events.md#printingprogress-event-gridex-control), [PrinterProperties Property](Properties.md#printerproperties-property-gridex-control), [PrintGrid Method](#printgrid-method-gridex-control), [JSPrinterProperties](Objects/JSPrinterProperties-Object.md#jsprinterproperties), [BeforePrinting Event](Events.md#beforeprinting-event-gridex-control)  
**Example:** [PrinterProperties Example](../Examples.md#printerproperties-example)

## PrintPreview Method (GridEX Control)

Sends the contents of a **GridEX** control to be printed to a **GEXPreviewControl**.

### Syntax

 *object*.**PrintPreview** *previewcontrol, printselecteditems*  
 The **PrintPreview** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *previewcontrol* | A GEXPreview object to be used as the print preview for the control. |
| *printselecteditems* | A Boolean expressions that determines whether only the selected records are printed or all the records in the control. |

### Settings

 The settings for *printselecteditems* are:

| Setting | Description |
| --- | --- |
| **True** | Only selected items in a GridEX control are printed. |
| **False** | Default. All Items in the GridEX control are printed. |

### Remarks

 Use this method to enable the users to preview the data to be printed.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [PrintGrid Method](#printgrid-method-gridex-control), [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control), [PrinterProperties Property](Properties.md#printerproperties-property-gridex-control)  
**Example:** [PrintPreview Example](../Examples.md#printpreview-example)

## Rebind Method (GridEX Control)

Forces re-creation of the **Recordset** in a **GridEX** control.

### Syntax

 *object*.**Rebind** *holdsortsettings*  
 The **Rebind** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *holdsortsettings* | A boolean expression that determines whether the current group and sort settings are preserved. |

The settings for *holdsortsettings* are:

| Setting | Description |
| --- | --- |
| **True** | Groups and SortKeys are preserved. |
| **False** | Default. Groups and SortKeys are cleared. |

### Remarks

 The Rebind method causes the **GridEX** control to close the current **Recordset** and create a new one with the parameters specified in the **DatabaseName**, **Exclusive**, **LockType**, **ReadOnly**, **RecordsetType** and **Connect** properties.  
 If you have not modified the grid columns at design time, then executing the **Rebind** method will reset the columns, headings, and other properties based on the current data source.  
 However, if you have altered the columns in any way at design time, then the **GridEX** will assume that you wish to maintain the modified grid layout and will not automatically reset the columns.  
 The **Rebind** method also clears all groups and sort keys in a **GridEX** control, unless holdsortsettings parameter is **True**.  
 If the columns are reset, it also clears the formatting conditions.

**Notes** To force the **GridEX** to reset the column bindings even if the columns were modified at design time, invoke the **ClearFields** method immediately before **Rebind**.

 In order to cancel the automatic layout response of the control and force the grid to use the current column/field layout, invoke the **HoldFields** method immediately before **Rebind**.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [HoldFields Method](#holdfields-method-gridex-control), [ADORecordset Property](Properties.md#adorecordset-property-gridex-control), [HoldSortSettings Property](Properties.md#holdsortsettings-property-gridex-control), [ClearFields Method](#clearfields-method-gridex-control), [Recordset Property](Properties.md#recordset-property-gridex-control)  
**Example:** [RebindSample Example](../Examples.md#rebindsample-example)

## Refetch Method (GridEX Control)

Forces a **GridEX** control to refresh its contents without re-opening the recordset.

### Syntax

 *object*.**Refetch** *holdsortsettings*  
 The **Refetch** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *holdsortsettings* | A boolean expression that determines whether group and sort settings currently in the grid are preserved. |

The settings for *holdsortsettings* are:

| Setting | Description |
| --- | --- |
| **True** | Groups and sort keys are preserved. |
| **False** | Default. Groups and sort keys are cleared. |

### Remarks

 Use this method when several changes have been made to the underlying recordset data forcing the control to refresh the data displayed in the records and recalculate the groups and sort positions.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RefreshRowBookmark Method](#refreshrowbookmark-method-gridex-control), [RefreshRowIndex Method](#refreshrowindex-method-gridex-control), [Refresh Method](#refresh-method-gridex-control), [Rebind Method](#rebind-method-gridex-control)

## Refresh Method (GridEX Control)

Forces a complete repaint of a **GridEX** control.

### Syntax

 *object*.**Refresh**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 Generally, the painting of the control is automatically handled while no events occur. However, there may be situations where you want to update the control immediately.  
 The **Refresh** method does not reopen or refresh the underlying **Recordset** object; to do this you must use the **Rebind** method.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Refetch Method](#refetch-method-gridex-control), [RefreshRowBookmark Method](#refreshrowbookmark-method-gridex-control), [RefreshRowIndex Method](#refreshrowindex-method-gridex-control), [Redraw Property](Properties.md#redraw-property-gridex-control)

## RefreshGroups Method (GridEX Control)

Forces recalculation of groups in a **GridEX** control.

### Syntax

 *object*.**RefreshGroups** *allcollapsed*  
 The **RefreshGroups** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *allcollapsed* | Optional. A Boolean expression that determines if groups must be collapsed or expanded as described in settings. The default value is **False**. |

### Settings

 The settings for *allcollapsed* are:

| Setting | Description |
| --- | --- |
| **True** | The group rows will be collapsed. |
| **False** | The group rows will be expanded. |

### Remarks

 Changes to the data in the underlying **Recordset** made by other users may affect the consistency of the current groups. However, the **GridEX** cannot know when this happens.  
 In order to force recalculation of groups with the current group settings you may use the **RefreshGroups** method. An example of this is as follows:

```vb
Private Sub mnuRefreshGroups_Click()
If GridEX1.Groups.Count > 0 then
GridEX1.RefreshGroups
End if
End Sub
```

Use the **RefreshGroups** method when there is a reason to believe that other processes have changed data in the underlying **Recordset**, possibly invalidating the current group configuration.  
 If you try to use this method while the **GridEX** control is in card view, an error occurs.

**Note** The **RefreshGroups** method also refreshes the sort order of groups in the **GridEX** control.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RefreshSort Method](#refreshsort-method-gridex-control)  
**Example:** [Groups Example](../Examples.md#groups-example)

## RefreshRowBookmark Method (GridEX Control)

Refreshes the data of the record that matches the specified bookmark. If needed, it also changes the position of the record being refreshed.

### Syntax

 *object*.**RefreshRowBookmark** *bookmark*  
 The **RefreshRowBookmark** method syntax has these parts:

| Part | Description |
| --- | --- |
| object | An object expression that evaluates to an object in the Applies To list. |
| *bookmark* | A Variant value identifying a bookmark. |

### Remarks

 Use this method when you want to update data in a record.  
 This method retrieves data from the **Recordset** and, if the refreshed data does not concur with the current grouping and/or sorting criteria, the control will recalculate the group and/or sort position for the record, and move it to another group if necessary.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Refetch Method](#refetch-method-gridex-control), [RefreshRowIndex Method](#refreshrowindex-method-gridex-control), [Refresh Method](#refresh-method-gridex-control)  
**Example:** [Redraw Example](../Examples.md#redraw-example)

## RefreshRowIndex Method (GridEX Control)

Refreshes the data of the record that matches the specified Index. If needed, it also changes the position of the record being refreshed.

### Syntax

 *object*.**RefreshRowIndex** *rowindex*  
 The **RefreshRowIndex** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *rowindex* | A Long expression identifying a row index. |

### Remarks

 Use this method when you want to update data in a record.  
 This method retrieves data from the **Recordset** and, if the refreshed data does not concur with the current grouping and/or sorting criteria, the control will recalculate the group and/or sort position for the record, and move it to another group if necessary.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Refetch Method](#refetch-method-gridex-control), [RefreshRowBookmark Method](#refreshrowbookmark-method-gridex-control), [Refresh Method](#refresh-method-gridex-control)  
**Example:** [RefreshRowIndex Example](../Examples.md#refreshrowindex-example)

## RefreshSort Method (GridEX Control)

Forces a **GridEX** control to re-sort the records.

### Syntax

 *object*.**RefreshSort**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 Changes to the data in the underlying Recordset made by other users may affect the continuity of the current sort order. However, the **GridEX** cannot know when this happens. In order to ensure the validity of the current sort settings you may use the **RefreshSort** method.  
 An example of this is as follows:

```vb
Private Sub mnuRefreshSort_Click()
If GridEX1.SortKeys.Count>0 then
GridEX1.RefreshSort
End if
End sub
```

Use the **RefreshSort** method when there is a reason to believe that other users have changed data in the underlying **Recordset**, therefore invalidating the current sort order.

**Note** The **RefreshSort** method doesn’t refresh groups in the **GridEX** control.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RefreshGroups Method](#refreshgroups-method-gridex-control)  
**Example:** [SortKeys Example](../Examples.md#sortkeys-example)

## RowFromPoint Method (GridEX Control)

Returns the row whose position matches the specified coordinates.

### Syntax

 *object*.**RowFromPoint** *x, y*  
 The **RowFromPoint** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *x,y* | Coordinates of a target row. |

Remarks  
 If the point specified in the x and y parameters is not contained in any row the return value is 0.  
 Hidden rows (under collapsed groups) are not considered when calculating the target row.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [HitTest Method](#hittest-method-gridex-control), [GroupFromPoint Method](#groupfrompoint-method-gridex-control), [ColFromPoint Method](#colfrompoint-method-gridex-control)  
**Example:** [ColFromPoint Example](../Examples.md#colfrompoint-example)

## RowIndex Method (GridEX Control)

Returns a value containing the original index of the selected row in a **GridEX** control.

### Syntax

 *object*.**RowIndex** ( *row* )

The **RowIndex** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *row* | A long expression that represents the position of the row being tested. |

### Remarks

 The **RowIndex** method returns the original index of a row.  
 If the selected row is a “group row” the **RowIndex** property will return 0.

### Data Type

 Long

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RowIndex Property](Objects/JSRowData-Object.md#rowindex-property-jsrowdata-object), [RowIndex Property](Objects/JSSelectedItem-Object.md#rowindex-property-jsselecteditem-object), [RowBookmark Property](Properties.md#rowbookmark-property-gridex-control), [Row Property](Properties.md#row-property-gridex-control)  
**Example:** [RowIndex Example](../Examples.md#rowindex-example)

## SaveLayout Method (GridEX Control)

Saves the current **GridEX** layout in a file.

### Syntax

 *object*.**SaveLayout** *filename*  
 The **SaveLayout** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *filename* | The path of the file in wich the current **GridEX** layout is saved. |

### Remarks

 You can use this method to save the current layout in a **GridEX** control.  
 Saving a layout, you can preserve user setting such as groups, sort keys, column positions and column widths and restore those settings when your application is restarted.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [LoadLayoutString Method](#loadlayoutstring-method-gridex-control), [LoadLayout Method](#loadlayout-method-gridex-control), [LoadLayoutFromURL Method](#loadlayoutfromurl-method-gridex-control), [LayoutString Method](#layoutstring-method-gridex-control)  
**Example:** [SaveLayout Example](../Examples.md#savelayout-example)

## SearchNewRecords Method (GridEX Control)

Searches for records that may have been added after a grouping or sorting operation.

### Syntax

 *object*.**SearchNewRecords**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 Use this method whenever you add records to **Recordset** clone of a **GridEX** control, and want those records to be displayed in the control.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [FullyLoaded Property](Properties.md#fullyloaded-property-gridex-control), [LoadEntireRecordset Method](#loadentirerecordset-method-gridex-control)  
**Example:** [SearchNewRecords Example](../Examples.md#searchnewrecords-example)

## SetFocusRecordNavigator Method (GridEX Control)

Set focus to the specific record box in the record navigator.

### Syntax

 *object*.**SetFocusRecordNavigator**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 This method is ignored if **RecordNavigator** property is **False**.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RecordNavigator Property](Properties.md#recordnavigator-property-gridex-control)

## Update Method (GridEX Control)

Commits changes made to the current row, writes to the database and re-positions the record if necessary.

### Syntax

 *object*.**Update**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The **Update** method also determines if the current row position is still valid.  
 If the current position is not longer valid and the **AutomaticArrange** property is set to **True**, the **GridEX** control will reposition the row. This may be necessary if changes to the record where made by other users or to the **Recordset** clone.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [DataChanged Property](Properties.md#datachanged-property-gridex-control), [RefreshRowBookmark Method](#refreshrowbookmark-method-gridex-control), [RefreshRowIndex Method](#refreshrowindex-method-gridex-control)  
**Example:** [DataChanged Example](../Examples.md#datachanged-example)
