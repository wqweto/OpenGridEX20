# GridEX Control — Events

## AfterColEdit Event (GridEX Control)

Occurs just before the user enters edit mode by typing a character or clicking in a cell.

### Syntax

 **Private Sub** *object*_**BeforeColEdit**([ *index* **As Integer**,] **ByVal** *colindex* **As Integer**, **ByVal** *cancel* **As JSRetBoolean**)  
 The **BeforeColEdit** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *Index* | An integer that identifies a control if it is in a control array. |
| *colindex* | An integer that identifies the index of the column to be edited. |
| *cancel* | A **JSRetBoolean** object that may be set to True to prevent the user from editing the cell, as described in Settings. |

### Settings

 The settings for cancel are:

| Setting | Description |
| --- | --- |
| **True** | The cell will not enter edit mode. |
| **False** | (Default) The cell will enter edit mode and the edit window will appear. |

### Remarks

 You can use this event to control which cell can be edited, on a per-cell basis. If you want to prevent edition on a per-column basis set the **EditType** property of the **JSColumn** to **jgexEditNone**.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AfterColUpdate Event](#aftercolupdate-event-gridex-control), [AfterUpdate Event](#afterupdate-event-gridex-control), [BeforeColEdit Event](#beforecoledit-event-gridex-control), [BeforeColUpdate Event](#beforecolupdate-event-gridex-control), [BeforeUpdate Event](#beforeupdate-event-gridex-control)

## AfterColMove Event (GridEX Control)

Occurs after the user has dropped a column into a new position.

### Syntax

 **Private Sub** *object*_**AfterColMove** ([ *index* **As Integer** ])  
 The **AfterColMove** event syntax has these parts:

| Part | Description |
| --- | --- |
| object | An object expression that evaluates to an object in the Applies To list. |
| index | An integer that identifies a control if it is in a control array. |

### Remarks

 The event is triggered after the position of a column has changed in response to a user action.  
 The event is preceded by the **BeforeColMove** event and will be fired only if this event is not canceled.  
 This event will be fired whether the column dragged was originally in the header row or it was placed in the “group by” box.  
 The **AfterColMove** event is fired only if the **View** property is set to **jgexTable**.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AfterGroupChange Event](#aftergroupchange-event-gridex-control), [BeforeColumnDrag Event](#beforecolumndrag-event-gridex-control), [BeforeGroupChange Event](#beforegroupchange-event-gridex-control), [BeforeGroupDrag Event](#beforegroupdrag-event-gridex-control), [BeforeColMove Event](#beforecolmove-event-gridex-control)

## AfterColUpdate Event (GridEX Control)

Occurs after data is moved from a cell in the **GridEX** control to the control's copy buffer.

### Syntax

 **Private Sub** *object*_**AfterColUpdate**([ *index* **As Integer**, ] **ByVal** *colindex* **As Integer**)  
 The **AfterColUpdate** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *colindex* | An integer that identifies the index of the column that was updated. |

### Remarks

 When a user completes an editing operation in a cell of the **GridEX** control, the **BeforeColUpdate** event is executed, and unless canceled, data from the cell is moved to the control's copy buffer. Once the data is moved, the **AfterColUpdate** event is executed.  
 This event is fired only if the cancel argument of the **BeforeColUpdate** event is not set to **True**.  
 Once the **AfterColUpdate** event procedure is invoked, the cell data has already been moved to the control's copy buffer and the operation cannot be canceled. However, other updates can occur before the data is committed to the **Recordset**.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AfterUpdate Event](#afterupdate-event-gridex-control), [BeforeColEdit Event](#beforecoledit-event-gridex-control), [BeforeColUpdate Event](#beforecolupdate-event-gridex-control), [BeforeUpdate Event](#beforeupdate-event-gridex-control), [AfterColEdit Event](#aftercoledit-event-gridex-control)

## AfterDelete Event (GridEX Control)

Occurs after the user deletes the selected records in the **GridEX** control.

### Syntax

 **Private Sub** *object*_**AfterDelete**([ *index* **As Integer** ])  
 The **AfterDelete** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |

### Remarks

 When the user selects one or more rows in the **GridEX** control and presses DEL, these rows are deleted. Before the delete operation begins, the **BeforeDelete** event is fired and, unless canceled, the **GridEX** will try to delete the selected records in the **Recordset**. Once the delete operation ends, the **AfterDelete** event is triggered.  
 **Note**: To select the entire row if the **AllowEdit** property setting is set to **True**, move the cursor to the extreme left of the row and the cursor image will change to an inverted arrow, allowing the selection of the entire row.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AllowDelete Property](Properties.md#allowdelete-property-gridex-control), [BeforeDeleteEx Event](#beforedeleteex-event-gridex-control), [BeforeDelete Event](#beforedelete-event-gridex-control), [Delete Method](Methods.md#delete-method-gridex-control)

## AfterGroupChange Event (GridEX Control)

Occurs after the user has changed, added or removed a group in the **GridEX** Control.

### Syntax

 **Private Sub** *object*_**AfterGroupChange**([ *index* **As Integer** ])  
 The **AfterGroupChange** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |

### Remarks

 Actions made by the user, such as dragging a column header into the “group by” box, changing the position of a group, or dragging a group outside the “group by” box, changes the group settings. The **BeforeGroupChange** event is fired when group settings are about to change in response to a user action and, unless canceled, the control will group the rows according to the new settings. Afterward the new grouping is applied, the **AfterGroupChange** event will be fired.  
 **Note:** The **AfterGroupChange** event will not be fired if the group settings are changed through code. This event is fired only in response to user interaction with the control.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AfterColMove Event](#aftercolmove-event-gridex-control), [BeforeColMove Event](#beforecolmove-event-gridex-control), [BeforeColumnDrag Event](#beforecolumndrag-event-gridex-control), [BeforeGroupChange Event](#beforegroupchange-event-gridex-control), [BeforeGroupDrag Event](#beforegroupdrag-event-gridex-control)

## AfterUpdate Event (GridEX Control)

Occurs after changes made by the user have been written to the database from a **GridEX** control.

### Syntax

 **Private Sub** *object*_**AfterUpdate**([ *index* **As Integer** ])  
 The **AfterUpdate** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |

### Remarks

 When the user moves to a different row after making changes to the field(s), data from the **GridEX** control's copy buffer is written to the database. Once the write is complete, the **AfterUpdate** event is triggered.  
 The **AfterUpdate** event occurs after the **BeforeUpdate** event if the **BeforeUpdate** event is not cancelled.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AfterColUpdate Event](#aftercolupdate-event-gridex-control), [BeforeColEdit Event](#beforecoledit-event-gridex-control), [BeforeColUpdate Event](#beforecolupdate-event-gridex-control), [BeforeUpdate Event](#beforeupdate-event-gridex-control), [AfterColEdit Event](#aftercoledit-event-gridex-control)

## BeforeColEdit Event (GridEX Control)

Occurs just before the user enters edit mode by typing a character or clicking in a cell.

### Syntax

 **Private Sub** *object*_**BeforeColEdit**([ *index* **As Integer**,] **ByVal** *colindex* **As Integer**, **ByVal** *cancel* **As JSRetBoolean**)  
 The **BeforeColEdit** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *Index* | An integer that identifies a control if it is in a control array. |
| *colindex* | An integer that identifies the index of the column to be edited. |
| *cancel* | A **JSRetBoolean** object that may be set to **True** to prevent the user from editing the cell, as described in Settings. |

### Settings

 The settings for cancel are:

| Setting | Description |
| --- | --- |
| **True** | The cell will not enter edit mode. |
| **False** | (Default) The cell will enter edit mode and the edit window will appear. |

### Remarks

 You can use this event to control which cell can be edited, on a per-cell basis. If you want to prevent edition on a per-column basis set the **EditType** property of the **JSColumn** to **jgexEditNone**.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AfterColUpdate Event](#aftercolupdate-event-gridex-control), [AfterUpdate Event](#afterupdate-event-gridex-control), [BeforeColUpdate Event](#beforecolupdate-event-gridex-control), [BeforeUpdate Event](#beforeupdate-event-gridex-control), [AfterColEdit Event](#aftercoledit-event-gridex-control)  
**Example:** [BeforeColEdit Example](../Examples.md#beforecoledit-example)

## BeforeColMove Event (GridEX Control)

Occurs before the **GridEX** control repaints after the user has dropped a column into a new position.

### Syntax

 **Private Sub** *object*_**BeforeColMove**([ *index* **As Integer**, ] *column* **As JSColumn**, **ByVal** *newposition* **As Integer**, **ByVal** *cancel* **As JSRetBoolean**)  
 The **BeforeColMove** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *column* | A **JSColumn** object that identifies the column that is going to be moved. |
| *newposition* | An integer that indicates the new position of the column. |
| *cancel* | A **JSRetBoolean** object that may be set to **True**, to prevent the user to move the column as described in settings. |

### Settings

 The settings for cancel are:

| Setting | Description |
| --- | --- |
| **True** | The column will not be moved to the new position. |
| **False** | (Default) The column will be moved to the new position. |

### Remarks

 The event is triggered after the user has dropped a column into a new position and before the **GridEX** control has been repainted.  
 When this event is fired, the **ColPosition** property of the **JSColumn** object has not been changed, so you can obtain the original position of the column that is being moved using the **ColPosition** property of the column argument.  
 If the cancel parameter is set to **True**, the reposition operation is canceled and the **AfterColMove** event will not be fired.  
 The **BeforeColMove** event is fired only if the control is in table view.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AfterColMove Event](#aftercolmove-event-gridex-control), [AfterGroupChange Event](#aftergroupchange-event-gridex-control), [BeforeColumnDrag Event](#beforecolumndrag-event-gridex-control), [BeforeGroupChange Event](#beforegroupchange-event-gridex-control), [BeforeGroupDrag Event](#beforegroupdrag-event-gridex-control)  
**Example:** [BeforeColMove Example](../Examples.md#beforecolmove-example)

## BeforeColumnDrag Event (GridEX Control)

Occurs before the user begins a drag operation with a header in the column headers row.

### Syntax

 **Private Sub** *object*_**BeforeColumnDrag** ( [ *index* **As Integer**,] *column* **As JSColumn**, **ByVal** *cancel* **As JSRetBoolean**)  
 The **BeforeColumnDrag** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *column* | A **JSColumn** object that identifies the column that is going to be dragged. |
| *cancel* | A **JSRetBoolean** object that may be set to **True**, to prevent the user to begin the dragging of a column as described in settings. |

### Settings

 The settings for cancel are:

| Setting | Description |
| --- | --- |
| **True** | The column can not be dragged by the user. |
| **False** | (Default) The user can drag the column. |

### Remarks

 The event is triggered before a column is dragged from the column headers row.  
 Use this event in order to prevent the user from interactively moving columns in a per-column basis. If you want to prevent dragging of all the columns, use the **AllowColumnDrag** property instead.  
 If the header that is about to be dragged is in the “group by” box the **BeforeGroupDrag** event is fired instead.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AfterColMove Event](#aftercolmove-event-gridex-control), [AfterGroupChange Event](#aftergroupchange-event-gridex-control), [BeforeColMove Event](#beforecolmove-event-gridex-control), [BeforeGroupChange Event](#beforegroupchange-event-gridex-control), [BeforeGroupDrag Event](#beforegroupdrag-event-gridex-control)  
**Example:** [BeforeColumnDrag Example](../Examples.md#beforecolumndrag-example)

## BeforeColUpdate Event (GridEX Control)

Occurs after editing is completed in a cell, but before data is moved from the cell to the **GridEX** control's copy buffer.

### Syntax

 **Private Sub** *object*_**BeforeColUpdate**( [ *index* **As Integer**,] **ByVal** *row* **As Long**, **ByVal** *colindex* **As Integer**, **ByVal** *oldvalue* **As String**, **ByVal** *cancel* **As JSRetBoolean**)  
 The **BeforeColUpdate** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *row* | A Long that identifies the row position. |
| *colindex* | An integer that represents the index of the column. |
| *oldvalue* | A String that contains the value contained in the cell before the change. |
| *cancel* | A **JSRetBoolean** object that specifies whether the change occurs, as described in Settings. |

### Settings

 The settings for cancel are:

| Setting | Description |
| --- | --- |
| **True** | Cancels the change and restores the cell to oldvalue. |
| **False** | (Default) Continues with the update operation. |

### Remarks

 The data moves from the cell to the control's copy buffer when a user completes editing within a cell, but the current row does not change. Before the data has been moved from the cell into the control's copy buffer, the **BeforeColUpdate** event is triggered. This event gives your application an opportunity to check the value(s) of the individual cells before they are committed to the control's copy buffer.  
 If your event procedure set the cancel argument to **True**, the previous value is restored to the cell and the **AfterColUpdate** event is not triggered.  
 The **AfterColUpdate** event occurs after the **BeforeColUpdate** event.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AfterColUpdate Event](#aftercolupdate-event-gridex-control), [AfterUpdate Event](#afterupdate-event-gridex-control), [BeforeColEdit Event](#beforecoledit-event-gridex-control), [BeforeUpdate Event](#beforeupdate-event-gridex-control), [AfterColEdit Event](#aftercoledit-event-gridex-control)  
**Example:** [BeforeColUpdate Example](../Examples.md#beforecolupdate-example)

## BeforeDelete Event (GridEX Control)

Occurs before a delete operation begins in a **GridEX** control.

### Syntax

 **Private Sub** *object*_**BeforeDelete** ( [ *index* **As Integer**,] **ByVal** *cancel* **As JSRetBoolean**)  
 The **BeforeDelete** event syntax has these parts:

| Part | Description |
| --- | --- |
| object | An object expression that evaluates to an object in the Applies To list. |
| index | An integer that identifies a control if it is in a control array. |
| cancel | A **JSRetBoolean** object that determines whether the record is allowed to be deleted, as described in Settings. |

### Settings

 The settings for cancel are:

| Setting | Description |
| --- | --- |
| True | The GridEX control will not delete the record. |
| False | (Default) Continues with the delete operation. |

### Remarks

 When the user selects an entire row and presses DEL, the **BeforeDelete** event is triggered before the selected row is deleted.  
 Once the row is deleted, the **AfterDelete** event is fired.  
 If your event procedure sets the cancel argument to **True**, the row is not deleted and the **AfterDelete** event will not be fired.

**Note**: To select the entire row if the **AllowEdit** property setting is set to **True**, move the cursor to the extreme left of the row and the cursor image will change to an inverted arrow, allowing the selection of the entire row.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AllowDelete Property](Properties.md#allowdelete-property-gridex-control), [BeforeDeleteEx Event](#beforedeleteex-event-gridex-control), [Delete Method](Methods.md#delete-method-gridex-control), [AfterDelete Event](#afterdelete-event-gridex-control)  
**Example:** [BeforeDelete Example](../Examples.md#beforedelete-example)

## BeforeDeleteEx Event (GridEX Control)

Occurs before a selected record is deleted in a **GridEX** control.

### Syntax

 **Private Sub** *object*_**BeforeDeleteEX** ( [ *index* **As Integer**,] **ByVal** *rowindex* **As Long**, **ByVal** *bookmark* **As Variant**, **ByVal** *cancel* **As JSRetBoolean**)  
 The **BeforeDeleteEX** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *rowindex* | A Long representing the RowIndex of the row that is about to be deleted. |
| *bookmark* | A variant representing the bookmark of the row that is about to be deleted. |
| *cancel* | A **JSRetBoolean** object whose Value determines whether a record is deleted or not, as described in Settings. |

### Settings

 The settings for cancel are:

| Setting | Description |
| --- | --- |
| **True** | The GridEX control will not delete the record. |
| **False** | (Default) Continues with the delete operation. |

### Remarks

 When the user selects one or more rows and presses DEL, the **BeforeDelete** event is triggered once before the delete operation is initiated. In addition to the **BeforeDelete** event, the **BeforeDeleteEX** event is triggered once for each record being deleted (this event is useful in a multi-select **GridEX** control).  
 Using this event you can selectively cancel the deletion of individual records within a range of selected rows in a **GridEx** control.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AllowDelete Property](Properties.md#allowdelete-property-gridex-control), [BeforeDelete Event](#beforedelete-event-gridex-control), [Delete Method](Methods.md#delete-method-gridex-control), [AfterDelete Event](#afterdelete-event-gridex-control)  
**Example:** [BeforeDeleteEx Example](../Examples.md#beforedeleteex-example)

## BeforeGroupChange Event (GridEX Control)

Occurs after the user has changed, added or removed a group in the **GridEX** control but before the change is committed.

### Syntax

 **Private Sub** *object*_**BeforeGroupChange**( [ *index* **As Integer**,] **ByVal** *group* **As JSGroup**, **ByVal** *changeoperation* **As jsgexGroupChange**, **ByVal** *groupposition* **As Integer**, **ByVal** *cancel* **As JSRetBoolean**)  
 The **BeforeGroupChange** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *group* | A **JSGroup** object that identifies the group that is about to change. |
| *changeoperation* | A value or constant that identifies the type of change in the group as described in Settings. |
| *groupposition* | An integer that indicates the position of the changing group. |
| *cancel* | A **JSRetBoolean** object that determines whether or not the change operation will be committed as described in Settings. |

### Settings

 The settings for changeoperation are:

| Constant | Setting | Description |  |
| --- | --- | --- | --- |
|  | **jgexGroupInsert** | 0 | A header dragged from the column header row was dropped into the “group by” box and is about to be inserted in the **JSGroups** collection. |
|  | **jgexGroupDelete** | 1 | A header dragged from the “group by” box was dropped outside the control or in the column header row and is about to be removed from the **JSGroups** collection. |
|  | **jgexGroupMove** | 2 | A header dragged from the “group by” box was dropped in the “group by” box again, but in a new position for the group. |

The settings for cancel are:

| Setting | Description |
| --- | --- |
| **True** | The group change operation is canceled and the group settings remain the same. |
| **False** | (Default) The group change operation will be committed. |

### Remarks

 The **BeforeGroupChange** event is fired whenever group settings are changed by the user, as when the user drags a column into the “group by” box, changes the position of a group, or drags a group outside the “group by” box. Unless canceled, the control will group the rows with the new settings, done this; the **AfterGroupChange** event is fired.  
 The **BeforeGroupChange** will not be fired if the group settings are changed through code.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AfterColMove Event](#aftercolmove-event-gridex-control), [AfterGroupChange Event](#aftergroupchange-event-gridex-control), [BeforeColumnDrag Event](#beforecolumndrag-event-gridex-control), [BeforeGroupChange Event](#beforegroupchange-event-gridex-control), [BeforeGroupDrag Event](#beforegroupdrag-event-gridex-control), [BeforeColMove Event](#beforecolmove-event-gridex-control)  
**Example:** [BeforeGroupChange Example](../Examples.md#beforegroupchange-example)

## BeforeGroupDrag Event (GridEX Control)

Occurs before the user begins a drag operation with a column header in the “group by” box.

### Syntax

 **Private Sub** *object*_**BeforeGroupDrag**( [ *index* **As Integer**, ] **ByVal** *group* **As JSGroup**, **ByVal** *cancel* **As JSRetBoolean**)  
 The **BeforeGroupDrag** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *group* | A JSGroup object that identifies the group that is going to be dragged. |
| *cancel* | A **JSRetBoolean** object that may be set to True, to prevent the user to begin the dragging of a group as described in Settings. |

### Settings

 The settings for cancel are:

| Setting | Description |
| --- | --- |
| True | The user cannot drag the group header. |
| False | (Default) The user can drag the group header. |

### Remarks

 This event is triggered before the dragging of a group begins.  
 You can use this event to prevent the dragging of individual group headers by the user. If you want to disable dragging in all the columns and all groups use the **AllowColumnDrag** property instead.  
 If the column header that is about to be dragged is in the column header row, the **BeforeColumnDrag** event is fired instead.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AfterColMove Event](#aftercolmove-event-gridex-control), [AfterGroupChange Event](#aftergroupchange-event-gridex-control), [BeforeColumnDrag Event](#beforecolumndrag-event-gridex-control), [BeforeGroupChange Event](#beforegroupchange-event-gridex-control), [BeforeColMove Event](#beforecolmove-event-gridex-control)

## BeforePrinting Event (GridEX Control)

Occurs before the **GridEX** control prints its contents.

### Syntax

 **Private Sub** *object*_**BeforePrinting**( [ *index* **As Integer**, ] **ByVal** *npages* **As Long**, **ByVal** *cancel* **As JSRetBoolean**)  
 The **BeforePrinting** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *npages* | A Long that represents the number of pages to be printed in a document. |
| *cancel* | A **JSRetBoolean** object that may be set to **True**, to prevent the printing of a document, as described in settings. |

### Settings

 The settings for cancel are:

| Setting | Description |
| --- | --- |
| **True** | The print operation in a **GridEX** control is canceled. |
| **False** | (Default) The **GridEX** control can start the print operation. |

### Remarks

 The **BeforePrinting** event is triggered before a print operation starts.  
 You can use this event to know how many pages will be printed in a document and to cancel the print operation in case you wanted so.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BeforePrintPage Event](#beforeprintpage-event-gridex-control), [PrintingProgress Event](#printingprogress-event-gridex-control), [PrinterProperties Property](Properties.md#printerproperties-property-gridex-control), [PrintGrid Method](Methods.md#printgrid-method-gridex-control), [JSPrinterProperties](Objects/JSPrinterProperties-Object.md#jsprinterproperties)  
**Example:** [BeforePrinting Example](../Examples.md#beforeprinting-example)

## BeforePrintPage Event (GridEX Control)

Occurs before the **GridEX** prints a page.

### Syntax

 **Private Sub** *object*_**BeforePrintPage**([ *index* **As Integer**, ] **ByVal** *pagenumber* **As Long**, **ByVal** *nPages* **As Long**)  
 The **BeforePrintPage** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *pagenumber* | A Long that represents the number of the page that is about to be printed. |
| *nPages* | A Long that represents the number of pages to be printed in a document. |

### Remarks

 The **BeforePrinting** event is triggered before a page is printed and is fired for every page in a document.  
 You can use this event to know the progress of the printing operation and to change the **PageFooter** of **PageHeader** property to print the page number in every page in a document.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [PrintingProgress Event](#printingprogress-event-gridex-control), [PrinterProperties Property](Properties.md#printerproperties-property-gridex-control), [PrintGrid Method](Methods.md#printgrid-method-gridex-control), [JSPrinterProperties](Objects/JSPrinterProperties-Object.md#jsprinterproperties), [BeforePrinting Event](#beforeprinting-event-gridex-control)  
**Example:** [BeforePrintPage Example](../Examples.md#beforeprintpage-example)

## BeforeUpdate Event (GridEX Control)

Occurs before data is committed from a **GridEX** control to the database.

### Syntax

 **Private Sub** *object*_**BeforeUpdate**( [ *index* **As Integer**,] **ByVal** *cancel* **As JSRetBoolean**)  
 The **BeforeUpdate** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *cancel* | A **JSRetBoolean** object that determines if data is written to the database, as described in Settings. |

### Settings

 The settings for cancel are:

| Setting | Description |
| --- | --- |
| **True** | The **GridEX** control does not write changes to the database. |
| **False** | (Default) Continues with the update operation. |

### Remarks

 When the user moves to another row or the control is about to update the sorting or grouping settings, the **Update** method is executed and data from the **GridEX** control's copy buffer is written to the database.  
 Just before the changes are committed to the database, the **BeforeUpdate** event is triggered. Unless the update operation is canceled, the **AfterUpdate** event is fired after the data has been written to the database.  
 If you set the **BeforeUpdate** event cancel argument to **True**, the **AfterUpdate** event will not be fired, and the record will not saved to the database.  
 You can use this event to validate data in the current record before allowing the user to commit the changes to the database. If you set the cancel argument to **True**, the user will not be able to move to another record in the control. If you want to restore the record values and cancel the update operation, you can set the **GridEX** control’s **DataChanged** property to **False**.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AfterColUpdate Event](#aftercolupdate-event-gridex-control), [AfterUpdate Event](#afterupdate-event-gridex-control), [BeforeColEdit Event](#beforecoledit-event-gridex-control), [BeforeColUpdate Event](#beforecolupdate-event-gridex-control), [AfterColEdit Event](#aftercoledit-event-gridex-control)  
**Example:** [BeforeUpdate Example](../Examples.md#beforeupdate-example)

## CardResize Event (GridEX Control)

Occurs when the user changes the **CardWith** by dragging the bar draw between cards in a **GridEX** control.

### Syntax

 **Private Sub** *object*_**CardResize**([ *index* **As Integer**,] **ByVal** *newcardwidth* **As Long**, **ByVal** *cancel* **As JSRetBoolean**)

The **CardResize** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *newcardwidth* | A Long expression that represents the new width for cards. |
| *cancel* | A **JSRetBoolean** object that determines whether cards' width is changed, as described in Settings. |

### Settings

 The settings for *cancel* are:

| Setting | Description |
| --- | --- |
| **True** | Cancels the change and cards remain with their original width. |
| **False** | (Default) Continues the width change operation. |

### Remarks

 This event is only triggered if **AllowCardSizing** is **True**.  
 The **CardResize** event is triggered when the user is about to change the width of cards in a **GridEX** control. Your event procedure can accept, cancel, or alter the amount of change.  
 If you set the *cancel* argument to **True**, the card width is restored. To alter the degree of change, set the **CardWidth** property to the desired value.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [CardWidth Property](Properties.md#cardwidth-property-gridex-control), [AllowCardSizing Property](Properties.md#allowcardsizing-property-gridex-control)

## Change Event (GridEX Control)

Occurs when the contents of the current cell have changed.

### Syntax

 **Private Sub** *object*_**Change**( [ *index* **As Integer** ] )  
 The **Change** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that uniquely identifies a control if it's in a control array |

### Remarks

 Once the user has entered edit mode, the **Change** event is fired every time the information on the cell is modified. For TextBox cells, this occurs as the user enters or changes the characters in the cell. For combo-style cell dropdowns, such as date columns, the event is fired only once, when the user selects an item from the dropdown.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AfterColEdit Event](#aftercoledit-event-gridex-control), [BeforeColEdit Event](#beforecoledit-event-gridex-control)

## Click Event (GridEX Control)

Occurs when the user presses and then releases a mouse button over an object

### Syntax

 **Private Sub** *object*_**Click**( [ *index* **As Integer**])  
 The Click event syntax has these parts:

| Part | Description |
| --- | --- |
| object | An object expression that evaluates to an object in the Applies To list. |
| index | An integer that uniquely identifies a control if it's in a control array. |

### Remarks

 Clicking the **GridEX** control, generates **MouseDown** and **MouseUp** events in addition to the **Click** event.  
 The events occur in this order: **MouseDown**, **MouseUp**, and **Click**.  
 When you are attaching code to these related events procedures, be sure that their actions do not conflict.

**Note** To distinguish between the left, right, and middle mouse buttons, use the MouseDown and MouseUp events.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [DblClick Event](#dblclick-event-gridex-control), [MouseUp Event](#mouseup-event-gridex-control), [MouseDown Event](#mousedown-event-gridex-control), [MouseMove Event](#mousemove-event-gridex-control)

## ColButtonClick Event (GridEX Control)

Occurs when the user has clicked a button in a cell.

### Syntax

 **Private Sub** *object*_**ColButtonClick** ([ *index* **As Integer**, ] **ByVal** *colindex* **As Integer**)  
 The **ColButtonClick** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *colindex* | An integer that identifies the index of the column clicked. |

### Remarks

 This event is triggered only for cells whose **ButtonStyle** property is not **jgexNobutton**.  
 In columns with its **EditType** set to **jgexEditDropDown** or **jgexEditCalendarDropDown** this event is not triggered even if a cell has a button.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ButtonStyle Property](Objects/JSColumn-Object.md#buttonstyle-property-jscolumn-object)  
**Example:** [ColButtonClick Example](../Examples.md#colbuttonclick-example)

## ColResize Event (GridEX Control)

Occurs when a user resizes a column in a **GridEX** control.

### Syntax

 **Private Sub** *object*_**ColResize**([ *index* **As Integer**,] **ByVal** *colindex* **As Integer**, **ByVal** *newcolwidth* **As Long**, **ByVal** *cancel* **As JSRetBoolean**)  
 The **ColResize** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *colindex* | An integer that represents the index of the column that is about to be resized. |
| *newcolwidth* | A Long expression that represents the new width for the column being resized. |
| *cancel* | A **JSRetBoolean** object that determines whether the column is resized, as described in Settings. |

### Settings

 The settings for cancel are:

| Setting | Description |
| --- | --- |
| **True** | Cancels the change and restores the column to its original width. |
| **False** | (Default) Continues the width change operation. |

### Remarks

 When the user resizes a column, the **ColResize** event is triggered. Your event procedure can accept, cancel, or alter the amount of change.  
 If you set the cancel argument to **True**, the column width is restored. To alter the degree of change, set the **Width** property of the **JSColumn** object to the desired value.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ColumnAutoResize Property](Properties.md#columnautoresize-property-gridex-control), [Autosize Method](Objects/JSColumn-Object.md#autosize-method-jscolumn-object), [AllowSizing Property](Objects/JSColumn-Object.md#allowsizing-property-jscolumn-object)

## ColumnHeaderClick Event (GridEX Control)

Occurs when the user clicks on the header for a particular column of a **GridEX** control while it is in table view.

### Syntax

 **Private Sub** *object*_**ColumnHeaderClick**([ *index* **As Integer**,] **ByVal** *column* **As JSColumn**)  
 The **ColumnHeaderClick** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *column* | A **JSColumn** object that identifies the column. |

### Remarks

 One possible use for this event is to sort the records in the **GridEX** control based on the selected column. While the control groups automatically, due to user interaction, the actual sorting of records has to be done by your program. An example of a standard procedure that does this could similar to this:

```vb
Private Sub GridEX1_ColumnHeaderClick(ByVal Column As JSColumn)

Dim SortOrder as integer

SortOrder = Column.SortOrder
'Remove the existing sorting criteria, if any
GridEX1.SortKeys.Clear
If SortOrder = jgexSortAscending Then
GridEX1.SortKeys.Add Column.Index, jgexSortDescending
Else 'if SortOrder is none or is descending
GridEX1.SortKeys.Add Column.Index, jgexSortAscending
End If
End Sub
```

**Applies To:** [ColumnHeaderClick Event](#columnheaderclick-event-gridex-control)  
**See Also:** [GroupByBoxHeaderClick Event](#groupbyboxheaderclick-event-gridex-control)  
**Example:** [ColumnHeaderClick Example](../Examples.md#columnheaderclick-example)

## DblClick Event (GridEX Control)

Occurs when the user presses and releases a mouse button two times in succession over a **GridEX** control.

### Syntax

 **Private Sub** *object*_**DblClick**([ *index* **As Integer** ])

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | Identifies the control if it is in a control array. |

### Remarks

 The mouse events in a **GridEX** control occur in this order: **MouseDown**, **MouseUp**, **Click**, **DblClick**, and **MouseUp**.  
 If **DblClick** does not occur within the system's double-click time limit, the object recognizes another **Click** event.  
 You can change your system’s double click time using the Keyboard applet of the Control Panel.

**Note** To distinguish between the left, right, and middle mouse buttons use the **MouseDown** and **MouseUp** events.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [MouseUp Event](#mouseup-event-gridex-control), [MouseDown Event](#mousedown-event-gridex-control), [MouseMove Event](#mousemove-event-gridex-control), [Click Event](#click-event-gridex-control)

## DropList Event (GridEX Control)

Occurs before a drop down list appears.

### Syntax

 **Private Sub** *object*_**DropList**([ *index* **As Integer**,] **ByVal** *colindex* **As Integer**)  
 The **DropList** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *colindex* | An integer that represents the index of the column. |

### Remarks

 Use this event if you require more control over how lists are displayed. For example, if you want to change list entries before the user opens a drop down list.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ListSelected Event](#listselected-event-gridex-control)  
**Example:** [DropList Example](../Examples.md#droplist-example)

## EndCustomEdit Event (GridEX Control)

Occurs when a custom edit operation is ended. This event is only raised for columns with **EditType** equal to **jgexEditCustom**.

### Syntax

 **Private Sub** *object*_**EndCustomEdit** ([ *index* **As Integer**, ] **ByVal** *colindex* **As Integer**)  
 The **EndCustomEdit** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *colindex* | An integer that identifies the index of the custom edit column that is about to end the edit operation. |

### Remarks

 In this event you must set value of the current cell to the value in the control being used for the edit operation.  
 This event is triggered only for cells with its **EditType** property equal to **jgexEditcustom**.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [EditType Property](Objects/JSColumn-Object.md#edittype-property-jscolumn-object), [InitCustomEdit Event](#initcustomedit-event-gridex-control), [ShowCustomEdit Event](#showcustomedit-event-gridex-control)  
**Example:** [CustomEdit Example](../Examples.md#customedit-example)

## Error Event (GridEX Control)

Occurs as the result of a data access error.

### Syntax

 **Private Sub** *object*_**Error**([ *index* **As Integer**,] **ByVal** *errnumber* **As Long**, *displaymessage* **As Boolean**)  
 The **Error** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An Integer that identifies a control if it is in a control array. |
| *errnumber* | A Long value that identifies the error that occurred. |
| *displaymessage* | A Boolean that may be set to False to suppress any error message displays, as described in Settings. |

### Settings

 The settings for displaymessage are:

| Setting | Description |
| --- | --- |
| **False** | No error message will be displayed. |
| **True** | (Default) The message associated with the error will be displayed by the control. |

### Remarks

 Even if your application handles run time errors in code, errors can still occur when none of your code is executing.  
 An example of this is when the **GridEX** control tries to create a **Recordset** and one or more properties had been set incorrectly.  
 Another example is when the **GridEX** control tries to write changes made by the user into a record that is locked by another user. If a data access error results from such an action, the **Error** event is fired.

**Note** Use the **ErrorText** property to retrieve or modify the error string that will be displayed.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ErrorText Property](Properties.md#errortext-property-gridex-control)  
**Example:** [Error Example](../Examples.md#error-example)

## FetchData Event (GridEX Control)

Fetches unbound column data when the **GridEX** is in DAO or ADO mode.

### Syntax

 **Private Sub** *object*_**FetchData**([ *index* **As Integer**,] **ByVal** *rowindex* **As Long**, **ByVal** *colindex* **As Integer**, **ByVal** *rowbookmark* **As Variant**, **ByVal** *value* **As JSRetVariant**)  
 The **FetchData** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *rowindex* | A long that identifies the index of the record being fetched. |
| *colindex* | An integer that identifies the index of the column being fetched. |
| *rowbookmark* | A variant that identifies the bookmark of the record being fetched. |
| *value* | A **JSRetVariant** object that can be set to the value of the cell being fetched. |

**Remarks**:  
 This event is only fired for the columns whose **FetchData** property is set to **True**.  
 The *rowindex* parameter represents the index of the record as it was originally in the **Recordset**.  
 Because the position of a record could be changed if the control is sorted or grouped, the row index will not always match with its position.  
 The *colindex* parameter also represents the index of the column and not its position, so that your code in the event procedure can reference the object without worrying if the user has changed the column position.  
 The *rowbookmark* parameter could be used if you want to access the record values via the **Recordset** or the **ADORecordset** property. A **FetchData** event procedure for a calculated field could be coded as follows:

```vb
Private Sub GridEX1_FetchData(ByVal RowIndex As Long, _
ByVal ColIndex As Integer, _
ByVal RowBookmark As Variant, _
Value As GridEX16.JSRetVariant)

dim rsTemp as Recordset
Set rsTemp=GridEX1.Recordset

rsTemp.Bookmark = RowBookmark
'The value is calculated adding the value of field 1
'to the value of field 3
Value = rsTemp.Fields(1).Value + rsTemp.Fields(3).Value
End Sub
```

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [FetchData Property](Objects/JSColumn-Object.md#fetchdata-property-jscolumn-object), [FetchIcon Event](#fetchicon-event-gridex-control), [FetchIcon Property](Objects/JSColumn-Object.md#fetchicon-property-jscolumn-object)  
**Example:** [FetchData Example](../Examples.md#fetchdata-example)

## FetchIcon Event (GridEX Control)

Fetches icons for any column whose **FetchIcon** property is set to **True**.

### Syntax

 **Private Sub** *object*_**FetchIcon**([ *index* **As Integer**,] **ByVal** *rowindex* **As Long**, **ByVal** *colindex* **As Integer**, **ByVal** *rowbookmark* **As Variant**, **ByVal** *iconindex* **As JSRetInteger**)  
 The **FetchIcon** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *rowindex* | A long that identifies the index of the record being fetched. |
| *colindex* | An integer that identifies the index of the column being fetched. |
| *rowbookmark* | A variant that identifies the bookmark of the record being fetched. |
| *iconindex* | A **JSRetInteger** object that can be set to the icon index of the cell being fetched. |

### Remarks

 This event is only fired for the columns whose property **FetchIcon** is set to **True**.  
 The *rowindex* parameter represents the index of the record as it was originally in the **Recordset**. Because the position of a record could be changed if the control is sorted or grouped, the row index will not always match with its position.  
 The *colindex* parameter also represents the index of the column and not its actual position. Because of this, you can write the event procedure without worrying if the user has changed the column position.  
 If you want to access the record values via the **Recordset** or the **ADORecordset** property, you could use the *rowbookmark* parameter. A **FetchIcon** event procedure could be coded as follows:

```vb
Private Sub GridEX1_FetchIcon (ByVal RowIndex As Long, _
ByVal ColIndex As Integer, _
ByVal RowBookmark As Variant, _
IconIndex As GridEX16.JSRetInteger)

Dim rsTemp as Recordset
const IconIndexPaid = 1

const IconIndexUnPaid = 2
set rsTemp = GridEX1.Recordset
rsTemp.Bookmark = RowBookmark
'The icon depends of the value of "Paid" Field
if rsTemp![Paid] = True then
IconIndex = IconIndexPaid
Else

IconIndex = IconIndexUnpaid

End if
End Sub
```

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [FetchData Event](#fetchdata-event-gridex-control), [FetchData Property](Objects/JSColumn-Object.md#fetchdata-property-jscolumn-object), [FetchIcon Property](Objects/JSColumn-Object.md#fetchicon-property-jscolumn-object)  
**Example:** [FetchIcon Example](../Examples.md#fetchicon-example)

## FirstItemChange Event (GridEX Control)

Fired when the **FirstItem** property changes.

### Syntax

 **Private Sub** *object*_**FirstItemChange**([ *index* **As Integer**])  
 The **FirstItemChange** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |

### Remarks

 The **FirstItemChange** event is fired when the **GridEX** control’s first item displayed changes; for example, when the user scrolls the control.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [FirstItem Property](Properties.md#firstitem-property-gridex-control)

## GroupByBoxHeaderClick Event (GridEX Control)

Occurs when the user clicks on a header group of a **GridEX** control.

### Syntax

 **Private Sub** *object*_**GroupByBoxHeaderClick**([ index **As Integer**,] *group* **As JSGroup**)  
 The **GroupByBoxHeaderClick** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *group* | A **JSGroup** object that identifies the group. |

### Remarks

 One possible use for this event is to change the sort order of the group being clicked. While the control groups automatically based on user interaction, actually changing the sorting criteria of a group has to be done by application.  
 A standard procedure that does this could be as follows:

```vb
'This code toggles the sorting order for the JSGroup object passed in the Group argument of the event.
Private Sub GridEX1_GroupByBoxHeaderClick(Group As JSGroup)

If Group.SortOrder = jgexSortAscending Then
Group.SortOrder = jgexSortDescending
Else
Group.SortOrder = jgexSortAscending

End If
End Sub
```

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ColumnHeaderClick Event](#columnheaderclick-event-gridex-control)  
**Example:** [ColumnHeaderClick Example](../Examples.md#columnheaderclick-example)

## InitCustomEdit Event (GridEX Control)

Occurs when an edit operation is about to begin in a custom edit column.

### Syntax

 **Private Sub** *object*_**InitCustomEdit** ([ *index* **As Integer**, ] **ByVal** *colindex* **As Integer**, **ByVal** *editbackcolor* **As OLE_COLOR**, **ByVal** *editforecolor* **As OLE_COLOR**, **ByVal** *editfont* **As StdFont**)  
 The **InitCustomEdit** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *colindex* | An integer that identifies the index of the custom edit column that is about to initialize the edit operation. |
| *editbackcolor* | An OLE_COLOR that represents the background color suggested for the control used as custom edit. |
| *editforecolor* | An OLE_COLOR that represents the foreground color suggested for the control used as custom edit. |
| *editfont* | A StdFont object that represents the font suggested for the control used as custom edit. |

### Remarks

 In this event you must set value of the control being used for the edit operation as the value of the current cell in **GridEX** control.

**Note**: This event is triggered only for cells whose **EditType** property is set to **jgexEditcustom**.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [EditType Property](Objects/JSColumn-Object.md#edittype-property-jscolumn-object), [EndCustomEdit Event](#endcustomedit-event-gridex-control), [ShowCustomEdit Event](#showcustomedit-event-gridex-control)  
**Example:** [CustomEdit Example](../Examples.md#customedit-example)

## KeyDown Event (GridEX Control)

Occur when the user presses (**KeyDown**) a key while an object has the focus.

### Syntax

 **Private Sub** *object*_**KeyDown**( [ *index* **As Integer**,] *keycode* **As Integer**, *shift* **As Integer**)  
 The **KeyDown** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that uniquely identifies a control if it's in a control array. |
| *keycode* | An integer that represents the key code. |
| *shift* | An integer that corresponds to the state of the SHIFT, CTRL, and ALT keys at the time of the event. The shift argument is a bit field with the least-significant bits corresponding to the SHIFT key (bit 0), the CTRL key (bit 1), and the ALT key (bit 2). These bits correspond to the values 1, 2, and 4, respectively. Some, all, or none of the bits can be set, indicating that some, all, or none of the keys are pressed. For example, if both CTRL and ALT are pressed, the value of shift is 6. |

### Remarks

 Use **KeyDown** and **KeyUp** event procedures if you need to respond to both the pressing(**KeyDown**) and releasing(**KeyUp**) of a key.  
 You test for a condition by first assigning each result to a temporary integer variable and then comparing shift to a bit mask.  
 Use the And operator with the shift argument to test whether the condition is greater than 0, indicating that the modifier was pressed, as in this example:

```vb
ShiftDown = (Shift And 1) > 0

CtrlDown = (Shift And 2) > 0

AltDown = (Shift And 4) > 0
```

In a procedure, you can test for any combination of conditions, as in this example:

```vb
If AltDown And CtrlDown Then
```

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [KeyUp Event](#keyup-event-gridex-control), [KeyPress Event](#keypress-event-gridex-control)

## KeyPress Event (GridEX Control)

Occurs when the user presses and releases an ANSI key.

### Syntax

 **Private Sub** *object*_**KeyPress**( [ *index* **As Integer**,] *keyascii* **As Integer**)  
 The **KeyPress** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that uniquely identifies a control if it's in a control array. |
| *keycode* | An integer that represents the key code. |
| *keyascii* | An integer that returns a standard numeric ANSI keycode. *keyascii* is passed by reference; changing it sends a different character to the object. <br> <br> Changing *keyascii* to 0 cancels the keystroke so the object receives no character. |

### Remarks

 The **KeyPress** event enables you to immediately test keystrokes for validity or formatting of characters as they are typed.  
 Changing the value of the keyascii argument changes the character displayed.  
 Use **KeyDown** and **KeyUp** event procedures to handle any keystroke not recognized by **KeyPress**, such as function keys, editing keys, navigation keys, and any combinations of these with keyboard modifiers. Unlike the **KeyDown** and KeyUp **events**, **KeyPress** does not indicate the physical state of the keyboard; instead, it just passes a character value.  
 **KeyPress** interprets the uppercase and lowercase of each character as separate key codes and, therefore, as two separate characters.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [KeyDown Event](#keydown-event-gridex-control), [KeyUp Event](#keyup-event-gridex-control)

## KeyUp Event (GridEX Control)

Occur when the user releases (**KeyUp**) a key while an object has the focus.

### Syntax

 **Private Sub** *object*_**KeyUp**( [ *index* **As Integer**,] *keycode* **As Integer**, *shift* **As Integer**)  
 The **KeyUp** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that uniquely identifies a control if it's in a control array. |
| *keycode* | An integer that represents the key code. |
| *shift* | An integer that corresponds to the state of the SHIFT, CTRL, and ALT keys at the time of the event. The shift argument is a bit field with the least-significant bits corresponding to the SHIFT key (bit 0), the CTRL key (bit 1), and the ALT key (bit 2). These bits correspond to the values 1, 2, and 4, respectively. Some, all, or none of the bits can be set, indicating that some, all, or none of the keys are pressed. For example, if both CTRL and ALT are pressed, the value of shift is 6. |

### Remarks

 Use KeyDown and KeyUp event procedures if you need to respond to both the pressing(**KeyDown**) and releasing(**KeyUp**) of a key.  
 You test for a condition by first assigning each result to a temporary integer variable and then comparing shift to a bit mask.  
 Use the And operator with the shift argument to test whether the condition is greater than 0, indicating that the modifier was pressed, as in this example:

```vb
ShiftDown = (Shift And 1) > 0

CtrlDown = (Shift And 2) > 0

AltDown = (Shift And 4) > 0
```

In a procedure, you can test for any combination of conditions, as in this example:

```vb
If AltDown And CtrlDown Then
```

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [KeyDown Event](#keydown-event-gridex-control), [KeyPress Event](#keypress-event-gridex-control)

## LeftColChange Event (GridEX Control)

Fired when the **LeftCol** property changes.

### Syntax

 **Private Sub** *object*_**LeftColChange**([ *index* **As Integer** ] )  
 The **LeftColChange** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |

### Remarks

 The **LeftColChange** event is fired at any time the **GridEX** control’s displayed leftmost column changes; for example, when the user horizontally scrolls the control.  
 This event is not fired when the **GridEX** control is in card view.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [FirstItem Property](Properties.md#firstitem-property-gridex-control), [FirstItemChange Event](#firstitemchange-event-gridex-control), [LeftCol Property](Properties.md#leftcol-property-gridex-control)

## ListSelected Event (GridEX Control)

Occurs after the user has selected an entry in a drop down list.

### Syntax

 **Private Sub** *object*_**ListSelected**( [ *index* **As Integer**,] **ByVal** *colindex* **As Integer**, **ByVal** *valuelistindex* **As Long**, **ByVal** *value* **As Variant**)  
 The **ListSelected** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *colindex* | An integer that represents the index of the column. |
| *valuelistindex* | A long that represents the index of the selected **JSValueItem**. |
| *value* | A variant expression that represents the value of the selected **JSValueItem**. |

### Remarks

 If you need to retrieve the text of the selected list item, you must obtain it from the **JSValueItem** object’s Text property as follows:

```vb
Private Sub GridEX1_ListSelected(ByVal ColIndex As Integer, ByVal ValueListIndex As Long, ByVal Value As Variant)

Dim strText As String

strText = GridEX1.Columns(ColIndex).ValueList(ValueListIndex).Text

Debug.Print "Text Selected: "; strText

End Sub
```

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [DropList Event](#droplist-event-gridex-control)

## MouseDown Event (GridEX Control)

Occur when the user presses (**MouseDown**) a mouse button.

### Syntax

 **Private Sub** *object*_**MouseDown**([ *index* **As Integer**,] *button* **As Integer**, *shift* **As Integer**, *x* **As Single**, *y* **As Single**)  
 The **MouseDown** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that uniquely identifies a control if it's in a control array. |
| *button* | An integer that identifies the button that was pressed (MouseDown) to cause the event. |
| *shift* | An integer that corresponds to the state of the SHIFT, CTRL, and ALT keys when the button specified in the button argument is pressed or released. |
| *x, y* | A single that specifies the current location of the mouse pointer. The *x* and *y* values are always expressed in twips. |

### Remarks

 You can use the **MouseDown** or **MouseUp** event procedure to specify actions that will occur when a mouse button is pressed or released.  
 Unlike the **Click** and **DblClick** events, **MouseDown** and **MouseUp** events enable you to distinguish between the left, right, and middle mouse buttons.  
 You can also write code for mouse-keyboard combinations that use the SHIFT, CTRL, and ALT keyboard modifiers.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [DblClick Event](#dblclick-event-gridex-control), [MouseUp Event](#mouseup-event-gridex-control), [MouseMove Event](#mousemove-event-gridex-control), [Click Event](#click-event-gridex-control)

## MouseMove Event (GridEX Control)

Occurs when the user moves the mouse over the control, or outside the boundaries of the control if any of the mouse buttons are pressed.

### Syntax

 **Private Sub** *object*_**MouseMove**( [ *index* **As Integer**,] *button* **As Integer**, *shift* **As Integer**, *x* **As Single**, *y* **As Single**)  
 The **MouseMove** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| index | An integer that uniquely identifies a control if it is in a control array. |
| button | An integer that corresponds to the state of the mouse buttons in which a bit is set if the button is down. |
| shift | An integer that corresponds to the state of the SHIFT, CTRL, and ALT keys. |
| x, y | A single that specifies the current location of the mouse pointer. The x and y values are always expressed in twips |

### Remarks

 The **MouseMove** event is generated continually as the mouse pointer moves across objects.  
 Unless another object has captured the mouse, an object recognizes a **MouseMove** event whenever the mouse position is within its borders.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [DblClick Event](#dblclick-event-gridex-control), [MouseUp Event](#mouseup-event-gridex-control), [MouseDown Event](#mousedown-event-gridex-control), [Click Event](#click-event-gridex-control)

## MouseUp Event (GridEX Control)

Occur when the user releases (**MouseUp**) a mouse button.

### Syntax

 **Private Sub** *object*_**MouseUp**([ *index* **As Integer**,] *button* **As Integer**, *shift* **As Integer**, *x* **As Single**, *y* **As Single**)  
 The **MouseUp** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that uniquely identifies a control if it's in a control array. |
| *button* | An integer that identifies the button that was released (MouseUp) to cause the event. |
| *shift* | An integer that corresponds to the state of the SHIFT, CTRL, and ALT keys when the button specified in the button argument is pressed or released. |
| *x, y* | A single that specifies the current location of the mouse pointer. The *x* and *y* values are always expressed in twips. |

### Remarks

 You can use the **MouseDown** or **MouseUp** event procedure to specify actions that will occur when a mouse button is pressed or released.  
 Unlike the **Click** and **DblClick** events, **MouseDown** and **MouseUp** events enable you to distinguish between the left, right, and middle mouse buttons.  
 You can also write code for mouse-keyboard combinations that use the SHIFT, CTRL, and ALT keyboard modifiers.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [DblClick Event](#dblclick-event-gridex-control), [MouseDown Event](#mousedown-event-gridex-control), [MouseMove Event](#mousemove-event-gridex-control), [Click Event](#click-event-gridex-control)

## NotInList Event (GridEX Control)

Occurs when the text entered by the user in a combo column is not present in the list.

### Syntax

 **Private Sub** *object*_**NotInList**([ *index* **As Integer**,] **ByVal** *colindex* **As Integer**, **ByVal** *newdata* **As String**, **ByVal** *cancelupdate* **As JSRetBoolean**)  
 The **NotInList** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *colindex* | An integer that represents the index of the column. |
| *newdata* | A string that represents the data entered that wan't encountered in the list. |
| *cancelupdate* | A **JSRetBoolean** object that may be set to **True**, to prevent the update of the cell, as described in settings. |

### Settings

 The settings for *cancelupdate* are:

| Setting | Description |
| --- | --- |
| **True** | Default. The **GridEX** control will not accept the entry and the update operation will be canceled. |
| **False** | The **GridEX** control assumes you have added the entry into the Drop down grid and it will search again for a match. |

### Remarks

 Use this event to indicate your users that there is no match in the list for the data they entered or to add the entry to the **GridEX** control that acts as the drop down control for the column.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [ActAsDropDown Property](Properties.md#actasdropdown-property-gridex-control), [EditType Property](Objects/JSColumn-Object.md#edittype-property-jscolumn-object)  
**Example:** [NotInList Example](../Examples.md#notinlist-example)

## OLECompleteDrag Event (GridEX Control)

Occurs when a source component is dropped onto a target component, informing the source component that a drag action was either performed or canceled.

### Syntax

 **Private Sub** *object*_**OLECompleteDrag** ([ *index* **As Integer**,] *effect* **As Long**)  
 The **OLECompleteDrag** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *effect* | A long set by the source object identifying the action that has been performed, thus allowing the source to take appropriate action if the component was moved (such as the source deleting data if it is moved from one component to another). The possible values are listed in Settings. |

### Settings

 The settings for *effect* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexDropEffectNone** | 0 | Drop target cannot accept the data. |
|  **jgexDropEffectCopy** | 1 | Drop results in a copy of data from the source to the target. The original data is unaltered by the drag operation. |
|  **jgexDropEffectMove** | 2 | Drop results in data being moved from drag source to drop source. The drag source should remove the data from itself after the move. |

### Remarks

 The **OLECompleteDrag** event is the final event to be called in an OLE drag/drop operation.  
 This event informs the source component of the action that was performed when the object was dropped onto the target component. The target sets this value through the effect parameter of the **OLEDragDrop** event. Based on this, the source can then determine the appropriate action it needs to take. For example, if the object was moved into the target (**jgexDropEffectMove**), the source needs to delete the object from itself after the move.  
 The **GridEX** control only supports manual OLE drag and drop events.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [OLEDragDrop Event](#oledragdrop-event-gridex-control), [OLEDrag Method](Methods.md#oledrag-method-gridex-control), [OLEDragOver Event](#oledragover-event-gridex-control), [OLEDropMode Property](Properties.md#oledropmode-property-gridex-control), [OLEGiveFeedback Event](#olegivefeedback-event-gridex-control), [OLESetData Event](#olesetdata-event-gridex-control), [OLEStartDrag Event](#olestartdrag-event-gridex-control)

## OLEDragDrop Event (GridEX Control)

Occurs when a source component is dropped onto a target component when the source component determines that a drop can occur.

### Syntax

 **Private Sub** *object*_**OLEDragDrop**([ *index* **As Integer**,] *data* **As JSDataObject**, *effect* **As Long**, *button* **As Integer**, *shift* **As Integer**, *x* **As Single**, *y* **As Single**)  
 The **OLEDragDrop** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *data* | A **JSDataObject** object containing formats that the source will provide and, in addition, possibly the data for those formats. If no data is contained in the **JSDataObject**, it is provided when the control calls the **GetData** method. The **SetData** and **Clear** methods cannot be used here. |
| *effect* | A Long set by the target component identifying the action that has been performed (if any), thus allowing the source to take appropriate action if the component was moved (such as the source deleting the data). The possible values are listed in Settings. |
| *button* | An integer which acts as a bit field corresponding to the state of a mouse button when it is depressed. The left button is bit 0, the right button is bit 1, and the middle button is bit 2. These bits correspond to the values 1, 2, and 4, respectively. It indicates the state of the mouse buttons; some, all, or none of these three bits can be set, indicating that some, all, or none of the buttons are depressed. |
| *shift* | An integer which acts as a bit field corresponding to the state of the SHIFT, CTRL, and ALT keys when they are depressed. The SHIFT key is bit 0, the CTRL key is bit 1, and the ALT key is bit 2. These bits correspond to the values 1, 2, and 4, respectively. The shift parameter indicates the state of these keys; some, all, or none of the bits can be set, indicating that some, all, or none of the keys are depressed. For example, if both the CTRL and ALT keys were depressed, the value of shift would be 6. |
| *x,y* | A number which specifies the current location of the mouse pointer. The x and y values are always expressed in container coordinates. |

### Settings

 The settings for *effect* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexDropEffectNone** | 0 | Drop target cannot accept the data. |
|  **jgexDropEffectCopy** | 1 | Drop results in a copy of data from the source to the target. The original data is unaltered by the drag operation. |
|  **jgexDropEffectMove** | 2 | Drop results in data being moved from drag source to drop source. The drag source should remove the data from itself after the move. |

### Remarks

 The source should always mask values from the effect parameter to ensure compatibility with future implementations of ActiveX components. Presently, only three of the 32 bits in the effect parameter are used. In future versions of Visual Basic, however, these other bits may be used. Therefore, as a precaution against future problems, drag sources and drop targets should mask these values appropriately before performing any comparisons.  
 For example, a source component should not compare an effect against, say, **jgexDropEffectCopy**, such as the following code:

```vb
If Effect = jgexDropEffectCopy Then...
Instead, the source component should mask for the value or values being sought, such as this:
If Effect And jgexDropEffectCopy = jgexDropEffectCopy...
-or-
If (Effect And jgexDropEffectCopy)...
```

This allows for the definition of new drop effects in future versions of Visual Basic while preserving backwards compatibility with your existing code.  
 The **GridEX** control only supports manual OLE drag and drop events.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [OLECompleteDrag Event](#olecompletedrag-event-gridex-control), [OLEDrag Method](Methods.md#oledrag-method-gridex-control), [OLEDragOver Event](#oledragover-event-gridex-control), [OLEDropMode Property](Properties.md#oledropmode-property-gridex-control), [OLEGiveFeedback Event](#olegivefeedback-event-gridex-control), [OLESetData Event](#olesetdata-event-gridex-control), [OLEStartDrag Event](#olestartdrag-event-gridex-control)

## OLEDragOver Event (GridEX Control)

Occurs when one component is dragged over another.

### Syntax

 **Private Sub** *object*_**OLEDragOver** ([ *index* **As Integer**,] *data* **As JSDataObject**, *effect* **As Long**, *button* **As Integer**, *shift* **As Integer**, *x* **As Single**, *y* **As Single**, *state* **As Integer**)  
 The **OLEDragDrop** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *data* | A **JSDataObject** object containing formats that the source will provide and, in addition, possibly the data for those formats. If no data is contained in the **JSDataObject**, it is provided when the control calls the **GetData** method. The **SetData** and **Clear** methods cannot be used here. |
| *effect* | A long integer initially set by the source object identifying all effects it supports. This parameter must be correctly set by the target component during this event. The value of effect is determined by logically **Or**'ing together all active effects (as listed in Settings). The target component should check these effects and other parameters to determine which actions are appropriate for it, and then set this parameter to one of the allowable effects (as specified by the source) to specify which actions will be performed if the user drops the selection on the component. The possible values are listed in Settings. |
| *button* | An integer which acts as a bit field corresponding to the state of a mouse button when it is depressed. The left button is bit 0, the right button is bit 1, and the middle button is bit 2. These bits correspond to the values 1, 2, and 4, respectively. It indicates the state of the mouse buttons; some, all, or none of these three bits can be set, indicating that some, all, or none of the buttons are depressed. |
| *shift* | An integer which acts as a bit field corresponding to the state of the SHIFT, CTRL, and ALT keys when they are depressed. The SHIFT key is bit 0, the CTRL key is bit 1, and the ALT key is bit 2. These bits correspond to the values 1, 2, and 4, respectively. The shift parameter indicates the state of these keys; some, all, or none of the bits can be set, indicating that some, all, or none of the keys are depressed. For example, if both the CTRL and ALT keys were depressed, the value of shift would be 6. |
| *x,y* | A number which specifies the current location of the mouse pointer. The x and y values are always expressed in container coordinates. |
| *state* | An integer that corresponds to the transition state of the control being dragged in relation to a target form or control. The possible values are listed in Settings. |

### Settings

 The settings for *effect* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexDropEffectNone** | 0 | Drop target cannot accept the data. |
|  **jgexDropEffectCopy** | 1 | Drop results in a copy of data from the source to the target. The original data is unaltered by the drag operation. |
|  **jgexDropEffectMove** | 2 | Drop results in data being moved from drag source to drop source. The drag source should remove the data from itself after the move. |
|  **jgexDropEffectScroll** | -2147483648<br> (&H80000000) | Scrolling is occurring or about to occur in the target component. This value is used in conjunction with the other values. Note Use only if you are performing your own scrolling in the target component. |

### Settings

 The settings for *state* are:

| Value | Meaning | Description |
| --- | --- | --- |
| 0 | Enter | Source component is being dragged within the range of a target. |
| 1 | Leave | Source component is being dragged out of the range of a target. |
| 2 | Over | Source component has moved from one position in the target to another. |

### Remarks

 Note If the state parameter is 1, indicating that the mouse pointer has left the target, then the x and y parameters will contain zeros.  
 The source component should always mask values from the effect parameter to ensure compatibility with future implementations of ActiveX components. As a precaution against future problems, drag sources and drop targets should mask these values appropriately before performing any comparisons.  
 For example, a source component should not compare an effect against, say, **jgexDropEffectCopy**, such as the following code:

```vb
If Effect = jgexDropEffectCopy Then...
```

Instead, the source component should mask for the value or values being sought, such as this:

```vb
If Effect And jgexDropEffectCopy = jgexDropEffectCopy...
-or-
If (Effect And jgexDropEffectCopy)...
```

This allows for the definition of new drop effects in future versions of Visual Basic while preserving backwards compatibility with your existing code.  
  The **GridEX** control only supports manual OLE drag and drop events.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [OLECompleteDrag Event](#olecompletedrag-event-gridex-control), [OLEDragDrop Event](#oledragdrop-event-gridex-control), [OLEDrag Method](Methods.md#oledrag-method-gridex-control), [OLEDropMode Property](Properties.md#oledropmode-property-gridex-control), [OLEGiveFeedback Event](#olegivefeedback-event-gridex-control), [OLESetData Event](#olesetdata-event-gridex-control), [OLEStartDrag Event](#olestartdrag-event-gridex-control)

## OLEGiveFeedback Event (GridEX Control)

Occurs after every **OLEDragOver** event. **OLEGiveFeedback** allows the source component to provide visual feedback to the user, such as changing the mouse cursor to indicate what will happen if the user drops the object, or provide visual feedback on the selection (in the source component) to indicate what will happen.

### Syntax

 **Private Sub** *object*_**OLEGiveFeedback** ([ *index* **As Integer**,] *effect* **As Long**, *defaultcursors* **As Boolean**)  
 The **OLEGiveFeedback** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *effect* | A long integer set by the target component in the **OLEDragOver** event specifying the action to be performed if the user drops the selection on it. This allows the source to take the appropriate action (such as giving visual feedback). The possible values are listed in Settings. |
| *defaultcursors* | A Boolean value that determines whether to use the default mouse cursor, or to use a user-defined mouse cursor.True (default) = use default mouse cursor.False = do not use default cursor. Mouse cursor must be set with the MousePointer property of the Screen object. |

### Settings

 The settings for *effect* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexDropEffectNone** | 0 | Drop target cannot accept the data. |
|  **jgexDropEffectCopy** | 1 | Drop results in a copy of data from the source to the target. The original data is unaltered by the drag operation. |
|  **jgexDropEffectMove** | 2 | Drop results in data being moved from drag source to drop source. The drag source should remove the data from itself after the move. |
|  **jgexDropEffectScroll** | -2147483648<br> (&H80000000) | Scrolling is occurring or about to occur in the target component. This value is used in conjunction with the other values. Note Use only if you are performing your own scrolling in the target component. |

### Settings

 The settings for *state* are:

| Value | Meaning | Description |
| --- | --- | --- |
| 0 | Enter | Source component is being dragged within the range of a target. |
| 1 | Leave | Source component is being dragged out of the range of a target. |
| 2 | Over | Source component has moved from one position in the target to another. |

### Remarks

 If there is no code in the **OLEGiveFeedback** event, or if the defaultcursors parameter is set to **True**, the mouse cursor will be set to the default cursor provided by the **GridEX** control.  
 The source component should always mask values from the effect parameter to ensure compatibility with future implementations of ActiveX components. As a precaution against future problems, drag sources and drop targets should mask these values appropriately before performing any comparisons.  
 For example, a source component should not compare an effect against, say, **jgexDropEffectCopy**, such as the following code:

```vb
If Effect = jgexDropEffectCopy Then...
```

Instead, the source component should mask for the value or values being sought, such as this:

```vb
If Effect And jgexDropEffectCopy = jgexDropEffectCopy...
-or-
If (Effect And jgexDropEffectCopy)...
```

This allows for the definition of new drop effects in future versions of Visual Basic while preserving backwards compatibility with your existing code.  
 The **GridEX** control only supports manual OLE drag and drop events.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [OLECompleteDrag Event](#olecompletedrag-event-gridex-control), [OLEDragDrop Event](#oledragdrop-event-gridex-control), [OLEDrag Method](Methods.md#oledrag-method-gridex-control), [OLEDragOver Event](#oledragover-event-gridex-control), [OLEDropMode Property](Properties.md#oledropmode-property-gridex-control), [OLESetData Event](#olesetdata-event-gridex-control), [OLEStartDrag Event](#olestartdrag-event-gridex-control)

## OLESetData Event (GridEX Control)

Occurs on a source component when a target component performs the **GetData** method on the source’s **JSDataObject** object, but the data for the specified format has not yet been loaded.

### Syntax

 **Private Sub** *object*_**OLESetData** ([ *index* **As Integer**,] *data* **As JSDataObject**, *dataformat* **As Integer**)  
 The **OLESetData** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *data* | A **JSDataObject** object in which to place the requested data. The component calls the **SetData** method to load the requested format. |
| *dataformat* | An integer specifying the format of the data that the target component is requesting. The source component uses this value to determine what to load into the **DataObject** object. |

### Remarks

 In certain cases, you may wish to defer loading data into the **JSDataObject** object of a source component to save time, especially if the source component supports many formats or when too many rows are selected and data is obtained using the **GetClipString** method.  
 This event allows the source to respond to only one request for a given format of data.  
 When this event is called, the source should check the format parameter to determine what needs to be loaded and then perform the **SetData** method on the **JSDataObject** object to load the data which is then passed back to the target component.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [OLECompleteDrag Event](#olecompletedrag-event-gridex-control), [OLEDragDrop Event](#oledragdrop-event-gridex-control), [OLEDrag Method](Methods.md#oledrag-method-gridex-control), [OLEDragOver Event](#oledragover-event-gridex-control), [OLEDropMode Property](Properties.md#oledropmode-property-gridex-control), [OLEGiveFeedback Event](#olegivefeedback-event-gridex-control), [OLEStartDrag Event](#olestartdrag-event-gridex-control)  
**Example:** [RowDrag Example](../Examples.md#rowdrag-example)

## OLEStartDrag Event (GridEX Control)

Occurs when a **GridEX**’s **OLEDrag** method is performed.  
 This event specifies the data formats and drop effects that the source component supports. It can also be used to insert data into the **JSDataObject** object.

### Syntax

 **Private Sub** *object*_**OLEStartDrag** ([ *index* **As Integer**,] *data* **As JSDataObject**, *allowedeffects* **As Integer**)  
 The **OLEStartDrag** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *data* | A **JSDataObject** object containing formats that the source will provide and, optionally, the data for those formats. If no data is contained in the **JSDataObject**, it is provided when the control calls the **GetData** method. The programmer should provide the values for this parameter in this event. The **SetData** and **Clear** methods cannot be used here. |
| *allowedeffects* | A long containing the effects that the source component supports. The possible values are listed in Settings. The programmer should provide the values for this parameter in this event. |

### Settings

 The settings for *allowedeffects* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexDropEffectNone** | 0 | Drop target cannot accept the data. |
|  **jgexDropEffectCopy** | 1 | Drop results in a copy of data from the source to the target. The original data is unaltered by the drag operation. |
|  **jgexDropEffectMove** | 2 | Drop results in data being moved from drag source to drop source. The drag source should remove the data from itself after the move. |

### Remarks

 The source component should logically Or together the supported values and places the result in the allowedeffects parameter.  
 The target component can use this value to determine the appropriate action (and what the appropriate user feedback should be).  
 You may wish to defer putting data into the **JSDataObject** object until the target component requests it. This allows the source component to save time. When the target performs the **GetData** method on the **JSDataObject**, the source’s **OLESetData** event will occur if the requested data is not contained in the **JSDataObject**. At this point, the data can be loaded into the **JSDataObject**, which will in turn provide the data to the target.  
 If the user does not load any formats into the **JSDataObject**, then the drag/drop operation is canceled.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [OLECompleteDrag Event](#olecompletedrag-event-gridex-control), [OLEDragDrop Event](#oledragdrop-event-gridex-control), [OLEDrag Method](Methods.md#oledrag-method-gridex-control), [OLEDragOver Event](#oledragover-event-gridex-control), [OLEDropMode Property](Properties.md#oledropmode-property-gridex-control), [OLEGiveFeedback Event](#olegivefeedback-event-gridex-control), [OLESetData Event](#olesetdata-event-gridex-control)  
**Example:** [RowDrag Example](../Examples.md#rowdrag-example)

## PrintingProgress Event (GridEX Control)

Occurs continuously while the **GridEX** control is printing a document.

### Syntax

 **Private Sub** *object*_**PrintingProgress**([ *index* **As Integer**, ] **ByVal** *printprogress* **As Single**, **ByVal** *cancel* **As JSRetBoolean**)  
 The **PrintingProgress** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *printprogress* | A single in the range from 0 to 1 that indicates the portion of the document already printed. |
| *cancel* | A **JSRetBoolean** object that may be set to **True**, to cancel the printing of a document, as described in settings. |

### Settings

 The settings for *cancel* are:

| Setting | Description |
| --- | --- |
| **True** | The print operation in a **GridEX** control is canceled. |
| **False** | (Default) The **GridEX** control can continue the print operation. |

### Remarks

 The **PrintingProgress** event is fired continuously while a the **GridEX** control is printing a document.  
 You can use this event to inform the user about the progress of a printing operation or to cancel the printing operation at any time.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [BeforePrintPage Event](#beforeprintpage-event-gridex-control), [PrintingProgress Event](#printingprogress-event-gridex-control), [PrinterProperties Property](Properties.md#printerproperties-property-gridex-control), [PrintGrid Method](Methods.md#printgrid-method-gridex-control), [JSPrinterProperties](Objects/JSPrinterProperties-Object.md#jsprinterproperties), [BeforePrinting Event](#beforeprinting-event-gridex-control)  
**Example:** [PrintingProgress Example](../Examples.md#printingprogress-example)

## RowColChange Event (GridEX Control)

Occurs when the current cell changes to a different cell in the same row on a different cell in another row.

### Syntax

 **Private Sub** *object*_**RowColChange**( [ *index* **As Integer**, ] **ByVal** *lastrow* **As Long**, **ByVal** *lastcol* **As Integer**)  
 The **RowColChange** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that uniquely identifies a control if it is in a control array. |
| *lastrow* | A long specifying the previous row position. |
| *lastcol* | An integer specifying the previous column position. |

### Remarks

 This event occurs whenever the user clicks a cell other than the current cell, uses the keyboard to navigate to another cell, or when you programmatically change the current cell using the **Col** and **Row** properties.  
 The previous cell position is specified by lastrow and lastcol.  
 If you edit data and then move the current cell position to a different row, the update events for the original row are completed before another cell becomes the current cell.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [Row Property](Properties.md#row-property-gridex-control), [Col Property](Properties.md#col-property-gridex-control)  
**Example:** [EditMode Example](../Examples.md#editmode-example)

## RowDrag Event (GridEX Control)

Occurs when the user attempts to drag a row or rows to begin an OLE drag operation.  
 This event is only triggered when a **GridEX** control’s **DetectRowDrag** property is set to **True**.

### Syntax

 **Private Sub** *object*_**RowDrag** ([ *index* **As Integer**, ] **ByVal** *button* **As Integer**, **ByVal** *shift* **As Integer**)  
 The **RowDrag** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *button* | An integer that corresponds to the state of the mouse buttons in which a bit is set if the button is down. |
| *shift* | An integer that corresponds to the state of the SHIFT, CTRL, and ALT keys. |

### Remarks

 The event is triggered only in controls with its **DetectRowDrag** property is equal to **True**.  
 Use this event to start an OLE drag operation instead of simply using the **MouseDown** event.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [DetectRowDrag Property](Properties.md#detectrowdrag-property-gridex-control), [OLEDrag Method](Methods.md#oledrag-method-gridex-control)  
**Example:** [RowDrag Example](../Examples.md#rowdrag-example)

## RowFormat Event (GridEX Control)

Occurs when The **GridEX** control loads a row.

### Syntax

 **Private Sub** *object*_**RowFormat** ( [ *index* **As Integer**, ] *rowbuffer* **As JSRowData**)  
 The **RowFormat** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *rowbuffer* | A **JSRowData** object that contains data and format information about a row. |

### Remarks

 Use this event to specify row and/or cell based format styles in **GridEX** control's cells.  
 For group rows, you can change the group caption to include sub totals for the group.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**Example:** [RowFormat Example](../Examples.md#rowformat-example)

## RowResize Event (GridEX Control)

Occurs when the user changes the **RowHeight** property by dragging the bar drawn between each row header.

### Syntax

 **Private Sub** *object*_**RowResize**([ *index* **As Integer**,] **ByVal** *newrowheight* **As Long**, **ByVal** *cancel* **As JSRetBoolean**)  
 The **RowResize** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *newrowheight* | A Long expression that represents the new **RowHeight**. |
| *cancel* | A **JSRetBoolean** object that determines whether the column is resized, as described in Settings. |

### Settings

 The settings for *cancel* are:

| Setting | Description |
| --- | --- |
| **True** | Cancels the change and **RowHeight** is restored. |
| **False** | (Default) Continues the height change operation. |

### Remarks

 This event is only triggered if **AllowRowSizing** and **RowHeaders** property are both **True**.  
 The **RowResize** event is triggered when the user is about to change the **RowHeight** property in a GridEX control. Your event procedure can accept, cancel, or alter the amount of change.  
 If you set the *cancel* argument to **True**, the **RowHeight** is restored. To alter the degree of change, set the **RowHeight** property to the desired value.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [RowHeight Property](Properties.md#rowheight-property-gridex-control)

## SelectionChange Event (GridEX Control)

Occurs when the user selects a different range of rows.

### Syntax

 **Private Sub** *object*_**SelectionChange**([ *index* **As Integer** ])  
 The **SelectionChange** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that uniquely identifies a control if it's in a control array |

### Remarks

 The selection change event is raised whenever a row selected state changes.  
 In a multiselect **GridEX** control, the event informs you when a row has been added to or removed from the **SelectedItems** collection.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [MultiSelect Property](Properties.md#multiselect-property-gridex-control), [JSSelectedItems Collection](Objects/JSSelectedItems-Collection.md#jsselecteditems-collection), [RowSelected Property](Properties.md#rowselected-property-gridex-control), [SelectedItems Property](Properties.md#selecteditems-property-gridex-control)  
**Example:** [SelectionChange Example](../Examples.md#selectionchange-example)

## ShowCustomEdit Event (GridEX Control)

Occurs when a custom edit column is about to hide or show the cell editor.  
 This event is only raised for columns whose **EditType** is set to **jgexEditCustom**.

### Syntax

 **Private Sub** *object*_**ShowCustomEdit** ([ *index* **As Integer**, ] **ByVal** *colindex* **As Integer**, **ByVal** *editleft* **As Single**, **ByVal** *edittop* **As Single**, **ByVal** *editwidth* **As Single**, **ByVal** *editheight* **As Single**, **ByVal** *editvisible* **As Boolean**)  
 The **ShowCustomEdit** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *colindex* | An integer that identifies the index of the custom edit column that is about to end the edit operation. |
| *editleft, edittop* | A number that specifies the suggested location of the custom cell editor. The editleft and edittop values are always expressed in twips and offset from the upper left corner of the control. |
| *editwidth, editheight* | A number that specifies the suggested size of the custom cell editor. The editwidth and editheight values are always expressed in twips. |
| *editvisible* | A Boolean expression that determines if the custom cell editor must be visible or hidden. |

### Remarks

 In this event you must place and size the control being used for edition in a custom edit column.  
 To place a control properly you must add the **Left** and **Top** properties of the GridEX control to the values of the *editleft* and *edittop* arguments, respectively.  
 This event is triggered only for cells whose **EditType** property is set to **jgexEditcustom**.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [EditType Property](Objects/JSColumn-Object.md#edittype-property-jscolumn-object), [InitCustomEdit Event](#initcustomedit-event-gridex-control), [EndCustomEdit Event](#endcustomedit-event-gridex-control)  
**Example:** [CustomEdit Example](../Examples.md#customedit-example)

## UnboundAddNew Event (GridEX Control)

Occurs in an unbound **GridEX** control when a new row is added to it.  
 This event alerts your application that a new row of data must be added.

### Syntax

 **Private Sub** *object*_**UnboundAddNew**( [ *index* **As Integer**,] **ByVal** *newrowbookmark* **As JSRetVariant**, **ByVal** *values* **As JSRowData**)  
 The **UnboundAddNew** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it’s in a control array. |
| *newrowbookmark* | A **JSRetVariant** object that can be set to the bookmark for the new row. |
| *varvalues* | A **JSRowData** that holds the values entered by the users in the new record. |

### Remarks

 A new row of data can be added only if the **AllowAddNew** property is **True**.  
 When the **UnboundAddNew** event is fired the **UnboundUpdate** event will not be triggered.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [UnboundDelete Event](#unbounddelete-event-gridex-control), [UnboundReadData Event](#unboundreaddata-event-gridex-control), [UnboundUpdate Event](#unboundupdate-event-gridex-control), [DataMode Property](Properties.md#datamode-property-gridex-control), [JSRowData Object](Objects/JSRowData-Object.md#jsrowdata-object), [AllowAddNew Property](Properties.md#allowaddnew-property-gridex-control)  
**Example:** [UnboundEvents Example](../Examples.md#unboundevents-example)

## UnboundDelete Event (GridEX Control)

Occurs whenever a row of data is deleted from the unbound **GridEX**.  
 This event alerts your application that a row of data must be deleted.

### Syntax

 **Private Sub** *object*_**UnboundDelete**([ *index* **As Integer**,] **ByVal** *rowindex* **As Long**, **ByVal** *bookmark* **As Variant**)  
 The **UnboundDelete** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it’s in a control array. |
| *rowindex* | A long representing the index of the row to be deleted. |
| *bookmark* | A value representing the bookmark of the row to be deleted. |

### Remarks

 This event is fired to alert your application to that it must delete the specified row.  
 This event cannot be canceled. However, if you want to cancel the deletion of a particular record you can do it in the **BeforeDeleteEx** event.  
 The *rowindex* parameter is the same as the value of the **RowIndex** property of the **GridEX** control as follows:

```vb
Private Sub GridEX1_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As variant)

Dim lngIndex as Long
Dim lngRow as Long

lngRow = GridEX1.Row
lngIndex = GridEX1.RowIndex(lngRow)
Debug.Print lngIndex = RowIndex 'Prints "True"
End sub
```

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [AllowDelete Property](Properties.md#allowdelete-property-gridex-control), [UnboundReadData Event](#unboundreaddata-event-gridex-control), [UnboundUpdate Event](#unboundupdate-event-gridex-control), [DataMode Property](Properties.md#datamode-property-gridex-control), [BeforeDeleteEx Event](#beforedeleteex-event-gridex-control), [BeforeDelete Event](#beforedelete-event-gridex-control), [UnboundAddNew Event](#unboundaddnew-event-gridex-control)  
**Example:** [UnboundEvents Example](../Examples.md#unboundevents-example)

## UnboundReadData Event (GridEX Control)

Occurs whenever an unbound **GridEX** control requires data for display, sorting or grouping.

### Syntax

 **Private Sub** *object*_**UnboundReadData**( [ *index* **As Integer**, ] **ByVal** *rowindex* **As Long**, **ByVal** *bookmark* **As Variant**, **ByVal** *values* **As JSRowData**)  
 The **UnboundReadData** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it’s in a control array. |
| *rowindex* | A long that identifies the original index of the row being retrieved. |
| *bookmark* | A variant bookmark which specifies the bookmark set to row using the **RowBookmark** property |
| *values* | A **JSRowData** where the application must set the values corresponding to the row being fetched. <br> <br> The bounds of the value property are 1 to the value of the **ColCount** property. |

### Remarks

 This event retrieves all the column values for any particular row.  
 The row position is irrelevant for this event.  
 The *rowindex* parameter represents the original index of the row. As in rows, the position of columns doesn’t matter either.  
 For example, to set the value of the third column in the columns collection you may write code as follows:

```vb
Values(3) = "Column 3"
```

The *bookmark* parameter is provided as a reference. You can not change the bookmark in this event; to do so you must use **RowBookmark** property.  
 The *rowindex* parameter is in the range of 1 to the value of the **ItemCount** property.

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [UnboundDelete Event](#unbounddelete-event-gridex-control), [RowBookmark Property](Properties.md#rowbookmark-property-gridex-control), [UnboundUpdate Event](#unboundupdate-event-gridex-control), [DataMode Property](Properties.md#datamode-property-gridex-control), [JSRowData Object](Objects/JSRowData-Object.md#jsrowdata-object), [ItemCount Property](Properties.md#itemcount-property-gridex-control), [UnboundAddNew Event](#unboundaddnew-event-gridex-control)  
**Example:** [UnboundEvents Example](../Examples.md#unboundevents-example)

## UnboundUpdate Event (GridEX Control)

Occurs when an unbound **GridEX** control has an entire row of modified data to be written in the data set.  
 It alerts your application that it must update an edited row of data.

### Syntax

 **Private Sub** *object*_**UnboundUpdate**( [ *index* **As Integer**, ] **ByVal** *rowindex* **As Long**, **ByVal** *bookmark* **As Variant**, **ByVal** *values* **As JSRowData**)  
 The **UnboundUpdate** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it’s in a control array. |
| *rowindex* | A long that identifies the original index of the row being updated. |
| *bookmark* | A bookmark which acts as a unique identifier for each row of data. |
| *values* | A **JSRowData** object that holds the values entered by the users in the new record. |

### Remarks

 You cannot cancel this event. If you want to cancel an update, even in unbound mode, you must do so in the **BeforeUpdate** event.  
 If the row that is to be updated is a new record, the **UnboundAddNew** event will be fired instead.  
 If you don’t use bookmarks but you need the index of the row to be updated you could write code as follows:

```vb
Dim RowIndex As Long

RowIndex = GridEX1.RowIndex(GridEX1.Row)
Debug.print RowIndex
```

**Applies To:** [GridEX Control](../GridEX-Control.md#gridex-control)  
**See Also:** [UnboundDelete Event](#unbounddelete-event-gridex-control), [UnboundReadData Event](#unboundreaddata-event-gridex-control), [BeforeUpdate Event](#beforeupdate-event-gridex-control), [DataMode Property](Properties.md#datamode-property-gridex-control), [JSRowData Object](Objects/JSRowData-Object.md#jsrowdata-object), [Update Method](Methods.md#update-method-gridex-control), [UnboundAddNew Event](#unboundaddnew-event-gridex-control)  
**Example:** [UnboundEvents Example](../Examples.md#unboundevents-example)
