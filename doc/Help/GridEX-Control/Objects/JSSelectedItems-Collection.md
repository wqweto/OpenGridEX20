# JSSelectedItems Collection

## JSSelectedItems Collection

![](../../images/JSSelectedItems.jpg)  
 Represents a collection of selected rows in a **GridEX** control.

### Syntax

 *object*.**SelectedItems**  
 The *object* placeholder represents an object expression that evaluates to a **GridEX** control.

### Remarks

 With the **JSSelectedItems** collection you can add and remove selected rows, enumerate or count **JSSelectedItem** objects, and access individual **JSSelectedItem** objects.  
 The **JSSelectedItems** collection can be accessed through the **SelectedItems** property of the **GridEX** control.  
 To select a row you can use any of the three different methods provided by this collection: **Add**, **AddBookmark** or **AddRowIndex**.  
 To remove a selected item form the collection you may use one of four different methods: **Remove**, **RemoveRowIndex**, **RemoveBookmark**, **RemoveRowPosition**.  
 Determining which of the methods to use depends on the criteria that you want the control to use when searching for and selecting (or removing) individual records.

- [JSSelectedItem Object](JSSelectedItem-Object.md#jsselecteditem-object)

**See Also:** [JSSelectedItem Object](JSSelectedItem-Object.md#jsselecteditem-object), [SelectedItems Property](../Properties.md#selecteditems-property-gridex-control), [Item Property](#item-property-jsselecteditems-collection)

## Count Property (JSSelectedItems Collection)

Returns the number of objects in a collection.

### Syntax

 *object*.**Count**  
 The object placeholder is an object expression that evaluates to an object in the Applies To list.

### Remarks

 You can use this property with a **For...Next** statement to carry out operations on objects in a collection.  
 For example, the following code prints in the immediate window the **RowPosition** property of all the **JSSelectedItem** objects in the collection.

```vb
For I = 1 To GridEX1.SelectedItems.Count
Debug.Print "SelectedItems(" & I & ") =" & _
GridEX1.SelectedItems(I).RowPosition
Next I
```

### Data Type

 Long

**Note** Since collections are 1-based, there is no need to use the *Collections.Count-1*counter expression in **For…Next** loops.

**Applies To:** [JSSelectedItems Collection](#jsselecteditems-collection)  
**See Also:** [Item Property](#item-property-jsselecteditems-collection), [RowIndex Property](JSSelectedItem-Object.md#rowindex-property-jsselecteditem-object)

## Item Property (JSSelectedItems Collection)

Returns a specific **JSSelectedItem** object in the **JSSelectedItems** collection.

### Syntax

 *object*.**Item(***index***)**  
 The **Item** property syntax has the following parts:

| Part | Description |
| --- | --- |
| *object* | Required. An object expression that evaluates to an object in the Applies To list. |
| *index* | Required. A long that specifies the index of a JSSelectedItem in the collection. |

### Remarks

 If the value provided as index is larger than the Count property, an error occurs.

### Data Type

 **JSSelectedItem**

**Applies To:** [JSSelectedItems Collection](#jsselecteditems-collection)  
**See Also:** [Count Property](#count-property-jsselecteditems-collection), [Remove Method](#remove-method-jsselecteditems-collection)

## Add Method (JSSelectedItems Collection)

Selects a row in a **GridEX** control given its current position. The row then becomes part of the **JSSelectedItems** collection.

### Syntax

 *object*.**Add** *rowposition*  
 The **Add** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | Required. An object expression that evaluates to an object in the Applies To list. |
| *rowposition* | A Long specifying the position of the row to be selected. |

### Remarks

 Use this method to select a row when you know its current position.  
 This method is equivalent to using the **RowSelected** property in a **GridEX** control.  
 To select a row in a multi-select **GridEX** control you can also use the **AddBookmark** or **AddRowIndex** methods.

**Applies To:** [JSSelectedItems Collection](#jsselecteditems-collection)  
**See Also:** [AddRowIndex Method](#addrowindex-method-jsselecteditems-collection), [MultiSelect Property](../Properties.md#multiselect-property-gridex-control), [RowSelected Property](../Properties.md#rowselected-property-gridex-control), [AddBookmark Method](#addbookmark-method-jsselecteditems-collection)

## AddBookmark Method (JSSelectedItems Collection)

Selects a row in a **GridEX** control given its bookmark. The row then becomes part of the **JSSelectedItems** collection.  
 Syntax  
 *object*.**AddBookmark** *bookmark*  
 The **AddBookmark** method syntax has these parts:

**Part** **Description**

*object* Required. An object expression that evaluates to an object in the Applies To list.

*bookmark* A variant specifying the bookmark of the record to be selected.

### Remarks

 Use this method to select a row when you know its bookmark. If the record to be selected is in a collapsed group, the group will be expanded automatically and the row selected.  
 To select a row in a multi-select **GridEX** control you can also use the **Add** or **AddRowIndex** methods.

**Applies To:** [JSSelectedItems Collection](#jsselecteditems-collection)  
**See Also:** [Add Method](#add-method-jsselecteditems-collection), [AddRowIndex Method](#addrowindex-method-jsselecteditems-collection), [MultiSelect Property](../Properties.md#multiselect-property-gridex-control), [RowSelected Property](../Properties.md#rowselected-property-gridex-control)

## AddRowIndex Method (JSSelectedItems Collection)

Selects a row in a **GridEX** control given its original row index. The row then becomes part of the **JSSelectedItems** collection.

### Syntax

 *object*.**AddRowIndex** *rowindex*  
 The **AddRowIndex** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | Required. An object expression that evaluates to an object in the Applies To list. |
| *rowindex* | A Long specifying the original row index of the row to be selected. |

### Remarks

 Use this method to select a row when you know its row index. If the record to be selected is in a collapsed group, the group will be expanded and the row selected.  
 To select a row in a multi-select **GridEX** control you can also use the **AddBookmark** or **AddRowIndex** methods.  
 This method is especially useful when using a **GridEX** control in unbound mode without bookmarks.

**Applies To:** [JSSelectedItems Collection](#jsselecteditems-collection)  
**See Also:** [Add Method](#add-method-jsselecteditems-collection), [AddBookmark Method](#addbookmark-method-jsselecteditems-collection), [MultiSelect Property](../Properties.md#multiselect-property-gridex-control), [RowSelected Property](../Properties.md#rowselected-property-gridex-control)

## Clear Method (JSSelectedItems Collection)

Removes all objects in a collection.

### Syntax

 *object*.**Clear**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 To remove only one object from a collection, use the **Remove** method of the collection.

**Applies To:** [JSSelectedItems Collection](#jsselecteditems-collection)  
**See Also:** [Remove Method](#remove-method-jsselecteditems-collection)

## Remove Method (JSSelectedItems Collection)

Removes an item from the **JSSelectedItems** collection in a **GridEX** control given its positional index within the collection.  
 The row is deselected and removed from the collection.

### Syntax

 *object*.**Remove** *index*  
 The **Remove** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | A long specifying the index of the item in the collection to be removed. |

### Remarks

 To remove a selected item from the JSSelectedItems collection you can also use the **RemoveRowPosition**, **RemoveBookmark** or **RemoveRowIndex** methods.

**Applies To:** [GridEX Control](../../GridEX-Control.md#gridex-control)  
**See Also:** [Item Property](#item-property-jsselecteditems-collection), [RowIndex Property](JSSelectedItem-Object.md#rowindex-property-jsselecteditem-object), [Clear Method](#clear-method-jsselecteditems-collection), [RemoveRowIndex Method](#removerowindex-method-jsselecteditems-collection), [RemoveRowPosition Method](#removerowposition-method-jsselecteditems-collection), [RemoveBookmark Method](#removebookmark-method-jsselecteditems-collection)

## RemoveBookmark Method (JSSelectedItems Collection)

Removes an item from the **JSSelectedItems** collection in a **GridEX** control given its bookmark.  
 The row is deselected and removed from the collection.

### Syntax

 *object*.**RemoveBookmark** *bookmark*  
 The **RemoveBookmark** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *bookmark* | A variant specifying the bookmark of the record to be deselected. |

### Remarks

 To remove a selected item from the **JSSelectedItems** collection you can also use the **RemoveRowPosition**, **Remove** or **RemoveRowIndex** methods.

**Applies To:** [JSSelectedItems Collection](#jsselecteditems-collection)  
**See Also:** [RemoveRowIndex Method](#removerowindex-method-jsselecteditems-collection), [RemoveRowPosition Method](#removerowposition-method-jsselecteditems-collection), [Remove Method](#remove-method-jsselecteditems-collection), [Clear Method](#clear-method-jsselecteditems-collection), [RowSelected Property](../Properties.md#rowselected-property-gridex-control), [SelectionChange Event](../Events.md#selectionchange-event-gridex-control)

## RemoveRowIndex Method (JSSelectedItems Collection)

Removes an item from the **JSSelectedItems** collection in a **GridEX** control given its original row index.  
 The row is deselected and removed from the collection.

### Syntax

 *object*.**RemoveRowIndex** *rowindex*  
 The **RemoveRowIndex** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *rowindex* | A Long specifying the original row index of the row to be deselected. |

### Remarks

 To remove a selected item from the **JSSelectedItems** collection you can also use the **RemoveRowPosition**, **RemoveBookmark** or **Remove** methods.

**Applies To:** [JSSelectedItems Collection](#jsselecteditems-collection)  
**See Also:** [RemoveRowPosition Method](#removerowposition-method-jsselecteditems-collection), [RemoveBookmark Method](#removebookmark-method-jsselecteditems-collection), [Remove Method](#remove-method-jsselecteditems-collection), [Clear Method](#clear-method-jsselecteditems-collection), [RowSelected Property](../Properties.md#rowselected-property-gridex-control), [SelectionChange Event](../Events.md#selectionchange-event-gridex-control)

## RemoveRowPosition Method (JSSelectedItems Collection)

Removes an item from the **JSSelectedItems** collection in a **GridEX** control given its current row position.  
 The row is deselected and removed from the collection.

### Syntax

 *object*.**RemoveRowPosition** *rowposition*  
 The **RemoveRowPosition** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *rowposition* | A Long specifying the current position of the row to be deselected. |

### Remarks

 To remove a selected item from the **JSSelectedItems** collection you can also use the **RemoveRowPosition**, **RemoveBookmark** or **Remove** methods.

**Applies To:** [JSSelectedItems Collection](#jsselecteditems-collection)  
**See Also:** [RemoveRowIndex Method](#removerowindex-method-jsselecteditems-collection), [RemoveBookmark Method](#removebookmark-method-jsselecteditems-collection), [Remove Method](#remove-method-jsselecteditems-collection), [Clear Method](#clear-method-jsselecteditems-collection), [RowSelected Property](../Properties.md#rowselected-property-gridex-control), [SelectionChange Event](../Events.md#selectionchange-event-gridex-control)
