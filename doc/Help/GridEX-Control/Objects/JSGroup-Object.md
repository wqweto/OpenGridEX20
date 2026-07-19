# JSGroup Object

## JSGroup Object

![](../../images/JSGroup.jpg)  
 Contains information for grouping rows in a **GridEX** control.

### Syntax

 *object*.**Groups** (*index*)  
 The **JSGroup** object syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to a **GridEX** control. |
| *index* | An integer that represents the value of the **Index** property. |

### Remarks

 With a **JSGroup** object, you can modify the group settings for a **GridEX** control.  
 To use a **JSGroup** object, you can use the **Groups** property of the **GridEX** control. You can also assign a **JSGroup** to a separate variable dimensioned as **JSGroup**.  
 The following code illustrates both methods:

```vb
Dim grpTemp as JSGroup

Set grpTemp = GridEX1.Groups(1)
Debug.Print (grpTemp Is GridEX1.Groups(1))
'Prints True
```

**See Also:** [Groups Property](../Properties.md#groups-property-gridex-control), [JSGroups Collection](JSGroups-Collection.md#jsgroups-collection), [Item Property](JSGroups-Collection.md#item-property-jsgroups-collection)

## ColIndex Property (JSGroup Object)

Returns or sets the index of the column referenced by an object.

### Syntax

 *object*.**ColIndex** [ = *value*]  
 The **ColIndex** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | An integer that refers to the index of the column to which the object is attached. |

**Remarks**:  
 If you set the **ColIndex** property to a value that does not represent a **JSColumn** in the **JSColumns** collection, an error occurs.  
 Changing the **ColIndex** property, in a **JSGroup** or **JSSortKey** object, forces the recalculation of group rows and/or sort positions for all records.

### Data Type

 Integer

**Applies To:** [JSGroup Object](#jsgroup-object)  
**See Also:** [Index Property](JSColumn-Object.md#index-property-jscolumn-object), [JSColumn Object](JSColumn-Object.md#jscolumn-object)

## Index Property (JSGroup Object)

Returns a value that represents the index of an object in a collection. Read only.

### Syntax

 *object*.**Index**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The Index property for all collections is 1-based.

### Data Type

 Integer

**Applies To:** [JSGroup Object](#jsgroup-object)  
**See Also:** [Count Property](JSGroups-Collection.md#count-property-jsgroups-collection), [Item Property](JSGroups-Collection.md#item-property-jsgroups-collection), [Remove Method](JSGroups-Collection.md#remove-method-jsgroups-collection)

## SortOrder Property (JSGroup Object)

Returns a value that represents the sort order of an object in a **GridEX** control. This property is Read-Write for **JSGroup** object.

### Syntax

 *object*.**SortOrder** [ = *value* ]  
 The **SortOrder** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A value or constant that represents the sort order of an object as described in settings. |

### Settings

 The settings for *value* are:

| Constant | Value | Description |
| --- | --- | --- |
|  **jgexSortNone** | 0 | The column is not sorted. |
|  **jgexSortAscending** | 1 | The column is sorted in ascending order. |
|  **jgexSortDescending** | -1 | The column is sorted in descending order. |

### Remarks

 An error occurs if you try to set the **SortOrder** property to **jgexSortNone** for a **JSGroup** object.

### Data Type

 **jgexSortOrderConstants**

**Applies To:** [JSGroup Object](#jsgroup-object)  
**See Also:** [SortOrder Property](JSSortKey-Object.md#sortorder-property-jssortkey-object), [SortType Property](JSColumn-Object.md#sorttype-property-jscolumn-object), [SortOrder Property](JSColumn-Object.md#sortorder-property-jscolumn-object)  
**Example:** [ColumnHeaderClick Example](../../Examples.md#columnheaderclick-example)
