# JSSortKey Object

## JSSortKey Object

![](../../images/JSSortKey.jpg)  
 Contains information for sorting rows within a **GridEX** control.

### Syntax

 *object*.**SortKeys**(*index*)  
 The **JSSortKey** object syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to a **GridEX** control. |
| *index* | An integer that represents the value of the **Index** property. |

### Remarks

 With a **JSSortKey** object, you can modify the sort settings for a **GridEX** control.  
 To use a **JSSortKey** object, you can use the **SortKeys** property of the **GridEX** control. You can also assign a **JSSortKey** to a separate variable dimensioned as **JSSortKey**.  
 The following code illustrates both methods:

```vb
Dim stkTemp as JSSortKey

Set stkTemp = GridEX1.SortKeys(1)
Debug.Print (stkTemp Is GridEX1.SortKeys(1))

'Prints True
```

**See Also:** [SortKeys Property](../Properties.md#sortkeys-property-gridex-control), [JSSortKeys Collection](JSSortKeys-Collection.md#jssortkeys-collection), [Item Property](JSSortKeys-Collection.md#item-property-jssortkeys-collection)

## ColIndex Property (JSSortKey Object)

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

**See Also:** [Index Property](JSColumn-Object.md#index-property-jscolumn-object), [JSColumn Object](JSColumn-Object.md#jscolumn-object)

## Index Property (JSSortKey Object)

Returns a value that represents the index of an object in a collection. Read only.

### Syntax

 *object*.**Index**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The Index property for all collections is 1-based.

### Data Type

 Integer

**Applies To:** [JSSortKey Object](#jssortkey-object)  
**See Also:** [Count Property](JSSortKeys-Collection.md#count-property-jssortkeys-collection), [Item Property](JSSortKeys-Collection.md#item-property-jssortkeys-collection), [Remove Method](JSSortKeys-Collection.md#remove-method-jssortkeys-collection)

## SortOrder Property (JSSortKey Object)

Returns a value that represents the sort order of an object in a **GridEX** control. This property is Read-Write for **JSSortKey** object.

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

 An error occurs if you try to set the **SortOrder** property to **jgexSortNone** for a **JSSortKey** object.

### Data Type

 **jgexSortOrderConstants**

**Applies To:** [JSSortKey Object](#jssortkey-object)  
**See Also:** [SortOrder Property](JSGroup-Object.md#sortorder-property-jsgroup-object), [SortType Property](JSColumn-Object.md#sorttype-property-jscolumn-object), [SortOrder Property](JSColumn-Object.md#sortorder-property-jscolumn-object)  
**Example:** [ColumnHeaderClick Example](../../Examples.md#columnheaderclick-example)
