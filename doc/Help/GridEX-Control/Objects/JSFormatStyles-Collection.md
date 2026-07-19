# JSFormatStyles Collection

## JSFormatStyles Collection

![](../../images/JSFormatStyles.jpg)  
 Contains a collection of **JSFmtStyle** objects.

### Syntax

 *gridex*.**FormatStyles**  
 The *gridex* placeholder represents an object expression that evaluates to a **GridEX** control.

### Remarks

 With the **JSFormatStyles** collection you can add and remove **JSFormatStyle** objects, count the number of **JSFormatStyle** objects, and address individual **JSFormatStyle** objects.  
 The **JSFormatStyles** collection can be accessed through the **FormatStyles** property of the **GridEX** control.  
 To get a specific **JSFormatStyle** object in the collection you can use the **Item** property.

- [JSFormatStyle Object](JSFormatStyle-Object.md#jsformatstyle-object)

**See Also:** [FormatStyles Property](../Properties.md#formatstyles-property-gridex-control), [Item Property](#item-property-jsformatstyles-collection), [JSFormatStyle Object](JSFormatStyle-Object.md#jsformatstyle-object)

## Count Property (JSFormatStyles Collection)

Returns the number of objects in a collection.

### Syntax

 *object*.**Count**  
 The object placeholder is an object expression that evaluates to an object in the Applies To list.

### Remarks

 You can use this property with a **For...Next** statement to carry out operations on objects in a collection.  
 For example, the following code print in the immediate window the names of all the **JSFormatStyle** objects in the collection.

```vb
For I = 1 To GridEX1.FormatStyles.Count
Debug.Print "FormatStyles(" & I & ")", _
GridEX1.FormatStyles(I).Name
Next I
```

### Data Type

 Long

**Applies To:** [JSFormatStyles Collection](#jsformatstyles-collection)  
**See Also:** [Item Property](#item-property-jsformatstyles-collection), [Name Property](JSFormatStyle-Object.md#name-property-jsformatstyle-object)

## Item Property (JSFormatStyles Collection)

Returns a specific **JSFormatStyle** of the **JSFormatStyles** collection either by index or by name.

### Syntax

 *object*.**Item**( *index* )  
 The **Item** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | Required. An object expression that evaluates to an object in the Applies To list. |
| *index* | Required. An expression that specifies the position of a member of the collection. <br> If a numeric expression, index must be a number from 1 to the value of the **Count** property. <br> If a string expression, index must correspond to the **Name** property of the member. |

### Remarks

 If the value specified in the index argument does not match any existing member of the collection, an error occurs.  
 Item is the default property for a collection. Therefore, the following lines of code are equivalent:

```vb
Debug.Print GridEX1.FormatStyles(1).BackColor
Debug.Print GridEX1.FormatStyles.Item(1).BackColor
```

**Data Type  
 JSFormatStyle**

**Applies To:** [JSFormatStyles Collection](#jsformatstyles-collection)  
**See Also:** [Count Property](#count-property-jsformatstyles-collection), [Remove Method](#remove-method-jsformatstyles-collection), [Name Property](JSFormatStyle-Object.md#name-property-jsformatstyle-object)

## Add Method (JSFormatStyles Collection)

Adds a **JSFormatStyle** object to the collection and returns a reference to the newly created object.

### Syntax

 *object*.**Add** *name*  
 The **Add** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | Required. An object expression that evaluates to an object in the Applies To list. |
| *name* | Required. A unique string that identifies the **JSFormatStyle** object. Use this value to retrieve a specific **JSFormatStyle** object. |

### Remarks

 The name of a predefined style cannot be used as a name for a new **JSFormatStyle**.  
 The predefined styles are: Default, OddRow, EvenRow, RowGroup and PreviewRow  
 Use the **Add** method to add **JSFormatStyle** objects as in the following code:

```vb
Dim fstStyle as JSFormatStyle

Set fstStyle = GridEX1.FormatStyles.Add("Style1")
```

### Data Type

 **JSFormatStyle**

**Applies To:** [JSFormatStyles Collection](#jsformatstyles-collection)  
**See Also:** [Name Property](JSFormatStyle-Object.md#name-property-jsformatstyle-object), [JSFormatStyle Object](JSFormatStyle-Object.md#jsformatstyle-object), [RowStyle Property](JSRowData-Object.md#rowstyle-property-jsrowdata-object), [CellStyle Property](JSRowData-Object.md#cellstyle-property-jsrowdata-object), [CellStyle Property](JSColumn-Object.md#cellstyle-property-jscolumn-object), [HeaderStyle Property](JSColumn-Object.md#headerstyle-property-jscolumn-object)  
**Example:** [FormatStyles Example](../../Examples.md#formatstyles-example)

## Clear Method (JSFormatStyles Collection)

Removes all the **JSFormatStyle** objects from the **JSFormatStyles** collection.

### Syntax

 *object*.**Clear**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 When **Clear** method is called, predefined styles are not removed from the collection.  
 The predefined styles are: Default, OddRow, EvenRow, RowGroup and PreviewRow  
 To remove only one object from the collection, use the **Remove** method.

**Applies To:** [JSFormatStyles Collection](#jsformatstyles-collection)  
**See Also:** [Remove Method](#remove-method-jsformatstyles-collection)

## Remove Method (JSFormatStyles Collection)

Removes a specific member from the **JSFormatStyles** collection.

### Syntax

 *object*.**Remove** *index*  
 The **Remove** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer or string that uniquely identifies the **JSFormatStyle** within the collection. <br> Use an integer to specify the value of the **Index** property; use a string to specify the value of the **Name** property. |

### Remarks

 To remove all the members of a collection, use the **Clear** method instead.  
 Any attempt to remove a predefined format style like the 'Default' one, will result in an error.  
 The predefined styles are: Default, OddRow, EvenRow, RowGroup and PreviewRow

**Applies To:** [JSFormatStyles Collection](#jsformatstyles-collection)  
**See Also:** [Item Property](#item-property-jsformatstyles-collection), [Name Property](JSFormatStyle-Object.md#name-property-jsformatstyle-object), [Clear Method](#clear-method-jsformatstyles-collection)
