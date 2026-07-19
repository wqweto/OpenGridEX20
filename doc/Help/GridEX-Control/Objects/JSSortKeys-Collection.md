# JSSortKeys Collection

## JSSortKeys Collection

![](../../images/JSSortKeys.jpg)  
 Contains a collection of **JSSortKey** objects.

### Syntax

 *gridex*.**SortKeys**  
 The *gridex* placeholder represents an object expression that evaluates to a **GridEX** control.

### Remarks

 With the **JSSortKeys** collection you can add and remove **JSSortKey** objects, enumerate or count **JSSortKey** objects, and access individual **JSSortKey** objects.  
 The **JSSortKeys** collection can be accessed through the **SortKeys** property of the **GridEX** control.  
 To obtain a specific **JSSortKey** object in the collection you can use the **Item** property.

**Note:** The **JSSortKeys** collection can contain only four items at any given time. If you try to add more than four items, an error occurs.

- [JSSortKey Object](JSSortKey-Object.md#jssortkey-object)

**See Also:** [SortKeys Property](../Properties.md#sortkeys-property-gridex-control), [JSSortKey Object](JSSortKey-Object.md#jssortkey-object), [Item Property](#item-property-jssortkeys-collection)

## Count Property (JSSortKeys Collection)

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

**Applies To:** [JSSortKeys Collection](#jssortkeys-collection)  
**See Also:** [Item Property](#item-property-jssortkeys-collection), [Index Property](JSSortKey-Object.md#index-property-jssortkey-object)

## Item Property (JSSortKeys Collection)

Returns a specific **JSSortKey** of the **JSSortKeys** collection.

### Syntax

 *object*.**Item**(*index*)  
 The **Item** property syntax has the following parts:

| Part | Description |
| --- | --- |
| *object* | Required. An object expression that evaluates to an object in the Applies To list. |
| *index* | Required. An integer that specifies the position of a **JSSortKey** in the collection. |

### Remarks

 If the value provided as index does not match any existing member of the collection, an error occurs.  
 **Item** is the default property for a collection. Therefore, the following lines of code are equivalent:

```vb
Debug.Print GridEX1.SortKeys(1).ColIndex

Debug.Print GridEX1.SortKeys.Item(1).ColIndex
```

### Data Type

 **JSSortKey**

**Applies To:** [JSSortKeys Collection](#jssortkeys-collection)  
**See Also:** [Count Property](#count-property-jssortkeys-collection), [Remove Method](#remove-method-jssortkeys-collection), [Index Property](JSSortKey-Object.md#index-property-jssortkey-object)

## Add Method (JSSortKeys Collection)

Adds a **JSSortKey** object to a **JSSortKeys** collection and returns a reference to the newly created object.  
 Syntax  
 *object*.**Add** *colindex, sortorder, index*  
 The **Add** method syntax has these parts:

| Part | Description |
| --- | --- |
| object | Required. An object expression that evaluates to an object in the Applies To list. |
| colindex | Required. An integer that represents the index of the column whose values will be grouped. |
| sortorder | Required. A value or constant that specifies the sort order of an object. The available sort orders are detailed in the **SortOrder** property. |
| index | Optional. An integer that specifies the index in which the **JSSortKey** will be added. |

### Remarks

 The **JSSortKeys** collection is a 1-based collection.

### Data Type

 **JSSortKey**

**Applies To:** [JSSortKeys Collection](#jssortkeys-collection)  
**See Also:** [ColIndex Property](JSSortKey-Object.md#colindex-property-jssortkey-object), [Index Property](JSSortKey-Object.md#index-property-jssortkey-object), [JSSortKey Object](JSSortKey-Object.md#jssortkey-object), [SortOrder Property](JSSortKey-Object.md#sortorder-property-jssortkey-object)  
**Example:** [SortKeys Example](../../Examples.md#sortkeys-example)

## Clear Method (JSSortKeys Collection)

Removes all objects in a collection.

### Syntax

 *object*.**Clear**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 To remove only one object from a collection, use the **Remove** method of the collection.

**Applies To:** [JSSortKeys Collection](#jssortkeys-collection)  
**See Also:** [Remove Method](#remove-method-jssortkeys-collection)

## Remove Method (JSSortKeys Collection)

Removes a specific member from the **JSSortKeys** collection.

### Syntax

 *object*.**Remove** *index*  
 The **Remove** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that represents the index of a **JSSortKey** within the **JSSortKeys** collection. |

### Remarks

 To remove all the members of a collection, use the **Clear** method.

**Applies To:** [JSSortKeys Collection](#jssortkeys-collection)  
**See Also:** [Item Property](#item-property-jssortkeys-collection), [Index Property](JSSortKey-Object.md#index-property-jssortkey-object), [Clear Method](#clear-method-jssortkeys-collection)
