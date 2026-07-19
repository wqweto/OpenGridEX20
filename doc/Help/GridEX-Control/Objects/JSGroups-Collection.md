# JSGroups Collection

## JSGroups Collection

![](../../images/JSGroups.jpg)  
 Contains a collection of **JSGroup** objects.

### Syntax

 *gridex*.**Groups**  
 The *gridex* placeholder represents an object expression that evaluates to a **GridEX** control.

### Remarks

 With the **JSGroups** collection you can add and remove **JSGroup** objects, enumerate or count **JSGroup** objects, and access individual **JSGroup** objects.  
 The **JSGroups** collection can be accessed through the **Groups** property of the **GridEX** control.  
 To obtain a specific **JSGroup** object in the collection you can use the **Item** property of the collection.

**Note** The **JSGroups** collection can contain only four items at any given time. If you try to add more than four items, an error occurs.

- [JSGroup Object](JSGroup-Object.md#jsgroup-object)

**See Also:** [Groups Property](../Properties.md#groups-property-gridex-control), [JSGroup Object](JSGroup-Object.md#jsgroup-object), [Item Property](#item-property-jsgroups-collection)

## Count Property (JSGroups Collection)

Returns the number of objects in a collection.

### Syntax

 *object*.**Count**  
 The object placeholder is an object expression that evaluates to an object in the Applies To list.

### Remarks

 You can use this property with a **For...Next** statement to carry out operations on objects in a collection.  
 For example, the following code prints in the immediate window the **ColIndex** of all the **JSGroup** objects in the collection.

```vb
For I = 1 To GridEX1.Groups.Count
Debug.Print "Groups(" & I & ") =" & _
GridEX1.Groups(I).ColIndex
Next I
```

### Data Type

 Integer

**Note** Since collections are 1-based, there is no need to use the *Collections.Count-1*counter expression in **For…Next** loops.

**Applies To:** [JSGroups Collection](#jsgroups-collection)  
**See Also:** [Item Property](#item-property-jsgroups-collection), [Index Property](JSGroup-Object.md#index-property-jsgroup-object)

## Item Property (JSGroups Collection)

Returns a specific **JSGroup** of the **JSGroups** Collection.

### Syntax

 *object*.**Item(***index***)**  
 The **Item** property syntax has the following parts:

| Part | Description |
| --- | --- |
| *object* | Required. An object expression that evaluates to an object in the Applies To list. |
| *index* | Required. An integer that specifies the position of a **JSGroup** in the collection. |

### Remarks

 If the value provided as index does not match any existing member of the collection, an error occurs.  
 **Item** is the default property for a collection. Therefore, the following lines of code are equivalent:

```vb
Debug.Print GridEX1.Groups(1).ColIndex

Debug.Print GridEX1.Groups.Item(1).ColIndex
```

### Data Type

 **JSGroup**

**Applies To:** [JSGroups Collection](#jsgroups-collection)  
**See Also:** [Count Property](#count-property-jsgroups-collection), [Remove Method](#remove-method-jsgroups-collection), [Index Property](JSGroup-Object.md#index-property-jsgroup-object)

## Add Method (JSGroups Collection)

Adds a **JSGroup** object to the collection and returns a reference to the newly created object.

### Syntax

 *object*.**Add** *colindex, sortorder, index*  
 The **Add** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | Required. An object expression that evaluates to an object in the Applies To list. |
| *colindex* | Required. An integer that represents the index of the column whose values will be grouped. |
| *sortorder* | Required. A value or constant that specifies the sort order of an object. The available sort orders are detailed in the **SortOrder** property. |
| *index* | Optional. An integer that specifies the index in which the group will be added. |

### Remarks

 The **JSGroups** collection is a 1-based collection.

### Data Type

 **JSGroup**

**Applies To:** [JSGroups Collection](#jsgroups-collection)  
**See Also:** [Index Property](JSGroup-Object.md#index-property-jsgroup-object), [JSGroup Object](JSGroup-Object.md#jsgroup-object), [SortOrder Property](JSGroup-Object.md#sortorder-property-jsgroup-object), [ColIndex Property](JSGroup-Object.md#colindex-property-jsgroup-object)  
**Example:** [Groups Example](../../Examples.md#groups-example)

## Clear Method (JSGroups Collection)

Removes all objects in a collection.

### Syntax

 *object*.**Clear**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 To remove only one object from a collection, use the **Remove** method of the collection.

**Applies To:** [JSGroups Collection](#jsgroups-collection)  
**See Also:** [Remove Method](#remove-method-jsgroups-collection)

## Remove Method (JSGroups Collection)

Removes a specific member from the **JSGroups** collection.

### Syntax

 *object*.**Remove** *index*  
 The **Remove** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that represents the index of a **JSGroup** within the collection. |

### Remarks

 To remove all the members of a collection, use the **Clear** method.

**Applies To:** [JSGroups Collection](#jsgroups-collection)  
**See Also:** [Item Property](#item-property-jsgroups-collection), [Index Property](JSGroup-Object.md#index-property-jsgroup-object), [Clear Method](#clear-method-jsgroups-collection)
