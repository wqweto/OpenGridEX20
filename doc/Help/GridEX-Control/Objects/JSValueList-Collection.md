# JSValueList Collection

## JSValueList Collection

![](../../images/JSValueList.jpg)  
 Represents a list of **JSValueItem** objects.

### Syntax

 *column*.**ValueList**  
 The *column* placeholder represents an object expression that evaluates to a **JSColumn** object.

### Remarks

 With the **JSValueList** collection you can add and remove **JSValueItem** objects, enumerate or count **JSValueItem** objects, and access individual **JSValueItem** objects.  
 The **JSValueList** collection can be accessed through the **ValueList** property of a **JSColumn** object.  
 To obtain a specific **JSValueItem** object in the collection you can use the **Item** or the **ItemByValue** properties.

- [JSValueItem Object](JSValueItem-Object.md#jsvalueitem-object)

**See Also:** [ValueList Property](JSColumn-Object.md#valuelist-property-jscolumn-object), [Item Property](#item-property-jsvaluelist-collection), [ItemByValue Property](#itembyvalue-property-jsvaluelist-collection), [JSValueItem Object](JSValueItem-Object.md#jsvalueitem-object)

## Count Property (JSValueList Collection)

Returns the number of objects in a collection.

### Syntax

 *object*.**Count**  
 The object placeholder is an object expression that evaluates to an object in the Applies To list.

### Remarks

 You can use this property with a **For...Next** statement to carry out operations on objects in a collection.

### Data Type

 Long

**Note** Since collections are 1-based, there is no need to use the *Collections.Count-1*counter expression in **For…Next** loops.

**Applies To:** [JSValueList Collection](#jsvaluelist-collection)  
**See Also:** [Item Property](#item-property-jsvaluelist-collection), [Index Property](JSValueItem-Object.md#index-property-jsvalueitem-object)

## Item Property (JSValueList Collection)

Returns a specific **JSValueItem** of the **JSValueList** collection.

### Syntax

 *object*.**Item(***index***)**  
 The **Item** property syntax has the following parts:

| Part | Description |
| --- | --- |
| *object* | Required. An object expression that evaluates to an object in the Applies To list. |
| *index* | Required. A long that specifies the position of a **JSValueItem** in the **JSValueList** collection. |

### Remarks

 If the value provided in the index argument does not match an existing member of the collection, an error occurs.  
 **Item** is the default property for a collection. Therefore, the following lines of code are equivalent:

```vb
Debug.Print GridEX1.Columns(1).ValueList(1).Text

Debug.Print GridEX1Columns.item(1).ValueList.Item(1).Text
```

**Data Type  
 JSValueItem**

**Applies To:** [JSValueList Collection](#jsvaluelist-collection)  
**See Also:** [Count Property](#count-property-jsvaluelist-collection), [Remove Method](#remove-method-jsvaluelist-collection), [ItemByValue Property](#itembyvalue-property-jsvaluelist-collection), [Index Property](JSValueItem-Object.md#index-property-jsvalueitem-object)

## ItemByValue Property (JSValueList Collection)

Returns a **JSValueItem** given its value.

### Syntax

 *object*.**ItemByValue**( *value* )  
 The **ItemByValue** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | Required. A variant that specifies the unique value of a **JSValueItem** in the **JSValueList** collection. |

### Remarks

 If the value parameter does not match an existing member of the collection, an error occurs.

### Data Type

 **JSValueItem**

**Applies To:** [JSValueList Collection](#jsvaluelist-collection)  
**See Also:** [Index Property](JSValueItem-Object.md#index-property-jsvalueitem-object), [Item Property](#item-property-jsvaluelist-collection), [Value Property](JSValueItem-Object.md#value-property-jsvalueitem-object)

## VisibleCount Property (JSValueList Collection)

Returns the number of visible value items in a **JSValueList**.

### Syntax

 *object*.**VisibleCount**  
 The object placeholder is an object expression that evaluates to an object in the Applies To list.

### Remarks

### Data Type

 Long

**Applies To:** [JSValueList Collection](#jsvaluelist-collection)  
**See Also:** [Visible Property](JSValueItem-Object.md#visible-property-jsvalueitem-object), [Count Property](#count-property-jsvaluelist-collection)

## Add Method (JSValueList Collection)

Adds a **JSValueItem** object to a **JSValueList** collection and returns a reference to the newly created object.

### Syntax

 *object*.**Add** *value, text, iconindex*  
 The **Add** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | Required. An object expression that evaluates to an object in the Applies To list. |
| *value* | Required. A variant expression that represents the unique value for the **JSValueItem** object. |
| *text* | Optional. A string expression that represents the text used to replace the value when is displayed. |
| *iconindex* | Optional. An integer that specifies the index of a **JSGridImage** in the **JSGridImages** collection. |

### Remarks

 The **JSValueList** collection is a 1-based collection.

### Data Type

 **JSValueItem**

**Applies To:** [JSValueList Collection](#jsvaluelist-collection)  
**See Also:** [IconIndex Property](JSValueItem-Object.md#iconindex-property-jsvalueitem-object), [JSValueItem Object](JSValueItem-Object.md#jsvalueitem-object), [Text Property](JSValueItem-Object.md#text-property-jsvalueitem-object), [Value Property](JSValueItem-Object.md#value-property-jsvalueitem-object)  
**Example:** [ValueList Example](../../Examples.md#valuelist-example)

## Clear Method (JSValueList Collection)

Removes all objects in a collection.

### Syntax

 *object*.**Clear**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 To remove only one object from a collection, use the **Remove** method of the collection.

**Applies To:** [GridEX Control](../../GridEX-Control.md#gridex-control)  
**See Also:** [Remove Method](#remove-method-jsvaluelist-collection)  
**Example:** [DropList Example](../../Examples.md#droplist-example)

## Remove Method (JSValueList Collection)

Removes a specific **JSValueItem** from the **JSValueList** collection.

### Syntax

 *object*.**Remove** *index*  
 The **Remove** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | A Long value that represents the **Index** property of a **JSValueItem** object. |

### Remarks

 To remove all the members of a collection, use the **Clear** method.

**Applies To:** [JSValueList Collection](#jsvaluelist-collection)  
**See Also:** [Item Property](#item-property-jsvaluelist-collection), [Index Property](JSValueItem-Object.md#index-property-jsvalueitem-object), [Clear Method](#clear-method-jsvaluelist-collection), [RemoveItemByValue Method](#removeitembyvalue-method-jsvaluelist-collection)

## RemoveItemByValue Method (JSValueList Collection)

Removes a specific **JSValueItem** from the **JSValueList**.

### Syntax

 *object*.**RemoveItemByValue** *value*  
 The **RemoveItemByValue** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A variant that represents the **Value** of a **JSValueItem** object. |

### Remarks

 To remove all the members, use the **Clear** method instead.

**Applies To:** [JSValueList Collection](#jsvaluelist-collection)  
**See Also:** [Remove Method](#remove-method-jsvaluelist-collection), [Value Property](JSValueItem-Object.md#value-property-jsvalueitem-object), [Clear Method](#clear-method-jsvaluelist-collection)
