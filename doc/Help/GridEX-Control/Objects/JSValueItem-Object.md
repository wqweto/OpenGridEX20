# JSValueItem Object

## JSValueItem Object

![](../../images/JSValueItem.jpg)  
 Represents the text and/or image that will be displayed instead a cell value.

### Syntax

 *object*.**Item**(*index*)  
 The **JSValueItem** object syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that represents the value of the **Index** property. |

### Remarks

 With a **JSValueItem** object, you can associate text and/or an image to a given value.  
 To use a **JSValueItem** object, you can use the **ValueItem** property for a column in the **GridEX** control.

**See Also:** [ValueList Property](JSColumn-Object.md#valuelist-property-jscolumn-object), [JSValueList Collection](JSValueList-Collection.md#jsvaluelist-collection), [Item Property](JSValueList-Collection.md#item-property-jsvaluelist-collection), [ItemByValue Property](JSValueList-Collection.md#itembyvalue-property-jsvaluelist-collection)

## IconIndex Property (JSValueItem Object)

Returns or sets the index of the **JSGridImage** associated with the value in a **JSValueItem** object.

### Syntax

 *object*.**IconIndex** [ = *value* ]  
 The **IconIndex** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | An integer that represents the index of a **JSGridImage** in the **JSGridImages** collection. |

### Remarks

 If you set the **IconIndex** property to 0, no image is used.

### Data Type

 Integer

**Applies To:** [JSValueItem Object](#jsvalueitem-object)  
**See Also:** [Index Property](JSGridImage-Object.md#index-property-jsgridimage-object), [EditType Property](JSColumn-Object.md#edittype-property-jscolumn-object)

## Index Property (JSValueItem Object)

Returns a value that represents the index of an object in a collection. Read only.

### Syntax

 *object*.**Index**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The Index property for all collections is 1-based.

### Data Type

 Long

**Applies To:** [JSValueItem Object](#jsvalueitem-object)  
**See Also:** [Count Property](JSValueList-Collection.md#count-property-jsvaluelist-collection), [Item Property](JSValueList-Collection.md#item-property-jsvaluelist-collection), [Remove Method](JSValueList-Collection.md#remove-method-jsvaluelist-collection), [ItemByValue Property](JSValueList-Collection.md#itembyvalue-property-jsvaluelist-collection)

## Text Property (JSValueItem Object)

Returns or sets the text associated with a value in a **JSValueItem** object.

### Syntax

 *object*.**Text** [ = *value* ]  
 The **Text** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A string expression that will be used to replace the cell contents when the cell value matches the **Value** property setting in a **JSValueItem** object. |

### Data Type

 String

**Applies To:** [JSValueItem Object](#jsvalueitem-object)  
**See Also:** [Add Method](JSValueList-Collection.md#add-method-jsvaluelist-collection), [Value Property](#value-property-jsvalueitem-object), [ReplaceValues Property](JSColumn-Object.md#replacevalues-property-jscolumn-object)

## Value Property (JSValueItem Object)

Returns or sets the value of a **JSValueItem** object.

### Syntax

 *object*.**Value** [ = *value* ]  
 The **Value** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A variant expression that represents the value of a JSValueItem object. |

### Remarks

 The **Value** property setting must be unique in a **JSValueList** collection.  
 Any attempt to set the **Value** property to a value that is already in the collection will cause an error.

### Data Type

 Variant

**Applies To:** [JSValueItem Object](#jsvalueitem-object)  
**See Also:** [Add Method](JSValueList-Collection.md#add-method-jsvaluelist-collection), [Text Property](#text-property-jsvalueitem-object), [ReplaceValues Property](JSColumn-Object.md#replacevalues-property-jscolumn-object)

## Visible Property (JSValueItem Object)

Returns or sets a value indicating whether a **JSValueItem** is visible or hidden in a drop-down list.

### Syntax

 *object*.**Visible** [ = *boolean* ]  
 The **Visible** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *boolean* | A Boolean expression specifying whether the object is visible or hidden. |

### Settings

 The settings for *boolean* are:

| Setting | Description |
| --- | --- |
| **True** | Default) Object is visible. |
| **False** | Object is hidden. |

### Data Type

 Boolean

**Applies To:** [JSValueItem Object](#jsvalueitem-object)
