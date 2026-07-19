# JSGridImage Object

## JSGridImage Object

![](../../images/JSGridImage.jpg)  
 Represents an image used in a **GridEX** control.

### Syntax

 *object*.**GridImages** ( *index* )  
 The **JSGridImage** object syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to a **GridEX** control. |
| *index* | An integer that represents the value of the **Index** property. |

### Remarks

 A **JSGridImage** object is used in a **GridEX** control to display icons or bitmaps in cells.  
 To use a **JSGridImage** object, you can use the **GridImages** property of the **GridEX** control. You can also assign a **JSGridImage** to a separate variable dimensioned as **JSGridImage**.

**See Also:** [GridImages Property](../Properties.md#gridimages-property-gridex-control), [JSGridImages Collection](JSGridImages-Collection.md#jsgridimages-collection), [Item Property](JSGridImages-Collection.md#item-property-jsgridimages-collection)

## Index Property (JSGridImage Object)

Returns a value that represents the index of an object in a collection. Read only.

### Syntax

 *object*.**Index**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 The Index property for all collections is 1-based.

### Data Type

 Integer

**Applies To:** [JSGridImage Object](#jsgridimage-object)  
**See Also:** [Count Property](JSGridImages-Collection.md#count-property-jsgridimages-collection), [Item Property](JSGridImages-Collection.md#item-property-jsgridimages-collection), [Remove Method](JSGridImages-Collection.md#remove-method-jsgridimages-collection), [IconIndex Property](JSValueItem-Object.md#iconindex-property-jsvalueitem-object), [IconIndex Property](JSRowData-Object.md#iconindex-property-jsrowdata-object), [DefaultIcon Property](JSColumn-Object.md#defaulticon-property-jscolumn-object), [FetchIcon Event](../Events.md#fetchicon-event-gridex-control), [HeaderIcon Property](JSColumn-Object.md#headericon-property-jscolumn-object)

## Picture Property (JSGridImage Object)

Returns or sets a graphic to be displayed.

### Syntax

 *object*.**Picture** [ = *picture*]  
 The **Picture** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *picture* | A **Picture** object containing an icon or bitmap. |

### Remarks

 The picture dimensions remains unchanged even when the **JSGridImages** object resizes the picture to accommodate the graphic to the **ImageWidth** and **ImageHeight** properties.

### Data Type

 **StdPicture**

**Applies To:** [JSGridImage Object](#jsgridimage-object)  
**See Also:** [ImageHeight Property](../Properties.md#imageheight-property-gridex-control), [ImageWidth Property](../Properties.md#imagewidth-property-gridex-control)
