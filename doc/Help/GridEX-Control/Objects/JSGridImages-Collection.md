# JSGridImages Collection

## JSGridImages Collection

![](../../images/JSGridImages.jpg)  
 Contains a collection of **JSGridImage** objects.

### Syntax

 *object*.**GridImages**  
 The *object* placeholder represents an object expression that evaluates to a **GridEX** control.

### Remarks

 With the **JSGridImages** collection you can add and remove **JSGridImage** objects, count or enumerate **JSGridImage** objects, and access individual **JSGridImage** objects.  
 The **JSGridImages** collection can be accessed through the **GridImages** property of the **GridEX** control.  
 To obtain a specific **JSGridImage** object in the collection you can use the **Item** property.

- [JSGridImage Object](JSGridImage-Object.md#jsgridimage-object)

**See Also:** [GridImages Property](../Properties.md#gridimages-property-gridex-control), [JSGridImage Object](JSGridImage-Object.md#jsgridimage-object), [Item Property](#item-property-jsgridimages-collection)

## Count Property (JSGridImages Collection)

Returns the number of objects in a collection.

### Syntax

 *object*.**Count**  
 The object placeholder is an object expression that evaluates to an object in the Applies To list.

### Remarks

 You can use this property with a **For...Next** statement to carry out operations on objects in a collection.

### Data Type

 Integer

**Note** Since collections are 1-based, there is no need to use the *Collections.Count-1*counter expression in **For…Next** loops.

**Applies To:** [JSGridImages Collection](#jsgridimages-collection)  
**See Also:** [Item Property](#item-property-jsgridimages-collection), [Index Property](JSGridImage-Object.md#index-property-jsgridimage-object)

## hImageList Property (JSGridImages Collection)

Returns or sets the handle of an external **ImageList**.

### Syntax

 *object*.**hImageList**  
 The **hImageList** property syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *value* | A Long expression that specifies the handle of an external image list. |

### Remarks

 Use this property when the images that are going to be used in a **GridEX** control are in an external image list control

### Data Type

 Long

**Applies To:** [JSGridImages Collection](#jsgridimages-collection)  
**See Also:** [Add Method](#add-method-jsgridimages-collection)  
**Example:** [hImageList Example](../../Examples.md#himagelist-example)

## Item Property (JSGridImages Collection)

Returns a specific **JSGridImage** object of a **JSGridImages** collection.

### Syntax

 *object*.**Item(***index***)**  
 The **Item** property syntax has the following parts:

| Part | Description |
| --- | --- |
| *object* | Required. An object expression that evaluates to an object in the Applies To list. |
| *index* | Required. An integer that specifies the index of a **JSGridImage** in the collection. |

### Remarks

 If the value provided as index does not match any existing member of the collection, an error occurs.  
 **Item** is the default property for a collection. Therefore, the following lines of code are equivalent:

```vb
Debug.Print GridEX1.GridImages(1).Index

Debug.Print GridEX1.GridImages.Item(1).Index
```

### Data Type

 **JSGridImage**

**Applies To:** [JSGridImages Collection](#jsgridimages-collection)  
**See Also:** [Count Property](#count-property-jsgridimages-collection), [Remove Method](#remove-method-jsgridimages-collection), [Index Property](JSGridImage-Object.md#index-property-jsgridimage-object)

## Add Method (JSGridImages Collection)

Adds a JSGridImage object to the collection and returns a reference to the newly created object.

### Syntax

 *object*.**Add** *picture*  
 The **Add** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | Required. An object expression that evaluates to a Buttons collection. |
| *picture* | Required. Specifies the **Picture** to be added to the collection |

### Remarks

 The **JSGridImages** collection is a 1-based collection.  
 You can load either bitmaps or icons into a **JSGridImage** object. You can also assign **Picture** objects returned from image loading functions, such as **LoadPicture** and **LoadResPicture**. See the Visual Basic help file for more information on these functions.

### Data Type

 **JSGridImage**

**Applies To:** [JSGridImages Collection](#jsgridimages-collection)  
**See Also:** [Picture Property](JSGridImage-Object.md#picture-property-jsgridimage-object), [JSGridImage Object](JSGridImage-Object.md#jsgridimage-object)  
**Example:** [GridImages Example](../../Examples.md#gridimages-example)

## Clear Method (JSGridImages Collection)

Removes all objects in a collection.

### Syntax

 *object*.**Clear**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 To remove only one object from a collection, use the **Remove** method of the collection.

**Applies To:** [JSGridImages Collection](#jsgridimages-collection)  
**See Also:** [Remove Method](#remove-method-jsgridimages-collection)

## Remove Method (JSGridImages Collection)

Removes a specific object from the **JSGridImages** collection.

### Syntax

 *object*.**Remove** *index*  
 The **Remove** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that represents the index of a **JSGridImage** within the **JSGridImages** collection. |

### Remarks

 To remove all the members of a collection, use the **Clear** method.

**Applies To:** [JSGridImages Collection](#jsgridimages-collection)  
**See Also:** [Item Property](#item-property-jsgridimages-collection), [Index Property](JSGridImage-Object.md#index-property-jsgridimage-object), [Clear Method](#clear-method-jsgridimages-collection)
