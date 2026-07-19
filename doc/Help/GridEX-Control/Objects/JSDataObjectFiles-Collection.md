# JSDataObjectFiles Collection

## JSDataObjectFiles Collection

![](../../images/JSDataObjectFiles.jpg)  
 Represents a collection of strings which is the type of the **Files** property on the **JSDataObject** object.

### Syntax

 *jsdataobject*.**Files**  
 The *jsdataobject* placeholder represents an object expression that evaluates to a **JSDataObject** object.

### Remarks

 The **JSDataObjectFiles** collection is a collection of strings which represent a set of files that have been selected either through the **GetData** method, or through selection in an application such as the Windows Explorer.  
 Although the **JSDataObjectFiles** collection has methods and properties of its own, you should use the **Files** property of the **DataObject** object to view and manipulate the contents of the **JSDataObjectFiles** collection.

**Note** This collection is used by the **Files** property only when the data in the **JSDataObject** object is in the **jgexCFFiles** format.

**See Also:** [JSDataObject Object](JSDataObject-Object.md#jsdataobject-object), [Files Property](JSDataObject-Object.md#files-property-jsdataobject-object)

## Count Property (JSDataObjectFiles Collection)

Returns the number of objects in a collection.

### Syntax

 *object*.**Count**  
 The object placeholder is an object expression that evaluates to an object in the Applies To list.

### Remarks

 You can use this property with a **For...Next** statement to carry out operations on objects in a collection.

### Data Type

 Long

**Applies To:** [JSDataObjectFiles Collection](#jsdataobjectfiles-collection)  
**See Also:** [Item Property](#item-property-jsdataobjectfiles-collection)

## Item Property (JSDataObjectFiles Collection)

Returns an specific file name of the **JSDataObjectFiles** collection.

### Syntax

 *object*.**Item**(*index*)  
 The **Item** property syntax has the following parts:

| Part | Description |
| --- | --- |
| *object* | Required. An object expression that evaluates to an object in the Applies To list. |
| *index* | Required. A long that specifies the index of a file name in the collection. |

### Remarks

 If the value provided as index does not match any existing member of the collection, an error occurs.

### Data Type

 **String**

**Applies To:** [JSDataObjectFiles Collection](#jsdataobjectfiles-collection)  
**See Also:** [Count Property](#count-property-jsdataobjectfiles-collection), [Remove Method](#remove-method-jsdataobjectfiles-collection)

## Add Method (JSDataObjectFiles Collection)

Adds a filename to the Files collection of a **JSDataObject** object (**jgexCFFiles** format only).

### Syntax

 *object*.**Add** *filename, index*  
 The **Add** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | Required. An object expression that evaluates to an object in the Applies To list. |
| *filename* | A string expression specifying the file name to add to the collection. |
| *index* | Optional. A long that specifies the index in which the file will be added. |

### Remarks

 This method is available only for component drag sources. If **Add** is called from a component drop target, an error is generated.

**Applies To:** [JSDataObjectFiles Collection](#jsdataobjectfiles-collection)  
**See Also:** [Files Property](JSDataObject-Object.md#files-property-jsdataobject-object)

## Clear Method (JSDataObjectFiles Collection)

Removes all objects in a collection.

### Syntax

 *object*.**Clear**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 To remove only one object from a collection, use the **Remove** method of the collection.

**Applies To:** [JSDataObjectFiles Collection](#jsdataobjectfiles-collection)  
**See Also:** [Remove Method](#remove-method-jsdataobjectfiles-collection)

## Remove Method (JSDataObjectFiles Collection)

Removes an specific file name from the **JSDataObjectFiles** collection.

### Syntax

 *object*.**Remove** *index*  
 The **Remove** method syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | A long that identifies the index of a file name within the collection. |

### Remarks

 To remove all the members of a collection, use the **Clear** method.

**Applies To:** [JSDataObjectFiles Collection](#jsdataobjectfiles-collection)  
**See Also:** [Item Property](#item-property-jsdataobjectfiles-collection), [Clear Method](#clear-method-jsdataobjectfiles-collection)
