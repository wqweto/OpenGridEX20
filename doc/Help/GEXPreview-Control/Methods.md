# GEXPreview Control — Methods

## NextPage Method (GEXPreview Control)

Displays the next page in a document.

### Syntax

 *object*.**NextPage**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 Calling this method is equivalent to click in the "Next Page" button in the **GEXPreview** control's tool bar.

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [PreviousPage Method](#previouspage-method-gexpreview-control), [Repaginate Method](#repaginate-method-gexpreview-control), [TotalPages Property](Properties.md#totalpages-property-gexpreview-control), [CurrentPage Property](Properties.md#currentpage-property-gexpreview-control)

## PreviousPage Method (GEXPreview Control)

Displays the previous page in a document.

### Syntax

 *object*.**PreviousPage**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 Calling this method is equivalent to click in the "Previous Page" button in the **GEXPreview** control's tool bar.

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [NextPage Method](#nextpage-method-gexpreview-control), [Repaginate Method](#repaginate-method-gexpreview-control), [TotalPages Property](Properties.md#totalpages-property-gexpreview-control), [CurrentPage Property](Properties.md#currentpage-property-gexpreview-control), [PageChanged Event](Events.md#pagechanged-event-gexpreview-control)

## Repaginate Method (GEXPreview Control)

Forces the **GEXPreview** control to re-calculate the layout for pages in a document.

### Syntax

 *object*.**Repaginate**  
 The object placeholder represents an object expression that evaluates to an object in the Applies To list.

### Remarks

 You must call this method if the user has made change to the page layout or the layout in the **GridEX** control has changed.

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [NextPage Method](#nextpage-method-gexpreview-control), [PreviousPage Method](#previouspage-method-gexpreview-control), [BeforePaginating Event](Events.md#beforepaginating-event-gexpreview-control), [AfterPaginating Event](Events.md#afterpaginating-event-gexpreview-control)
