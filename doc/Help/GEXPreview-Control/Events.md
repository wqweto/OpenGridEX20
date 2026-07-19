# GEXPreview Control — Events

## AfterPaginating Event (GEXPreview Control)

Occurs when paginating of a document has finished.

### Syntax

 **Private Sub** *object*_**AfterPaginating****(** [ *index* **As Integer** ]**)**  
 The **AfterPaginating** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |

### Remarks

 When the **PrintPreview** method in a **GridEX** control is called or when the user changes the paper size or margins using the Page Setup dialog in a **GEXPreview** control, the control paginates the document.  
 After paginating the document the **AfterPaginating** event is fired.

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [Repaginate Method](Methods.md#repaginate-method-gexpreview-control), [BeforePaginating Event](#beforepaginating-event-gexpreview-control), [PageSetup Method](../GridEX-Control/Objects/JSPrinterProperties-Object.md#pagesetup-method-jsprinterproperties-object)

## BeforePaginating Event (GEXPreview Control)

Occurs when the **GEXPreview** control is about to paginate the document.

### Syntax

 **Private Sub** *object*_**BeforePaginating****(** [ *index* **As Integer** ]**)**  
 The **BeforePaginating** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |

### Remarks

 When the **PrintPreview** method in a **GridEX** control is called or when the user changes the paper size or margins using the Page Setup dialog in a **GEXPreview** control, the control paginates the document.  
 Before paginating the document the **BeforePaginating** event is fired.  
 Use this event if you want to send the user a message indicating that the grid is working.

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [Repaginate Method](Methods.md#repaginate-method-gexpreview-control), [AfterPaginating Event](#afterpaginating-event-gexpreview-control), [PageSetup Method](../GridEX-Control/Objects/JSPrinterProperties-Object.md#pagesetup-method-jsprinterproperties-object)

## OnCloseClick Event (GEXPreview Control)

Occurs when the "Close" button, in the **GEXPreview** control's tool bar, has been clicked.

### Syntax

 **Private Sub** *object*_**OnCloseClick (** [ *index* **As Integer** ]**)**  
 The **OnCloseClick** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |

### Remarks

 When the user clicks in the close button, The **GEXPreview** control releases the current document.  
 In this event, you should close the parent form of the **GEXPreview** control.

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [OnPrintClick Event](#onprintclick-event-gexpreview-control)  
**Example:** [PrintPreview Example](../Examples.md#printpreview-example)

## OnPrintClick Event (GEXPreview Control)

Occurs when the "**Close**" button, in the **GEXPreview** control's tool bar, has been clicked.

### Syntax

 **Private Sub** *object*_**OnPrintClick (** [ *index* **As Integer**, ] *useprintsetupdlg***)**  
 The **OnPrintClick** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |
| *useprintsetupdlg* | A **JSRetBoolean** object that can be set to **True** to display the Print Setup dialog before printing the document, as described in settings |

### Settings

 The settings for *useprintsetupdlg* are:

| Setting | Description |
| --- | --- |
| **True** | A page setup dialog is displayed before the **GridEX** control's contents is printed. |
| **False** | Default. **GridEX** control's contents is printed using the settings in the **JSPrinterProperties** object. |

### Remarks

 When the user clicks in the print button, The **GEXPreview** control sends the document to the printer and releases the current document.  
 In this event, you should close the parent form of the **GEXPreview** control.

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [OnCloseClick Event](#oncloseclick-event-gexpreview-control)

## PageChanged Event (GEXPreview Control)

Occurs when the current page in a **GEXPreview** control changes.

### Syntax

 **Private Sub** *object*_**PageChanged****(** [ *index* **As Integer** ]**)**  
 The **PageChanged** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |

### Remarks

 When the **CurrentPage** property in a **GEXPreview** control is changed or when the user changes it using the "Previous Page" or "Next Page" buttons in the toolbar, the **PageChanged** event is fired.  
 Use this event if you want to indicate, in a status bar for instance, the current page in the **GEXPreview** control.

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [NextPage Method](Methods.md#nextpage-method-gexpreview-control), [PreviousPage Method](Methods.md#previouspage-method-gexpreview-control), [Repaginate Method](Methods.md#repaginate-method-gexpreview-control), [TotalPages Property](Properties.md#totalpages-property-gexpreview-control), [CurrentPage Property](Properties.md#currentpage-property-gexpreview-control)

## ZoomChanged Event (GEXPreview Control)

Occurs when the Zoom property in a GEXPreview control changes.

### Syntax

 **Private Sub** *object*_**ZoomChanged (** [ *index* **As Integer** ]**)**  
 The **ZoomChanged** event syntax has these parts:

| Part | Description |
| --- | --- |
| *object* | An object expression that evaluates to an object in the Applies To list. |
| *index* | An integer that identifies a control if it is in a control array. |

### Remarks

 When the **Zoom** property in a **GEXPreview** control is changed or when the user changes it using the **Zoom** buttons in the toolbar or clicking in the page, the **ZoomChanged** event is fired.

**Applies To:** [GEXPreview Control](../GEXPreview-Control.md#gexpreview-control)  
**See Also:** [Zoom Property](Properties.md#zoom-property-gexpreview-control)
