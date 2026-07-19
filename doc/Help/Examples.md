# Examples

## ActAsDropDown Example

Example

This example uses a GridEX control as the drop down list for a column in another GridEX control.

 To try this example follow the next steps:

 Step 1) Start a new project.

 Step 2) Select "Janus GridEX Control 2.0" from VB Components dialog.

 Step 3) Place two GridEX controls in Form1

 Step 4) In the Load event of Form1 form, paste the following code:  
 (Change DBPath const to the correct JSNWind.MDB path)

```vb
Private Sub Form_Load()

Dim strDBName As String

    Const DBPath = "C:\JSNWind.MDB"
    If GridEX1.DataMode = jgexADO Then
        strDBName = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath
    Else
        strDBName = DBPath
    End If
    GridEX1.DatabaseName = strDBName
    GridEX1.RecordSource = "SELECT * FROM Products"
    GridEX1.ClearFields
    GridEX1.Rebind
    If GridEX2.DataMode = jgexADO Then
        strDBName = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath
    Else
        strDBName = DBPath
    End If
    GridEX2.DatabaseName = strDBName
    GridEX2.RecordSource = "SELECT SupplierID, CompanyName FROM Suppliers"
    GridEX2.ColumnAutoResize = True
    GridEX2.ClearFields
    GridEX2.Rebind

    'GridEX2 will act as the drop down list
    'for column 'SupplierID' in GridEX1
    GridEX2.ActAsDropDown = True
    GridEX2.BoundColumnIndex = "SupplierID"
    GridEX2.ReplaceColumnIndex = "CompanyName"
    
    'Set column 'SupplierID' in GridEX1 properties
    'to be a combo box with GridEX2 as its DropDown list
    With GridEX1.Columns("SupplierID")
        .TextAlignment = jgexAlignLeft
        .EditType = jgexEditCombo
        Set .DropDownControl = GridEX2
    End With
    
End Sub
``` 
 Step 5) Run the program and see how editing works in the SupplierID column.

## ADORecordset Example

Example

This example opens an ADO Recordset and sets it to the **ADORecordset** property of a **GridEX** control.

 To try this example, paste the code into the declarations section of a form with a **GridEX** control in it and add "Microsoft ActiveX Data Objects" in the References Dialog of the project.

```vb
Private Sub Form_Load()

Dim strConn As String
Dim cn As Connection
Dim RS As Recordset

    'Change DBPath const to a valid mdb file name
    Const DBPath = "C:\JSNWind.mdb"
    Const SQLStatement = "SELECT * FROM Customers"
    
    'Opening the Recordset
    Set cn = New Connection
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath
    cn.Open

    Set RS = New Recordset

    RS.Open SQLStatement, cn, adOpenKeyset, adLockOptimistic
    'Setting ADORecordset property has the same effect as calling the Rebind method.
    Set GridEX1.ADORecordset = RS
    
End Sub
```

## BeforeColEdit Example

Example

This example prevents edition in column 3 of a GridEX control.

```vb
Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, _
                                ByVal Cancel As JSRetBoolean)

    If ColIndex = 3 Then
        Cancel = True
    End If

End Sub
```

## BeforeColMove Example

Example

This example prevents moving any column to the leftmost position. It also prevents the leftmost column to be moved to any other position.

```vb
Private Sub GridEX1_BeforeColMove(ByVal Column As JSColumn, _
ByVal NewPosition As Integer, ByVal Cancel As JSRetBoolean)

    If Column.ColPosition = 1 Then
        'Leftmost column can't be moved to other positions.
        MsgBox "Column '" & Column.Caption & "' can't be moved.", vbExclamation
        Cancel = True
    ElseIf NewPosition = 1 Then
        'other columns can't be moved to the first position.
        MsgBox "Column '" & Column.Caption & "' can't be at the left.", vbExclamation
        Cancel = True
    End If
    
End Sub
```

## BeforeColumnDrag Example

Example

This example prevents frozen columns to be dragged by the user.

 To try this example, paste the code into the declarations section of a form that contains a GridEX control with at least one frozen column.

```vb
Private Sub GridEX1_BeforeColumnDrag _
(ByVal Column As JSColumn, ByVal Cancel As JSRetBoolean)

    If Column.ColPosition <= GridEX1.FrozenColumns Then
        Cancel = True
    End If
    
End Sub
```

## BeforeColUpdate Example

Example

This example validates data entered by the user.

 Data in column 1 should be numeric or empty while Data in column 2 should be a date.
```vb
Private Sub GridEX1_BeforeColUpdate(ByVal Row As Long, _
ByVal ColIndex As Integer, ByVal OldValue As String, _
ByVal Cancel As JSRetBoolean)

Dim strValue As String

    Select Case ColIndex
        Case 1
            strValue = GridEX1.Value(ColIndex)
            If Len(varValue) > 0 Then
                If Not IsNumeric(strValue) Then
                    MsgBox "Only numeric entries are allowed in this column.", _
                    vbExclamation
                    Cancel = True
                End If
            End If
        Case 2
            strValue = GridEX1.Value(ColIndex)
            If Len(strValue) > 0 Then
                If Not IsDate(strValue) Then
                    MsgBox strValue & " is not a valid date.", _
                    vbExclamation
                    Cancel = True
                End If
            End If
    End Select
End Sub
```

## BeforeDelete Example

Example

This example prompts for confirmation before a delete operation takes place.
```vb
Private Sub GridEX1_BeforeDelete(ByVal Cancel As JSRetBoolean)

Dim strMessage As String

    'the sample assumes MultiSelect property is set to True.
    strMessage = "The application is about to delete " & _
    GridEX1.SelectedItems.Count & " record(s)."
    strMessage = strMessage & vbCrLf & "Are you sure you want to delete these records?"
    
    If MsgBox(strMessage, vbQuestion + vbYesNo) = vbNo Then
        Cancel = True
    End If
    
End Sub
```

## BeforeDeleteEx Example

Example

This example prompts for confirmation before a a record is deleted.

 If more than one record are deleted at a time, confirmation will be requested for each record.
```vb
Private Sub GridEX1_BeforeDeleteEX(ByVal RowIndex As Long, _
ByVal Bookmark As Variant, ByVal Cancel As JSRetBoolean)

Dim Rs As Recordset
Dim strMessage As String

    Set Rs = GridEX1.ADORecordset
    Rs.Bookmark = Bookmark
    strMessage = "Are you sure you want to delete Order No.: " & Rs![OrderID] & "?"
    If MsgBox(strMessage, vbQuestion + vbYesNo) = vbNo Then
        Cancel = True
    End If
    
End Sub
```

## BeforeGroupChange Example

Example

This example prevents grouping records by an icon column while other columns can still be grouped by the user.

```vb
Private Sub GridEX1_BeforeGroupChange _
(ByVal Group As JSGroup, ByVal ChangeOperation As jgexGroupChange, _
ByVal GroupPosition As Integer, ByVal Cancel As JSRetBoolean)

Dim colToBeGrouped As JSColumn

    If ChangeOperation = jgexGroupInsert Then
        Set colToBeGrouped = GridEX1.Columns(Group.ColIndex)
        If colToBeGrouped.ColumnType = jgexIcon Then
            MsgBox "Records can not grouped by this column.", vbExclamation
            'don't allow users to group by an icon column
            Cancel = True
        End If
    End If
    
End Sub
```

## BeforePrinting Example

Example

This example prompts for confirmation before printing a document.

```vb
Private Sub GridEX1_BeforePrinting _
(ByVal nPages As Long, ByVal Cancel As JSRetBoolean)

Dim strMessage As String

    strMessage = "You are about to print " _
    & nPages & " page(s)."
    strMessage = strMessage & vbCrLf _
    & "Do you want to continue?"
    If MsgBox(strMessage, vbQuestion + vbYesNo, _
    "Before Print...") = vbYes Then
        Cancel = True
    End If
    
End Sub
```

## BeforePrintPage Example

Example

This example demonstrates how BeforePrintPage event can be used to print page numbers in pages while printing a document.

```vb
Private Sub GridEX1_BeforePrintPage _
(ByVal PageNumber As Long, ByVal nPages As Long)

    GridEX1.PrinterProperties.FooterString(jgexHFRight) _
    = "Page " & PageNumber & " of " & nPages
    
End Sub
```

## BeforeUpdate Example

Example

This example performs a simple validation before a record is updated.

```vb
Private Sub GridEX1_BeforeUpdate(ByVal Cancel As JSRetBoolean)
Dim bIsEmpty As Boolean

    'Checking non null values in first column.
    If IsNull(GridEX1.Value(1)) Then
        bIsEmpty = True
    Else
        If GridEX1.Value(1) = "" Then
            bIsEmpty = True
        End If
    End If

    If bIsEmpty Then
        MsgBox "The field '" & GridEX1.Columns(1).Caption & "' can not be null.", vbExclamation
        Cancel = True
    End If
    
End Sub
```

## CellStyle Example

Example

This example demonstrates the use of the CellStyle and HeaderStyle properties in a JSColumn object.

```vb
Private Sub SetColumnStyles()
Dim fs As JSFormatStyle
Dim col As JSColumn
    
    'adding the style to be used in 'ColumnX' cells
    Set fs = GridEX1.FormatStyles.Add("ColumnCells")
    fs.BackColor = vbYellow
    fs.ForeColor = vbBlue
    fs.FontBold = True
    fs.FontItalic = True
    'adding the style to be used in 'ColumnX' header
    Set fs = GridEX1.FormatStyles.Add("ColumnHeader")
    fs.FontBold = True
    fs.FontUnderline = True
    fs.ForeColor = vbRed
    
    'setting column properties
    Set col = GridEX1.Columns("ColumnX")
    col.CellStyle = "ColumnCells"
    col.HeaderStyle = "ColumnHeader"
    
End Sub
```

## Col Example

Example

This example gets the current column's caption using the Col property and the JSColumns' ItemByPosition method.

```vb
Dim colCurrent As JSColumn

    If GridEX1.col = 0 Then
        MsgBox "There is no current column."
    Else
        Set colCurrent = GridEX1.Columns.ItemByPosition(GridEX1.col)
        MsgBox "The current column is '" & colCurrent.Caption & "'"
    End If
```

## ColButtonClick Example

Example

This example opens a modal form with a text box to edit large fields when user presses a button in a column which ButtonStyle property is set to jgexButtonEllipsis.

```vb
Private Sub Form_Load()
   GridEX1.Columns("Notes").ButtonStyle = jgexButtonEllipsis
End Sub

Private Sub GridEX1_ColButtonClick(ByVal ColIndex As Integer)

    Form2.Text1.Text = GridEX1.Value(ColIndex)
    Form2.Caption = GridEX1.Columns(ColIndex).Caption
    Form2.Show vbModal
    
End Sub
```

## ColFormat Example

Example

This example demonstrates the use of some column properties.

```vb
Private Sub SetDateColumnProperties()
Dim col As JSColumn

    Set col = GridEX1.Columns("Date")
    col.Format = "Medium Date"
    col.GroupFormat = "Long Date"
    col.GroupPrefix = "Date:"
    col.EditType = jgexEditCalendarDropDown
    col.GroupEmptyStringCaption = "(Not assigned)"
    col.NullBehavior = jgexNBNull
    col.SortType = jgexSortTypeDate
    col.DefaultValue = Date
    col.TotalRowFormat = "Medium Date"
    col.TotalRowPrefix = "MAX="
    col.AggregateFunction = jgexMax
    
End Sub
```

## ColFromPoint Example

Example

This example displays a context menu when user right clicked in a column header, a group by box header or a row header.

 The menu used depends on where the user clicked.

```vb
Private Sub GridEX1_MouseUp(Button As Integer, Shift As Integer, _
X As Single, Y As Single)

Dim colClicked As JSColumn
Dim grpClicked As JSGroup
    
    If Button = vbRightButton Then
        Select Case GridEX1.HitTest(X, Y)
            Case jgexHTRowHeader
                'saving the row where the user clicked
                'to be used as a reference in mnuRowHeader
                mnuRowHeaders.Tag = GridEX1.RowFromPoint(X, Y)
                Me.PopupMenu mnuRowHeaders

            Case jgexHTColumnHeader
                Set colClicked = GridEX1.ColFromPoint(X, Y)
                If Not colClicked Is Nothing Then
                'saving the column index where the user clicked
                'to be used as a reference in mnuColumnHeader
                    mnuColumnHeader.Tag = colClicked.Index
                    Me.PopupMenu mnuColumnHeader
                End If

            Case jgexHTGroupByBox
                Set grpClicked = GridEX1.GroupFromPoint(X, Y)
                If Not grpClicked Is Nothing Then
                'saving the group index where the user clicked
                'to be used as a reference in mnuGroupbyBoxHeader
                    mnuGroupByBoxHeader.Tag = grpClicked.Index
                    Me.PopupMenu mnuGroupByBoxHeader
                End If
        
        End Select
    End If
End Sub
```

## CollapseAll Example

Example

This example groups records by the first column and collapses or expands all the groups in the control.

```vb
Private Sub GroupRecords()

    GridEX1.Groups.Add 1, jgexSortAscending
    
    'RefreshGroups method must be called before CollapseAll or ExpandAll
    'if Groups collection has changed. If you didn't add or remove any group,
    'there is no need to refresh the groups in the control.
    
    GridEX1.RefreshGroups

    If MsgBox("Do you want the groups collapsed?", vbQuestion + vbYesNo) = vbYes Then
        GridEX1.CollapseAll
    Else
        GridEX1.ExpandAll
    End If
    
End Sub
```

## ColumnAdd Example

Example

This example adds a column to a GridEX control and sets some of the column properties.

```vb
Private Sub AddColumn()

Dim colNew As JSColumn

    Set colNew = GridEX1.Columns.Add
    colNew.Caption = "Name"
    colNew.ColumnType = jgexText
    colNew.DataField = "Name"
    colNew.EditType = jgexEditTextBox
    colNew.GroupPrefix = "Name:"
    colNew.HeaderAlignment = jgexAlignCenter
    colNew.HeaderToolTip = "Insert the name in this column"
    colNew.MaxLength = 50
    colNew.NullBehavior = jgexNBEmptyString
    colNew.SortType = jgexSortTypeString
    colNew.Width = 3000 'twips

End Sub
```

## ColumnAutosize Example

Example

This example uses the AutoSize method to make each column in a GridEX control as wide as its widest cell.

```vb
Private Sub Form_Load()
Dim colTemp As JSColumn

    'Calling AutoSize in each column of a
    'GridEX control to make the columns
    'width fit their longest entry
    
    For Each colTemp In GridEX1.Columns
        colTemp.AutoSize
    Next
    
End Sub
```

## ColumnCaption Example

Example

This example sets the GroupPrefix property of each column in a GridEX control equal to its Caption property.

```vb
Private Sub Form_Load()
Dim colTemp As JSColumn

    For Each colTemp In GridEX1.Columns
        colTemp.GroupPrefix = colTemp.Caption & ": "
    Next
    
End Sub
```

## ColumnHeaderClick Example

Example

This example sorts records in a GridEX control when the user clicks in a column header or a group by box header.

 If the column clicked hasn't been sorted, it should sort the column in ascending order.

 If the column is sorted, the sort order changes and if the column is grouped, the sort order of the group changes.

```vb
Private Sub GridEX1_ColumnHeaderClick(ByVal Column As JSColumn)

Dim intSortOrder As Integer
Dim jgrTemp As JSGroup

    If Column.IsGrouped Then
        'Get the JSGroup object for this column
        
        For Each jgrTemp In GridEX1.Groups
            If jgrTemp.ColIndex = Column.Index Then
                'Make the same as when the group is clicked and exit
                
                GridEX1_GroupByBoxHeaderClick jgrTemp
                Exit Sub
            End If
        Next
    Else
        'get the sort order of the clicked column 
        
        intSortOrder = Column.SortOrder
        'remove all sortkeys
        
        GridEX1.SortKeys.Clear
        If intSortOrder = jgexSortAscending Then
            'if the clicked column is sorted in ascending order then
            'sort it in descending order
            
            GridEX1.SortKeys.Add Column.Index, jgexSortDescending
        Else
            'if the clicked column  is not sorted
            'or sorted in descending order then
            'sort it in ascending order
            
            GridEX1.SortKeys.Add Column.Index, jgexSortAscending
        End If
    End If
    
End Sub

Private Sub GridEX1_GroupByBoxHeaderClick(ByVal Group As JSGroup)

    'This changes the group's sort order
    
    Group.SortOrder = -Group.SortOrder
    
    'The previous line has the same effect as the following lines:
    
'    If Group.SortOrder = jgexSortAscending Then
'        Group.SortOrder = jgexSortDescending
'    Else
'        Group.SortOrder = jgexSortAscending
'    End If

End Sub
```

## Columns Example

Example

This example, inspects all the columns in GridEX control.

 Column properties are printed in the immediate window. They are retrieved by its index and by its position in the control.

```vb
Private Sub InspectColumns()

Dim col As JSColumn
Dim i As Integer

    'inspecting column by its order in the JSColumns collection
    For i = 1 To GridEX1.Columns.Count
        Set col = GridEX1.Columns.Item(i)
        Debug.Print "Index:=" & col.Index, "Caption:=" _
        & col.Caption, "Position:=" & col.ColPosition
    Next

    'the same can be done using For Each...Next
    For Each col In GridEX1.Columns
        Debug.Print "Index:=" & col.Index, "Caption:=" _
        & col.Caption, "Position:=" & col.ColPosition
    Next

    'inspecting columns by its position
    
    For i = 1 To GridEX1.Columns.Count
        Set col = GridEX1.Columns.ItemByPosition(i)
        Debug.Print "Position:=" & col.ColPosition, _
        "Index:=" & col.Index, "Caption:=" _
        & col.Caption, "Visible:=" & col.Visible
    Next
    
End Sub
```

## ColumnTotalRowFormat Example

Example

This example customize a GridEX control to show the Sum of a column named "Amount" in the group footer.

```vb
Private Sub TotalRowProperties()

Dim colTemp As JSColumn

    'Making the control to show the sum of
    'the amount column in group footers
    
    GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
    Set colTemp = GridEX1.Columns("Amount")
    colTemp.AggregateFunction = jgexSum
    colTemp.TotalRowFormat = "Currency"
    colTemp.TotalRowPrefix = "SUM="

End Sub
```

## CustomEdit Example

Example

This example demonstrates how custom edit events should be handled.

 In the example, column 2 is a CustomEdit column and the edit control used to edit this column is Text1.

 To run this example, create a new project, place a GridEX control in Form1, enter DatabaseName and RecordSource properties at design time and paste the following code into the form module.

```vb
Private Sub Form_Load()

    GridEX1.Columns(2).EditType = jgexEditCustom
    
End Sub

Private Sub GridEX1_EndCustomEdit(ByVal ColIndex As Integer)

    If ColIndex = 2 Then
        If Text1.Tag = "Changed" Then
            GridEX1.Value(ColIndex) = Text1.Text
        End If
    End If
    
End Sub

Private Sub GridEX1_InitCustomEdit(ByVal ColIndex As Integer, _
ByVal EditBackColor As stdole.OLE_COLOR, _
ByVal EditForeColor As stdole.OLE_COLOR, _
ByVal EditFont As stdole.Font)

    If ColIndex = 2 Then
        Set Text1.Font = EditFont
        Text1.BackColor = EditBackColor
        Text1.ForeColor = EditForeColor
        Text1.Text = GridEX1.Value(ColIndex)
        Text1.Tag = ""
    End If
    
End Sub

Private Sub GridEX1_ShowCustomEdit(ByVal ColIndex As Integer, _
ByVal EditLeft As Single, ByVal EditTop As Single, _
ByVal EditWidth As Single, ByVal EditHeight As Single, _
ByVal EditVisible As Boolean)

    'In this event you show or hide the control
    'you are using for making the edits.
    'Also, you must move the control to the
    'appropriate position and size
    
    'NOTE: EditLeft and EditTop parameters are given
    'based on the left-top corner of the control
        
    If ColIndex = 2 Then
        If EditVisible Then
        
            'Set HideSelection to jgexHighLightNormal.
            GridEX1.HideSelection = jgexHighLightNormal
            
            'Doing this the control will be shown highlighted
            'when your edit control gain the focus.
            Text1.Move GridEX1.Left + ScaleX(EditLeft, vbTwips, Me.ScaleMode), _
            GridEX1.Top + ScaleY(EditTop, vbTwips, Me.ScaleMode), _
            ScaleX(EditWidth, vbTwips, Me.ScaleMode), _
            ScaleY(EditHeight, vbTwips, Me.ScaleMode)
            Text1.Visible = True
            Text1.SetFocus
        Else
            GridEX1.SetFocus
            
            'Re set the HideSelection to the default
            GridEX1.HideSelection = jgexHideSelection
            Text1.Visible = False
        End If
    End If
End Sub

Private Sub Text1_Change()

    Text1.Tag = "Changed"
    
End Sub
```

## DataChanged Example

Example

This example prompts the user for the action to take when a form is unloaded and data in the current record of a GridEX control has changed.

 If user wants to save the changes, the Update method is called.

```vb
Private Sub Form_Unload(Cancel As Integer)
Dim intRes As Integer
Dim strMessage As String

    If GridEX1.DataChanged Then
        strMessage = "Data in the current record has changed."
        strMessage = strMessage & vbCrLf & _
        "Do you want to save the changes?"
        intRest = MsgBox(strMessage, vbYesNoCancel + vbInformation)
        Select Case intRes
            Case vbYes
                GridEX1.Update
                If GridEX1.DataChanged Then
                    'If Datachanged is still True,
                    'an error occurred when updating the record.
                    Cancel = True
                End If
            Case vbNo
                'do nothing
            Case vbCancel
                Cancel = True
        End Select
    End If
    
End Sub
```

## DropList Example

Example

This example demonstrates how the DropList event can be used to have different ValueLists for the same column depending on the value of another column.

```vb
Private Sub Form_Load()    

    With GridEX1.Columns(2)
        .EditType = jgexEditDropDown
        .ReplaceValues = False
        .HasValueList = True
    End With

End Sub

Private Sub GridEX1_DropList(ByVal ColIndex As Integer)
Dim VL As JSValueList

    If ColIndex = 2 Then
        Set VL = GridEX1.Columns(2).ValueList
        VL.Clear
        Select Case GridEX1.Value(1)
            Case "A"
                VL.Add "A1", "A1"
                VL.Add "A2", "A2"
                VL.Add "A3", "A3"
            Case "B"
                VL.Add "B1", "B1"
                VL.Add "B2", "B2"
                VL.Add "B3", "B3"
            Case "C"
                VL.Add "C1", "C1"
                VL.Add "C2", "C2"
                VL.Add "C3", "C3"
                VL.Add "C4", "C4"
        End Select
    End If

End Sub
```

## EditMode Example

Example

This example forces the GridEX control to enter into edit mode immediately after the user changes from cell.

 It also selects the text in the cell once the control is in EditMode.

```vb
Private Sub GridEX1_RowColChange(ByVal LastRow As Long, _
ByVal LastCol As Integer)

Dim colCurrent As JSColumn

    'if GridEX isn't in edit mode and there is a current column
    If GridEX1.EditMode = jgexEditModeOff And GridEX1.col <> 0 Then
        'enter edit mode
        GridEX1.EditMode = jgexEditModeOn
        'get the current column
        Set colCurrent = GridEX1.Columns.ItemByPosition(GridEX1.col)
        'select all the text in the cell
        GridEX1.SelStart = 0
        GridEX1.SelLength = Len(GridEX1.Value(colCurrent.Index))
    End If
End Sub
```

## Error Example

Example

This example demonstrates the use of the Error event.

 In this event, you can customize your error messages and instruct GridEX control not to display the default error message .

```vb
Private Sub GridEX1_Error(ByVal ErrNumber As Long, _
ByVal DisplayMessage As JSRetBoolean)

    'display customized error message
    MsgBox "Error No.: " & ErrNumber & vbCrLf & _
    GridEX1.ErrorText
    
    If GridEX1.DataChanged Then
    
        'If Datachanged is True, you can also ask
        'your users if they want to drop their changes.
        If MsgBox("Do you want to drop the changes?", _
        vbQuestion + vbYesNo) = vbYes Then
            GridEX1.DataChanged = False
        End If
    End If
    
    'Force GridEX control to not display its
    'default error message.
    DisplayMessage = False
    
End Sub
```

## FetchData Example

Example

This example demonstrates how FetchData event works.

 In the FetchData event, values from the underlying Recordset are taken to calculate a FetchData column.

```vb
Private Sub Form_Load()

    GridEX1.Columns(6).FetchData = True
    
End Sub

For ADO DataMode:

Private Sub GridEX1_FetchData(ByVal RowIndex As Long, _
ByVal ColIndex As Integer, ByVal RowBookmark As Variant, _
ByVal Value As JSRetVariant)

Dim Rs As Recordset

    If ColIndex = 6 Then
        Set Rs = GridEX1.ADORecordset
        Rs.Bookmark = RowBookmark
        Value = Rs![Quantity] * Rs![UnitPrice]
    End If
    
End Sub

For DAO DataMode:

Private Sub GridEX1_FetchData(ByVal RowIndex As Long, _
ByVal ColIndex As Integer, ByVal RowBookmark As Variant, _
ByVal Value As JSRetVariant)

Dim Rs As Recordset

    If ColIndex = 6 Then
        Set Rs = GridEX1.Recordset
        Rs.Bookmark = RowBookmark
        Value = Rs![Quantity] * Rs![UnitPrice]
    End If
    
End Sub
```

## FetchIcon Example

Example

This example demonstrates how FetchIcon event works.

 In the FetchIcon event, the icon index for the cell is assigned based on the value a field named Status in the underlying Recordset.

```vb
Private Sub Form_Load()

    'used for rows where Status = "Not Started"
    GridEX1.GridImages.Add LoadResPicture(101, vbResBitmap)
    
    'used for rows where Status = "Incomplete"
    GridEX1.GridImages.Add LoadResPicture(102, vbResBitmap)
    GridEX1.Columns(6).ColumnType = jgexIcon
    GridEX1.Columns(6).FetchIcon = True

End Sub

Private Sub GridEX1_FetchIcon(ByVal RowIndex As Long, _
ByVal ColIndex As Integer, ByVal RowBookmark As Variant, _
ByVal IconIndex As JSRetInteger)

Dim Rs As Recordset

    If ColIndex = 6 Then
        Set Rs = GridEX1.ADORecordset
        Rs.Bookmark = RowBookmark
        Select Case Rs![Status]
            Case "Not started"
                IconIndex = 1
            Case "Incomplete"
                IconIndex = 2
            Case Else
                IconIndex = 0
        End Select
    End If
End Sub
```

## Find Example

Example  (Find)

This example demonstrates the use of the Find method.  
 In order to run this example, create a new project, place a GridEX control in Form1, set DatabaseName property to JSNwind.MDB file and RecordSource equal to "Customers". Add a command button and in the Click event for the button call FindCustomer procedure.

```vb
Private Sub FindCustomer()
Dim strFind As String

    strFind = InputBox("Enter the Customer ID you want to find:")
    If Len(strFind) > 0 Then
        If Not GridEX1.Find(1, jgexEqual, strFind) Then
            MsgBox "Customer '" & strFind & "' was not found."
        End If
    End If
    
End Sub
```

## FmtConditions Example

Example

This example ilustrates the use of FmtConditions in a GridEX control.

 To run this example, create a new project, place a GridEX control in Form1, set DatabaseName property to JSNwind.MDB file and RecordSource equal to "Products"

```vb
Private Sub Form_Load()

    AddConditionalFormatting

End Sub

Private Sub AddConditionalFormatting()
Dim col As JSColumn
Dim fmtCon As JSFmtCondition

    'This procedure sets bold style to all the products on sale

    'Get the OnSale column
    Set col = GridEX1.Columns("OnSale")
    'add a format condition for the products on sale
    Set fmtCon = GridEX1.FmtConditions.Add(col.Index, jgexEqual, True)
    'Set FontBold property to True in the FormatStyle for this condition
    fmtCon.FormatStyle.FontBold = True
    'set the Group Condition too
    With GridEX1.FmtConditions
        .ApplyGroupCondition = True
        .ShowGroupConditionCount = True
        .GroupConditionCountTitle = "On Sale"
        Set fmtCon = .GroupCondition
    End With

    fmtCon.SetCondition col.Index, jgexEqual, True
    fmtCon.FormatStyle.FontBold = True

    'Add another format condition for discontinued products
    Set col = GridEX1.Columns("Discontinued")
    Set fmtCon = GridEX1.FmtConditions.Add(col.Index, jgexEqual, True)

    'Set properties in the FormatStyle for this condition
    fmtCon.FormatStyle.FontStrikethru = True
    fmtCon.FormatStyle.ForeColor = vbGrayText
    
End Sub
```

## FormatStyles Example

Example

This example illustrates how FormatStyles can be used in a GridEX control to format cells or rows.

```vb
Private Sub Form_Load()
Dim fstCustom As JSFormatStyle

    'Adding custom FormatStyles
    
    'adding a format style to be used in
    'cells with a negative amount.
    Set fstCustom = GridEX1.FormatStyles.Add("Negative")
    fstCustom.ForeColor = vbRed
    
    'Adding a format style to be used in
    'rows where Paid field is equal to True
    Set fstCustom = GridEX1.FormatStyles.Add("Paid")
    fstCustom.BackColor = vbButtonFace
    
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As JSRowData)

    Const IndexPaid = 15
    Const IndexAmount = 8
    
    If RowBuffer.Value(IndexPaid) = True Then
        'Setting the Format style for all cells in the row
        RowBuffer.RowStyle = "Paid"
    End If
    If RowBuffer.Value(IndexAmount) < 0 Then
        'setting the Format style just for
        'the cell in the Amount column.
        RowBuffer.CellStyle(IndexAmount) = "Negative"
    End If
    
End Sub
```

## FrozenColumns Example

Example

This example illustrates the use of the FrozenColumn property that lets you have non-scrollable columns in the GridEX control. The example, freeze the current column and, in case the column is already frozen, it will be unfreeze.

```vb
Private Sub FreezeCurrentColumn()
Dim colCurrent As JSColumn

    If GridEX1.col = 0 Then
        MsgBox "There is no current column.", vbExclamation
    Else
        Set colCurrent = GridEX1.Columns.ItemByPosition(GridEX1.col)
        If colCurrent.ColPosition <= GridEX1.FrozenColumns Then
            If MsgBox("Current column is already non-scrollable." & _
            vbCrLf & "Do you want to make it scrollable?", _
            vbQuestion + vbYesNo) = vbYes Then
                GridEX1.FrozenColumns = 0
            End If
        Else
            'first move it to the left most position
            colCurrent.ColPosition = 1
            'freeze the column
            GridEX1.FrozenColumns = 1
        End If
    End If
        
End Sub
```

## GetRowData Example

Example

This example illustrates the use of GetRowData method.

 In the example, when the user updates column 15, the image in column 16 is changed.

 In addition, when updating column 5, if the user entered a negative number, the CellStyle of cell in column 5 changes to "Red"

 ( "Red" is the name of a JSFormatStyle in the JSFormatStyles collection)

```vb
Private Sub GridEX1_AfterColUpdate(ByVal ColIndex As Integer)
Dim rdCurrent As JSRowData

    Select Case ColIndex
        Case 5
            Set rdCurrent = GridEX1.GetRowData(GridEX1.Row)
            If CCur(rdCurrent.Value(ColIndex)) < 0 Then
                rdCurrent.CellStyle(ColIndex) = "Red"
            Else
                rdCurrent.CellStyle(ColIndex) = ""
            End If
        Case 15
            Set rdCurrent = GridEX1.GetRowData(GridEX1.Row)
            If CBool(rdCurrent.Value(ColIndex)) = True Then
                rdCurrent.IconIndex(16) = 1
            Else
                rdCurrent.IconIndex(16) = 0
            End If
    End Select
    
End Sub
```

## GetSubTotal Example

Example

This example demonstrates the use of the GetSubTotal method of the JSRowData object.

```vb
Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
Dim strCaption As String
Dim curAmount As Currency
Dim colAmount As JSColumn

    If RowBuffer.RowType = jgexRowTypeGroupHeader Then
        'changing the GroupCaption to show the
        'Sum of the Amount column in the group caption
        Set colAmount = GridEX1.Columns("Amount")
        curAmount = RowBuffer.GetSubTotal(colAmount.Index, jgexSum)
        strCaption = RowBuffer.GroupCaption
        strCaption = strCaption & "(Total amount = " & _
                Format(curAmount, "Currency") & ")"
        RowBuffer.GroupCaption = strCaption
    End If
    
End Sub
```

## GridImages Example

Example

This example illustrates how JSGridImage objects are added to the JSGridImages collection at run time.

```vb
Private Sub Form_Load()

Dim gimAdd As JSGridImage

    'Setting the MaskColor for pictures
    
    'black areas in pictures will be transparent
    GridEX1.MaskColor = vbBlack
    
    'Setting image size in pixels
    GridEX1.ImageHeight = 16
    GridEX1.ImageWidth = 16
    
    'adding a picture from the resource to the JSGridImages collection
    Set gimAdd = GridEX1.GridImages.Add(LoadResPicture(101, vbResBitmap))
    'index of the JSGridImage will be 1
    GridEX1.Columns(1).HeaderIcon = gimAdd.Index
    
    'adding a picture, loading it from a file, to the JSGridImages collection
    
    'Index of the JSGridImage will be 2
    Set gimAdd = GridEX1.GridImages.Add(LoadPicture("C:\Pic.bmp"))
    
    GridEX1.Columns(1).ColumnType = jgexIcon
    
    'All cells in column 1 will display the picture loaded from the file
    GridEX1.Columns(1).DefaultIcon = gimAdd
    
    
End Sub
```

## Groups Example

Example

This example illustrates the use of the Groups property, Grouping a GridEX control by column 3 first and then by column 2 in code.

```vb
Private Sub Form_Load()

    'Grouping the control when is loaded by
    'column 3 in ascending order first
    'and by column 2 in descending order
    
    GridEX1.Groups.Add 3, jgexSortAscending
    GridEX1.Groups.Add 2, jgexSortDescending
    
    'Call RefreshGroups to force the control
    'to immediately groups the records.
    GridEX1.RefreshGroups
    
    'the current row will be the first row in the Recordset.
    'because of the group operation, the current row
    'may not be the one at the top
    'To ensure the record at the top is selected after a
    'group operation set the Row property to 1
    
    GridEX1.Row = 1
        
End Sub
```

## hImageList Example

Example

This example demonstrate how GridEX can be linked to an external ImageList control instead of using its own JSGridImages collection.

```vb
Private Sub Form_Load()

    'If you are going to use an external image list
    'in a GridEX control, be sure ImageHeight and ImageWidth
    'are the same in both controls.
    GridEX1.ImageHeight = ImageList1.ImageHeight
    GridEX1.ImageWidth = ImageList1.ImageWidth
    
    GridEX1.GridImages.hImageList = ImageList1.hImageList
    'Setting using IconIndexes
    
    GridEX1.Columns(1).HeaderIcon = ImageList1.ListImages("Header").Index
    GridEX1.Columns(2).DefaultIcon = ImageList1.ListImages("Default").Index
    
End Sub
```

## HoldFields Example

Example

This example illustrates the use of HoldFields method and HoldSortSettings property.

```vb
Public Sub RequeryGridEX()
Dim rs As Recordset

    Set rs = GridEX1.ADORecordset
    'requery the Recordset used by the GridEX control
    rs.Requery
    'to maintain the column layout
    GridEX1.HoldFields
    'to maintain SortKeys and groups after setting ADORecordset
    GridEX1.HoldSortSettings = True
    'reset the ADORecordset to the requeried Recordset
    Set GridEX1.ADORecordset = rs
        
End Sub
```

## ItemCount Example

Example

This example illustrates how ItemCount is used to tell GridEX control how many rows it should read in Unbound mode.

```vb
'DataMode property should be set to jgexUnbound 
'at design time to make this code works

Dim mData() As Variant

Private Sub Form_Load()

Dim i As Long
Dim j As Long

    Const nRows = 10
    ReDim mData(1 To 2, 1 To nRows)
    For i = 1 To nRows
        For j = 1 To 2
            mData(j, i) = "Row:=" & i & ", Col:=" & j
        Next j
    Next i
    'ItemCount property can only be set in jgexUnbound DataMode
    GridEX1.ItemCount = nRows
    'force the control to read the items.
    GridEX1.Rebind
    
    
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, _
ByVal Bookmark As Variant, ByVal Values As JSRowData)

Dim i As Long

    For i = 1 To Values.ColCount
        Values(i) = mData(i, RowIndex)
    Next
    
End Sub
```

## LayoutString Example

Example

This example demonstrates the use of LayoutString and LoadLayoutString methods.

```vb
Private Sub SaveCurrentLayoutInDB(LayoutID As String)

Dim Rs As Recordset
Dim cn As Connection

    Const DBPath = "C:\JSNWind.MDB"
    Set cn = New Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath
    
    Set Rs = New Recordset
    Rs.CursorLocation = adUseClient
    Rs.Open "SELECT * FROM Layouts WHERE LayoutID='" & LayoutID _
    & "'", cn, adOpenStatic, adLockOptimistic
    
    If Not Rs.EOF Then
        'Layout Data is a BLOB field
        Rs.Fields("LayoutData").Value = GridEX1.LayoutString(False)
        Rs.Update
    End If
    
End Sub

Private Sub LoadLayoutFromDB(LayoutID As String)
Dim Rs As Recordset
Dim cn As Connection

    Const DBPath = "C:\Layouts.MDB"
    Set cn = New Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath
    
    Set Rs = New Recordset
    Rs.CursorLocation = adUseClient
    Rs.Open "SELECT * FROM Layouts WHERE LayoutID='" _
    & LayoutID & "'", cn, adOpenStatic, adLockReadOnly
    
    If Not Rs.EOF Then
        GridEX1.LoadLayoutString Rs.Fields("LayoutData").Value
    End If
    
End Sub
```

## MoveToBookmark Example

Example

This example demonstrates how MoveToBookmark method can be used to set the current record in a GridEX control when you know its bookmark.

```vb
Private Sub FindRecordInGridEX(FirstName As String, LastName As String)
Dim Rs As Recordset

    Set Rs = GridEX1.Recordset
    Rs.FindFirst "[FirstName] = '" & FirstName _
    & "' AND [LastName] = '" & LastName & "'"
    If Rs.NoMatch Then
        MsgBox "No record found."
    Else
        GridEX1.MoveToBookmark Rs.Bookmark
    End If

End Sub
```

## NotInList Example

Example

This example demonstrates the use of the NotInList event to add new entries into a GridEX control acting as the drop down control for a column in GridEX1.

```vb
Private Sub GridEX1_NotInList(ByVal ColIndex As Integer, _
ByVal NewData As String, ByVal DataAdded As JSRetBoolean)

Dim rs As Recordset
Dim strCustomerID As String

    If ColIndex = 2 Then
        If MsgBox("'" & NewData & "' is not present in the list." _
        & vbCrLf & vbCrLf & "Do you want to add it as a Customer?", _
        vbYesNo) = vbYes Then
            strCustomerID = InputBox("Enter the Customer ID for '" _
            & NewData & "'", "Enter Customer ID")
            If Len(strCustomerID) > 0 Then
                Set rs = Me.grCustDropDown.ADORecordset
                rs.AddNew
                rs![CustomerID] = strCustomerID
                rs![CompanyName] = NewData
                rs.Update
                Me.grCustDropDown.SearchNewRecords
                DataAdded = True
            End If
        End If
    End If
            
End Sub
```

## PreviewColumn Example

Example

This example sets a column named "Notes" as the column to be displayed as the PreviewRow.

```vb
Private Sub Form_Load()

    'preview row displays three lines at most.
    GridEX1.PreviewRowLines = 3
    'Notes column will be displayed in the Preview row
    GridEX1.PreviewColumn = "Notes"    
    GridEX1.PreviewRowIndent = 300 'twips
    'Hide "Notes" column because it'll be
    'displayed in the preview row anyway.
    GridEX1.Columns("Notes").Visible = False
    
End Sub
```

## PrinterProperties Example

Example

This example changes some of the PrinterProperties properties of a GridEX control before printing the contents of the GridEX control .

```vb
Private Sub PrintGridEX()

    'setting some properties in before
    'printing the GridEX control
    
    With GridEX1.PrinterProperties
        'all columns in the grid will be
        'printed in the same page.
        .FitColumns = True
        'column headers will be printed in
        'every page and not just in the first one.
        .RepeatHeaders = True
        'Preview rows are also printed in
        'case the GridEX control show Preview rows.
        .PrintPreviewRows = True
        'System colors will be translated to
        'black, white and gray colors.
        .TranslateColors = True
        'frozen columns will be printed in
        'every page and not just in the first one.
        .RepeatFrozenCols = True
    End With
    
    'the page setup dialog will be shown before printing.
    GridEX1.PrintGrid True
    
End Sub
```

## PrinterPropertiesMargins Example

Example

This example demonstrates how page margins can be set in the PrinterProperties object.

```vb
Private Sub PageMarginsSetup()

Dim pp As JSPrinterProperties

    Set pp = GridEX1.PrinterProperties

    'setting page margins

    pp.LeftMargin = 0.5 * 1440 'half inch
    pp.TopMargin = 1 * 1440 'one inch
    pp.RightMargin = 0.5 * 1440 'half inch
    pp.BottomMargin = 1 * 1440 'one inch
    
    'setting Header and Footer distance
    
    'HeaderDistance should be less than TopMargin
    pp.HeaderDistance = 0.5 * 1440 'half inch
    'FooterDistance should be less than BottomMargin
    pp.FooterDistance = 0.5 * 1440 'half inch
    
    'setting header strings
    pp.HeaderString(jgexHFCenter) = "Header at the center"
    pp.HeaderString(jgexHFLeft) = "Header at the left"
    pp.HeaderString(jgexHFRight) = "Header at the right"
    
    'setting footer strings
    pp.FooterString(jgexHFCenter) = "Footer at the center"
    pp.FooterString(jgexHFLeft) = "Footer at the left"
    pp.FooterString(jgexHFRight) = "Footer at the right"
    
End Sub
```

## PrintingProgress Example

Example

This example demonstrates how a custom "Print Cancel" dialog can be implemented.

 To try this example follow the next steps:

 Step 1) Start a new project.

 Step 2) Select "Janus GridEX 2000 Control" from VB Components dialog.

 Step 3) Place one GridEX control in Form1

 Step 4) At design time, Set DatabaseName and RecordSource properties in GridEX1 control

 Step 5) Add a CommandButton named cmdPrint and change its caption property to "Print..."

 Step 6) Paste the following code in the declarations section

```vb
Private Sub cmdPrint_Click()

    Load frmPrintCancel
    frmPrintCancel.Canceled = False
    GridEX1.PrinterProperties.PrintProgressDialog = False
    With frmPrintCancel
        .lblPrintWhere = "Printing Document in '" _
        & GridEX1.PrinterProperties.DeviceName & "'"
        .lblPercent = "Paginating the document"
        .Show , Me
        DoEvents
    End With
    GridEX1.PrintGrid
    Unload frmPrintCancel
    
End Sub

Private Sub GridEX1_BeforePrintPage _
(ByVal PageNumber As Long, ByVal nPages As Long)

    frmPrintCancel.lblPage = "Page " & PageNumber & " of " & nPages
    
End Sub

Private Sub GridEX1_PrintingProgress _
(ByVal PrintProgress As Single, ByVal Cancel As GridEX20.JSRetBoolean)

    frmPrintCancel.lblPercent = Format(PrintProgress, "Percent")
    If frmPrintCancel.Canceled = True Then
        Cancel = True
    End If
    
End Sub
```

Step 7) Add a new Form to the project and set its name as frmPrintCancel

 Step 8) Add a CommandButton to frmPrintCancel and change the following properties:

Name = "cmdCancel"

Caption = "Cancel"

Cancel = True

Default = True

Step 9) Add three labels: lblPrintWhere, lblPage and lblPercent and clear their caption property

 Step 10) Paste the following code into the declarations section of frmPrintCancel

```vb
Public Canceled As Boolean

Private Sub cmdCancel_Click()
    
    Hide
    Canceled = True
    
End Sub 
``` 
 Step 11) Run the Project and click at the Print button.

## PrintPreview Example

Example

This example demonstrates how a "Print Preview Window" can be implemented with a GEXPreview Control.

 To try this example follow the next steps:

 Step 1) Start a new project.

 Step 2) Select "Janus GridEX 2000 Control" from VB Components dialog.

 Step 3) Place one GridEX control in Form1

 Step 4) At design time, Set DatabaseName and RecordSource properties in GridEX1 control

 Step 5) Add a CommandButton named cmdPrintPreview and change its caption property to "Print Preview..."

 Step 6) Paste the following code in the declarations section

```vb
Private Sub cmdPrintPreview_Click()

    Load frmPreview
    frmPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    GridEX1.PrintPreview frmPreview.grPrev
    frmPreview.Show
    Me.Hide
    
End Sub
```

Step 7) Add a new Form to the project and set its name as frmPreview

 Step 8) Add GEXPreview control to the Form and change its name to grPrev

 Step 9) Paste the following code into the declarations section of frmPrintCancel

```vb
Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        grPrev.Move 0, 0, ScaleWidth, ScaleHeight
    End If
End Sub
``` 
```vb
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If Me.WindowState = 0 Then
        Form1.Move Left, Top, Width, Height
    Else
        Form1.WindowState = Me.WindowState
    End If
    'Show the form that called the Print Preview window again
    Form1.Show
    Hide    
End Sub
``` 
```vb
Private Sub grPrev_OnCloseClick()
    Unload Me
 End Sub
``` 
 Step 10) Run the Project and click the Print Preview button.

## RebindSample Example

Example

This example, changes some of the data source related properties of a GridEX control and calls the Rebind method to force the control to create its underlying Recordset based on the new settings.

```vb
Public Sub RebindGridEX()
Dim strConn As String

'Change DBPath const to a valid mdb file name
    Const DBPath = "C:\JSNWind.mdb"
    Const SQLStatement = "SELECT * FROM Orders"
    If GridEX1.DataMode = jgexADO Then
        With GridEX1
            strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
            & DBPath
            .CursorLocation = jgexUseClient
            .DatabaseName = strConn
            .LockType = jgexLockOptimistic
            .RecordsetType = jgexRSADOStatic
            .RecordSource = SQLStatement
            .ClearFields
            .Rebind
        End With
    ElseIf GridEX1.DataMode = jgexDAO Then
        With GridEX1
            .DatabaseName = DBPath
            .Exclusive = False
            .LockType = jgexLockOptimistic
            .ReadOnly = False
            .RecordsetType = jgexRSDAODynaset
            .RecordSource = SQLStatement
            .ClearFields
            .Rebind
        End With
    End If

End Sub
```

## Recordset Example

Example

This example opens an Recordset and sets it to the Recordset property of a GridEX control.

 To try this example, paste the code into the declarations section of a form with a GridEX control in it and add "Microsoft DAO" in the References Dialog of the project.

```vb
Private Sub Form_Load()

Dim DB As Database
Dim RS As Recordset

    'Change DBPath const to a valid mdb file name
    Const DBPath = "C:\JSNWind.mdb"
    Const SQLStatement = "SELECT * FROM Customers"
    
    'Opening the Recordset
    Set DB = OpenDatabase(DBPath)
    
    Set RS = DB.OpenRecordset(SQLStatement)
    
    'Setting Recordset property is equivalent to call Rebind method.
    Set GridEX1.Recordset = RS
    
End Sub
```

## Redistributables Example

Files required for Redistribution

Janus GridEX 2000 Redistributable File:

 **GRIDEX20.OCX**

 The dependency files required to be redistributed with applications using the Janus GridEX 2000 ActiveX Control vary according to the DataMode in which the control is being used.

 **For All Data Modes (Unbound/DAO & ADO) :**The following files are required by the control regardless the DataMode used.  
Visual Basic 6 Runtime Library

MSVBVM60.DLL 6.00.8495

OLE 2.4 Files

OLEAUT32.DLL 2.40.4275

OLEPRO32.DLL 5.0.4275

STDOLE2.TLB 2.40.4275

ASYCFILT.DLL 2.40.4275

COMCAT.DLL 4.71 **For ADO DataMode:** In addition to the files outlined above, you must redistribute the Microsoft Data Access Components Files (Mdac_typ.exe).

 You can download the latest available MDAC Files from Microsoft's **Universal Data Access Web Site [http://www.microsoft.com/data/](http://www.microsoft.com/data/).**

 For detailed information about installing MDAC please refer to the following Microsoft Knowledge Base article:

 Article ID: Q232053 [INFO: List of Files Installed by MDAC 2.1 (GA)](http://support.microsoft.com/support/kb/articles/Q232/0/53.ASP)

 **For DAO DataMode:**The DAO 3.6 (Dao360.dll) library files are required by the Janus GridEX 2000 control when DAO DataMode is being used.For detailed information about installing DAO 3.6 please refer to the following Microsoft Knowledge Base article:  
 Article ID: Q233002 [HOWTO: Redistribute DAO 3.6](http://support.microsoft.com/support/kb/articles/Q233/0/02.ASP?LNG=ENG&SA=PER)  
 Note: DAO 3.6 & MDAC are covered by their own License Agreement. Please consult your Microsoft Product Documentation and/or End User's License Agreement (EULA) for more information about licensing and redistributing these files.

## Redraw Example

Example

This example uses the Redraw property to avoid flickering of the control when data in the underlying Recordset of a GridEX control is changed.

```vb
Private Sub ChangingValuesInCode()
Dim simTemp As JSSelectedItem
Dim rs As Recordset

    GridEX1.Redraw = False
    Set rs = GridEX1.ADORecordset
    
    For Each simTemp In GridEX1.SelectedItems
        If simTemp.RowType = jgexRowTypeRecord Then
            rs.Bookmark = simTemp.Bookmark
            rs![Status] = 0
            rs.Update
            GridEX1.RefreshRowBookmark simTemp.Bookmark
        End If
    Next
    'always set Redraw property back to True.
    GridEX1.Redraw = True

End Sub
```

## RefreshRowIndex Example

Example

This examples demonstrates, changes values of a GridEX control and calls the RefreshRowIndex method to force the GridEX control to re-read those values and shows the new data immediately.

```vb
'DataMode property should be set to jgexUnbound
'at design time to make this code works

Dim mData() As Variant

Private Sub cmdUpdateValues_Click()

    'This procedure updates values directly into the mData array.
    'After updating a row, it calls RefreshRowIndex method
    'to force the GridEX control to refresh the values in the updated row.
    
    'updating row 4
    mData(1, 4) = "Updated Value"
    mData(2, 4) = Now
    
    GridEX1.RefreshRowIndex 4
    
End Sub

Private Sub Form_Load()
Dim i As Long
Dim j As Long

    Const nRows = 10
    ReDim mData(1 To 2, 1 To nRows)
    For i = 1 To nRows
        For j = 1 To 2
            mData(j, i) = "Row:=" & i & ", Col:=" & j
        Next j
    Next i
    'ItemCount property can only be set in jgexUnbound DataMode
    GridEX1.ItemCount = nRows
    'force the control to read the items.
    GridEX1.Rebind
    
    
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, _
ByVal Bookmark As Variant, ByVal Values As JSRowData)

Dim i As Long
    For i = 1 To Values.ColCount
        Values(i) = mData(i, RowIndex)
    Next
End Sub
```

## RowDrag Example

Example

This example performs a simple OLE Drag operation with a GridEX control.

 To run this test create a new project, place a GridEX control into the form, select a valid DatabaseName and RecordSource for the control and paste the following code into the form module.

```vb
Private Sub Form_Load()

   'Set DetectRowDrag to true in order
   'to receive the RowDrag event.
    GridEX1.DetectRowDrag = True
    
End Sub

Private Sub GridEX1_OLESetData(Data As JSDataObject, _
DataFormat As Integer)

    If DataFormat = vbCFText Then
        Data.SetData GridEX1.GetClipString(True), vbCFText
    End If
    
End Sub

Private Sub GridEX1_OLEStartDrag(Data As JSDataObject, _
AllowedEffects As Long)

    AllowedEffects = vbDropEffectCopy
    Data.SetData , vbCFText
    
End Sub

Private Sub GridEX1_RowDrag(ByVal Button As Integer, _
ByVal Shift As Integer)

    'The ExpandSelection is called to expand and
    'select all the children rows of selected group rows.
    GridEX1.ExpandSelection
    'Start the OLE Drag operation
    GridEX1.OLEDrag
    
End Sub
```

## RowExpanded Example

Example

This example expands all the groups in the first level while higher levels are collapsed.

```vb
Private Sub ExpandLevel1()
Dim i As Long

    'expanding all groups in level 1 and collapsing other levels
    If GridEX1.Groups.Count > 0 Then
        GridEX1.Redraw = False
        For i = 1 To GridEX1.RowCount
            If GridEX1.IsGroupItem(i) Then
                If GridEX1.GroupRowLevel(i) = 1 Then
                    GridEX1.RowExpanded(i) = True
                Else
                    GridEX1.RowExpanded(i) = False
                End If
            End If
        Next
        GridEX1.Redraw = True
    End If
    
End Sub
```

## RowFormat Example

Example

This example demonstrates the use of the RowFormat event to format individual cells or rows or to change the caption of a group header to display some extra information.

```vb
Private Sub Form_Load()
Dim fmsTemp As JSFormatStyle

    Set fmsTemp = GridEX1.FormatStyles.Add("Alert")
    fmsTemp.ForeColor = vbRed
    
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As JSRowData)

Dim strNewGroupCaption As String
    If RowBuffer.RowType = jgexRowTypeGroupHeader Then
        'Changing group caption to give more
        'information about the group to users
        strNewGroupCaption = RowBuffer.GroupCaption
        'show the count of records in the group
        strNewGroupCaption = strNewGroupCaption _
        & " (" & RowBuffer.RecordCount & " item(s)) "
        'also show the total amount for records
        'in the group assuming Amount is in column 4
        strNewGroupCaption = strNewGroupCaption _
        & " (Total Amount = " & Format(RowBuffer.GetSubTotal(4, jgexSum), "Currency") & ")"
        'If GroupCaption was originally "Customer: Alfreds Futterkiste"
        'the new GroupCaption will be something like:
        '"Customer: Alfreds Futterkiste (8 item(s)) (Total Amount = $456.00)"
        RowBuffer.GroupCaption = strNewGroupCaption
    ElseIf RowBuffer.RowType = jgexRowTypeRecord Then
        'Setts icon index to a column based on the value of another column
        If RowBuffer(2) = True Then
            RowBuffer.IconIndex(1) = 1
        End If
        
        'Hides preview rows for rows where value in column 5 is False
        If RowBuffer(5) = False Then
            RowBuffer.PreviewRowVisible = False
        End If
        
        'assuming column 4 display a "Due Date", and Due Date
        'is expired set the format to the cell to Alert
        If DateDiff("d", RowBuffer(4), Date) > 0 Then
            RowBuffer.CellStyle(4) = "Alert"
        End If
    End If
        
End Sub
```

## RowIndex Example

Example

This example gets the underlying Recordset of a GridEX control and synchronize the current record of the GridEX control with the current record in the Recordset.

```vb
Private Sub SynchronizeRecordset()

Dim rs As Recordset
Dim lngRowIndex As Long
Dim vBookmark As Variant
Dim i As Long

    'synchronizing the recordset to the
    'current record in a GridEX control
    
    'Gets the RowIndex of the current row
    lngRowIndex = GridEX1.RowIndex(GridEX1.Row)
    
    'Gets the bookmark
    vBookmark = GridEX1.RowBookmark(lngRowIndex)
    
    'get the Recordset used by the GridEX control
    
    Set rs = GridEX1.ADORecordset
    
    'synchronizes the current record in the Recordset
    'with the current record in the GridEX control
    rs.Bookmark = vBookmark
    
    For i = 0 To rs.Fields.Count - 1
        Debug.Print rs.Fields(i).Value
    Next
    
End Sub
```

## RowSelected Example

Example

This example selects all the rows in a GridEX control.

```vb
Private Sub SelectAllRecords()

    Dim i As Long
    'make sure the control has loaded all records
    If Not GridEX1.FullyLoaded Then
        GridEX1.LoadEntireRecordset
    End If
    GridEX1.MultiSelect = True
    For i = 1 To GridEX1.RowCount
        GridEX1.RowSelected(i) = True
    Next
    
End Sub
```

## SaveLayout Example

Example

This example demonstrates the use of SaveLayout and LoadLayout methods.

```vb
This example demonstrates the use of SaveLayout and LoadLayout methods.

Private Sub SaveCurrentLayout()
On Error GoTo EH_SaveCurrentLayout
Dim strFileName As String

    'cdlSaveAs is a common dialog control
    cdlSaveAs.ShowSave
    GridEX1.SaveLayout cdlSaveAs.FileName
    Exit Sub
    
EH_SaveCurrentLayout:
    If Err.Number <> cdlCancel Then
        MsgBox Err.Description
    End If
End Sub

Private Sub LoadLayout()
On Error GoTo EH_OpenLayout
Dim strFileName As String
Dim strConnString As String

    strConnString = GetSetting("MyApp", "Settings", "DBPath")
    strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
    & strConnString
    'cdlOpen is a common dialog control
    cdlOpen.ShowOpen
    GridEX1.LoadLayout cdlSaveAs.FileName, True, strConnString
    Exit Sub
    
EH_OpenLayout:
    If Err.Number <> cdlCancel Then
        MsgBox Err.Description
    End If
    
End Sub
```

## ScrollTollTips Example

Example

This example sets a column named "CompanyName" as the column to be displayed as ToolTip while the user drags the scroll thumb along the vertical scroll bar.

```vb
Private Sub Form_Load()

    GridEX1.ContinuousScroll = False
    GridEX1.ScrollToolTips = True
    GridEX1.ScrollToolTipColumn = "CompanyName"
    
End Sub
```

## SearchNewRecords Example

Example

This example demonstrates the use of the SearchNewRecords method when records are added directly into the recordset used by the GridEX control.

```vb
Private Sub AddRecords()
Dim rs As Recordset

    Set rs = GridEX1.ADORecordset
    rs.AddNew
    rs.Fields(0).Value = "Field0Value"
    rs.Fields(1).Value = "Field1Value"
    rs.Update
    'after adding records in code
    'SearchNewRecords method should be called
    'to force the GridEX control to read those
    'new records
    GridEX1.SearchNewRecords
    
End Sub
```

## SelectedItems Example

Example

This example the value of each column in the selected rows in a GridEX control.

```vb
Private Sub GetValuesFromSelectedRows()
Dim simTemp As JSSelectedItem
Dim RowData As JSRowData
Dim i As Long

    If GridEX1.MultiSelect = False Then
        MsgBox "SelectedItems property can not be used.", vbExclamation
    Else
        For Each simTemp In GridEX1.SelectedItems
            Set RowData = GridEX1.GetRowData(simTemp.RowPosition)
            Debug.Print "Row Selected:=" & simTemp.RowPosition
            For i = 1 To RowData.ColCount
                Debug.Print "Column(" & i & ")", RowData(i)
            Next
        Next
    End If
    
End Sub
```

## SelectionChange Example

Example

This example demonstrates how SelectionChange event can be used to inform the user how many rows and/or which rows has been selected.

```vb
Private Sub GridEX1_SelectionChange()
Dim SelItem As JSSelectedItem
Dim rd As JSRowData
Dim strCaption As String

    strCaption = GridEX1.SelectedItems.Count & " items selected"
    strCaption = strCaption & vbCrLf & vbCrLf
    For Each SelItem In GridEX1.SelectedItems
        Set rd = GridEX1.GetRowData(SelItem.RowPosition)
        strCaption = strCaption & "Item " & rd.Value(1) & vbCrLf
    Next
    lblSelItems.Caption = strCaption
    
End Sub
```

## SortKeys Example

Example

This example illustrates the use of the SortKeys property, first sorting a GridEX control by column 3 and then by column 2 in code.

```vb
Private Sub Form_Load()

    'sorting the control when is loaded by column 3 in ascending 
    'order (first) and then by column 2 in descending order (second)
    
    GridEX1.SortKeys.Add 3, jgexSortAscending
    GridEX1.SortKeys.Add 2, jgexSortDescending
    
    'Call RefreshSort to force the control to
    'immediately sort the records.
    GridEX1.RefreshSort
    
    'the current row will be the first row in the recordset.
    'because of the sort operation, the current row
    'may not be the one at the top
    
    'To ensure the record at the top is selected after a
    'sort operation set the Row property to 1
    
    GridEX1.Row = 1
    
    
End Sub
```

## UnboundEvents Example

Example

This examples illustrates how Unbound events should be handled when the GridEX control is working in Unbound mode.

```vb
Dim mavarData() As Variant
Dim mRowCount As Long

Private Sub Form_Load()

Dim colTemp As JSColumn

    'Set the Column layout (This can be done at design time
    
    GridEX1.Columns.Clear
    Set colTemp = GridEX1.Columns.Add("Last Name", _
    jgexText, jgexEditTextBox, "LastName")
    colTemp.Width = 3000
    Set colTemp = GridEX1.Columns.Add("First Name", _
    , , "FirstName")
    colTemp.Width = 3000
    Set colTemp = GridEX1.Columns.Add("Age", , , "Age")
    colTemp.Width = 600
    colTemp.SortType = jgexSortTypeNumeric
    
    'Redim our dataset
    'first dimension is used for columns
    'second dimension is used for rows
    mRowCount = 3
    ReDim mavarData(1 To 3, 1 To mRowCount)
    'Filling the dataset
    mavarData(1, 1) = "Smith"
    mavarData(2, 1) = "John"
    mavarData(3, 1) = 22
    
    mavarData(1, 2) = "Wright"
    mavarData(2, 2) = "Steven"
    mavarData(3, 2) = 35
    
    mavarData(1, 3) = "Connors"
    mavarData(2, 3) = "Henry"
    mavarData(3, 3) = 47
    
    'Tell GridEX how many rows you have in your dataset
    GridEX1.ItemCount = mRowCount
    
    'See Unbound events
    
    
End Sub

Private Sub GridEX1_UnboundAddNew _
(ByVal NewRowBookmark As JSRetVariant, ByVal Values As JSRowData)

    'Increment rowcount and redim the array
    mRowCount = mRowCount + 1
    ReDim Preserve mavarData(1 To 3, 1 To mRowCount)
    
    'set values in the last row in the array
    mavarData(1, mRowCount) = Values(1)
    mavarData(2, mRowCount) = Values(2)
    mavarData(3, mRowCount) = Values(3)
    
    
End Sub

Private Sub GridEX1_UnboundDelete _
(ByVal RowIndex As Long, ByVal Bookmark As Variant)

Dim i As Long
Dim j As Long

    
    'use the BeforeDelete event if you
    'want to prevent deletion of a row
    
    'First shift the rows
    For i = RowIndex To mRowCount - 1
        For j = 1 To 3
            mavarData(j, i) = mavarData(j, i + 1)
        Next
    Next
    
    'decrement rowcount and redim array
    mRowCount = mRowCount - 1
    If mRowCount > 0 Then ReDim Preserve _
    mavarData(1 To 3, 1 To mRowCount)
    
End Sub

Private Sub GridEX1_UnboundReadData _
(ByVal RowIndex As Long, ByVal Bookmark As Variant, _
ByVal Values As JSRowData)

    'don't worry if the grid is sorted or if the
    'column positions are changed by the user
    
    Values(1) = mavarData(1, RowIndex)
    Values(2) = mavarData(2, RowIndex)
    Values(3) = mavarData(3, RowIndex)
    
End Sub

Private Sub GridEX1_UnboundUpdate _
(ByVal RowIndex As Long, ByVal Bookmark As Variant, _
ByVal Values As JSRowData)

    mavarData(1, RowIndex) = Values(1)
    mavarData(2, RowIndex) = Values(2)
    mavarData(3, RowIndex) = Values(3)
    
End Sub
```

## ValueList Example

Example

This example ilustrates how a ValueList can be filled to be used as the dropdown list in a column.

```vb
Private Sub Form_Load()
Dim colTemp As JSColumn
Dim vl As JSValueList

    Set colTemp = GridEX1.Columns("Status")
    colTemp.HasValueList = True
    
    'Filling the value list
    Set vl = colTemp.ValueList
    'it is assumed that GridImages collection contains at 3 pictures
    vl.Add 0, "Done", 1
    vl.Add 1, "Incomplete", 2
    vl.Add 2, "Not started", 3
    
    'changing the ColumnType to show icons and text
    
    colTemp.ColumnType = jgexIconAndText
    
    'making the column a drop down list
    
    colTemp.EditType = jgexEditDropDown
    
End Sub
```

## View Example

Example

This example demonstrates the use of the View property and some other properties that depend on the View used.

```vb
Private Sub ViewProperties(View As String)

    Select Case View
    Case "Table"
        With GridEX1
            .View = jgexTable
            'Setting some properties that are used
            'only when GridEX is in table view
            .AllowRowSizing = True
            .GroupByBoxVisible = True
            .RowHeaders = True
            .ColumnAutoResize = True
            .EmptyRows = True
            .NewRowPos = jgexTop
            .GridLines = jgexGLBoth
            .GridLineStyle = jgexGLSSmallDots
            .PreviewColumn = "Notes"
            .PreviewRowLines = 3
        End With
    Case "Card"
        With GridEX1
            .View = jgexCard
            'Setting some properties that are used
            'only when GridEX is in card view
            .AllowCardSizing = True
            .CardBorders = False
            .CardCaptionPrefix = "Item:"
            .CardSpacing = 300 'twips
            .CardWidth = 3000 'twips
            .ShowEmptyFields = True
        End With
    End Select
    
End Sub
```

## WordWrap Example

Example

This example converts a column named "Notes" in a GridEX control into a word-wrap column.

 It also specify how many rows should be displayed when the control is in card view.

```vb
Private Sub Form_Load()
Dim colTemp As JSColumn

    GridEX1.View = jgexCard
    Set colTemp = GridEX1.Columns("Notes")
    colTemp.WordWrap = True
    colTemp.MinRowsInCardView = 2
    colTemp.MaxRowsInCardView = 6
    
End Sub
```
