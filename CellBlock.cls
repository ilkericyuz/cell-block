VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' Cell Block Class defines a cell block and lets you reach values in it
''' A Cell Block is a group of cells which can be defined as a table.
'''
''' Referencing from an initial cell, class finds the edges of the cell block and defines:
''' topRow:         Range object which refers to top row of the cell block
''' bottomRow:      Range object which refers to bottom row of the cell block
''' leftColumn:     Range object which refers to left column of the cell block
''' rightColumn:    Range object which refers to right column of the cell block
''' titleRow:       Range object which refers to entire top row of the cell block.
''' titleColumn:    Range object which refers to entire left column of the cell block
''' titleRowCells:  Range object which refers to cell block part of the title row
''' titleColumnCells:   Range object which refers to cell block part of the title column
''' top:            Row number of the top row
''' left:           Column number of the left column
''' width:          Width of the cell block
''' height:         Height of the cell block
''' size:           Total number of cells in the cell block
''' activeRowCells: Dictionary object which let you reach any value in the selected row
'''                 using title of the row as key
'''     Example:
'''         Date        A   B   C   D
'''         20170101    1   2   3   ronaldinho
'''         20170102   [4]  5   6   messi
'''         20170103    7   8   9   c. ronaldo
'''
'''         As 4 is selected, activeRowCells("D") gives you messi, activeRowCells("C") gives you 6
'''
''' activeColumnCells:  Just like activeRowCells, but refers to cell values according to their
'''                     title columns

Option Explicit

Dim pActiveCell As Range, pFirstCell As Range, pLastCell As Range, _
    pTopLeftCell As Range, pTopRightCell As Range, _
    pBottomLeftCell As Range, pBottomRightCell As Range
    
Dim pActiveSheet As Worksheet

Dim pTopRow As Range, pBottomRow As Range, pLeftColumn As Range, pRightColumn As Range
Dim pTitleRow As Range, pTitleColumn As Range, pTitleRowCells As Range, pTitleColumnCells As Range

Dim pTop As Double, pLeft As Double, pWidth As Double, pHeight As Double, pSize As Double

Dim pActiveRowCells As Dictionary, pActiveColumnCells As Dictionary

Dim pRangeBlock As Range

Private Sub class_initialize()

'    Set pActiveCell = Selection
'    Set pActiveSheet = ActiveSheet
    
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    Set pActiveSheet = Nothing
    Set pActiveRowCells = Nothing
    Set pActiveColumnCells = Nothing
    
End Sub


' =========================================
'   INITIATE OBJECT PROPERTIES
' -----------------------------------------
Public Sub InitiateProperties(Optional theSelection As Range, Optional theActiveSheet As Worksheet)
    ''' initiates cell block properties.
    ''' a remote cell or current selected cell can be assigned as a member of desired cell block
    ''' another worksheet or current worksheet can be assigned as desired worksheet
    
On Error Resume Next

    ' if theActiveSheet is not sent, assign current selected worksheet to the pActiveSheet
    If theActiveSheet Is Nothing Then
        Set pActiveSheet = activeSheet
    Else
        Set pActiveSheet = theActiveSheet
        pActiveSheet.Activate
    End If

    ' if selection is not sent, assign current selected cell to the pActiveCell
    If theSelection Is Nothing Then
        Set pActiveCell = Selection
    Else
        Set pActiveCell = theSelection
    End If
    
    
    Set pRangeBlock = pActiveCell
    
    Call SetTopRow
    Call SetBottomRow
    Call SetLeftColumn
    Call SetRightColumn
    
    Set pTopLeftCell = activeSheet.Cells(pTopRow.Row, pLeftColumn.Column)
    Set pTopRightCell = activeSheet.Cells(pTopRow.Row, pRightColumn.Column)
    Set pBottomLeftCell = activeSheet.Cells(pBottomRow.Row, pLeftColumn.Column)
    Set pBottomRightCell = activeSheet.Cells(pBottomRow.Row, pRightColumn.Column)
    
    Set pFirstCell = pTopLeftCell
    Set pLastCell = pBottomRightCell
    
    'title row must be called after firstcell and lastcell are defined
    Call SetTitleRow
    Call SetTitleColumn
    Call SetTitleRowCells
    Call SetTitleColumnCells
    
    
    pTop = pFirstCell.Row - 1
    pLeft = pFirstCell.Column - 1
    pHeight = pLastCell.Row - pFirstCell.Row + 1
    pWidth = pLastCell.Column - pFirstCell.Column + 1
    pSize = pHeight * pWidth
    
    Call SetActiveRowCells
    Call SetActiveColumnCells
    
End Sub

Public Sub InitiatePropertiesByConnectedCells(Optional theSelection As Range, Optional theActiveSheet As Worksheet)
    
    ''' !!! THIS SUB IS IN BETA, MAY CAUSE INFINITE LOOPS OR CRASH EXCEL !!!
    
    ''' initiates cell block properties.
    ''' a remote cell or current selected cell can be assigned as a member of desired cell block
    ''' another worksheet or current worksheet can be assigned as desired worksheet
    
On Error Resume Next
    
    ' if theActiveSheet is not sent, assign current selected worksheet to the pActiveSheet
    If theActiveSheet Is Nothing Then
        Set pActiveSheet = activeSheet
    Else
        Set pActiveSheet = theActiveSheet
        pActiveSheet.Activate
    End If
    
    ' if selection is not sent, assign current selected cell to the pActiveCell
    If theSelection Is Nothing Then
        Set pActiveCell = Selection
    Else
        Set pActiveCell = theSelection
    End If

    Call DefineByConnectedCells(theSelection)
    
    Set pTopLeftCell = pActiveSheet.Cells(pTopRow.Row, pLeftColumn.Column)
    Set pTopRightCell = pActiveSheet.Cells(pTopRow.Row, pRightColumn.Column)
    Set pBottomLeftCell = pActiveSheet.Cells(pBottomRow.Row, pLeftColumn.Column)
    Set pBottomRightCell = pActiveSheet.Cells(pBottomRow.Row, pRightColumn.Column)
    
    Set pFirstCell = pTopLeftCell
    Set pLastCell = pBottomRightCell
    
    'title row must be called after firstcell and lastcell are defined
    Call SetTitleRow
    Call SetTitleColumn
    Call SetTitleRowCells
    Call SetTitleColumnCells
    
    pTop = pFirstCell.Row - 1
    pLeft = pFirstCell.Column - 1
    pHeight = pLastCell.Row - pFirstCell.Row + 1
    pWidth = pLastCell.Column - pFirstCell.Column + 1
    pSize = pHeight * pWidth
    
    Call SetActiveRowCells
    Call SetActiveColumnCells

End Sub
    
    
' =========================================
'   GETTERS
' -----------------------------------------
    
Public Property Get activeCell() As Range
    Set activeCell = pActiveCell
End Property

Public Property Get topRow() As Range
    Set topRow = pTopRow
End Property

Public Property Get bottomRow() As Range
    Set bottomRow = pBottomRow
End Property

Public Property Get leftColumn() As Range
    Set leftColumn = pLeftColumn
End Property

Public Property Get topLeftCell() As Range
    Set topLeftCell = pTopLeftCell
End Property

Public Property Get topRightCell() As Range
    Set topRightCell = pTopRightCell
End Property

Public Property Get bottomLeftCell() As Range
    Set bottomLeftCell = pBottomLeftCell
End Property

Public Property Get bottomRightCell() As Range
    Set bottomRightCell = pBottomRightCell
End Property

Public Property Get firstCell() As Range
    Set firstCell = pFirstCell
End Property

Public Property Get lastCell() As Range
    Set lastCell = pLastCell
End Property

Public Property Get top() As Double
    top = pTop
End Property

Public Property Get Left() As Double
    Left = pLeft
End Property

Public Property Get height() As Double
    height = pHeight
End Property

Public Property Get size() As Double
    size = pSize
End Property

Public Property Get CountA() As Double
    for each cell in
End Property


Public Property Get width() As Double
    width = pWidth
End Property

Public Property Get rangeBlock() As Range
    Set rangeBlock = pRangeBlock
End Property

Public Property Get titleRow() As Range
    Set titleRow = pTitleRow
End Property

Public Property Get titleColumn() As Range
    Set titleColumn = pTitleColumn
End Property

Public Property Get titleRowCells() As Range
    Set titleRowCells = pTitleRowCells
End Property

Public Property Get titleColumnCells() As Range
    Set titleColumnCells = pTitleColumnCells
End Property

Public Property Get activeRowCells() As Dictionary
    Set activeRowCells = pActiveRowCells
End Property



' =========================================
'   NEW COLUMN
' -----------------------------------------
' NewColumn_Recordset is not written because it would be the same as NewRow_Recordset except for insert direction.
' While NewRow_Recordset shifts cells down in insert, NewColumn_Recordset should shift them right.


Public Sub NewColumn_Dictionary(theValues As Dictionary, Optional theColumnTitle As Variant = "", Optional isInsert As Boolean = False)
    ''' adds new column to current cell block
    
    Dim key As Variant, rowCounter As Double, columnCounter As Double, rowCount As Double
    
    columnCounter = 0
    rowCount = 0
    
    'find how many columns will be pasted
    If theColumnTitle <> "" Then
        rowCount = theValues.Count + 1
    Else
        rowCount = theValues.Count
    End If
    
    
    If isInsert Then
        pActiveSheet.Range(Cells(pActiveCell.Row, pActiveCell.Column + 1), _
                            Cells(pActiveCell.Row + rowCount - 1, pActiveCell.Column + 1)). _
                            Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        pActiveSheet.Range(Cells(pActiveCell.Row, pActiveCell.Column + 1), _
                            Cells(pActiveCell.Row + rowCount - 1, pActiveCell.Column + 1)). _
                            Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    
    
    If theColumnTitle <> "" Then
        pActiveSheet.Cells(pActiveCell.Row, pActiveCell.Column + columnCounter + 1) = theColumnTitle
        
        rowCounter = 1
        For Each key In theValues.Keys()
        
            pActiveSheet.Cells(pActiveCell.Row + rowCounter, pActiveCell.Column + columnCounter) = key
            pActiveSheet.Cells(pActiveCell.Row + rowCounter, pActiveCell.Column + columnCounter + 1) = theValues(key)
         
            rowCounter = rowCounter + 1
         
        Next key
        
    Else
        
        rowCounter = 0
        For Each key In theValues.Keys()
        
            pActiveSheet.Cells(pActiveCell.Row + rowCounter, pActiveCell.Column + columnCounter) = key
            pActiveSheet.Cells(pActiveCell.Row + rowCounter, pActiveCell.Column + columnCounter + 1) = theValues(key)
         
            rowCounter = rowCounter + 1
         
        Next key
        
    End If
    
    
    ' modify cell block to be fit with the new row
    Call InitiateProperties
    
End Sub


Public Sub NewColumn_Collection(theValues As Collection, Optional theTitles As Collection, Optional theColumnTitle As Variant = "", Optional isInsert As Boolean = False)
    ''' writes new column
    ''' if theTitles is sent, it will write titles, too
    
    
    Dim key As Variant, i As Integer, cell As Range, rowCount As Double, columnCounter As Double
    
    
    rowCount = 0
    columnCounter = 0
    
    
    'find how many columns will be pasted
    If theColumnTitle <> "" Then
        rowCount = theValues.Count + 1
    Else
        rowCount = theValues.Count
    End If
    
    
    ' if titles collection is sent, write both titles and values
    If Not theTitles Is Nothing Then
    
        ' make two inserts
        If isInsert Then
            pActiveSheet.Range(Cells(pActiveCell.Row, pActiveCell.Column + 1), _
                        Cells(pActiveCell.Row + rowCount - 1, pActiveCell.Column + 1)). _
                        Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            pActiveSheet.Range(Cells(pActiveCell.Row, pActiveCell.Column + 1), _
                        Cells(pActiveCell.Row + rowCount - 1, pActiveCell.Column + 1)). _
                        Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
        
        ' if column title is sent
        If theColumnTitle <> "" Then
            pActiveSheet.Cells(pActiveCell.Row, pActiveCell.Column + columnCounter + 1) = theColumnTitle
            
            ' paste titles and values
            For i = 1 To theTitles.Count
                pActiveSheet.Cells(pActiveCell.Row + i, pActiveCell.Column + columnCounter) = theTitles(i)
                pActiveSheet.Cells(pActiveCell.Row + i, pActiveCell.Column + columnCounter + 1) = theValues(i)
            Next i
            
        ' if column title is not sent
        Else
            
            ' paste titles and values
            For i = 1 To theTitles.Count
                pActiveSheet.Cells(pActiveCell.Row + i - 1, pActiveCell.Column + columnCounter) = theTitles(i)
                pActiveSheet.Cells(pActiveCell.Row + i - 1, pActiveCell.Column + columnCounter + 1) = theValues(i)
            Next i
            
        End If

    
    ' if title is not sent,
    Else
    
        ' make one insert
        If isInsert Then
            pActiveSheet.Range(Cells(pActiveCell.Row, pActiveCell.Column + 1), _
                        Cells(pActiveCell.Row + rowCount - 1, pActiveCell.Column + 1)). _
                        Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        End If

        
        ' check if column title is sent
        If theColumnTitle <> "" Then
        
            pActiveSheet.Cells(pActiveCell.Row, pActiveCell.Column + columnCounter) = theColumnTitle
            
            For i = 1 To theValues.Count
                pActiveSheet.Cells(pActiveCell.Row + i, pActiveCell.Column + columnCounter) = theValues(i)
            Next i
            
        Else
        
            For i = 1 To theValues.Count
                pActiveSheet.Cells(pActiveCell.Row + i - 1, pActiveCell.Column + columnCounter) = theValues(i)
            Next i
        
        End If
        
    End If
    

    ' modify cell block to be fit with the new row
    Call InitiateProperties

End Sub



' =========================================
'   NEW ROW
' -----------------------------------------
Public Sub NewRow_Recordset(theValues As ADODB.Recordset, Optional isInsert As Boolean = False)
    ''' adds new row to current cell block
    ''' looks at the titles of the current cell block and inserts new rows according to them.
    ''' if the cell block is empty, inserts nothing
    
    Dim i As Double, rowCounter As Double, columnCount As Double, columnCounter As Double
    
    ' for the key values of theValues which doesn't exist in pActiveRowCells, put an error handling
    On Error Resume Next
    If theValues.EOF Then theValues.MoveFirst
    
    columnCount = 0
    rowCounter = 0
    columnCounter = 0
    
    'find how many columns will be pasted
    columnCount = theValues.fields.Count

    ' ===== write titles =====
    If isInsert Then
        pActiveSheet.Range(Cells(pActiveCell.Row + rowCounter + 1, pActiveCell.Column), _
                        Cells(pActiveCell.Row + rowCounter + 1, pActiveCell.Column + columnCount - 1)). _
                        Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If

    For i = 0 To theValues.fields.Count - 1
        pActiveSheet.Cells(pActiveCell.Row + rowCounter, pActiveCell.Column + columnCounter) = theValues.fields(i).Name
        columnCounter = columnCounter + 1
    Next i
    
    rowCounter = rowcountern + 1


    ' ===== write values =====
    Do Until theValues.EOF
    
        ' insert for each row
        If isInsert Then
            pActiveSheet.Range(Cells(pActiveCell.Row + rowCounter + 1, pActiveCell.Column), _
                            Cells(pActiveCell.Row + rowCounter + 1, pActiveCell.Column + columnCount - 1)). _
                            Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
    
        columnCounter = 0
    
        For i = 0 To theValues.fields.Count - 1
            pActiveSheet.Cells(pActiveCell.Row + rowCounter, pActiveCell.Column + columnCounter) = theValues.fields(i).Value
            columnCounter = columnCounter + 1
        Next i
        
        rowCounter = rowCounter + 1
        theValues.MoveNext
        
    Loop
    
    ' modify cell block to be fit with the new row
    Call InitiateProperties

End Sub


Public Sub NewRow_Dictionary(theValues As Dictionary, Optional theRowTitle As Variant = "", Optional isInsert As Boolean = False)
    ''' creates new cell block
    
    Dim key As Variant, rowCounter As Double, columnCount As Double, columnCounter As Double
    
    columnCount = 0
    rowCounter = 0
    
    'find how many columns will be pasted
    If theRowTitle <> "" Then
        columnCount = theValues.Count + 1
    Else
        columnCount = theValues.Count
    End If
    
    
    If isInsert Then
        pActiveSheet.Range(Cells(pActiveCell.Row + 1, pActiveCell.Column), _
                            Cells(pActiveCell.Row + 1, pActiveCell.Column + columnCount - 1)). _
                            Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        pActiveSheet.Range(Cells(pActiveCell.Row + 1, pActiveCell.Column), _
                            Cells(pActiveCell.Row + 1, pActiveCell.Column + columnCount - 1)). _
                            Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    
    
    If theRowTitle <> "" Then
        pActiveSheet.Cells(pActiveCell.Row + rowCounter + 1, pActiveCell.Column) = theRowTitle
        
        columnCounter = 1
        For Each key In theValues.Keys()
        
            pActiveSheet.Cells(pActiveCell.Row + rowCounter, pActiveCell.Column + columnCounter) = key
            pActiveSheet.Cells(pActiveCell.Row + rowCounter + 1, pActiveCell.Column + columnCounter) = theValues(key)
         
            columnCounter = columnCounter + 1
         
        Next key
        
    Else
        
        columnCounter = 0
        For Each key In theValues.Keys()
        
            pActiveSheet.Cells(pActiveCell.Row + rowCounter, pActiveCell.Column + columnCounter) = key
            pActiveSheet.Cells(pActiveCell.Row + rowCounter + 1, pActiveCell.Column + columnCounter) = theValues(key)
         
            columnCounter = columnCounter + 1
         
        Next key
        
    End If
    
    
    
    ' modify cell block to be fit with the new row
    Call InitiateProperties

End Sub


Public Sub NewRow_Collection(theValues As Collection, Optional theTitles As Collection, Optional theRowTitle As Variant = "", Optional isInsert As Boolean = False)
    ''' writes new row
    ''' if theTitles is sent, it will write titles, too
    
    Dim key As Variant, i As Integer, cell As Range, rowCounter As Double, columnCount As Double
    
    rowCounter = 0
    columnCount = 0
    
    
    'find how many columns will be pasted
    If theRowTitle <> "" Then
        columnCount = theValues.Count + 1
    Else
        columnCount = theValues.Count
    End If
    
    ' if titles collection is sent, write both titles and values
    If Not theTitles Is Nothing Then
    
        ' make two inserts
        If isInsert Then
            pActiveSheet.Range(Cells(pActiveCell.Row + 1, pActiveCell.Column), _
                            Cells(pActiveCell.Row + 1, pActiveCell.Column + columnCount - 1)). _
                            Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            pActiveSheet.Range(Cells(pActiveCell.Row + 1, pActiveCell.Column), _
                            Cells(pActiveCell.Row + 1, pActiveCell.Column + columnCount - 1)). _
                            Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
        
        ' if row title is sent
        If theRowTitle <> "" Then
            pActiveSheet.Cells(pActiveCell.Row + rowCounter + 1, pActiveCell.Column) = theRowTitle
            
            ' paste titles and values
            For i = 1 To theTitles.Count
                pActiveSheet.Cells(pActiveCell.Row + rowCounter, pActiveCell.Column + i) = theTitles(i)
                pActiveSheet.Cells(pActiveCell.Row + rowCounter + 1, pActiveCell.Column + i) = theValues(i)
            Next i
            
        ' if row title is not sent
        Else
            
            ' paste titles and values
            For i = 1 To theTitles.Count
                pActiveSheet.Cells(pActiveCell.Row + rowCounter, pActiveCell.Column + i - 1) = theTitles(i)
                pActiveSheet.Cells(pActiveCell.Row + rowCounter + 1, pActiveCell.Column + i - 1) = theValues(i)
            Next i
            
        End If

    
    ' if title is not sent,
    Else
    
        ' make two inserts
        If isInsert Then
            pActiveSheet.Range(Cells(pActiveCell.Row + 1, pActiveCell.Column), _
                            Cells(pActiveCell.Row + 1, pActiveCell.Column + columnCount - 1)). _
                            Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If

        
        ' check if row title is sent
        If theRowTitle <> "" Then
        
            pActiveSheet.Cells(pActiveCell.Row + rowCounter, pActiveCell.Column) = theRowTitle
            
            For i = 1 To theValues.Count
                pActiveSheet.Cells(pActiveCell.Row + rowCounter, pActiveCell.Column + i) = theValues(i)
            Next i
            
        Else
        
            For i = 1 To theValues.Count
                pActiveSheet.Cells(pActiveCell.Row + rowCounter, pActiveCell.Column + i - 1) = theValues(i)
            Next i
        
        End If
        
    End If
    
    
    ' modify cell block to be fit with the new row
    Call InitiateProperties
    

End Sub




' =========================================
'   ADD ROW
' -----------------------------------------
' PASSED
Public Sub AddRow_Recordset(theValues As ADODB.Recordset, Optional isInsert As Boolean = False)
    ''' adds new row to current cell block
    ''' looks at the titles of the current cell block and inserts new rows according to them.
    ''' if the cell block is empty, inserts nothing
    
    Dim i As Double, rowCounter As Double
    
    ' for the key values of theValues which doesn't exist in pActiveRowCells, put an error handling
    On Error Resume Next
    If theValues.EOF Then theValues.MoveFirst
    
    rowCounter = 1
    
    Do Until theValues.EOF
    
        ' insert for each row
        If isInsert Then
            pActiveSheet.Range(Cells(pBottomLeftCell.Row + rowCounter, pBottomLeftCell.Column), _
                            Cells(pBottomRightCell.Row + rowCounter, pBottomRightCell.Column)). _
                            Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
    
        For i = 0 To theValues.fields.Count - 1
            pActiveSheet.Cells(pLastCell.Row + rowCounter, pActiveRowCells(theValues.fields(i).Name).Column) = theValues.fields(i).Value
        Next i
        
        rowCounter = rowCounter + 1
        theValues.MoveNext
        
    Loop
    
    ' modify cell block to be fit with the new row
    Call InitiateProperties

End Sub

' PASSED
Public Sub AddRow_Dictionary(theValues As Dictionary, Optional theRowTitle As Variant = "", Optional isInsert As Boolean = False)
    ''' adds new row to current cell block
    
    Dim key As Variant, rowCounter As Double
    
    If isInsert Then
        pActiveSheet.Range(Cells(pBottomLeftCell.Row + 1, pBottomLeftCell.Column), _
                        Cells(pBottomRightCell.Row + 1, pBottomRightCell.Column)). _
                        Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    
    rowCounter = 1
    
    If theRowTitle <> "" Then
        pActiveSheet.Cells(pBottomLeftCell.Row + 1, pTitleColumn.Column) = theRowTitle
    End If
    
    ' for the key values of theValues which doesn't exist in pActiveRowCells, put an error handling
    On Error Resume Next
    For Each key In theValues.Keys()
    
        pActiveSheet.Cells(pLastCell.Row + rowCounter, pActiveRowCells(key).Column) = theValues(key)
    
    Next key
    
    ' modify cell block to be fit with the new row
    Call InitiateProperties

End Sub


Public Sub AddRow_Collection(theValues As Collection, Optional theTitles As Collection, Optional theRowTitle As Variant = "", Optional isInsert As Boolean = False)
    ''' adds new row to current cell block
    ''' if theTitles is sent, it will find the column of corresponding title and add values accordingly
    ''' otherwise, it will begin adding adding row title and values from the first colum. if there is no row title,
    ''' it will begin adding values only from the first column
    
    Dim key As Variant, i As Integer, cell As Range, rowCounter As Double
    
    rowCounter = 1
    
    If isInsert Then
        pActiveSheet.Range(Cells(pBottomLeftCell.Row + 1, pBottomLeftCell.Column), _
                        Cells(pBottomRightCell.Row + 1, pBottomRightCell.Column)). _
                        Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    

    If theRowTitle <> "" Then
        pActiveSheet.Cells(pBottomLeftCell.Row + rowCounter, pTitleColumn.Column) = theRowTitle
    End If
    
    ' if title collection is sent, find columns and add new row
    If Not theTitles Is Nothing Then
    
        ' paste row according to titles in the existing cell block
        For i = 1 To theTitles.Count
    
            For Each cell In pTitleRowCells
                If CStr(cell.Value) = theTitles(i) Then
                    pActiveSheet.Cells(pLastCell.Row + rowCounter, cell.Column) = theValues(i)
                End If
            Next cell
        
        Next i
    
    ' if title is not sent,
    Else
    
        ' check if row title is sent
        If theRowTitle <> "" Then
        
            For i = 1 To theValues.Count
                pActiveSheet.Cells(pLastCell.Row + rowCounter, pFirstCell.Column + i) = theValues(i)
            Next i
            
        Else
        
            For i = 1 To theValues.Count
                pActiveSheet.Cells(pLastCell.Row + rowCounter, pFirstCell.Column - 1 + i) = theValues(i)
            Next i
        
        End If
    End If
    
    ' modify cell block to be fit with the new row
    Call InitiateProperties
    

End Sub


' =========================================
'   ADD COLUMN
' -----------------------------------------
Public Sub AddColumn_Recordset(theValues As ADODB.Recordset, Optional isInsert As Boolean = False)
    ''' adds new row to current cell block
    Dim i As Double, rowCounter As Double
    
    ' if isInsert, insert number of columns
    If isInsert Then
        For i = 0 To theValues.fields.Count - 1
            pActiveSheet.Range(Cells(pTopRightCell.Row, pTopRightCell.Column + 1), _
                        Cells(pBottomRightCell.Row, pBottomRightCell.Column + 1)). _
                        Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Next i
    End If
    
    ' enter titles
    For i = 0 To theValues.fields.Count - 1
        pActiveSheet.Cells(pTitleRow.Row, pRightColumn.Column + i + 1) = theValues.fields(i).Name
    Next i
    
    ' enter values
    On Error Resume Next
    If theValues.EOF Then theValues.MoveFirst
    
    rowCounter = 1
    Do Until theValues.EOF
    
        ' if record count is more than height of cell block, insert for each record seperately
        If pTitleRow.Row + rowCounter > pBottomRow.Row Then
            If isInsert Then
    
                pActiveSheet.Range(Cells(pTitleRow.Row + rowCounter, pTopRightCell.Column + 1), _
                    Cells(pTitleRow.Row + rowCounter, pBottomRightCell.Column + theValues.fields.Count)). _
                    Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                   
            End If
        End If
    
        For i = 0 To theValues.fields.Count - 1
            
            pActiveSheet.Cells(pTitleRow.Row + rowCounter, pRightColumn.Column + i + 1) = theValues.fields(i).Value
            
        Next i
        
        rowCounter = rowCounter + 1
        theValues.MoveNext
        
    Loop
    
    ' modify cell block to be fit with the new row
    Call InitiateProperties
    
End Sub


Public Sub AddColumn_Dictionary(theValues As Dictionary, Optional theColumnTitle As Variant = "", Optional isInsert As Boolean = False)
    ''' adds new column to current cell block
    Dim key As Variant
    
    If isInsert Then
        pActiveSheet.Range(Cells(pTopRightCell.Row, pTopRightCell.Column + 1), _
                        Cells(pBottomRightCell.Row, pBottomRightCell.Column + 1)). _
                        Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    
    If theColumnTitle <> "" Then
        pActiveSheet.Cells(pTitleRow.Row, pTopRightCell.Column + 1) = theColumnTitle
    End If
    
    ' for the key values of theValues which doesn't exist in pActiveColumnCells, put an error handling
    On Error Resume Next
    For Each key In theValues.Keys()
    
        pActiveSheet.Cells(pActiveColumnCells(key).Row, pLastCell.Column + 1) = theValues(key)
    
    Next key
    
    ' modify cell block to be fit with the new row
    Call InitiateProperties
    
End Sub


Public Sub AddColumn_Collection(theValues As Collection, Optional theTitles As Collection, Optional theColumnTitle As Variant = "", Optional isInsert As Boolean = False)
    ''' adds new Column to current cell block
    ''' if theTitles is sent, it will find the row of corresponding title and add values accordingly
    ''' otherwise, it will begin adding adding row title and values from the first colum. if there is no row title,
    ''' it will begin adding values only from the first column
    
    
    Dim key As Variant, i As Integer, cell As Range
    
    If isInsert Then
        pActiveSheet.Range(Cells(pTopRightCell.Row, pTopRightCell.Column + 1), _
                        Cells(pBottomRightCell.Row, pBottomRightCell.Column + 1)). _
                        Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    End If

    If theColumnTitle <> "" Then
        pActiveSheet.Cells(pTitleRow.Row, pLastCell.Column + 1) = theColumnTitle
    End If
    
    ' if title collection is sent, find columns and add new Column
    If Not theTitles Is Nothing Then
        For i = 1 To theTitles.Count
        
            For Each cell In pTitleColumnCells
                If CStr(cell.Value) = theTitles(i) Then
                    pActiveSheet.Cells(cell.Row, pLastCell.Column + 1) = theValues(i)
                End If
            Next cell
        
        Next i
        
    ' if title is not sent,
    Else
    
        ' check if title is sent
        If theColumnTitle <> "" Then
        
            For i = 1 To theValues.Count
                pActiveSheet.Cells(pTitleRow.Row + i, pLastCell.Column + 1) = theValues(i)
            Next i
            
        Else
        
            For i = 1 To theValues.Count
                pActiveSheet.Cells(pTitleRow.Row - 1 + i, pLastCell.Column + 1) = theValues(i)
            Next i
        
        End If
    End If
    
    ' modify cell block to be fit with the new row
    Call InitiateProperties

End Sub


' =========================================
'   PUBLIC FUNCTIONS
' -----------------------------------------
Public Function isEdgeCell(theRange As Range) As Boolean
    ''' check if the cell is located in the edge of a cell block
    
    ' if the cell is at the edge of the spreadsheet, return true
    If theRange.Row - 1 = 0 Or theRange.Row + 1 = 1048577 Then
        isEdgeCell = True
        Exit Function
    End If
    
    If theRange.Column - 1 = 0 Or theRange.Column + 1 = 16385 Then
        isEdgeCell = True
        Exit Function
    End If
    
    ' check if surrounding cells are empty if any one of them is empty, then this is an edge cell
    If IsEmpty(theRange.Offset(-1, 0)) Then
        isEdgeCell = True
        Exit Function
    End If
    
    If IsEmpty(theRange.Offset(-1, 1)) Then
        isEdgeCell = True
        Exit Function
    End If
    
    If IsEmpty(theRange.Offset(0, 1)) Then
        isEdgeCell = True
        Exit Function
    End If
    
    If IsEmpty(theRange.Offset(1, 1)) Then
        isEdgeCell = True
        Exit Function
    End If
    
    If IsEmpty(theRange.Offset(1, 0)) Then
        isEdgeCell = True
        Exit Function
    End If
    
    If IsEmpty(theRange.Offset(1, -1)) Then
        isEdgeCell = True
        Exit Function
    End If
    
    If IsEmpty(theRange.Offset(0, -1)) Then
        isEdgeCell = True
        Exit Function
    End If
    
    If IsEmpty(theRange.Offset(-1, -1)) Then
        isEdgeCell = True
        Exit Function
    End If
    
    isEdgeCell = False

End Function


' =========================================
'   SET ATTRIBUTES
' -----------------------------------------
Private Sub SetActiveRowCells()
    On Error Resume Next
    
    Dim cell As Range

    Set pActiveRowCells = New Dictionary
    
    For Each cell In pTitleRowCells
    
        pActiveRowCells.Add CStr(cell.Value), pActiveSheet.Cells(pActiveCell.Row, cell.Column)
        
    Next cell

    Set cell = Nothing
End Sub

Private Sub SetActiveColumnCells()
    On Error Resume Next
    
    Dim cell As Range

    Set pActiveColumnCells = New Dictionary
    
    For Each cell In pTitleColumnCells
    
        pActiveColumnCells.Add CStr(cell.Value), pActiveSheet.Cells(cell.Row, pActiveCell.Column)
        
    Next cell

    Set cell = Nothing
End Sub


Private Sub SetTitleRowCells()
    
    Set pTitleRowCells = pActiveSheet.Range(Cells(pTitleRow.Row, pLeftColumn.Column), Cells(pTitleRow.Row, pRightColumn.Column))

End Sub


Private Sub SetTitleColumnCells()
    
    Set pTitleColumnCells = pActiveSheet.Range(Cells(pTitleRow.Row + 1, pTitleColumn.Column), Cells(pBottomRow.Row, pTitleColumn.Column))

End Sub


Private Sub SetTitleRow()
    ''' selects first row which has non-empty cells. Number of non-empty cells should be more than total column number - 2
    
    Dim i As Double
    
    For i = pTopRow.Row To pBottomRow.Row Step 1
        If WorksheetFunction.CountA(pActiveSheet.Range(Cells(i, pFirstCell.Column), Cells(i, pLastCell.Column))) > pLastCell.Column - pFirstCell.Column - 1 Then
            Set pTitleRow = pActiveSheet.Rows(i)
            Exit Sub
        End If
    Next i
    
    'if not found until 1, set it the first row
    Set pTitleRow = pActiveSheet.Rows(pFirstCell.Row)

End Sub


Private Sub SetTitleColumn()
    ''' selects first column which has non-empty cells. Number of non-empty cells should be equal to total number of rows - 2

    Dim i As Double
    
    'when setting title column, count rows after title row
    For i = pLeftColumn.Column To pRightColumn.Column Step 1
        If WorksheetFunction.CountA(pActiveSheet.Range(Cells(pTitleRow.Row + 1, i), Cells(pLastCell.Row, i))) > pLastCell.Row - pTitleRow.Row - 1 Then
            Set pTitleColumn = pActiveSheet.columns(i)
            Exit Sub
        End If
    Next i
    
    'if not found until 1, set it the first column
    Set pTitleColumn = pActiveSheet.columns(pFirstCell.Column)

End Sub


Private Sub SetTopRow()
    Dim i As Double
    
    On Error GoTo Sortie
    
    For i = pActiveCell.Row - 1 To 1 Step -1
        If WorksheetFunction.CountA(pActiveSheet.Rows(i)) = 0 Then
            Set pTopRow = pActiveSheet.Rows(i + 1)
            Exit Sub
        End If
    Next i
    
Sortie:
    'if not found until 1, set it the first row
    Set pTopRow = pActiveSheet.Rows(1)
End Sub


Private Sub SetBottomRow()
    Dim i As Double
    
    On Error GoTo Sortie
    
    For i = pActiveCell.Row + 1 To 1048576 Step 1
        If WorksheetFunction.CountA(pActiveSheet.Rows(i)) = 0 Then
            Set pBottomRow = pActiveSheet.Rows(i - 1)
            Exit Sub
        End If
    Next i
    
Sortie:
    'if not found until 1048576, set it the last row
    Set pBottomRow = pActiveSheet.Rows(1048576)
End Sub


Private Sub SetLeftColumn()
    Dim i As Double
    
    On Error GoTo Sortie
    
    For i = pActiveCell.Column - 1 To 1 Step -1
        If WorksheetFunction.CountA(pActiveSheet.Range(Cells(pTopRow.Row, i), Cells(pBottomRow.Row, i))) = 0 Then
            Set pLeftColumn = pActiveSheet.columns(i + 1)
            Exit Sub
        End If
    Next i
    
Sortie:
    'if not found until 1, set it the first column
    Set pLeftColumn = pActiveSheet.columns(1)
End Sub

Private Sub SetRightColumn()
    Dim i As Double
    
    On Error GoTo Sortie
    
    For i = pActiveCell.Column + 1 To 16384 Step 1
        If WorksheetFunction.CountA(pActiveSheet.Range(Cells(pTopRow.Row, i), Cells(pBottomRow.Row, i))) = 0 Then
            Set pRightColumn = pActiveSheet.columns(i - 1)
            Exit Sub
        End If
    Next i
    
Sortie:
    'if not found until 1, set it the first column
    Set pRightColumn = pActiveSheet.columns(16384)
End Sub



' =========================================
'   PUBLIC FUNCTIONS
' -----------------------------------------

Private Function DefineByConnectedCells(Optional theSelection As Range)
    ' iterates through all connected cells to theSelection.
    ' function ends when nextCell - currentCell pair is iterated for the second time
    
    Dim a As Variant

    Dim i As Integer, columnDifference As Integer, rowDifference As Integer
    Dim topRowNumber As Double, bottomRowNumber As Double, leftColumnNumber As Double, rightColumnNumber As Double
    Dim direction As String
    Dim currentCell As Range, nextCell As Range
    Dim cellChangePattern(0 To 3) As Integer
    
    Dim counter As Double
    
    If theSelection Is Nothing Then
        Set currentCell = pActiveCell
    Else
        Set currentCell = theSelection
    End If
    
    If currentCell.Value = "" Then Exit Function
    
    ' first get to the top edge cell
    For i = currentCell.Row To 1 Step -1
        If pActiveSheet.Cells(i, currentCell.Column) = "" Then
            Set currentCell = pActiveSheet.Cells(i + 1, currentCell.Column)
            direction = "up"
            GoTo TopCellFound
        End If
    Next i
    ' if edge is not found until row 1, then row 1 is the edge
    Set currentCell = pActiveSheet.Cells(1, currentCell.Column)
    
TopCellFound:

    ' first get to the left edge cell
    For i = currentCell.Column To 1 Step -1
        If pActiveSheet.Cells(currentCell.Column, i) = "" Then
            Set currentCell = pActiveSheet.Cells(currentCell.Row, i + 1)
            direction = "left"
            GoTo LeftCellFound
        End If
    Next i
    ' if edge is not found until column 1, then column 1 is the edge
    Set currentCell = pActiveSheet.Cells(currentCell.Row, 1)

LeftCellFound:
    ' set first direction as right because we want FindNextCell function to begin looking neighbour cells from up
    direction = "right"
    
    ' find nextcell and define change pattern
    Set nextCell = FindNextCell(currentCell, direction, False)
    
    cellChangePattern(0) = nextCell.Row
    cellChangePattern(1) = nextCell.Column
    cellChangePattern(2) = currentCell.Row
    cellChangePattern(3) = currentCell.Column
    
    ' initiate border rows and columns
    topRowNumber = currentCell.Row
    bottomRowNumber = currentCell.Row
    leftColumnNumber = currentCell.Column
    rightColumnNumber = currentCell.Column
    
    counter = 0
    
    Do
        
        columnDifference = nextCell.Column - currentCell.Column
        rowDifference = nextCell.Row - currentCell.Row
        
        direction = ""
        
        Select Case rowDifference
            Case 1
                direction = "down"
            Case -1
                direction = "up"
        End Select
        
        Select Case columnDifference
            Case 1
                If Len(direction) > 0 Then
                    direction = direction & "_" & "right"
                Else
                    direction = "right"
                End If
            Case -1
                If Len(direction) > 0 Then
                    direction = direction & "_" & "left"
                Else
                    direction = "left"
                End If
        End Select
    
        Set currentCell = nextCell
        
        ' find next cell
        Set nextCell = FindNextCell(currentCell, direction)
        
'        nextCell.Select
        
        'define farest border rows and columns
        topRowNumber = WorksheetFunction.Min(topRowNumber, nextCell.Row)
        bottomRowNumber = WorksheetFunction.Max(bottomRowNumber, nextCell.Row)
        leftColumnNumber = WorksheetFunction.Min(leftColumnNumber, nextCell.Column)
        rightColumnNumber = WorksheetFunction.Max(rightColumnNumber, nextCell.Column)
        
        If counter > 10000 Then Exit Function
        
        counter = counter + 1
        
    Loop Until (cellChangePattern(0) = nextCell.Row _
                And cellChangePattern(1) = nextCell.Column _
                And cellChangePattern(2) = currentCell.Row _
                And cellChangePattern(3) = currentCell.Column)
                
EndLoop:

    Set pTopRow = pActiveSheet.Rows(topRowNumber)
    Set pBottomRow = pActiveSheet.Rows(bottomRowNumber)
    Set pLeftColumn = pActiveSheet.columns(leftColumnNumber)
    Set pRightColumn = pActiveSheet.columns(rightColumnNumber)

End Function



Private Function FindNextCell(theSelection As Range, theDirection As String, Optional isEdge As Boolean = True) As Range
    ''' When defining cell block through connected cells, this function finds the next cell of the current cell


    Dim counter As Integer
    Dim nextCell As Range
    
    counter = 0

    If theDirection = "up" Then
        GoTo Left
    ElseIf theDirection = "up_right" Then
        GoTo UpLeft
    ElseIf theDirection = "right" Then
        GoTo Up
    ElseIf theDirection = "down_right" Then
        GoTo UpRight
    ElseIf theDirection = "down" Then
        GoTo Right
    ElseIf theDirection = "down_left" Then
        GoTo DownRight
    ElseIf theDirection = "left" Then
        GoTo Down
    ElseIf theDirection = "up_left" Then
        GoTo DownLeft
    Else
        GoTo ExitSub
    End If

Up:
    ' UP
    If theSelection.Row - 1 > 0 Then
        If pActiveSheet.Cells(theSelection.Row - 1, theSelection.Column) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row - 1, theSelection.Column)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub

UpRight:
    ' UP RIGHT
    If theSelection.Row - 1 > 0 And theSelection.Column + 1 <= 16384 Then
        If pActiveSheet.Cells(theSelection.Row - 1, theSelection.Column + 1) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row - 1, theSelection.Column + 1)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub

Right:
    ' RIGHT
    If theSelection.Column + 1 <= 16384 Then
        If pActiveSheet.Cells(theSelection.Row, theSelection.Column + 1) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row, theSelection.Column + 1)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub

DownRight:
    ' DOWN RIGHT
    If theSelection.Row + 1 < 1048576 And theSelection.Column + 1 <= 16384 Then
        If pActiveSheet.Cells(theSelection.Row + 1, theSelection.Column + 1) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row + 1, theSelection.Column + 1)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub

Down:
    ' DOWN
    If theSelection.Row + 1 < 1048576 Then
        If pActiveSheet.Cells(theSelection.Row + 1, theSelection.Column) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row + 1, theSelection.Column)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub

DownLeft:
    ' DOWN LEFT
    If theSelection.Row + 1 < 1048576 And theSelection.Column - 1 > 0 Then
        If pActiveSheet.Cells(theSelection.Row + 1, theSelection.Column - 1) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row + 1, theSelection.Column - 1)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub

Left:
    ' LEFT
    If theSelection.Column - 1 > 0 Then
        If pActiveSheet.Cells(theSelection.Row, theSelection.Column - 1) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row, theSelection.Column - 1)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub

UpLeft:
    ' UP LEFT
    If theSelection.Row - 1 > 0 And theSelection.Column - 1 > 0 Then
        If pActiveSheet.Cells(theSelection.Row - 1, theSelection.Column - 1) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row - 1, theSelection.Column - 1)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub


    ' UP
    If theSelection.Row - 1 > 0 Then
        If pActiveSheet.Cells(theSelection.Row - 1, theSelection.Column) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row - 1, theSelection.Column)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub


    ' UP RIGHT
    If theSelection.Row - 1 > 0 And theSelection.Column + 1 <= 16384 Then
        If pActiveSheet.Cells(theSelection.Row - 1, theSelection.Column + 1) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row - 1, theSelection.Column + 1)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub


    ' RIGHT
    If theSelection.Column + 1 <= 16384 Then
        If pActiveSheet.Cells(theSelection.Row, theSelection.Column + 1) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row, theSelection.Column + 1)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub


    ' DOWN RIGHT
    If theSelection.Row + 1 < 1048576 And theSelection.Column + 1 <= 16384 Then
        If pActiveSheet.Cells(theSelection.Row + 1, theSelection.Column + 1) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row + 1, theSelection.Column + 1)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub


    ' DOWN
    If theSelection.Row + 1 < 1048576 Then
        If pActiveSheet.Cells(theSelection.Row + 1, theSelection.Column) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row + 1, theSelection.Column)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub


    ' DOWN LEFT
    If theSelection.Row + 1 < 1048576 And theSelection.Column - 1 > 0 Then
        If pActiveSheet.Cells(theSelection.Row + 1, theSelection.Column - 1) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row + 1, theSelection.Column - 1)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub


    ' LEFT
    If theSelection.Column - 1 > 0 Then
        If pActiveSheet.Cells(theSelection.Row, theSelection.Column - 1) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row, theSelection.Column - 1)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub


    ' UP LEFT
    If theSelection.Row - 1 > 0 And theSelection.Column - 1 > 0 Then
        If pActiveSheet.Cells(theSelection.Row - 1, theSelection.Column - 1) <> "" Then

        Set nextCell = pActiveSheet.Cells(theSelection.Row - 1, theSelection.Column - 1)
        GoTo ExitSub
        End If
    End If

    counter = counter + 1
    If counter = 8 And isEdge Then GoTo ExitSub


ExitSub:
    If nextCell Is Nothing Then Set nextCell = theSelection
    Set FindNextCell = nextCell
    Set nextCell = Nothing


End Function

