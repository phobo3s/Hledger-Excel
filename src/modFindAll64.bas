Attribute VB_Name = "modFindAll64"
Option Explicit
Option Compare Text
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modFindAll
' By Chip Peasron, chip@cpearson.com. www.cpearson.com
' Web page for this module: www.cpearson.com/Excel/FindAll.aspx
' 24-October-2007
' Revised 5-January-2010
' This module is described at www.cpearson.com/Excel/FindAll.aspx
' Requires Excel 2000 or later.
'
' This module contains two functions, FindAll and FindAllOnWorksheets that are use
' to find values on a worksheet or multiple worksheets.
'
' FindAll searches a range and returns a range containing the cells in which the
'   searched for text was found. If the string was not found, it returns Nothing.

' FindAllOnWorksheets searches the same range on one or more workshets. It return
'   an array of ranges, each of which is the range on that worksheet in which the
'   value was found. If the value was not found on a worksheet, that worksheet's
'   element in the returned array will be Nothing.
'
' In both functions, the parameters that control the search have the same meaning
' and effect as they do in the Range.Find method.
' This module is compatible with 64-bit Excel.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function FindAll(SearchRange As Range, _
                FindWhat As Variant, _
               Optional LookIn As XlFindLookIn = xlValues, _
                Optional LookAt As XlLookAt = xlWhole, _
                Optional SearchOrder As XlSearchOrder = xlByRows, _
                Optional MatchCase As Boolean = False, _
                Optional BeginsWith As String = vbNullString, _
                Optional EndsWith As String = vbNullString, _
                Optional BeginEndCompare As VbCompareMethod = vbTextCompare) As Range
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FindAll
' This searches the range specified by SearchRange and returns a Range object
' that contains all the cells in which FindWhat was found. The search parameters to
' this function have the same meaning and effect as they do with the
' Range.Find method. If the value was not found, the function return Nothing. If
' BeginsWith is not an empty string, only those cells that begin with BeginWith
' are included in the result. If EndsWith is not an empty string, only those cells
' that end with EndsWith are included in the result. Note that if a cell contains
' a single word that matches either BeginsWith or EndsWith, it is included in the
' result.  If BeginsWith or EndsWith is not an empty string, the LookAt parameter
' is automatically changed to xlPart. The tests for BeginsWith and EndsWith may be
' case-sensitive by setting BeginEndCompare to vbBinaryCompare. For case-insensitive
' comparisons, set BeginEndCompare to vbTextCompare. If this parameter is omitted,
' it defaults to vbTextCompare. The comparisons for BeginsWith and EndsWith are
' in an OR relationship. That is, if both BeginsWith and EndsWith are provided,
' a match if found if the text begins with BeginsWith OR the text ends with EndsWith.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim foundCell As Range
Dim FirstFound As Range
Dim LastCell As Range
Dim resultRange As Range
Dim XLookAt As XlLookAt
Dim Include As Boolean
Dim CompMode As VbCompareMethod
Dim Area As Range
Dim MaxRow As Long
Dim MaxCol As Long
Dim BeginB As Boolean
Dim EndB As Boolean


CompMode = BeginEndCompare
If BeginsWith <> vbNullString Or EndsWith <> vbNullString Then
    XLookAt = xlPart
Else
    XLookAt = LookAt
End If

' this loop in Areas is to find the last cell
' of all the areas. That is, the cell whose row
' and column are greater than or equal to any cell
' in any Area.
For Each Area In SearchRange.Areas
    With Area
        If .Cells(.Cells.count).Row > MaxRow Then
            MaxRow = .Cells(.Cells.count).Row
        End If
        If .Cells(.Cells.count).Column > MaxCol Then
            MaxCol = .Cells(.Cells.count).Column
        End If
    End With
Next Area
Set LastCell = SearchRange.Worksheet.Cells(MaxRow, MaxCol)


'On Error Resume Next
On Error GoTo 0
Set foundCell = SearchRange.find(what:=FindWhat, _
        After:=LastCell, _
        LookIn:=LookIn, _
        LookAt:=XLookAt, _
        SearchOrder:=SearchOrder, _
        MatchCase:=MatchCase)

If Not foundCell Is Nothing Then
    Set FirstFound = foundCell
    'Set ResultRange = FoundCell
    'Set FoundCell = SearchRange.FindNext(after:=FoundCell)
    Do Until False ' Loop forever. We'll "Exit Do" when necessary.
        Include = False
        If BeginsWith = vbNullString And EndsWith = vbNullString Then
            Include = True
        Else
            If BeginsWith <> vbNullString Then
                If StrComp(left(foundCell.Text, Len(BeginsWith)), BeginsWith, BeginEndCompare) = 0 Then
                    Include = True
                End If
            End If
            If EndsWith <> vbNullString Then
                If StrComp(right(foundCell.Text, Len(EndsWith)), EndsWith, BeginEndCompare) = 0 Then
                    Include = True
                End If
            End If
        End If
        If Include = True Then
            If resultRange Is Nothing Then
                Set resultRange = foundCell
            Else
                Set resultRange = Application.Union(resultRange, foundCell)
            End If
        End If
        Set foundCell = SearchRange.FindNext(After:=foundCell)
        If (foundCell Is Nothing) Then
            Exit Do
        End If
        If (foundCell.Address = FirstFound.Address) Then
            Exit Do
        End If

    Loop
End If
    
Set FindAll = resultRange

End Function

Function FindAllOnWorksheets(InWorkbook As Workbook, _
                InWorksheets As Variant, _
                SearchAddress As String, _
                FindWhat As Variant, _
                Optional LookIn As XlFindLookIn = xlValues, _
                Optional LookAt As XlLookAt = xlWhole, _
                Optional SearchOrder As XlSearchOrder = xlByRows, _
                Optional MatchCase As Boolean = False, _
                Optional BeginsWith As String = vbNullString, _
                Optional EndsWith As String = vbNullString, _
                Optional BeginEndCompare As VbCompareMethod = vbTextCompare) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FindAllOnWorksheets
' This function searches a range on one or more worksheets, in the range specified by
' SearchAddress.
'
' InWorkbook specifies the workbook in which to search. If this is Nothing, the active
'   workbook is used.
'
' InWorksheets specifies what worksheets to search. InWorksheets can be any of the
' following:
'   - Empty: This will search all worksheets of the workbook.
'   - String: The name of the worksheet to search.
'   - String: The names of the worksheets to search, separated by a ':' character.
'   - Array: A one dimensional array whose elements are any of the following:
'           - Object: A worksheet object to search. This must be in the same workbook
'               as InWorkbook.
'           - String: The name of the worksheet to search.
'           - Number: The index number of the worksheet to search.
' If any one of the specificed worksheets is not found in InWorkbook, no search is
' performed. The search takes place only after everything has been validated.
'
' The other parameters have the same meaning and effect on the search as they do
' in the Range.Find method.
'
' Most of the code in this procedure deals with the InWorksheets parameter to give
' the absolute maximum flexibility in specifying which sheet to search.
'
' This function requires the FindAll procedure, also in this module or avaialable
' at www.cpearson.com/Excel/FindAll.aspx.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim WSArray() As String
Dim ws As Worksheet
Dim wb As Workbook
Dim resultRange() As Range
Dim WSNdx As Long
Dim r As Range
Dim SearchRange As Range
Dim FoundRange As Range
Dim WSS As Variant
Dim n As Long


'''''''''''''''''''''''''''''''''''''''''''
' Determine what Workbook to search.
'''''''''''''''''''''''''''''''''''''''''''
If InWorkbook Is Nothing Then
    Set wb = ActiveWorkbook
Else
    Set wb = InWorkbook
End If

'''''''''''''''''''''''''''''''''''''''''''
' Determine what sheets to search
'''''''''''''''''''''''''''''''''''''''''''
If IsEmpty(InWorksheets) = True Then
    ''''''''''''''''''''''''''''''''''''''''''
    ' Empty. Search all sheets.
    ''''''''''''''''''''''''''''''''''''''''''
    With wb.Worksheets
        ReDim WSArray(1 To .count)
        For WSNdx = 1 To .count
            WSArray(WSNdx) = .item(WSNdx).name
        Next WSNdx
    End With

Else
    '''''''''''''''''''''''''''''''''''''''
    ' If Object, ensure it is a Worksheet
    ' object.
    ''''''''''''''''''''''''''''''''''''''
    If IsObject(InWorksheets) = True Then
        If TypeOf InWorksheets Is Excel.Worksheet Then
            ''''''''''''''''''''''''''''''''''''''''''
            ' Ensure Worksheet is in the WB workbook.
            ''''''''''''''''''''''''''''''''''''''''''
            If StrComp(InWorksheets.parent.name, wb.name, vbTextCompare) <> 0 Then
                ''''''''''''''''''''''''''''''
                ' Sheet is not in WB. Get out.
                ''''''''''''''''''''''''''''''
                Exit Function
            Else
                ''''''''''''''''''''''''''''''
                ' Same workbook. Set the array
                ' to the worksheet name.
                ''''''''''''''''''''''''''''''
                ReDim WSArray(1 To 1)
                WSArray(1) = InWorksheets.name
            End If
        Else
            '''''''''''''''''''''''''''''''''''''
            ' Object is not a Worksheet. Get out.
            '''''''''''''''''''''''''''''''''''''
        End If
    Else
        '''''''''''''''''''''''''''''''''''''''''''
        ' Not empty, not an object. Test for array.
        '''''''''''''''''''''''''''''''''''''''''''
        If IsArray(InWorksheets) = True Then
            '''''''''''''''''''''''''''''''''''''''
            ' It is an array. Test if each element
            ' is an object. If it is a worksheet
            ' object, get its name. Any other object
            ' type, get out. Not an object, assume
            ' it is the name.
            ''''''''''''''''''''''''''''''''''''''''
            ReDim WSArray(LBound(InWorksheets) To UBound(InWorksheets))
            For WSNdx = LBound(InWorksheets) To UBound(InWorksheets)
                If IsObject(InWorksheets(WSNdx)) = True Then
                    If TypeOf InWorksheets(WSNdx) Is Excel.Worksheet Then
                        ''''''''''''''''''''''''''''''''''''''
                        ' It is a worksheet object, get name.
                        ''''''''''''''''''''''''''''''''''''''
                        WSArray(WSNdx) = InWorksheets(WSNdx).name
                    Else
                        ''''''''''''''''''''''''''''''''
                        ' Other type of object, get out.
                        ''''''''''''''''''''''''''''''''
                        Exit Function
                    End If
                Else
                    '''''''''''''''''''''''''''''''''''''''''''
                    ' Not an object. If it is an integer or
                    ' long, assume it is the worksheet index
                    ' in workbook WB.
                    '''''''''''''''''''''''''''''''''''''''''''
                    Select Case UCase(TypeName(InWorksheets(WSNdx)))
                        Case "LONG", "INTEGER"
                            Err.Clear
                            '''''''''''''''''''''''''''''''''''
                            ' Ensure integer if valid index.
                            '''''''''''''''''''''''''''''''''''
                            Set ws = wb.Worksheets(InWorksheets(WSNdx))
                            If Err.Number <> 0 Then
                                '''''''''''''''''''''''''''''''
                                ' Invalid index.
                                '''''''''''''''''''''''''''''''
                                Exit Function
                            End If
                            ''''''''''''''''''''''''''''''''''''
                            ' Valid index. Get name.
                            ''''''''''''''''''''''''''''''''''''
                            WSArray(WSNdx) = wb.Worksheets(InWorksheets(WSNdx)).name
                        Case "STRING"
                            Err.Clear
                            '''''''''''''''''''''''''''''''''''''
                            ' Ensure valid name.
                            '''''''''''''''''''''''''''''''''''''
                            Set ws = wb.Worksheets(InWorksheets(WSNdx))
                            If Err.Number <> 0 Then
                                '''''''''''''''''''''''''''''''''
                                ' Invalid name, get out.
                                '''''''''''''''''''''''''''''''''
                                Exit Function
                            End If
                            WSArray(WSNdx) = InWorksheets(WSNdx)
                    End Select
                End If
                'WSArray(WSNdx) = InWorksheets(WSNdx)
            Next WSNdx
        Else
            ''''''''''''''''''''''''''''''''''''''''''''
            ' InWorksheets is neither an object nor an
            ' array. It is either the name or index of
            ' the worksheet.
            ''''''''''''''''''''''''''''''''''''''''''''
            Select Case UCase(TypeName(InWorksheets))
                Case "INTEGER", "LONG"
                    '''''''''''''''''''''''''''''''''''''''
                    ' It is a number. Ensure sheet exists.
                    '''''''''''''''''''''''''''''''''''''''
                    Err.Clear
                    Set ws = wb.Worksheets(InWorksheets)
                    If Err.Number <> 0 Then
                        '''''''''''''''''''''''''''''''
                        ' Invalid index, get out.
                        '''''''''''''''''''''''''''''''
                        Exit Function
                    Else
                        WSArray = Array(wb.Worksheets(InWorksheets).name)
                    End If
                Case "STRING"
                    '''''''''''''''''''''''''''''''''''''''''''''''''''
                    ' See if the string contains a ':' character. If
                    ' so, the InWorksheets contains a string of multiple
                    ' worksheets.
                    '''''''''''''''''''''''''''''''''''''''''''''''''''
                    If InStr(1, InWorksheets, ":", vbBinaryCompare) > 0 Then
                        ''''''''''''''''''''''''''''''''''''''''''
                        ' ":" character found. split apart sheet
                        ' names.
                        ''''''''''''''''''''''''''''''''''''''''''
                        WSS = Split(InWorksheets, ":")
                        Err.Clear
                        n = LBound(WSS)
                        If Err.Number <> 0 Then
                            '''''''''''''''''''''''''''''
                            ' Unallocated array. Get out.
                            '''''''''''''''''''''''''''''
                            Exit Function
                        End If
                        If LBound(WSS) > UBound(WSS) Then
                            '''''''''''''''''''''''''''''
                            ' Unallocated array. Get out.
                            '''''''''''''''''''''''''''''
                            Exit Function
                        End If
                            
                                                
                        ReDim WSArray(LBound(WSS) To UBound(WSS))
                        For n = LBound(WSS) To UBound(WSS)
                            Err.Clear
                            Set ws = wb.Worksheets(WSS(n))
                            If Err.Number <> 0 Then
                                Exit Function
                            End If
                            WSArray(n) = WSS(n)
                         Next n
                    Else
                        Err.Clear
                        Set ws = wb.Worksheets(InWorksheets)
                        If Err.Number <> 0 Then
                            '''''''''''''''''''''''''''''''''
                            ' Invalid name, get out.
                            '''''''''''''''''''''''''''''''''
                            Exit Function
                        Else
                            ReDim WSArray(0)
                            WSArray(0) = InWorksheets
                        End If
                    End If
            End Select
        End If
    End If
End If
'''''''''''''''''''''''''''''''''''''''''''
' Ensure SearchAddress is valid
'''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
For WSNdx = LBound(WSArray) To UBound(WSArray)
    Err.Clear
    Set ws = wb.Worksheets(WSArray(WSNdx))
    ''''''''''''''''''''''''''''''''''''''''
    ' Worksheet does not exist
    ''''''''''''''''''''''''''''''''''''''''
    If Err.Number <> 0 Then
        Exit Function
    End If
    Err.Clear
    Set r = wb.Worksheets(WSArray(WSNdx)).Range(SearchAddress)
    If Err.Number <> 0 Then
        ''''''''''''''''''''''''''''''''''''
        ' Invalid Range. Get out.
        ''''''''''''''''''''''''''''''''''''
        Exit Function
    End If
Next WSNdx

''''''''''''''''''''''''''''''''''''''''
' SearchAddress is valid for all sheets.
' Call FindAll to search the range on
' each sheet.
''''''''''''''''''''''''''''''''''''''''
ReDim resultRange(LBound(WSArray) To UBound(WSArray))
For WSNdx = LBound(WSArray) To UBound(WSArray)
    Set ws = wb.Worksheets(WSArray(WSNdx))
    Set SearchRange = ws.Range(SearchAddress)
    Set FoundRange = FindAll(SearchRange:=SearchRange, _
                    FindWhat:=FindWhat, _
                    LookIn:=LookIn, LookAt:=LookAt, _
                    SearchOrder:=SearchOrder, _
                    MatchCase:=MatchCase, _
                    BeginsWith:=BeginsWith, _
                    EndsWith:=EndsWith)
    
    If FoundRange Is Nothing Then
        Set resultRange(WSNdx) = Nothing
    Else
        Set resultRange(WSNdx) = FoundRange
    End If
Next WSNdx

FindAllOnWorksheets = resultRange

End Function













