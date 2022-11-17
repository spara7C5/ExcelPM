Attribute VB_Name = "StoreAndLoad"
Option Explicit

Public Const PRENAME As String = "n_"
Public Const POSTNAME As String = "_st"
Public Const SHEETVARS As String = "Variables"

Public arrFirst()   'Need to be public variables at the top of a standard module
Public arrSecond()  'Need to be public variables at the top of a standard module

Sub CreateArray()

    Dim a As Long
    Dim b As Long
    Dim i As Long
    Dim j As Long

   
    a = 10
    b = 1
     
    ClearAllStored
   
    'Dimension and create first test array
    ReDim arrFirst(1 To a, 1 To b)   'Two dimensional one based array
   
    For i = 1 To a
        For j = 1 To b
            arrFirst(i, j) = "First " & Chr(i + 64) & "_" & Chr(j + 64)
        Next j
    Next i



    'Dimension and create second test array
    ReDim arrSecond(1 To a, 1 To b)   'Two dimensional one based array
   
    For i = 1 To a
        For j = 1 To b
            arrSecond(i, j) = "Second " & Chr(i + 64) & "_" & Chr(j + 64)
        Next j
    Next i

    ReDim arrThird(1 To a, 1 To b)   'Two dimensional one based array
   
    For i = 1 To a
        For j = 1 To b
            arrThird(i, j) = "Third " & Chr(i + 64) & "_" & Chr(j + 64)
        Next j
    Next i
    
    ReDim arrFourth(1 To a, 1 To b)   'Two dimensional one based array
   
    For i = 1 To a
        For j = 1 To b
            arrFourth(i, j) = "Fourth " & Chr(i + 64) & "_" & Chr(j + 64)
        Next j
    Next i
    'Following for testing if required
    'Debug.Print 'Adds line feed
    'For i = 1 To a
    '    For j = 1 To 3
    '        Debug.Print arrSecond(i, j) & ", ";
    '    Next j
    '    Debug.Print
    'Next i
   
    StoreArray arrFirst, "firstarray"
    StoreArray arrSecond, "secondarray"
    StoreArray arrThird, "thirdarray"
    StoreArray arrFourth, "fourtharray"
    
End Sub

Function LastRowOrCol(bolRowOrCol As Boolean, Optional rng As Range) As Long
    'Finds the last used row or column in a worksheet
    'First parameter is True for Last Row or False for last Column
    'Third parameter is optional
        'Must be specified if not ActiveSheet
   
    Dim lngRowCol As Long
    Dim rngToFind As Range
   
    If rng Is Nothing Then
        Set rng = ThisWorkbook.Worksheets(SHEETVARS).Cells
    End If
   
    If bolRowOrCol Then
        lngRowCol = xlByRows
    Else
        lngRowCol = xlByColumns
    End If
   
    With rng
        Set rngToFind = rng.Find(What:="*", _
                LookIn:=xlFormulas, _
                LookAt:=xlPart, _
                SearchOrder:=lngRowCol, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False)
    End With
   
    If Not rngToFind Is Nothing Then
        If bolRowOrCol Then
            LastRowOrCol = rngToFind.Row
        Else
            LastRowOrCol = rngToFind.Column
        End If
    End If
   
End Function

Sub StoreArray(ByVal arr As Variant, ByVal arrname As String)

    Dim wbthis As Workbook
    Dim wsArrays As Worksheet
    Dim fullname As String
    Dim lngNextRow As Long
    Dim tmprng As Range
    
        
    Set wbthis = ThisWorkbook
    Set wsArrays = wbthis.Worksheets(SHEETVARS)
    
    fullname = PRENAME & arrname & POSTNAME
    
    Set tmprng = SearchStored(arrname)
    
    If tmprng Is Nothing Then 'name not exists, create new one
    
         With wsArrays
    
           
            'First array saves at cell(1,1). (Resized range to fit the array)
            '.Cells(1, 1).Resize(UBound(arrFirst, 1), UBound(arrFirst, 2)).name = arrname
            'Range(arrname).Value = arr
       
            lngNextRow = LastRowOrCol(True) + 2 'Starting row for next array (spaced one row)
           
            
            'Next array saves 2 rows below previous array
            .Cells(lngNextRow, 1).Resize(UBound(arr, 1)).name = fullname
            Range(fullname).value = arr
            
        End With
        
    
    Else 'name already exists, just update
   
        Range(fullname).value = arr
        
    End If
    

End Sub



Sub ClearAllStored()

    Dim wbthis As Workbook
    Dim wsArrays As Worksheet
    Dim nme As name
    Dim WsStr As String
    Dim tmprng As Range
    
        
    Set wbthis = ThisWorkbook
    Set wsArrays = wbthis.Worksheets(SHEETVARS)
    
    With wsArrays
        .Cells.Clear    'Clears all data
        'Delete any existing names in the worksheet

        For Each nme In wbthis.Names
        
            On Error GoTo del
            WsStr = nme.RefersToRange.Parent.name 'assignemnt for error trigger in case of sheet cancelling
            
            If WsStr = wsArrays.name Then 'extra check the name belongs to the dedicated sheet
del:             nme.Delete
            
            End If
        Next nme
            
        
    End With
End Sub


Private Sub Workbook_Open()

    LoadObjects
   

End Sub

Public Sub LoadObjects()


    Dim nme As name
    Dim wbthis As Workbook
    Dim wsArrays As Worksheet
    Dim loadedrng As New Collection
    Dim loadedfr As New Collection
    Dim collfr As New Collection
    Dim tmprng As Variant
    Dim tmpFR As FR
    Dim tmpPRJ As PRJ
    Dim tmpUSR As USR
    
    Set wbthis = ThisWorkbook
    Set wsArrays = wbthis.Worksheets(SHEETVARS)
    
    For Each nme In wbthis.Names
        If nme.RefersToRange.Parent.name = wsArrays.name Then

            loadedrng.Add Range(nme.name)

        End If
    Next nme


    For Each tmprng In loadedrng
    
        

        Select Case tmprng(1, 1).value
        
        Case FR_OBJ
        
            Set tmpFR = New FR
            tmpFR.Load ThisWorkbook.Worksheets(CStr(tmprng(2, 1).value)).Range(CStr(tmprng(4, 1).value)), _
            tmprng(3, 1).value, _
            ThisWorkbook.Worksheets(CStr(tmprng(2, 1).value)).Range(CStr(tmprng(5, 1).value)), _
            tmprng(2, 1).value
            Debug.Print "Loaded obj fr: " & tmprng(3, 1).value
            collfr.Add tmpFR
            
        
        Case PRJ_OBJ
        
            Set tmpPRJ = New PRJ
            tmpPRJ.Load tmprng(2, 1).value, _
            ThisWorkbook.Worksheets(CStr(tmprng(2, 1).value)).Range(CStr(tmprng(3, 1).value)), _
            ThisWorkbook.Worksheets(Main.FRTEMP).Range(CStr(tmprng(4, 1).value))
            
            Debug.Print "Loaded obj prj: " & tmprng(2, 1).value
            Main.prjlist.Add Item:=tmpPRJ, Key:=CStr(tmprng(2, 1).value)
            
        Case USR_OBJ
       
            Set tmpUSR = New USR
            tmpUSR.Load tmprng(2, 1).value
            
            Debug.Print "Loaded obj usr: " & tmprng(2, 1).value
            Main.Usrlist.Add Item:=tmpUSR, Key:=CStr(tmprng(2, 1).value)
        
        End Select
    
    Next tmprng
    
    Dim frx As FR
    
    For Each frx In collfr
    
        Main.prjlist.Item(frx.m_prj).LoadFR frx
    
    Next frx
    
    
End Sub


Function SearchStored(ByVal str As String) As Range

    Dim nme As name
    Dim wbthis As Workbook
    Dim wsArrays As Worksheet
    Dim fullstr As String
    
    Set wbthis = ThisWorkbook
    Set wsArrays = wbthis.Worksheets(SHEETVARS)
    
    fullstr = PRENAME & str & POSTNAME
    
    Set SearchStored = Nothing
    
    For Each nme In wbthis.Names
        If nme.RefersToRange.Parent.name = wsArrays.name And nme.name = fullstr Then
                Set SearchStored = Range(nme.name)
                Exit For
            
        End If
        
    Next nme


End Function

Function DeleteStored(ByVal str As String) As Boolean

    Dim wbthis As Workbook
    Dim wsArrays As Worksheet
    Dim fullstr As String
    DeleteStored = True
    
    Set wbthis = ThisWorkbook
    Set wsArrays = wbthis.Worksheets(SHEETVARS)
    
    fullstr = PRENAME & str & POSTNAME
    
    
    Range(wbthis.Names(fullstr).name).Delete
    wbthis.Names(fullstr).Delete

End Function
