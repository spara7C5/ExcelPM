Attribute VB_Name = "Main"
Public MainSheets As New Scripting.Dictionary
Public prjlist As New Scripting.Dictionary
Public Usrlist As New Scripting.Dictionary
Dim DeleteSheets As New Collection
Public Const CPTEMP As String = "Control Panel"
Public Const FRTEMP As String = "FRTemplate"
Public Const VARTEMP As String = "Variables"
Const PRJSheet As String = "P2048v62"


Public FRTsht, CPsht, VRsht As Worksheet


Public newFR As New FR
Public newFR2 As New FR
Public newFR3 As New FR

Public Const TASKOWN As Integer = 1

Const ObjectType_n As Integer = 2
Enum ObjectType
 FR_OBJ = 0
 PRJ_OBJ = 1
 USR_OBJ = ObjectType_n
End Enum


'''''''''''''''''''''''''''''''''''''
'Initialize
'Called from "ThisWorkbook module"
'''''''''''''''''''''''''''''''''''''
Sub Initialize()
    
    ControlPanel_Form.Height = 280
    ControlPanel_Form.Width = 700
    Set FRTsht = ActiveWorkbook.Sheets(FRTEMP)
    Set CPsht = ActiveWorkbook.Sheets(CPTEMP)
    Set VRsht = ActiveWorkbook.Sheets(VARTEMP)
    Main.VRsht.Visible = xlSheetHidden
    On Error GoTo ex
    MainSheets.Add Item:=FRTsht, Key:=FRTsht.name
    MainSheets.Add Item:=CPsht, Key:=CPsht.name
    MainSheets.Add Item:=VRsht, Key:=VRsht.name
    
    
    
    StoreAndLoad.LoadObjects
    
ex:
End Sub
Sub Button3_Click()

    ControlPanel_Form.Show

End Sub
Sub ClearProject()

    Initialize
    Dim wsht As Worksheet
    Dim res As Long
    
    For Each wsht In ActiveWorkbook.Worksheets
        If Not MainSheets.Exists(wsht.name) Then
            wsht.Delete
        End If
        
        
    Next
    
    Main.prjlist.RemoveAll
End Sub
Sub CreateProject()

    
    Dim tabrange As Range
    Dim taskCol, lastRow, lastCol As Integer

    

End Sub


Sub UndoProject()
    Set FRTsht = ActiveWorkbook.Sheets(FRTEMP)
    FRTsht.ListObjects("Test1").Unlist
End Sub

Public Function CheckSpecialCharacters(nTextValue As String) As Boolean
CheckSpecialCharacters = nTextValue Like "*[!A-Za-z0-9]*"
End Function

Public Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function

Sub test()

    Dim tmprng As Range
    Dim tmpwsht As Worksheet
    
    Set tmpwsht = ThisWorkbook.Worksheets("FRTemplate")
    Set tmprng = tmpwsht.Range("$A$1:$E$9")
    
    tmprng.Copy
    
    ThisWorkbook.Worksheets("8001").Cells(10, 10).PasteSpecial Paste:=xlPasteColumnWidths
    ThisWorkbook.Worksheets("8001").Cells(10, 10).PasteSpecial Paste:=xlPasteAll
    
   'tmprng.Select
   
    
End Sub

Sub test2()

    test3 Cells(8, 1)

End Sub

Sub test3()

    Dim myusr As New USR
    
    myusr.Create "pippo"
    
End Sub

Sub testGet()

Dim auxprj As PRJ


Set auxprj = Main.prjlist.Item("8000v10")
auxprj.EditFR "8000v10last", "Task", "Modelling", "Owner", "12345"

End Sub




