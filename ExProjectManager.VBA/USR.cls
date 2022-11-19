VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "USR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CLASS MODULE CODE = USR (user)
' AUTHOR = Stefano Paracchino
' DATE = 9th november 2022
' DESCRIPTION = Class for the management of the user (= developer/consultant/contributor)


'''''''''''''''''''''''''''''''''''''''''''''''''''
' Member variables declaration section
'''''''''''''''''''''''''''''''''''''''''''''''''''
Const COLWID As Integer = 20

Private m_descriptor()
Private m_name As String
Private m_ws As Worksheet 'sheet name and name are the same
Private m_tab_dayactivity As ListObject
Private m_tab_dayactivity_orig As Range
Public butadd As Button
Public butdeliv As Button

'''''''''''''''''''''''''''''''''''''''''''''''''''
' End Member variables declaration section
'''''''''''''''''''''''''''''''''''''''''''''''''''

Enum work_cols
    COL_DATE = 1
    COL_PRJ = 2
    COL_FR = 3
    COL_ACTIVITY = 4
    COL_HOURS = 5

End Enum



'''''''''''''''''''''''''''''''''''''''''''''''''''
' Subroutine (class methods with no return) section
'''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Create(ByVal name As String)
    
    Dim tabrange, taskrange, btnrange As Range
    Dim prjliststr As String
    
    m_name = name
    
    ActiveWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count)).name = m_name
    Set m_ws = ActiveWorkbook.Worksheets(m_name)
    
    Set btnrange = m_ws.Cells(2, 2)
    Set butadd = m_ws.Buttons.Add(btnrange.Left, btnrange.Top, btnrange.Width, btnrange.Height)

    With butadd
        .OnAction = "Addday_cb"
        .Caption = "Add Day"
        .name = "Addday_" & m_name 'in order to retrieve object name in button callback
    
    End With
    
    Set btnrange = m_ws.Cells(2, 3)
    Set butdeliv = m_ws.Buttons.Add(btnrange.Left, btnrange.Top, btnrange.Width, btnrange.Height)

    With butdeliv
        .OnAction = "Deliv_cb"
        .Caption = "Deliver"
        .name = "Deliv_" & m_name 'in order to retrieve object name in button callback
    
    End With
    
    Application.ScreenUpdating = True
    Set m_tab_dayactivity_orig = m_ws.Cells(4, 2)
    m_ws.ListObjects.Add(xlSrcRange, m_tab_dayactivity_orig).name = m_name
    Set m_tab_dayactivity = m_ws.ListObjects(m_name)
    m_tab_dayactivity.TableStyle = "TableStyleDark10"
    
    AddAndSetColToTab m_tab_dayactivity, "date", xlValidateDate, xlGreater, DateSerial(Year(Now), 1, 1)
    AddAndSetColToTab m_tab_dayactivity, "Project", xlValidateList, xlEqual, Join(Main.prjlist.Keys, ",")
    AddAndSetColToTab m_tab_dayactivity, "FR"
    AddAndSetColToTab m_tab_dayactivity, "Activity"
    AddAndSetColToTab m_tab_dayactivity, "Hours"

    
    'Main.Prjlist.Item(m_tab_dayactivity.ListColumns(m_tab_dayactivity.ListColumns.Count - 1).DataBodyRange.Cells(1, 1).value).frlist
    
    addday  'Add first day
    m_tab_dayactivity.ListColumns(6).Delete
    UpdateAndStore
      
End Sub

Private Sub AddAndSetColToTab(tabx As ListObject, ByVal namex As String, Optional ByVal typex As XlDVType, Optional ByVal operx As Variant, Optional ByVal formux As String)

    Dim colx As ListColumns
    Dim abscol As Integer
    
    Set colx = tabx.ListColumns
    colx.Add
    colx(colx.Count - 1).name = namex
    colx(colx.Count - 1).Range.ColumnWidth = COLWID
    If Not IsMissing(typex) And Not IsMissing(operx) And Not IsMissing(formux) Then
        colx(colx.Count - 1).Range.Validation.Add Type:=typex, AlertStyle:=xlValidAlertStop, Operator:=operx, Formula1:=formux
    End If

End Sub

Function getfrlistfromprj() As String

    Dim out As String
    out = ""
    
    On Error GoTo exi
    out = Main.prjlist.Item(m_tab_dayactivity.ListColumns(m_tab_dayactivity.ListColumns.Count - 1).DataBodyRange.Cells(1, 1).value).getAllFRstr
     
exi:
End Function
Public Sub Load(title As String)
    

    Dim wstDest As Worksheet
    
    m_name = title
    Set wstDest = ThisWorkbook.Worksheets(title)
   
    Set m_tab_dayactivity = wstDest.ListObjects(title)
    m_tab_dayactivity.ShowTotals = True
    
    'Set butadd = wsdest.Buttons("Addday_" & m_name)
    'Set butdeliv = wsdest.Buttons("Deliv_" & m_name)
    
    'Dim LC As ListColumn
    'For Each LC In m_tab.ListColumns
    '    LC.TotalsCalculation = xlTotalsCalculationSum
    'Next

    
    
End Sub
Private Sub Class_Initialize()
  
    
    
End Sub

Public Sub UpdateAndStore()

    ReDim m_descriptor(1 To 2, 1 To 1)
    
    m_descriptor(1, 1) = USR_OBJ
    m_descriptor(2, 1) = m_name
    'm_descriptor(3, 1) = m_tab_dayactivity.Range.Address
    
    StoreArray m_descriptor, m_name
    
End Sub

Public Sub updatevalues()

    With m_tab_dayactivity.ListColumns("Project").Range.Validation

    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlEqual, Formula1:=Join(Main.prjlist.Keys, ",")
    
    End With
  
    'With m_tab_dayactivity.ListColumns("Task").Range.Validation

    '.Delete
    '.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlEqual, Formula1:=Join(Main.prjlist.Keys, ",")
    
    'End With
    
End Sub

Public Sub addday()

    m_tab_dayactivity.ListRows.Add
    UpdateAndStore
    
End Sub

Public Sub deliv()

    Dim rw As ListRow
    Dim tmpval As Integer
    

    For Each rw In m_tab_dayactivity.ListRows
    
        If rw.Range.Cells(1, COL_DATE).Interior.ColorIndex <> IDX_GREEN Then
        
            On Error GoTo exi
            
                tmpval = Main.prjlist(rw.Range.Cells(1, COL_PRJ).value).GetFR( _
                rw.Range.Cells(1, COL_PRJ).value & rw.Range.Cells(1, COL_FR).value, _
                "Task", _
                rw.Range.Cells(1, COL_ACTIVITY).value, _
                "Actual Effort" _
                )
                
                tmpval = tmpval + rw.Range.Cells(1, COL_HOURS).value
            
                Main.prjlist(rw.Range.Cells(1, COL_PRJ).value).EditFR _
                rw.Range.Cells(1, COL_PRJ).value & rw.Range.Cells(1, COL_FR).value, _
                "Task", _
                rw.Range.Cells(1, COL_ACTIVITY).value, _
                "Actual Effort", _
                tmpval
                 
                rw.Range.Cells(1, COL_DATE).Interior.ColorIndex = IDX_GREEN

        End If
    
     Next rw
exi:
 
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''
' End of Subroutine section
'''''''''''''''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''
'Function section
'''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''
'End of function section
'''''''''''''''''''''''''''''''''''''''''''''''''''


