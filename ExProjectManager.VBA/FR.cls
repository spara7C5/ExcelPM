VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CLASS MODULE CODE = FR (functional request)
' AUTHOR = Stefano Paracchino
' DATE = 12th October 2022
' DESCRIPTION =


'''''''''''''''''''''''''''''''''''''''''''''''''''
' Member variables declaration section
'''''''''''''''''''''''''''''''''''''''''''''''''''
Private m_descriptor()
Private m_origin As Range
Private m_title As Range
Private m_ws As Worksheet
Public m_prj As String
Private test As Integer

Public m_tab As ListObject
'''''''''''''''''''''''''''''''''''''''''''''''''''
' End Member variables declaration section
'''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''
'Property section (functions as variables)
'''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get origin() As Range

    Set origin = m_origin
    
End Property

Property Get name() As String
    name = m_title.Cells(0, 1).value
End Property
'''''''''''''''''''''''''''''''''''''''''''''''''''
' End of Property section
'''''''''''''''''''''''''''''''''''''''''''''''''''





'''''''''''''''''''''''''''''''''''''''''''''''''''
' Subroutine (class methods with no return) section
'''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Create(ByVal rng As Range, title As String, orig As Range, ByVal PRJ As String)
    
    Dim tabOrig, tabrange As Range
    Dim wstDest As Worksheet
    Set m_title = orig.Resize(, rng.Columns.Count)
    m_title.Interior.Color = rgbBisque
    m_title.Merge
    m_title.Cells(1, 1) = title
    m_title.HorizontalAlignment = xlHAlignCenterAcrossSelection
    m_title.Font.Bold = True
    Set wstDest = orig.Worksheet
    With wstDest
        Set tabOrig = .Range(.Cells(orig.Row + 1, orig.Column), .Cells(orig.Row + 1, orig.Column))
    End With
    
    rng.Copy
    tabOrig.PasteSpecial Paste:=xlPasteColumnWidths
    tabOrig.PasteSpecial Paste:=xlPasteAll
    
    Set tabrange = tabOrig.Resize(rng.Rows.Count, rng.Columns.Count)
    wstDest.ListObjects.Add(xlSrcRange, tabrange).name = title
    Set m_tab = wstDest.ListObjects(title)
    m_tab.ShowTotals = True
    
    Dim LC As ListColumn
    For Each LC In m_tab.ListColumns
        LC.TotalsCalculation = xlTotalsCalculationSum
    Next


    Set m_origin = orig
    m_prj = PRJ
    
    UpdateAndStore
    
End Sub

Public Sub Load(ByVal tabrng As Range, title As String, orig As Range, ByVal PRJ As String)
    
    Dim tabOrig, tabrange As Range
    Dim wstDest As Worksheet
    
    Set m_title = orig.Resize(, tabrng.Columns.Count)
    Set wstDest = orig.Worksheet
    
    'wstDest.ListObjects.Add(xlSrcRange, tabrng).name = title
    Set m_tab = wstDest.ListObjects(title)
    m_tab.ShowTotals = True
    
    Dim LC As ListColumn
    For Each LC In m_tab.ListColumns
        LC.TotalsCalculation = xlTotalsCalculationSum
    Next


    Set m_origin = orig
    m_prj = PRJ
    
    
End Sub

Private Sub Class_Initialize()

    Set m_origin = ActiveSheet.Cells(100, 100)
    
End Sub

Public Sub UpdateAndStore()

    ReDim m_descriptor(1 To 5, 1 To 1)
    m_descriptor(1, 1) = FR_OBJ
    m_descriptor(2, 1) = m_prj
    m_descriptor(3, 1) = m_title.Cells(1, 1).value
    m_descriptor(4, 1) = m_origin.Address
    m_descriptor(5, 1) = m_tab.Range.Address
    
    StoreArray m_descriptor, m_descriptor(3, 1)
    
End Sub

Public Sub move(new_origin As Range)

    m_origin.Cut
    Set m_origin = new_origin
    m_origin.PasteSpecial Paste:=xlPasteColumnWidths
    m_origin.PasteSpecial Paste:=xlPasteAll
    
End Sub

Sub EditTable(COLSearch As String, COLValue As String, COLTarget As String, NEWVal As Integer)
    Dim x, y As Integer
    '' controllare che match funzioni anche con gli spazi!
        y = Application.WorksheetFunction.Match(COLValue, m_tab.ListColumns(COLSearch).Range, 0)
        x = Application.WorksheetFunction.Match(COLTarget, m_tab.HeaderRowRange, 0)
        m_tab.DataBodyRange.Cells(y - 1, x) = NEWVal
'    With m_tab
'
'        .ListRows(.ListColumns(COLSearch).Range.Find(COLValue).Row).Range.Cells(.ListColumns(COLTarget).Index) = NEWVal
'
'    End With
    
End Sub


Sub Clear()
    
    m_title.Cells.Delete
    m_tab.DataBodyRange.Delete
    m_tab.Delete

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
