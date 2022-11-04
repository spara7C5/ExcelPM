VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CLASS MODULE CODE = PRJ (Project = sum of FRs)
' AUTHOR = Stefano Paracchino
' DATE = 12th October 2022
' DESCRIPTION =


'''''''''''''''''''''''''''''''''''''''''''''''''''
' Member variables declaration section
'''''''''''''''''''''''''''''''''''''''''''''''''''
Private m_descriptor()
Private m_name As String
Private m_FRlist As New Scripting.Dictionary
Private test As Integer
Private m_frnum As Integer
Private m_ws As Worksheet 'sheet name and prj name are the same
Private m_frOrigPointer As Range


Public m_frrange As Range

'''''''''''''''''''''''''''''''''''''''''''''''''''
' End Member variables declaration section
'''''''''''''''''''''''''''''''''''''''''''''''''''

Property Get frnum() As Integer
    frnum = m_frnum
End Property

Property Let frnum(value As Integer)
    m_frnum = value
End Property

Public Property Get frrange() As Range
    frrange = m_frrange
End Property

Public Property Set frrange(ByVal value As Range)
    Set m_frrange = value
End Property



'''''''''''''''''''''''''''''''''''''''''''''''''''
' Subroutine (class methods with no return) section
'''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Create(ByVal name As String)
    
    m_name = name
    
    ActiveWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count)).name = m_name
    Set m_ws = ActiveWorkbook.Worksheets(m_name)
    
    Set m_frOrigPointer = m_ws.Cells(1, 1)
        

   
End Sub

Public Sub Load(ByVal name As String, ByVal org As Range, frrng As Range)
    
    m_name = name
    
    Set m_frrange = frrng
    
    Set m_ws = ActiveWorkbook.Worksheets(m_name)
    
    Set m_frOrigPointer = org
        
   
End Sub
Private Sub Class_Initialize()
   
End Sub

Public Sub UpdateAndStore()

    ReDim m_descriptor(1 To 4, 1 To 1)
    
    m_descriptor(1, 1) = PRJ_OBJ
    m_descriptor(2, 1) = m_name
    m_descriptor(3, 1) = m_frOrigPointer.Address
    m_descriptor(4, 1) = m_frrange.Address
    
    StoreArray m_descriptor, m_name
    
End Sub



Public Sub AddFR(title As String)
    
    Dim auxFR As New FR
    Dim fullname As String
    Dim auxrng As Range
    
    fullname = m_name + title 'in order to create a unique FR name
    Set auxrng = SearchStored(fullname)
    If auxrng Is Nothing Then

        auxFR.Create m_frrange, fullname, m_frOrigPointer, m_name
        m_FRlist.Add Item:=auxFR, Key:=fullname
        Set m_frOrigPointer = m_ws.Cells(auxFR.m_tab.TotalsRowRange.Rows.Row + 1, 1)
        UpdateAndStore
        
    Else
        MsgBox "FR already exists"
    
    End If
    

    
End Sub

Function RemoveFR(title As String) As Boolean
    
    Dim auxFR As New FR
    Dim fullname As String
    Dim auxrng As Range
    Dim auxsht As Worksheet
    
    
    fullname = m_name + title 'in order to create a unique FR name
    Set auxrng = SearchStored(fullname)
    
    If auxrng Is Nothing Then ' FR does not exist
        
        RemoveFR = False
        
    Else 'FR to be removed
        
        Set auxFR = m_FRlist.Item(fullname)
        auxFR.Clear
        m_FRlist.Remove fullname
        RemoveFR = DeleteStored(fullname)
    End If
    

    
End Function

Public Sub LoadFR(frx As FR)

    m_FRlist.Add Item:=frx, Key:=frx.name


End Sub

Public Sub EditFR(frn As String, COLSearch As String, COLValue As String, COLTarget As String, NEWVal As Integer)
    
    Dim frx As FR
    
    Set frx = m_FRlist.Item(frn)
    
    frx.EditTable COLSearch, COLValue, COLTarget, NEWVal

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

