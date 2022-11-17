VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    Call Main.Initialize
    
End Sub

Sub Workbook_SheetActivate(ByVal Sh As Object)

    Dim tmpUSR As USR
    For Each shtnm In Main.Usrlist.Keys
        If shtnm = Sh.name Then
            Set tmpUSR = Main.Usrlist.Item(shtnm)
            tmpUSR.updatevalues
        End If
        
    Next
       
    
        
End Sub


