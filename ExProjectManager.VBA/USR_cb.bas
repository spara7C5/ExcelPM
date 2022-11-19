Attribute VB_Name = "USR_cb"
Public Sub Addday_cb()

    Dim tmpUSR As New USR
    Dim tmpstr As String
    Dim stridx, strl As Integer
    
    tmpstr = Application.Caller
    strl = Len(tmpstr)
    stridx = InStr(tmpstr, "_")
    tmpstr = Right(tmpstr, strl - stridx)
    
    Set tmpUSR = Main.Usrlist.Item(tmpstr)
    
    tmpUSR.addday


End Sub

Public Sub Deliv_cb()

    Dim tmpUSR As New USR
    Dim tmpstr As String
    Dim stridx, strl As Integer
    
    tmpstr = Application.Caller
    strl = Len(tmpstr)
    stridx = InStr(tmpstr, "_")
    tmpstr = Right(tmpstr, strl - stridx)
    
    Set tmpUSR = Main.Usrlist.Item(tmpstr)

    tmpUSR.deliv
    
End Sub
