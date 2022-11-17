Attribute VB_Name = "ControlPanel_cb"
'''''''''''''''''''''''''''''''''
' shared variable among callbacks
'''''''''''''''''''''''''''''''''
Private prjfocus As String
Private frcounter As String
'''''''''''''''''''''''''''''''''
' end of shared vars
'''''''''''''''''''''''''''''''''


Public Sub comm1_callback() 'Add project

    Dim auxprj As New PRJ
    
    With ControlPanel_Form
        If .RefEd_FRrange.Text = "" Then
            MsgBox ("Select a FR template range")
          Exit Sub
        End If
        
        If IsNumeric(.TextB_FRnum.Text) = False Then
           MsgBox ("FR num must be a valid number")
           Exit Sub
        ElseIf CInt(.TextB_FRnum.Text) < 1 Then
            MsgBox ("FR num must be > 0")
            Exit Sub
        
        End If
        
              
        Set frrange = Range(.RefEd_FRrange.value)
        With frrange
            If .Rows(Main.TASKOWN).Find("Task") Is Nothing Then
                MsgBox "FR Range: missing Task column at row 1"
                Exit Sub
            End If
            
            'lastRow = .Cells(Rows.Count, .Rows(TASKOWN).Find("Task").Column).End(xlUp).Row
            'lastCol = .Cells(TASKOWN, Columns.Count).End(xlToLeft).Column
        
        End With

        
        'Set Prjsht = ActiveWorkbook.Worksheets(.TextB_name.Text)
        'Set tabrange = FRTsht.Range(FRTsht.Cells(1, 1), FRTsht.Cells(lastRow, lastCol))
        'newFR.Create FRrange, "fr4", Prjsht.Cells(1, 1)
        'newFR2.Create FRrange, "fr5", Prjsht.Cells(newFR.m_tab.TotalsRowRange.Rows.Row + 1, 1)
        'newFR3.Create FRrange, "fr6", Prjsht.Cells(newFR2.m_tab.TotalsRowRange.Rows.Row + 1, 1)
        
        If Main.prjlist.Exists(.TextB_name.Text) = True Then
            MsgBox "project already exists!"
            Exit Sub
        End If
        
       auxprj.Create CStr(.TextB_name.Text)
       auxprj.frnum = .TextB_FRnum.Text
       Set auxprj.frrange = Range(.RefEd_FRrange.value)
       frcounter = CStr(.TextB_FRnum.Text)
       prjfocus = CStr(.TextB_name.Text)
       
        Main.prjlist.Add Item:=auxprj, Key:=CStr(.TextB_name.Text)
    
        GreyFRup
        UngreyFRdown
        GreyCPleft
   
    End With
    
End Sub


Public Sub comm2_callback() 'Remove Project

    Dim prjname As String
    
    With ControlPanel_Form
    
        If Main.prjlist.Exists(.TextB_name.Text) = False Then
            MsgBox "project does not exist!Check the name"
            GoTo exi
        End If
    prjname = .TextB_name.Text
    Main.prjlist.Item(prjname).RemoveAllFR
    Main.prjlist.Remove prjname
    DeleteStored (prjname)
    ThisWorkbook.Worksheets(prjname).Delete
    
    End With
   
exi:

End Sub

Public Sub comm3_callback() 'Add FR
    Dim tmpPRJ As New PRJ
    
    
    With ControlPanel_Form
    
        If .TextB_name = "" Then
            MsgBox "project name cannot be empty!"
            GoTo exi
        End If
        
        If .TextB_frname = "" Or Main.CheckSpecialCharacters(.TextB_frname.Text) Then
            MsgBox "FR name cannot be empty or contain specail characters!"
            GoTo exi
        End If
        
        If Main.prjlist.Exists(.TextB_name.Text) = False Then
            MsgBox "project does not exist!Check the name"
            GoTo exi
        End If
       
        
        Set tmpPRJ = Main.prjlist.Item(CStr(.TextB_name.Text))
        
        tmpPRJ.AddFR (CStr(.TextB_frname))
        
    End With
exi:
    'UngreyFRup
    'GreyFRdown
End Sub

Public Sub comm4_callback() 'Remove FR

    Dim tmpPRJ As New PRJ
    
    
    With ControlPanel_Form
    
        If .TextB_name = "" Then
            MsgBox "project name cannot be empty!"
            GoTo exi
        End If
        
        If .TextB_frname = "" Or Main.CheckSpecialCharacters(.TextB_frname.Text) Then
            MsgBox "FR name cannot be empty or contain specail characters!"
            GoTo exi
        End If
        
        If Main.prjlist.Exists(.TextB_name.Text) = False Then
            MsgBox "project does not exist!Check the name"
            GoTo exi
        End If
        
        Set tmpPRJ = Main.prjlist.Item(CStr(.TextB_name.Text))
        
        
        res = tmpPRJ.RemoveFR(CStr(.TextB_name.Text) + CStr(.TextB_frname))
        
        If res = True Then
            MsgBox "FR " & CStr(.TextB_frname) & " removed"
        Else
            MsgBox "FR " & CStr(.TextB_frname) & " not present in selected project!"
        End If
    End With
    
exi:

End Sub

Public Sub comm5_callback() 'Add developer

    Dim tmpUSR As New USR
   
    With ControlPanel_Form
    
        If Main.prjlist.Exists(.TextB_name.Text) = True Then
            MsgBox "username already exists!"
            Exit Sub
        End If
        
        If Main.prjlist.Count < 1 Then
            MsgBox "Cannot create user with no projects!"
            Exit Sub
        End If
        
        
       tmpUSR.Create CStr(.TextB_name.Text)
       
        Main.Usrlist.Add Item:=tmpUSR, Key:=CStr(.TextB_name.Text)
    
   
    End With
    
End Sub
Public Sub comm6_callback() 'Remove usr

    Dim usrname As String
    
    With ControlPanel_Form
    
        If Main.Usrlist.Exists(.TextB_name.Text) = False Then
            MsgBox "userr does not exist!Check the name"
            GoTo exi
        End If
    usrname = .TextB_name.Text
    Main.Usrlist.Remove usrname
    'DeleteStored (usrname)
    ThisWorkbook.Worksheets(usrname).Delete
    
    End With
   
exi:

End Sub


Private Sub Butt_FRok_callback()

    Dim tempprj As New PRJ
    
    
    With ControlPanel_Form
    
        If .TextB_frname = "" Or Main.CheckSpecialCharacters(.TextB_frname) Then
        
            MsgBox ("FR name cannot be empty or contains special characters!")
            
            Exit Sub
        
        End If
        
       
       Set tempprj = Main.prjlist.Item(prjfocus)
       
       tempprj.AddFR .TextB_frname
     
    
        frcounter = frcounter - 1
    
        If frcounter = 0 Then
        
            UngreyCPleft
            UngreyFRup
            GreyFRdown
            
            Exit Sub
        
        End If
   
   
    End With
End Sub

Sub GreyFRup()

    With ControlPanel_Form
        .RefEd_FRrange.Enabled = False
        .TextB_FRnum.Enabled = False
        .TextB_FRnum.BackColor = &H80000016
        .Lab_FRnum.ForeColor = &H80000015
        .Lab_FRsel.ForeColor = &H80000015
    End With

End Sub

Sub UngreyFRup()

    With ControlPanel_Form
        .RefEd_FRrange.Enabled = True
        .TextB_FRnum.Enabled = True
        .TextB_FRnum.BackColor = vbWhite
        .Lab_FRnum.ForeColor = vbBlack
        .Lab_FRsel.ForeColor = vbBlack
    
    End With

End Sub

Sub GreyFRdown()

    With ControlPanel_Form
        .TextB_frname.Enabled = False
        .TextB_frname.BackColor = &H80000016
        .Lab_frname.ForeColor = &H80000015
        .Butt_FRok.Enabled = False
    End With

End Sub

Sub UngreyFRdown()

    With ControlPanel_Form
        .TextB_frname.Enabled = True
        .TextB_frname.BackColor = vbWhite
        .Lab_frname.ForeColor = vbBlack
        .Butt_FRok.Enabled = True
    End With

End Sub

Sub GreyCPleft()

    With ControlPanel_Form
        .Combo_command.Enabled = False
        .Lab_command.ForeColor = &H80000015
        .TextB_name.Enabled = False
        .TextB_name.BackColor = &H80000016
        .Butt_Run.Enabled = False
    End With

End Sub

Sub UngreyCPleft()

    With ControlPanel_Form
        .Combo_command.Enabled = True
        .Lab_command.ForeColor = vbBlack
        .TextB_name.Enabled = True
        .TextB_name.BackColor = vbWhite
        .Butt_Run.Enabled = True
    End With

End Sub

Public Sub comm1_change()

    UngreyFRup
    GreyFRdown

End Sub
Public Sub comm2_change()

    UngreyCPleft
    GreyFRup
    GreyFRdown
    
End Sub

Public Sub comm3_change()

    GreyFRup
    UngreyFRdown
    ControlPanel_Form.Butt_FRok.Enabled = False
    
End Sub
Public Sub comm4_change()

    comm3_change
    
End Sub
Public Sub comm5_change()

    GreyFRup
    GreyFRdown
End Sub

Public Sub comm6_change()

    GreyFRup
    GreyFRdown
End Sub
