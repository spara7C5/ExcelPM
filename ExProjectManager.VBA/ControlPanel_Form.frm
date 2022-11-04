VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ControlPanel_Form 
   Caption         =   "Control Panel"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   408
   ClientWidth     =   13788
   OleObjectBlob   =   "ControlPanel_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ControlPanel_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const comm_num As Integer = 6

Const comm1 As String = "Add Project"
Const comm2 As String = "Remove Project"
Const comm3 As String = "Add FR"
Const comm4 As String = "Remove FR"
Const comm5 As String = "Modify all FRs"
Const comm6 As String = "Assign FR"

Private comm_dic As New Scripting.Dictionary

Dim frrange As Range


Private Sub Butt_debug_Click()

    If Main.VRsht.Visible = xlSheetHidden Then
        Main.VRsht.Visible = xlSheetVisible
    Else
        Main.VRsht.Visible = xlSheetHidden
    End If
End Sub

Private Sub Butt_FRok_Click()
    
     If TextB_frname.Text = "" Then
       MsgBox ("FR Name is empty!")
       Exit Sub
    End If
    
    Application.Run ("Butt_FRok" + "_callback")
    
End Sub

Private Sub Butt_reset_Click()
    Call Main.ClearProject
    Call StoreAndLoad.ClearAllStored
End Sub

Private Sub Butt_Run_Click()


        
Dim x As String
    

    If Not comm_dic.Exists(Combo_command.Text) Then
       MsgBox ("Command does not exist!")
       Exit Sub
    End If
    
    If TextB_name.Text = "" Or Main.CheckSpecialCharacters(TextB_name.Text) Then
       MsgBox ("Poject Name is empty or contains special characters!")
       Exit Sub
    End If
    
      
    Application.Run (comm_dic(Combo_command.Text) + "_callback")
End Sub




Private Sub Combo_command_Change()
    Application.Run (comm_dic(Combo_command.Text) + "_change")
End Sub

Private Sub Frame_FR_Click()

End Sub

Private Sub Lab_frname_Click()

End Sub

Private Sub UserForm_Initialize()

    Combo_command.List = Array(comm1, comm2, comm3, comm4, comm5, comm6)
    
    comm_dic.Add Item:="comm1", Key:=comm1
    comm_dic.Add Item:="comm2", Key:=comm2
    comm_dic.Add Item:="comm3", Key:=comm3
    comm_dic.Add Item:="comm4", Key:=comm4
    comm_dic.Add Item:="comm5", Key:=comm5
    comm_dic.Add Item:="comm6", Key:=comm6
    
    ControlPanel_cb.GreyFRdown
    
End Sub



