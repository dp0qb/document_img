VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "LAM-ICPMS Common Lead Correction - version 3.15"
   ClientHeight    =   9708.001
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox3_Click()
If CheckBox3.Value = True Then
UserForm1.TextBox2.Enabled = False
Else: UserForm1.TextBox2.Enabled = True
End If
End Sub

Private Sub CheckBox4_Click()
If UserForm1.CheckBox4.Value = True Then
UserForm1.CheckBox5.Enabled = False
UserForm1.CheckBox6.Enabled = False
Else
UserForm1.CheckBox5.Enabled = True
UserForm1.CheckBox6.Enabled = True
End If
End Sub

Private Sub CheckBox5_Click()
If UserForm1.CheckBox5.Value = True And UserForm1.CheckBox6.Value = True Then
UserForm1.CheckBox4.Enabled = False
ElseIf UserForm1.CheckBox5.Value = True And UserForm1.CheckBox6.Value = False Then
UserForm1.CheckBox6.Enabled = False
UserForm1.CheckBox4.Enabled = False
Else
UserForm1.CheckBox6.Enabled = True
UserForm1.CheckBox4.Enabled = True
End If
End Sub

Private Sub CheckBox6_Click()
If UserForm1.CheckBox6.Value = True Then

UserForm1.CheckBox4.Enabled = False
Else

UserForm1.CheckBox4.Enabled = True
End If

End Sub

Private Sub CommandButton1_Click()
UserForm1.Hide
gemoc_on = False
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub OptionButton1_Click()
UserForm1.TextBox10.Visible = False
End Sub

Private Sub OptionButton2_Click()
UserForm1.TextBox10.Visible = True
End Sub

Private Sub OptionButton3_Click()
UserForm1.TextBox10.Visible = False
End Sub

Private Sub CommandButton2_Click()

gemoc_on = False
End

End Sub

Private Sub TextBox2_Change()

If UserForm1.TextBox2.Text = "." Or UserForm1.TextBox2.Value = "" Or UserForm1.TextBox2.Value = " " Then

Else
   

   t = UserForm1.TextBox2.Value
 
    UserForm1.TextBox7.Text = 18.7 - 9.74 * (Exp(l8 * t) - 1)
    UserForm1.TextBox8.Text = 15.628 - 9.74 * (Exp(l5 * t) - 1) / 137.88
    UserForm1.TextBox9.Text = 38.63 - 36.84 * (Exp(l2 * t) - 1)

End If

End Sub

Private Sub UserForm_Click()

End Sub
