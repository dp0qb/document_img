VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "ComPbCorr#3 - GEMOC version"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
UserForm3.Hide
End Sub

Private Sub CommandButton2_Click()

End
End Sub

Private Sub OptionButton1_Click()
If UserForm3.OptionButton1.Value = True Then
UserForm3.TextBox1.Value = 1
UserForm3.TextBox2.Value = 0
UserForm3.Label3.Visible = False
UserForm3.Label4.Visible = False
UserForm3.TextBox2.Visible = False
TextBox1.Enabled = False
End If
End Sub

Private Sub OptionButton2_Click()
'If Worksheets("input").Cells(1, 1) = "" Then
UserForm3.TextBox1.Value = mdisc
UserForm3.TextBox2.Value = mdiscerr
'Else
'UserForm3.TextBox1.Text = Worksheets("input").Cells(1, 1)
'UserForm3.TextBox2.Text = Worksheets("input").Cells(1, 2)
'End If
TextBox1.Enabled = True
UserForm3.TextBox2.Visible = True
UserForm3.Label3.Visible = True
UserForm3.Label4.Visible = True

End Sub

