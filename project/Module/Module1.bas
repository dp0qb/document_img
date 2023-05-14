Attribute VB_Name = "Module1"

Public gemoc_on As Boolean
Dim gemoc_input(14) As Variant
Dim converted(11) As Variant
Public Const mdisc = 0.85
Public Const mdiscerr = 0.05

Sub convert()
'setting default options for mass discrimination
'If Worksheets("input").Cells(1, 1) = "" Then

UserForm3.TextBox1.Value = mdisc
UserForm3.TextBox2.Value = mdiscerr
'Else
'UserForm3.TextBox1.Text = Worksheets("input").Cells(1, 1)
'UserForm3.TextBox2.Text = Worksheets("input").Cells(1, 2)
'End If

UserForm3.OptionButton1.Value = True
UserForm3.Show


massdisc = UserForm3.TextBox1.Value
massdiscerr = UserForm3.TextBox2.Value
[a1] = massdisc
[b1] = massdiscerr

j = 7

Application.ScreenUpdating = False
    Do Until Cells(j, 1) = ""
    Sheets("input").Select
    'Reading data
        i = 1
        Do Until i = 15
            gemoc_input(i) = Cells(j, i)
            i = i + 1
        Loop
        i = 1
        Do Until i = 10
            converted(i) = gemoc_input(i)
            i = i + 1
        Loop
    ' calculating U/Th ratio and error
    
        converted(10) = massdisc * gemoc_input(14) / gemoc_input(13)
        'converted(11) = Sqr(gemoc_input(14) + gemoc_input(14) ^ 2 / gemoc_input(13)) / gemoc_input(13)

        converted(11) = Sqr(massdiscerr ^ 2 * gemoc_input(14) ^ 2 / gemoc_input(13) ^ 2 + gemoc_input(14) * massdisc ^ 2 / gemoc_input(13) ^ 2 + massdiscerr ^ 2 * gemoc_input(14) ^ 2 / gemoc_input(13) ^ 3)
        
    
    'writing data
        Sheets("data").Activate
        i = 1
        Do Until i = 12
        Cells(j, i) = Format(converted(i), "###.#####")
         i = i + 1
        Loop
  
       j = j + 1
     Sheets("input").Select
   Loop
   
    Worksheets("data").Visible = True
    Worksheets("data").Select
    Application.ScreenUpdating = True
   
    PbCorr2
End Sub
