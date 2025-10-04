                                          VB EXP1
Private Sub clear_Click()
output.Text = " "
End Sub
Private Sub exit_Click()
Dim d As Variant
d = MsgBox("do you want to exit", vbYesNo + vbQuestion, "thank you")
If d = vbYes Then
End
Exit Sub
End If
End Sub
Private Sub sum_Click()
s = 0
i = 1
Do While i <= 100
s = s + i
i = i + 1
Loop
output.Text = s
End Sub


                                         VB EXP3

Private Sub CLEAR_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
avgvalue.Caption = ""
valresult.Caption = ""
End Sub
Private Sub EXIT_Click()
Dim d As Integer
d = MsgBox("Do you want to exit", vbYesNo + vbQuestion, "Thank You")
If d = vbYes Then
End
Else
Exit Sub
End If
End SubPrivate Sub result_Click()
Dim t As Boolean
t = True
Dim m1, m2, m3, m4, m5, m6, avg As Variant
m1 = Val(Text1.Text)
m2 = Val(Text2.Text)
m3 = Val(Text3.Text)
m4 = Val(Text4.Text)
m5 = Val(Text5.Text)
m6 = Val(Text6.Text)
avg = (m1 + m2 + m3 + m4 + m5 + m6) / 6
avgvalue = avg
Select Case t
Case m1 < 40 Or m2 < 40 Or m3 < 40 Or m4 < 40 Or m5 < 40 Or m6 < 40
valresult.Caption = "Fail"
Case avg >= 40 And avg < 60 And m1 >= 40 And m2 >= 40 And m3 >= 40 And m4 >= 40 And m5 >= 40 And 
m6 >= 40
valresult.Caption = "II-Class"
Case avg >= 60
valresult.Caption = "I-Class"
End Select
End Sub


                                        VB EXP 4

Private Sub AREA_Click()
Dim a As Double
If shape.Text = "CIRCLE" Then
a = 3.14 * CDbl(RADIUS.Text) * CDbl(RADIUS.Text)
MsgBox ("Area of the circle is = " & a)
Else
a = CDbl(LENGTH.Text) * CDbl(BREADTH.Text)
MsgBox ("Area of the rectangle is = " & a)
End If
End Sub
Private Sub EXIT_Click()
Dim d As Variant
d = MsgBox("Do you want to exit", vbYesNo + vbQuestion, "Thank You")
If d = vbYes Then
End
Exit Sub
End If
End Sub
Private Sub shape_Click()
If shape.Text = "CIRCLE" Then
RADIUS.Visible = True
Label2.Visible = True
Label3.Visible = False
Label4.Visible = False
LENGTH.Visible = False
BREADTH.Visible = False
ElseIf shape.Text = "RECTANGLE" Then
RADIUS.Visible = False
Label2.Visible = False
Label3.Visible = True
Label4.Visible = True
LENGTH.Visible = True
BREADTH.Visible = True
End If
End Sub

                                        VB EXP 5


Private Sub backblue_Click(Index As Integer)
Shape1.BackColor = vbBlue
End Sub
Private Sub backred_Click(Index As Integer)
Shape1.BackColor = vbRed
End Sub
Private Sub backyellow_Click()
Shape1.BackColor = vbYellow
End Sub
Private Sub blue_Click(Index As Integer)
Shape1.BorderColor = vbBlue
End Sub
Private Sub five_Click()
Shape1.BorderWidth = 5
End Sub
Private Sub green_Click()
Shape1.BorderColor = vbGreen
End Sub
Private Sub red_Click(Index As Integer)
Shape1.BorderColor = vbRed
End Sub
Private Sub ten_Click()
Shape1.BorderWidth = 10
End Sub
Private Sub two_Click()
Shape1.BorderWidth = 2
End Sub

                                            EXP7
Private Sub close_Click()
End
End Sub
Private Sub copy_Click()
Clipboard.SetText (tb.SelText)
End Sub
Private Sub cut_Click()
Clipboard.SetText (tb.SelText)
tb.SelText = ""
End Sub
Private Sub new_Click()
tb.Text = ""
End Sub
Private Sub open_Click()
CommonDialog1.ShowOpen
tb.LoadFile (CommonDialog1.FileName)
End Sub
Private Sub paste_Click()
tb.SelText = Clipboard.GetText
End Sub
Private Sub save_Click()
CommonDialog1.ShowSave
tb.SaveFile (CommonDialog1.FileName)
MsgBox ("Your file has been saved")
End Sub




