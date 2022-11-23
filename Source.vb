Dim num1 As Double
Dim num2 As Double
Dim chr As String
Dim res As Boolean

Function calc_last()
    num2 = textfrm.Text

    If chr = "*" Then
        num1 = num1 * num2
        num2 = 0
    End If

    If chr = "/" Then
        If num2 = "0" Then
            textfrm.Text = "Err: DV/0"
        Else
            num1 = num1 / num2
            num2 = 0
        End If
    End If

    If chr = "-" Then
        num1 = num1 - num2
        num2 = 0
    End If

    If chr = "+" Then
        num1 = num1 + num2
        num2 = 0
    End If
End Function

Private Sub b1_Click()
    If textfrm.Text = "0" Then
        textfrm.Text = "1"
    Else
        If res Then
            textfrm = "1"
            res = False
        Else
            textfrm.Text = textfrm.Text + "1"
        End If
    End If
End Sub

Private Sub b2_Click()
    If textfrm.Text = "0" Then
        textfrm.Text = "2"
    Else
        If res Then
            textfrm = "2"
            res = False
        Else
            textfrm.Text = textfrm.Text + "2"
        End If
    End If
End Sub

Private Sub b3_Click()
    If textfrm.Text = "0" Then
        textfrm.Text = "3"
    Else
        If res Then
            textfrm = "3"
            res = False
        Else
            textfrm.Text = textfrm.Text + "3"
        End If
    End If
End Sub

Private Sub b4_Click()
    If textfrm.Text = "0" Then
        textfrm.Text = "4"
    Else
        If res Then
            textfrm = "4"
            res = False
        Else
            textfrm.Text = textfrm.Text + "4"
        End If
    End If
End Sub

Private Sub b5_Click()
    If textfrm.Text = "0" Then
        textfrm.Text = "5"
    Else
        If res Then
            textfrm = "5"
            res = False
        Else
            textfrm.Text = textfrm.Text + "5"
        End If
    End If
End Sub

Private Sub b6_Click()
    If textfrm.Text = "0" Then
        textfrm.Text = "6"
    Else
        If res Then
            textfrm = "6"
            res = False
        Else
            textfrm.Text = textfrm.Text + "6"
        End If
    End If
End Sub

Private Sub b7_Click()
    If textfrm.Text = "0" Then
        textfrm.Text = "7"
    Else
        If res Then
            textfrm = "7"
            res = False
        Else
            textfrm.Text = textfrm.Text + "7"
        End If
    End If
End Sub

Private Sub b8_Click()
    If textfrm.Text = "0" Then
        textfrm.Text = "8"
    Else
        If res Then
            textfrm = "8"
            res = False
        Else
            textfrm.Text = textfrm.Text + "8"
        End If
    End If
End Sub

Private Sub b9_Click()
    If textfrm.Text = "0" Then
        textfrm.Text = "9"
    Else
        If res Then
            textfrm = "9"
            res = False
        Else
            textfrm.Text = textfrm.Text + "9"
        End If
    End If
End Sub

Private Sub b0_Click()
    If textfrm.Text = "0" Then
        textfrm.Text = "0"
    Else
        If res Then
            textfrm = "0"
            res = False
        Else
            textfrm.Text = textfrm.Text + "0"
        End If
    End If
End Sub

Private Sub bclr_Click()
    num1 = 0
    num2 = 0
    chr = ""
    textfrm.Text = "0"
    res = False
End Sub

Private Sub bins_Click()
    On Error GoTo except
    Range(insertbox.Text).Value = textfrm.Text
    Exit Sub
except:
    insertbox.Text = "Err"
End Sub

Private Sub bptr_Click()
    If textfrm.Text > 0 Then
        textfrm.Text = textfrm.Text + "."
    End If
End Sub

Private Sub br_Click()
    num2 = textfrm.Text

    If chr = "*" Then
        textfrm.Text = num1 * num2
        res = True
        num1 = 0
        num2 = 0
    End If

    If chr = "/" Then
        If num2 = "0" Then
            textfrm.Text = "Err: DV/0"
        Else
            textfrm.Text = num1 / num2
            res = True
            num1 = 0
            num2 = 0
        End If
    End If

    If chr = "-" Then
        textfrm.Text = num1 - num2
        res = True
        num1 = 0
        num2 = 0
    End If

    If chr = "+" Then
        textfrm.Text = num1 + num2
        res = True
        num1 = 0
        num2 = 0
    End If
End Sub

Private Sub bym_Click()
    If num1 = 0 Then
        num1 = textfrm.Text
        chr = "*"
        textfrm.Text = "0"
    Else
        calc_last()
        num2 = textfrm.Text
        textfrm.Text = "0"
        chr = "*"
    End If

    If res Then
        res = False
    End If
End Sub

Private Sub bdv_Click()
    If num1 = 0 Then
        num1 = textfrm.Text
        chr = "/"
        textfrm.Text = "0"
    Else
        calc_last()
        num2 = textfrm.Text
        textfrm.Text = "0"
        chr = "/"
    End If

    If res Then
        res = False
    End If
End Sub

Private Sub bp_Click()
    If num1 = 0 Then
        num1 = textfrm.Text
        chr = "+"
        textfrm.Text = "0"
    Else
        calc_last()
        num2 = textfrm.Text
        textfrm.Text = "0"
        chr = "+"
    End If

    If res Then
        res = False
    End If
End Sub

Private Sub bm_Click()
    If num1 = 0 Then
        num1 = textfrm.Text
        chr = "-"
        textfrm.Text = "0"
    Else
        calc_last()
        num2 = textfrm.Text
        textfrm.Text = "0"
        chr = "-"
    End If

    If res Then
        res = False
    End If
End Sub

Private Sub bperc_Click()
    If textfrm.Text > 0 Then
        num1 = textfrm.Text / 100
        textfrm.Text = num1
    End If
End Sub

Private Sub bactinsert_Click()
    ActiveCell.Value = textfrm.Text
End Sub
