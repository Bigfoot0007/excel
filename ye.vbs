Function getflag1(preresult)
   preresult = Trim(preresult)
   Rem preresult=right(preresult,2)
   temp = Split(preresult)
   getflag1 = temp(0)
End Function

Function getflag(preresult)
   preresult = Trim(preresult)
   Rem preresult=right(preresult,2)
   temp = Split(preresult)
   getflag = temp(1)
End Function
Rem ＋3 －1
Function updown(preresult, value1, value2)
    preflag = getflag(preresult)
    preflag1 = getflag1(preresult)
    preflagf = Left(preflag, 1)
    preflagn = Mid(preflag, 2, 10)


    Rem MsgBox (preresult & ",LEN:" & Len(preresult) & ">>" & preflag & ",LEN:" & Len(preflag) & "-preflagf," & preflagf & " preflagn," & preflagn)

    a = Val(value1)
    b = Val(value2)
    
    updown = "ERROR ERROR"
    If (a < b) Then
        Rem MsgBox (preflagf)
        If (preflagf = "＋" Or preflagf = "+") Then
        
            updown = preflag1 & " ＋" & preflagn + 1
        End If
        If (preflagf = "－" Or preflagf = "-") Then
            updown = preflag & " ＋1"
        End If
        If (preflagf = "平") Then
            updown = preflag & " ＋1"
        End If
    End If
    If (a = b) Then
        If (preflagf = "＋" Or preflagf = "+") Then
            updown = preflag & " 平1"
        End If
        If (preflagf = "－" Or preflagf = "-") Then
            updown = preflag & " 平1"
        End If
        If (preflagf = "平") Then
            updown = preflag1 & " 平" & preflagn + 1
        End If
    End If
    If (a > b) Then
        If (preflagf = "＋" Or preflagf = "+") Then
            updown = preflag & " －1"
        End If
        If (preflagf = "－" Or preflagf = "-") Then
            updown = preflag1 & " －" & preflagn + 1
        End If
        If (preflagf = "平") Then
            updown = preflag & " －1"
        End If
    End If
    

End Function


Function updownflag(preresult, value1, value2)
    preflag = getflag(preresult)
    preflagf = Left(preflag, 1)
    preflagn = Mid(preflag, 2, 10)
    a = Val(value1)
    b = Val(value2)
    
    updownflag = FormatNumber(b, 2)
    If (a < b) Then

        If (preflagf = "－" Or preflagf = "-") Then
            updownflag = value2 & "↑"
        End If
        If (preflagf = "平") Then
            updownflag = value2 & "↑"
        End If
    End If
    If (a > b) Then
        If (preflagf = "＋" Or preflagf = "+") Then
            updownflag = value2 & "↓"
        End If

        If (preflagf = "平") Then
            updownflag = value2 & "↓"
        End If
    End If
    
End Function
