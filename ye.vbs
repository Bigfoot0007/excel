
Function getflag(preresult)
   preresult = Trim(preresult)
   Rem preresult=right(preresult,2)
   getflag = Mid(preresult, 4, 10)
End Function

Function updown(preresult, value1, value2)
    preflag = getflag(preresult)
    preflagf = Left(preflag, 1)
    preflagn = Mid(preflag, 2, 10)
    a = Val(value1)
    b = Val(value2)

    If (a < b) Then
        
        If (preflagf = "+") Then
            updown = preflag & " +" & preflagn + 1
        End If
        If (preflagf = "-") Then
            updown = preflag & " +1"
        End If
        If (preflagf = "平") Then
            updown = preflag & " +1"
        End If
    End If
    If (a = b) Then
   
        If (preflagf = "+") Then
            updown = preflag & " 平1"
        End If
        If (preflagf = "-") Then
            updown = preflag & " 平1"
        End If
        If (preflagf = "平") Then
            updown = preflag & " 平" & preflagn + 1
        End If
    End If
    If (a > b) Then
        If (preflagf = "+") Then
            updown = preflag & " -1"
        End If
        If (preflagf = "-") Then
            updown = preflag & " -" & preflagn + 1
        End If
        If (preflagf = "平") Then
            updown = preflag & " -1"
        End If
    End If
    
End Function


Function updownflag(preresult, value1, value2)
    preflag = getflag(preresult)
    preflagf = Left(preflag, 1)
    preflagn = Mid(preflag, 2, 10)
    a = Val(value1)
    b = Val(value2)
    
    updownflag = ""
    If (a < b) Then
        
        Rem If (preflagf = "+") Then
        Rem     updown = preflag & " +" & preflagn + 1
        Rem End If
        If (preflagf = "-") Then
            updownflag = "↑"
        End If
        If (preflagf = "平") Then
            updownflag = "↑"
        End If
    End If
    If (a > b) Then
        If (preflagf = "+") Then
            updownflag = "↓"
        End If

        If (preflagf = "平") Then
            updownflag = "↓"
        End If
    End If
    
End Function

