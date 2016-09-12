Private Sub DelayMs(ms As Long)
    Debug.Print TimeValue(Now)
    Application.Wait (Now + (ms * 0.00000001))
    Debug.Print TimeValue(Now)
End Sub


Sub CLEAR()
Set aRange = Sheets("DATA").Range("A5.S50000")
aRange.ClearContents
End Sub


Sub Desk_Shuffle()


Dim CurrentHost As Object
Set CurrentHost = GetObject(, "ATWin32.AccuTerm")
Set CurrentHost = CurrentHost.ActiveSession
Dim irow As Long
irow = 5
  file = Range("A" & irow).Value
  Desk = Range("B" & irow).Value
  CRAR = Range("D" & irow).Value
  
  Do
  If Range("A" & irow).Value = "" Then
    Application.StatusBar = "Credit AR Add Complete"
    MsgBox "Add Complete"
    Exit Sub
    End If
  
    If Range("B" & irow).Value = 801 Or Range("B" & irow).Value = 800 Or Range("B" & irow).Value = 802 Or Range("B" & irow).Value = 803 Or Range("B" & irow).Value = 804 Or Range("B" & irow).Value = 805 Or Range("B" & irow).Value = 806 Or Range("B" & irow).Value = 807 Or Range("B" & irow).Value = 808 Or Range("B" & irow).Value = 809 Or Range("B" & irow).Value = 831 Or Range("B" & irow).Value = 832 Or Range("B" & irow).Value = 833 Or Range("B" & irow).Value = 834 Or Range("B" & irow).Value = 835 Or Range("B" & irow).Value = 848 Then
    file = Range("A" & irow).Value
    Desk = Range("B" & irow).Value
    CRAR = Range("D" & irow).Value
     
        
        If CurrentHost.GetText(0, 22, 52) = "ENTER SELECTION (.,FILE#,/,STATUS,-nnnnn,Tn,/R,HELP)" Then
        CurrentHost.Output file & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output file & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 48) = "ENTER SELECTION, FILE#,HELP,W,V,C,S,Dn,GC#,/,-,." Then
        CurrentHost.Output "14" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "14" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 20) = "Enter Command,HELP,/" Then
        CurrentHost.Output "5-"
        Call DelayMs(200)
        CurrentHost.Output CRAR & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "5-"
        Call DelayMs(200)
        CurrentHost.Output CRAR & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 73) = "POSTDATES EXIST FOR THIS ACCOUNT.  DO YOU STILL WISH TO DESK CHANGE (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "Y" & ChrW$(13)
        End If
                
        Call DelayMs(200)
                
        If CurrentHost.GetText(0, 22, 51) = "ENTER SELECTION, FILE#,HELP,W,V,C,S,Dn,GC#,/,-,." Then
        CurrentHost.Output "/" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "/" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 15) = "ENTER WHAT (nn)" Then
        CurrentHost.Output "16" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "16" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 14) = "ENTER WHO (nn)" Then
        CurrentHost.Output "17" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "17" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER WHO (nn)" Then
        CurrentHost.Output "12" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "12" & ChrW$(13)
        End If
    
                    
        Range("B" & irow).Value = CRAR
               
    End If
    irow = irow + 1
    Loop

End Sub

Sub CRAR_Replace()
  For Each rw In UsedRange.Rows
    If rw.Columns("D") = "" Then
    rw.Columns("D") = rw.Columns("T")
    End If
    Next rw
    
End Sub

Sub Add_CreditAR()
  Dim CurrentHost As Object
  Set CurrentHost = GetObject(, "ATWin32.AccuTerm")
  Set CurrentHost = CurrentHost.ActiveSession
  Dim irow As Long
  irow = 5
  Do
    If Range("A" & irow).Value = "" Then
    Application.StatusBar = "Credit AR Add Complete"
    MsgBox "Add Complete"
    Exit Sub
    End If
    
    If Range("D" & irow).Value = "" Then
    Range("D" & irow).Value = Range("T" & irow).Value
    file = Range("A" & irow).Value
    CRAR = Range("T" & irow).Value
    
    Call DelayMs(200)
    
        If CurrentHost.GetText(0, 22, 52) = "ENTER SELECTION (.,FILE#,/,STATUS,-nnnnn,Tn,/R,HELP)" Then
        CurrentHost.Output file & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output file & ChrW$(13)
        End If
      
        Call DelayMs(200)
    
        If CurrentHost.GetText(0, 22, 48) = "ENTER SELECTION, FILE#,HELP,W,V,C,S,Dn,GC#,/,-,." Then
        CurrentHost.Output "12" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "12" & ChrW$(13)
        End If
        
        Call DelayMs(200)
    
        If CurrentHost.GetText(0, 22, 11) = "ENTER (n,/)" Then
        CurrentHost.Output "9" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "9" & ChrW$(13)
        End If
      
        Call DelayMs(200)
    
        If CurrentHost.GetText(0, 22, 59) = "ENTER PAYOFF DAYS or DATE (nn,mm/dd/yy,A,S,B,H,/Dn,/F,/H,/)" Then
        CurrentHost.Output "A" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "A" & ChrW$(13)
        End If
     
        Call DelayMs(200)
    
        If CurrentHost.GetText(0, 22, 64) = "THE REHAB FALLOUT COUNT FOR THIS BORROWER IS 1. <CR> TO CONTINUE" Then
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output ChrW$(13)
        End If
          
        If Range("E" & irow).Value = "REHAB3" Then
            
            Call DelayMs(200)
               
            If CurrentHost.GetText(0, 22, 15) = "ENTER (nn,/F,/)" Then
            CurrentHost.Output "/F" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/F" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 18) = "ENTER (nn,/F,/)" Then
            CurrentHost.Output "/F" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/F" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 15) = "ENTER (nn,/B,/)" Then
            CurrentHost.Output "24" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "24" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 32) = "ENTER CREDIT A/R (nnn,X,/n,/,//)" Then
            CurrentHost.Output CRAR & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output CRAR & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 22) = "OK TO FILE (Y,nn,/B,/)" Then
            CurrentHost.Output "Y" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "Y" & ChrW$(13)
            End If
            
            Call DelayMs(1700)
            
            If CurrentHost.GetText(0, 22, 59) = "ENTER PAYOFF DAYS or DATE (nn,mm/dd/yy,A,S,B,H,/Dn,/F,/H,/)" Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 48) = "ENTER SELECTION, FILE#,HELP,W,V,C,S,Dn,GC#,/,-,." Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
    
        Else
        
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 15) = "ENTER (nn,/F,/)" Then
            CurrentHost.Output "/F" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/F" & ChrW$(13)
            End If
        
            Call DelayMs(200)
        
            If CurrentHost.GetText(0, 22, 18) = "ENTER (nn,/F,/B,/)" Then
            CurrentHost.Output "/F" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/F" & ChrW$(13)
            End If
        
            Call DelayMs(200)
        
            If CurrentHost.GetText(0, 22, 15) = "ENTER (nn,/B,/)" Then
            CurrentHost.Output "26" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "26" & ChrW$(13)
            End If
        
            Call DelayMs(200)
        
            If CurrentHost.GetText(0, 22, 32) = "ENTER CREDIT A/R (nnn,X,/n,/,//)" Then
            CurrentHost.Output CRAR & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output CRAR & ChrW$(13)
            End If
            
            Call DelayMs(200)
        
            If CurrentHost.GetText(0, 22, 22) = "OK TO FILE (Y,nn,/B,/)" Then
            CurrentHost.Output "Y" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "Y" & ChrW$(13)
            End If
            
            Call DelayMs(1700)
        
            If CurrentHost.GetText(0, 22, 59) = "ENTER PAYOFF DAYS or DATE (nn,mm/dd/yy,A,S,B,H,/Dn,/F,/H,/)" Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
            
            Call DelayMs(200)
        
            If CurrentHost.GetText(0, 22, 48) = "ENTER SELECTION, FILE#,HELP,W,V,C,S,Dn,GC#,/,-,." Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
            
                  
        End If
  
    End If
    
    If (Range("D" & irow).Value >= 900 Or Range("D" & irow) <= 799 Or Range("D" & irow) = 814 Or Range("D" & irow) = 831 Or Range("D" & irow) = 821 Or Range("D" & irow) = 832 Or Range("D" & irow) = 833 Or Range("D" & irow) = 834 Or Range("D" & irow) = 835 Or Range("D" & irow) = 848) Then
    box = InputBox("What desk does " & Range("A" & irow).Value & " belong to?")
    file = Range("A" & irow).Value
    Range("D" & irow).Value = box
    
            Call DelayMs(200)
        
            If CurrentHost.GetText(0, 22, 52) = "ENTER SELECTION (.,FILE#,/,STATUS,-nnnnn,Tn,/R,HELP)" Then
            CurrentHost.Output file & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output file & ChrW$(13)
            End If
            
            Call DelayMs(200)
        
            If CurrentHost.GetText(0, 22, 48) = "ENTER SELECTION, FILE#,HELP,W,V,C,S,Dn,GC#,/,-,." Then
            CurrentHost.Output "12" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "12" & ChrW$(13)
            End If
    
            Call DelayMs(200)
        
            If CurrentHost.GetText(0, 22, 11) = "ENTER (n,/)" Then
            CurrentHost.Output "9" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "9" & ChrW$(13)
            End If
            
            Call DelayMs(200)
        
            If CurrentHost.GetText(0, 22, 59) = "ENTER PAYOFF DAYS or DATE (nn,mm/dd/yy,A,S,B,H,/Dn,/F,/H,/)" Then
            CurrentHost.Output "A" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "A" & ChrW$(13)
            End If
            
            Call DelayMs(200)
        
            If CurrentHost.GetText(0, 22, 64) = "THE REHAB FALLOUT COUNT FOR THIS BORROWER IS 1. <CR> TO CONTINUE" Then
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output ChrW$(13)
            End If
        
     
        If Range("E" & irow).Value = "REHAB3" Then
        
            Call DelayMs(200)
        
            If CurrentHost.GetText(0, 22, 15) = "ENTER (nn,/F,/)" Then
            CurrentHost.Output "/F" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/F" & ChrW$(13)
            End If
        
            Call DelayMs(200)
        
            If CurrentHost.GetText(0, 22, 18) = "ENTER (nn,/F,/B,/)" Then
            CurrentHost.Output "/F" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/F" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 15) = "ENTER (nn,/B,/)" Then
            CurrentHost.Output "24" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "24" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 32) = "ENTER CREDIT A/R (nnn,X,/n,/,//)" Then
            CurrentHost.Output box & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output box & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 22) = "OK TO FILE (Y,nn,/B,/)" Then
            CurrentHost.Output "Y" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "Y" & ChrW$(13)
            End If
                        
            Call DelayMs(1700)
            
            If CurrentHost.GetText(0, 22, 59) = "ENTER PAYOFF DAYS or DATE (nn,mm/dd/yy,A,S,B,H,/Dn,/F,/H,/)" Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 48) = "ENTER PAYOFF DAYS or DATE (nn,mm/dd/yy,A,S,B,H,/Dn,/F,/H,/)" Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
                      
        Else
        
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 15) = "ENTER (nn,/F,/)" Then
            CurrentHost.Output "/F" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/F" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 18) = "ENTER (nn,/F,/B,/)" Then
            CurrentHost.Output "/F" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/F" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 15) = "ENTER (nn,/B,/)" Then
            CurrentHost.Output "26" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "26" & ChrW$(13)
            End If
        
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 32) = "ENTER CREDIT A/R (nnn,X,/n,/,//)" Then
            CurrentHost.Output box & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output box & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 22) = "OK TO FILE (Y,nn,/B,/)" Then
            CurrentHost.Output "Y" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "Y" & ChrW$(13)
            End If
            
            Call DelayMs(1700)
            
            If CurrentHost.GetText(0, 22, 59) = "ENTER PAYOFF DAYS or DATE (nn,mm/dd/yy,A,S,B,H,/Dn,/F,/H,/)" Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 59) = "ENTER PAYOFF DAYS or DATE (nn,mm/dd/yy,A,S,B,H,/Dn,/F,/H,/)" Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
            
                  
        End If

    End If
    
irow = irow + 1
Loop
    
End Sub

Sub PasswordBreaker()

Dim i As Integer, j As Integer, k As Integer
Dim l As Integer, m As Integer, n As Integer
 Dim i1 As Integer, i2 As Integer, i3 As Integer
 Dim i4 As Integer, i5 As Integer, i6 As Integer
 On Error Resume Next
 For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
 For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
 For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
 For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
 ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
 Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
 If ActiveSheet.ProtectContents = False Then
 MsgBox "One usable password is " & Chr(i) & Chr(j) & _
 Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
 Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
 Exit Sub
 End If
 Next: Next: Next: Next: Next: Next
 Next: Next: Next: Next: Next: Next
 End Sub




