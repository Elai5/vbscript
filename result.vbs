Dim result
result = MsgBox("An error occured, how would you lke to proceed?", vbAbortRetryIgnore, "Error")

if result = vbAbort Then
     MsgBox "You chose abort"
    MsgBox "A serious error occured!" ,vbCritical, "critical error"
ElseIf result = vbRetry Then
    MsgBox "You chose retry"
    MsgBox "Do you wish to retry?" ,vbQuestion +vbYesNo, "save changes"
ElseIf result = vbIgnore Then       
    MsgBox "Ignore please"
    MsgBox"Operation terminated.", vbInformation, "Information"
End If