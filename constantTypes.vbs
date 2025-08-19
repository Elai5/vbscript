Dim let1
Dim let2

let1 = "This is a red text"
let2 = "This is a green text"
Msgbox"Current date and time: "& Now

Dim response
response = MsgBox("Are you above 18 years old?", vbYesNo, "Confirmation")

If response = vbYes Then
    MsgBox "Proced to vote"
Else 
    MsgBox "You cannot vote yet."
End If