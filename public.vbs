Dim var1
Dim var2
Public var3

Call add()
Function add()
var1 = 10
var2 = 15
var3 = var1 + var2

Msgbox "sum var1 and var2 is " & var3
End Function

Msgbox "declared at script level " & var1
Msgbox "declared at script level " &  var2
Msgbox "declared as public "& var3

