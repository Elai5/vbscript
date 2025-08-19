Dim Var1
Dim Var2

Call add()
Function add()
	var1 = 10
	var2 = 15
	Dim var3
	var3 = var1 + var2

	Msgbox "this is for variant declared at procedure level,hence answer is " & var3
End Function

Msgbox "this is declared on script level "&  var1
Msgbox "this is declared at script level " & var2
Msgbox "var3 has no scope outside the procedure. Prints empty" & var3