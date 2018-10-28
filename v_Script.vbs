Class v_Script
	Private pScriptHost, _
		pScriptEngine, _
		pScriptLanguage

	Private Sub Class_Initialize()
		Set pScriptHost = CreateObject("HTMLFile")
		Set pScriptEngine = pScriptHost.parentWindow
		pScriptLanguage = ""
	End Sub


	' Properties:


	Public Property Get Language()
		Language = pScriptLanguage
	End Property

	Public Property Let Language(strLang)
		If LCase(strLang) = "vbscript" Or LCase(strLang) = "jscript" Then pScriptLanguage = strLang
	End Property

	Public Property Get Variable(strVar)
		If Exists(strVar) Then
			If TypeName(Eval("pScriptEngine." & strVar)) = "JScriptTypeInfo" Then
				Set Variable = Eval("pScriptEngine." & strVar)
			Else
				Variable = Eval("pScriptEngine." & strVar)
			End If
		End If
	End Property

	Public Property Let Variable(strVar, strNewVal)
		If Not Exists(strVar) Then	
			If LCase(pScriptLanguage) = "jscript" Then
				pScriptEngine.execScript "var " & strVar & ";", pScriptLanguage
			ElseIf LCase(pScriptLanguage) = "vbscript" Then
				pScriptEngine.execScript "Dim " & strVar, pLanguage
			End If
		End If

		If TypeName(strNewVal) = "String" And Left(strNewVal, 1) = "{" And Right(strNewVal, 1) = "}" Then
			pScriptEngine.execScript strVar & " = " & strNewVal, pScriptLanguage
		Else
			Execute("pScriptEngine." & strVar & " = strNewVal")
		End If
	End Property

	Public Property Set Variable(strVar, objNewObj)
		If Not Exists(strVar) Then
			If LCase(pScriptLanguage) = "jscript" Then
				pScriptEngine.execScript "var " & strVar & ";", pScriptLanguage
			ElseIf LCase(pScriptLanguage) = "vbscript" Then
				pScriptEngine.execScript "Dim " & strVar, pScriptLanguage
			End If
		End If

		Execute("Set pScriptEngine." & strVar & " = objNewObj")
	End Property


	' Methods:


	Public Sub AddCode(strCode)
		On Error Resume Next

		If pScriptLanguage <> "" Then pScriptEngine.execScript strCode, pScriptLanguage

		If Err.Number <> 0 Then
			WScript.Echo "Error occured in 'AddCode()'."
		End If
	End Sub

	Public Function Exists(strVar)
		On Error Resume Next

		Eval("pScriptEngine." & strVar)

		If Err.Number <> 0 Then
			Err.Clear
			Exists = False
		Else
			Exists = True
		End If
	End Function

	Public Function Run(strProcedure, arrArgs)
		Dim strArgs, _
			i

		If Right(strProcedure, 2) = "()" Then strProcedure = Left(strProcedure, Len(strProcedure) - 2)

        'This variable is used to create the string of arguments passed to the pScriptEngine variable 
        '(which is the one that ultimatelly executes the function)
		strArgs = "("

        Dim itHasArgs ' <-- I ADDED THIS LINE
        itHasArgs = false

        'You'll notice in every v_JSON function that does not require arguments, 
        'the Run(strProcedure, arrArgs) function gets called like this, for instance:
         
'	    Public Property Get Items()
'		    Items = Deserialize(pScript.Run("getItems", Array())) <--SEE HERE HOW AN EMPTY ARRAY IS PASSED
'	    End Property

        'while if the function does require arguments, this happens, example:

'       Public Function ValueExists(varValue, blnDeep)
'		    ValueExists = pScript.Run("valueExists", Array(varValue, blnDeep)) <-- SEE HERE HOW A NON-EMPTY ARRAY IS PASSED
'	    End Function
		
        'Here starts the problem because it will only execute line 119 IF THERE ARE ARGUMENTS, which there aren't with some functions, as
        'explained before
        If IsArray(arrArgs) Then
			For i = 0 to UBound(arrArgs)
				strArgs = strArgs & "arrArgs(" & i & "), "
                itHasArgs = true  ' <-- I ADDED THIS LINE
			Next
		End If

        'causing strArgs to arrive to this point with only "("

        If itHasArgs Then ' <-- I ADDED THIS LINE
            strArgs = Left(strArgs, Len(strArgs) - 2) & ")"
        Else ' <-- I ADDED THIS LINE
		    strArgs = strArgs & ")" ' <-- I ADDED THIS LINE
        End If ' <-- I ADDED THIS LINE

        'in the original code only line 127 existed, this used to happen:
        '1. Len(strArgs) = Len("(") = 1
        '2. Len(strArgs) - 2 = -1
        '3. Left(strArgs, -1) made it all explode since it can't stand that negative argument
        
        'and that is why some functions worked and some didn't
 
		If TypeName(Eval("pScriptEngine." & strProcedure & strArgs)) = "JScriptTypeInfo" Then
			Set Run = Eval("pScriptEngine." & strProcedure & strArgs)
		Else
			Run = Eval("pScriptEngine." & strProcedure & strArgs)
		End If
	End Function

	Public Sub Reset()
		Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
		Set pScriptHost = Nothing
		Set pScriptEngine = Nothing
	End Sub
End Class