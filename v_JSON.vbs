Include "v_Script"

Class v_JSON
	Private pScript, _
		pJSON

	Private Sub Class_Initialize()
		Set pScript = New v_Script
    		Set pJSON = CreateObject("Scripting.Dictionary")

		With pScript
			.Language = "JScript"
			.AddCode("function getKeys(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; }")
			.AddCode("function getKeyValue(jsonObj, key) { return jsonObj[key]; }")
			.AddCode("function getKeyType(jsonObj, key) { return typeof(jsonObj[key]); }")
			.AddCode("function isArray(jsonObj) { return Object.prototype.toString.call(jsonObj) === '[object Array]'; }")
			.AddCode("function getArrayLength(jsonArr) { return jsonArr.length; }")
			.AddCode("function getArrayItem(jsonArr, i) { return jsonArr[i]; }")
			.AddCode("function getArrayItemType(jsonArr, i) { return typeof(jsonArr[i]); }")
			.AddCode("function stringify(jsonObj) { var t = typeof(jsonObj); if (t != ""object"" || jsonObj === null) { if (t == ""string"") jsonObj = '""'+jsonObj+'""'; return String(jsonObj); } else { var n, v, jsonStr = [], arr = (jsonObj && jsonObj.constructor == Array); for (n in jsonObj) { v = jsonObj[n]; t = typeof(v); if (t == ""string"") v = '""'+v+'""'; else if (t == ""object"" && v !== null) v = stringify(v); jsonStr.push((arr ? """" : '""' + n + '"":') + String(v)); } return (arr ? ""["" : ""{"") + String(jsonStr) + (arr ? ""]"" : ""}""); } };")
		End With
	End Sub


	' Properties


	Public Default Property Get Item(strKey)
    		If IsObject(pJSON(strKey)) Then
        		Set Item = pJSON(strKey)
    		Else
        		Item = pJSON(strKey)
    		End If
	End Property

	Public Property Get Items()
		Items = pJSON.Items
	End Property

	Public Property Let Key(strKey, strNewKey)
		pJSON.Key(strKey) = strNewKey
	End Property

	Public Property Get Keys()
		Keys = pJSON.Keys
	End Property

	Public Property Get Count()
		Count = pJSON.Count 
	End Property

	
	' Methods


	Public Sub Add(strKey, varItem)
		If VerifyContent(varItem) And Not pJSON.Exists(strKey) Then pJSON.Add strKey, varItem
	End Sub

	Public Sub Remove(strKey)
		pJSON.Remove(strKey)
	End Sub

	Public Sub Clear()
		pJSON.RemoveAll()
	End Sub

	Public Function Exists(strKey, blnDeep)
		If blnDeep Then
			If IsEmpty(SearchContent(Me, strKey)) Then
				Exists = False
			Else
				Exists = True
			End If
		Else
			Exists = pJSON.Exists(strKey)
		End If
	End Function

	Public Function Find(strKey)
		If IsObject(SearchContent(Me, strKey)) Then
			Set Find = SearchContent(Me, strKey)
		Else
			Find = SearchContent(Me, strKey)
		End If	
	End Function

	Public Sub Load(strFile)
		Dim objFSO
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		FromString objFSO.OpenTextFile(strFile, 1).ReadAll()
		Set FSO = Nothing
	End Sub

	Public Sub Save(strFilename)
		Dim objFSO, _
			objJsonFile

		Set objFSO = CreateObject("Scripting.FileSystemObject")

		If objFSO.FileExists(strFilename) Then
			Set objJsonFile = objFSO.OpenTextFile(strFilename, 2, True)			
		Else
			Set objJsonFile = objFSO.CreateTextFile(strFilename, True)
		End If

		With objJsonFile
			.WriteLine Me.ToString()
			.Close()
		End With
	End Sub

	Public Sub FromString(strJSON)
		If TypeName(strJSON) = "String" Then
			pScript.Variable("json") = strJSON
			Serialize pScript.Variable("json")
		End If
	End Sub

	Public Function ToString()
		ToString = Deserialize(pJSON)
	End Function

	
	' Helper Methods


	Private Sub Serialize(objJSON)
		Dim k

		For Each k In GetKeys(objJSON)
			Select Case GetKeyType(objJSON, k)
				Case "object":
					pJSON.Add k, CreateJSONObject(GetKeyValue(objJSON, k))
				Case "array":
					pJSON.Add k, CreateJSONArray(GetKeyValue(objJSON, k))
				Case "null":
					pJSON.Add k, Null
				Case "string", "number", "boolean":
					pJSON.Add k, GetKeyValue(objJSON, k)
			End Select
		Next
	End Sub

	Private Function Deserialize(varJSON)
		Dim strReturn

		Select Case TypeName(varJSON)
			Case "v_JSON":
				strReturn = varJSON.ToString()
			Case "Dictionary":
				Dim key

				For Each key in varJSON.Keys
					Select Case TypeName(varJSON.Item(key))
						Case "v_JSON", "ArrayList":
							strReturn = strReturn & ", """ & key & """: " & Deserialize(varJSON.Item(key))
						Case "Null":
							strReturn = strReturn & ", """ & key & """: null"
						Case "String":
							strReturn = strReturn & ", """ & key & """: """ & varJSON.Item(key) & """"
						Case "Boolean":
							strReturn = strReturn & ", """ & key & """: " & LCase(varJSON.Item(key))
						Case Else:
							strReturn = strReturn & ", """ & key & """: " & varJSON.Item(key)
					End Select
				Next

				strReturn = "{ " & Right(strReturn, Len(strReturn) - 2)
				strReturn = strReturn & " }"
			Case "ArrayList":
				Dim i

				For i = 0 To varJSON.Count - 1
					Select Case TypeName(varJSON.Item(i))
						Case "v_JSON", "ArrayList":
							strReturn = strReturn & ", " & Deserialize(varJSON.Item(i))
						Case "Null":
							strReturn = strReturn & ", null"
						Case "String":
							strReturn = strReturn & ", """ & varJSON.Item(i) & """"
						Case "Boolean":
							strReturn = strReturn & ", " & LCase(varJSON.Item(i))
						Case Else:
							strReturn = strReturn & ", " & varJSON.Item(i)
					End Select
				Next

				strReturn = "[ " & Right(strReturn, Len(strReturn) - 2)
				strReturn = strReturn & " ]"
		End Select

		Deserialize = strReturn
	End Function

	Private Function GetKeys(objJSON)
		GetKeys = Split(pScript.Run("getKeys", Array(objJSON)), ",")
	End Function

	Private Function GetKeyType(objJSON, strKey)
		If pScript.Run("getKeyType", Array(objJSON, strKey)) = "object" Then
			If TypeName(GetKeyValue(objJSON, strKey)) = "Null" Then
				GetKeyType = "null"
			ElseIf IsArray(GetKeyValue(objJSON, strKey)) Then
				GetKeyType = "array"
			Else
				GetKeyType = "object"
			End If
		Else
			GetKeyType = pScript.Run("getKeyType", Array(objJSON, strKey))
		End If
	End Function

	Private Function GetKeyValue(objJSON, strKey)
		If IsObject(pScript.Run("getKeyValue", Array(objJSON, strKey))) Then
			Set GetKeyValue = pScript.Run("getKeyValue", Array(objJSON, strKey))
		Else
			GetKeyValue = pScript.Run("getKeyValue", Array(objJSON, strKey))
		End If
	End Function

	Private Function IsArray(objJSON)
		IsArray = pScript.Run("isArray", Array(objJSON))
	End Function

	Private Function GetArrayLength(objJSONArr)
		GetArrayLength = pScript.Run("getArrayLength", Array(objJSONArr))
	End Function

	Private Function GetArrayItem(objJSONArr, intIndex)
		If IsObject(pScript.Run("getArrayItem", Array(objJSONArr, intIndex))) Then
			Set GetArrayItem = pScript.Run("getArrayItem", Array(objJSONArr, intIndex))
		Else
			GetArrayItem = pScript.Run("getArrayItem", Array(objJSONArr, intIndex))
		End If
	End Function

	Private Function GetArrayItemType(objJSONArr, intIndex)
		If pScript.Run("getArrayItemType", Array(objJSONArr, intIndex)) = "object" Then
			If TypeName(GetArrayItem(objJSONArr, intIndex)) = "Null" Then
				GetArrayItemType = "null"
			ElseIf IsArray(GetArrayItem(objJSONArr, intIndex)) Then
				GetArrayItemType = "array"
			Else
				GetArrayItemType = "object"
			End If
		Else
			GetArrayItemType = pScript.Run("getArrayItemType", Array(objJSONArr, intIndex))
		End If
	End Function

	Private Function CreateJSONObject(objJSON)
		Dim objJsonObj: Set objJsonObj = New v_JSON
		objJsonObj.FromString pScript.Run("stringify", Array(objJSON))
		Set CreateJSONObject = objJsonObj
	End Function

	Private Function CreateJSONArray(objJSONArr)
		Dim objArray, _
			i

		Set objArray = CreateObject("System.Collections.ArrayList")

		For i = 0 To GetArrayLength(objJSONArr) - 1
			Select Case GetArrayItemType(objJSONArr, i)
				Case "object":
					objArray.Add CreateJSONObject(GetArrayItem(objJSONArr, i))
				Case "array":
					objArray.Add CreateJSONArray(GetArrayItem(objJSONArr, i))
				Case "null":
					objArray.Add Null
				Case "string", "number", "boolean":
					objArray.Add GetArrayItem(objJSONArr, i)
			End Select
		Next

		Set CreateJSONArray = objArray
	End Function

	Private Function VerifyContent(varContent)
		Dim blnVerified 

		Select Case TypeName(varContent)
			Case "String", "Null", "Boolean", "Integer", "Long", "Single", "Double", "Date", "Currency":
				blnVerified = True
			Case "v_JSON":
				Dim key

				blnVerified = True

				For Each key in varContent.Keys()
					If Not VerifyContent(varContent.Item(key)) Then
						blnVerified = False
						Exit For
					End If
				Next
			Case "ArrayList":
				Dim i

				blnVerified = True

				For i = 0 To varContent.Count - 1
					If Not VerifyContent(varContent.Item(i)) Then
						blnVerified = False
						Exit For
					End If
				Next
			Case Else:
				blnVerified = False
		End Select

		VerifyContent = blnVerified
	End Function

	Private Function SearchContent(varContent, strKey)
		Dim varReturn

		If TypeName(varContent) = "v_JSON" Then
			If varContent.Exists(strKey, False) Then
				If IsObject(varContent.Item(strKey)) Then
					Set SearchContent = varContent.Item(strKey)
					Exit Function
				Else
					SearchContent = varContent.Item(strKey)
					Exit Function
				End If
			Else
				Dim key

				For Each key in varContent.Keys()
					If TypeName(varContent.Item(key)) = "v_JSON" Or TypeName(varContent.Item(key)) = "ArrayList" Then
						If IsObject(SearchContent(varContent.Item(key), strKey)) Then
							Set SearchContent = SearchContent(varContent.Item(key), strKey)
							Exit Function
						Else
							varReturn = SearchContent(varContent.Item(key), strKey)

							If Not IsEmpty(varReturn) Then
								SearchContent = varReturn
								Exit Function
							End If
						End If
					End If
				Next
			End If 
		ElseIf TypeName(varContent) = "ArrayList" Then
			Dim i

			For i = 0 To varContent.Count - 1
				If TypeName(varContent.Item(i)) = "v_JSON" Or TypeName(varContent.Item(i)) = "ArrayList" Then
					If IsObject(SearchContent(varContent.Item(i), strKey)) Then
						Set SearchContent = SearchContent(varContent.Item(i), strKey)
						Exit Function
					Else
						varReturn = SearchContent(varContent.Item(i), strKey)

						If Not IsEmpty(varReturn) Then
							SearchContent = varReturn
							Exit Function
						End If
					End If
				End If
			Next
		End If
	End Function

	Private Sub Class_Terminate()
		Set pScript = Nothing
    		Set pJSON = Nothing
	End Sub 
End Class
