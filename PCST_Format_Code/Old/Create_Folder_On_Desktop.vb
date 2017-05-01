Sub Create_Folder()

	Dim User_Name As String
	Dim Project_Name As String
	Dim Save_Path As String
	Dim Name_Input_Checker As Integer
	Dim Project_Name_Checker As Integer

	Do 'Checks to confirm the user entered an incorrect userid
		Name_Input_Checker = 0
			  User_Name = InputBox("Please enter your Cerner userID. ex. BE042983")

		If Len(User_Name) <> 8 Then
			Msgbox("Lets try this again..." & vbNewLine & "Please enter your user_ID. No spaces" & vbNewLine & "ex. BE042983")

		Else
			Name_Input_Checker = 1
		End If

	Loop While Name_Input_Checker = 0

	Do 'Checks to confirm user entered correct project name
		Project_Name_Checker = 0
			  Project_Name = InputBox("Please enter the abbreviation for this project. ex. NBRO")

		If Len(Project_Name) = 4 Or Len(Project_Name) = 7 Then 'If length of user inut incorrect, prompt user to try again.
			Project_Name_Checker = 1
		Else
			MsgBox("Lets try this again.... Please enter the project name..." & vbNewLine & "ex. NBRO")
		End If

	Loop While Project_Name_Checker = 0

	Save_Path = "C:\Users\" & User_Name & "\Documents\" & Project_Name & "_" & "PCST_Files"

	If Len(Dir(Save_Path, vbDirectory)) = 0 Then 'If the file already exists then do nothing. Else make it.
		MkDir Save_Path 'Creates the folder
		Name_Input_Checker = 1
	Else
		Msgbox("Looks like the folder already exists... Moving on!") 'Folder already exists so continuing on.
	End If

End Sub
