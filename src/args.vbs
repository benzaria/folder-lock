' ================ ARGUMENTS ================

Dim lock, unlock, changeLock
Dim folders() : ReDim folders(-1) ' dynamic array for folder paths

Sub ParseArgs(args)
    If args.Count = 0 Then
        MsgPop "No Arguments provided.", vbExclamation, "Error"
        WScript.Quit
    End If

    Dim i
    ' Loop through each argument
    For i = 0 To args.Count - 1
        Select Case LCase(args(i))
            Case "--lock", "-lock", "/lock", "-l", "/l"
                lock = True
            Case "--unlock", "-unlock", "/unlock", "-u", "/u"
                unlock = True
            Case "--change-lock", "-change", "/change", "-c", "/c"
                changeLock = True
            Case "--everyone", "-everyone", "/everyone", "-e", "/e"
                user = "Everyone"
            Case Else
                ' Append folder to dynamic array
                ReDim Preserve folders(UBound(folders) + 1)
                folders(UBound(folders)) = CleanPath(args(i))
        End Select
    Next

    ' Check if any folders were provided
    If UBound(folders) = -1 Then
        MsgPop "No Folders provided.", vbExclamation, "Error"
        WScript.Quit
    End If    
End Sub
