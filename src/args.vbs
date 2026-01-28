' ================ ARGUMENTS ================

Dim lock, unlock, changeLock, folders(), joinedFolders
Dim args, arg_i, arg_j : arg_j = 0
Set args = WScript.Arguments

If args.Count <> 0 Then
    For arg_i = 0 To args.Count - 1
        Select Case LCase(args(arg_i))
            Case "--lock", "-lock", "/lock", "-l", "/l"
                lock = True : arg_jpush(1)
            Case "--unlock", "-unlock", "/unlock", "-u", "/u"
                unlock = True : arg_jpush(1)
            Case "--change-lock", "-change", "/change", "-c", "/c"
                changeLock = True : arg_jpush(1)
            Case "--everyone", "-everyone", "/everyone", "-e", "/e"
                user = "Everyone" : arg_jpush(1)
            Case Else
                ReDim Preserve folders(args.Count - 1 - arg_j)
                folders(arg_i - arg_j) = CleanPath(args(arg_i))
        End Select
    Next
Else
    MsgBox "No Arguments provided.", vbExclamation, "Error"
    WScript.Quit
End If

If UBound(folders) = -1 Then
    MsgBox "No Folders provided.", vbExclamation, "Error"
    WScript.Quit
End If

joinedFolders = """" & Join(folders, """ """) & """"

Sub arg_jpush(affset)
    arg_j = arg_j + 1
    ReDim Preserve folders(args.Count - 1 - arg_j)
End Sub

' WScript.Echo lock, unlock, changeLock, user
' WScript.Echo joinedFolders, UBound(folders)

' WScript.Quit
