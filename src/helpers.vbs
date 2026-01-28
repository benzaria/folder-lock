' ================= HELPERS =================

Function CleanPath(path)
    If Right(path, 1) = "\" Or Right(path, 1) = "/" Then
        CleanPath = Left(path, Len(path) - 1)
    Else
        CleanPath = path
    End If
End Function

Function GetAccess(path, user)
    Dim command, hasN
    
    command = "cmd /c icacls """ & path & """ | findstr /i /c:""" & user & """ /c:Eveyone | findstr /i ""(N)"""
    hasN = shell.Run(command, 0, True)
    
    ' 0 = no access
    If hasN = 0 Then
        GetAccess = True
    Else
        GetAccess = False
    End If
End Function

Function GetPassword(path)
    Dim pass : pass = GetPassword_prv(path)
    
    If pass = "" Then
        pass = InputBox("No password found." & vbCrLf & "Please enter a password to use:", "Set Folder Password")
        SaveIni path, pass
    End If

    GetPassword	= pass
End Function

Function GetPassword_prv(path)
    Dim iniPath, ini, inSection
    iniPath = path & "\Desktop.ini" : inSection = False

    If Not fso.FileExists(iniPath) Then
        GetPassword_prv = ""
        Exit Function
    End If

    Set ini = fso.OpenTextFile(iniPath, 1)

    Do While Not ini.AtEndOfStream
        line = Trim(ini.ReadLine)

        If LCase(line) = "[folderlock]" Then 
            inSection = True
        ElseIf LCase(line) = "[" Then
            inSection = False
        End If

        If inSection And LCase(Left(line, 9)) = "password=" Then
            GetPassword_prv = Mid(line, 10)
            ini.Close
            Exit Function
        End If
    Loop

    ini.Close
    GetPassword_prv = ""
End Function

Sub SetPassword(path)
    Dim pass, input : pass = GetPassword_prv(path)

    If pass = "" Then
        pass = InputBox("No password found." & vbCrLf & "Please enter a password to use:", "Set Folder Password")
        SaveIni path, pass
    Else
        input = InputBox("Enter the current password to unlock folders:", "Set Folder Password")
        If input = admin Or input = pass Then
            input = InputBox("Enter the new password to unlock folders:", "Set Folder Password")
            SaveIni path, input
        Else
            MsgBox "Incorrect password!" & vbCrLf & "Current password remained.", vbCritical, "Set Folder Password"
        End If
    End If
End Sub

Sub SaveIni(path, pass)
    Dim iniPath, ini, inSection, foundSection, foundPass, content
    iniPath = path & "\Desktop.ini" : inSection = False
    foundSection = False : foundPass = False
    
    If fso.FileExists(iniPath) Then
        Set ini = fso.OpenTextFile(iniPath, 1)
        content = "" 

        Do While Not ini.AtEndOfStream
            line = Trim(ini.ReadLine)

            If LCase(line) = "[folderlock]" Then 
                inSection = True
                foundSection = True
            ElseIf LCase(line) = "[" Then
                inSection = False
            End If

            If inSection And LCase(Left(line, 9)) = "password=" Then
                line = "password=" & pass
                foundPass = True
                inSection = False
            End If

            content = content & line & vbCrLf
        Loop

        If Not (foundSection And foundPass) Then
            content = content & BuildIni(pass, False)
        End If

        ini.Close
    Else
        content = BuildIni(pass, True)
    End If

    shell.Run "cmd /c attrib -s -h """ & iniPath & """", 0, True
    Set ini = fso.OpenTextFile(iniPath, 2, True)
    ini.Write content : ini.Close
    shell.Run "cmd /c attrib +s +h """ & iniPath & """", 0, True
    
    MsgBox "Password saved", vbInformation, "Set Folder Password"
End Sub

Function BuildIni(pass, addIcon)
    Dim scriptPath, content : content = ""
    scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

    If addIcon Then content = _
        "[.ShellClassInfo]" & vbCrLf & _
        "IconResource=" & scriptPath & "\media\vault.ico" & vbCrLf

    BuildIni = content & vbCrLf & _
        "[FolderLock]" & vbCrLf & _
        "password=" & pass & vbCrLf
End Function

