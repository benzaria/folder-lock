' ================= HELPERS =================

' Alternative to MsgBox
Sub MsgPop(msg, style, title)
    shell.Popup msg, 1, title, style
End Sub

' Removes trailing slash/backslash from a folder path
Function CleanPath(path)
    If Right(path, 1) = "\" Or Right(path, 1) = "/" Then
        CleanPath = Left(path, Len(path) - 1)
    Else
        CleanPath = path
    End If
End Function

' Checks if a user has "no access" (denied) to a folder using icacls
Function GetAccess(path, user)
    Dim command, hasN
    
    ' Build the command to check ACLs and filter for (N) = deny
    command = "cmd /c icacls """ & path & """ | findstr /i /c:""" & user & """ /c:Eveyone | findstr /i ""(N)"""
    hasN = shell.Run(command, 0, True)

    ' 0 = no access
    If hasN = 0 Then
        GetAccess = True
    Else
        GetAccess = False
    End If
End Function

' Returns the folder password, prompting the user if none is set
Function GetPassword(path)
    Dim pass : pass = GetPassword_prv(path)
    
    ' If no password found, ask user and save it
    If pass = "" Then
        pass = InputBox("No password found." & vbCrLf & "Please enter a password to use:", "Set Folder Password")
        SaveIni path, pass
    End If

    GetPassword = pass
End Function

' Reads the password from Desktop.ini in the folder (private function)
Function GetPassword_prv(path)
    Dim iniPath, ini, inSection
    iniPath = path & "\Desktop.ini" : inSection = False

    ' If Desktop.ini does not exist, return empty
    If Not fso.FileExists(iniPath) Then
        GetPassword_prv = ""
        Exit Function
    End If

    Set ini = fso.OpenTextFile(iniPath, 1)

    ' Loop through each line
    Do While Not ini.AtEndOfStream
        line = Trim(ini.ReadLine)

        ' Detect [FolderLock] section
        If LCase(line) = "[folderlock]" Then 
            inSection = True
        ElseIf Left(line, 1) = "[" Then
            inSection = False
        End If

        ' If inside FolderLock section, look for "password="
        If inSection And LCase(Left(line, 9)) = "password=" Then
            GetPassword_prv = Mid(line, 10) ' Return password value
            ini.Close
            Exit Function
        End If
    Loop

    ini.Close
    GetPassword_prv = ""
End Function

' Sub to set a new password (prompts for current password if exists)
Sub SetPassword(path)
    Dim pass, input : pass = GetPassword_prv(path)

    ' No current password → ask for new one
    If pass = "" Then
        pass = InputBox("No password found." & vbCrLf & "Please enter a password to use:", "Set Folder Password")
        SaveIni path, pass
    Else
        ' Current password exists → verify before changing
        input = InputBox("Enter the current password to unlock folders:", "Set Folder Password")
        If input = admin Or input = pass Then
            input = InputBox("Enter the new password to unlock folders:", "Set Folder Password")
            SaveIni path, input
        Else
            MsgPop "Incorrect password!" & vbCrLf & "Current password remained.", vbCritical, "Set Folder Password"
        End If
    End If
End Sub

' Writes the password (and optionally icon) to Desktop.ini
Sub SaveIni(path, pass)
    Dim iniPath, ini, inSection, foundSection, foundPass, content
    iniPath = path & "\Desktop.ini" : inSection = False
    foundSection = False : foundPass = False
    
    ' If Desktop.ini exists, read & update it
    If fso.FileExists(iniPath) Then
        Set ini = fso.OpenTextFile(iniPath, 1)
        content = "" 

        ' Process each line
        Do While Not ini.AtEndOfStream
            line = Trim(ini.ReadLine)

            ' Detect FolderLock section
            If LCase(line) = "[folderlock]" Then 
                inSection = True
                foundSection = True
            ElseIf Left(line, 1) = "[" Then
                inSection = False
            End If

            ' Replace password line if found
            If inSection And LCase(Left(line, 9)) = "password=" Then
                line = "password=" & pass
                foundPass = True
                inSection = False
            End If

            content = content & line & vbCrLf
        Loop

        ' If section or password was missing, append it
        If Not (foundSection And foundPass) Then
            content = content & BuildIni(pass, False)
        End If

        ini.Close
    Else
        ' Desktop.ini does not exist → create full content
        content = BuildIni(pass, True)
    End If

    ' Ensure Desktop.ini is writable and hidden/system
    shell.Run "cmd /c attrib -s -h """ & iniPath & """", 0, True
    Set ini = fso.OpenTextFile(iniPath, 2, True)
    ini.Write content : ini.Close
    shell.Run "cmd /c attrib +s +h """ & iniPath & """", 0, True
    
    MsgPop "Password saved", vbInformation, "Set Folder Password"
End Sub

' Builds the content for Desktop.ini including optional icon path
Function BuildIni(pass, addIcon)
    Dim scriptPath, content : content = ""
    scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

    ' Add folder icon section if requested
    If addIcon Then content = _
        "[.ShellClassInfo]" & vbCrLf & _
        "IconResource=" & scriptPath & "\media\vault.ico" & vbCrLf

    ' Add FolderLock password section
    BuildIni = content & vbCrLf & _
        "[FolderLock]" & vbCrLf & _
        "password=" & pass & vbCrLf
End Function
