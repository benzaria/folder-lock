' ============== ICACLS ACTIONS ==============

' Shorthand for shell.Run
Sub RunCmd(cmd, wait)
    shell.Run cmd, 0, wait
End Sub

' Lock multiple folders
Sub LockFolders(paths)
    For Each path In paths
        LockFolder path, user
    Next
End Sub

' Unlock multiple folders
Sub UnlockFolders(paths)
    For Each path In paths
        UnlockFolder path
    Next
End Sub

' Lock a single folder by denying full access to a specific user
Sub LockFolder(path)
    ' Build ICACLS command to deny full access (F) for the given user
    RunCmd "attrib +i +h """ & path & """", False
    RunCmd "icacls " & path & " /deny """ & user & ":F""", False
End Sub

' Unlock a single folder by resetting their permissions to default
Sub UnlockFolder(path)
    ' Reset all ACLs on the folder
    RunCmd "icacls " & path & " /reset", False
    RunCmd "attrib -i -h """ & path & """", False
End Sub

Sub BypassPerm(path, iniPath, content)
    ' Temporarily grant full access
    UnlockFolder path
    RunCmd "attrib -s -h """ & iniPath & """", True

    Set ini = fso.OpenTextFile(iniPath, 2, True)
    ini.Write content
    ini.Close

    ' Restore original protections
    RunCmd "attrib +s +h """ & iniPath & """", True
    user = "Everyone" : LockFolder path

    MsgPop "Password saved", vbInformation, "Set Folder Password"
End Sub
