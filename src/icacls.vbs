' ============== ICACLS ACTIONS ==============

Dim command

' Lock multiple folders
Sub LockFolders(paths, user)
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
Sub LockFolder(path, user)
    ' Build ICACLS command to deny full access (F) for the given user
    command = "icacls " & path & " /deny """ & user & ":(F)"""
    shell.Run command, 0, False
End Sub

' Unlock a single folder by resetting their permissions to default
Sub UnlockFolder(path)
    ' Reset all ACLs on the folder
    command = "icacls " & path & " /reset"
    shell.Run command, 0, False
End Sub
