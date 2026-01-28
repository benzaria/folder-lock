' ============== ACACLS ACTIONS ==============

Sub LockFolders(paths, user)
    Dim command
    For Each path In paths
        command = "icacls " & path & " /deny """ & user & ":(F)"""
        shell.Run command, 0, False
    Next
End Sub

Sub UnlockFolders(paths)
    Dim command
    For Each path In paths
        command = "icacls " & path & " /reset"
        shell.Run command, 0, False
    Next
End Sub

Sub LockFolder(path, user)
    Dim paths(0) : paths(0) = path
    LockFolders paths, user
End Sub

Sub UnlockFolder(path)
    Dim paths(0) : paths(0) = path
    UnlockFolders paths
End Sub
