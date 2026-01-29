' ================== INIT ==================

Option Explicit
On Error Goto 0

' Create COM objects for shell operations and file system operations
Dim shell, fso, admin, user
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Get environment variables for admin password and current user
admin = shell.ExpandEnvironmentStrings("%ADMINPASSWORD%")
user = shell.ExpandEnvironmentStrings("%USERDOMAIN%\%USERNAME%")

ParseArgs(WScript.Arguments)

' ================== MAIN ==================

' Check if only one folder was provided
If UBound(folders) = 0 Then
    Dim folder : folder = folders(0)
    Dim isLocked : isLocked = GetAccess(folder) ' Check if folder is currently locked

    ' ---------- Folder is currently locked ----------
    If isLocked Then
        ' Trying to lock an already locked folder → notify and exit
        If lock Then
            MsgPop "Folder already locked.", vbQuestion, "Access Revoked"
            WScript.Quit
        End If

        ' User wants to change password → call SetPassword
        If changeLock Then
            SetPassword folder
        Else
            ' Prompt for password to unlock folder
            Dim input : input = InputBox("Enter the password to unlock folder:", "Unlock Protected Folder")
            Dim pass : pass = GetPassword(folder)  ' Get stored password

            ' If input matches admin override or folder password → unlock
            If input = admin Or input = pass Then
                UnlockFolder folder
                MsgPop "Folder unlocked.", vbInformation, "Access Granted"
            Else
                ' Wrong password → ensure folder stays locked
                LockFolder folder
                MsgPop "Incorrect password!" & vbCrLf & "Folder remain locked.", vbCritical, "Access Denied"
            End If
        End If

    ' ---------- Folder is currently unlocked ----------
    Else
        ' Trying to unlock an already unlocked folder → notify and exit
        If unlock Then
            MsgPop "Folder already unlocked.", vbQuestion, "Access Granted"
            WScript.Quit
        End If

        ' Change password on unlocked folder
        If changeLock Then
            SetPassword folder
        Else
            ' Ensure a password is set
            GetPassword folder
            ' Lock the folder
            LockFolder folder
            MsgPop "Folder locked.", vbInformation, "Access Revoked"
        End If
    End If

' ---------- Multiple folders ----------
Else
    ' If folder is unlocked and no unlock requested → we will lock
    If Not (GetAccess(folders(0)) Or unlock) Then lock = True

    ' ---------- Lock all folders ----------
    If lock Then
        LockFolders folders
        MsgPop "Folders locked.", vbInformation, "Access Revoked"
    ' ---------- Unlock all folders ----------
    Else
        ' Prompt for password to unlock all folders
        input = InputBox("Enter the password to unlock folders:", "Unlock Protected Folders")

        ' Admin override → unlock all
        If input = admin Then
            UnlockFolders folders
            MsgPop "Folders unlocked.", vbInformation, "Access Granted"
        Else
            ' Wrong password → keep all folders locked
            LockFolders folders
            MsgPop "Incorrect password!" & vbCrLf & "Folders remain locked.", vbCritical, "Access Denied"
        End If
    End If
End If
