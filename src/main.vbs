' ================= CONFIG =================

Option Explicit
On Error Goto 0

Dim shell, fso, admin, user
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

admin = shell.ExpandEnvironmentStrings("%ADMINPASSWORD%")
user = shell.ExpandEnvironmentStrings("%USERDOMAIN%\%USERNAME%")

' ================== MAIN ==================

If UBound(folders) = 0 Then
    Dim folder : folder = folders(0)
    Dim isLocked : isLocked = GetAccess(folder, user)

    If isLocked Then
        If lock Then
            MsgBox "Folder already locked.", vbQuestion, "Access Revoked"
            WScript.Quit
        End If

        If changeLock Then
            SetPassword folder
        Else
            Dim input : input = InputBox("Enter the password to unlock folder:", "Unlock Protected Folder")
            Dim pass : pass = GetPassword(folder)

            If input = admin Or input = pass Then
                UnlockFolder folder
                MsgBox "Folder unlocked.", vbInformation, "Aceess Granted"
            Else
                LockFolder folder, user
                MsgBox "Incorrect password!" & vbCrLf & "Folder remain locked.", vbCritical, "Access Denied"
            End If
        End If
        
    Else
        If unlock Then
            MsgBox "Folder already unlocked.", vbQuestion, "Access Granted"
            WScript.Quit
        End If

        If changeLock Then
            SetPassword folder
        Else
            GetPassword folder
            LockFolder folder, user
            MsgBox "Folder locked.", vbInformation, "Aceess Revoked"
        End If
    End If

Else
    folder = folders(0)
    isLocked = GetAccess(folder, user)

    If Not (isLocked Or unlock) Then lock = True

    If lock Then
        LockFolders folders, user
        MsgBox "Folders locked.", vbInformation, "Aceess Revoked"
    Else
        input = InputBox("Enter the password to unlock folders:", "Unlock Protected Folders")

        If input = admin Then
            UnlockFolders folders
            MsgBox "Folders unlocked.", vbInformation, "Aceess Granted"
        Else
            LockFolders folders, user
            MsgBox "Incorrect password!" & vbCrLf & "Folders remain locked.", vbCritical, "Access Denied"
        End If
    End If
End If
