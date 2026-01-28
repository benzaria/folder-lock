; Folder Lock Hotkey

scriptFolder := A_ScriptDir
lockScript := scriptFolder . "\lock.wsf"

^l::
WinGetClass, class, A

if (class != "CabinetWClass" and class != "ExploreWClass")
    return

for window in ComObjCreate("Shell.Application").Windows
{
    if (window.hwnd = WinActive("A"))
    {
        selected := window.Document.SelectedItems
        count := selected.Count
        Loop %count%
        {
            item := selected.Item(A_Index-1)
            folderPath := item.Path
            Run, wscript.exe "%lockScript%" -e "%folderPath%"
        }
    }
}

return
