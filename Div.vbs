MsgBox "DivSecurity'e Hosgeldiniz! Masaustunde tarama yapiliyor...", vbInformation, "Bilgilendirme"

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(objShell.ExpandEnvironmentStrings("%USERPROFILE%\Desktop"))

' Silinen dosya sayacı
deletedCount = 0
deletedFiles = ""
fileExtensions = ""

' Masaustunde ve klasorlerdeki tum .SCR dosyalarini sil ve tum dosya uzantilarini listele
Sub ScanAndDelete(folder)
    For Each file In folder.Files
        ext = LCase(objFSO.GetExtensionName(file.Name))
        fileExtensions = fileExtensions & ext & vbCrLf

        If ext = "scr" Then
            deletedFiles = deletedFiles & file.Name & vbCrLf
            objFSO.DeleteFile file.Path, True
            deletedCount = deletedCount + 1
        End If
    Next

    For Each subFolder In folder.SubFolders
        ScanAndDelete subFolder
    Next
End Sub

ScanAndDelete objFolder

If deletedCount > 0 Then
    MsgBox "Toplam " & deletedCount & " adet .SCR dosyasi silindi!" & vbCrLf & _
           "Silinen dosyalar:" & vbCrLf & deletedFiles & vbCrLf & _
           "Masaustundeki tum dosya uzantilari:" & vbCrLf & fileExtensions, vbInformation, "DivSecurity"
Else
    MsgBox "Masaustunde silinecek .SCR dosyasi bulunamadi!" & vbCrLf & _
           "Masaustundeki tum dosya uzantilari:" & vbCrLf & fileExtensions, vbInformation, "DivSecurity"
End If