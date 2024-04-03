On Error Resume Next

Const wdExportFormatPDF = 32
Set oPptApp = WScript.CreateObject("PowerPoint.Application")
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set fds = fso.GetFolder(".")
Set ffs = fds.Files

For Each ff In ffs
	If (LCase(Right(ff.Name, 4))=".ppt" Or Lcase(Right(ff.Name, 4))="pptx") And Left(ff.Name, 1)<>"~" Then

			Set oPpt = oPptApp.Presentations.Open(ff.Path, false, false, false)
			oPpt.Saveas Left(ff.Path, InStrRev(ff.Path, "."))&"pdf",wdExportFormatPDF,false

			If Err.Number Then
				MsgBox Err.Description
			End If
	End If
Next

oPpt.Close
oPptApp.Quit

Set oPpt=Nothing
Set oPptApp=Nothing

MsgBox "Completed"
