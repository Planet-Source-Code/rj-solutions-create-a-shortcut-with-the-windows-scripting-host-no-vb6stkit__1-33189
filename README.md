<div align="center">

## Create a shortcut with the Windows Scripting Host \(no VB6STKIT\)


</div>

### Description

I wanted code that used components that were mostly likely to be found on a user's machine. VB5STKIT, VB4STKIT, and VB6STKIT all could be used for creating shortcuts, but there is a good chance they weren't already on the user's machine, meaning I'd have to include it in my install package. The Windows Scripting Host is a default installation on Windows 98 and higher, and is likely on a Windows 95 machine. Plus you can include an object test to easily verify the user has the Scripting Host installed.
 
### More Info
 
Shortcut path (including the ".lnk" part of the filename), pop-up description, target path (the file to execute), command-line arguments, working directory, Hot Key to execute shortcut, Icon path, Window Style.

Arguments, Working Directory, hot key, icon location, and window style are all optional. This code is a simple adaptation of the Windows Scripting Documentation, which can be found on Microsoft's site. The default window style is to run in a normal window. Other values allow Minimized and Maximized operation. This can be used for URL shortcuts as well. This code verifies that the Scripting Host is installed, and if not, it returns False (no error produced).

Boolean (True or False) as to whether it was successful.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[RJ Solutions](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rj-solutions.md)
**Level**          |Beginner
**User Rating**    |4.7 (47 globes from 10 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rj-solutions-create-a-shortcut-with-the-windows-scripting-host-no-vb6stkit__1-33189/archive/master.zip)

### API Declarations

None!


### Source Code

```
Public Function CreateShortcut( _
 ByVal sShortcutPath As String, _
 ByVal sDescription As String, _
 ByVal sTargetPath As String, _
 Optional ByVal sArguments As String, _
 Optional ByVal sWorkingDirectory As String, _
 Optional ByVal sHotKey As String, _
 Optional ByVal sIconLocation As String, _
 Optional ByVal iWindowStyle As Integer = 3) As Boolean
'To get this to work for VB Script, Change the two lines below to: Dim sh, link.
Dim sh As Object
Dim link As Object
'Dynamically create the Script Object
Set sh = CreateObject("WScript.Shell")
'Check the path supplied and make sure the correct extension is on it.
If LCase(Right(sShortcutPath, 4)) = ".lnk" Or LCase(Right(sShortcutPath, 4)) = ".url" Then
Else
 sShortcutPath = sShortcutPath & ".lnk"
End If
'Check that the Scripting Host is installed by confirming that an object was truly created.
If IsObject(sh) Then
 Set link = sh.CreateShortcut(sShortcutPath)
 If IsObject(link) Then
  If IsMissing(sArguments) Then
  Else
   link.Arguments = sArguments
  End If
  link.Description = sDescription
  If IsMissing(sHotKey) Then
  Else
   link.HotKey = sHotKey
  End If
  If IsMissing(sIconLocation) Then
   sIconLocation = sTargetPath & ",1"
  End If
  link.IconLocation = sIconLocation
  link.TargetPath = sTargetPath
  link.WindowStyle = iWindowStyle
  If IsMissing(sWorkingDirectory) Then
   link.WorkingDirectory = sTargetPath
  Else
   link.WorkingDirectory = sWorkingDirectory
  End If
  'Now that the shortcut is fully created, you must save it.
  link.Save
  CreateShortcut = True
 End If
End If
End Function
```

