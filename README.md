<div align="center">

## Mp3AudioDll


</div>

### Description

This ActiveX Dll allows you to easily play Mp3 Audio using the Mci Api. You simply create an Audio object with a filename and control the playback through the objects interface. Gives you control over the speed, volume (both left and right) and so on... If enough of you are interested I will gladly expand and improve this project. Comments and suggestions would be greatly appreciated.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2004-12-16 13:36:30
**By**             |[Jason George](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jason-george.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Sound/MP3](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/sound-mp3__1-45.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Mp3AudioDl18299312162004\.zip](https://github.com/Planet-Source-Code/jason-george-mp3audiodll__1-57748/archive/master.zip)

### API Declarations

```
Private Declare Function CoCreateGuid Lib "ole32" (id As Any) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
```





