VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudioObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************
'*Created By: Jason George                   *
'*Date: Dec. 16 2004                         *
'*Email: mescalito@adelphia.net              *
'*********************************************

Option Explicit

Private lFileName As String     'holds file name property
Private lGuid As String         'holds unique id for each object created
Private lName As String         'so we can describe the object
Private lPlayFrom As Long       'used for custom start position
Private lPlayTo As Long         'used for custom end position
Private lRepeat As Boolean      'holds repeat property

Private Sub Class_Terminate()
On Error Resume Next
  
  CloseDevice                   'close the device when the class Terminates
  
End Sub

Public Property Let PlayFrom(ByVal nPos As Long)
On Error Resume Next

  lPlayFrom = nPos
  
End Property

Public Property Get PlayFrom() As Long
On Error Resume Next

  PlayFrom = lPlayFrom
  
End Property

Public Property Let PlayTo(ByVal nPos As Long)
On Error Resume Next

  lPlayTo = nPos
  
End Property

Public Property Get PlayTo() As Long
On Error Resume Next

  PlayTo = lPlayTo
  
End Property

Public Property Let Name(ByVal StrName As String)
On Error Resume Next

  lName = StrName
  
End Property

Public Property Get Name() As String
On Error Resume Next

  Name = lName
  
End Property

Public Property Let Repeat(ByVal bValue As Boolean)
On Error Resume Next

  lRepeat = bValue
  
End Property

Public Property Get Repeat() As Boolean
On Error Resume Next

  Repeat = lRepeat
  
End Property

Public Property Get GUID() As String
On Error Resume Next

  GUID = lGuid
  
End Property

Public Property Let FileName(ByVal Path As String)
On Error Resume Next

  lFileName = Path
  OpenFile lFileName
  
End Property

Public Property Get FileName() As String
On Error Resume Next

  FileName = lFileName
  
End Property

Public Property Get Length() As Long          'returns the length in milliseconds
On Error GoTo ErrHnd

  Dim RetVal As Long
  Dim Temp As String * 128
  
    RetVal = mciSendString("status " & lGuid & " length", Temp, 128, 0)
      If RetVal <> 0 Then GoTo ErrHnd
      
    Length = CLng(Temp)
    
Exit Property
ErrHnd:
  Length = 0
  DisplayMciError RetVal, "Length"
  
End Property

Private Sub InitDevice()
On Error Resume Next

  Dim Temp As String
  Dim RetVal As Long
    
    'tell the device to return time and position information in milliseconds
    Temp = "set " & lGuid & " time format milliseconds"
    RetVal = mciSendString(Temp, 0&, 0, 0)
      If RetVal <> 0 Then GoTo ErrHnd
      
Exit Sub
ErrHnd:
  DisplayMciError RetVal, "InitDevice"
  
End Sub

Public Function OpenFile(ByVal Path As String) As Boolean
On Error GoTo ErrHnd
  
  Dim Temp As String * 128
  Dim RetVal As Long
  Dim Fl As Boolean
    
    If lGuid <> "" Then         'if Guid isn't null then(a device is open)
      CloseDevice               'close this device
    Else
      lFileName = Path          'set the local file name property
    End If
    
    lGuid = SS_CreateGUID()     'grab a new Guid
    
    'get the short path name and quote it
    Path = Chr(34) & SS_GetShortPath(Path) & Chr(34)
    'open device with file(path), use MPEGVideo Codec and name the device with the Guid
    Temp = "open " & Path & " type " & DevType & " Alias " & lGuid
    RetVal = mciSendString(Temp, 0&, 0, 0)
      If RetVal <> 0 Then GoTo ErrHnd
      
    InitDevice
    
  OpenFile = True                     'opened device without error
Exit Function
ErrHnd:
  OpenFile = False                    'opened device without error
  DisplayMciError RetVal, "OpenFile"  'send error code(RetVal) with the function name
  
End Function

Public Function Play(Optional ByVal nPlayFrom As Long = 0, Optional ByVal nPlayTo As Long) As Boolean
On Error GoTo ErrHnd
  
  Dim Temp As String
  Dim RetVal As Long
    
    If lPlayFrom = 0 Then                         'if the PlayFrom property is not set
      If nPlayFrom > nPlayTo Then nPlayFrom = 0   'and if start pos > end pos then set it to 0
    Else                                          'if the PlayFrom property is set
      nPlayFrom = lPlayFrom                       'lPlayFrom overrides play function nPlayFrom position
    End If
    
    If lPlayTo = 0 Then                                             'if the PlayTo property is not set
      If nPlayTo = 0 Or nPlayTo < PlayFrom Then nPlayTo = Me.Length 'set nplayto to length if properties are invalid
    Else                                                            'if the PlayTo property is set
      nPlayTo = lPlayTo                                             'lPlayTo overrides play function nPlayTo position
    End If
                                                                    'if lRepeat is true append " repeat" to the send string
    Temp = "play " & lGuid & " from " & nPlayFrom & " to " & nPlayTo & IIf(lRepeat, " repeat", "")
    RetVal = mciSendString(Temp, 0&, 0, 0)
      If RetVal <> 0 Then GoTo ErrHnd
      
  Play = True
Exit Function
ErrHnd:
  Play = False
  DisplayMciError RetVal, "Play"
  
End Function

Public Function Pause() As Boolean
On Error GoTo ErrHnd
  
  Dim RetVal As Long
  
    RetVal = mciSendString("pause " & lGuid, 0&, 0, 0)          'pause playing
      If RetVal <> 0 Then GoTo ErrHnd
      
  Pause = True
Exit Function
ErrHnd:
  Pause = False
  DisplayMciError RetVal, "Pause"
  
End Function

Public Property Get LeftVolume() As Long
On Error GoTo ErrHnd
  
  Dim RetVal As Long
  Dim Temp As String * 128
  
    RetVal = mciSendString("status " & lGuid & " left volume", Temp, 128, 0)          'Get left channel volume
      If RetVal <> 0 Then GoTo ErrHnd
      
    LeftVolume = CLng(Temp)
    
Exit Property
ErrHnd:
  LeftVolume = -1
  DisplayMciError RetVal, "LeftVolume"
  
End Property

Public Property Let LeftVolume(ByVal nVolume As Long)
On Error GoTo ErrHnd
  
  Dim RetVal As Long
    
    RetVal = mciSendString("setaudio " & lGuid & " left volume to " & nVolume, 0&, 0, 0)  'set left channel volume
      If RetVal <> 0 Then GoTo ErrHnd
      
Exit Property
ErrHnd:
  DisplayMciError RetVal, "LeftVolume"
  
End Property

Public Property Get RightVolume() As Long
On Error GoTo ErrHnd
  
  Dim RetVal As Long
  Dim Temp As String * 128
  
    RetVal = mciSendString("status " & lGuid & " right volume", Temp, 128, 0) 'Get right channel volume
      If RetVal <> 0 Then GoTo ErrHnd
      
    RightVolume = CLng(Temp)
    
Exit Property
ErrHnd:
  RightVolume = -1
  DisplayMciError RetVal, "RightVolume"
  
End Property

Public Property Let RightVolume(ByVal nVolume As Long)
On Error GoTo ErrHnd
  
  Dim RetVal As Long
  
    RetVal = mciSendString("setaudio " & lGuid & " right volume to " & nVolume, 0&, 0, 0) 'set right channel volume
      If RetVal <> 0 Then GoTo ErrHnd
      
Exit Property
ErrHnd:
  DisplayMciError RetVal, "RightVolume"
  
End Property


Public Property Get Volume() As Long
On Error GoTo ErrHnd
  
  Dim RetVal As Long
  Dim Temp As String * 128
  
    RetVal = mciSendString("status " & lGuid & " volume", Temp, 128, 0)         'get the average volume for both channels
      If RetVal <> 0 Then GoTo ErrHnd
      
    Volume = CLng(Temp)
    
Exit Property
ErrHnd:
  Volume = -1
  DisplayMciError RetVal, "Volume"
  
End Property

Public Property Let Volume(ByVal nVolume As Long)
On Error GoTo ErrHnd
  
  Dim RetVal As Long
    
    RetVal = mciSendString("setaudio " & lGuid & " volume to " & nVolume, 0&, 0, 0) 'set the average volume for both channels
      If RetVal <> 0 Then GoTo ErrHnd
      
Exit Property
ErrHnd:
  DisplayMciError RetVal, "Volume"
  
End Property

Public Property Let Speed(ByVal nSpeed As Long)
On Error GoTo ErrHnd
  
  Dim RetVal As Long
  Dim Temp As String
    
    RetVal = mciSendString("set " & lGuid & " speed " & nSpeed, 0&, 0, 0)       'set the playback speed
      If RetVal <> 0 Then GoTo ErrHnd
      
Exit Property
ErrHnd:
  DisplayMciError RetVal, "Speed"
  
End Property

Public Property Get Speed() As Long
On Error GoTo ErrHnd
  
  Dim RetVal As Long
  Dim Temp As String * 128
    
    RetVal = mciSendString("status " & lGuid & " speed", Temp, 128, 0)          'get the playback speed
      If RetVal <> 0 Then GoTo ErrHnd
      
    Speed = CLng(Temp)
    
Exit Property
ErrHnd:
  DisplayMciError RetVal, "Speed"
  
End Property

Public Property Get LeftChannel() As SwitchEnum         'get the left channel on/off state
On Error GoTo ErrHnd
  
  Dim RetVal As Long
  Dim Temp As String * 128
  
    RetVal = mciSendString("status " & lGuid & " audio left", Temp, 128, 0)
      LeftChannel = IIf(Trim(Temp) = "on", TurnOn, TurnOff)
    
Exit Property
ErrHnd:
  LeftChannel = -1
  DisplayMciError RetVal, "LeftChannel"
  
End Property

Public Property Let LeftChannel(ByVal OnOff As SwitchEnum)
On Error GoTo ErrHnd
  
  Dim RetVal As Long
  Dim Temp As String
    
    If OnOff = TurnOn Then
      Temp = "set " & lGuid & " audio left on"          'turn left channel on
    ElseIf OnOff = TurnOff Then
      Temp = "set " & lGuid & " audio left off"         'turn left channel off
    End If
    
    RetVal = mciSendString(Temp, 0&, 0, 0)
      If RetVal <> 0 Then GoTo ErrHnd
      
Exit Property
ErrHnd:
  
  DisplayMciError RetVal, "LeftChannel"
  
End Property

Public Property Get RightChannel() As SwitchEnum          'get the right channel on/off state
On Error GoTo ErrHnd
  
  Dim RetVal As Long
  Dim Temp As String * 128
  
    RetVal = mciSendString("status " & lGuid & " audio right", Temp, 128, 0)
      RightChannel = IIf(Trim(Temp) = "on", 1, 0)
    
Exit Property
ErrHnd:
  RightChannel = -1
  DisplayMciError RetVal, "RightChannel"
  
End Property

Public Property Let RightChannel(ByVal OnOff As SwitchEnum)
On Error GoTo ErrHnd
  
  Dim RetVal As Long
  Dim Temp As String
  
    If OnOff = TurnOn Then
      Temp = "set " & lGuid & " audio right on"           'turn right channel on
    ElseIf OnOff = TurnOff Then
      Temp = "set " & lGuid & " audio right off"          'turn right channel off
    End If
    
    RetVal = mciSendString(Temp, 0&, 0, 0)
      If RetVal <> 0 Then GoTo ErrHnd
    
Exit Property
ErrHnd:
  DisplayMciError RetVal, "RightChannel"
  
End Property

Public Function SetLeftRightChannel(Optional ByVal OnOff As SwitchEnum = TurnOn) As Boolean
On Error GoTo ErrHnd
  
  Dim RetVal As Long
  Dim Temp As String
  
    If OnOff = TurnOn Then
      Temp = "set " & lGuid & " audio all on"           'turn left & right channels on
    ElseIf OnOff = TurnOff Then
      Temp = "set " & lGuid & " audio all off"          'turn left & right channels off
    End If
    
    RetVal = mciSendString(Temp, 0&, 0, 0)
      If RetVal <> 0 Then GoTo ErrHnd
      
  SetLeftRightChannel = True
Exit Function
ErrHnd:
  SetLeftRightChannel = False
  DisplayMciError RetVal, "SetLeftRightChannel"
  
End Function

Public Function ResumePlaying() As Boolean
On Error GoTo ErrHnd
  
  Dim RetVal As Long
  
    RetVal = mciSendString("Resume " & lGuid, 0&, 0, 0)         'resume playback during a paused state
      If RetVal <> 0 Then GoTo ErrHnd
      
  ResumePlaying = True
Exit Function
ErrHnd:
  ResumePlaying = False
  DisplayMciError RetVal, "ResumePlaying"
  
End Function

Public Function StopPlaying() As Boolean
On Error GoTo ErrHnd
  
  Dim RetVal As Long
  
    RetVal = mciSendString("stop " & lGuid, 0&, 0, 0)         'stop playing audio
      If RetVal <> 0 Then GoTo ErrHnd
      
  StopPlaying = True
Exit Function
ErrHnd:
  StopPlaying = False
  DisplayMciError RetVal, "StopPlaying"
  
End Function


Public Function CloseDevice() As Boolean
On Error GoTo ErrHnd
  
  Dim RetVal As Long
      
    If Len(lGuid) = 0 Then Exit Function                'if no Guid the device is not open so exit function
    
    RetVal = mciSendString("close " & lGuid, 0&, 0, 0)  'close the open device
      If RetVal <> 0 Then GoTo ErrHnd                   'on error report it
    
    lGuid = ""                                          'reset the default properties
    lFileName = ""
    lName = ""
    lPlayFrom = 0
    lPlayTo = 0
    lRepeat = False
    
  CloseDevice = True
Exit Function
ErrHnd:
  CloseDevice = False
  DisplayMciError RetVal, "CloseDevice"
  
End Function

Public Function IsPlaying() As Boolean  'returns true if playing otherwise returns false
On Error GoTo ErrHnd
  
  If Mode = "playing" Then
    IsPlaying = True
  Else
    IsPlaying = False
  End If
  
Exit Function
ErrHnd:
  IsPlaying = False
  
End Function

Public Function IsStopped() As Boolean  'returns true if stopped otherwise returns false
On Error GoTo ErrHnd
  
  If Mode = "stopped" Then
    IsStopped = True
  Else
    IsStopped = False
  End If
  
Exit Function
ErrHnd:
  IsStopped = False
  
End Function

Public Function IsPaused() As Boolean 'returns true if paused otherwise returns false
On Error GoTo ErrHnd
  
  If Mode = "paused" Then
    IsPaused = True
  Else
    IsPaused = False
  End If
  
  
Exit Function
ErrHnd:
  IsPaused = False
  
End Function

Public Function Position() As Long  'returns the current position during playback
On Error GoTo ErrHnd

  Dim RetVal As Long
  Dim Temp As String * 128
  
    RetVal = mciSendString("status " & lGuid & " position", Temp, 128, 0)
      If RetVal <> 0 Then GoTo ErrHnd
      
    Position = CLng(Temp)
    
Exit Function
ErrHnd:
  Position = 0
  DisplayMciError RetVal, "Position"
  
End Function

Public Function Mode() As String  'returns the current device Mode(playing,paused,stopped...)
On Error GoTo ErrHnd

  Dim RetVal As Long
  Dim Temp As String * 128
  
    RetVal = mciSendString("status " & lGuid & " mode", Temp, 128, 0)
      If RetVal <> 0 Then GoTo ErrHnd
      
    Mode = Trim(Temp)
    
Exit Function
ErrHnd:
  Mode = 0
  DisplayMciError RetVal, "Mode"
  
End Function


