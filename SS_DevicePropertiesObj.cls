VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DevicePropertiesObj"
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

'private vars to hold property values
Private lFileName As String
Private lName As String
Private lAverageVolume As Long
Private lLeftVolume As Long
Private lRightVolume As Long
Private lLeftChannel As Long
Private lRightChannel As Long
Private lSpeed As Long
Private lPlayFrom As Long
Private lPlayTo As Long
Private lGuid As String
Private lLength As Long
Private lRepeat As Boolean

Private Sub Class_Initialize()
On Error Resume Next
  
  'set default values
  lFileName = ""
  lName = ""
  lAverageVolume = 1000
  lLeftVolume = 1000
  lRightVolume = 1000
  lLeftChannel = 0
  lRightChannel = 0
  lSpeed = 1000
  lPlayFrom = 0
  lPlayTo = 0
  lLength = 0
  lRepeat = False
  
End Sub

Public Property Let Repeat(ByVal bValue As Boolean)
On Error Resume Next

  lRepeat = bValue
  
End Property

Public Property Get Repeat() As Boolean
On Error Resume Next

  Repeat = lRepeat
  
End Property

Public Property Let Length(ByVal nLength As Long)
On Error Resume Next
  
  lLength = nLength

End Property

Public Property Get Length() As Long
On Error Resume Next
  
  Length = lLength

End Property

Public Property Let GUID(ByVal GUID As String)
On Error Resume Next

  lGuid = GUID
  
End Property

Public Property Get GUID() As String
On Error Resume Next

  GUID = lGuid
  
End Property

Public Property Let FileName(ByVal StrFile As String)
On Error Resume Next

  lFileName = StrFile
  
End Property

Public Property Get FileName() As String
On Error Resume Next

  FileName = lFileName
  
End Property

Public Property Let Name(ByVal StrName As String)
On Error Resume Next
  
  lName = StrName
  
End Property

Public Property Get Name() As String
On Error Resume Next
  
  Name = lName
  
End Property

Public Property Let AverageVolume(ByVal nVolume As Long)
On Error Resume Next

  lAverageVolume = nVolume
  
End Property

Public Property Get AverageVolume() As Long
On Error Resume Next

  AverageVolume = lAverageVolume
  
End Property

Public Property Let LeftVolume(ByVal nVolume As Long)
On Error Resume Next

  lLeftVolume = nVolume
  
End Property

Public Property Get LeftVolume() As Long
On Error Resume Next

  LeftVolume = lLeftVolume
  
End Property

Public Property Let RightVolume(ByVal nVolume As Long)
On Error Resume Next

  lRightVolume = nVolume
  
End Property

Public Property Get RightVolume() As Long
On Error Resume Next

  RightVolume = lRightVolume
  
End Property

Public Property Let LeftChannel(ByVal OnOff As SwitchEnum)
On Error Resume Next

  lLeftChannel = OnOff
  
End Property

Public Property Get LeftChannel() As SwitchEnum
On Error Resume Next

  LeftChannel = lLeftChannel
  
End Property

Public Property Let RightChannel(ByVal OnOff As SwitchEnum)
On Error Resume Next

  lRightChannel = OnOff
  
End Property

Public Property Get RightChannel() As SwitchEnum
On Error Resume Next

  RightChannel = lRightChannel
  
End Property

Public Property Let Speed(ByVal nSpeed As Long)
On Error Resume Next
  
  lSpeed = nSpeed
  
End Property

Public Property Get Speed() As Long
On Error Resume Next
  
  Speed = lSpeed
  
End Property

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
