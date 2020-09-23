VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mp3AudioCls"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************
'*Created By: Jason George                   *
'*Date: Dec. 16 2004                         *
'*Email: mescalito@adelphia.net              *
'*********************************************

Option Explicit

Public Enum SwitchEnum
  TurnOff = 0
  TurnOn = 1
End Enum

Public Property Let DisplayError(ByVal bValue As Boolean) 'sets option to display Mci Errors
On Error Resume Next
  
  pDisplayError = bValue
  
End Property

Public Property Get DisplayError() As Boolean 'gets option to display Mci Errors
On Error Resume Next
  
  DisplayError = pDisplayError
  
End Property

'returns a device properties object from a open device
Public Function GetPropertiesFromDevice(ByVal AudioDevice As AudioObj) As DevicePropertiesObj
On Error GoTo ErrHnd
  
  Dim Props As New DevicePropertiesObj          'create a new DeviceProperty object
  
    With AudioDevice                            'set the new DeviceProperty objects values
      Props.AverageVolume = .Volume             'to the properties of the audiodevice object
      Props.FileName = .FileName                'passed to the function
      Props.GUID = .GUID
      Props.LeftChannel = .LeftChannel
      Props.LeftVolume = .LeftVolume
      Props.Length = .Length
      Props.Name = .Name
      Props.PlayFrom = .PlayFrom
      Props.PlayTo = .PlayTo
      Props.Repeat = .Repeat
      Props.RightChannel = .RightChannel
      Props.RightVolume = .RightVolume
      Props.Speed = .Speed
    End With
    
  Set GetPropertiesFromDevice = Props         'set the function to new DeviceProperty object
  Set Props = Nothing                         'destroy the new DeviceProperty object
  
Exit Function
ErrHnd:
  If pDisplayError = True Then
    MsgBox "GetPropertiesFromDevice Failed.", , "Error"
  End If
  
  GetPropertiesFromDevice = Nothing
  
End Function

'creates a new AudioObj with a given filename
Public Function CreateAudioDevice(ByVal FileName As String) As AudioObj
On Error GoTo ErrHnd
  
  Dim RetVal As Boolean
  Dim NewAudioDevice As New AudioObj                      'create a new AudioObj
    
    RetVal = NewAudioDevice.OpenFile(FileName)            'open the AudioObj with the given filename
      If RetVal = False Then                              'error
        If pDisplayError = True Then                      'if error display is on
          MsgBox "CreateAudioDevice Failed.", , "Error"   'display the error
        End If
      End If
    
    Set CreateAudioDevice = NewAudioDevice                'set the function to the NewAudioDevice
    Set NewAudioDevice = Nothing                          'destroy the new AudioDevice
    
Exit Function
ErrHnd:
  CreateAudioDevice = Nothing
  
End Function

'creates a new AudioObj with a given DevicePropertiesObj
Public Function CreateAudioDeviceEx(ByVal Props As DevicePropertiesObj) As AudioObj
On Error GoTo ErrHnd
  
  Dim RetVal As Boolean
  Dim NewAudioDevice As New AudioObj                        'create a new AudioObj
    
    RetVal = NewAudioDevice.OpenFile(Props.FileName)        'open the AudioObj with the given filename
      If RetVal = False Then
        If pDisplayError = True Then
          MsgBox "CreateAudioDeviceEx Failed.", , "Error"
        End If
      Else
        With NewAudioDevice                                 'set the new device properties to the
          Props.FileName = .FileName                        'DevicePropertiesObj passed to the function
          Props.GUID = .GUID
          Props.Length = .Length
          .Name = Props.Name
          .PlayFrom = Props.PlayFrom
          .PlayTo = Props.PlayTo
          .Repeat = Props.Repeat
          .LeftChannel = Props.LeftChannel
          .RightChannel = Props.RightChannel
          .LeftVolume = Props.LeftVolume
          .RightVolume = Props.RightVolume
          .Speed = Props.Speed
          .Volume = Props.AverageVolume
        End With
      End If
      
    Set CreateAudioDeviceEx = NewAudioDevice                'set the function to the NewAudioDevice
    Set NewAudioDevice = Nothing                            'destroy the new AudioDevice
    
Exit Function
ErrHnd:
  CreateAudioDeviceEx = Nothing
  
End Function

