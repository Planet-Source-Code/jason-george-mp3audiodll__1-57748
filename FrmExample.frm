VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FrmExample 
   Caption         =   "Play Mp3s with an ActiveX Dll"
   ClientHeight    =   3525
   ClientLeft      =   3735
   ClientTop       =   1740
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   7245
   Begin VB.Frame Frame2 
      Caption         =   "Audio 2"
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   7215
      Begin MSComctlLib.Slider Slider4 
         Height          =   495
         Left            =   4680
         TabIndex        =   11
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   393216
         Max             =   2000
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   495
         Left            =   1800
         TabIndex        =   9
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         Max             =   1000
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Right Channel"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Left Channel"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Speed"
         Height          =   195
         Left            =   4800
         TabIndex        =   13
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Volume"
         Height          =   195
         Left            =   1920
         TabIndex        =   8
         Top             =   360
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Audio 1"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin MSComctlLib.Slider Slider3 
         Height          =   495
         Left            =   4680
         TabIndex        =   10
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   393216
         Max             =   2000
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         Max             =   1000
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Right Channel"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Left Channel"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Speed"
         Height          =   195
         Left            =   4800
         TabIndex        =   12
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Volume:"
         Height          =   195
         Left            =   1920
         TabIndex        =   7
         Top             =   360
         Width           =   570
      End
   End
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   7200
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu MnuAudio1 
         Caption         =   "Audio 1"
         Begin VB.Menu MnuOpen1 
            Caption         =   "Open..."
         End
         Begin VB.Menu MnuClose1 
            Caption         =   "Close"
         End
         Begin VB.Menu MnuSep2 
            Caption         =   "-"
         End
         Begin VB.Menu MnuPlay1 
            Caption         =   "Play"
         End
         Begin VB.Menu MnuPause1 
            Caption         =   "Pause"
         End
         Begin VB.Menu MnuStop1 
            Caption         =   "Stop"
         End
         Begin VB.Menu MnuSep4 
            Caption         =   "-"
         End
         Begin VB.Menu MnuProperties1 
            Caption         =   "Properties"
         End
      End
      Begin VB.Menu MnuAudio2 
         Caption         =   "Audio 2"
         Begin VB.Menu Open2 
            Caption         =   "Open..."
         End
         Begin VB.Menu MnuClose2 
            Caption         =   "Close"
         End
         Begin VB.Menu MnuSep3 
            Caption         =   "-"
         End
         Begin VB.Menu MnuPlay2 
            Caption         =   "Play"
         End
         Begin VB.Menu MnuPause2 
            Caption         =   "Pause"
         End
         Begin VB.Menu MnuStop2 
            Caption         =   "Stop"
         End
         Begin VB.Menu MnuSep5 
            Caption         =   "-"
         End
         Begin VB.Menu MnuProperties2 
            Caption         =   "Properties"
         End
      End
      Begin VB.Menu MnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FrmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************
'*Created By: Jason George                   *
'*Date: Dec. 16 2004                         *
'*Email: mescalito@adelphia.net              *
'*********************************************

Option Explicit

Private Sub Check1_Click()
On Error Resume Next

  Device1.LeftChannel = Check1.Value    'set left channel off or on
  
End Sub

Private Sub Check2_Click()
On Error Resume Next
  
  Device1.RightChannel = Check2.Value   'set right channel off or on
  
End Sub

Private Sub Check3_Click()
On Error Resume Next

  Device2.LeftChannel = Check3.Value    'set left channel off or on

End Sub

Private Sub Check4_Click()
On Error Resume Next
  
  Device2.RightChannel = Check4.Value   'set right channel off or on

End Sub

Private Sub Form_Load()
On Error Resume Next

  mp3audio.DisplayError = True          'tells the mp3audio.dll to display Mci Errors
  UpdateObjects                         'update the frame controls and menus
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  
                                        'remove objects from memory
  Set Device1 = Nothing                 'destroy device1
  Set Device2 = Nothing                 'destroy device2
  Set AudioProps1 = Nothing             'destroy AudioProps1
  Set AudioProps2 = Nothing             'destroy AudioProps2
  
End Sub

Private Sub MnuClose1_Click()
On Error Resume Next

  Device1.CloseDevice                   'close the device
  mp3audio.DisplayError = False         'turn off error display while closing the device
  UpdateObjects                         'update the frame controls and menus
    
End Sub

Private Sub MnuClose2_Click()
On Error Resume Next

  Device2.CloseDevice                   'close the device
  mp3audio.DisplayError = False         'turn off error display while closing the device
  UpdateObjects                         'update the frame controls and menus

End Sub

Private Sub MnuExit_Click()
On Error Resume Next
  
  Unload Me                             'unload this form from memory
  
End Sub

Private Sub MnuOpen1_Click()
On Error Resume Next
  
  With Cdlg
    .Filter = "Mp3 (*.mp3)|*.mp3"
    .ShowOpen                           'open a dialog to search for Mp3 files
    If Len(.FileName) <> 0 Then         'if a valid file name exist
      lFileName1 = .FileName            'set the global file name Var
      'create device method 1
      Set Device1 = mp3audio.CreateAudioDevice(lFileName1)
      'update the frame controls and menus
      UpdateObjects
    End If
  End With

End Sub

Private Sub MnuPause1_Click()
On Error Resume Next

  Device1.Pause
  MnuPause1.Enabled = False
  MnuPlay1.Enabled = True
  
End Sub

Private Sub MnuPause2_Click()
On Error Resume Next

  Device2.Pause
  MnuPause2.Enabled = False
  MnuPlay2.Enabled = True
  
End Sub

Private Sub MnuPlay1_Click()
On Error Resume Next
  
  If Device1.IsPaused Then              'if device is paused
    Device1.ResumePlaying               'resume playback
  Else
    Device1.Play                        'start playing
  End If
  
  MnuPlay1.Enabled = False
  MnuPause1.Enabled = True
  MnuStop1.Enabled = True
  
End Sub

Private Sub MnuPlay2_Click()
On Error Resume Next
  
  If Device2.IsPaused = True Then       'if device is paused
    Device2.ResumePlaying               'resume playback
  Else
    Device2.Play                        'start playing
  End If
  
  MnuPlay2.Enabled = False
  MnuPause2.Enabled = True
  MnuStop2.Enabled = True
  
End Sub

Private Sub MnuProperties1_Click()
On Error Resume Next

  Dim StrProps As String
  
    Set AudioProps1 = mp3audio.GetPropertiesFromDevice(Device1)
    
    With AudioProps1
      StrProps = "AverageVolume: " & .AverageVolume & vbCrLf
      StrProps = StrProps & "FileName: " & .FileName & vbCrLf
      StrProps = StrProps & "Guid: " & .Guid & vbCrLf
      StrProps = StrProps & "LeftChannel: " & .LeftChannel & vbCrLf
      StrProps = StrProps & "RightChannel: " & .RightChannel & vbCrLf
      StrProps = StrProps & "LeftVolume: " & .LeftVolume & vbCrLf
      StrProps = StrProps & "RightVolume: " & .RightVolume & vbCrLf
      StrProps = StrProps & "Length: " & .Length & vbCrLf
      StrProps = StrProps & "Name: " & .Name & vbCrLf
      StrProps = StrProps & "PlayFrom: " & .PlayFrom & vbCrLf
      StrProps = StrProps & "PlayTo: " & .PlayTo & vbCrLf
      StrProps = StrProps & "Repeat: " & .Repeat & vbCrLf
      StrProps = StrProps & "Speed: " & .Speed & vbCrLf
      StrProps = StrProps & "Mode: " & Device1.Mode
    End With
    
    MsgBox StrProps, vbInformation, "Device 1 Proprties"
    
End Sub

Private Sub MnuProperties2_Click()
On Error Resume Next

  Dim StrProps As String
  
    Set AudioProps2 = mp3audio.GetPropertiesFromDevice(Device2)
    
    With AudioProps2
      StrProps = "AverageVolume: " & .AverageVolume & vbCrLf
      StrProps = StrProps & "FileName: " & .FileName & vbCrLf
      StrProps = StrProps & "Guid: " & .Guid & vbCrLf
      StrProps = StrProps & "LeftChannel: " & .LeftChannel & vbCrLf
      StrProps = StrProps & "RightChannel: " & .RightChannel & vbCrLf
      StrProps = StrProps & "LeftVolume: " & .LeftVolume & vbCrLf
      StrProps = StrProps & "RightVolume: " & .RightVolume & vbCrLf
      StrProps = StrProps & "Length: " & .Length & vbCrLf
      StrProps = StrProps & "Name: " & .Name & vbCrLf
      StrProps = StrProps & "PlayFrom: " & .PlayFrom & vbCrLf
      StrProps = StrProps & "PlayTo: " & .PlayTo & vbCrLf
      StrProps = StrProps & "Repeat: " & .Repeat & vbCrLf
      StrProps = StrProps & "Speed: " & .Speed & vbCrLf
      StrProps = StrProps & "Mode: " & Device2.Mode
    End With
    
    MsgBox StrProps, vbInformation, "Device 2 Proprties"

End Sub

Private Sub MnuStop1_Click()
On Error Resume Next

  Device1.StopPlaying                   'stop playing
  MnuStop1.Enabled = False
  MnuPlay1.Enabled = True
  MnuPause1.Enabled = False
  
End Sub

Private Sub MnuStop2_Click()
On Error Resume Next

  Device2.StopPlaying                   'stop playing
  MnuStop2.Enabled = False
  MnuPlay2.Enabled = True
  MnuPause2.Enabled = False
  
End Sub

Private Sub Open2_Click()
On Error Resume Next
  
  With Cdlg
    .Filter = "Mp3 (*.mp3)|*.mp3"
    .ShowOpen                           'open a dialog to search for Mp3 files
    If Len(.FileName) <> 0 Then         'if a valid file name exist
      lFileName2 = .FileName            'set the global file name Var
      'create device method 2
      With AudioProps2                  'set some properties for the device first
        .AverageVolume = 750
        .FileName = lFileName2          'filename must be set
        .Name = "Audio Device Two"
        .Speed = 750
        .LeftChannel = TurnOn
        .RightChannel = TurnOn
      End With
      'create device with the above properties
      Set Device2 = mp3audio.CreateAudioDeviceEx(AudioProps2)
      'update the frame controls and menus
      UpdateObjects
    End If
  End With

End Sub

'updates the frame controls and menus
Public Sub UpdateObjects()
On Error Resume Next

  With Me
    If Not (Device1 Is Nothing) Then          'if device1 is a valid object-has been set
      If Len(Device1.FileName) <> 0 Then      'if it is-does it have a valid filename
        Frame1.Caption = "File Name: " & Device1.FileName
        .Frame1.Enabled = True
        .MnuClose1.Enabled = True
        .MnuPlay1.Enabled = True
        .MnuStop1.Enabled = False
        .MnuPause1.Enabled = False
        .MnuProperties1.Enabled = True
        .Check1.Value = Not Device1.LeftChannel
        .Check2.Value = Not Device1.RightChannel
        .Slider1.Value = Device1.Volume
        .Slider3.Value = Device1.Speed
      Else                                    'does not have a valid filename
        Frame1.Caption = "File Name: "
        .Frame1.Enabled = False
        .MnuClose1.Enabled = False
        .MnuPlay1.Enabled = False
        .MnuStop1.Enabled = False
        .MnuPause1.Enabled = False
        .MnuProperties1.Enabled = False
        .Check1.Value = Not Device1.LeftChannel
        .Check2.Value = Not Device1.RightChannel
        .Slider1.Value = Device1.Volume
        .Slider3.Value = Device1.Speed
      End If
    Else                                      'device1 is not valid-has not been set
      Frame1.Caption = "File Name: "
      .Frame1.Enabled = False
      .MnuClose1.Enabled = False
      .MnuPlay1.Enabled = False
      .MnuStop1.Enabled = False
      .MnuPause1.Enabled = False
      .MnuProperties1.Enabled = False
      .Check1.Value = 1
      .Check2.Value = 1
      .Slider1.Value = 1000
      .Slider3.Value = 1000
    End If
  
    If Not (Device2 Is Nothing) Then          'if device2 is a valid object
      If Len(Device2.FileName) <> 0 Then      'if it is-does it have a valid filename
        Frame2.Caption = "File Name: " & Device2.FileName
        .Frame2.Enabled = True
        .MnuClose2.Enabled = True
        .MnuPlay2.Enabled = True
        .MnuStop2.Enabled = False
        .MnuPause2.Enabled = False
        .MnuProperties2.Enabled = True
        .Check3.Value = Not Device2.LeftChannel
        .Check4.Value = Not Device2.RightChannel
        .Slider2.Value = Device2.Volume
        .Slider4.Value = Device2.Speed
      Else                                     'does not have a valid filename
        Frame2.Caption = "File Name: "
        .Frame2.Enabled = False
        .MnuClose2.Enabled = False
        .MnuPlay2.Enabled = False
        .MnuStop2.Enabled = False
        .MnuPause2.Enabled = False
        .MnuProperties2.Enabled = False
        .Check3.Value = Not Device2.LeftChannel
        .Check4.Value = Not Device2.RightChannel
        .Slider2.Value = Device2.Volume
        .Slider4.Value = Device2.Speed
      End If
    Else                                       'device2 is not valid-has not been set
      Frame2.Caption = "File Name: "
      .Frame2.Enabled = False
      .MnuClose2.Enabled = False
      .MnuPlay2.Enabled = False
      .MnuStop2.Enabled = False
      .MnuPause2.Enabled = False
      .MnuProperties2.Enabled = False
      .Check3.Value = 1
      .Check4.Value = 1
      .Slider2.Value = 1000
      .Slider4.Value = 1000
    End If
  End With
  
  mp3audio.DisplayError = True
  
End Sub

Private Sub Slider1_Change()
On Error Resume Next
  
  Device1.Volume = Slider1.Value          'set the volume
  
End Sub

Private Sub Slider1_Scroll()
On Error Resume Next

  Device1.Volume = Slider1.Value          'set the volume
  
End Sub

Private Sub Slider2_Change()
On Error Resume Next
  
  Device2.Volume = Slider2.Value          'set the volume

End Sub

Private Sub Slider2_Scroll()
On Error Resume Next
  
  Device2.Volume = Slider2.Value          'set the volume

End Sub

Private Sub Slider3_Change()
On Error Resume Next

  Device1.Speed = Slider3.Value           'set the speed
  
End Sub

Private Sub Slider3_Scroll()
On Error Resume Next

  Device1.Speed = Slider3.Value           'set the speed

End Sub

Private Sub Slider4_Change()
On Error Resume Next

  Device2.Speed = Slider4.Value           'set the speed

End Sub

Private Sub Slider4_Scroll()
On Error Resume Next

  Device2.Speed = Slider4.Value           'set the speed

End Sub
