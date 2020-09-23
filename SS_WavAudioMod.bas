Attribute VB_Name = "Mp3AudioMod"
'*********************************************
'*Created By: Jason George                   *
'*Date: Dec. 16 2004                         *
'*Email: mescalito@adelphia.net              *
'*********************************************

Option Explicit

'declare Api function
Private Declare Function CoCreateGuid Lib "ole32" (id As Any) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Public Const DevType As String = " MPEGVideo "          'tells the Mci device what drivers to use
                                                        'to play Mp3 Files with...
                                                        'Note: limited wave support/not all properties
                                                        'available when playing waves

Public Type MciErrorType                                'support for a get last function error/not implemented
  ErrNum As Long
  ErrStr As String
End Type
Public pLastMciError As MciErrorType                    'support for a get last function error/not implemented

Public pDisplayError As Boolean                         'current Display Error value

'returns an error string from a Mci error number
Private Function GetMCIErrorString(lErr As Long) As String
On Error GoTo ErrHnd
       
    GetMCIErrorString = Space$(255)
    mciGetErrorString lErr, GetMCIErrorString, Len(GetMCIErrorString)
    GetMCIErrorString = Trim$(GetMCIErrorString)
    
Exit Function
ErrHnd:
  GetMCIErrorString = ""
  
End Function

Public Sub DisplayMciError(lErr As Long, Optional ByVal ProcName As String)
On Error Resume Next
  
  pLastMciError.ErrNum = lErr
  pLastMciError.ErrStr = GetMCIErrorString(lErr)
  
  If pDisplayError = True Then    'if Mci error displaying is on
    If lErr <> 0 Then             'if Mci error NOT success
      'display the error number/string and optional Sub or Function name
      MsgBox pLastMciError.ErrStr, vbInformation, ProcName & "(" & lErr & ")"
    End If
  End If
  
End Sub

'create a Unique 16 character string
Public Function SS_CreateGUID() As String
On Error GoTo ErrHnd
  
  Dim id(0 To 15) As Byte
  Dim Cnt As Long, GUID As String
    
    If CoCreateGuid(id(0)) = 0 Then
      For Cnt = 0 To 15
        SS_CreateGUID = SS_CreateGUID + IIf(id(Cnt) < 16, "0", "") + Hex$(id(Cnt))
      Next Cnt
      
      SS_CreateGUID = Left$(SS_CreateGUID, 8) + "-" + Mid$(SS_CreateGUID, 9, 4) + "-" + Mid$(SS_CreateGUID, 13, 4) + "-" + Mid$(SS_CreateGUID, 17, 4) + "-" + Right$(SS_CreateGUID, 12)
      
    Else
      GoTo ErrHnd
    End If
    
Exit Function
ErrHnd:
  SS_CreateGUID = ""
  
End Function

Public Function SS_GetShortPath(StrFileName As String) As String
On Local Error GoTo ErrHnd

  Dim lngRes As Long
  Dim strPath As String
    
    strPath = String$(165, 0)
    
    lngRes = GetShortPathName(StrFileName, strPath, 164)
    
    SS_GetShortPath = Left$(strPath, lngRes)
    
Exit Function
ErrHnd:
  SS_GetShortPath = ""
  
End Function
