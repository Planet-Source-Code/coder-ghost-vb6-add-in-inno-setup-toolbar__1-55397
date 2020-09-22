Attribute VB_Name = "modFunction"
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWDEFAULT = 10

Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public InnoEXE As String

Public Function Find_Inno() As String
  Dim Location As String
  
  'Find_Inno = "C:\Program Files\Inno Setup 3\Compil32.exe"
  
  Find_Inno = Get_Key("ForcePath")
  
  If Find_Inno = "" Then
     Find_Inno = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\CLASSES\InnoSetupScriptFile\DefaultIcon", "")
  
     If Find_Inno <> "" Then Find_Inno = Mid(Find_Inno, 1, InStrRev(Find_Inno, ",") - 1)
  End If
  
  'HKEY_LOCAL_MACHINE\Software\CLASSES\InnoSetupScriptFile\DefaultIcon     :: (Default)
End Function

Public Function File_Exists(ByVal Path As String) As Boolean
  On Error GoTo Fallout
  
  File_Exists = False
  
  Open Path For Input As #1
  Close #1
    
  File_Exists = True
  
Fallout:
End Function

Public Sub API_WinExec(ByVal Command As String, ByVal Hidden As Boolean)
  Dim Mode As Integer
  
  Debug.Print Command
  
  Mode = SW_SHOWDEFAULT
  If Hidden Then Mode = SW_SHOWMINNOACTIVE
  
  WinExec Command, Mode
End Sub

Public Function Get_Key(ByVal Name As String) As String
  Get_Key = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\Final Stand\InnoToolbar", Name)
End Function

Public Function Write_Key(ByVal Name As String, ByVal Value As String) As Long
  Dim Rtn As Long
  
  Rtn = UpdateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Final Stand\InnoToolbar", Name, Value)
  
  Write_Key = Rtn
End Function

Public Sub ShellURL(ByVal Msg As String)
    Call ShellExecute(0&, vbNullString, Msg, vbNullString, vbNullString, vbNormalFocus)
End Sub
