VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10155
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   13395
   _ExtentX        =   23627
   _ExtentY        =   17912
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "Inno Setup Toolbar"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Visual Basic Interface
Public VBInstance                     As VBIDE.VBE

'the command bar that will hold the buttons
Private cmdBar                        As CommandBar
'lets add a few buttons
Private cmdBarBtn1                    As CommandBarButton
Private cmdBarBtn2                    As CommandBarButton
Private cmdBarBtn3                    As CommandBarButton
'we now need to enable the buttons to
'receive click event or their useless
Private WithEvents cmdBarBtnEvents1   As CommandBarEvents
Attribute cmdBarBtnEvents1.VB_VarHelpID = -1
Private WithEvents cmdBarBtnEvents2   As CommandBarEvents
Attribute cmdBarBtnEvents2.VB_VarHelpID = -1
Private WithEvents cmdBarBtnEvents3   As CommandBarEvents
Attribute cmdBarBtnEvents3.VB_VarHelpID = -1

Private tmpConfig As frmConfig

'
' This method adds the Add-In to VB.
'
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    'On Error GoTo Err_Handler:
    Dim RtnA As String, RtnB As String, RtnC As String, RtnD As String
    Dim CtlType As MsoControlType
    
    'Save the VB Instance
    Set VBInstance = Application
       
    'Locate the Inno Executable, if its installed.
    InnoEXE = Find_Inno
 
    'Create the toolbar
    'Set cmdBar = VBInstance.CommandBars.Add("Inno Setup Toolbar", msoBarFloating)
    'Set cmdBar = VBInstance.CommandBars.Add("Inno Setup", msoBarTop)
    Set cmdBar = VBInstance.CommandBars.Add("Inno Setup", msoBarTop)
    cmdBar.Protection = msoBarNoCustomize Or msoBarNoResize
    
    RtnA = Get_Key("Position")
    RtnB = Get_Key("RowIndex")
    RtnC = Get_Key("Left")
    RtnD = Get_Key("Top")
    
    If RtnA <> "" Then cmdBar.Position = CLng(RtnA)
    If RtnB <> "" Then cmdBar.RowIndex = CLng(RtnB)
    If RtnC <> "" Then cmdBar.Left = CLng(RtnC)
    If RtnD <> "" Then cmdBar.Top = CLng(RtnD)
       
    'Make it visible
    cmdBar.Visible = True
    
    'Set the control type were adding to the command bar
    CtlType = msoControlButton
    
    'We now need to add the buttons to the toolbar
    Set cmdBarBtn1 = cmdBar.Controls.Add(CtlType)
    Set cmdBarBtn2 = cmdBar.Controls.Add(CtlType)
    Set cmdBarBtn3 = cmdBar.Controls.Add(CtlType)
      
    'Create the properties for the toolbar.
    With cmdBarBtn1
        .Caption = "Script Editor"
        .ToolTipText = "Launch the Inno Script Editor"
        .Style = msoButtonIcon
        .FaceId = 593
    End With
    
    With cmdBarBtn2
        .Caption = "Compile Script"
        .ToolTipText = "Compile the Script"
        .Style = msoButtonIcon
        .FaceId = 1396
    End With
    
    With cmdBarBtn3
        .Caption = "Configuration"
        .ToolTipText = "Configure the Addin"
        .Style = msoButtonIcon
        .FaceId = 2946
    End With
    
   '-------------------------------------------
   ' we now need to link the buttons to events
   '-------------------------------------------
   With VBInstance
        Set cmdBarBtnEvents1 = .Events.CommandBarEvents(cmdBarBtn1)
        Set cmdBarBtnEvents2 = .Events.CommandBarEvents(cmdBarBtn2)
        Set cmdBarBtnEvents3 = .Events.CommandBarEvents(cmdBarBtn3)
   End With
 
 Exit Sub
Err_Handler:
     Err.Source = Err.Source & "." & VarType(Me) & ".AddinInstance_OnConnection"
     Debug.Print Err.Number & vbTab & Err.Source & Err.Description
     Err.Clear
     Resume Next
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'this is the time to destroy objects
'and references
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
 On Error GoTo Err_Handler:
   
     Call Write_Key("Position", CStr(cmdBar.Position))
     Call Write_Key("RowIndex", CStr(cmdBar.RowIndex))
     Call Write_Key("Left", CStr(cmdBar.Left))
     Call Write_Key("Top", CStr(cmdBar.Top))
     
     If tmpConfig Is Nothing Then
        'DO Nothing
     Else
        Unload tmpConfig
     End If
     
     'delete the buttons
     cmdBarBtn1.Delete
     cmdBarBtn2.Delete
     cmdBarBtn3.Delete
     
     'unset buttons reference
     Set cmdBarBtn1 = Nothing
     Set cmdBarBtn2 = Nothing
     Set cmdBarBtn3 = Nothing
     
     'unset events reference
     Set cmdBarBtnEvents1 = Nothing
     Set cmdBarBtnEvents2 = Nothing
     Set cmdBarBtnEvents3 = Nothing
     
     'destroy toolbar and its  variable
     cmdBar.Delete
     Set cmdBar = Nothing

     'kill core reference
     Set VBInstance = Nothing
     
 Exit Sub
Err_Handler:
     Err.Source = Err.Source & "." & VarType(Me) & ".AddinInstance_OnDisconnection"
     Debug.Print Err.Number & vbTab & Err.Source & Err.Description
     Err.Clear
     Resume Next
End Sub


Private Sub cmdBarBtnEvents1_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  'Edit Script
  '------------------------------------------------
  Dim ScriptPath As String, Rtn As Integer
  
  ScriptPath = VBInstance.ActiveVBProject.FileName
  If Len(ScriptPath) > 3 Then ScriptPath = Mid(ScriptPath, 1, Len(ScriptPath) - 3) & "iss"
  
  'Autodetect Inno
  If InnoEXE = "" Then
     InnoEXE = Find_Inno
     If InnoEXE = "" Then
        MsgBox "Unable to locate the Inno Compiler Program. Please configure the toolbar.", vbCritical + vbOKOnly, "VB6 - Inno Setup Toolbar"
        Set tmpConfig = New frmConfig
        tmpConfig.Show 1, Me
        Exit Sub
     End If
  End If
  
  If ScriptPath = "" Then
     MsgBox "This project has not been saved yet.", vbCritical + vbOKOnly, "VB6 - Inno Setup Toolbar"
     Exit Sub
  End If
  
  If VBInstance.ActiveVBProject.IsDirty Then
     Rtn = MsgBox("This project has been changed since your last save. Continue Anyway?", vbExclamation + vbYesNo, "VB6 - Inno Setup Toolbar")
     If Rtn <> vbYes Then Exit Sub
  End If
  
  If Not File_Exists(ScriptPath) Then
     'Create a default script
     Call Default_Script(ScriptPath)
  End If
  
  'Open the Script
  Call API_WinExec(Chr(34) & InnoEXE & Chr(34) & " " & Chr(34) & ScriptPath & Chr(34), False)
End Sub

Private Sub cmdBarBtnEvents2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  'Compile Script
  '------------------------------------------------
  Dim ScriptPath As String, Rtn As Integer
  
  ScriptPath = VBInstance.ActiveVBProject.FileName
  If Len(ScriptPath) > 3 Then ScriptPath = Mid(ScriptPath, 1, Len(ScriptPath) - 3) & "iss"
   
  'Autodetect Inno
  If InnoEXE = "" Then
     InnoEXE = Find_Inno
     If InnoEXE = "" Then
        MsgBox "Unable to locate the Inno Compiler Program. Please configure the toolbar.", vbCritical + vbOKOnly, "VB6 - Inno Setup Toolbar"
        Set tmpConfig = New frmConfig
        tmpConfig.Show 1, Me
        Exit Sub
     End If
  End If
  
  If ScriptPath = "" Then
     MsgBox "This project has not been saved yet.", vbCritical + vbOKOnly, "VB6 - Inno Setup Toolbar"
     Exit Sub
  End If
  
  If VBInstance.ActiveVBProject.IsDirty Then
     Rtn = MsgBox("This project has been changed since your last save. Continue Anyway?", vbExclamation + vbYesNo, "VB6 - Inno Setup Toolbar")
     If Rtn <> vbYes Then Exit Sub
  End If
  
  If Not File_Exists(ScriptPath) Then
     MsgBox "Unable to locate the script. You must generate a script first.", vbCritical + vbOKOnly, "VB6 - Inno Setup Toolbar"
     Exit Sub
  End If
  
  'Recompile the Project
  If Get_Key("C0") = "1" Then
     Rtn = vbYes
  ElseIf Get_Key("C2") = "1" Then
     Rtn = vbNo
  Else
     Rtn = MsgBox("Would you like to recompile the project before continuing?", vbExclamation + vbYesNo, "VB6 - Inno Setup Toolbar")
  End If
  
  If Rtn = vbYes Then
     On Error Resume Next
     
     Err.Clear
     
     VBInstance.ActiveVBProject.MakeCompiledFile
     
     If Err.Number <> 0 Then
        Rtn = MsgBox("Failed to compile the project. It may be running. Continue anyway?", vbCritical + vbYesNo, "VB6 - Inno Setup Toolbar")
        If Rtn = vbNo Then Exit Sub
     End If
     
     On Error GoTo 0
  End If
  
  'Compile the Script
  Call API_WinExec(Chr(34) & InnoEXE & Chr(34) & " /cc " & Chr(34) & ScriptPath & Chr(34), True)
  
  'MsgBox "Script Compile Finished.", vbExclamation + vbOKOnly, "VB6 - Inno Setup Toolbar"
End Sub

Private Sub cmdBarBtnEvents3_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  'Configuration
  '------------------------------------------------
  
  Set tmpConfig = New frmConfig
  
  tmpConfig.Show 1, Me
End Sub

'
' Create a Default Inno Script
'
Public Sub Default_Script(ByVal Path As String)
  Open Path For Output As #1
  
  Dim ShortName As String, SourceDir As String
  
  ShortName = Mid(VBInstance.ActiveVBProject.BuildFileName, InStrRev(VBInstance.ActiveVBProject.BuildFileName, "\") + 1)
  
  SourceDir = Mid(VBInstance.ActiveVBProject.BuildFileName, 1, InStrRev(VBInstance.ActiveVBProject.BuildFileName, "\") - 1)
  
  ' Output Header
  Print #1, "; "
  Print #1, "; Install Script for " & VBInstance.ActiveVBProject.Name
  Print #1, "; "
  Print #1, "; Generated by the 'Inno Setup Toolbar for VB6'"
  Print #1, "; Written and Programmed by Brian Haase"
  Print #1, "; "
  Print #1, " "
  
  If Get_Key("C1") <> "1" Then
    Print #1, "[Setup]"
    Print #1, "AppName=" & VBInstance.ActiveVBProject.Name
    Print #1, "AppVerName=" & VBInstance.ActiveVBProject.Name & " 1.0"
    Print #1, "AppPublisher=InnoSetupAddin"
    Print #1, "DefaultDirName={pf}\" & VBInstance.ActiveVBProject.Name
    Print #1, "DefaultGroupName=" & VBInstance.ActiveVBProject.Name
    Print #1, "SourceDir=" & SourceDir
    Print #1, "OutputDir=" & SourceDir & "\Output"
    Print #1, " "
    Print #1, "[Tasks]"
    Print #1, "Name: ""desktopicon""; Description: ""Create a &desktop icon""; GroupDescription: ""Additional icons:"""
    Print #1, " "
    Print #1, "[Files]"
    Print #1, "Source: """ & VBInstance.ActiveVBProject.BuildFileName & """; DestDir: ""{app}""; Flags: ignoreversion"
    Print #1, "; NOTE: Don't use ""Flags: ignoreversion"" on any shared system files"
    Print #1, " "
    Print #1, "[icons]"
    Print #1, "Name: ""{group}\" & VBInstance.ActiveVBProject.Name & """; Filename: ""{app}\" & ShortName & """"
    Print #1, "Name: ""{group}\Uninstall My Program""; Filename: ""{uninstallexe}"""
    Print #1, "Name: ""{userdesktop}\" & VBInstance.ActiveVBProject.Name & """; Filename: ""{app}\" & ShortName & """; Tasks: desktopicon"
    Print #1, " "
    Print #1, "[Run]"
    Print #1, "Filename: ""{app}\" & ShortName & """; Description: ""Launch My Program""; Flags: nowait postinstall skipifsilent"
  End If
  
  Close #1
End Sub
