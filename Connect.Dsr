VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   6735
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   14475
   _ExtentX        =   25532
   _ExtentY        =   11880
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "Cyber Rayaneh Search Project Explorer"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSafe     =   -1  'True
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed As Boolean
Dim mcbMenuCommandBar As Office.CommandBarControl
'Dim mfrmAddIn As New frmAddIn

Public WithEvents MenuHandler As CommandBarEvents 'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Private WithEvents ProjEvent    As VBProjectsEvents
Attribute ProjEvent.VB_VarHelpID = -1
Private WithEvents CompEvent    As VBComponentsEvents
Attribute CompEvent.VB_VarHelpID = -1

Private Const mstrGuid As String = "{787C322C-DD0C-4e72-9AAA-F31EE8620163}"

Sub Hide()
    FormDisplayed = False
End Sub

Sub Show()
    On Error Resume Next
    
    Set VBInstance = VBInstance
    Set Connect = Me
    FormDisplayed = True
End Sub

Private Sub ProjEvent_ItemActivated(ByVal VBProject As VBIDE.VBProject)
    'The one and only project (or a new project in a multi-project app) was selected
    Set CompEvent = VBInstance.Events.VBComponentsEvents(VBProject) 'so point to the components events of this newly selected project
End Sub

Private Sub CompEvent_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
    'A new component was selected
    UnhookKeyboard
    HookKeyboard
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application

    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    Debug.Print VBInstance.FullName

    With VBInstance
        Set ProjEvent = .Events.VBProjectsEvents 'ensure that we're kept up to date about project events
        ProjEvent_ItemActivated .ActiveVBProject 'and components events of this project also (VB should fire this event initially also, but doesn't - so we fake it)
    End With 'VBINSTANCE

    Show_Disign = GetSetting(App.ProductName, "Setting", "Show_Disign", True)
    Close_Disign = GetSetting(App.ProductName, "Setting", "Close_Disign", True)
    Show_Code = GetSetting(App.ProductName, "Setting", "Show_Code", True)
    Close_Code = GetSetting(App.ProductName, "Setting", "Close_Code", True)

    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    End If
  
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            Me.Show
        End If
    End If
  

  'Convert the ActiveX document into a dockable tool window in the VB IDE
  'Uses the CreateToolWindow function
  'VB help doesn't explain this line very clearly so here's my explanation.
  'Set to an object of type Window (this can be private to the Add-in designer, it's only purpose is to
  '              allow you to make the add-in visible either when the menu button is clicked or during VB startup (see below)
  
  '1st parameter comes form this routine's parameters (don't touch)
  '2nd parameter is project name (up at top of project window on right (usually) plus a '.' connector plus the name of your UserDocument (bottom of project window)
  '3rp parameter the name you want to appear on tool (I use a routine to keep the Ver number up to date)
  '4th parameter a Guid number (you must generate a new one for each program DO NOT just cut and paste the one in the help file
  '              Use a tool called Guidgen.exe, which is located in the \tools\idgen directory of Visual Basic CD.
  '5th parameter The name of your Userdocument (or a variable holding a reference to it which is declared  As <Type = UserDocument's name>
  '              I used 'Public mobjDoc           As docfind' which I placed in a bas Module so that it was public to other parts of code
    Set mWindow = VBInstance.Windows.CreateToolWindow(AddInInst, "CyberRayaneh_SearchObjects.UD_Search", AppDetails, mstrGuid, mObjDoc)
    mWindow.Visible = True

    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
End Sub

Public Function AppDetails() As String
  'VB tab width and font
  With App
    AppDetails = "Search Project Explorer" & " V" & .Major & "." & .Minor & "." & .Revision
  End With
End Function

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

