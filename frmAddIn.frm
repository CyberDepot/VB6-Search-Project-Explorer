VERSION 5.00
Begin VB.Form frmAddIn 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Project Explorer - Setiing"
   ClientHeight    =   4860
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3045
      Left            =   150
      TabIndex        =   9
      Top             =   1020
      Width           =   7785
      Begin VB.Image Image1 
         Height          =   885
         Left            =   6360
         Picture         =   "frmAddIn.frx":0000
         Stretch         =   -1  'True
         Top             =   450
         Width           =   1245
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Close All Opened Windows, With Using This Future Press ""Left Alt+C"""
         Height          =   195
         Left            =   300
         TabIndex        =   16
         Top             =   2070
         Width           =   5190
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "It Clears immediate Window & Codes Smart Immediate Windows. To Use This Future, Press ""Left Alt+Z"""
         Height          =   195
         Left            =   300
         TabIndex        =   15
         Top             =   2580
         Width           =   7230
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Then It Does Search Your Object Name & With Pressing Enter,"
         Height          =   195
         Left            =   300
         TabIndex        =   14
         Top             =   1230
         Width           =   4365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The Add-in List Will Show All Results And You Can Select Each One You Like."
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   1560
         Width           =   5565
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "In Order To Accelerate Switching Of Add-in With Pressing ""Left Alt+S"""
         Height          =   195
         Left            =   300
         TabIndex        =   12
         Top             =   930
         Width           =   4965
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search In Project Explorer With Using Regular Expression"
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   630
         Width           =   4080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This Tools Can:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   270
         Width           =   1335
      End
   End
   Begin VB.CheckBox Chk_Close_Design 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close Designer Window"
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   180
      TabIndex        =   5
      Top             =   4470
      Width           =   2145
   End
   Begin VB.CheckBox Chk_Close_Code 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close Code Module"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   4470
      Width           =   2175
   End
   Begin VB.Timer TimerImmadiate 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4620
      Top             =   4230
   End
   Begin VB.CheckBox Chk_Show_Code 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Code Module"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   4110
      Width           =   2175
   End
   Begin VB.CheckBox Chk_Show_Design 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Designer Window"
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   4110
      Width           =   2145
   End
   Begin VB.CommandButton Btn_Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5310
      TabIndex        =   1
      Top             =   4260
      Width           =   1215
   End
   Begin VB.CommandButton Btn_OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6690
      TabIndex        =   0
      Top             =   4260
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email: (Admin@CyberRayaneh.com)"
      Height          =   195
      Left            =   2767
      TabIndex        =   8
      Top             =   750
      Width           =   2550
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created By: Mohsen Mousavi © 2016"
      Height          =   195
      Left            =   2707
      TabIndex        =   7
      Top             =   420
      Width           =   2670
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visual Basic 6 Search Project Explorer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2407
      TabIndex        =   6
      Top             =   90
      Width           =   3270
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Btn_Cancel_Click()
    Unload Me
End Sub

Private Sub Btn_OK_Click()
    If Chk_Show_Code.Value = vbUnchecked And Chk_Show_Design.Value = vbUnchecked Then
        MsgBox "To Show Windows You Must Select One Of Theme!", vbOKOnly + vbExclamation, "www.CyberRayaneh.com"
        Exit Sub
    End If
    
    If Chk_Close_Code.Value = vbUnchecked And Chk_Close_Design.Value = vbUnchecked Then
        MsgBox "To Close Windows You Must Select One Of Theme!", vbOKOnly + vbExclamation, "www.CyberRayaneh.com"
        Exit Sub
    End If
    
    Call SaveSetting(App.ProductName, "Setting", "Show_Disign", Chk_Show_Design.Value)
    Call SaveSetting(App.ProductName, "Setting", "Close_Disign", Chk_Close_Design.Value)
    Call SaveSetting(App.ProductName, "Setting", "Show_Code", Chk_Show_Code.Value)
    Call SaveSetting(App.ProductName, "Setting", "Close_Code", Chk_Close_Code.Value)
    
    Show_Disign = Chk_Show_Design.Value
    Close_Disign = Chk_Close_Design.Value
    Show_Code = Chk_Show_Code.Value
    Close_Code = Chk_Close_Code.Value
    
    Unload Me
End Sub

Private Sub Form_Load()
    Chk_Show_Design.Value = Abs(Show_Disign)
    Chk_Close_Design.Value = Abs(Close_Disign)
    Chk_Show_Code.Value = Abs(Show_Code)
    Chk_Close_Code.Value = Abs(Close_Code)
End Sub

'Clear Immediate Windows
Private Sub TimerImmadiate_Timer()
    Dim winActive As VBIDE.Window
    Dim winImm As VBIDE.Window
    Dim winImmSMART As VBIDE.Window
    
    Set winImm = VBInstance.Windows.Item("Immediate")
    Set winImmSMART = VBInstance.Windows.Item("Immediate - by CodeSMART")
    
    'Save the currently active window
    Set winActive = VBInstance.ActiveWindow

    If Not winImm Is Nothing Then
        'Do not clear if Window Not Visible
        If winImm.Visible = True Then
            winImm.SetFocus
            SendKeys "^({Home})", True
            SendKeys "^(+({End}))", True
            SendKeys "{Del}", True
        End If
    End If
   
    If Not winImmSMART Is Nothing Then
        'Do not clear if Window Not Visible
        If winImmSMART.Visible = True Then
            winImmSMART.SetFocus
            SendKeys "^({Home})", True
            SendKeys "^(+({End}))", True
            SendKeys "{Del}", True
        End If
    End If
    
    'Return to active window
    winActive.SetFocus
    
    Set winImm = Nothing
    Set winImmSMART = Nothing
    Set winActive = Nothing
    
    TimerImmadiate.Enabled = False
End Sub
