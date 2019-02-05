VERSION 5.00
Object = "{02353968-C1C9-4E0A-88D3-18759BDC60FE}#1.0#0"; "AeroSuite.ocx"
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormTentang 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tentang"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormTentang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin AeroSuite.AeroGroupBox AeroGroupBox1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8070
      BorderColor     =   11908533
      BackColor       =   14737632
      BackColor2      =   13882323
      HeadColor1      =   14737632
      HeadColor2      =   13092807
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Dacara_dcButton.dcButton cmTutup 
         Height          =   375
         Left            =   6000
         TabIndex        =   14
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16751432
         ButtonStyle     =   1
         Caption         =   "&Tutup"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   600
         Picture         =   "FormTentang.frx":0442
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RikySoft Simple Account Manager"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   600
         Width           =   3060
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   360
         X2              =   360
         Y1              =   600
         Y2              =   1320
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   360
         X2              =   1200
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   1200
         X2              =   1200
         Y1              =   600
         Y2              =   1320
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         X1              =   360
         X2              =   1200
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Developer/Programmer by"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   11
         Top             =   1680
         Width           =   1830
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright ©_2013. RikySoft Software House Foundation. All Rights Reserved."
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   10
         Top             =   2760
         Width           =   5025
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Homepage"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   9
         Top             =   2040
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   8
         Top             =   2400
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check update at :"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   7
         Top             =   3480
         Width           =   1110
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rizky Khafitsyah"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3320
         MouseIcon       =   "FormTentang.frx":0884
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://rikymetalist.blogspot.com"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   2280
         MouseIcon       =   "FormTentang.frx":154E
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   2040
         Width           =   2145
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "rikymetal10@gmail.com"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   2280
         MouseIcon       =   "FormTentang.frx":2218
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   2400
         Width           =   1515
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   3
         Top             =   2040
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   2
         Top             =   2400
         Width           =   45
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://rikymetalist.blogspot.com/p/software-ku.html"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   2640
         MouseIcon       =   "FormTentang.frx":2EE2
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   3480
         Width           =   3570
      End
   End
   Begin XPEngine.XPControl XP_Engine 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
End
Attribute VB_Name = "FormTentang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    DisableCloseBtn Me
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        cmTutup.Caption = "Tutup"
        Me.Caption = "Tentang"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        cmTutup.Caption = "Close"
        Me.Caption = "About"
    End If
    XP_Engine.StartEngine
End Sub

Private Sub Label10_Click()
    Kalimat = "rikymetal10@gmail.com"
    EMAIL = ShellExecute(0, vbNullString, "mailto:" & Kalimat, "", "", vbNormalFocus)
    Label10.ForeColor = vbBlack
End Sub

Private Sub Label13_Click()
    Kalimat = "http://rikymetalist.blogspot.com/p/software-ku.html"
    SITUS = ShellExecute(0, vbNullString, Kalimat, "", "", vbNormalFocus)
    Label13.ForeColor = vbBlack
End Sub

Private Sub Label8_Click()
    Kalimat = "http://facebook.com/RizkyKhafitsyah"
    SITUS = ShellExecute(0, vbNullString, Kalimat, "", "", vbNormalFocus)
    Label8.ForeColor = vbBlack
End Sub

Private Sub Label9_Click()
    Kalimat = "http://rikymetalist.blogspot.com"
    SITUS = ShellExecute(0, vbNullString, Kalimat, "", "", vbNormalFocus)
    Label9.ForeColor = vbBlack
End Sub
