VERSION 5.00
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form Form_JEJARING_SOSIAL 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "---------------"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_JEJARING_SOSIAL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   0
      Top             =   5160
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Height          =   1215
         Left            =   0
         Picture         =   "Form_JEJARING_SOSIAL.frx":74F2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   7095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   6855
      Begin XPControls.XPText textJejaringSosial 
         Height          =   330
         Index           =   0
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         Text            =   "XPText1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   0
         Left            =   4965
         TabIndex        =   4
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Salin"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   1
         Left            =   4965
         TabIndex        =   5
         Top             =   600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Salin"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   2
         Left            =   4965
         TabIndex        =   6
         Top             =   960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Salin"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   3
         Left            =   4965
         TabIndex        =   7
         Top             =   1320
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Salin"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   4
         Left            =   4965
         TabIndex        =   8
         Top             =   1680
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Salin"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   5
         Left            =   4965
         TabIndex        =   9
         Top             =   2040
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Salin"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   6
         Left            =   4965
         TabIndex        =   10
         Top             =   2400
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Salin"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   7
         Left            =   4965
         TabIndex        =   11
         Top             =   2760
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Salin"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   0
         Left            =   5610
         TabIndex        =   12
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Hapus"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   1
         Left            =   5610
         TabIndex        =   13
         Top             =   600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Hapus"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   2
         Left            =   5610
         TabIndex        =   14
         Top             =   960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Hapus"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   3
         Left            =   5610
         TabIndex        =   15
         Top             =   1320
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Hapus"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   4
         Left            =   5610
         TabIndex        =   16
         Top             =   1680
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Hapus"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   5
         Left            =   5610
         TabIndex        =   17
         Top             =   2040
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Hapus"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   6
         Left            =   5610
         TabIndex        =   18
         Top             =   2400
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Hapus"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   7
         Left            =   5610
         TabIndex        =   19
         Top             =   2760
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Hapus"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmPergi 
         Height          =   330
         Index           =   0
         Left            =   6255
         TabIndex        =   20
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "Go >"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmPergi 
         Height          =   330
         Index           =   1
         Left            =   6255
         TabIndex        =   21
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "Go >"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPText textJejaringSosial 
         Height          =   330
         Index           =   1
         Left            =   1800
         TabIndex        =   22
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         Text            =   "XPText1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPText textJejaringSosial 
         Height          =   330
         Index           =   2
         Left            =   1800
         TabIndex        =   23
         Top             =   960
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         Text            =   "XPText1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPText textJejaringSosial 
         Height          =   330
         Index           =   3
         Left            =   1800
         TabIndex        =   24
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         Text            =   "XPText1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPText textJejaringSosial 
         Height          =   330
         Index           =   4
         Left            =   1800
         TabIndex        =   25
         Top             =   1680
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         Text            =   "XPText1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPText textJejaringSosial 
         Height          =   330
         Index           =   5
         Left            =   1800
         TabIndex        =   26
         Top             =   2040
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         Text            =   "XPText1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPText textJejaringSosial 
         Height          =   330
         Index           =   6
         Left            =   1800
         TabIndex        =   27
         Top             =   2400
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         Text            =   "XPText1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPText textJejaringSosial 
         Height          =   330
         Index           =   7
         Left            =   1800
         TabIndex        =   28
         Top             =   2760
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         Text            =   "XPText1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmSet 
         Height          =   330
         Left            =   6255
         TabIndex        =   29
         Top             =   2400
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "Set"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cmbDataJejaringSosialLalu 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   240
         Width           =   3135
      End
      Begin VB.ComboBox cmbDataJejaringSosialLalu 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   600
         Width           =   3135
      End
      Begin VB.ComboBox cmbDataJejaringSosialLalu 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   960
         Width           =   3135
      End
      Begin VB.ComboBox cmbDataJejaringSosialLalu 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1320
         Width           =   3135
      End
      Begin VB.ComboBox cmbDataJejaringSosialLalu 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1680
         Width           =   3135
      End
      Begin VB.ComboBox cmbDataJejaringSosialLalu 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   2040
         Width           =   3135
      End
      Begin VB.ComboBox cmbDataJejaringSosialLalu 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   2400
         Width           =   3135
      End
      Begin VB.ComboBox cmbDataJejaringSosialLalu 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1680
         TabIndex        =   53
         Top             =   2760
         Width           =   45
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1680
         TabIndex        =   52
         Top             =   2400
         Width           =   45
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1680
         TabIndex        =   51
         Top             =   2040
         Width           =   45
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1680
         TabIndex        =   50
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1680
         TabIndex        =   49
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1680
         TabIndex        =   48
         Top             =   960
         Width           =   45
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1680
         TabIndex        =   47
         Top             =   600
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1680
         TabIndex        =   46
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -480
         TabIndex        =   45
         Top             =   2760
         Width           =   1995
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -480
         TabIndex        =   44
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pemilik Akun"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -480
         TabIndex        =   43
         Top             =   2040
         Width           =   1995
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "URL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -480
         TabIndex        =   42
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -480
         TabIndex        =   41
         Top             =   1320
         Width           =   1995
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat E-Mail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -480
         TabIndex        =   40
         Top             =   960
         Width           =   1995
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pengguna"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -480
         TabIndex        =   39
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Jejaring"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -480
         TabIndex        =   38
         Top             =   240
         Width           =   1995
      End
   End
   Begin Dacara_dcButton.dcButton cmSimpan 
      Height          =   375
      Left            =   120
      TabIndex        =   54
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   12632256
      ButtonStyle     =   3
      Caption         =   "&Simpan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "Form_JEJARING_SOSIAL.frx":20794
      PicHot          =   "Form_JEJARING_SOSIAL.frx":20AE6
      PicNormal       =   "Form_JEJARING_SOSIAL.frx":20E38
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmReset 
      Height          =   375
      Left            =   1440
      TabIndex        =   55
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   12632256
      ButtonStyle     =   3
      Caption         =   "&Reset"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "Form_JEJARING_SOSIAL.frx":2118A
      PicHot          =   "Form_JEJARING_SOSIAL.frx":21CD4
      PicNormal       =   "Form_JEJARING_SOSIAL.frx":2281E
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmBatal 
      Height          =   375
      Left            =   5760
      TabIndex        =   56
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   12632256
      ButtonStyle     =   3
      Caption         =   "&Batal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "Form_JEJARING_SOSIAL.frx":23368
      PicHot          =   "Form_JEJARING_SOSIAL.frx":237BA
      PicNormal       =   "Form_JEJARING_SOSIAL.frx":23C0C
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmVerifikasi 
      Height          =   375
      Left            =   2760
      TabIndex        =   57
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   12632256
      ButtonStyle     =   3
      Caption         =   "Verifikasi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "Form_JEJARING_SOSIAL.frx":2405E
      PicHot          =   "Form_JEJARING_SOSIAL.frx":244B0
      PicNormal       =   "Form_JEJARING_SOSIAL.frx":24902
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmBantuan 
      Height          =   375
      Left            =   4080
      TabIndex        =   58
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   12632256
      ButtonStyle     =   3
      Caption         =   "&Bantuan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "Form_JEJARING_SOSIAL.frx":24D54
      PicHot          =   "Form_JEJARING_SOSIAL.frx":251A6
      PicNormal       =   "Form_JEJARING_SOSIAL.frx":255F8
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin XPEngine.XPControl XP_Engine 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
End
Attribute VB_Name = "Form_JEJARING_SOSIAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub SambungkanKontrolKeADODC_UTAMA()
    If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
        CN_FormUtama.CursorLocation = adUseClient
        CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
        With FORM_UTAMA.ADODC_UTAMA
            .ConnectionString = CN_FormUtama.ConnectionString
            .RecordSource = "Select * From tbJejaringSosial order by Nama_Jejaring asc;"
            .Refresh
        End With
End Sub
Sub AturKontrol()
IsiTextBoxKosong_ID(0) = "Misal : Facebook..."
IsiTextBoxKosong_ID(1) = "(Nama Pengguna untuk Login ke Jejaring)..."
IsiTextBoxKosong_ID(2) = "Misal : email.saya@blabla.com"
IsiTextBoxKosong_ID(3) = ""
IsiTextBoxKosong_ID(4) = "Misal : http://rikymetalist.blogspot.com"
IsiTextBoxKosong_ID(5) = "(Tulis Nama Anda)..."
IsiTextBoxKosong_ID(6) = "(Klik 'Set' untuk mengatur..._"
IsiTextBoxKosong_ID(7) = "(Tambahkan keterangan lainnya)..."
IsiTextBoxKosong_EN(0) = "Eg : Facebook..."
IsiTextBoxKosong_EN(1) = "(User Name for login to social net)..."
IsiTextBoxKosong_EN(2) = "Eg : my.email@blabla.com"
IsiTextBoxKosong_EN(3) = ""
IsiTextBoxKosong_EN(4) = "Eg : http://rikymetalist.blogspot.com"
IsiTextBoxKosong_EN(5) = "(Put in your name)..."
IsiTextBoxKosong_EN(6) = "(Click 'Set' for set date...)"
IsiTextBoxKosong_EN(7) = "(Add other description)..."
    For NomorIndex = 0 To 7
        With textJejaringSosial.Item(NomorIndex)
            .ForeColor = SilverTua
            .MaxLength = 254
        End With
    Next
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textJejaringSosial
            .Item(0).Text = IsiTextBoxKosong_ID(0)
            .Item(1).Text = IsiTextBoxKosong_ID(1)
            .Item(2).Text = IsiTextBoxKosong_ID(2)
            .Item(3).Text = IsiTextBoxKosong_ID(3)
            .Item(4).Text = IsiTextBoxKosong_ID(4)
            .Item(5).Text = IsiTextBoxKosong_ID(5)
            .Item(6).Text = IsiTextBoxKosong_ID(6)
            .Item(7).Text = IsiTextBoxKosong_ID(7)
        End With
    Else
        With textJejaringSosial
            .Item(0).Text = IsiTextBoxKosong_EN(0)
            .Item(1).Text = IsiTextBoxKosong_EN(1)
            .Item(2).Text = IsiTextBoxKosong_EN(2)
            .Item(3).Text = IsiTextBoxKosong_EN(3)
            .Item(4).Text = IsiTextBoxKosong_EN(4)
            .Item(5).Text = IsiTextBoxKosong_EN(5)
            .Item(6).Text = IsiTextBoxKosong_EN(6)
            .Item(7).Text = IsiTextBoxKosong_EN(7)
        End With
    End If
    For NomorIndex = 0 To 7
        With cmbDataJejaringSosialLalu.Item(NomorIndex)
            .Clear
        End With
    Next
    IsiCMBDataLalu
    DisableCloseBtn Me
    'PENGATURAN UNTUK ALWAYS ON TOP
    If FormPengaturan.cekAlwaysOnTop.Value = Checked Then
        SetOnTop (Me.hwnd)
    ElseIf FormPengaturan.cekAlwaysOnTop.Value = Unchecked Then
        NotOnTop (Me.hwnd)
    End If
    For Each Objek In Me
        If TypeName(Objek) = "Label" Or TypeName(Objek) = "dcButton" Or TypeName(Objek) = "AeroCheckBox" Or TypeName(Objek) = "TextBox" Or TypeName(Objek) = "ComboBox" Then
            With Objek
                .Font.Name = "Agency FB"
                .Font.Size = 11
            End With
        End If
        If TypeName(Objek) = "XPText" Then Objek.Font.Name = "Agency FB"
    Next
    XP_Engine.StartEngine
End Sub
Sub IsiCMBDataLalu()
    SambungkanKontrolKeADODC_UTAMA
    With cmbDataJejaringSosialLalu
        .Item(0).Clear
        .Item(1).Clear
        .Item(2).Clear
        .Item(3).Clear
        .Item(4).Clear
        .Item(5).Clear
        .Item(6).Clear
        .Item(7).Clear
    End With
    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
        With cmbDataJejaringSosialLalu
            .Item(0).AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(0).Value
            .Item(1).AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(1).Value
            .Item(2).AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(2).Value
            .Item(3).AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(3).Value
            .Item(4).AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(4).Value
            .Item(5).AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(5).Value
            .Item(6).AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(6).Value
            .Item(7).AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(7).Value
        End With
        FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
    Loop
    FORM_UTAMA.ADODC_UTAMA.Refresh
End Sub
Sub PENGATURAN_BAHASA()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Label1.Caption = "Nama Jejaring"
        Label2.Caption = "Nama Pengguna"
        Label3.Caption = "Alamat E-Mail"
        Label4.Caption = "Password"
        Label5.Caption = "URL"
        Label6.Caption = "Pemilik Akun"
        Label7.Caption = "Tanggal"
        Label8.Caption = "Keterangan"
        For NomorIndex = 0 To 7
            cmSalin.Item(NomorIndex).Caption = "&Salin"
            cmHapus.Item(NomorIndex).Caption = "&Hapus"
        Next
        cmSimpan.Caption = "&Simpan"
        cmReset.Caption = "&Reset"
        cmBatal.Caption = "&Batal"
        cmVerifikasi.Caption = "&Verifikasi"
        cmBantuan.Caption = "&Bantuan"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Label1.Caption = "Social Name"
        Label2.Caption = "User Name"
        Label3.Caption = "Mail Address"
        Label4.Caption = "Passwords"
        Label5.Caption = "URL"
        Label6.Caption = "Account Owner"
        Label7.Caption = "Date"
        Label8.Caption = "Description"
        For NomorIndex = 0 To 7
            cmSalin.Item(NomorIndex).Caption = "&Copy"
            cmHapus.Item(NomorIndex).Caption = "&Delete"
        Next
        cmSimpan.Caption = "&Save"
        cmReset.Caption = "&Reset"
        cmBatal.Caption = "&Cancel"
        cmVerifikasi.Caption = "&Verify"
        cmBantuan.Caption = "&Help"
    End If
End Sub
Sub KosongkanTextBox()
    For NomorIndex = 0 To 7
        With textJejaringSosial.Item(NomorIndex)
            .ForeColor = SilverTua
            .MaxLength = 254
        End With
    Next
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textJejaringSosial
            .Item(0).Text = IsiTextBoxKosong_ID(0)
            .Item(1).Text = IsiTextBoxKosong_ID(1)
            .Item(2).Text = IsiTextBoxKosong_ID(2)
            .Item(3).Text = IsiTextBoxKosong_ID(3)
            .Item(4).Text = IsiTextBoxKosong_ID(4)
            .Item(5).Text = IsiTextBoxKosong_ID(5)
            .Item(6).Text = IsiTextBoxKosong_ID(6)
            .Item(7).Text = IsiTextBoxKosong_ID(7)
        End With
    Else
        With textJejaringSosial
            .Item(0).Text = IsiTextBoxKosong_EN(0)
            .Item(1).Text = IsiTextBoxKosong_EN(1)
            .Item(2).Text = IsiTextBoxKosong_EN(2)
            .Item(3).Text = IsiTextBoxKosong_EN(3)
            .Item(4).Text = IsiTextBoxKosong_EN(4)
            .Item(5).Text = IsiTextBoxKosong_EN(5)
            .Item(6).Text = IsiTextBoxKosong_EN(6)
            .Item(7).Text = IsiTextBoxKosong_EN(7)
        End With
    End If
End Sub
Sub SIMPAN_KE_DATABASE()
On Error GoTo HancurkanError
If FormPengaturan.cekPesanKonfirmasi.Value = Checked Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Pesan = MsgBox("Anda yakin isian Anda sudah benar?", vbQuestion + vbYesNo, "Konfirmasi")
    Else
        Pesan = MsgBox("Are you sure with your entry?", vbQuestion + vbYesNo, "Confirmation")
    End If
        If Pesan = vbYes Then
            SambungkanKontrolKeADODC_UTAMA
                If Me.cmSimpan.Caption = "&Save" Or Me.cmSimpan.Caption = "&Simpan" Then
                        With FORM_UTAMA.ADODC_UTAMA
                            .Recordset.AddNew
                            .Recordset.Fields(0).Value = textJejaringSosial(0).Text
                            .Recordset.Fields(1).Value = textJejaringSosial(1).Text
                            .Recordset.Fields(2).Value = textJejaringSosial(2).Text
                            .Recordset.Fields(3).Value = textJejaringSosial(3).Text
                            .Recordset.Fields(4).Value = textJejaringSosial(4).Text
                            .Recordset.Fields(5).Value = textJejaringSosial(5).Text
                            .Recordset.Fields(6).Value = textJejaringSosial(6).Text
                            .Recordset.Fields(7).Value = textJejaringSosial(7).Text
                            .Recordset.Update
                            .Refresh
                        End With
                ElseIf Me.cmSimpan.Caption = "&Update" Or Me.cmSimpan.Caption = "&Perbarui" Then
                        With FormManage.AdodcMain
                            .Recordset.Delete
                            .Recordset.AddNew
                            .Recordset.Fields(0).Value = textJejaringSosial(0).Text
                            .Recordset.Fields(1).Value = textJejaringSosial(1).Text
                            .Recordset.Fields(2).Value = textJejaringSosial(2).Text
                            .Recordset.Fields(3).Value = textJejaringSosial(3).Text
                            .Recordset.Fields(4).Value = textJejaringSosial(4).Text
                            .Recordset.Fields(5).Value = textJejaringSosial(5).Text
                            .Recordset.Fields(6).Value = textJejaringSosial(6).Text
                            .Recordset.Fields(7).Value = textJejaringSosial(7).Text
                            .Recordset.Update
                            .Refresh
                        End With
                        FormManage.AturDatabase
                End If
                With FormPengaturan
                    If .cekAutoRefresh.Value = Checked Then FORM_UTAMA.cmJejaringSosial_Click
                    If .cekTampilkanPesanSimpan.Value = Checked Then
                        If .cmbBahasa.ListIndex = 0 Then
                            MsgBox "Data berhasil disimpan!", vbInformation + vbOKOnly, "Sukses"
                        Else
                            MsgBox "Data saved successed!", vbInformation + vbOKOnly, "Success"
                        End If
                    End If
                    If .cekKosongkanInput.Value = Checked Then KosongkanTextBox
                    If .cekTutupForm.Value = Checked Then Unload Me
                    If .cmbBahasa.ListIndex = 0 Then
                        cmBatal.Caption = "&Tutup"
                    Else
                        cmBatal.Caption = "&Close"
                    End If
                End With
        End If
Else
    SambungkanKontrolKeADODC_UTAMA
        If Me.cmSimpan.Caption = "&Save" Or Me.cmSimpan.Caption = "&Simpan" Then
                With FORM_UTAMA.ADODC_UTAMA
                    .Recordset.AddNew
                    .Recordset.Fields(0).Value = textJejaringSosial(0).Text
                    .Recordset.Fields(1).Value = textJejaringSosial(1).Text
                    .Recordset.Fields(2).Value = textJejaringSosial(2).Text
                    .Recordset.Fields(3).Value = textJejaringSosial(3).Text
                    .Recordset.Fields(4).Value = textJejaringSosial(4).Text
                    .Recordset.Fields(5).Value = textJejaringSosial(5).Text
                    .Recordset.Fields(6).Value = textJejaringSosial(6).Text
                    .Recordset.Fields(7).Value = textJejaringSosial(7).Text
                    .Recordset.Update
                    .Refresh
                End With
                FormManage.AturDatabase
        ElseIf Me.cmSimpan.Caption = "&Update" Or Me.cmSimpan.Caption = "&Perbarui" Then
                With FormManage.AdodcMain
                    .Recordset.Delete
                    .Recordset.AddNew
                    .Recordset.Fields(0).Value = textJejaringSosial(0).Text
                    .Recordset.Fields(1).Value = textJejaringSosial(1).Text
                    .Recordset.Fields(2).Value = textJejaringSosial(2).Text
                    .Recordset.Fields(3).Value = textJejaringSosial(3).Text
                    .Recordset.Fields(4).Value = textJejaringSosial(4).Text
                    .Recordset.Fields(5).Value = textJejaringSosial(5).Text
                    .Recordset.Fields(6).Value = textJejaringSosial(6).Text
                    .Recordset.Fields(7).Value = textJejaringSosial(7).Text
                    .Recordset.Update
                    .Refresh
                End With
                FormManage.AturDatabase
        End If
        With FormPengaturan
            If .cekAutoRefresh.Value = Checked Then FORM_UTAMA.cmJejaringSosial_Click
            If .cekTampilkanPesanSimpan.Value = Checked Then
                If .cmbBahasa.ListIndex = 0 Then
                    MsgBox "Data berhasil disimpan!", vbInformation + vbOKOnly, "Sukses"
                Else
                    MsgBox "Data saved successed!", vbInformation + vbOKOnly, "Success"
                End If
            End If
            If .cekKosongkanInput.Value = Checked Then KosongkanTextBox
            If .cekTutupForm.Value = Checked Then Unload Me
            If .cmbBahasa.ListIndex = 0 Then
                cmBatal.Caption = "&Tutup"
            Else
                cmBatal.Caption = "&Close"
            End If
        End With
End If
Exit Sub
HancurkanError:
    PusatError
End Sub

Private Sub cmBantuan_Click()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Kalimat = App.Path & "\bantuan\html\JejaringSosial.html"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Kalimat = App.Path & "\bantuan\html\SocialNetwork.html"
    End If
    
    If Dir$(Kalimat) <> "" Then
        OpenLocation Kalimat, SHOWNORMAL
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf, file untuk menampilkan petunjuk bantuan tidak ditemukan!" & vbCrLf & _
                    "Silahkan instal ulang aplikasi ini.", vbCritical + vbOKOnly, "Error"
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry, the file to display the help manual can not be found!" & vbCrLf & _
                    "Please reinstall this application.", vbCritical + vbOKOnly, "Error"
        End If
    End If
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmbDataJejaringSosialLalu_Click(Index As Integer)
With textJejaringSosial
    Select Case Index
        Case Is = 0
            .Item(0).Text = cmbDataJejaringSosialLalu(0).Text
            .Item(0).ForeColor = Hitam
            .Item(0).SetFocus
        Case Is = 1
            .Item(1).Text = cmbDataJejaringSosialLalu(1).Text
            .Item(1).ForeColor = Hitam
            .Item(1).SetFocus
        Case Is = 2
            .Item(2).Text = cmbDataJejaringSosialLalu(2).Text
            .Item(2).ForeColor = Hitam
            .Item(2).SetFocus
        Case Is = 3
            .Item(3).Text = cmbDataJejaringSosialLalu(3).Text
            .Item(3).ForeColor = Hitam
            .Item(3).SetFocus
        Case Is = 4
            .Item(4).Text = cmbDataJejaringSosialLalu(4).Text
            .Item(4).ForeColor = Hitam
            .Item(4).SetFocus
        Case Is = 5
            .Item(5).Text = cmbDataJejaringSosialLalu(5).Text
            .Item(5).ForeColor = Hitam
            .Item(5).SetFocus
        Case Is = 6
            .Item(6).Text = cmbDataJejaringSosialLalu(6).Text
            .Item(6).ForeColor = Hitam
            .Item(6).SetFocus
        Case Is = 7
            .Item(7).Text = cmbDataJejaringSosialLalu(7).Text
            .Item(7).ForeColor = Hitam
            .Item(7).SetFocus
    End Select
End With
End Sub

Private Sub cmHapus_Click(Index As Integer)
If FormPengaturan.cmbBahasa.ListIndex = 0 Then
    With textJejaringSosial
        Select Case Index
            Case Is = 0
                If .Item(0).Text = IsiTextBoxKosong_ID(0) Then
                    .Item(0).Text = IsiTextBoxKosong_ID(0)
                Else
                    .Item(0).Text = ""
                    .Item(0).ForeColor = Hitam
                End If
            Case Is = 1
                If .Item(1).Text = IsiTextBoxKosong_ID(1) Then
                    .Item(1).Text = IsiTextBoxKosong_ID(1)
                Else
                    .Item(1).Text = ""
                    .Item(1).ForeColor = Hitam
                End If
            Case Is = 2
                If .Item(2).Text = IsiTextBoxKosong_ID(2) Then
                    .Item(2).Text = IsiTextBoxKosong_ID(2)
                Else
                    .Item(2).Text = ""
                    .Item(2).ForeColor = Hitam
                End If
            Case Is = 3
                If .Item(3).Text = IsiTextBoxKosong_ID(3) Then
                    .Item(3).Text = IsiTextBoxKosong_ID(3)
                Else
                    .Item(3).Text = ""
                    .Item(3).ForeColor = Hitam
                End If
            Case Is = 4
                If .Item(4).Text = IsiTextBoxKosong_ID(4) Then
                    .Item(4).Text = IsiTextBoxKosong_ID(4)
                Else
                    .Item(4).Text = ""
                    .Item(4).ForeColor = Hitam
                End If
            Case Is = 5
                If .Item(5).Text = IsiTextBoxKosong_ID(5) Then
                    .Item(5).Text = IsiTextBoxKosong_ID(5)
                Else
                    .Item(5).Text = ""
                    .Item(5).ForeColor = Hitam
                End If
            Case Is = 6
                If .Item(6).Text = IsiTextBoxKosong_ID(6) Then
                    .Item(6).Text = IsiTextBoxKosong_ID(6)
                Else
                    .Item(6).Text = ""
                    .Item(6).ForeColor = Hitam
                End If
            Case Is = 7
                If .Item(7).Text = IsiTextBoxKosong_ID(7) Then
                    .Item(7).Text = IsiTextBoxKosong_ID(7)
                Else
                    .Item(7).Text = ""
                    .Item(7).ForeColor = Hitam
                End If
        End Select
    End With
Else
    With textJejaringSosial
        Select Case Index
            Case Is = 0
                If .Item(0).Text = IsiTextBoxKosong_EN(0) Then
                    .Item(0).Text = IsiTextBoxKosong_EN(0)
                Else
                    .Item(0).Text = ""
                    .Item(0).ForeColor = Hitam
                End If
            Case Is = 1
                If .Item(1).Text = IsiTextBoxKosong_EN(1) Then
                    .Item(1).Text = IsiTextBoxKosong_EN(1)
                Else
                    .Item(1).Text = ""
                    .Item(1).ForeColor = Hitam
                End If
            Case Is = 2
                If .Item(2).Text = IsiTextBoxKosong_EN(2) Then
                    .Item(2).Text = IsiTextBoxKosong_EN(2)
                Else
                    .Item(2).Text = ""
                    .Item(2).ForeColor = Hitam
                End If
            Case Is = 3
                If .Item(3).Text = IsiTextBoxKosong_EN(3) Then
                    .Item(3).Text = IsiTextBoxKosong_EN(3)
                Else
                    .Item(3).Text = ""
                    .Item(3).ForeColor = Hitam
                End If
            Case Is = 4
                If .Item(4).Text = IsiTextBoxKosong_EN(4) Then
                    .Item(4).Text = IsiTextBoxKosong_EN(4)
                Else
                    .Item(4).Text = ""
                    .Item(4).ForeColor = Hitam
                End If
            Case Is = 5
                If .Item(5).Text = IsiTextBoxKosong_EN(5) Then
                    .Item(5).Text = IsiTextBoxKosong_EN(5)
                Else
                    .Item(5).Text = ""
                    .Item(5).ForeColor = Hitam
                End If
            Case Is = 6
                If .Item(6).Text = IsiTextBoxKosong_EN(6) Then
                    .Item(6).Text = IsiTextBoxKosong_EN(6)
                Else
                    .Item(6).Text = ""
                    .Item(6).ForeColor = Hitam
                End If
            Case Is = 7
                If .Item(7).Text = IsiTextBoxKosong_EN(7) Then
                    .Item(7).Text = IsiTextBoxKosong_EN(7)
                Else
                    .Item(7).Text = ""
                    .Item(7).ForeColor = Hitam
                End If
        End Select
    End With
End If
End Sub

Private Sub cmPergi_Click(Index As Integer)
Select Case Index
    Case Is = 0
        If textJejaringSosial.Item(2).Text = "" Or textJejaringSosial.Item(2).Text = IsiTextBoxKosong_ID(2) Or textJejaringSosial.Item(2).Text = IsiTextBoxKosong_EN(2) Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                MsgBox "Silahkan isi alamat email yang ingin di cek!", vbExclamation + vbOKOnly, ""
            Else
                MsgBox "Put in Email Address for checking!", vbExclamation + vbOKOnly, ""
            End If
                textJejaringSosial.Item(2).SetFocus
        Else
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                AlamatEmail = ShellExecute(0, vbNullString, _
                   "kirim pesan ke :" & textJejaringSosial.Item(2).Text, "", "", vbNormalFocus)
            Else
                AlamatEmail = ShellExecute(0, vbNullString, _
                   "send mail to :" & textJejaringSosial.Item(2).Text, "", "", vbNormalFocus)
            End If
        End If
    Case Is = 1
        If textJejaringSosial.Item(4).Text = "" Or textJejaringSosial.Item(4).Text = IsiTextBoxKosong_ID(4) Or textJejaringSosial.Item(4).Text = IsiTextBoxKosong_EN(4) Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                MsgBox "Silahkan isi alamat URL yang ingin di cek!", vbExclamation + vbOKOnly, ""
            Else
                MsgBox "Put in URL Address for checking!", vbExclamation + vbOKOnly, ""
            End If
                textJejaringSosial.Item(4).SetFocus
        Else
            AlamatSitus = ShellExecute(0, vbNullString, _
                textJejaringSosial.Item(4).Text, "", "", vbNormalFocus)
        End If
    End Select
End Sub

Private Sub cmReset_Click()
    KosongkanTextBox
End Sub
Sub KhususCmSalin()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Tidak dapat disalin karena input masih kosong.", vbExclamation + vbOKOnly, ""
    Else
        MsgBox "Cannot copy because input still be empty.", vbExclamation + vbOKOnly, ""
    End If
End Sub
Private Sub cmSalin_Click(Index As Integer)
If FormPengaturan.cmbBahasa.ListIndex = 0 Then
    With textJejaringSosial
        Select Case Index
            Case Is = 0
                If .Item(0).Text = "" Or .Item(0).Text = IsiTextBoxKosong_ID(0) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(0).SetFocus
                        .Item(0).SelStart = 0
                        .Item(0).SelLength = Len(textJejaringSosial(0).Text)
                        Clipboard.SetText .Item(0).Text
                End If
            Case Is = 1
                If .Item(1).Text = "" Or .Item(1).Text = IsiTextBoxKosong_ID(1) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(1).SetFocus
                        .Item(1).SelStart = 0
                        .Item(1).SelLength = Len(textJejaringSosial(1).Text)
                        Clipboard.SetText .Item(1).Text
                End If
            Case Is = 2
                If .Item(2).Text = "" Or .Item(2).Text = IsiTextBoxKosong_ID(2) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(2).SetFocus
                        .Item(2).SelStart = 0
                        .Item(2).SelLength = Len(textJejaringSosial(2).Text)
                        Clipboard.SetText .Item(2).Text
                End If
            Case Is = 3
                If .Item(3).Text = "" Or .Item(3).Text = IsiTextBoxKosong_ID(3) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(3).SetFocus
                        .Item(3).SelStart = 0
                        .Item(3).SelLength = Len(textJejaringSosial(3).Text)
                        Clipboard.SetText .Item(3).Text
                End If
            Case Is = 4
                If .Item(4).Text = "" Or .Item(4).Text = IsiTextBoxKosong_ID(4) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(4).SetFocus
                        .Item(4).SelStart = 0
                        .Item(4).SelLength = Len(textJejaringSosial(4).Text)
                        Clipboard.SetText .Item(4).Text
                End If
            Case Is = 5
                If .Item(5).Text = "" Or .Item(5).Text = IsiTextBoxKosong_ID(5) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(5).SetFocus
                        .Item(5).SelStart = 0
                        .Item(5).SelLength = Len(textJejaringSosial(5).Text)
                        Clipboard.SetText .Item(5).Text
                End If
            Case Is = 6
                If .Item(6).Text = "" Or .Item(6).Text = IsiTextBoxKosong_ID(6) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(6).SetFocus
                        .Item(6).SelStart = 0
                        .Item(6).SelLength = Len(textJejaringSosial(6).Text)
                        Clipboard.SetText .Item(6).Text
                End If
            Case Is = 7
                If .Item(7).Text = "" Or .Item(7).Text = IsiTextBoxKosong_ID(7) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(7).SetFocus
                        .Item(7).SelStart = 0
                        .Item(7).SelLength = Len(textJejaringSosial(7).Text)
                        Clipboard.SetText .Item(7).Text
                End If
        End Select
    End With
ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
    With textJejaringSosial
        Select Case Index
            Case Is = 0
                If .Item(0).Text = "" Or .Item(0).Text = IsiTextBoxKosong_EN(0) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(0).SetFocus
                        .Item(0).SelStart = 0
                        .Item(0).SelLength = Len(textJejaringSosial(0).Text)
                        Clipboard.SetText .Item(0).Text
                End If
            Case Is = 1
                If .Item(1).Text = "" Or .Item(1).Text = IsiTextBoxKosong_EN(1) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(1).SetFocus
                        .Item(1).SelStart = 0
                        .Item(1).SelLength = Len(textJejaringSosial(1).Text)
                        Clipboard.SetText .Item(1).Text
                End If
            Case Is = 2
                If .Item(2).Text = "" Or .Item(2).Text = IsiTextBoxKosong_EN(2) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(2).SetFocus
                        .Item(2).SelStart = 0
                        .Item(2).SelLength = Len(textJejaringSosial(2).Text)
                        Clipboard.SetText .Item(2).Text
                End If
            Case Is = 3
                If .Item(3).Text = "" Or .Item(3).Text = IsiTextBoxKosong_EN(3) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(3).SetFocus
                        .Item(3).SelStart = 0
                        .Item(3).SelLength = Len(textJejaringSosial(3).Text)
                        Clipboard.SetText .Item(3).Text
                End If
            Case Is = 4
                If .Item(4).Text = "" Or .Item(4).Text = IsiTextBoxKosong_EN(4) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(4).SetFocus
                        .Item(4).SelStart = 0
                        .Item(4).SelLength = Len(textJejaringSosial(4).Text)
                        Clipboard.SetText .Item(4).Text
                End If
            Case Is = 5
                If .Item(5).Text = "" Or .Item(5).Text = IsiTextBoxKosong_EN(5) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(5).SetFocus
                        .Item(5).SelStart = 0
                        .Item(5).SelLength = Len(textJejaringSosial(5).Text)
                        Clipboard.SetText .Item(5).Text
                End If
            Case Is = 6
                If .Item(6).Text = "" Or .Item(6).Text = IsiTextBoxKosong_EN(6) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(6).SetFocus
                        .Item(6).SelStart = 0
                        .Item(6).SelLength = Len(textJejaringSosial(6).Text)
                        Clipboard.SetText .Item(6).Text
                End If
            Case Is = 7
                If .Item(7).Text = "" Or .Item(7).Text = IsiTextBoxKosong_EN(7) Then
                    KhususCmSalin
                Else
                    Clipboard.Clear
                        .Item(7).SetFocus
                        .Item(7).SelStart = 0
                        .Item(7).SelLength = Len(textJejaringSosial(7).Text)
                        Clipboard.SetText .Item(7).Text
                End If
        End Select
    End With
End If
End Sub

Private Sub cmSet_Click()
With FormKalender
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        .Caption = "Set Tanggal"
    Else
        .Caption = "Setting Date"
    End If
    .textTanggal.Text = .Kalender.Day & " - " & .Kalender.Month & " - " & .Kalender.Year
    .Show vbModal, Me
End With
End Sub

Private Sub cmSimpan_Click()
With textJejaringSosial
    Select Case FormPengaturan.cmbBahasa.ListIndex
    Case Is = 0
        If .Item(0).Text = "" Or .Item(0).Text = IsiTextBoxKosong_ID(0) Then
            MsgBox "Mohon isi Nama Jejaring Sosial!" & vbCrLf & _
                    "Jika ingin dikosongkan tambahkan tanda '-' atau klik tombol 'Verifikasi'", vbExclamation + vbOKOnly, ""
            .Item(0).SetFocus
        ElseIf .Item(1).Text = "" Or .Item(1).Text = IsiTextBoxKosong_ID(1) Then
            MsgBox "Mohon isi Nama Pengguna!" & vbCrLf & _
                    "Jika ingin dikosongkan tambahkan tanda '-' atau klik tombol 'Verifikasi'", vbExclamation + vbOKOnly, ""
            .Item(1).SetFocus
        ElseIf .Item(2).Text = "" Or .Item(2).Text = IsiTextBoxKosong_ID(2) Then
            MsgBox "Mohon isi Alamat E-Mail!" & vbCrLf & _
                    "Jika ingin dikosongkan tambahkan tanda '-' atau klik tombol 'Verifikasi'", vbExclamation + vbOKOnly, ""
            .Item(2).SetFocus
        ElseIf .Item(3).Text = "" Or .Item(3).Text = IsiTextBoxKosong_ID(3) Then
            MsgBox "Mohon isi Password Akun!" & vbCrLf & _
                    "Jika ingin dikosongkan tambahkan tanda '-' atau klik tombol 'Verifikasi'", vbExclamation + vbOKOnly, ""
            .Item(3).SetFocus
        ElseIf .Item(4).Text = "" Or .Item(4).Text = IsiTextBoxKosong_ID(4) Then
            MsgBox "Mohon isi URL atau Alamat Website!" & vbCrLf & _
                    "Jika ingin dikosongkan tambahkan tanda '-' atau klik tombol 'Verifikasi'", vbExclamation + vbOKOnly, ""
            .Item(4).SetFocus
        ElseIf .Item(5).Text = "" Or .Item(5).Text = IsiTextBoxKosong_ID(5) Then
            MsgBox "Mohon isi nama Pemilik Akun!" & vbCrLf & _
                    "Jika ingin dikosongkan tambahkan tanda '-' atau klik tombol 'Verifikasi'", vbExclamation + vbOKOnly, ""
            .Item(5).SetFocus
        ElseIf .Item(6).Text = "" Or .Item(6).Text = IsiTextBoxKosong_ID(6) Then
            MsgBox "Mohon isi Tanggal!" & vbCrLf & _
                    "Jika ingin dikosongkan tambahkan tanda '-' atau klik tombol 'Verifikasi'", vbExclamation + vbOKOnly, ""
            .Item(6).SetFocus
        ElseIf .Item(7).Text = "" Or .Item(7).Text = IsiTextBoxKosong_ID(7) Then
            MsgBox "Mohon isi Keterangan!" & vbCrLf & _
                    "Jika ingin dikosongkan tambahkan tanda '-' atau klik tombol 'Verifikasi'", vbExclamation + vbOKOnly, ""
            .Item(7).SetFocus
        Else
            SIMPAN_KE_DATABASE
            IsiCMBDataLalu
        End If
    Case Is = 1
        If .Item(0).Text = "" Or .Item(0).Text = IsiTextBoxKosong_EN(0) Then
            MsgBox "Please put in Social Network Name!" & vbCrLf & _
                    "If you want be clear the entry please insert '-' or click the 'Verify' button", vbExclamation + vbOKOnly, ""
            .Item(0).SetFocus
        ElseIf .Item(1).Text = "" Or .Item(1).Text = IsiTextBoxKosong_EN(1) Then
            MsgBox "Please write User Name!" & vbCrLf & _
                    "If you want be clear the entry please insert '-' or click the 'Verify' button", vbExclamation + vbOKOnly, ""
            .Item(1).SetFocus
        ElseIf .Item(2).Text = "" Or .Item(2).Text = IsiTextBoxKosong_EN(2) Then
            MsgBox "Please write Email Address!" & vbCrLf & _
                    "If you want be clear the entry please insert '-' or click the 'Verify' button", vbExclamation + vbOKOnly, ""
            .Item(2).SetFocus
        ElseIf .Item(3).Text = "" Or .Item(3).Text = IsiTextBoxKosong_EN(3) Then
            MsgBox "Please write Passwords!" & vbCrLf & _
                    "If you want be clear the entry please insert '-' or click the 'Verify' button", vbExclamation + vbOKOnly, ""
            .Item(3).SetFocus
        ElseIf .Item(4).Text = "" Or .Item(4).Text = IsiTextBoxKosong_EN(4) Then
            MsgBox "Please write URL or Web Address!" & vbCrLf & _
                    "If you want be clear the entry please insert '-' or click the 'Verify' button", vbExclamation + vbOKOnly, ""
            .Item(4).SetFocus
        ElseIf .Item(5).Text = "" Or .Item(5).Text = IsiTextBoxKosong_EN(5) Then
            MsgBox "Please write Owner Accounts Name!" & vbCrLf & _
                    "If you want be clear the entry please insert '-' or click the 'Verify' button", vbExclamation + vbOKOnly, ""
            .Item(5).SetFocus
        ElseIf .Item(6).Text = "" Or .Item(6).Text = IsiTextBoxKosong_EN(6) Then
            MsgBox "MPlease write the Date!" & vbCrLf & _
                    "If you want be clear the entry please insert '-' or click the 'Verify' button", vbExclamation + vbOKOnly, ""
            .Item(6).SetFocus
        ElseIf .Item(7).Text = "" Or .Item(7).Text = IsiTextBoxKosong_EN(7) Then
            MsgBox "Please write the Description!" & vbCrLf & _
                    "If you want be clear the entry please insert '-' or click the 'Verify' button", vbExclamation + vbOKOnly, ""
            .Item(7).SetFocus
        Else
            SIMPAN_KE_DATABASE
            IsiCMBDataLalu
        End If
    End Select
End With
End Sub

Private Sub cmVerifikasi_Click()
For NomorIndex = 0 To 7
    textJejaringSosial(NomorIndex).ForeColor = Hitam
Next
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textJejaringSosial
            If .Item(0).Text = "" Or .Item(0).Text = IsiTextBoxKosong_ID(0) Then .Item(0).Text = "-"
            If .Item(1).Text = "" Or .Item(1).Text = IsiTextBoxKosong_ID(1) Then .Item(1).Text = "-"
            If .Item(2).Text = "" Or .Item(2).Text = IsiTextBoxKosong_ID(2) Then .Item(2).Text = "-"
            If .Item(3).Text = "" Or .Item(3).Text = IsiTextBoxKosong_ID(3) Then .Item(3).Text = "-"
            If .Item(4).Text = "" Or .Item(4).Text = IsiTextBoxKosong_ID(4) Then .Item(4).Text = "-"
            If .Item(5).Text = "" Or .Item(5).Text = IsiTextBoxKosong_ID(5) Then .Item(5).Text = "-"
            If .Item(6).Text = "" Or .Item(6).Text = IsiTextBoxKosong_ID(6) Then .Item(6).Text = "-"
            If .Item(7).Text = "" Or .Item(7).Text = IsiTextBoxKosong_ID(7) Then .Item(7).Text = "-"
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textJejaringSosial
            If .Item(0).Text = "" Or .Item(0).Text = IsiTextBoxKosong_EN(0) Then .Item(0).Text = "-"
            If .Item(1).Text = "" Or .Item(1).Text = IsiTextBoxKosong_EN(1) Then .Item(1).Text = "-"
            If .Item(2).Text = "" Or .Item(2).Text = IsiTextBoxKosong_EN(2) Then .Item(2).Text = "-"
            If .Item(3).Text = "" Or .Item(3).Text = IsiTextBoxKosong_EN(3) Then .Item(3).Text = "-"
            If .Item(4).Text = "" Or .Item(4).Text = IsiTextBoxKosong_EN(4) Then .Item(4).Text = "-"
            If .Item(5).Text = "" Or .Item(5).Text = IsiTextBoxKosong_EN(5) Then .Item(5).Text = "-"
            If .Item(6).Text = "" Or .Item(6).Text = IsiTextBoxKosong_EN(6) Then .Item(6).Text = "-"
            If .Item(7).Text = "" Or .Item(7).Text = IsiTextBoxKosong_EN(7) Then .Item(7).Text = "-"
        End With
    End If
End Sub

Private Sub Form_Load()
    AturKontrol
    PENGATURAN_BAHASA
    PENGATURAN_WARNA
End Sub

Private Sub textJejaringSosial_DblClick(Index As Integer)
Select Case Index
    Case Is = 0
       R = SendMessageLong(cmbDataJejaringSosialLalu(0).hwnd, CB_SHOWDROPDOWN, True, 0)
    Case Is = 1
       R = SendMessageLong(cmbDataJejaringSosialLalu(1).hwnd, CB_SHOWDROPDOWN, True, 0)
    Case Is = 2
       R = SendMessageLong(cmbDataJejaringSosialLalu(2).hwnd, CB_SHOWDROPDOWN, True, 0)
    Case Is = 3
       R = SendMessageLong(cmbDataJejaringSosialLalu(3).hwnd, CB_SHOWDROPDOWN, True, 0)
    Case Is = 4
       R = SendMessageLong(cmbDataJejaringSosialLalu(4).hwnd, CB_SHOWDROPDOWN, True, 0)
    Case Is = 5
       R = SendMessageLong(cmbDataJejaringSosialLalu(5).hwnd, CB_SHOWDROPDOWN, True, 0)
    Case Is = 6
       R = SendMessageLong(cmbDataJejaringSosialLalu(6).hwnd, CB_SHOWDROPDOWN, True, 0)
    Case Is = 7
       R = SendMessageLong(cmbDataJejaringSosialLalu(7).hwnd, CB_SHOWDROPDOWN, True, 0)
End Select
End Sub

Private Sub textJejaringSosial_GotFocus(Index As Integer)
With textJejaringSosial
    Select Case Index
        Case Is = 0
            If .Item(0).Text = IsiTextBoxKosong_ID(0) Or .Item(0).Text = IsiTextBoxKosong_EN(0) Then
                .Item(0).Text = ""
                .Item(0).ForeColor = Hitam
            End If
        Case Is = 1
            If .Item(1).Text = IsiTextBoxKosong_ID(1) Or .Item(1).Text = IsiTextBoxKosong_EN(1) Then
                .Item(1).Text = ""
                .Item(1).ForeColor = Hitam
            End If
        Case Is = 2
            If .Item(2).Text = IsiTextBoxKosong_ID(2) Or .Item(2).Text = IsiTextBoxKosong_EN(2) Then
                .Item(2).Text = ""
                .Item(2).ForeColor = Hitam
            End If
        Case Is = 3
            If .Item(3).Text = IsiTextBoxKosong_ID(3) Or .Item(3).Text = IsiTextBoxKosong_EN(3) Then
                .Item(3).Text = ""
                .Item(3).ForeColor = Hitam
           End If
        Case Is = 4
            If .Item(4).Text = IsiTextBoxKosong_ID(4) Or .Item(4).Text = IsiTextBoxKosong_EN(4) Then
                .Item(4).Text = ""
                .Item(4).ForeColor = Hitam
           End If
        Case Is = 5
            If .Item(5).Text = IsiTextBoxKosong_ID(5) Or .Item(5).Text = IsiTextBoxKosong_EN(5) Then
                .Item(5).Text = ""
                .Item(5).ForeColor = Hitam
           End If
        Case Is = 6
            If .Item(6).Text = IsiTextBoxKosong_ID(6) Or .Item(6).Text = IsiTextBoxKosong_EN(6) Then
                .Item(6).Text = ""
                .Item(6).ForeColor = Hitam
            End If
        Case Is = 7
            If .Item(7).Text = IsiTextBoxKosong_ID(7) Or .Item(7).Text = IsiTextBoxKosong_EN(7) Then
                .Item(7).Text = ""
                .Item(7).ForeColor = Hitam
            End If
    End Select
End With
End Sub

Private Sub textJejaringSosial_LostFocus(Index As Integer)
With textJejaringSosial
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Select Case Index
            Case Is = 0
                If .Item(0).Text = "" Then
                    .Item(0).Text = IsiTextBoxKosong_ID(0)
                    .Item(0).ForeColor = SilverTua
                End If
            Case Is = 1
                If .Item(1).Text = "" Then
                    .Item(1).Text = IsiTextBoxKosong_ID(1)
                    .Item(1).ForeColor = SilverTua
                End If
            Case Is = 2
                If .Item(2).Text = "" Then
                    .Item(2).Text = IsiTextBoxKosong_ID(2)
                    .Item(2).ForeColor = SilverTua
                End If
            Case Is = 3
                If .Item(3).Text = "" Then
                    .Item(3).Text = IsiTextBoxKosong_ID(3)
                    .Item(3).ForeColor = SilverTua
                End If
            Case Is = 4
                If .Item(4).Text = "" Then
                    .Item(4).Text = IsiTextBoxKosong_ID(4)
                    .Item(4).ForeColor = SilverTua
                End If
            Case Is = 5
                If .Item(5).Text = "" Then
                    .Item(5).Text = IsiTextBoxKosong_ID(5)
                    .Item(5).ForeColor = SilverTua
                End If
            Case Is = 6
                If .Item(6).Text = "" Then
                    .Item(6).Text = IsiTextBoxKosong_ID(6)
                    .Item(6).ForeColor = SilverTua
                End If
            Case Is = 7
                If .Item(7).Text = "" Then
                    .Item(7).Text = IsiTextBoxKosong_ID(7)
                    .Item(7).ForeColor = SilverTua
                End If
        End Select
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Select Case Index
            Case Is = 0
                If .Item(0).Text = "" Then
                    .Item(0).Text = IsiTextBoxKosong_EN(0)
                    .Item(0).ForeColor = SilverTua
                End If
            Case Is = 1
                If .Item(1).Text = "" Then
                    .Item(1).Text = IsiTextBoxKosong_EN(1)
                    .Item(1).ForeColor = SilverTua
                End If
            Case Is = 2
                If .Item(2).Text = "" Then
                    .Item(2).Text = IsiTextBoxKosong_EN(2)
                    .Item(2).ForeColor = SilverTua
                End If
            Case Is = 3
                If .Item(3).Text = "" Then
                    .Item(3).Text = IsiTextBoxKosong_EN(3)
                    .Item(3).ForeColor = SilverTua
                End If
            Case Is = 4
                If .Item(4).Text = "" Then
                    .Item(4).Text = IsiTextBoxKosong_EN(4)
                    .Item(4).ForeColor = SilverTua
                End If
            Case Is = 5
                If .Item(5).Text = "" Then
                    .Item(5).Text = IsiTextBoxKosong_EN(5)
                    .Item(5).ForeColor = SilverTua
                End If
            Case Is = 6
                If .Item(6).Text = "" Then
                    .Item(6).Text = IsiTextBoxKosong_EN(6)
                    .Item(6).ForeColor = SilverTua
                End If
            Case Is = 7
                If .Item(7).Text = "" Then
                    .Item(7).Text = IsiTextBoxKosong_EN(7)
                    .Item(7).ForeColor = SilverTua
                End If
        End Select
    End If
End With
End Sub

Sub PENGATURAN_WARNA()
    'PENGATURAN WARNA UNTUK FORM INI
    For Each Objek In Me
        Select Case FormPengaturan.cmbWarnaTampilan.ListIndex
        Case Is = 0 'Ungu Natural
            Me.BackColor = UnguNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguNatural
            If TypeName(Objek) = "Frame" Then Objek.BackColor = UnguNatural
        Case Is = 1 'Merah
            Me.BackColor = Merah
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Merah
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Merah
        Case Is = 2 'Pink
            Me.BackColor = Pink
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Pink
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Pink
        Case Is = 3 'HijauMuda
            Me.BackColor = HijauMuda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = HijauMuda
            If TypeName(Objek) = "Frame" Then Objek.BackColor = HijauMuda
        Case Is = 4 'Hitam
            Me.BackColor = Hitam
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Hitam
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Hitam
        Case Is = 5 'Silver
            Me.BackColor = Silver
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Silver
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Silver
        Case Is = 6 'SilverNatural
            Me.BackColor = SilverNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = SilverNatural
            If TypeName(Objek) = "Frame" Then Objek.BackColor = SilverNatural
        Case Is = 7 'Orange
            Me.BackColor = Orange
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Orange
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Orange
        Case Is = 8 'UnguJanda
            Me.BackColor = UnguJanda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguJanda
            If TypeName(Objek) = "Frame" Then Objek.BackColor = UnguJanda
        End Select
    Next
    'PENGATURAN THEMA UNTUK FORM INI
    For Each Objek In Me
        If TypeName(Objek) = "dcButton" Then
            Select Case FormPengaturan.cmbTemaTampilan.ListIndex
            Case Is = 0 'RST_Office 2003
                Objek.ButtonStyle = 3
                Objek.BackColor = &HC0C0C0
            Case Is = 1 'RST_Office XP
                Objek.ButtonStyle = 4
            Case Is = 2 'RST_Opera Browser
                Objek.ButtonStyle = 5
            Case Is = 3 'RST_Classic
                Objek.ButtonStyle = 6
            Case Is = 4 'RST_XP Blue
                Objek.ButtonStyle = 7
            Case Is = 5 'RST_XP Olive Green
                Objek.ButtonStyle = 8
            Case Is = 6 'RST_XP Silver
                Objek.ButtonStyle = 9
            Case Is = 7 'RST_XP Toolbar
                Objek.ButtonStyle = 10
            Case Is = 8 'RST_Yahoo
                Objek.ButtonStyle = 11
                Objek.BackColor = &H12BCFF
            Case Is = 9 'RST_Mac
                Objek.ButtonStyle = 1
                Objek.BackColor = &HFF9B48
            Case Is = 10 'RST_MacOSX
                Objek.ButtonStyle = 2
            End Select
        End If
    Next
End Sub
