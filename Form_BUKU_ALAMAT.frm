VERSION 5.00
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form Form_BUKU_ALAMAT 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "-------------"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_BUKU_ALAMAT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1335
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Height          =   1215
         Left            =   0
         Picture         =   "Form_BUKU_ALAMAT.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   0
         Width           =   7095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6855
      Begin XPControls.XPText textNamaKontak 
         Height          =   330
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   16
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
         TabIndex        =   17
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
         Left            =   6240
         TabIndex        =   18
         Top             =   2400
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
         Left            =   6240
         TabIndex        =   19
         Top             =   2760
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
      Begin XPControls.XPText textNamaPanggilan 
         Height          =   330
         Left            =   2040
         TabIndex        =   20
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
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
      Begin XPControls.XPText textNomorTeleponPribadi 
         Height          =   330
         Left            =   2040
         TabIndex        =   21
         Top             =   960
         Width           =   2895
         _ExtentX        =   5106
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
      Begin XPControls.XPText textNomorTeleponRumah 
         Height          =   330
         Left            =   2040
         TabIndex        =   22
         Top             =   1320
         Width           =   2895
         _ExtentX        =   5106
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
      Begin XPControls.XPText textNomorTeleponKantor 
         Height          =   330
         Left            =   2040
         TabIndex        =   23
         Top             =   1680
         Width           =   2895
         _ExtentX        =   5106
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
      Begin XPControls.XPText textFax 
         Height          =   330
         Left            =   2040
         TabIndex        =   24
         Top             =   2040
         Width           =   2895
         _ExtentX        =   5106
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
      Begin XPControls.XPText textAlamatEmail 
         Height          =   330
         Left            =   2040
         TabIndex        =   25
         Top             =   2400
         Width           =   2895
         _ExtentX        =   5106
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
      Begin XPControls.XPText textWebsite 
         Height          =   330
         Left            =   2040
         TabIndex        =   26
         Top             =   2760
         Width           =   2895
         _ExtentX        =   5106
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
      Begin XPControls.XPText textZIPPostalCode 
         Height          =   330
         Left            =   2040
         TabIndex        =   27
         Top             =   3120
         Width           =   2895
         _ExtentX        =   5106
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
      Begin XPControls.XPText textAlamatRumah 
         Height          =   330
         Left            =   2040
         TabIndex        =   28
         Top             =   3480
         Width           =   2895
         _ExtentX        =   5106
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
         Index           =   8
         Left            =   4965
         TabIndex        =   29
         Top             =   3120
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
         Index           =   9
         Left            =   4965
         TabIndex        =   30
         Top             =   3480
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
         Index           =   8
         Left            =   5610
         TabIndex        =   31
         Top             =   3120
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
         Index           =   9
         Left            =   5610
         TabIndex        =   32
         Top             =   3480
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
      Begin XPControls.XPText textKeterangan 
         Height          =   330
         Left            =   2040
         TabIndex        =   55
         Top             =   3840
         Width           =   2895
         _ExtentX        =   5106
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
         Index           =   10
         Left            =   4965
         TabIndex        =   56
         Top             =   3840
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
         Index           =   10
         Left            =   5610
         TabIndex        =   57
         Top             =   3840
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
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   59
         Top             =   3840
         Width           =   45
      End
      Begin VB.Label Label11 
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
         Left            =   -240
         TabIndex        =   58
         Top             =   3840
         Width           =   1995
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   52
         Top             =   2760
         Width           =   45
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   51
         Top             =   2400
         Width           =   45
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   50
         Top             =   2040
         Width           =   45
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   49
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   48
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label Label110 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   47
         Top             =   960
         Width           =   45
      End
      Begin VB.Label Label100 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   46
         Top             =   600
         Width           =   45
      End
      Begin VB.Label Label90 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   45
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Website"
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
         Left            =   -240
         TabIndex        =   44
         Top             =   2760
         Width           =   1995
      End
      Begin VB.Label Label7 
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
         Left            =   -240
         TabIndex        =   43
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
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
         Left            =   -240
         TabIndex        =   42
         Top             =   2040
         Width           =   1995
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telepon (Kantor)"
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
         Left            =   -240
         TabIndex        =   41
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telepon (Rumah)"
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
         Left            =   -240
         TabIndex        =   40
         Top             =   1320
         Width           =   1995
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telepon (Pribadi)"
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
         Left            =   -240
         TabIndex        =   39
         Top             =   960
         Width           =   1995
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Panggilan"
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
         Left            =   -240
         TabIndex        =   38
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Kontak"
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
         Left            =   -240
         TabIndex        =   37
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   36
         Top             =   3480
         Width           =   45
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   35
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Rumah"
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
         Left            =   -240
         TabIndex        =   34
         Top             =   3480
         Width           =   1995
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ZIP / Postal Code"
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
         Left            =   -240
         TabIndex        =   33
         Top             =   3120
         Width           =   1995
      End
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   0
      Top             =   5400
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
   Begin VB.ComboBox cmbDataLalu11 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   60
      Top             =   5040
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu10 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   61
      Top             =   4680
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu9 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   62
      Top             =   4320
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu8 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   63
      Top             =   3960
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu7 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   64
      Top             =   3600
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu6 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   65
      Top             =   3240
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu5 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   66
      Top             =   2880
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu4 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   67
      Top             =   2520
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu3 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   68
      Top             =   2160
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu2 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   69
      Top             =   1800
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu1 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   70
      Top             =   1440
      Width           =   2895
   End
   Begin Dacara_dcButton.dcButton cmSimpan 
      Height          =   375
      Left            =   120
      TabIndex        =   71
      Top             =   5640
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
      PicDown         =   "Form_BUKU_ALAMAT.frx":1EA84
      PicHot          =   "Form_BUKU_ALAMAT.frx":1EDD6
      PicNormal       =   "Form_BUKU_ALAMAT.frx":1F128
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmReset 
      Height          =   375
      Left            =   1440
      TabIndex        =   72
      Top             =   5640
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
      PicDown         =   "Form_BUKU_ALAMAT.frx":1F47A
      PicHot          =   "Form_BUKU_ALAMAT.frx":1FFC4
      PicNormal       =   "Form_BUKU_ALAMAT.frx":20B0E
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmBatal 
      Height          =   375
      Left            =   5760
      TabIndex        =   73
      Top             =   5640
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
      PicDown         =   "Form_BUKU_ALAMAT.frx":21658
      PicHot          =   "Form_BUKU_ALAMAT.frx":21AAA
      PicNormal       =   "Form_BUKU_ALAMAT.frx":21EFC
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmVerifikasi 
      Height          =   375
      Left            =   2760
      TabIndex        =   74
      Top             =   5640
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
      PicDown         =   "Form_BUKU_ALAMAT.frx":2234E
      PicHot          =   "Form_BUKU_ALAMAT.frx":227A0
      PicNormal       =   "Form_BUKU_ALAMAT.frx":22BF2
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmBantuan 
      Height          =   375
      Left            =   4080
      TabIndex        =   75
      Top             =   5640
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
      PicDown         =   "Form_BUKU_ALAMAT.frx":23044
      PicHot          =   "Form_BUKU_ALAMAT.frx":23496
      PicNormal       =   "Form_BUKU_ALAMAT.frx":238E8
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
Attribute VB_Name = "Form_BUKU_ALAMAT"
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
            .RecordSource = "Select * From tbBukuAlamat order by Nama_Kontak asc;"
            .Refresh
        End With
End Sub
Sub AturKontrol()
    SambungkanKontrolKeADODC_UTAMA
    IsiTextBoxKosong_ID(0) = "(Cth : Rizky Khafitsyah)..."
    IsiTextBoxKosong_ID(1) = "(Cth : RikyMetall)..."
    IsiTextBoxKosong_ID(2) = "(Cth : 081234567890)..."
    IsiTextBoxKosong_ID(3) = "(Cth : 012-345678)..."
    IsiTextBoxKosong_ID(4) = "(Cth : 012-345678)..."
    IsiTextBoxKosong_ID(5) = "(Cth : (+62 12) 345-6789)..."
    IsiTextBoxKosong_ID(6) = "(Cth : email.saya@blabla.com)..."
    IsiTextBoxKosong_ID(7) = "(Cth : http://rikymetalist.blogspot.com)..."
    IsiTextBoxKosong_ID(8) = "(Cth : 20122)..."
    IsiTextBoxKosong_ID(9) = "(Alamat Rumah Anda)..."
    IsiTextBoxKosong_ID(10) = "(Keterangan-Keterangan lain)..."
    IsiTextBoxKosong_EN(0) = "(Eg : Rizky Khafitsyah)..."
    IsiTextBoxKosong_EN(1) = "(Eg : RikyMetall)..."
    IsiTextBoxKosong_EN(2) = "(Eg : 081234567890)..."
    IsiTextBoxKosong_EN(3) = "(Eg : 012-345678)..."
    IsiTextBoxKosong_EN(4) = "(Eg : 012-345678)..."
    IsiTextBoxKosong_EN(5) = "(Eg : (+62 12) 345-6789)..."
    IsiTextBoxKosong_EN(6) = "(Eg : my.email@blabla.com)..."
    IsiTextBoxKosong_EN(7) = "(Eg : http://rikymetalist.blogspot.com)..."
    IsiTextBoxKosong_EN(8) = "(Eg : 20122)..."
    IsiTextBoxKosong_EN(9) = "(Your address)..."
    IsiTextBoxKosong_EN(10) = "(Others description)..."
    For Each Objek In Me
        If TypeName(Objek) = "XPText" Then
            With Objek
                .ForeColor = SilverTua
                .MaxLength = 254
            End With
        End If
    Next
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With Me
            .textNamaKontak.Text = IsiTextBoxKosong_ID(0)
            .textNamaPanggilan.Text = IsiTextBoxKosong_ID(1)
            .textNomorTeleponPribadi.Text = IsiTextBoxKosong_ID(2)
            .textNomorTeleponRumah.Text = IsiTextBoxKosong_ID(3)
            .textNomorTeleponKantor.Text = IsiTextBoxKosong_ID(4)
            .textFax.Text = IsiTextBoxKosong_ID(5)
            .textAlamatEmail.Text = IsiTextBoxKosong_ID(6)
            .textWebsite.Text = IsiTextBoxKosong_ID(7)
            .textZIPPostalCode.Text = IsiTextBoxKosong_ID(8)
            .textAlamatRumah.Text = IsiTextBoxKosong_ID(9)
            .textKeterangan.Text = IsiTextBoxKosong_ID(10)
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With Me
            .textNamaKontak.Text = IsiTextBoxKosong_EN(0)
            .textNamaPanggilan.Text = IsiTextBoxKosong_EN(1)
            .textNomorTeleponPribadi.Text = IsiTextBoxKosong_EN(2)
            .textNomorTeleponRumah.Text = IsiTextBoxKosong_EN(3)
            .textNomorTeleponKantor.Text = IsiTextBoxKosong_EN(4)
            .textFax.Text = IsiTextBoxKosong_EN(5)
            .textAlamatEmail.Text = IsiTextBoxKosong_EN(6)
            .textWebsite.Text = IsiTextBoxKosong_EN(7)
            .textZIPPostalCode.Text = IsiTextBoxKosong_EN(8)
            .textAlamatRumah.Text = IsiTextBoxKosong_EN(9)
            .textKeterangan.Text = IsiTextBoxKosong_EN(10)
        End With
    End If
    IsiCMBDataLalu
    DisableCloseBtn Me
    'PENGATURAN UNTUK ALWAYS ON TOP
    If FormPengaturan.cekAlwaysOnTop.Value = Checked Then
        SetOnTop (Me.hwnd)
    ElseIf FormPengaturan.cekAlwaysOnTop.Value = Unchecked Then
        NotOnTop (Me.hwnd)
    End If
    For Each Objek In Me
        If TypeName(Objek) = "Label" Or TypeName(Objek) = "dcButton" Then
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
    With Me
        .cmbDataLalu1.Clear
        .cmbDataLalu2.Clear
        .cmbDataLalu3.Clear
        .cmbDataLalu4.Clear
        .cmbDataLalu5.Clear
        .cmbDataLalu6.Clear
        .cmbDataLalu7.Clear
        .cmbDataLalu8.Clear
        .cmbDataLalu9.Clear
        .cmbDataLalu10.Clear
        .cmbDataLalu11.Clear
        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
            .cmbDataLalu1.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(0).Value
            .cmbDataLalu2.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(1).Value
            .cmbDataLalu3.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(2).Value
            .cmbDataLalu4.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(3).Value
            .cmbDataLalu5.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(4).Value
            .cmbDataLalu6.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(5).Value
            .cmbDataLalu7.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(6).Value
            .cmbDataLalu8.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(7).Value
            .cmbDataLalu9.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(8).Value
            .cmbDataLalu10.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(9).Value
            .cmbDataLalu11.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(10).Value
            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
        Loop
        FORM_UTAMA.ADODC_UTAMA.Refresh
    End With
End Sub
Sub KhususCmSalin()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Tidak dapat disalin karena input masih kosong.", vbExclamation + vbOKOnly, ""
    Else
        MsgBox "Cannot copy because input still be empty.", vbExclamation + vbOKOnly, ""
    End If
End Sub
Sub PENGATURAN_BAHASA()
If FormPengaturan.cmbBahasa.ListIndex = 0 Then
    With Me
        .Label1.Caption = "Nama Kontak"
        .Label2.Caption = "Nama Panggilan"
        .Label3.Caption = "No. Telepon (Pribadi)"
        .Label4.Caption = "No. Telepon (Rumah)"
        .Label5.Caption = "No. Telepon (Kantor)"
        .Label6.Caption = "Fax"
        .Label7.Caption = "Alamat E-Mail"
        .Label8.Caption = "Website"
        .Label9.Caption = "ZIP/Postal Code"
        .Label10.Caption = "Alamat Rumah"
        .Label11.Caption = "Keterangan"
        For NomorIndex = 0 To 10
            For Each ObjekArray(NomorIndex) In Me
                If TypeName(ObjekArray(NomorIndex)) = "dcButton" Then
                    If ObjekArray(NomorIndex).Caption = "&Copy" Then ObjekArray(NomorIndex).Caption = "&Salin"
                    If ObjekArray(NomorIndex).Caption = "&Delete" Then ObjekArray(NomorIndex).Caption = "&Hapus"
                End If
            Next
        Next
        .cmSimpan.Caption = "&Simpan"
        .cmReset.Caption = "&Reset"
        .cmVerifikasi.Caption = "&Verifikasi"
        .cmBatal.Caption = "&Batal"
        .cmBantuan.Caption = "&Bantuan"
    End With
ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
    With Me
        .Label1.Caption = "Contact Name"
        .Label2.Caption = "Cool Name"
        .Label3.Caption = "Phone (Private)"
        .Label4.Caption = "Phone (Home)"
        .Label5.Caption = "Phone (Office)"
        .Label6.Caption = "Fax"
        .Label7.Caption = "E-Mail Address"
        .Label8.Caption = "Website"
        .Label9.Caption = "ZIP/Postal Code"
        .Label10.Caption = "Home Address"
        .Label11.Caption = "Description"
        For NomorIndex = 0 To 10
            For Each ObjekArray(NomorIndex) In Me
                If TypeName(ObjekArray(NomorIndex)) = "dcButton" Then
                    If ObjekArray(NomorIndex).Caption = "&Salin" Then ObjekArray(NomorIndex).Caption = "&Copy"
                    If ObjekArray(NomorIndex).Caption = "&Hapus" Then ObjekArray(NomorIndex).Caption = "&Delete"
                End If
            Next
        Next
        .cmSimpan.Caption = "&Save"
        .cmReset.Caption = "&Reset"
        .cmVerifikasi.Caption = "&Verify"
        .cmBatal.Caption = "&Cancel"
        .cmBantuan.Caption = "&Help"
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
                If Me.cmSimpan.Caption = "&Simpan" Or Me.cmSimpan.Caption = "&Save" Then
                    With FORM_UTAMA.ADODC_UTAMA
                        .Recordset.AddNew
                        .Recordset.Fields(0).Value = textNamaKontak.Text
                        .Recordset.Fields(1).Value = textNamaPanggilan.Text
                        .Recordset.Fields(2).Value = textNomorTeleponPribadi.Text
                        .Recordset.Fields(3).Value = textNomorTeleponRumah.Text
                        .Recordset.Fields(4).Value = textNomorTeleponKantor.Text
                        .Recordset.Fields(5).Value = textFax.Text
                        .Recordset.Fields(6).Value = textAlamatEmail.Text
                        .Recordset.Fields(7).Value = textWebsite.Text
                        .Recordset.Fields(8).Value = textZIPPostalCode.Text
                        .Recordset.Fields(9).Value = textAlamatRumah.Text
                        .Recordset.Fields(10).Value = textKeterangan.Text
                        .Recordset.Update
                        .Refresh
                    End With
                ElseIf Me.cmSimpan.Caption = "&Perbarui" Or Me.cmSimpan.Caption = "&Update" Then
                    With FormManage.AdodcMain
                        .Recordset.Delete
                        .Recordset.AddNew
                        .Recordset.Fields(0).Value = textNamaKontak.Text
                        .Recordset.Fields(1).Value = textNamaPanggilan.Text
                        .Recordset.Fields(2).Value = textNomorTeleponPribadi.Text
                        .Recordset.Fields(3).Value = textNomorTeleponRumah.Text
                        .Recordset.Fields(4).Value = textNomorTeleponKantor.Text
                        .Recordset.Fields(5).Value = textFax.Text
                        .Recordset.Fields(6).Value = textAlamatEmail.Text
                        .Recordset.Fields(7).Value = textWebsite.Text
                        .Recordset.Fields(8).Value = textZIPPostalCode.Text
                        .Recordset.Fields(9).Value = textAlamatRumah.Text
                        .Recordset.Fields(10).Value = textKeterangan.Text
                        .Recordset.Update
                        .Refresh
                    End With
                    FormManage.AturDatabase
                End If
                With FormPengaturan
                    If .cekAutoRefresh.Value = Checked Then FORM_UTAMA.cmBukuAlamat_Click
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
            If Me.cmSimpan.Caption = "&Simpan" Or Me.cmSimpan.Caption = "&Save" Then
                With FORM_UTAMA.ADODC_UTAMA
                    .Recordset.AddNew
                    .Recordset.Fields(0).Value = textNamaKontak.Text
                    .Recordset.Fields(1).Value = textNamaPanggilan.Text
                    .Recordset.Fields(2).Value = textNomorTeleponPribadi.Text
                    .Recordset.Fields(3).Value = textNomorTeleponRumah.Text
                    .Recordset.Fields(4).Value = textNomorTeleponKantor.Text
                    .Recordset.Fields(5).Value = textFax.Text
                    .Recordset.Fields(6).Value = textAlamatEmail.Text
                    .Recordset.Fields(7).Value = textWebsite.Text
                    .Recordset.Fields(8).Value = textZIPPostalCode.Text
                    .Recordset.Fields(9).Value = textAlamatRumah.Text
                    .Recordset.Fields(10).Value = textKeterangan.Text
                    .Recordset.Update
                    .Refresh
                End With
            ElseIf Me.cmSimpan.Caption = "&Perbarui" Or Me.cmSimpan.Caption = "&Update" Then
                With FormManage.AdodcMain
                    .Recordset.Delete
                    .Recordset.AddNew
                    .Recordset.Fields(0).Value = textNamaKontak.Text
                    .Recordset.Fields(1).Value = textNamaPanggilan.Text
                    .Recordset.Fields(2).Value = textNomorTeleponPribadi.Text
                    .Recordset.Fields(3).Value = textNomorTeleponRumah.Text
                    .Recordset.Fields(4).Value = textNomorTeleponKantor.Text
                    .Recordset.Fields(5).Value = textFax.Text
                    .Recordset.Fields(6).Value = textAlamatEmail.Text
                    .Recordset.Fields(7).Value = textWebsite.Text
                    .Recordset.Fields(8).Value = textZIPPostalCode.Text
                    .Recordset.Fields(9).Value = textAlamatRumah.Text
                    .Recordset.Fields(10).Value = textKeterangan.Text
                    .Recordset.Update
                    .Refresh
                End With
                FormManage.AturDatabase
            End If
            With FormPengaturan
            If .cekAutoRefresh.Value = Checked Then FORM_UTAMA.cmBukuAlamat_Click
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
Sub KosongkanTextBox()
For Each Objek In Me
    If TypeName(Objek) = "XPText" Then
        With Objek
            .MaxLength = 254
            .ForeColor = SilverTua
        End With
    End If
Next
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With Me
            .textNamaKontak.Text = IsiTextBoxKosong_ID(0)
            .textNamaPanggilan.Text = IsiTextBoxKosong_ID(1)
            .textNomorTeleponPribadi.Text = IsiTextBoxKosong_ID(2)
            .textNomorTeleponRumah.Text = IsiTextBoxKosong_ID(3)
            .textNomorTeleponKantor.Text = IsiTextBoxKosong_ID(4)
            .textFax.Text = IsiTextBoxKosong_ID(5)
            .textAlamatEmail.Text = IsiTextBoxKosong_ID(6)
            .textWebsite.Text = IsiTextBoxKosong_ID(7)
            .textZIPPostalCode.Text = IsiTextBoxKosong_ID(8)
            .textAlamatRumah.Text = IsiTextBoxKosong_ID(9)
            .textKeterangan.Text = IsiTextBoxKosong_ID(10)
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With Me
            .textNamaKontak.Text = IsiTextBoxKosong_EN(0)
            .textNamaPanggilan.Text = IsiTextBoxKosong_EN(1)
            .textNomorTeleponPribadi.Text = IsiTextBoxKosong_EN(2)
            .textNomorTeleponRumah.Text = IsiTextBoxKosong_EN(3)
            .textNomorTeleponKantor.Text = IsiTextBoxKosong_EN(4)
            .textFax.Text = IsiTextBoxKosong_EN(5)
            .textAlamatEmail.Text = IsiTextBoxKosong_EN(6)
            .textWebsite.Text = IsiTextBoxKosong_EN(7)
            .textZIPPostalCode.Text = IsiTextBoxKosong_EN(8)
            .textAlamatRumah.Text = IsiTextBoxKosong_EN(9)
            .textKeterangan.Text = IsiTextBoxKosong_EN(10)
        End With
    End If
End Sub


Private Sub cmBantuan_Click()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Kalimat = App.Path & "\bantuan\html\BukuAlamat.html"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Kalimat = App.Path & "\bantuan\html\AddressBook.html"
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

Private Sub cmbDataLalu1_Click()
    With textNamaKontak
        .Text = cmbDataLalu1.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu10_Click()
    With textAlamatRumah
        .Text = cmbDataLalu10.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu11_Click()
    With textKeterangan
        .Text = cmbDataLalu11.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu2_Click()
    With textNamaPanggilan
        .Text = cmbDataLalu2.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu3_Click()
    With textNomorTeleponPribadi
        .Text = cmbDataLalu3.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu4_Click()
    With textNomorTeleponRumah
        .Text = cmbDataLalu4.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu5_Click()
    With textNomorTeleponKantor
        .Text = cmbDataLalu5.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu6_Click()
    With textFax
        .Text = cmbDataLalu6.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu7_Click()
    With textAlamatEmail
        .Text = cmbDataLalu7.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu8_Click()
    With textWebsite
        .Text = cmbDataLalu8.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu9_Click()
    With textZIPPostalCode
        .Text = cmbDataLalu9.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmHapus_Click(Index As Integer)
Select Case Index
    Case Is = 0
        If textNamaKontak.Text = IsiTextBoxKosong_ID(0) Or textNamaKontak.Text = IsiTextBoxKosong_EN(0) Then
            With textNamaKontak
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(0)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(0)
                End If
            End With
        Else
            With textNamaKontak
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 1
        If textNamaPanggilan.Text = IsiTextBoxKosong_ID(1) Or textNamaPanggilan.Text = IsiTextBoxKosong_EN(1) Then
            With textNamaPanggilan
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(1)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(1)
                End If
            End With
        Else
            With textNamaPanggilan
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 2
        If textNomorTeleponPribadi.Text = IsiTextBoxKosong_ID(2) Or textNomorTeleponPribadi.Text = IsiTextBoxKosong_EN(2) Then
            With textNomorTeleponPribadi
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(2)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(2)
                End If
            End With
        Else
            With textNomorTeleponPribadi
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 3
        If textNomorTeleponRumah.Text = IsiTextBoxKosong_ID(3) Or textNomorTeleponRumah.Text = IsiTextBoxKosong_EN(3) Then
            With textNomorTeleponRumah
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(3)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(3)
                End If
            End With
        Else
            With textNomorTeleponRumah
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 4
        If textNomorTeleponKantor.Text = IsiTextBoxKosong_ID(4) Or textNomorTeleponKantor.Text = IsiTextBoxKosong_EN(4) Then
            With textNomorTeleponKantor
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(4)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(4)
                End If
            End With
        Else
            With textNomorTeleponKantor
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 5
        If textFax.Text = IsiTextBoxKosong_ID(5) Or textFax.Text = IsiTextBoxKosong_EN(5) Then
            With textFax
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(5)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(5)
                End If
            End With
        Else
            With textFax
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 6
        If textAlamatEmail.Text = IsiTextBoxKosong_ID(6) Or textAlamatEmail.Text = IsiTextBoxKosong_EN(6) Then
            With textAlamatEmail
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(6)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(6)
                End If
            End With
        Else
            With textAlamatEmail
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 7
        If textWebsite.Text = IsiTextBoxKosong_ID(7) Or textWebsite.Text = IsiTextBoxKosong_EN(7) Then
            With textWebsite
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(7)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(7)
                End If
            End With
        Else
            With textWebsite
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 8
        If textZIPPostalCode.Text = IsiTextBoxKosong_ID(8) Or textZIPPostalCode.Text = IsiTextBoxKosong_EN(8) Then
            With textZIPPostalCode
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(8)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(8)
                End If
            End With
        Else
            With textZIPPostalCode
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 9
        If textAlamatRumah.Text = IsiTextBoxKosong_ID(9) Or textAlamatRumah.Text = IsiTextBoxKosong_EN(9) Then
            With textAlamatRumah
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(9)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(9)
                End If
            End With
        Else
            With textAlamatRumah
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 10
        If textKeterangan.Text = IsiTextBoxKosong_ID(10) Or textKeterangan.Text = IsiTextBoxKosong_EN(10) Then
            With textKeterangan
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(10)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(10)
                End If
            End With
        Else
            With textKeterangan
                .Text = ""
                .SetFocus
            End With
        End If
End Select
End Sub

Private Sub cmPergi_Click(Index As Integer)
Select Case Index
    Case Is = 0
        If textAlamatEmail.Text = "" Or textAlamatEmail.Text = IsiTextBoxKosong_ID(6) Or textAlamatEmail.Text = IsiTextBoxKosong_EN(6) Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                MsgBox "Silahkan isi alamat email yang ingin di cek!", vbExclamation + vbOKOnly, ""
            Else
                MsgBox "Put in Email Address for checking!", vbExclamation + vbOKOnly, ""
            End If
                textAlamatEmail.SetFocus
        Else
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                AlamatEmail = ShellExecute(0, vbNullString, _
                   "kirim pesan ke :" & textAlamatEmail.Text, "", "", vbNormalFocus)
            Else
                AlamatEmail = ShellExecute(0, vbNullString, _
                   "send mail to :" & textAlamatEmail.Text, "", "", vbNormalFocus)
            End If
        End If
    Case Is = 1
        If textWebsite.Text = "" Or textWebsite.Text = IsiTextBoxKosong_ID(7) Or textWebsite.Text = IsiTextBoxKosong_EN(7) Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                MsgBox "Silahkan isi alamat URL yang ingin di cek!", vbExclamation + vbOKOnly, ""
            Else
                MsgBox "Put in URL Address for checking!", vbExclamation + vbOKOnly, ""
            End If
                textWebsite.SetFocus
        Else
            AlamatSitus = ShellExecute(0, vbNullString, _
               textWebsite.Text, "", "", vbNormalFocus)
        End If
    End Select
End Sub

Private Sub cmReset_Click()
    KosongkanTextBox
End Sub

Private Sub cmSalin_Click(Index As Integer)
Select Case Index
    Case Is = 0
        If textNamaKontak.Text = "" Or textNamaKontak.Text = IsiTextBoxKosong_ID(0) Or textNamaKontak.Text = IsiTextBoxKosong_EN(0) Then
            KhususCmSalin
            textNamaKontak.SetFocus
        Else
            With textNamaKontak
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textNamaKontak.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 1
        If textNamaPanggilan.Text = "" Or textNamaPanggilan.Text = IsiTextBoxKosong_ID(1) Or textNamaPanggilan.Text = IsiTextBoxKosong_EN(1) Then
            KhususCmSalin
            textNamaPanggilan.SetFocus
        Else
            With textNamaPanggilan
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textNamaPanggilan.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 2
        If textNomorTeleponPribadi.Text = "" Or textNomorTeleponPribadi.Text = IsiTextBoxKosong_ID(2) Or textNomorTeleponPribadi.Text = IsiTextBoxKosong_EN(2) Then
            KhususCmSalin
            textNomorTeleponPribadi.SetFocus
        Else
            With textNomorTeleponPribadi
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textNomorTeleponPribadi.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 3
        If textNomorTeleponRumah.Text = "" Or textNomorTeleponRumah.Text = IsiTextBoxKosong_ID(3) Or textNomorTeleponRumah.Text = IsiTextBoxKosong_EN(3) Then
            KhususCmSalin
            textNomorTeleponRumah.SetFocus
        Else
            With textNomorTeleponRumah
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textNomorTeleponRumah.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 4
        If textNomorTeleponKantor.Text = "" Or textNomorTeleponKantor.Text = IsiTextBoxKosong_ID(4) Or textNomorTeleponKantor.Text = IsiTextBoxKosong_EN(4) Then
            KhususCmSalin
            textNomorTeleponKantor.SetFocus
        Else
            With textNomorTeleponKantor
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textNomorTeleponKantor.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 5
        If textFax.Text = "" Or textFax.Text = IsiTextBoxKosong_ID(5) Or textFax.Text = IsiTextBoxKosong_EN(5) Then
            KhususCmSalin
            textFax.SetFocus
        Else
            With textFax
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textFax.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 6
        If textAlamatEmail.Text = "" Or textAlamatEmail.Text = IsiTextBoxKosong_ID(6) Or textAlamatEmail.Text = IsiTextBoxKosong_EN(6) Then
            KhususCmSalin
            textAlamatEmail.SetFocus
        Else
            With textAlamatEmail
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textAlamatEmail.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 7
        If textWebsite.Text = "" Or textWebsite.Text = IsiTextBoxKosong_ID(7) Or textWebsite.Text = IsiTextBoxKosong_EN(7) Then
            KhususCmSalin
            textWebsite.SetFocus
        Else
            With textWebsite
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textWebsite.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 8
        If textZIPPostalCode.Text = "" Or textZIPPostalCode.Text = IsiTextBoxKosong_ID(8) Or textZIPPostalCode.Text = IsiTextBoxKosong_EN(8) Then
            KhususCmSalin
            textZIPPostalCode.SetFocus
        Else
            With textZIPPostalCode
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textZIPPostalCode.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 9
        If textAlamatRumah.Text = "" Or textAlamatRumah.Text = IsiTextBoxKosong_ID(9) Or textAlamatRumah.Text = IsiTextBoxKosong_EN(9) Then
            KhususCmSalin
            textAlamatRumah.SetFocus
        Else
            With textAlamatRumah
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textAlamatRumah.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 10
        If textKeterangan.Text = "" Or textKeterangan.Text = IsiTextBoxKosong_ID(10) Or textKeterangan.Text = IsiTextBoxKosong_EN(10) Then
            KhususCmSalin
            textKeterangan.SetFocus
        Else
            With textKeterangan
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textKeterangan.Text)
                Clipboard.SetText .Text
            End With
        End If
    End Select
End Sub

Private Sub cmSimpan_Click()
If textNamaKontak.Text = "" Or textNamaKontak.Text = IsiTextBoxKosong_ID(0) Or textNamaKontak.Text = IsiTextBoxKosong_EN(0) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Nama Kontak!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the Contact Name!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textNamaKontak.SetFocus
ElseIf textNamaPanggilan.Text = "" Or textNamaPanggilan.Text = IsiTextBoxKosong_ID(1) Or textNamaPanggilan.Text = IsiTextBoxKosong_EN(1) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Nama Panggilan!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the Cool Name!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textNamaPanggilan.SetFocus
ElseIf textNomorTeleponPribadi.Text = "" Or textNomorTeleponPribadi.Text = IsiTextBoxKosong_ID(2) Or textNomorTeleponPribadi.Text = IsiTextBoxKosong_EN(2) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Nomor Telepon Pribadi Anda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write your private phone number!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textNomorTeleponPribadi.SetFocus
ElseIf textNomorTeleponRumah.Text = "" Or textNomorTeleponRumah.Text = IsiTextBoxKosong_ID(3) Or textNomorTeleponRumah.Text = IsiTextBoxKosong_EN(3) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Nomor Telepon Rumah Anda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write your home phone number!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textNomorTeleponRumah.SetFocus
ElseIf textNomorTeleponKantor.Text = "" Or textNomorTeleponKantor.Text = IsiTextBoxKosong_ID(4) Or textNomorTeleponKantor.Text = IsiTextBoxKosong_EN(4) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Nomor Telepon Kantor Anda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write your office phone number!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textNomorTeleponKantor.SetFocus
ElseIf textFax.Text = "" Or textFax.Text = IsiTextBoxKosong_ID(5) Or textFax.Text = IsiTextBoxKosong_EN(5) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Nomor Fax Pribadi Anda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write your private fax number!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textFax.SetFocus
ElseIf textAlamatEmail.Text = "" Or textAlamatEmail.Text = IsiTextBoxKosong_ID(6) Or textAlamatEmail.Text = IsiTextBoxKosong_EN(6) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi alamat email Anda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write your email address!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textAlamatEmail.SetFocus
ElseIf textWebsite.Text = "" Or textWebsite.Text = IsiTextBoxKosong_ID(7) Or textWebsite.Text = IsiTextBoxKosong_EN(7) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi alamat website anda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write your website!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textWebsite.SetFocus
ElseIf textZIPPostalCode.Text = "" Or textZIPPostalCode.Text = IsiTextBoxKosong_ID(8) Or textZIPPostalCode.Text = IsiTextBoxKosong_EN(8) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Nomor ZIP/Postal Code daerah Anda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write your zip postal code" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textZIPPostalCode.SetFocus
ElseIf textAlamatRumah.Text = "" Or textAlamatRumah.Text = IsiTextBoxKosong_ID(9) Or textAlamatRumah.Text = IsiTextBoxKosong_EN(9) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Alamat Rumah Anda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write your home address" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textAlamatRumah.SetFocus
ElseIf textKeterangan.Text = "" Or textKeterangan.Text = IsiTextBoxKosong_ID(10) Or textKeterangan.Text = IsiTextBoxKosong_EN(10) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi keterangan tambahan (jika ada)" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the other descriptions" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textKeterangan.SetFocus
Else
    SIMPAN_KE_DATABASE
    IsiCMBDataLalu
End If
End Sub

Private Sub cmVerifikasi_Click()
If textNamaKontak.Text = "" Or textNamaKontak.Text = IsiTextBoxKosong_ID(0) Or textNamaKontak.Text = IsiTextBoxKosong_EN(0) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Pesan = MsgBox("Nama Kontak belum terisi, yakin ingin mem-verifikasi?", vbQuestion + vbYesNo, "Nama Kontak")
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Pesan = MsgBox("Cool Name is Empty!, Are you sure to Verify entry?", vbQuestion + vbYesNo, "Contact Name?")
    End If
        If Pesan = vbYes Then
            For Each Objek In Me
                If TypeName(Objek) = "XPText" Then
                    If Objek.Text = "" Or Objek.ForeColor = SilverTua Then
                        With Objek
                            .Text = "-"
                            .ForeColor = Hitam
                        End With
                    End If
                End If
            Next
        End If
Else
    For Each Objek In Me
        If TypeName(Objek) = "XPText" Then
            If Objek.Text = "" Or Objek.ForeColor = SilverTua Then
                With Objek
                    .Text = "-"
                    .ForeColor = Hitam
                End With
            End If
        End If
    Next
End If
End Sub

Private Sub Form_Load()
    AturKontrol
    PENGATURAN_BAHASA
    PENGATURAN_WARNA
End Sub

Private Sub textNomorTeleponPribadi_DblClick()
       R = SendMessageLong(cmbDataLalu3.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textNomorTeleponPribadi_GotFocus()
If textNomorTeleponPribadi.Text = IsiTextBoxKosong_ID(2) Or textNomorTeleponPribadi.Text = IsiTextBoxKosong_EN(2) Then
    With textNomorTeleponPribadi
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textNomorTeleponPribadi_LostFocus()
If textNomorTeleponPribadi.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textNomorTeleponPribadi
            .Text = IsiTextBoxKosong_ID(2)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textNomorTeleponPribadi
            .Text = IsiTextBoxKosong_EN(2)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub textFax_DblClick()
       R = SendMessageLong(cmbDataLalu6.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textFax_GotFocus()
If textFax.Text = IsiTextBoxKosong_ID(5) Or textFax.Text = IsiTextBoxKosong_EN(5) Then
    With textFax
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textFax_LostFocus()
If textFax.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textFax
            .Text = IsiTextBoxKosong_ID(5)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textFax
            .Text = IsiTextBoxKosong_EN(5)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub textAlamatRumah_DblClick()
       R = SendMessageLong(cmbDataLalu10.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textAlamatRumah_GotFocus()
If textAlamatRumah.Text = IsiTextBoxKosong_ID(9) Or textAlamatRumah.Text = IsiTextBoxKosong_EN(9) Then
    With textAlamatRumah
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textAlamatRumah_LostFocus()
If textAlamatRumah.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textAlamatRumah
            .Text = IsiTextBoxKosong_ID(9)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textAlamatRumah
            .Text = IsiTextBoxKosong_EN(9)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub textNamaPanggilan_DblClick()
       R = SendMessageLong(cmbDataLalu2.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textNamaPanggilan_GotFocus()
If textNamaPanggilan.Text = IsiTextBoxKosong_ID(1) Or textNamaPanggilan.Text = IsiTextBoxKosong_EN(1) Then
    With textNamaPanggilan
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textNamaPanggilan_LostFocus()
If textNamaPanggilan.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textNamaPanggilan
            .Text = IsiTextBoxKosong_ID(1)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textNamaPanggilan
            .Text = IsiTextBoxKosong_EN(1)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub textNamaKontak_DblClick()
       R = SendMessageLong(cmbDataLalu1.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textNamaKontak_GotFocus()
If textNamaKontak.Text = IsiTextBoxKosong_ID(0) Or textNamaKontak.Text = IsiTextBoxKosong_EN(0) Then
    With textNamaKontak
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textNamaKontak_LostFocus()
If textNamaKontak.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textNamaKontak
            .Text = IsiTextBoxKosong_ID(0)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textNamaKontak
            .Text = IsiTextBoxKosong_EN(0)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub textNomorTeleponRumah_DblClick()
       R = SendMessageLong(cmbDataLalu4.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textNomorTeleponRumah_GotFocus()
If textNomorTeleponRumah.Text = IsiTextBoxKosong_ID(3) Or textNomorTeleponRumah.Text = IsiTextBoxKosong_EN(3) Then
    With textNomorTeleponRumah
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textNomorTeleponRumah_LostFocus()
If textNomorTeleponRumah.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textNomorTeleponRumah
            .Text = IsiTextBoxKosong_ID(3)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textNomorTeleponRumah
            .Text = IsiTextBoxKosong_EN(3)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub textWebsite_DblClick()
       R = SendMessageLong(cmbDataLalu8.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textWebsite_GotFocus()
If textWebsite.Text = IsiTextBoxKosong_ID(7) Or textWebsite.Text = IsiTextBoxKosong_EN(7) Then
    With textWebsite
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textWebsite_LostFocus()
If textWebsite.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textWebsite
            .Text = IsiTextBoxKosong_ID(7)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textWebsite
            .Text = IsiTextBoxKosong_EN(7)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub textNomorTeleponKantor_DblClick()
       R = SendMessageLong(cmbDataLalu5.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textNomorTeleponKantor_GotFocus()
If textNomorTeleponKantor.Text = IsiTextBoxKosong_ID(4) Or textNomorTeleponKantor.Text = IsiTextBoxKosong_EN(4) Then
    With textNomorTeleponKantor
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textNomorTeleponKantor_LostFocus()
If textNomorTeleponKantor.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textNomorTeleponKantor
            .Text = IsiTextBoxKosong_ID(4)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textNomorTeleponKantor
            .Text = IsiTextBoxKosong_EN(4)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub textZIPPostalCode_DblClick()
       R = SendMessageLong(cmbDataLalu9.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textZIPPostalCode_GotFocus()
If textZIPPostalCode.Text = IsiTextBoxKosong_ID(8) Or textZIPPostalCode.Text = IsiTextBoxKosong_EN(8) Then
    With textZIPPostalCode
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textZIPPostalCode_LostFocus()
If textZIPPostalCode.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textZIPPostalCode
            .Text = IsiTextBoxKosong_ID(8)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textZIPPostalCode
            .Text = IsiTextBoxKosong_EN(8)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub TextAlamatEmail_DblClick()
       R = SendMessageLong(cmbDataLalu7.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub TextAlamatEmail_GotFocus()
If textAlamatEmail.Text = IsiTextBoxKosong_ID(6) Or textAlamatEmail.Text = IsiTextBoxKosong_EN(6) Then
    With textAlamatEmail
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub TextAlamatEmail_LostFocus()
If textAlamatEmail.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textAlamatEmail
            .Text = IsiTextBoxKosong_ID(6)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textAlamatEmail
            .Text = IsiTextBoxKosong_EN(6)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub textKeterangan_DblClick()
       R = SendMessageLong(cmbDataLalu11.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textKeterangan_GotFocus()
If textKeterangan.Text = IsiTextBoxKosong_ID(10) Or textKeterangan.Text = IsiTextBoxKosong_EN(10) Then
    With textKeterangan
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textKeterangan_LostFocus()
If textKeterangan.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textKeterangan
            .Text = IsiTextBoxKosong_ID(10)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textKeterangan
            .Text = IsiTextBoxKosong_EN(10)
            .ForeColor = SilverTua
        End With
    End If
End If
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
