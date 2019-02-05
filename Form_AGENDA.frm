VERSION 5.00
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form Form_AGENDA 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "----------"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   405
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
   Icon            =   "Form_AGENDA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPEngine.XPControl XP_Engine 
      Left            =   5520
      Top             =   4680
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin Dacara_dcButton.dcButton cmSimpan 
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   12632256
      ButtonStyle     =   3
      Caption         =   "&Simpan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "Form_AGENDA.frx":0442
      PicHot          =   "Form_AGENDA.frx":0794
      PicNormal       =   "Form_AGENDA.frx":0AE6
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmReset 
      Height          =   375
      Left            =   1440
      TabIndex        =   43
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   12632256
      ButtonStyle     =   3
      Caption         =   "&Reset"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "Form_AGENDA.frx":0E38
      PicHot          =   "Form_AGENDA.frx":1982
      PicNormal       =   "Form_AGENDA.frx":24CC
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmBatal 
      Height          =   375
      Left            =   5760
      TabIndex        =   44
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   12632256
      ButtonStyle     =   3
      Caption         =   "&Batal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "Form_AGENDA.frx":3016
      PicHot          =   "Form_AGENDA.frx":3468
      PicNormal       =   "Form_AGENDA.frx":38BA
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmVerifikasi 
      Height          =   375
      Left            =   2760
      TabIndex        =   45
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   12632256
      ButtonStyle     =   3
      Caption         =   "Verifikasi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "Form_AGENDA.frx":3D0C
      PicHot          =   "Form_AGENDA.frx":415E
      PicNormal       =   "Form_AGENDA.frx":45B0
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   0
      Top             =   5280
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
      Height          =   1335
      Left            =   -120
      TabIndex        =   40
      Top             =   -120
      Width           =   7335
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Height          =   1215
         Left            =   120
         Picture         =   "Form_AGENDA.frx":4A02
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   120
         Width           =   7095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6855
      Begin XPControls.XPText textKodeAgenda 
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
         PicAlign        =   3
         PicDown         =   "Form_AGENDA.frx":1DCA4
         PicHot          =   "Form_AGENDA.frx":1DD23
         PicNormal       =   "Form_AGENDA.frx":1DDA2
         PicSizeH        =   14
         PicSizeW        =   11
      End
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   1
         Left            =   4965
         TabIndex        =   3
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
         PicAlign        =   3
         PicDown         =   "Form_AGENDA.frx":1DE21
         PicHot          =   "Form_AGENDA.frx":1DEA0
         PicNormal       =   "Form_AGENDA.frx":1DF1F
         PicSizeH        =   14
         PicSizeW        =   11
      End
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   2
         Left            =   4965
         TabIndex        =   4
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
         PicAlign        =   3
         PicDown         =   "Form_AGENDA.frx":1DF9E
         PicHot          =   "Form_AGENDA.frx":1E01D
         PicNormal       =   "Form_AGENDA.frx":1E09C
         PicSizeH        =   14
         PicSizeW        =   11
      End
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   3
         Left            =   4965
         TabIndex        =   5
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
         PicAlign        =   3
         PicDown         =   "Form_AGENDA.frx":1E11B
         PicHot          =   "Form_AGENDA.frx":1E19A
         PicNormal       =   "Form_AGENDA.frx":1E219
         PicSizeH        =   14
         PicSizeW        =   11
      End
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   4
         Left            =   4965
         TabIndex        =   6
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
         PicAlign        =   3
         PicDown         =   "Form_AGENDA.frx":1E298
         PicHot          =   "Form_AGENDA.frx":1E317
         PicNormal       =   "Form_AGENDA.frx":1E396
         PicSizeH        =   14
         PicSizeW        =   11
      End
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   5
         Left            =   4965
         TabIndex        =   7
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
         PicAlign        =   3
         PicDown         =   "Form_AGENDA.frx":1E415
         PicHot          =   "Form_AGENDA.frx":1E494
         PicNormal       =   "Form_AGENDA.frx":1E513
         PicSizeH        =   14
         PicSizeW        =   11
      End
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   6
         Left            =   4965
         TabIndex        =   8
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
         PicAlign        =   3
         PicDown         =   "Form_AGENDA.frx":1E592
         PicHot          =   "Form_AGENDA.frx":1E611
         PicNormal       =   "Form_AGENDA.frx":1E690
         PicSizeH        =   14
         PicSizeW        =   11
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   0
         Left            =   5610
         TabIndex        =   9
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
         PicAlign        =   3
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   1
         Left            =   5610
         TabIndex        =   10
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
         PicAlign        =   3
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   2
         Left            =   5610
         TabIndex        =   11
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
         PicAlign        =   3
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   3
         Left            =   5610
         TabIndex        =   12
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
         PicAlign        =   3
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   4
         Left            =   5610
         TabIndex        =   13
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
         PicAlign        =   3
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   5
         Left            =   5610
         TabIndex        =   14
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
         PicAlign        =   3
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   6
         Left            =   5610
         TabIndex        =   15
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
         PicAlign        =   3
      End
      Begin XPControls.XPText textNamaAgenda 
         Height          =   330
         Left            =   2040
         TabIndex        =   16
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
      Begin XPControls.XPText textTema 
         Height          =   330
         Left            =   2040
         TabIndex        =   17
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
      Begin XPControls.XPText textTanggal 
         Height          =   330
         Left            =   2040
         TabIndex        =   18
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
      Begin XPControls.XPText textWaktuMulai 
         Height          =   330
         Left            =   2040
         TabIndex        =   19
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
      Begin XPControls.XPText textWaktuAkhir 
         Height          =   330
         Left            =   2040
         TabIndex        =   20
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
      Begin XPControls.XPText textTempat 
         Height          =   330
         Left            =   2040
         TabIndex        =   21
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
      Begin XPControls.XPText textKeterangan 
         Height          =   330
         Left            =   2040
         TabIndex        =   22
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
      Begin Dacara_dcButton.dcButton cmSet 
         Height          =   330
         Left            =   6255
         TabIndex        =   23
         Top             =   1320
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
      Begin Dacara_dcButton.dcButton cmRefreshKode 
         Height          =   330
         Left            =   4965
         TabIndex        =   46
         Top             =   240
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "Refresh Kode"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   3
         PicDown         =   "Form_AGENDA.frx":1E70F
         PicHot          =   "Form_AGENDA.frx":2A776
         PicNormal       =   "Form_AGENDA.frx":367DD
         PicSize         =   1
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   39
         Top             =   2760
         Width           =   45
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   38
         Top             =   2400
         Width           =   45
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   37
         Top             =   2040
         Width           =   45
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   36
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   35
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   34
         Top             =   960
         Width           =   45
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   33
         Top             =   600
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   32
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan Lain"
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
         TabIndex        =   31
         Top             =   2760
         Width           =   1995
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat"
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
         TabIndex        =   30
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Akhir"
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
         TabIndex        =   29
         Top             =   2040
         Width           =   1995
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Mulai"
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
         TabIndex        =   28
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label Label4 
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
         Left            =   -240
         TabIndex        =   27
         Top             =   1320
         Width           =   1995
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tema"
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
         TabIndex        =   26
         Top             =   960
         Width           =   1995
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Agenda"
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
         TabIndex        =   25
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Agenda"
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
         TabIndex        =   24
         Top             =   240
         Width           =   1995
      End
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
      TabIndex        =   48
      Top             =   3960
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
      TabIndex        =   49
      Top             =   3600
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
      TabIndex        =   50
      Top             =   3240
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
      TabIndex        =   51
      Top             =   2880
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
      TabIndex        =   52
      Top             =   2520
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
      TabIndex        =   53
      Top             =   2160
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
      TabIndex        =   54
      Top             =   1800
      Width           =   2895
   End
   Begin Dacara_dcButton.dcButton cmBantuan 
      Height          =   375
      Left            =   4080
      TabIndex        =   55
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   12632256
      ButtonStyle     =   3
      Caption         =   "&Bantuan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "Form_AGENDA.frx":42844
      PicHot          =   "Form_AGENDA.frx":42C96
      PicNormal       =   "Form_AGENDA.frx":430E8
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.Label Label17 
      Height          =   255
      Left            =   5760
      TabIndex        =   47
      Top             =   6480
      Width           =   1095
   End
End
Attribute VB_Name = "Form_AGENDA"
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
            .RecordSource = "Select * From tbAgenda order by Nama_Agenda asc;"
            .Refresh
        End With
End Sub
Sub AturKontrol()
    SambungkanKontrolKeADODC_UTAMA
    IsiTextBoxKosong_ID(0) = "(Cth: Reuni, Arisan dll)..."
    IsiTextBoxKosong_ID(1) = "(Cth: Arisan, Rapat dll)..."
    IsiTextBoxKosong_ID(2) = "(Cth: 01-01-2013 atau klik 'Set')..."
    IsiTextBoxKosong_ID(3) = "(Cth: 20.00 PM)..."
    IsiTextBoxKosong_ID(4) = "(Cth: 20.00 PM)..."
    IsiTextBoxKosong_ID(5) = "(Cth: Gedung Auditorium dll)..."
    IsiTextBoxKosong_ID(6) = "(Keterangan-keterangan lain)..."
    IsiTextBoxKosong_EN(0) = "(Eg: Reunion, RSG dll)..."
    IsiTextBoxKosong_EN(1) = "(Eg: Arisan, Meeting etc)..."
    IsiTextBoxKosong_EN(2) = "(Eg: 01-01-2013 or click 'Set')..."
    IsiTextBoxKosong_EN(3) = "(Eg: 20.00 PM)..."
    IsiTextBoxKosong_EN(4) = "(Eg: 20.00 PM)..."
    IsiTextBoxKosong_EN(5) = "(Eg: Auditorium etc)..."
    IsiTextBoxKosong_EN(6) = "(Others description)..."
    For Each Objek In Me
        If TypeName(Objek) = "XPText" Then
            With Objek
                .ForeColor = Silver
                .MaxLength = 254
            End With
        End If
    Next
    AcakKodeAgenda
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With Me
            .textNamaAgenda.Text = IsiTextBoxKosong_ID(0)
            .textTema.Text = IsiTextBoxKosong_ID(1)
            .textTanggal.Text = IsiTextBoxKosong_ID(2)
            .textWaktuMulai.Text = IsiTextBoxKosong_ID(3)
            .textWaktuAkhir.Text = IsiTextBoxKosong_ID(4)
            .textTempat.Text = IsiTextBoxKosong_ID(5)
            .textKeterangan.Text = IsiTextBoxKosong_ID(6)
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With Me
            .textNamaAgenda.Text = IsiTextBoxKosong_EN(0)
            .textTema.Text = IsiTextBoxKosong_EN(1)
            .textTanggal.Text = IsiTextBoxKosong_EN(2)
            .textWaktuMulai.Text = IsiTextBoxKosong_EN(3)
            .textWaktuAkhir.Text = IsiTextBoxKosong_EN(4)
            .textTempat.Text = IsiTextBoxKosong_EN(5)
            .textKeterangan.Text = IsiTextBoxKosong_EN(6)
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
Sub AcakKodeAgenda()
    With textKodeAgenda
        .BackColor = Me.BackColor
        .ForeColor = Hitam
        .Locked = True
        .Text = Second(Time) & "AG" & Hour(Time) & Val(Day(Date) * 2) & "-" & (2 * Val(Second(Time))) * 2
    End With
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
        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
            .cmbDataLalu1.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(1).Value
            .cmbDataLalu2.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(2).Value
            .cmbDataLalu3.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(3).Value
            .cmbDataLalu4.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(4).Value
            .cmbDataLalu5.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(5).Value
            .cmbDataLalu6.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(6).Value
            .cmbDataLalu7.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(7).Value
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
        .Label1.Caption = "Kode Agenda"
        .Label2.Caption = "Nama Agenda"
        .Label3.Caption = "Tema"
        .Label4.Caption = "Tanggal"
        .Label5.Caption = "Waktu Mulai"
        .Label6.Caption = "Waktu Akhir"
        .Label7.Caption = "Tempat"
        .Label8.Caption = "Keterangan Lain"
        For NomorIndex = 0 To 9
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
        .cmRefreshKode.Caption = "Refresh Kode"
        .cmBantuan.Caption = "Bantuan"
    End With
ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
    With Me
        .Label1.Caption = "Agenda Code"
        .Label2.Caption = "Agenda Name"
        .Label3.Caption = "Thema"
        .Label4.Caption = "Date"
        .Label5.Caption = "Begin Time"
        .Label6.Caption = "End Time"
        .Label7.Caption = "Place"
        .Label8.Caption = "Other Description"
        For NomorIndex = 0 To 9
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
        .cmRefreshKode.Caption = "Refresh Code"
        .cmBantuan.Caption = "Help"
    End With
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
                        .Recordset.Fields(0).Value = textKodeAgenda.Text
                        .Recordset.Fields(1).Value = textNamaAgenda.Text
                        .Recordset.Fields(2).Value = textTema.Text
                        .Recordset.Fields(3).Value = textTanggal.Text
                        .Recordset.Fields(4).Value = textWaktuMulai.Text
                        .Recordset.Fields(5).Value = textWaktuAkhir.Text
                        .Recordset.Fields(6).Value = textTempat.Text
                        .Recordset.Fields(7).Value = textKeterangan.Text
                        .Recordset.Update
                        .Refresh
                    End With
                ElseIf Me.cmSimpan.Caption = "&Perbarui" Or Me.cmSimpan.Caption = "&Update" Then
                    With FormManage.AdodcMain
                        .Recordset.Delete
                        .Recordset.AddNew
                        .Recordset.Fields(0).Value = textKodeAgenda.Text
                        .Recordset.Fields(1).Value = textNamaAgenda.Text
                        .Recordset.Fields(2).Value = textTema.Text
                        .Recordset.Fields(3).Value = textTanggal.Text
                        .Recordset.Fields(4).Value = textWaktuMulai.Text
                        .Recordset.Fields(5).Value = textWaktuAkhir.Text
                        .Recordset.Fields(6).Value = textTempat.Text
                        .Recordset.Fields(7).Value = textKeterangan.Text
                        .Recordset.Update
                        .Refresh
                    End With
                    FormManage.AturDatabase
                End If
                AcakKodeAgenda
                With FormPengaturan
                    If .cekAutoRefresh.Value = Checked Then FORM_UTAMA.cmAgenda_Click
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
                .Recordset.Fields(0).Value = textKodeAgenda.Text
                .Recordset.Fields(1).Value = textNamaAgenda.Text
                .Recordset.Fields(2).Value = textTema.Text
                .Recordset.Fields(3).Value = textTanggal.Text
                .Recordset.Fields(4).Value = textWaktuMulai.Text
                .Recordset.Fields(5).Value = textWaktuAkhir.Text
                .Recordset.Fields(6).Value = textTempat.Text
                .Recordset.Fields(7).Value = textKeterangan.Text
                .Recordset.Update
                .Refresh
            End With
        ElseIf Me.cmSimpan.Caption = "&Perbarui" Or Me.cmSimpan.Caption = "&Update" Then
            With FormManage.AdodcMain
                .Recordset.Delete
                .Recordset.AddNew
                .Recordset.Fields(0).Value = textKodeAgenda.Text
                .Recordset.Fields(1).Value = textNamaAgenda.Text
                .Recordset.Fields(2).Value = textTema.Text
                .Recordset.Fields(3).Value = textTanggal.Text
                .Recordset.Fields(4).Value = textWaktuMulai.Text
                .Recordset.Fields(5).Value = textWaktuAkhir.Text
                .Recordset.Fields(6).Value = textTempat.Text
                .Recordset.Fields(7).Value = textKeterangan.Text
                .Recordset.Update
                .Refresh
            End With
            FormManage.AturDatabase
        End If
        AcakKodeAgenda
        With FormPengaturan
            If .cekAutoRefresh.Value = Checked Then FORM_UTAMA.cmAgenda_Click
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
    'code untuk riwayat pencatatan perogram
    If FormPengaturan.cekRiwayatAktivitas.Value = Checked Then
        With FormRiwayatAktivitas.AdodcUtama
            .Recordset.AddNew
            .Recordset.Fields(0).Value = Day(Date)
            .Recordset.Fields(1).Value = Month(Date)
            .Recordset.Fields(2).Value = Year(Date)
            .Recordset.Fields(3).Value = Hour(Time)
            .Recordset.Fields(4).Value = Minute(Time)
            .Recordset.Fields(5).Value = Second(Time)
            If FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 0 Then
                .Recordset.Fields(6).Value = "Data akun untuk kategori Agenda berhasil disimpan"
            ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                .Recordset.Fields(6).Value = "Account file for Agenda's Category saved successed"
            End If
            .Recordset.Fields(7).Value = GetComputerName
            .Recordset.Update
            .Refresh
            FORM_UTAMA.StatusBawah.Panels.Item(1).Text = .Recordset.Fields(6).Value
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
            .ForeColor = Silver
        End With
    End If
Next
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With Me
            .textNamaAgenda.Text = IsiTextBoxKosong_ID(0)
            .textTema.Text = IsiTextBoxKosong_ID(1)
            .textTanggal.Text = IsiTextBoxKosong_ID(2)
            .textWaktuMulai.Text = IsiTextBoxKosong_ID(3)
            .textWaktuAkhir.Text = IsiTextBoxKosong_ID(4)
            .textTempat.Text = IsiTextBoxKosong_ID(5)
            .textKeterangan.Text = IsiTextBoxKosong_ID(6)
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With Me
            .textNamaAgenda.Text = IsiTextBoxKosong_EN(0)
            .textTema.Text = IsiTextBoxKosong_EN(1)
            .textTanggal.Text = IsiTextBoxKosong_EN(2)
            .textWaktuMulai.Text = IsiTextBoxKosong_EN(3)
            .textWaktuAkhir.Text = IsiTextBoxKosong_EN(4)
            .textTempat.Text = IsiTextBoxKosong_EN(5)
            .textKeterangan.Text = IsiTextBoxKosong_EN(6)
        End With
    End If
    If FormPengaturan.cekRiwayatAktivitas.Value = Checked Then
        With FormRiwayatAktivitas.AdodcUtama
            .Recordset.AddNew
            .Recordset.Fields(0).Value = Day(Date)
            .Recordset.Fields(1).Value = Month(Date)
            .Recordset.Fields(2).Value = Year(Date)
            .Recordset.Fields(3).Value = Hour(Time)
            .Recordset.Fields(4).Value = Minute(Time)
            .Recordset.Fields(5).Value = Second(Time)
            If FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 0 Then
                .Recordset.Fields(6).Value = "Input data akun untuk kategori Agenda direset"
            ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                .Recordset.Fields(6).Value = "Input file for Agenda's Category has clean up"
            End If
            .Recordset.Fields(7).Value = GetComputerName
            .Recordset.Update
            .Refresh
            FORM_UTAMA.StatusBawah.Panels.Item(1).Text = .Recordset.Fields(6).Value
        End With
    End If
End Sub

Private Sub cmBantuan_Click()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Kalimat = App.Path & "\bantuan\html\Agenda.html"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Kalimat = App.Path & "\bantuan\html\Agenda1.html"
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
    With textNamaAgenda
        .Text = cmbDataLalu1.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu2_Click()
    With textTema
        .Text = cmbDataLalu2.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu3_Click()
    With textTanggal
        .Text = cmbDataLalu3.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu4_Click()
    With textWaktuMulai
        .Text = cmbDataLalu4.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu5_Click()
    With textWaktuAkhir
        .Text = cmbDataLalu5.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu6_Click()
    With textTempat
        .Text = cmbDataLalu6.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu7_Click()
    With textKeterangan
        .Text = cmbDataLalu7.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmHapus_Click(Index As Integer)
Select Case Index
    Case Is = 0
        If textNamaAgenda.Text = IsiTextBoxKosong_ID(0) Or textNamaAgenda.Text = IsiTextBoxKosong_EN(0) Then
            With textNamaAgenda
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(0)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(0)
                End If
            End With
        Else
            With textNamaAgenda
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 1
        If textTema.Text = IsiTextBoxKosong_ID(1) Or textTema.Text = IsiTextBoxKosong_EN(1) Then
            With textTema
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(1)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(1)
                End If
            End With
        Else
            With textTema
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 2
        If textTanggal.Text = IsiTextBoxKosong_ID(2) Or textTanggal.Text = IsiTextBoxKosong_EN(2) Then
            With textTanggal
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(2)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(2)
                End If
            End With
        Else
            With textTanggal
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 3
        If textWaktuMulai.Text = IsiTextBoxKosong_ID(3) Or textWaktuMulai.Text = IsiTextBoxKosong_EN(3) Then
            With textWaktuMulai
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(3)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(3)
                End If
            End With
        Else
            With textWaktuMulai
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 4
        If textWaktuAkhir.Text = IsiTextBoxKosong_ID(4) Or textWaktuAkhir.Text = IsiTextBoxKosong_EN(4) Then
            With textWaktuAkhir
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(4)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(4)
                End If
            End With
        Else
            With textWaktuAkhir
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 5
        If textTempat.Text = IsiTextBoxKosong_ID(5) Or textTempat.Text = IsiTextBoxKosong_EN(5) Then
            With textTempat
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(5)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(5)
                End If
            End With
        Else
            With textTempat
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 6
        If textKeterangan.Text = IsiTextBoxKosong_ID(6) Or textKeterangan.Text = IsiTextBoxKosong_EN(6) Then
            With textKeterangan
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(6)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(6)
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

Private Sub cmRefreshKode_Click()
    AcakKodeAgenda
    textNamaAgenda.SetFocus
End Sub

Private Sub cmReset_Click()
    KosongkanTextBox
End Sub

Private Sub cmSalin_Click(Index As Integer)
Select Case Index
    Case Is = 0
        If textNamaAgenda.Text = "" Or textNamaAgenda.Text = IsiTextBoxKosong_ID(0) Or textNamaAgenda.Text = IsiTextBoxKosong_EN(0) Then
            KhususCmSalin
            textNamaAgenda.SetFocus
        Else
            With textNamaAgenda
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textNamaAgenda.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 1
        If textTema.Text = "" Or textTema.Text = IsiTextBoxKosong_ID(1) Or textTema.Text = IsiTextBoxKosong_EN(1) Then
            KhususCmSalin
            textTema.SetFocus
        Else
            With textTema
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textTema.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 2
        If textTanggal.Text = "" Or textTanggal.Text = IsiTextBoxKosong_ID(2) Or textTanggal.Text = IsiTextBoxKosong_EN(2) Then
            KhususCmSalin
            textTanggal.SetFocus
        Else
            With textTanggal
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textTanggal.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 3
        If textWaktuMulai.Text = "" Or textWaktuMulai.Text = IsiTextBoxKosong_ID(3) Or textWaktuMulai.Text = IsiTextBoxKosong_EN(3) Then
            KhususCmSalin
            textWaktuMulai.SetFocus
        Else
            With textWaktuMulai
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textWaktuMulai.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 4
        If textWaktuAkhir.Text = "" Or textWaktuAkhir.Text = IsiTextBoxKosong_ID(4) Or textWaktuAkhir.Text = IsiTextBoxKosong_EN(4) Then
            KhususCmSalin
            textWaktuAkhir.SetFocus
        Else
            With textWaktuAkhir
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textWaktuAkhir.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 5
        If textTempat.Text = "" Or textTempat.Text = IsiTextBoxKosong_ID(5) Or textTempat.Text = IsiTextBoxKosong_EN(5) Then
            KhususCmSalin
            textTempat.SetFocus
        Else
            With textTempat
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textTempat.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 6
        If textKeterangan.Text = "" Or textKeterangan.Text = IsiTextBoxKosong_ID(6) Or textKeterangan.Text = IsiTextBoxKosong_EN(6) Then
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
If textNamaAgenda.Text = "" Or textNamaAgenda.Text = IsiTextBoxKosong_ID(0) Or textNamaAgenda.Text = IsiTextBoxKosong_EN(0) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Nama Agenda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the Agenda's Name!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textNamaAgenda.SetFocus
ElseIf textTema.Text = "" Or textTema.Text = IsiTextBoxKosong_ID(1) Or textTema.Text = IsiTextBoxKosong_EN(1) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Tema Agenda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the Agenda's Thema!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textTema.SetFocus
ElseIf textTanggal.Text = "" Or textTanggal.Text = IsiTextBoxKosong_ID(2) Or textTanggal.Text = IsiTextBoxKosong_EN(2) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Tanggal!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the Agenda's Date!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textTanggal.SetFocus
ElseIf textWaktuMulai.Text = "" Or textWaktuMulai.Text = IsiTextBoxKosong_ID(3) Or textWaktuMulai.Text = IsiTextBoxKosong_EN(3) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Waktu Mulai!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the Begin's Time!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textWaktuMulai.SetFocus
ElseIf textWaktuAkhir.Text = "" Or textWaktuAkhir.Text = IsiTextBoxKosong_ID(4) Or textWaktuAkhir.Text = IsiTextBoxKosong_EN(4) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Waktu Akhir!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write End of Agenda's Time!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textWaktuAkhir.SetFocus
ElseIf textTempat.Text = "" Or textTempat.Text = IsiTextBoxKosong_ID(5) Or textTempat.Text = IsiTextBoxKosong_EN(5) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Lokasi tempat dilaksanakan kegiatan agenda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the location of agenda's activity!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textTempat.SetFocus
ElseIf textKeterangan.Text = "" Or textKeterangan.Text = IsiTextBoxKosong_ID(6) Or textKeterangan.Text = IsiTextBoxKosong_EN(6) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Keterangan!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the Other description!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textKeterangan.SetFocus
Else
    SIMPAN_KE_DATABASE
    IsiCMBDataLalu
End If
End Sub

Private Sub cmVerifikasi_Click()
If textNamaAgenda.Text = "" Or textNamaAgenda.Text = IsiTextBoxKosong_ID(0) Or textNamaAgenda.Text = IsiTextBoxKosong_EN(0) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Pesan = MsgBox("NAma Agenda belum terisi, yakin ingin mem-verifikasi?", vbQuestion + vbYesNo, "Nama Agenda")
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Pesan = MsgBox("Agenda's Name is Empty!, Are you sure to Verify entry?", vbQuestion + vbYesNo, "Agenda's Name?")
    End If
        If Pesan = vbYes Then
            For Each Objek In Me
                If TypeName(Objek) = "XPText" Then
                    If Objek.Text = "" Or Objek.ForeColor = Silver Then
                        With Objek
                            .Text = "-"
                            .ForeColor = Hitam
                        End With
                    End If
                End If
            Next
            AcakKodeAgenda
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
    AcakKodeAgenda
        'kode untuk mencatat program ke program pencatatan
        If FormPengaturan.cekRiwayatAktivitas.Value = Checked Then
        With FormRiwayatAktivitas.AdodcUtama
            .Recordset.AddNew
            .Recordset.Fields(0).Value = Day(Date)
            .Recordset.Fields(1).Value = Month(Date)
            .Recordset.Fields(2).Value = Year(Date)
            .Recordset.Fields(3).Value = Hour(Time)
            .Recordset.Fields(4).Value = Minute(Time)
            .Recordset.Fields(5).Value = Second(Time)
            If FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 0 Then
                .Recordset.Fields(6).Value = "Input data akun untuk kategori Agenda diverifikasi"
            ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                .Recordset.Fields(6).Value = "Input file for Agenda's Category has verified"
            End If
            .Recordset.Fields(7).Value = GetComputerName
            .Recordset.Update
            .Refresh
            FORM_UTAMA.StatusBawah.Panels.Item(1).Text = .Recordset.Fields(6).Value
        End With
    End If
End If
End Sub

Private Sub Form_Load()
    AturKontrol
    PENGATURAN_BAHASA
    PENGATURAN_WARNA
End Sub

Private Sub textTanggal_DblClick()
       R = SendMessageLong(cmbDataLalu3.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textTanggal_GotFocus()
If textTanggal.Text = IsiTextBoxKosong_ID(2) Or textTanggal.Text = IsiTextBoxKosong_EN(2) Then
    With textTanggal
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textTanggal_LostFocus()
If textTanggal.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textTanggal
            .Text = IsiTextBoxKosong_ID(2)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textTanggal
            .Text = IsiTextBoxKosong_EN(2)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub textTempat_DblClick()
       R = SendMessageLong(cmbDataLalu6.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textTempat_GotFocus()
If textTempat.Text = IsiTextBoxKosong_ID(5) Or textTempat.Text = IsiTextBoxKosong_EN(5) Then
    With textTempat
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textTempat_LostFocus()
If textTempat.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textTempat
            .Text = IsiTextBoxKosong_ID(5)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textTempat
            .Text = IsiTextBoxKosong_EN(5)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub textTema_DblClick()
       R = SendMessageLong(cmbDataLalu2.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textTema_GotFocus()
If textTema.Text = IsiTextBoxKosong_ID(1) Or textTema.Text = IsiTextBoxKosong_EN(1) Then
    With textTema
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textTema_LostFocus()
If textTema.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textTema
            .Text = IsiTextBoxKosong_ID(1)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textTema
            .Text = IsiTextBoxKosong_EN(1)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub textNamaAgenda_DblClick()
       R = SendMessageLong(cmbDataLalu1.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textNamaAgenda_GotFocus()
If textNamaAgenda.Text = IsiTextBoxKosong_ID(0) Or textNamaAgenda.Text = IsiTextBoxKosong_EN(0) Then
    With textNamaAgenda
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textNamaAgenda_LostFocus()
If textNamaAgenda.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textNamaAgenda
            .Text = IsiTextBoxKosong_ID(0)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textNamaAgenda
            .Text = IsiTextBoxKosong_EN(0)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub textWaktuMulai_DblClick()
       R = SendMessageLong(cmbDataLalu4.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textWaktuMulai_GotFocus()
If textWaktuMulai.Text = IsiTextBoxKosong_ID(3) Or textWaktuMulai.Text = IsiTextBoxKosong_EN(3) Then
    With textWaktuMulai
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textWaktuMulai_LostFocus()
If textWaktuMulai.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textWaktuMulai
            .Text = IsiTextBoxKosong_ID(3)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textWaktuMulai
            .Text = IsiTextBoxKosong_EN(3)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub textWaktuAkhir_DblClick()
       R = SendMessageLong(cmbDataLalu5.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textWaktuAkhir_GotFocus()
If textWaktuAkhir.Text = IsiTextBoxKosong_ID(4) Or textWaktuAkhir.Text = IsiTextBoxKosong_EN(4) Then
    With textWaktuAkhir
        .Text = ""
        .ForeColor = Hitam
    End With
End If
End Sub

Private Sub textWaktuAkhir_LostFocus()
If textWaktuAkhir.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With textWaktuAkhir
            .Text = IsiTextBoxKosong_ID(4)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textWaktuAkhir
            .Text = IsiTextBoxKosong_EN(4)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub

Private Sub textKeterangan_DblClick()
       R = SendMessageLong(cmbDataLalu7.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textKeterangan_GotFocus()
If textKeterangan.Text = IsiTextBoxKosong_ID(6) Or textKeterangan.Text = IsiTextBoxKosong_EN(6) Then
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
            .Text = IsiTextBoxKosong_ID(6)
            .ForeColor = SilverTua
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With textKeterangan
            .Text = IsiTextBoxKosong_EN(6)
            .ForeColor = SilverTua
        End With
    End If
End If
End Sub



