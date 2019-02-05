VERSION 5.00
Object = "{02353968-C1C9-4E0A-88D3-18759BDC60FE}#1.0#0"; "AeroSuite.ocx"
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{A30DC858-670B-4336-A74E-10C38ADF5ADD}#1.0#0"; "xTab.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormPengaturan 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pengaturan"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPengaturan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjXTab.XTab XTab1 
      Height          =   3975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7011
      TabCaption(0)   =   "Umum"
      TabContCtrlCnt(0)=   1
      Tab(0)ContCtrlCap(1)=   "XPFrame2"
      TabCaption(1)   =   "System"
      TabContCtrlCnt(1)=   1
      Tab(1)ContCtrlCap(1)=   "XPFrame1"
      TabCaption(2)   =   "Penampilan"
      TabContCtrlCnt(2)=   1
      Tab(2)ContCtrlCap(1)=   "XPFrame3"
      ActiveTab       =   2
      TabTheme        =   1
      ActiveTabBackStartColor=   16777215
      ActiveTabBackEndColor=   -2147483626
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      Begin XPControls.XPFrame XPFrame3 
         Height          =   3375
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5953
         BackColor       =   14737632
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cmbWarnaTampilan 
            Height          =   390
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   600
            Width           =   2775
         End
         Begin VB.ComboBox cmbTemaTampilan 
            Height          =   390
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   120
            Width           =   2775
         End
         Begin AeroSuite.AeroCheckBox cekAlwaysOnTop 
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1200
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   450
            Align           =   0
            Caption         =   "Tampilan aplikasi selalu berada diatas"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   14737632
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "FormPengaturan.frx":0442
            Value           =   0
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   2160
            TabIndex        =   36
            Top             =   600
            Width           =   45
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Warna"
            Height          =   270
            Left            =   120
            TabIndex        =   35
            Top             =   600
            Width           =   555
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   2160
            TabIndex        =   33
            Top             =   120
            Width           =   45
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tema"
            Height          =   270
            Left            =   120
            TabIndex        =   32
            Top             =   120
            Width           =   465
         End
      End
      Begin XPControls.XPFrame XPFrame2 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5953
         BackColor       =   14737632
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin AeroSuite.AeroCheckBox cekAutoRefresh 
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   450
            Align           =   0
            Caption         =   "Auto Refresh"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   14737632
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "FormPengaturan.frx":045E
            Value           =   1
         End
         Begin AeroSuite.AeroCheckBox cekGarisGrid 
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   450
            Align           =   0
            Caption         =   "Garis Grid"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   14737632
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "FormPengaturan.frx":047A
            Value           =   1
         End
         Begin AeroSuite.AeroCheckBox cekTampilkanPesanSimpan 
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1440
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   450
            Align           =   0
            Caption         =   "Tampilkan Pesan saat menyimpan data"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   14737632
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "FormPengaturan.frx":0496
            Value           =   1
         End
         Begin AeroSuite.AeroCheckBox cekKosongkanInput 
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   1800
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   450
            Align           =   0
            Caption         =   "Kosongkan input saat berhasil menyimpan"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   14737632
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "FormPengaturan.frx":04B2
            Value           =   1
         End
         Begin AeroSuite.AeroCheckBox cekPesanKonfirmasi 
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   1080
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   450
            Align           =   0
            Caption         =   "Konfirmasi Saat data akan disimpan"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   14737632
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "FormPengaturan.frx":04CE
            Value           =   1
         End
         Begin AeroSuite.AeroCheckBox CekTutupForm 
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   2160
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   450
            Align           =   0
            Caption         =   "Tutup jendela saat data berhasil disimpan"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   14737632
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "FormPengaturan.frx":04EA
            Value           =   0
         End
         Begin AeroSuite.AeroCheckBox cekKunciTabel 
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   2520
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   450
            Align           =   0
            Caption         =   "Kunci Tabel"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   14737632
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "FormPengaturan.frx":0506
            Value           =   1
         End
         Begin AeroSuite.AeroCheckBox cekRiwayatAktivitas 
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   2880
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            Align           =   0
            Caption         =   "Catat Aktivitas Pengguna"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   14737632
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "FormPengaturan.frx":0522
            Value           =   1
         End
      End
      Begin XPControls.XPFrame XPFrame1 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5953
         BackColor       =   14737632
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cmbHasilPencarian 
            Height          =   390
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   2160
            Width           =   2775
         End
         Begin VB.ComboBox CmbDefaultTampilkanData 
            Height          =   390
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   2775
         End
         Begin VB.ComboBox cmbReferensiExcel 
            Height          =   390
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1680
            Width           =   2775
         End
         Begin VB.ComboBox cmbDefaultFormat 
            Height          =   390
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1200
            Width           =   2775
         End
         Begin VB.ComboBox cmbBahasa 
            Height          =   390
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   2175
         End
         Begin AeroSuite.AeroCheckBox CekTutupFormCAri 
            Height          =   735
            Left            =   2280
            TabIndex        =   27
            Top             =   2340
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   1296
            Align           =   0
            Caption         =   "Tutup Pencarian Saat Data Ditemukan"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Agency FB"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   14737632
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "FormPengaturan.frx":053E
            Value           =   1
         End
         Begin Dacara_dcButton.dcButton cmSID 
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   2880
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            BackColor       =   12230304
            ButtonStyle     =   3
            Caption         =   "Set Internal Databases"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Agency FB"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PicAlign        =   3
            PicDown         =   "FormPengaturan.frx":055A
            PicHot          =   "FormPengaturan.frx":28DC
            PicNormal       =   "FormPengaturan.frx":4C5E
            PicSize         =   2
            PicSizeH        =   24
            PicSizeW        =   24
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   2160
            TabIndex        =   26
            Top             =   2160
            Width           =   45
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasil Pencarian"
            Height          =   270
            Left            =   120
            TabIndex        =   25
            Top             =   2160
            Width           =   1245
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Tampil Data"
            Height          =   270
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1620
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   2160
            TabIndex        =   14
            Top             =   720
            Width           =   45
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referensi Excel"
            Height          =   270
            Left            =   120
            TabIndex        =   13
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   2160
            TabIndex        =   12
            Top             =   1680
            Width           =   45
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   2160
            TabIndex        =   11
            Top             =   1200
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Format Default"
            Height          =   270
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   270
            Left            =   2160
            TabIndex        =   9
            Top             =   240
            Width           =   45
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bahasa"
            Height          =   270
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   555
         End
      End
   End
   Begin Dacara_dcButton.dcButton cmOK 
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   5
      PicDown         =   "FormPengaturan.frx":6FE0
      PicHot          =   "FormPengaturan.frx":72FA
      PicNormal       =   "FormPengaturan.frx":7614
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmBatal 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Batal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FormPengaturan.frx":792E
      PicHot          =   "FormPengaturan.frx":7D80
      PicNormal       =   "FormPengaturan.frx":81D2
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   480
      Top             =   4680
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   2
      Bmp:1           =   "FormPengaturan.frx":8624
      Mask:1          =   6052895
      Key:1           =   "#MenuRegistry"
      Bmp:2           =   "FormPengaturan.frx":8976
      Mask:2          =   16777215
      Key:2           =   "#menuPF"
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
   Begin Dacara_dcButton.dcButton cmBantuan 
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Bantuan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   5
      PicDown         =   "FormPengaturan.frx":8CC8
      PicHot          =   "FormPengaturan.frx":911A
      PicNormal       =   "FormPengaturan.frx":956C
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
Attribute VB_Name = "FormPengaturan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
        With cmbBahasa
            .Clear
            .AddItem "Indonesia", 0
            .AddItem "Inggris", 1
            .ListIndex = GetSetting("rssamv1.0", "pengaturanUtama", cmbBahasa.Name, cmbBahasa.ListIndex)
        End With
        If cmbBahasa.ListIndex = 0 Then
            With CmbDefaultTampilkanData
                .Clear
                .AddItem "Identitas Pribadi", 0
                .AddItem "Buku Alamat", 1
                .AddItem "Ulang Tahun", 2
                .AddItem "Agenda", 3
                .AddItem "Registrasi Software", 4
                .AddItem "Jejaring Sosial", 5
                .AddItem "Electronic Mail", 6
                .AddItem "Forum Internet", 7
                .AddItem "File Transfer Protokol", 8
                .AddItem "Blogging/Website", 9
                .ListIndex = 0
            End With
        ElseIf cmbBahasa.ListIndex = 1 Then
            With CmbDefaultTampilkanData
                .Clear
                .AddItem "Personal Biodata", 0
                .AddItem "Address Book", 1
                .AddItem "Birthday", 2
                .AddItem "Agenda", 3
                .AddItem "Software Serial", 4
                .AddItem "Social Network", 5
                .AddItem "Electronic Mail", 6
                .AddItem "Internet Forums", 7
                .AddItem "FTP", 8
                .AddItem "Blogging/Site", 9
                .ListIndex = 0
            End With
        End If
        If cmbBahasa.ListIndex = 0 Then
            With cmbHasilPencarian
                .Clear
                .AddItem "Kotak Text", 0
                .AddItem "Kotak Pesan", 1
                .AddItem "Arahkan Ke Temuan", 2
                .ListIndex = 0
            End With
        ElseIf cmbBahasa.ListIndex = 1 Then
            With cmbHasilPencarian
                .Clear
                .AddItem "Text Box", 0
                .AddItem "Messages Box", 1
                .AddItem "Point to Result", 2
                .ListIndex = 0
            End With
        End If
    With cmbDefaultFormat
        .Clear
        .AddItem "RikySoft Catatan Files (*.rcf)", 0
        .AddItem "RikySoft Mail Files (*.rmd)", 1
        .AddItem "Text Files (*.txt)", 2
        .AddItem "Word 2003 Files (*.doc)", 3
        .AddItem "Excel 2003 Files (*.xls)", 4
        .AddItem "HTML Files (*.htm)", 5
        .ListIndex = 0
    End With
    With cmbReferensiExcel
        .Clear
        .AddItem "Microsoft Excel 11.0 Object Library", 0
        .ListIndex = 0
    End With
    With cmbTemaTampilan
        .Clear
        .AddItem "RST_Office 2003", 0
        .AddItem "RST_Office XP", 1
        .AddItem "RST_Opera Browser", 2
        .AddItem "RST_Classic", 3
        .AddItem "RST_XP Blue", 4
        .AddItem "RST_XP Olive Green", 5
        .AddItem "RST_XP Silver", 6
        .AddItem "RST_XP Toolbar", 7
        .AddItem "RST_Yahoo", 8
        .AddItem "RST_Mac", 9
        .AddItem "RST_MacOSX", 10
        .ListIndex = 0
    End With
    With cmbWarnaTampilan
        .Clear
        .AddItem "Ungu Natural", 0
        .AddItem "Merah", 1
        .AddItem "Pink", 2
        .AddItem "Hijau Muda", 3
        .AddItem "Hitam", 4
        .AddItem "Silver", 5
        .AddItem "Silver Natural", 6
        .AddItem "Orange", 7
        .AddItem "Ungu Janda", 8
        .ListIndex = 0
    End With
    AmbilPengaturanUtama
    PengaturanFormIni
    DisableCloseBtn Me
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

Sub PengaturanFORM_UTAMA()
    'PENGATURAN WARNA UNTUK FORM_UTAMA
    For Each Objek In FORM_UTAMA
        Select Case cmbWarnaTampilan.ListIndex
        Case Is = 0
            FORM_UTAMA.BackColor = UnguNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguNatural
            If TypeName(Objek) = "Frame" Then Objek.BackColor = UnguNatural
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = UnguNatural
        Case Is = 1
            FORM_UTAMA.BackColor = Merah
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Merah
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Merah
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = Merah
        Case Is = 2
            FORM_UTAMA.BackColor = Pink
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Pink
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Pink
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = Pink
        Case Is = 3
            FORM_UTAMA.BackColor = HijauMuda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = HijauMuda
            If TypeName(Objek) = "Frame" Then Objek.BackColor = HijauMuda
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = HijauMuda
        Case Is = 4
            FORM_UTAMA.BackColor = Hitam
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Hitam
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Hitam
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = Hitam
        Case Is = 5
            FORM_UTAMA.BackColor = Silver
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Silver
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Silver
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = Silver
        Case Is = 6
            FORM_UTAMA.BackColor = SilverNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = SilverFormUtama
            If TypeName(Objek) = "Frame" Then Objek.BackColor = SilverNatural
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = SilverNatural
            FORM_UTAMA.cmAkunDesktop.BackColor = &H808080
            FORM_UTAMA.cmAkunWeb.BackColor = &H808080
        Case Is = 7
            FORM_UTAMA.BackColor = Orange
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Orange
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Orange
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = Orange
        Case Is = 8
            FORM_UTAMA.BackColor = UnguJanda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguJanda
            If TypeName(Objek) = "Frame" Then Objek.BackColor = UnguJanda
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = UnguJanda
        End Select
    Next
    'PENGATURAN THEMA UNTUK FORM_UTAMA
    For Each Objek In FORM_UTAMA
        If TypeName(Objek) = "dcButton" Then
            Select Case cmbTemaTampilan.ListIndex
            Case Is = 0
                Objek.ButtonStyle = 3
            Case Is = 1
                Objek.ButtonStyle = 4
            Case Is = 2
                Objek.ButtonStyle = 5
            Case Is = 3
                Objek.ButtonStyle = 6
            Case Is = 4
                Objek.ButtonStyle = 7
            Case Is = 5
                Objek.ButtonStyle = 8
            Case Is = 6
                Objek.ButtonStyle = 9
            Case Is = 7
                Objek.ButtonStyle = 10
            Case Is = 8
                Objek.ButtonStyle = 11
                Objek.BackColor = &H12BCFF
            Case Is = 9
                Objek.ButtonStyle = 1
                Objek.BackColor = &HFF9B48
            Case Is = 10
                Objek.ButtonStyle = 2
            End Select
        End If
    Next
    If Me.cekGarisGrid.Value = Checked Then
        FORM_UTAMA.LV.GridLines = True
    ElseIf Me.cekGarisGrid.Value = Unchecked Then
        FORM_UTAMA.LV.GridLines = False
    End If
    'always on top
    If FormPengaturan.cekAlwaysOnTop.Value = Checked Then
        SetOnTop (FORM_UTAMA.hwnd)
    ElseIf FormPengaturan.cekAlwaysOnTop.Value = Unchecked Then
        NotOnTop (FORM_UTAMA.hwnd)
    End If
End Sub
Sub SimpanPengaturanUtama()
    With Me
        SaveSetting "rssamv1.0", "pengaturanUtama", .cmbWarnaTampilan.Name, .cmbWarnaTampilan.ListIndex
        SaveSetting "rssamv1.0", "pengaturanUtama", .cmbTemaTampilan.Name, .cmbTemaTampilan.ListIndex
        SaveSetting "rssamv1.0", "pengaturanUtama", .cekAutoRefresh.Name, .cekAutoRefresh.Value
        SaveSetting "rssamv1.0", "pengaturanUtama", .cekGarisGrid.Name, .cekGarisGrid.Value
        SaveSetting "rssamv1.0", "pengaturanUtama", .cekTampilkanPesanSimpan.Name, .cekTampilkanPesanSimpan.Value
        SaveSetting "rssamv1.0", "pengaturanUtama", .cekKosongkanInput.Name, .cekKosongkanInput.Value
        SaveSetting "rssamv1.0", "pengaturanUtama", .cekPesanKonfirmasi.Name, .cekPesanKonfirmasi.Value
        SaveSetting "rssamv1.0", "pengaturanUtama", .cekTutupForm.Name, .cekTutupForm.Value
        SaveSetting "rssamv1.0", "pengaturanUtama", .cekKunciTabel.Name, .cekKunciTabel.Value
        SaveSetting "rssamv1.0", "pengaturanUtama", .cekRiwayatAktivitas.Name, .cekRiwayatAktivitas.Value
        SaveSetting "rssamv1.0", "pengaturanUtama", .cmbHasilPencarian.Name, .cmbHasilPencarian.ListIndex
        SaveSetting "rssamv1.0", "pengaturanUtama", .CmbDefaultTampilkanData.Name, .CmbDefaultTampilkanData.ListIndex
        SaveSetting "rssamv1.0", "pengaturanUtama", .cmbReferensiExcel.Name, .cmbReferensiExcel.ListIndex
        SaveSetting "rssamv1.0", "pengaturanUtama", .cmbDefaultFormat.Name, .cmbDefaultFormat.ListIndex
        SaveSetting "rssamv1.0", "pengaturanUtama", .cmbBahasa.Name, .cmbBahasa.ListIndex
        SaveSetting "rssamv1.0", "pengaturanUtama", .CekTutupFormCAri.Name, .CekTutupFormCAri.Value
        SaveSetting "rssamv1.0", "pengaturanUtama", .XTab1.Name, .XTab1.ActiveTab
        SaveSetting "rssamv1.0", "pengaturanUtama", .cekAlwaysOnTop.Name, .cekAlwaysOnTop.Value
    End With
End Sub
Sub AmbilPengaturanUtama()
    With Me
        .cmbWarnaTampilan.ListIndex = GetSetting("rssamv1.0", "pengaturanUtama", .cmbWarnaTampilan.Name, .cmbWarnaTampilan.ListIndex)
        .cmbTemaTampilan.ListIndex = GetSetting("rssamv1.0", "pengaturanUtama", .cmbTemaTampilan.Name, .cmbTemaTampilan.ListIndex)
        .cekAutoRefresh.Value = GetSetting("rssamv1.0", "pengaturanUtama", .cekAutoRefresh.Name, .cekAutoRefresh.Value)
        .cekGarisGrid.Value = GetSetting("rssamv1.0", "pengaturanUtama", .cekGarisGrid.Name, .cekGarisGrid.Value)
        .cekTampilkanPesanSimpan.Value = GetSetting("rssamv1.0", "pengaturanUtama", .cekTampilkanPesanSimpan.Name, .cekTampilkanPesanSimpan.Value)
        .cekKosongkanInput.Value = GetSetting("rssamv1.0", "pengaturanUtama", .cekKosongkanInput.Name, .cekKosongkanInput.Value)
        .cekPesanKonfirmasi.Value = GetSetting("rssamv1.0", "pengaturanUtama", .cekPesanKonfirmasi.Name, .cekPesanKonfirmasi.Value)
        .cekTutupForm.Value = GetSetting("rssamv1.0", "pengaturanUtama", .cekTutupForm.Name, .cekTutupForm.Value)
        .cekKunciTabel.Value = GetSetting("rssamv1.0", "pengaturanUtama", .cekKunciTabel.Name, .cekKunciTabel.Value)
        .cekRiwayatAktivitas.Value = GetSetting("rssamv1.0", "pengaturanUtama", .cekRiwayatAktivitas.Name, .cekRiwayatAktivitas.Value)
        .cmbHasilPencarian.ListIndex = GetSetting("rssamv1.0", "pengaturanUtama", .cmbHasilPencarian.Name, .cmbHasilPencarian.ListIndex)
        .CmbDefaultTampilkanData.ListIndex = GetSetting("rssamv1.0", "pengaturanUtama", .CmbDefaultTampilkanData.Name, .CmbDefaultTampilkanData.ListIndex)
        .cmbReferensiExcel.ListIndex = GetSetting("rssamv1.0", "pengaturanUtama", .cmbReferensiExcel.Name, .cmbReferensiExcel.ListIndex)
        .cmbDefaultFormat.ListIndex = GetSetting("rssamv1.0", "pengaturanUtama", .cmbDefaultFormat.Name, .cmbDefaultFormat.ListIndex)
        .cmbBahasa.ListIndex = GetSetting("rssamv1.0", "pengaturanUtama", .cmbBahasa.Name, .cmbBahasa.ListIndex)
        .CekTutupFormCAri.Value = GetSetting("rssamv1.0", "pengaturanUtama", .CekTutupFormCAri.Name, .CekTutupFormCAri.Value)
        .XTab1.ActiveTab = GetSetting("rssamv1.0", "pengaturanUtama", .XTab1.Name, .XTab1.ActiveTab)
        .cekAlwaysOnTop.Value = GetSetting("rssamv1.0", "pengaturanUtama", .cekAlwaysOnTop.Name, .cekAlwaysOnTop.Value)
    End With

End Sub

Private Sub cmBantuan_Click()
    If FORM_UTAMA.cmBukuAlamat.Caption = "Buku Alamat" Then
        Select Case XTab1.ActiveTab
        Case Is = 0
            Kalimat = App.Path & "\bantuan\html\Umum.html"
        Case Is = 1
            Kalimat = App.Path & "\bantuan\html\System.html"
        Case Is = 2
            Kalimat = App.Path & "\bantuan\html\Penampilan.html"
        End Select
    ElseIf FORM_UTAMA.cmBukuAlamat.Caption = "Address Book" Then
        Select Case XTab1.ActiveTab
        Case Is = 0
            Kalimat = App.Path & "\bantuan\html\General.html"
        Case Is = 1
            Kalimat = App.Path & "\bantuan\html\System1.html"
        Case Is = 2
            Kalimat = App.Path & "\bantuan\html\Appearance.html"
        End Select
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

Private Sub cmOK_Click()
     PengaturanFORM_UTAMA
     SimpanPengaturanUtama
     FORM_UTAMA.PENGATURAN_BAHASA
     FORM_UTAMA.cmRefresh_Click
     Unload Me
End Sub

Private Sub cmSID_Click()
    FormMasukkanPassword.Show vbModal, Me
End Sub

Private Sub Form_Load()
    AturKontrol
    PENGATURAN_BAHASA
End Sub

Sub PengaturanFormIni()
    'PENGATURAN WARNA UNTUK FORM INI
    For Each Objek In Me
        Select Case cmbWarnaTampilan.ListIndex
        Case Is = 0
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguNatural
            If TypeName(Objek) = "XTab" Then Objek.ActiveTabBackEndColor = UnguNatural
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = UnguNatural
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = UnguNatural
        Case Is = 1
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Merah
            If TypeName(Objek) = "XTab" Then Objek.ActiveTabBackEndColor = Merah
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = Merah
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = Merah
        Case Is = 2
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Pink
            If TypeName(Objek) = "XTab" Then Objek.ActiveTabBackEndColor = Pink
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = Pink
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = Pink
        Case Is = 3
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = HijauMuda
            If TypeName(Objek) = "XTab" Then Objek.ActiveTabBackEndColor = HijauMuda
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = HijauMuda
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = HijauMuda
        Case Is = 4
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Hitam
            If TypeName(Objek) = "XTab" Then Objek.ActiveTabBackEndColor = Hitam
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = Hitam
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = Hitam
        Case Is = 5
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Silver
            If TypeName(Objek) = "XTab" Then Objek.ActiveTabBackEndColor = Silver
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = Silver
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = Silver
        Case Is = 6
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = SilverNatural
            If TypeName(Objek) = "XTab" Then Objek.ActiveTabBackEndColor = SilverNatural
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = SilverNatural
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = SilverNatural
        Case Is = 7
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Orange
            If TypeName(Objek) = "XTab" Then Objek.ActiveTabBackEndColor = Orange
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = Orange
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = Orange
        Case Is = 8
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguJanda
            If TypeName(Objek) = "XTab" Then Objek.ActiveTabBackEndColor = UnguJanda
            If TypeName(Objek) = "XPFrame" Then Objek.BackColor = UnguJanda
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = UnguJanda
        End Select
    Next
    'PENGATURAN THEMA UNTUK FORM INI
    For Each Objek In Me
        If TypeName(Objek) = "dcButton" Then
            Select Case cmbTemaTampilan.ListIndex
            Case Is = 0
                Objek.ButtonStyle = 3
            Case Is = 1
                Objek.ButtonStyle = 4
            Case Is = 2
                Objek.ButtonStyle = 5
            Case Is = 3
                Objek.ButtonStyle = 6
            Case Is = 4
                Objek.ButtonStyle = 7
            Case Is = 5
                Objek.ButtonStyle = 8
            Case Is = 6
                Objek.ButtonStyle = 9
            Case Is = 7
                Objek.ButtonStyle = 10
            Case Is = 8
                Objek.ButtonStyle = 11
                Objek.BackColor = &H12BCFF
            Case Is = 9
                Objek.ButtonStyle = 1
                Objek.BackColor = &HFF9B48
            Case Is = 10
                Objek.ButtonStyle = 2
            End Select
        End If
    Next
    'PENGATURAN UNTUK ALWAYS ON TOP
    If FormPengaturan.cekAlwaysOnTop.Value = Checked Then
        SetOnTop (Me.hwnd)
    ElseIf FormPengaturan.cekAlwaysOnTop.Value = Unchecked Then
        NotOnTop (Me.hwnd)
    End If
    'PENGATURAN UNTUK FORM TRANSPARAN
End Sub

Sub PENGATURAN_BAHASA()
    If cmbBahasa.ListIndex = 0 Then
        cekAutoRefresh.Caption = "Auto Refresh"
        cekGarisGrid.Caption = "Garis Grid"
        cekTampilkanPesanSimpan.Caption = "Tampilkan Pesan saat menyimpan data"
        cekKosongkanInput.Caption = "Kosongkan input saat berhasil menyimpan"
        cekPesanKonfirmasi.Caption = "Konfirmasi Saat data akan disimpan"
        cekTutupForm.Caption = "Tutup jendela saat data berhasil disimpan"
        cekKunciTabel.Caption = "Kunci Tabel"
        cekRiwayatAktivitas.Caption = "Catat Aktivitas Pengguna"
        CekTutupFormCAri.Caption = "Tutup Pencarian Saat Data Ditemukan"
        cmSID.Caption = "Set Internal Databases"
        cmBatal.Caption = "&Batal"
        XTab1.TabCaption(0) = "Umum"
        XTab1.TabCaption(1) = "System"
        XTab1.TabCaption(2) = "Penampilan"
        Label13.Caption = "Warna"
        Label12.Caption = "Tema"
        Label9.Caption = "Hasil Pencarian"
        Label8.Caption = "Default Tampil data"
        Label6.Caption = "Referensi Excel"
        Label3.Caption = "Format Default"
        Label1.Caption = "Bahasa"
        cekAlwaysOnTop.Caption = "Tampilan aplikasi selalu berada diatas"
        cmBantuan.Caption = "Bantuan"
    ElseIf cmbBahasa.ListIndex = 1 Then
        cekAutoRefresh.Caption = "Auto Refresh"
        cekGarisGrid.Caption = "Grid Lines"
        cekTampilkanPesanSimpan.Caption = "Show the message when saving the data"
        cekKosongkanInput.Caption = "Reset Inputs when saved successed"
        cekPesanKonfirmasi.Caption = "As confirmation of the data will be saved"
        cekTutupForm.Caption = "Close the window when the data is successfully saved"
        cekKunciTabel.Caption = "Lock Table"
        cekRiwayatAktivitas.Caption = "Record the User Activity"
        CekTutupFormCAri.Caption = "Close the search window When Data Found"
        cmSID.Caption = "Set Internal Databases"
        cmBatal.Caption = "&Cancel"
        XTab1.TabCaption(0) = "General"
        XTab1.TabCaption(1) = "System"
        XTab1.TabCaption(2) = "Appearance"
        Label13.Caption = "Color"
        Label12.Caption = "Thema"
        Label9.Caption = "Search Results"
        Label8.Caption = "Shown Default Data"
        Label6.Caption = "Reference of Excel"
        Label3.Caption = "Default Format"
        Label1.Caption = "Language"
        cekAlwaysOnTop.Caption = "Always On Top"
        cmBantuan.Caption = "Help"
    End If
End Sub

