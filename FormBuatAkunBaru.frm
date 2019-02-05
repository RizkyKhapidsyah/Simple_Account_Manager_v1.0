VERSION 5.00
Object = "{02353968-C1C9-4E0A-88D3-18759BDC60FE}#1.0#0"; "AeroSuite.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormBuatAkunBaru 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buat Akun Baru"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormBuatAkunBaru.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   480
      Top             =   5640
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   8
      Bmp:1           =   "FormBuatAkunBaru.frx":0442
      Key:1           =   "#menuSDKF"
      Bmp:2           =   "FormBuatAkunBaru.frx":0794
      Mask:2          =   -32640
      Key:2           =   "#menuEDT"
      Bmp:3           =   "FormBuatAkunBaru.frx":0BBC
      Key:3           =   "#menuED"
      Bmp:4           =   "FormBuatAkunBaru.frx":1924
      Mask:4          =   16645371
      Key:4           =   "#menuEDW"
      Bmp:5           =   "FormBuatAkunBaru.frx":1C76
      Mask:5          =   16383482
      Key:5           =   "#menuEDE"
      Bmp:6           =   "FormBuatAkunBaru.frx":1FC8
      Mask:6          =   16777215
      Key:6           =   "#menuPB"
      Bmp:7           =   "FormBuatAkunBaru.frx":2C1A
      Key:7           =   "#menuFAQ"
      Bmp:8           =   "FormBuatAkunBaru.frx":3982
      Key:8           =   "#menuKW"
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
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   330
      Left            =   240
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Keamanan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   49
      Top             =   7200
      Width           =   7575
      Begin Dacara_dcButton.dcButton cmCustom 
         Height          =   345
         Left            =   6240
         TabIndex        =   57
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BackColor       =   12230304
         ButtonStyle     =   3
         Caption         =   "&Custom"
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
         PicDown         =   "FormBuatAkunBaru.frx":3DAA
         PicHot          =   "FormBuatAkunBaru.frx":41FC
         PicNormal       =   "FormBuatAkunBaru.frx":464E
         PicSize         =   1
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin VB.ComboBox cmbPertanyaanRahasia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   240
         Width           =   3900
      End
      Begin AeroSuite.AeroTextBox textJawaban 
         Height          =   330
         Left            =   2280
         TabIndex        =   50
         Top             =   640
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   582
         BackColor       =   16777215
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "AeroTextBox1"
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pertanyaan Rahasia"
         Height          =   270
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   53
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jawaban"
         Height          =   270
         Left            =   120
         TabIndex        =   52
         Top             =   650
         Width           =   690
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   51
         Top             =   650
         Width           =   45
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Biodata Baru"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5055
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   7575
      Begin VB.ComboBox cmbJenisKelamin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ComboBox cmbAgama 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox textAlamat 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "FormBuatAkunBaru.frx":4AA0
         Top             =   2220
         Width           =   5155
      End
      Begin VB.ComboBox cmbAlamatWebsite 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3840
         Width           =   1215
      End
      Begin AeroSuite.AeroTextBox textNamaAsli 
         Height          =   330
         Left            =   2280
         TabIndex        =   11
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   582
         BackColor       =   16777215
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "AeroTextBox1"
      End
      Begin AeroSuite.AeroTextBox textTempat 
         Height          =   330
         Left            =   2280
         TabIndex        =   12
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         BackColor       =   16777215
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "AeroTextBox1"
      End
      Begin AeroSuite.AeroTextBox textTanggal 
         Height          =   330
         Left            =   5280
         TabIndex        =   13
         Top             =   720
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   582
         BackColor       =   16777215
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "12"
      End
      Begin AeroSuite.AeroTextBox textBulan 
         Height          =   330
         Left            =   5760
         TabIndex        =   14
         Top             =   720
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   582
         BackColor       =   16777215
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "12"
      End
      Begin AeroSuite.AeroTextBox textTahun 
         Height          =   330
         Left            =   6360
         TabIndex        =   15
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         BackColor       =   16777215
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "AeroTextBox1"
      End
      Begin AeroSuite.AeroTextBox textHobby 
         Height          =   330
         Left            =   2280
         TabIndex        =   16
         Top             =   1845
         Width           =   5155
         _ExtentX        =   9102
         _ExtentY        =   582
         BackColor       =   16777215
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "AeroTextBox1"
      End
      Begin AeroSuite.AeroTextBox textNomorTelepon 
         Height          =   330
         Left            =   2280
         TabIndex        =   17
         Top             =   3120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   582
         BackColor       =   16777215
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "AeroTextBox1"
      End
      Begin AeroSuite.AeroTextBox textAlamatEmail 
         Height          =   330
         Left            =   2280
         TabIndex        =   18
         Top             =   3480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   582
         BackColor       =   16777215
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "AeroTextBox1"
      End
      Begin AeroSuite.AeroTextBox textAlamatWebsite 
         Height          =   330
         Left            =   3480
         TabIndex        =   19
         Top             =   3840
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         BackColor       =   16777215
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "AeroTextBox1"
      End
      Begin AeroSuite.AeroTextBox textStatusAktivitas 
         Height          =   330
         Left            =   2280
         TabIndex        =   39
         Top             =   4200
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   582
         BackColor       =   16777215
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "AeroTextBox1"
      End
      Begin AeroSuite.AeroTextBox textStatusHubungan 
         Height          =   330
         Left            =   2280
         TabIndex        =   42
         Top             =   4560
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   582
         BackColor       =   16777215
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "AeroTextBox1"
      End
      Begin Dacara_dcButton.dcButton dcButton2 
         Height          =   345
         Left            =   4560
         TabIndex        =   60
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         BackColor       =   12230304
         ButtonStyle     =   3
         Caption         =   "&+ &Lain"
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
         PicDown         =   "FormBuatAkunBaru.frx":4AA6
         PicHot          =   "FormBuatAkunBaru.frx":4EF8
         PicNormal       =   "FormBuatAkunBaru.frx":534A
         PicSize         =   1
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin MSAdodcLib.Adodc AdodcDataLogin 
         Height          =   330
         Left            =   360
         Top             =   4560
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   44
         Top             =   4560
         Width           =   45
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Hubungan"
         Height          =   270
         Left            =   120
         TabIndex        =   43
         Top             =   4560
         Width           =   1365
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   41
         Top             =   4200
         Width           =   45
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Aktivitas"
         Height          =   270
         Left            =   120
         TabIndex        =   40
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Asli"
         Height          =   270
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   37
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat/Tanggal Lahir"
         Height          =   270
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   1785
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   35
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ","
         Height          =   270
         Left            =   3240
         TabIndex        =   34
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   33
         Top             =   1080
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         Height          =   270
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   31
         Top             =   1845
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hobby"
         Height          =   270
         Left            =   120
         TabIndex        =   30
         Top             =   1845
         Width           =   570
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agama"
         Height          =   270
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   28
         Top             =   1440
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   270
         Left            =   120
         TabIndex        =   27
         Top             =   2220
         Width           =   585
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   26
         Top             =   2220
         Width           =   45
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   25
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Telepon"
         Height          =   270
         Left            =   120
         TabIndex        =   24
         Top             =   3120
         Width           =   1305
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   23
         Top             =   3480
         Width           =   45
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat E-Mail"
         Height          =   270
         Left            =   120
         TabIndex        =   22
         Top             =   3480
         Width           =   1200
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   21
         Top             =   3840
         Width           =   45
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Website"
         Height          =   270
         Left            =   120
         TabIndex        =   20
         Top             =   3840
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Login Baru"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin AeroSuite.AeroCheckBox cekTampilkanPassword 
         Height          =   255
         Left            =   2280
         TabIndex        =   63
         Top             =   1440
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
         Align           =   0
         Caption         =   "Tampilkan Karakter Password"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         MouseIcon       =   "FormBuatAkunBaru.frx":579C
         Value           =   0
      End
      Begin VB.TextBox textKonfirmasiPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   1080
         Width           =   5175
      End
      Begin VB.TextBox textPasswordBaru 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   720
         Width           =   5175
      End
      Begin AeroSuite.AeroTextBox textNamaPengguna 
         Height          =   330
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   582
         BackColor       =   16777215
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "AeroTextBox1"
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Konfirmasi Password"
         Height          =   270
         Left            =   120
         TabIndex        =   48
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   47
         Top             =   1080
         Width           =   45
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   5
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password Baru"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pengguna Baru"
         Height          =   270
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1770
      End
   End
   Begin Dacara_dcButton.dcButton cmReset 
      Height          =   495
      Left            =   4440
      TabIndex        =   45
      Top             =   8400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BackColor       =   10591645
      ButtonStyle     =   2
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
      PicDown         =   "FormBuatAkunBaru.frx":57B8
      PicHot          =   "FormBuatAkunBaru.frx":5C0A
      PicNormal       =   "FormBuatAkunBaru.frx":605C
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
   End
   Begin Dacara_dcButton.dcButton cmSimpan 
      Height          =   495
      Left            =   6120
      TabIndex        =   46
      Top             =   8400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BackColor       =   10591645
      ButtonStyle     =   2
      Caption         =   "&Simpan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FormBuatAkunBaru.frx":64AE
      PicHot          =   "FormBuatAkunBaru.frx":6800
      PicNormal       =   "FormBuatAkunBaru.frx":6B52
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmVerifikasi 
      Height          =   495
      Left            =   2760
      TabIndex        =   56
      Top             =   8400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BackColor       =   10591645
      ButtonStyle     =   2
      Caption         =   "&Verifikasi"
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
      PicDown         =   "FormBuatAkunBaru.frx":6EA4
      PicHot          =   "FormBuatAkunBaru.frx":72F6
      PicNormal       =   "FormBuatAkunBaru.frx":7748
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
   End
   Begin Dacara_dcButton.dcButton cmBatal 
      Height          =   495
      Left            =   1080
      TabIndex        =   58
      Top             =   8400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BackColor       =   10591645
      ButtonStyle     =   2
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
      PicDown         =   "FormBuatAkunBaru.frx":7B9A
      PicHot          =   "FormBuatAkunBaru.frx":7FEC
      PicNormal       =   "FormBuatAkunBaru.frx":843E
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
   End
   Begin Dacara_dcButton.dcButton cmMenu 
      Height          =   495
      Left            =   120
      TabIndex        =   59
      Top             =   8400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BackColor       =   10591645
      ButtonStyle     =   2
      Caption         =   "V"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdodcPertanyaanRahasia 
      Height          =   330
      Left            =   120
      Top             =   3960
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdodcAgama 
      Height          =   330
      Left            =   120
      Top             =   4320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin XPEngine.XPControl XP_Engine 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.Menu menuTersembunyi 
      Caption         =   "Menu Tersembunyi"
      Begin VB.Menu menuSDKF 
         Caption         =   "Simpan Data ke File"
      End
      Begin VB.Menu menuEDT 
         Caption         =   "Extract Data ke Text"
      End
      Begin VB.Menu menuED 
         Caption         =   "Export Data"
         Begin VB.Menu menuEDE 
            Caption         =   "Excel 2003 Dokumen"
         End
         Begin VB.Menu menuEDW 
            Caption         =   "Word 2003 Dokumen"
         End
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu menuPS 
         Caption         =   "Prioritaskan Saya"
         Checked         =   -1  'True
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu menuPB 
         Caption         =   "Pusat Bantuan"
      End
      Begin VB.Menu menuKW 
         Caption         =   "Kunjungi Website"
      End
   End
End
Attribute VB_Name = "FormBuatAkunBaru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
NyambunggUtama
With AdodcUtama
    .ConnectionString = CN.ConnectionString
    .RecordSource = "Select * From tbLogin"
    .Refresh
End With
With AdodcPertanyaanRahasia
    .ConnectionString = CN.ConnectionString
    .RecordSource = "Select * from TbPertanyaanRahasia"
    .Refresh
End With
With AdodcAgama
    .ConnectionString = CN.ConnectionString
    .RecordSource = "Select * from TbAgama"
    .Refresh
End With
    
For Each Objek In Me
    If TypeName(Objek) = "AeroTextBox" Then
        With Objek
            .Text = ""
        End With
    ElseIf TypeName(Objek) = "TextBox" Then
        With Objek
            .Text = ""
            .MaxLength = "254"
        End With
    End If
Next
With Me
    .textTanggal.Text = Day(Date)
    .textBulan.Text = Month(Date)
    .textTahun.Text = Year(Date)
    .cmbJenisKelamin.Clear
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        .cmbJenisKelamin.AddItem "Laki-Laki", 0
        .cmbJenisKelamin.AddItem "Perempuan", 1
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        .cmbJenisKelamin.AddItem "Male", 0
        .cmbJenisKelamin.AddItem "Female", 1
    End If
    .cmbJenisKelamin.ListIndex = 0
    .cmbAgama.Clear
    Do Until AdodcAgama.Recordset.EOF
        .cmbAgama.AddItem AdodcAgama.Recordset.Fields(1).Value
        AdodcAgama.Recordset.MoveNext
    Loop
    .cmbAgama.Text = "Islam"
        
    .cmbAlamatWebsite.Clear
    .cmbAlamatWebsite.AddItem "http://", 0
    .cmbAlamatWebsite.AddItem "https://", 1
    .cmbAlamatWebsite.AddItem "ftp://", 2
    .cmbAlamatWebsite.AddItem "rtmp://", 3
    .cmbAlamatWebsite.AddItem "mms://", 4
    .cmbAlamatWebsite.ListIndex = 0
    .menuTersembunyi.Visible = False
End With
IsiCMBPertanyaanRahasia
    textPasswordBaru.PasswordChar = "*"
    textKonfirmasiPassword.PasswordChar = "*"
    PengaturanBahasa
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
Sub KosongkanInput()
For Each Objek In Me
    If TypeName(Objek) = "TextBox" Then
        Objek.Text = ""
    ElseIf TypeName(Objek) = "AeroTextBox" Then
        Objek.Text = ""
    End If
Next
End Sub
Sub IsiCMBPertanyaanRahasia()
With Me
    .cmbPertanyaanRahasia.Clear
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        .cmbPertanyaanRahasia.AddItem "Siapa teman masa kecil Anda?", 0
        .cmbPertanyaanRahasia.AddItem "Apa hoby yang paling Anda sukai?", 1
        .cmbPertanyaanRahasia.AddItem "Siapa nama ibu Anda?", 2
        .cmbPertanyaanRahasia.AddItem "Apa pekerjaan Ayah Anda?", 3
        .cmbPertanyaanRahasia.AddItem "Siapa tokoh kartun yang paling Anda Sukai?", 4
        .cmbPertanyaanRahasia.AddItem "Siapa musisi yang paling Anda sukai?", 5
        .cmbPertanyaanRahasia.AddItem "Siapa nama pacar atau istri/suami anda?", 6
        .cmbPertanyaanRahasia.AddItem "Siapa nama Adik ayah Anda?", 7
        .cmbPertanyaanRahasia.AddItem "Apa makanan dan minuman yang paling Anda sukai?", 8
        .cmbPertanyaanRahasia.AddItem "Apa merek rokok Anda?", 9
        .cmbPertanyaanRahasia.AddItem "Tanggal berapakah ulang tahun pacar atau istri/suami Anda?", 10
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        .cmbPertanyaanRahasia.AddItem "Who is Your Childhood Friend?", 0
        .cmbPertanyaanRahasia.AddItem "What hobbies do you like most?", 1
        .cmbPertanyaanRahasia.AddItem "Who is Your mother's name?", 2
        .cmbPertanyaanRahasia.AddItem "What is your father's job?", 3
        .cmbPertanyaanRahasia.AddItem "Who is the cartoon character You Like that?", 4
        .cmbPertanyaanRahasia.AddItem "Who are the musicians you like?", 5
        .cmbPertanyaanRahasia.AddItem "Who is the your girlfriend or wife/husband's name?", 6
        .cmbPertanyaanRahasia.AddItem "Who's your father's brother name?", 7
        .cmbPertanyaanRahasia.AddItem "What food and drink you like best?", 8
        .cmbPertanyaanRahasia.AddItem "What brand of you's cigarettes?", 9
        .cmbPertanyaanRahasia.AddItem "What is the anniversary date of your girlfriend or your wife/husband?", 10
    End If
        Do Until .AdodcPertanyaanRahasia.Recordset.EOF
            .cmbPertanyaanRahasia.AddItem AdodcPertanyaanRahasia.Recordset.Fields(0).Value
            .AdodcPertanyaanRahasia.Recordset.MoveNext
        Loop
        .AdodcPertanyaanRahasia.Refresh
    .cmbPertanyaanRahasia.ListIndex = 0
End With
End Sub
Sub PengaturanBahasa()
'PENGATURAN BAHASA
With Me
    Select Case FormPengaturan.cmbBahasa.ListIndex
    Case Is = 0
        Me.Caption = "Buat Akun Baru"
        Frame1.Caption = "Login Baru"
        Label24.Caption = "Nama Pengguna Baru"
        Label26.Caption = "Password Baru"
        Label29.Caption = "Konfirmasi Password"
        cekTampilkanPassword.Caption = "Tampilkan Karakter Password"
        Frame2.Caption = "Biodata Baru"
        Label1.Caption = "Nama Asli"
        Label3.Caption = "Tempat/Tanggal Lahir"
        Label7.Caption = "Jenis Kelamin"
        Label10.Caption = "Agama"
        Label9.Caption = "Hobby"
        Label12.Caption = "Alamat"
        Label15.Caption = "Nomor Telepon"
        Label17.Caption = "Alamat E-Mail"
        Label19.Caption = "Alamat Website"
        Label20.Caption = "Status Aktivitas"
        Label22.Caption = "Status Hubungan"
        Frame3.Caption = "Keamanan"
        Label35.Caption = "Pertanyaan Rahasia"
        Label33.Caption = "Jawaban"
        cmSimpan.Caption = "&Simpan"
        cmReset.Caption = "&Reset"
        cmVerifikasi.Caption = "&Verifikasi"
        cmBatal.Caption = "&Batal"
        dcButton2.Caption = "+ &Lain"
        menuSDKF.Caption = "Simpan Data ke File"
        menuEDT.Caption = "Ekstrak Data ke Teks"
        menuEDE.Caption = "Export Data ke Excel 2003"
        menuEDW.Caption = "Export Data ke Word 2003"
        menuPS.Caption = "Prioritaskan Saya"
        menuPB.Caption = "Pusat Bantuan"
        menuKW.Caption = "Kunjungi Website"
    Case Is = 1
        Me.Caption = "Create New Account"
        Frame1.Caption = "New Log In"
        Label24.Caption = "New UserName"
        Label26.Caption = "New Password"
        Label29.Caption = "Confirm of Password"
        cekTampilkanPassword.Caption = "Show Password Char"
        Frame2.Caption = "New Biodata"
        Label1.Caption = "Original Name"
        Label3.Caption = "Place/Date Birth"
        Label7.Caption = "Sex"
        Label10.Caption = "Religion"
        Label9.Caption = "Hobby"
        Label12.Caption = "Home Address"
        Label15.Caption = "Phone Number"
        Label17.Caption = "Mail Address"
        Label19.Caption = "Website Address"
        Label20.Caption = "Activity Status"
        Label22.Caption = "Relationship Status"
        Frame3.Caption = "Security"
        Label35.Caption = "Secret Question"
        Label33.Caption = "Answer"
        cmSimpan.Caption = "&Save"
        cmReset.Caption = "&Reset"
        cmVerifikasi.Caption = "&Verify"
        cmBatal.Caption = "&Cancel"
        dcButton2.Caption = "+ &Other"
        menuSDKF.Caption = "Save Entry to File.."
        menuEDT.Caption = "Extract Entry to Text"
        menuEDE.Caption = "Export Data to Excel 2003"
        menuEDW.Caption = "Export Data to Word 2003"
        menuPS.Caption = "I Prioritize"
        menuPB.Caption = "Help Center"
        menuKW.Caption = "Visit Website"
    End Select
End With
End Sub

Sub BUAT_DATABASE_BARU()
    Dim cn_tblogin As New ADODB.Connection
    Dim Posisi As Workspace
    
   
    On Error Resume Next
    'BAGIAN UNTUK MEMERIKSA DIREKTORI DAN MEMBUAT FOLDER
    Y = "C:\Windows\rssam\inc\pggn\" & textNamaPengguna.Text
    If DirectoryExist(Y) <> True Then Call CreateNewDirectory(Y) 'BUAT DIREKTORY



        DbLokasi = Y
        DbNama = "data.rdb"
        
        If Dir(DbLokasi & " \ " & DbNama) <> "" Then
            Kill DbLokasi & " \ " & DbNama
        End If
        
        Call CreateDB



        If cn_tblogin.State = adStateOpen Then cn_tblogin.Close
            cn_tblogin.CursorLocation = adUseClient
            cn_tblogin.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Y & "\data.rdb;Persist Security Info=False"
        With FormBuatAkunBaru.AdodcDataLogin
            .ConnectionString = cn_tblogin.ConnectionString
            .RecordSource = "Select * from tbDataLogin"
            .Refresh
            .Recordset.AddNew
            .Recordset.Fields(0).Value = FormBuatAkunBaru.textNamaPengguna.Text
            .Recordset.Fields(1).Value = FormBuatAkunBaru.textPasswordBaru.Text
            .Recordset.Fields(2).Value = FormBuatAkunBaru.textNamaAsli.Text
            .Recordset.Update
            .Refresh
            .Refresh
        End With
End Sub

Private Sub cekTampilkanPassword_Click()
If cekTampilkanPassword.Value = Checked Then
    textPasswordBaru.PasswordChar = ""
    textKonfirmasiPassword.PasswordChar = ""
Else
    textPasswordBaru.PasswordChar = "*"
    textKonfirmasiPassword.PasswordChar = "*"
End If
End Sub

Private Sub cmBatal_Click()
    If AdodcUtama.Recordset.RecordCount = 0 Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf, akun Anda belum dibuat. " & vbCrLf & _
                    "Persyaratan untuk menggunakan program ini adalah memiliki akun!" & vbCrLf & _
                    "Silahkan buat akun Anda!", vbExclamation + vbOKOnly, "Stop"
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry, your account has not been created" & vbCrLf & _
                    "Requirements for the use of this program is to have an account!" & vbCrLf & _
                    "Please create your account!", vbExclamation + vbOKOnly, "Stop"
        End If
    End If
        Unload Me
End Sub

Private Sub cmbPertanyaanRahasia_Change()
    cmbPertanyaanRahasia.ToolTipText = cmbPertanyaanRahasia.Text
End Sub

Private Sub cmbPertanyaanRahasia_Click()
    cmbPertanyaanRahasia.ToolTipText = cmbPertanyaanRahasia.Text
End Sub

Private Sub cmCustom_Click()
    With FormPertanyaanRahasiaCustom
        .Show vbModal, Me
    End With
End Sub

Private Sub cmMenu_Click()
    PopupMenu menuTersembunyi
End Sub

Sub SimpanDataPenggunaKeDatabase()
On Error GoTo HancurkanError
With AdodcUtama
    .Recordset.AddNew
    .Recordset.Fields(0).Value = textNamaAsli.Text
    .Recordset.Fields(1).Value = textTempat.Text
    .Recordset.Fields(2).Value = textTanggal.Text
    .Recordset.Fields(3).Value = textBulan.Text
    .Recordset.Fields(4).Value = textTahun.Text
    .Recordset.Fields(5).Value = cmbJenisKelamin.Text
    .Recordset.Fields(6).Value = cmbAgama.Text
    .Recordset.Fields(7).Value = textHobby.Text
    .Recordset.Fields(8).Value = textAlamat.Text
    .Recordset.Fields(9).Value = textNomorTelepon.Text
    .Recordset.Fields(10).Value = textAlamatEmail.Text
    .Recordset.Fields(11).Value = cmbAlamatWebsite.Text
    .Recordset.Fields(12).Value = textAlamatWebsite.Text
    .Recordset.Fields(13).Value = textStatusAktivitas.Text
    .Recordset.Fields(14).Value = textStatusHubungan.Text
    .Recordset.Fields(15).Value = textNamaPengguna.Text
    .Recordset.Fields(16).Value = textPasswordBaru.Text
    .Recordset.Fields(17).Value = cmbPertanyaanRahasia.Text
    .Recordset.Fields(18).Value = textJawaban.Text
    .Recordset.Update
    .Refresh
    .Refresh
End With
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Data Pengguna baru dengan nama pengguna '" & textNamaPengguna.Text & "' berhasil disimpan!", vbInformation + vbOKOnly, "Sukses!"
    Else
        MsgBox "New user with the name '" & textNamaPengguna.Text & "' is saved!", vbInformation + vbOKOnly, "Succcessed!"
    End If
Exit Sub
HancurkanError:
    PusatError
End Sub

Private Sub cmReset_Click()
    KosongkanInput
    textNamaPengguna.SetFocus
End Sub

Private Sub cmSimpan_Click()
Select Case FormPengaturan.cmbBahasa.ListIndex
Case Is = 0
    If textNamaPengguna.Text = "" Then
        MsgBox "Silahkan isi Nama dari Pengguna baru yang akan didaftarkan!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf textPasswordBaru.Text = "" Then
        MsgBox "Silahkan isi Password baru dari pengguna yang akan didaftarkan", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textKonfirmasiPassword.Text = "" Then
        MsgBox "Silahkan konfirmasikan password Anda!", vbExclamation + vbOKOnly, ""
        textKonfirmasiPassword.SetFocus
    ElseIf Len(textNamaPengguna.Text) <= 5 Then
        MsgBox "Nama Pengguna setidaknya minimal 6 karakter!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf Len(textPasswordBaru.Text) <= 5 Then
        MsgBox "Password setidaknya minimal 6 karakter!", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textPasswordBaru.Text <> textKonfirmasiPassword.Text Then
        MsgBox "Maaf, Konfirmasi password tidak sesuai dengan password baru yang diinputkan!", vbCritical + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textNamaAsli.Text = "" Then
        MsgBox "Silahkan isi nama asli dari pengguna!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textTempat.Text = "" Then
        MsgBox "Silahkan isi nama kota lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textTempat.SetFocus
    ElseIf textTanggal.Text = "" Then
        MsgBox "Silahkan isi tanggal lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textTanggal.SetFocus
    ElseIf textBulan.Text = "" Then
        MsgBox "Silahkan isi bulan lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textBulan.SetFocus
    ElseIf textTahun.Text = "" Then
        MsgBox "Silahkan isi tahun lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textTahun.SetFocus
    ElseIf textHobby.Text = "" Then
        MsgBox "Silahkan isi hobby Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textHobby.SetFocus
    ElseIf textAlamat.Text = "" Then
        MsgBox "Silahkan isi alamat Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textAlamat.SetFocus
    ElseIf textNomorTelepon.Text = "" Then
        MsgBox "Silahkan isi nomor telepon Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textNomorTelepon.SetFocus
    ElseIf textAlamatEmail.Text = "" Then
        MsgBox "Silahkan isi alamat email Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textAlamatEmail.SetFocus
    ElseIf textAlamatWebsite.Text = "" Then
        MsgBox "Silahkan isi alamat website Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textAlamatWebsite.SetFocus
    ElseIf textStatusAktivitas.Text = "" Then
        MsgBox "Silahkan isi status aktivitas Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textStatusAktivitas.SetFocus
    ElseIf textStatusHubungan.Text = "" Then
        MsgBox "Silahkan isi status hubungan Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textStatusHubungan.SetFocus
    ElseIf Len(textNamaAsli.Text) <= 5 Then
        MsgBox "Nama Asli setidaknya minimal 6 karakter!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textJawaban.Text = "" Then
        MsgBox "Silahkan isi jawaban dari pertanyaan rahasia Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textJawaban.SetFocus
    Else
        Pesan = MsgBox("Input sudah benar, Apakah Anda yakin dengan isian Anda?", vbQuestion + vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
            SimpanDataPenggunaKeDatabase
            BUAT_DATABASE_BARU
            KosongkanInput
            textNamaPengguna.SetFocus
            cmBatal.Caption = "&Tutup"
        End If
    End If
Case Is = 1
    If textNamaPengguna.Text = "" Then
        MsgBox "Please fill in the name of the new users will be registered!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf textPasswordBaru.Text = "" Then
        MsgBox "Please fill in the new password of the user to be registered", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textKonfirmasiPassword.Text = "" Then
        MsgBox "Please confirm your password!", vbExclamation + vbOKOnly, ""
        textKonfirmasiPassword.SetFocus
    ElseIf Len(textNamaPengguna.Text) <= 5 Then
        MsgBox "User Name at least a minimum of 6 characters!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf Len(textPasswordBaru.Text) <= 5 Then
        MsgBox "Password at least a minimum of 6 characters!", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textPasswordBaru.Text <> textKonfirmasiPassword.Text Then
        MsgBox "Sorry, Confirm password does not match the new password entered!", vbCritical + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textNamaAsli.Text = "" Then
        MsgBox "Please fill in the real name of the user!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textTempat.Text = "" Then
        MsgBox "Please fill in the name of the city of your birth!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textTempat.SetFocus
    ElseIf textTanggal.Text = "" Then
        MsgBox "Please fill in your date of birth", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textTanggal.SetFocus
    ElseIf textBulan.Text = "" Then
        MsgBox "Please fill in the month of your birth!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textBulan.SetFocus
    ElseIf textTahun.Text = "" Then
        MsgBox "Please fill in your birth year!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textTahun.SetFocus
    ElseIf textHobby.Text = "" Then
        MsgBox "Please fill in your hobby!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textHobby.SetFocus
    ElseIf textAlamat.Text = "" Then
        MsgBox "Please fill in your address!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textAlamat.SetFocus
    ElseIf textNomorTelepon.Text = "" Then
        MsgBox "Please fill in your phone number!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textNomorTelepon.SetFocus
    ElseIf textAlamatEmail.Text = "" Then
        MsgBox "Please fill in your email address!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textAlamatEmail.SetFocus
    ElseIf textAlamatWebsite.Text = "" Then
        MsgBox "Please fill in the address of your website!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textAlamatWebsite.SetFocus
    ElseIf textStatusAktivitas.Text = "" Then
        MsgBox "Please fill out the status of your activities!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textStatusAktivitas.SetFocus
    ElseIf textStatusHubungan.Text = "" Then
        MsgBox "Please fill in your relationship status!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textStatusHubungan.SetFocus
    ElseIf Len(textNamaAsli.Text) <= 5 Then
        MsgBox "Real Name at least a minimum of 6 characters!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textJawaban.Text = "" Then
        MsgBox "Please fill in the answer to your secret question!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textJawaban.SetFocus
    Else
        Pesan = MsgBox("Input is correct, Are you sure your entries?", vbQuestion + vbYesNo, "Confirmation")
        If Pesan = vbYes Then
            SimpanDataPenggunaKeDatabase
            BUAT_DATABASE_BARU
            KosongkanInput
            cmBatal.Caption = "&Close"
            textNamaPengguna.SetFocus
        End If
    End If
End Select
End Sub

Private Sub cmVerifikasi_Click()
If FormPengaturan.cmbBahasa.ListIndex = 0 Then
    If textNamaPengguna.Text = "" Then
        MsgBox "Tidak dapat mem-verifikasi input. " & vbCrLf & _
                "Silahkan isi Nama Pengguna yang dipakai untuk login!", vbExclamation + vbOKOnly, "Stop!"
        textNamaPengguna.SetFocus
    ElseIf textPasswordBaru.Text = "" Then
        MsgBox "Tidak dapat mem-verifikasi input. " & vbCrLf & _
                "Silahkan isi Password yang dipakai untuk login!", vbExclamation + vbOKOnly, "Stop!"
        textPasswordBaru.SetFocus
    ElseIf textKonfirmasiPassword.Text = "" Then
        MsgBox "Tidak dapat mem-verifikasi input. " & vbCrLf & _
                "Silahkan Konfirmasi Password baru yang dipakai untuk login!", vbExclamation + vbOKOnly, "Stop!"
        textKonfirmasiPassword.SetFocus
    ElseIf textNamaAsli.Text = "" Then
        MsgBox "Tidak dapat mem-verifikasi input. " & vbCrLf & _
                "Silahkan isi Nama Asli Anda!", vbExclamation + vbOKOnly, "Stop!"
        textNamaAsli.SetFocus
    ElseIf textJawaban.Text = "" Then
        MsgBox "Tidak dapat mem-verifikasi input. " & vbCrLf & _
                "Silahkan isi Jawaban dari pertanyaan rahasia Anda!", vbExclamation + vbOKOnly, "Stop!"
        textJawaban.SetFocus
    Else
        For Each Objek In Me
            If TypeName(Objek) = "TextBox" Then
                If Objek.Text = "" Then Objek.Text = "-"
            ElseIf TypeName(Objek) = "AeroTextBox" Then
                If Objek.Text = "" Then Objek.Text = "-"
            End If
        Next
    End If
ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
    If textNamaPengguna.Text = "" Then
        MsgBox "Unable to verify input. " & vbCrLf & _
                "Please fill in the user name used to log in!", vbExclamation + vbOKOnly, "Stop!"
        textNamaPengguna.SetFocus
    ElseIf textPasswordBaru.Text = "" Then
        MsgBox "Unable to verify input. " & vbCrLf & _
                "Please fill in the password used to log in!", vbExclamation + vbOKOnly, "Stop!"
        textPasswordBaru.SetFocus
    ElseIf textKonfirmasiPassword.Text = "" Then
        MsgBox "Unable to verify input. " & vbCrLf & _
                "Please Confirm the new password used to log in!", vbExclamation + vbOKOnly, "Stop!"
        textKonfirmasiPassword.SetFocus
    ElseIf textNamaAsli.Text = "" Then
        MsgBox "Unable to verify input. " & vbCrLf & _
                "Please fill in your Original Name!", vbExclamation + vbOKOnly, "Stop!"
        textNamaAsli.SetFocus
    ElseIf textJawaban.Text = "" Then
        MsgBox "Unable to verify input. " & vbCrLf & _
                "Please fill in the answer to your secret question!", vbExclamation + vbOKOnly, "Stop!"
        textJawaban.SetFocus
    Else
        For Each Objek In Me
            If TypeName(Objek) = "TextBox" Then
                If Objek.Text = "" Then Objek.Text = "-"
            ElseIf TypeName(Objek) = "AeroTextBox" Then
                If Objek.Text = "" Then Objek.Text = "-"
            End If
        Next
    End If
End If
End Sub


Private Sub dcButton2_Click()
With FormAgamaAnda
    .Show vbModal, Me
End With
End Sub

Private Sub Form_Load()
    AturKontrol
    PENGATURAN_WARNA
End Sub


Private Sub menuEDE_Click()
Select Case FormPengaturan.cmbBahasa.ListIndex
Case Is = 0
    If textNamaPengguna.Text = "" Then
        MsgBox "Silahkan isi Nama dari Pengguna baru yang akan didaftarkan!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf textPasswordBaru.Text = "" Then
        MsgBox "Silahkan isi Password baru dari pengguna yang akan didaftarkan", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textKonfirmasiPassword.Text = "" Then
        MsgBox "Silahkan konfirmasikan password Anda!", vbExclamation + vbOKOnly, ""
        textKonfirmasiPassword.SetFocus
    ElseIf Len(textNamaPengguna.Text) <= 5 Then
        MsgBox "Nama Pengguna setidaknya minimal 6 karakter!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf Len(textPasswordBaru.Text) <= 5 Then
        MsgBox "Password setidaknya minimal 6 karakter!", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textPasswordBaru.Text <> textKonfirmasiPassword.Text Then
        MsgBox "Maaf, Konfirmasi password tidak sesuai dengan password baru yang diinputkan!", vbCritical + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textNamaAsli.Text = "" Then
        MsgBox "Silahkan isi nama asli dari pengguna!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textTempat.Text = "" Then
        MsgBox "Silahkan isi nama kota lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textTempat.SetFocus
    ElseIf textTanggal.Text = "" Then
        MsgBox "Silahkan isi tanggal lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textTanggal.SetFocus
    ElseIf textBulan.Text = "" Then
        MsgBox "Silahkan isi bulan lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textBulan.SetFocus
    ElseIf textTahun.Text = "" Then
        MsgBox "Silahkan isi tahun lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textTahun.SetFocus
    ElseIf textHobby.Text = "" Then
        MsgBox "Silahkan isi hobby Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textHobby.SetFocus
    ElseIf textAlamat.Text = "" Then
        MsgBox "Silahkan isi alamat Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textAlamat.SetFocus
    ElseIf textNomorTelepon.Text = "" Then
        MsgBox "Silahkan isi nomor telepon Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textNomorTelepon.SetFocus
    ElseIf textAlamatEmail.Text = "" Then
        MsgBox "Silahkan isi alamat email Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textAlamatEmail.SetFocus
    ElseIf textAlamatWebsite.Text = "" Then
        MsgBox "Silahkan isi alamat website Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textAlamatWebsite.SetFocus
    ElseIf textStatusAktivitas.Text = "" Then
        MsgBox "Silahkan isi status aktivitas Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textStatusAktivitas.SetFocus
    ElseIf textStatusHubungan.Text = "" Then
        MsgBox "Silahkan isi status hubungan Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textStatusHubungan.SetFocus
    ElseIf Len(textNamaAsli.Text) <= 5 Then
        MsgBox "Nama Asli setidaknya minimal 6 karakter!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textJawaban.Text = "" Then
        MsgBox "Silahkan isi jawaban dari pertanyaan rahasia Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textJawaban.SetFocus
    Else
        With FormExportDataToExcel
            .Caption = "Export Data >> Excel 2003"
            .Show vbModal, Me
        End With
    End If
Case Is = 1
    If textNamaPengguna.Text = "" Then
        MsgBox "Please fill in the name of the new users will be registered!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf textPasswordBaru.Text = "" Then
        MsgBox "Please fill in the new password of the user to be registered", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textKonfirmasiPassword.Text = "" Then
        MsgBox "Please confirm your password!", vbExclamation + vbOKOnly, ""
        textKonfirmasiPassword.SetFocus
    ElseIf Len(textNamaPengguna.Text) <= 5 Then
        MsgBox "User Name at least a minimum of 6 characters!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf Len(textPasswordBaru.Text) <= 5 Then
        MsgBox "Password at least a minimum of 6 characters!", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textPasswordBaru.Text <> textKonfirmasiPassword.Text Then
        MsgBox "Sorry, Confirm password does not match the new password entered!", vbCritical + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textNamaAsli.Text = "" Then
        MsgBox "Please fill in the real name of the user!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textTempat.Text = "" Then
        MsgBox "Please fill in the name of the city of your birth!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textTempat.SetFocus
    ElseIf textTanggal.Text = "" Then
        MsgBox "Please fill in your date of birth", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textTanggal.SetFocus
    ElseIf textBulan.Text = "" Then
        MsgBox "Please fill in the month of your birth!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textBulan.SetFocus
    ElseIf textTahun.Text = "" Then
        MsgBox "Please fill in your birth year!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textTahun.SetFocus
    ElseIf textHobby.Text = "" Then
        MsgBox "Please fill in your hobby!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textHobby.SetFocus
    ElseIf textAlamat.Text = "" Then
        MsgBox "Please fill in your address!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textAlamat.SetFocus
    ElseIf textNomorTelepon.Text = "" Then
        MsgBox "Please fill in your phone number!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textNomorTelepon.SetFocus
    ElseIf textAlamatEmail.Text = "" Then
        MsgBox "Please fill in your email address!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textAlamatEmail.SetFocus
    ElseIf textAlamatWebsite.Text = "" Then
        MsgBox "Please fill in the address of your website!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textAlamatWebsite.SetFocus
    ElseIf textStatusAktivitas.Text = "" Then
        MsgBox "Please fill out the status of your activities!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textStatusAktivitas.SetFocus
    ElseIf textStatusHubungan.Text = "" Then
        MsgBox "Please fill in your relationship status!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textStatusHubungan.SetFocus
    ElseIf Len(textNamaAsli.Text) <= 5 Then
        MsgBox "Real Name at least a minimum of 6 characters!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textJawaban.Text = "" Then
        MsgBox "Please fill in the answer to your secret question!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textJawaban.SetFocus
    Else
        With FormExportDataToExcel
            .Caption = "Export Data >> Excel 2003"
            .Show vbModal, Me
        End With
    End If
End Select
End Sub

Private Sub menuEDT_Click()
Select Case FormPengaturan.cmbBahasa.ListIndex
Case Is = 0
    If textNamaPengguna.Text = "" Then
        MsgBox "Silahkan isi Nama dari Pengguna baru yang akan didaftarkan!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf textPasswordBaru.Text = "" Then
        MsgBox "Silahkan isi Password baru dari pengguna yang akan didaftarkan", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textKonfirmasiPassword.Text = "" Then
        MsgBox "Silahkan konfirmasikan password Anda!", vbExclamation + vbOKOnly, ""
        textKonfirmasiPassword.SetFocus
    ElseIf Len(textNamaPengguna.Text) <= 5 Then
        MsgBox "Nama Pengguna setidaknya minimal 6 karakter!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf Len(textPasswordBaru.Text) <= 5 Then
        MsgBox "Password setidaknya minimal 6 karakter!", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textPasswordBaru.Text <> textKonfirmasiPassword.Text Then
        MsgBox "Maaf, Konfirmasi password tidak sesuai dengan password baru yang diinputkan!", vbCritical + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textNamaAsli.Text = "" Then
        MsgBox "Silahkan isi nama asli dari pengguna!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textTempat.Text = "" Then
        MsgBox "Silahkan isi nama kota lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textTempat.SetFocus
    ElseIf textTanggal.Text = "" Then
        MsgBox "Silahkan isi tanggal lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textTanggal.SetFocus
    ElseIf textBulan.Text = "" Then
        MsgBox "Silahkan isi bulan lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textBulan.SetFocus
    ElseIf textTahun.Text = "" Then
        MsgBox "Silahkan isi tahun lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textTahun.SetFocus
    ElseIf textHobby.Text = "" Then
        MsgBox "Silahkan isi hobby Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textHobby.SetFocus
    ElseIf textAlamat.Text = "" Then
        MsgBox "Silahkan isi alamat Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textAlamat.SetFocus
    ElseIf textNomorTelepon.Text = "" Then
        MsgBox "Silahkan isi nomor telepon Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textNomorTelepon.SetFocus
    ElseIf textAlamatEmail.Text = "" Then
        MsgBox "Silahkan isi alamat email Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textAlamatEmail.SetFocus
    ElseIf textAlamatWebsite.Text = "" Then
        MsgBox "Silahkan isi alamat website Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textAlamatWebsite.SetFocus
    ElseIf textStatusAktivitas.Text = "" Then
        MsgBox "Silahkan isi status aktivitas Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textStatusAktivitas.SetFocus
    ElseIf textStatusHubungan.Text = "" Then
        MsgBox "Silahkan isi status hubungan Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textStatusHubungan.SetFocus
    ElseIf Len(textNamaAsli.Text) <= 5 Then
        MsgBox "Nama Asli setidaknya minimal 6 karakter!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textJawaban.Text = "" Then
        MsgBox "Silahkan isi jawaban dari pertanyaan rahasia Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textJawaban.SetFocus
    Else
        With FormEkstrakDataKeText
            .Show vbModal, Me
        End With
    End If
Case Is = 1
    If textNamaPengguna.Text = "" Then
        MsgBox "Please fill in the name of the new users will be registered!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf textPasswordBaru.Text = "" Then
        MsgBox "Please fill in the new password of the user to be registered", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textKonfirmasiPassword.Text = "" Then
        MsgBox "Please confirm your password!", vbExclamation + vbOKOnly, ""
        textKonfirmasiPassword.SetFocus
    ElseIf Len(textNamaPengguna.Text) <= 5 Then
        MsgBox "User Name at least a minimum of 6 characters!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf Len(textPasswordBaru.Text) <= 5 Then
        MsgBox "Password at least a minimum of 6 characters!", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textPasswordBaru.Text <> textKonfirmasiPassword.Text Then
        MsgBox "Sorry, Confirm password does not match the new password entered!", vbCritical + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textNamaAsli.Text = "" Then
        MsgBox "Please fill in the real name of the user!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textTempat.Text = "" Then
        MsgBox "Please fill in the name of the city of your birth!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textTempat.SetFocus
    ElseIf textTanggal.Text = "" Then
        MsgBox "Please fill in your date of birth", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textTanggal.SetFocus
    ElseIf textBulan.Text = "" Then
        MsgBox "Please fill in the month of your birth!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textBulan.SetFocus
    ElseIf textTahun.Text = "" Then
        MsgBox "Please fill in your birth year!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textTahun.SetFocus
    ElseIf textHobby.Text = "" Then
        MsgBox "Please fill in your hobby!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textHobby.SetFocus
    ElseIf textAlamat.Text = "" Then
        MsgBox "Please fill in your address!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textAlamat.SetFocus
    ElseIf textNomorTelepon.Text = "" Then
        MsgBox "Please fill in your phone number!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textNomorTelepon.SetFocus
    ElseIf textAlamatEmail.Text = "" Then
        MsgBox "Please fill in your email address!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textAlamatEmail.SetFocus
    ElseIf textAlamatWebsite.Text = "" Then
        MsgBox "Please fill in the address of your website!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textAlamatWebsite.SetFocus
    ElseIf textStatusAktivitas.Text = "" Then
        MsgBox "Please fill out the status of your activities!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textStatusAktivitas.SetFocus
    ElseIf textStatusHubungan.Text = "" Then
        MsgBox "Please fill in your relationship status!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textStatusHubungan.SetFocus
    ElseIf Len(textNamaAsli.Text) <= 5 Then
        MsgBox "Real Name at least a minimum of 6 characters!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textJawaban.Text = "" Then
        MsgBox "Please fill in the answer to your secret question!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textJawaban.SetFocus
    Else
        With FormEkstrakDataKeText
            .Show vbModal, Me
        End With
    End If
End Select
End Sub

Private Sub menuEDW_Click()
Select Case FormPengaturan.cmbBahasa.ListIndex
Case Is = 0
    If textNamaPengguna.Text = "" Then
        MsgBox "Silahkan isi Nama dari Pengguna baru yang akan didaftarkan!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf textPasswordBaru.Text = "" Then
        MsgBox "Silahkan isi Password baru dari pengguna yang akan didaftarkan", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textKonfirmasiPassword.Text = "" Then
        MsgBox "Silahkan konfirmasikan password Anda!", vbExclamation + vbOKOnly, ""
        textKonfirmasiPassword.SetFocus
    ElseIf Len(textNamaPengguna.Text) <= 5 Then
        MsgBox "Nama Pengguna setidaknya minimal 6 karakter!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf Len(textPasswordBaru.Text) <= 5 Then
        MsgBox "Password setidaknya minimal 6 karakter!", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textPasswordBaru.Text <> textKonfirmasiPassword.Text Then
        MsgBox "Maaf, Konfirmasi password tidak sesuai dengan password baru yang diinputkan!", vbCritical + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textNamaAsli.Text = "" Then
        MsgBox "Silahkan isi nama asli dari pengguna!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textTempat.Text = "" Then
        MsgBox "Silahkan isi nama kota lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textTempat.SetFocus
    ElseIf textTanggal.Text = "" Then
        MsgBox "Silahkan isi tanggal lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textTanggal.SetFocus
    ElseIf textBulan.Text = "" Then
        MsgBox "Silahkan isi bulan lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textBulan.SetFocus
    ElseIf textTahun.Text = "" Then
        MsgBox "Silahkan isi tahun lahir Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textTahun.SetFocus
    ElseIf textHobby.Text = "" Then
        MsgBox "Silahkan isi hobby Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textHobby.SetFocus
    ElseIf textAlamat.Text = "" Then
        MsgBox "Silahkan isi alamat Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textAlamat.SetFocus
    ElseIf textNomorTelepon.Text = "" Then
        MsgBox "Silahkan isi nomor telepon Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textNomorTelepon.SetFocus
    ElseIf textAlamatEmail.Text = "" Then
        MsgBox "Silahkan isi alamat email Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textAlamatEmail.SetFocus
    ElseIf textAlamatWebsite.Text = "" Then
        MsgBox "Silahkan isi alamat website Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textAlamatWebsite.SetFocus
    ElseIf textStatusAktivitas.Text = "" Then
        MsgBox "Silahkan isi status aktivitas Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textStatusAktivitas.SetFocus
    ElseIf textStatusHubungan.Text = "" Then
        MsgBox "Silahkan isi status hubungan Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textStatusHubungan.SetFocus
    ElseIf Len(textNamaAsli.Text) <= 5 Then
        MsgBox "Nama Asli setidaknya minimal 6 karakter!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textJawaban.Text = "" Then
        MsgBox "Silahkan isi jawaban dari pertanyaan rahasia Anda!", vbExclamation + vbOKOnly, ""
        MsgBox "Jika memang ingin dikosongkan, silahkan klik Verifikasi", vbInformation + vbOKOnly, ""
        textJawaban.SetFocus
    Else
        With FormExportDataToExcel
            .Caption = "Export Data >> Word 2003"
            .Show vbModal, Me
        End With
    End If
Case Is = 1
    If textNamaPengguna.Text = "" Then
        MsgBox "Please fill in the name of the new users will be registered!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf textPasswordBaru.Text = "" Then
        MsgBox "Please fill in the new password of the user to be registered", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textKonfirmasiPassword.Text = "" Then
        MsgBox "Please confirm your password!", vbExclamation + vbOKOnly, ""
        textKonfirmasiPassword.SetFocus
    ElseIf Len(textNamaPengguna.Text) <= 5 Then
        MsgBox "User Name at least a minimum of 6 characters!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf Len(textPasswordBaru.Text) <= 5 Then
        MsgBox "Password at least a minimum of 6 characters!", vbExclamation + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textPasswordBaru.Text <> textKonfirmasiPassword.Text Then
        MsgBox "Sorry, Confirm password does not match the new password entered!", vbCritical + vbOKOnly, ""
        textPasswordBaru.SetFocus
    ElseIf textNamaAsli.Text = "" Then
        MsgBox "Please fill in the real name of the user!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textTempat.Text = "" Then
        MsgBox "Please fill in the name of the city of your birth!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textTempat.SetFocus
    ElseIf textTanggal.Text = "" Then
        MsgBox "Please fill in your date of birth", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textTanggal.SetFocus
    ElseIf textBulan.Text = "" Then
        MsgBox "Please fill in the month of your birth!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textBulan.SetFocus
    ElseIf textTahun.Text = "" Then
        MsgBox "Please fill in your birth year!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textTahun.SetFocus
    ElseIf textHobby.Text = "" Then
        MsgBox "Please fill in your hobby!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textHobby.SetFocus
    ElseIf textAlamat.Text = "" Then
        MsgBox "Please fill in your address!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textAlamat.SetFocus
    ElseIf textNomorTelepon.Text = "" Then
        MsgBox "Please fill in your phone number!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textNomorTelepon.SetFocus
    ElseIf textAlamatEmail.Text = "" Then
        MsgBox "Please fill in your email address!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textAlamatEmail.SetFocus
    ElseIf textAlamatWebsite.Text = "" Then
        MsgBox "Please fill in the address of your website!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textAlamatWebsite.SetFocus
    ElseIf textStatusAktivitas.Text = "" Then
        MsgBox "Please fill out the status of your activities!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textStatusAktivitas.SetFocus
    ElseIf textStatusHubungan.Text = "" Then
        MsgBox "Please fill in your relationship status!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textStatusHubungan.SetFocus
    ElseIf Len(textNamaAsli.Text) <= 5 Then
        MsgBox "Real Name at least a minimum of 6 characters!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textNamaAsli.SetFocus
    ElseIf textJawaban.Text = "" Then
        MsgBox "Please fill in the answer to your secret question!", vbExclamation + vbOKOnly, ""
        MsgBox "If you really want evacuated, please click Verify", vbInformation + vbOKOnly, ""
        textJawaban.SetFocus
    Else
        With FormExportDataToExcel
            .Caption = "Export Data >> Word 2003"
            .Show vbModal, Me
        End With
    End If
End Select
End Sub


Private Sub menuKW_Click()
    Kalimat = "http://rikymetalist.blogspot.com/p/software-ku.html"
    SITUS = ShellExecute(0, vbNullString, Kalimat, "", "", vbNormalFocus)
End Sub

Private Sub menuPB_Click()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Kalimat = App.Path & "\bantuan\html\BuatAkunBaru.html"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Kalimat = App.Path & "\bantuan\html\CreateNewAccount.html"
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

Private Sub menuPS_Click()
Select Case menuPS.Checked
    Case Is = True
        menuPS.Checked = False
    Case Is = False
        menuPS.Checked = True
End Select
End Sub

Private Sub menuSDKF_Click()
On Error GoTo ErrorHandler
    With CommonDialog1
        .Filter = "All Files (*.*)|*.*|RikySoft Catatan Files (*.rcf)|*.rcf|RikySoft Mail Files (*.rmd)|*.rmd|Text Files (*.txt)|*.txt|Word 2003 Files (*.doc)|*.doc|Excel 2003 Files (*.xls)|*.xls|HTML Files (*.htm)|*.htm"
        AturDefaultFormat
        .ShowSave
        .FileName = CommonDialog1.FileName
    End With
    Dim iFile As Integer
    Dim SaveFileFromTB As Boolean
    Dim TxtBox As Object
    Dim FilePath As String
    Dim Append As Boolean
    iFile = FreeFile
        If Append Then
            Open CommonDialog1.FileName For Append As #iFile
        Else
            Open CommonDialog1.FileName For Output As #iFile
        End If
    Print #iFile, "==============" & Frame1.Caption & "=============="
    Print #iFile, Label24.Caption & " : " & textNamaPengguna.Text
    Print #iFile, Label26.Caption & " : " & textPasswordBaru.Text
    Print #iFile, Label29.Caption & " : " & textKonfirmasiPassword.Text
    Print #iFile, ""
    Print #iFile, "==============" & Frame2.Caption & "=============="
    Print #iFile, Label1.Caption & " : " & textNamaAsli.Text
    Print #iFile, Label3.Caption & " : " & textTempat.Text & ", " & textTanggal.Text & " - " & textBulan.Text & " - " & textTahun.Text
    Print #iFile, Label7.Caption & " : " & cmbJenisKelamin.Text
    Print #iFile, Label10.Caption & " : " & cmbAgama.Text
    Print #iFile, Label9.Caption & " : " & textHobby.Text
    Print #iFile, Label12.Caption & " : " & textAlamat.Text
    Print #iFile, Label15.Caption & " : " & textNomorTelepon.Text
    Print #iFile, Label17.Caption & " : " & textAlamatEmail.Text
    Print #iFile, Label19.Caption & " : " & cmbAlamatWebsite.Text & textAlamatWebsite.Text
    Print #iFile, Label20.Caption & " : " & textStatusAktivitas.Text
    Print #iFile, Label22.Caption & " : " & textStatusHubungan.Text
    Print #iFile, ""
    Print #iFile, "==============" & Frame3.Caption & "=============="
    Print #iFile, Label35.Caption & " : " & cmbPertanyaanRahasia.Text
    Print #iFile, Label33.Caption & " : " & textJawaban.Text

    SaveFileFromTB = True
ErrorHandler:
    Close #iFile
End Sub

Private Sub textKonfirmasiPassword_Change()
    textKonfirmasiPassword.ForeColor = Hitam
End Sub

Private Sub textKonfirmasiPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then textNamaAsli.SetFocus
End Sub

Private Sub textNamaPengguna_Change()
    textNamaPengguna.ForeColor = Hitam
End Sub

Private Sub textNamaPengguna_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then textPasswordBaru.SetFocus
End Sub

Private Sub textPasswordBaru_Change()
    textPasswordBaru.ForeColor = Hitam
End Sub

Private Sub textPasswordBaru_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then textKonfirmasiPassword.SetFocus
End Sub

Sub PENGATURAN_WARNA()
    'PENGATURAN WARNA UNTUK FORM INI
    For Each Objek In Me
        Select Case FormPengaturan.cmbWarnaTampilan.ListIndex
        Case Is = 0 'Ungu Natural
            Me.BackColor = UnguNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguNatural
            If TypeName(Objek) = "Frame" Then Objek.BackColor = UnguNatural
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = UnguNatural
        Case Is = 1 'Merah
            Me.BackColor = Merah
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Merah
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Merah
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = Merah
        Case Is = 2 'Pink
            Me.BackColor = Pink
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Pink
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Pink
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = Pink
        Case Is = 3 'HijauMuda
            Me.BackColor = HijauMuda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = HijauMuda
            If TypeName(Objek) = "Frame" Then Objek.BackColor = HijauMuda
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = HijauMuda
        Case Is = 4 'Hitam
            Me.BackColor = Hitam
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Hitam
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Hitam
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = Hitam
        Case Is = 5 'Silver
            Me.BackColor = Silver
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Silver
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Silver
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = Silver
        Case Is = 6 'SilverNatural
            Me.BackColor = SilverNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = SilverNatural
            If TypeName(Objek) = "Frame" Then Objek.BackColor = SilverNatural
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = SilverNatural
        Case Is = 7 'Orange
            Me.BackColor = Orange
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Orange
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Orange
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = Orange
        Case Is = 8 'UnguJanda
            Me.BackColor = UnguJanda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguJanda
            If TypeName(Objek) = "Frame" Then Objek.BackColor = UnguJanda
            If TypeName(Objek) = "AeroCheckBox" Then Objek.BackColor = UnguJanda
        End Select
    Next
    'PENGATURAN THEMA UNTUK FORM INI
    For Each Objek In Me
        If TypeName(Objek) = "dcButton" Then
            Select Case FormPengaturan.cmbTemaTampilan.ListIndex
            Case Is = 0 'RST_Office 2003
                Objek.ButtonStyle = 3
                Objek.BackColor = &HC0C0C0
                cmMenu.ButtonStyle = 2
                cmBatal.ButtonStyle = 2
                cmVerifikasi.ButtonStyle = 2
                cmReset.ButtonStyle = 2
                cmSimpan.ButtonStyle = 2
                cmMenu.BackColor = &HA19D9D
                cmBatal.BackColor = &HA19D9D
                cmVerifikasi.BackColor = &HA19D9D
                cmReset.BackColor = &HA19D9D
                cmSimpan.BackColor = &HA19D9D
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
