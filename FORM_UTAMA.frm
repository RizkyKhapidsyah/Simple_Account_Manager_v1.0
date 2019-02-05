VERSION 5.00
Object = "{02353968-C1C9-4E0A-88D3-18759BDC60FE}#1.0#0"; "AeroSuite.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FORM_UTAMA 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Account Manager - "
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12840
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FORM_UTAMA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   12840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timerDefaultTampilkanData 
      Interval        =   1
      Left            =   7080
      Top             =   3840
   End
   Begin VB.Timer TimerProgress 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   1920
      Top             =   6720
   End
   Begin AeroSuite.AeroProgressBar ProgressLogOut 
      Height          =   270
      Left            =   5760
      Top             =   7570
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   476
   End
   Begin VB.Timer TimerWaktu 
      Interval        =   10
      Left            =   1440
      Top             =   6720
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   4560
      Top             =   3480
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   30
      Bmp:1           =   "FORM_UTAMA.frx":0442
      Key:1           =   "#menuBaru"
      Bmp:2           =   "FORM_UTAMA.frx":0544
      Mask:2          =   16777215
      Key:2           =   "#menuBaruIP"
      Bmp:3           =   "FORM_UTAMA.frx":0896
      Mask:3          =   16711422
      Key:3           =   "#menuBaruBA"
      Bmp:4           =   "FORM_UTAMA.frx":0BE8
      Mask:4          =   16777215
      Key:4           =   "#menuBaruUT"
      Bmp:5           =   "FORM_UTAMA.frx":0F3A
      Mask:5          =   16382457
      Key:5           =   "#menuBaruA"
      Bmp:6           =   "FORM_UTAMA.frx":128C
      Mask:6          =   16579836
      Key:6           =   "#menuBaruRS"
      Bmp:7           =   "FORM_UTAMA.frx":15DE
      Mask:7          =   16777215
      Key:7           =   "#menuBaruJS"
      Bmp:8           =   "FORM_UTAMA.frx":1930
      Mask:8          =   16777215
      Key:8           =   "#menubaruEM"
      Bmp:9           =   "FORM_UTAMA.frx":1C82
      Mask:9          =   16777215
      Key:9           =   "#menuFI"
      Bmp:10          =   "FORM_UTAMA.frx":1FD4
      Mask:10         =   16777215
      Key:10          =   "#menuBaruFTP"
      Bmp:11          =   "FORM_UTAMA.frx":2326
      Key:11          =   "#menuBaruBlogging"
      Bmp:12          =   "FORM_UTAMA.frx":274E
      Mask:12         =   16777215
      Key:12          =   "#menuAkun"
      Bmp:13          =   "FORM_UTAMA.frx":2AA0
      Mask:13         =   16777215
      Key:13          =   "#menuBuatCadangan"
      Bmp:14          =   "FORM_UTAMA.frx":2DF2
      Mask:14         =   16777215
      Key:14          =   "#menuLogOut"
      Bmp:15          =   "FORM_UTAMA.frx":3144
      Mask:15         =   16777215
      Key:15          =   "#menuSelaluDiatas"
      Bmp:16          =   "FORM_UTAMA.frx":3496
      Mask:16         =   16777215
      Key:16          =   "#menuTabelFilter"
      Bmp:17          =   "FORM_UTAMA.frx":37E8
      Mask:17         =   16777215
      Key:17          =   "#menuRA"
      Bmp:18          =   "FORM_UTAMA.frx":3B3A
      Mask:18         =   16777215
      Key:18          =   "#menuASCG"
      Bmp:19          =   "FORM_UTAMA.frx":3E8C
      Mask:19         =   16316669
      Key:19          =   "#menuAR"
      Bmp:20          =   "FORM_UTAMA.frx":41DE
      Mask:20         =   8323072
      Key:20          =   "#menuES"
      Bmp:21          =   "FORM_UTAMA.frx":4530
      Mask:21         =   16777215
      Key:21          =   "#menuKV"
      Bmp:22          =   "FORM_UTAMA.frx":4882
      Mask:22         =   16777215
      Key:22          =   "#menuSN"
      Bmp:23          =   "FORM_UTAMA.frx":4BD4
      Mask:23         =   16777215
      Key:23          =   "#menuNotepad"
      Bmp:24          =   "FORM_UTAMA.frx":4F26
      Mask:24         =   12632256
      Key:24          =   "#menuaPengaturan"
      Bmp:25          =   "FORM_UTAMA.frx":5278
      Mask:25         =   16777215
      Key:25          =   "#menuPusatBantuan"
      Bmp:26          =   "FORM_UTAMA.frx":55CA
      Key:26          =   "#menuKH"
      Bmp:27          =   "FORM_UTAMA.frx":59F2
      Mask:27         =   16711679
      Key:27          =   "#menuDonasi"
      Bmp:28          =   "FORM_UTAMA.frx":5D44
      Mask:28         =   11842740
      Key:28          =   "#menuPVT"
      Bmp:29          =   "FORM_UTAMA.frx":6096
      Mask:29         =   16777215
      Key:29          =   "#menuTSAM"
      Bmp:30          =   "FORM_UTAMA.frx":63E8
      Mask:30         =   16777215
      Key:30          =   "#menuHubungiDeveloper"
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
   Begin MSAdodcLib.Adodc ADODC_UTAMA 
      Height          =   330
      Left            =   120
      Top             =   7440
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
   Begin MSComctlLib.ListView LV 
      Height          =   5295
      Left            =   2400
      TabIndex        =   0
      Top             =   1440
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9340
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame FrameAkunDesktop 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
      Begin Dacara_dcButton.dcButton cmAkunDesktop 
         Height          =   615
         Left            =   0
         TabIndex        =   8
         Top             =   75
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1085
         BackColor       =   8421504
         ButtonStyle     =   1
         Caption         =   "Akun Desktop"
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
   End
   Begin VB.Frame FrameAkunWeb 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -120
      TabIndex        =   2
      Top             =   3960
      Width           =   2535
      Begin Dacara_dcButton.dcButton cmAkunWeb 
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   80
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1085
         BackColor       =   8421504
         ButtonStyle     =   1
         Caption         =   "Akun Web"
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
   End
   Begin Dacara_dcButton.dcButton cmIdentitasPribadi 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2070
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
      BackColor       =   13815503
      ButtonStyle     =   5
      Caption         =   "Identitas Pribadi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FORM_UTAMA.frx":673A
      PicHot          =   "FORM_UTAMA.frx":6B8C
      PicNormal       =   "FORM_UTAMA.frx":6FDE
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmBukuAlamat 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2445
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
      BackColor       =   13815503
      ButtonStyle     =   5
      Caption         =   "Buku Alamat"
      Effects         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FORM_UTAMA.frx":7430
      PicHot          =   "FORM_UTAMA.frx":CC22
      PicNormal       =   "FORM_UTAMA.frx":12414
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmUlangTahun 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2820
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
      BackColor       =   13815503
      ButtonStyle     =   5
      Caption         =   "Ulang Tahun"
      Effects         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FORM_UTAMA.frx":17C06
      PicHot          =   "FORM_UTAMA.frx":19F88
      PicNormal       =   "FORM_UTAMA.frx":1C30A
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmAgenda 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3210
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
      BackColor       =   13815503
      ButtonStyle     =   5
      Caption         =   "Agenda"
      Effects         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FORM_UTAMA.frx":1E68C
      PicHot          =   "FORM_UTAMA.frx":1EADE
      PicNormal       =   "FORM_UTAMA.frx":1EF30
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmRegistrasiSoftware 
      Height          =   350
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   609
      BackColor       =   13815503
      ButtonStyle     =   5
      Caption         =   "Registrasi Software"
      Effects         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FORM_UTAMA.frx":1F382
      PicHot          =   "FORM_UTAMA.frx":26884
      PicNormal       =   "FORM_UTAMA.frx":2DD86
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmJejaringSosial 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4695
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
      BackColor       =   13815503
      ButtonStyle     =   5
      Caption         =   "Jejaring Sosial"
      Effects         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FORM_UTAMA.frx":35288
      PicHot          =   "FORM_UTAMA.frx":3C78A
      PicNormal       =   "FORM_UTAMA.frx":43C8C
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmElectronicMail 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5070
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
      BackColor       =   13815503
      ButtonStyle     =   5
      Caption         =   "Electronic Mail"
      Effects         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FORM_UTAMA.frx":4B18E
      PicHot          =   "FORM_UTAMA.frx":4B5E0
      PicNormal       =   "FORM_UTAMA.frx":4BA32
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmForumInternet 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   5460
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
      BackColor       =   13815503
      ButtonStyle     =   5
      Caption         =   "Forum Internet"
      Effects         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FORM_UTAMA.frx":4BE84
      PicHot          =   "FORM_UTAMA.frx":4C2D6
      PicNormal       =   "FORM_UTAMA.frx":4C728
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmFTP 
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   5850
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
      BackColor       =   13815503
      ButtonStyle     =   5
      Caption         =   "File Transfer Protocol"
      Effects         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FORM_UTAMA.frx":4CB7A
      PicHot          =   "FORM_UTAMA.frx":4CFCC
      PicNormal       =   "FORM_UTAMA.frx":4D41E
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmBlogging 
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   6240
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
      BackColor       =   13815503
      ButtonStyle     =   5
      Caption         =   "Blogging/Website"
      Effects         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FORM_UTAMA.frx":4D870
      PicHot          =   "FORM_UTAMA.frx":54D72
      PicNormal       =   "FORM_UTAMA.frx":5C274
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin MSAdodcLib.Adodc AdodcDataLogin 
      Height          =   330
      Left            =   120
      Top             =   7080
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
   Begin XPControls.XPFrame XPFrame1 
      Height          =   615
      Left            =   2400
      TabIndex        =   15
      Top             =   6840
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1085
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
      Begin Dacara_dcButton.dcButton cmBaru 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Baru"
         Effects         =   1
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
         PicDown         =   "FORM_UTAMA.frx":63776
         PicHot          =   "FORM_UTAMA.frx":63BC8
         PicNormal       =   "FORM_UTAMA.frx":6401A
         PicSize         =   1
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmManage 
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   120
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Manage"
         Effects         =   1
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
         PicDown         =   "FORM_UTAMA.frx":6446C
         PicHot          =   "FORM_UTAMA.frx":64786
         PicNormal       =   "FORM_UTAMA.frx":64AA0
         PicSize         =   1
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmRefresh 
         Height          =   375
         Left            =   3000
         TabIndex        =   18
         Top             =   120
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Refresh"
         Effects         =   1
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
         PicDown         =   "FORM_UTAMA.frx":64DBA
         PicHot          =   "FORM_UTAMA.frx":70E21
         PicNormal       =   "FORM_UTAMA.frx":7CE88
         PicSize         =   1
         PicSizeH        =   16
         PicSizeW        =   16
      End
   End
   Begin MSComctlLib.StatusBar StatusBawah 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   19
      Top             =   7545
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10054
            MinWidth        =   10054
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1323
            MinWidth        =   1323
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1323
            MinWidth        =   1323
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3775
            MinWidth        =   3775
         EndProperty
      EndProperty
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
   Begin XPEngine.XPControl XP_Engine 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.Menu menuMenu 
      Caption         =   "Menu"
      Begin VB.Menu menuBaru 
         Caption         =   "Baru"
         Begin VB.Menu menuBaruIP 
            Caption         =   "Identitas Pribadi"
         End
         Begin VB.Menu menuBaruBA 
            Caption         =   "Buku Alamat"
         End
         Begin VB.Menu menuBaruUT 
            Caption         =   "Ulang Tahun"
         End
         Begin VB.Menu menuBaruA 
            Caption         =   "Agenda"
         End
         Begin VB.Menu menuBaruRS 
            Caption         =   "Registrasi Software"
         End
         Begin VB.Menu sep3 
            Caption         =   "-"
         End
         Begin VB.Menu menuBaruJS 
            Caption         =   "Jejaring Sosial"
         End
         Begin VB.Menu menubaruEM 
            Caption         =   "Electronic Mail"
         End
         Begin VB.Menu menuFI 
            Caption         =   "Forum Internet"
         End
         Begin VB.Menu menuBaruFTP 
            Caption         =   "File Transfer Protocol"
         End
         Begin VB.Menu menuBaruBlogging 
            Caption         =   "Blogging/Website"
         End
      End
      Begin VB.Menu menuAkun 
         Caption         =   "Akun"
         Begin VB.Menu menuLogOut 
            Caption         =   "Log Out"
         End
      End
   End
   Begin VB.Menu menuView 
      Caption         =   "View"
      Begin VB.Menu menuBS 
         Caption         =   "Bar Status"
         Checked         =   -1  'True
      End
      Begin VB.Menu sep16 
         Caption         =   "-"
      End
      Begin VB.Menu menuTabelFilter 
         Caption         =   "Tabel Filter"
      End
      Begin VB.Menu menuRA 
         Caption         =   "Riwayat Aktivitas"
      End
   End
   Begin VB.Menu menuTools 
      Caption         =   "Tools"
      Begin VB.Menu menuASCG 
         Caption         =   "Access SQL Code Generator v2.0"
      End
      Begin VB.Menu menuAR 
         Caption         =   "Adress Register v1.0"
      End
      Begin VB.Menu menuES 
         Caption         =   "Encrypt String v1.3"
      End
      Begin VB.Menu sep15 
         Caption         =   "-"
      End
      Begin VB.Menu menuKV 
         Caption         =   "Keyboard Virtual"
      End
      Begin VB.Menu menuSN 
         Caption         =   "Sticky Notes"
      End
      Begin VB.Menu menuNotepad 
         Caption         =   "Notepad"
      End
      Begin VB.Menu sep18 
         Caption         =   "-"
      End
      Begin VB.Menu menuaPengaturan 
         Caption         =   "Pengaturan"
      End
   End
   Begin VB.Menu menuBantuan 
      Caption         =   "Bantuan"
      Begin VB.Menu menuPusatBantuan 
         Caption         =   "Pusat Bantuan"
      End
      Begin VB.Menu sep12 
         Caption         =   "-"
      End
      Begin VB.Menu menuKH 
         Caption         =   "Kunjungi HomePage"
      End
      Begin VB.Menu menuPVT 
         Caption         =   "Periksa Versi Terbaru.."
      End
      Begin VB.Menu sep14 
         Caption         =   "-"
      End
      Begin VB.Menu menuHubungiDeveloper 
         Caption         =   "Hubungi Developer"
      End
      Begin VB.Menu sep19 
         Caption         =   "-"
      End
      Begin VB.Menu menuTSAM 
         Caption         =   "Tentang - Simple Account Manager.."
      End
   End
   Begin VB.Menu MKK 
      Caption         =   "Menu Klik Kanan"
      Begin VB.Menu menuTampilkanPassword 
         Caption         =   "Tampilkan Password"
      End
   End
End
Attribute VB_Name = "FORM_UTAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    SambungkanADODC_UTAMA
    SambungkanADODC_DataLogin
    With AdodcDataLogin
        .ConnectionString = CN_FormUtamaLogin.ConnectionString
        .RecordSource = "Select * From tbDataLogin"
        .Refresh
    End With
    FrameAkunWeb.Enabled = False
    FrameAkunDesktop.Enabled = False
    AturStatusBawah
    MKK.Visible = False
    MatikanTombolNavigasi
    With StatusBawah.Panels
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            .Item(1).Text = "Database Siap!"
        Else
            .Item(1).Text = "Database Ready!"
        End If
    End With
    Me.Picture = LoadPicture(App.Path & "\image\banner_hitam.bmp")
    DisableCloseBtn Me
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
Sub PENGATURAN_FORM()
    If FormPengaturan.cekGarisGrid.Value = Checked Then LV.GridLines = True
    menuBS.Checked = GetSetting("rssamv1.0", "FormUtama", menuBS.Name, menuBS.Checked)
    'PENGATURAN UNTUK ALWAYS ON TOP
    If FormPengaturan.cekAlwaysOnTop.Value = Checked Then
        SetOnTop (Me.hwnd)
    ElseIf FormPengaturan.cekAlwaysOnTop.Value = Unchecked Then
        NotOnTop (Me.hwnd)
    End If
End Sub
Sub AturStatusBawah()
With StatusBawah.Panels
    .Item(1).Alignment = sbrLeft
    .Item(2).Alignment = sbrCenter
    .Item(3).Alignment = sbrCenter
    .Item(4).Alignment = sbrCenter
    .Item(5).Alignment = sbrCenter
    .Item(6).Alignment = sbrCenter
    .Item(1).ToolTipText = .Item(1).Text
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        .Item(2).ToolTipText = "Nama Pengguna"
        .Item(3).ToolTipText = "Jumlah Data"
        .Item(4).ToolTipText = "Jumlah Cell"
        .Item(5).ToolTipText = "Waktu Saat Ini"
        .Item(6).ToolTipText = "Tanggal Saat Ini"
    Else
        .Item(2).ToolTipText = "User Name"
        .Item(3).ToolTipText = "Data Count (Record)"
        .Item(4).ToolTipText = "Cell Count"
        .Item(5).ToolTipText = "Time is Now"
        .Item(6).ToolTipText = "Date is Now "
    End If
End With
End Sub
Sub MatikanTombolNavigasi()
    menuTampilkanPassword.Enabled = False
    cmBaru.Enabled = False
    cmManage.Enabled = False
    cmRefresh.Enabled = False
End Sub
Sub AktifkanTombolNavigasi()
    menuTampilkanPassword.Enabled = True
    cmBaru.Enabled = True
    cmManage.Enabled = True
    cmRefresh.Enabled = True
End Sub
Public Sub PENGATURAN_BAHASA()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        menuTampilkanPassword.Caption = "Tampilkan Password"
        cmAkunDesktop.Caption = "Akun Desktop"
        cmIdentitasPribadi.Caption = "Identitas Pribadi"
        cmBukuAlamat.Caption = "Buku Alamat"
        cmUlangTahun.Caption = "Ulang Tahun"
        cmAgenda.Caption = "Agenda"
        cmRegistrasiSoftware.Caption = "Serial Software"
        cmAkunWeb.Caption = "Akun Web"
        cmJejaringSosial.Caption = "Jejaring Sosial"
        cmElectronicMail.Caption = "E-Mail"
        cmForumInternet.Caption = "Forum Internet"
        cmFTP.Caption = "FTP"
        cmBlogging.Caption = "Blogging/Site"
        cmBaru.Caption = "&Baru"
        cmManage.Caption = "&Manage"
        cmRefresh.Caption = "&Refresh"
        
        menuBaru.Caption = "Baru"
        menuBaruIP.Caption = "Identitas Pribadi"
        menuBaruBA.Caption = "Buku Alamat"
        menuBaruUT.Caption = "Ulang Tahun"
        menuBaruA.Caption = "Agenda"
        menuBaruRS.Caption = "Registrasi Software"
        menuBaruJS.Caption = "Jejaring Sosial"
        menubaruEM.Caption = "Electronic Mail"
        menuFI.Caption = "Forum Internet"
        menuBaruFTP.Caption = "File Transfer Protocol"
        menuBaruBlogging.Caption = "Blogging/Website"
        menuAkun.Caption = "Akun"
        menuLogOut.Caption = "Log Out"
        menuBS.Caption = "Bar Status"
        menuTabelFilter.Caption = "Tabel Filter"
        menuRA.Caption = "Riwayat Aktivitas"
        menuaPengaturan.Caption = "Pengaturan"
        menuPusatBantuan.Caption = "Pusat Bantuan"
        menuKH.Caption = "Kunjungi Homepage"
        menuPVT.Caption = "Periksa Versi Terbaru"
        menuHubungiDeveloper.Caption = "Hubungi Developer"
        menuTSAM.Caption = "Tentang - Simple Account Manager"
        menuBantuan.Caption = "Bantuan"
        
    Else
        menuTampilkanPassword.Caption = "Show Passwords"
        cmAkunDesktop.Caption = "Desktop Account"
        cmIdentitasPribadi.Caption = "Personal Biodata"
        cmBukuAlamat.Caption = "Address Book"
        cmUlangTahun.Caption = "Birthday"
        cmAgenda.Caption = "Agenda"
        cmRegistrasiSoftware.Caption = "Software Serial"
        cmAkunWeb.Caption = "Web Account"
        cmJejaringSosial.Caption = "Social Network"
        cmElectronicMail.Caption = "Electronic Mail"
        cmForumInternet.Caption = "Internet Forum"
        cmFTP.Caption = "FTP"
        cmBlogging.Caption = "Blogging/Site"
        cmBaru.Caption = "&New"
        cmManage.Caption = "&Manage"
        cmRefresh.Caption = "&Refresh"
    
        menuBaru.Caption = "New"
        menuBaruIP.Caption = "Personal Identity"
        menuBaruBA.Caption = "Address Book"
        menuBaruUT.Caption = "Birthday"
        menuBaruA.Caption = "Agenda"
        menuBaruRS.Caption = "Software Serial"
        menuBaruJS.Caption = "Social Network"
        menubaruEM.Caption = "Electronic Mail"
        menuFI.Caption = "Internet Forums"
        menuBaruFTP.Caption = "File Transfer Protocol"
        menuBaruBlogging.Caption = "Blogging/Website"
        menuAkun.Caption = "Account"
        menuLogOut.Caption = "Log Out"
        menuBS.Caption = "Status Bar"
        menuTabelFilter.Caption = "Filter Table"
        menuRA.Caption = "History Activity"
        menuaPengaturan.Caption = "Settings"
        menuPusatBantuan.Caption = "Help Center"
        menuKH.Caption = "Visit Website"
        menuPVT.Caption = "Check for Update"
        menuHubungiDeveloper.Caption = "Contact Developer"
        menuTSAM.Caption = "About - Simple Account Manager"
        menuBantuan.Caption = "Help"
    End If
End Sub

Public Sub cmAgenda_Click()
    menuTabelFilter.Enabled = True
    cmIdentitasPribadi.FontBold = False
    cmBukuAlamat.FontBold = False
    cmUlangTahun.FontBold = False
    cmAgenda.FontBold = True
    cmRegistrasiSoftware.FontBold = False
    cmJejaringSosial.FontBold = False
    cmElectronicMail.FontBold = False
    cmForumInternet.FontBold = False
    cmFTP.FontBold = False
    cmBlogging.FontBold = False

    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Kode Agenda", 2000
            .ColumnHeaders.Add , , "Nama Agenda", 2000, vbCenter
            .ColumnHeaders.Add , , "Tema", 2000, vbCenter
            .ColumnHeaders.Add , , "Tanggal", 2000, vbCenter
            .ColumnHeaders.Add , , "Waktu Mulai", 2000, vbCenter
            .ColumnHeaders.Add , , "Waktu Akhir", 2000, vbCenter
            .ColumnHeaders.Add , , "Tempat", 2000, vbCenter
            .ColumnHeaders.Add , , "Keterangan Lain", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    Else
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Code", 2000
            .ColumnHeaders.Add , , "Agenda Name", 2000, vbCenter
            .ColumnHeaders.Add , , "Thema", 2000, vbCenter
            .ColumnHeaders.Add , , "Date", 2000, vbCenter
            .ColumnHeaders.Add , , "Begin Time", 2000, vbCenter
            .ColumnHeaders.Add , , "End Time", 2000, vbCenter
            .ColumnHeaders.Add , , "Place", 2000, vbCenter
            .ColumnHeaders.Add , , "Other Description", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    End If
        If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
            CN_FormUtama.CursorLocation = adUseClient
            CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & Me.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
        With ADODC_UTAMA
            .ConnectionString = CN_FormUtama.ConnectionString
            .RecordSource = "Select * From tbAgenda order by Nama_Agenda asc;"
            .Refresh
        End With
        LV.ListItems.Clear
        Do Until ADODC_UTAMA.Recordset.EOF
        Set LI = LV.ListItems.Add(, , ADODC_UTAMA.Recordset.Fields(0).Value)
            LI.SubItems(1) = ADODC_UTAMA.Recordset.Fields(1).Value
            LI.SubItems(2) = ADODC_UTAMA.Recordset.Fields(2).Value
            LI.SubItems(3) = ADODC_UTAMA.Recordset.Fields(3).Value
            LI.SubItems(4) = ADODC_UTAMA.Recordset.Fields(4).Value
            LI.SubItems(5) = ADODC_UTAMA.Recordset.Fields(5).Value
            LI.SubItems(6) = ADODC_UTAMA.Recordset.Fields(6).Value
            LI.SubItems(7) = ADODC_UTAMA.Recordset.Fields(7).Value
            ADODC_UTAMA.Recordset.MoveNext
        Loop
        ADODC_UTAMA.Refresh
        Me.Caption = "Simple Accounts Manager - " & AdodcDataLogin.Recordset.Fields(2).Value & " (" & cmAgenda.Caption & ")"
        AktifkanTombolNavigasi
        With StatusBawah.Panels
            .Item(1).Text = cmAgenda.Caption
            .Item(3).Text = ADODC_UTAMA.Recordset.RecordCount
            .Item(4).Text = ADODC_UTAMA.Recordset.RecordCount * ADODC_UTAMA.Recordset.Fields.Count
        End With
        AturStatusBawah
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            menuTampilkanPassword.Caption = "Tampilkan Password"
        Else
            menuTampilkanPassword.Caption = "Show Passwords"
        End If
            menuTampilkanPassword.Enabled = False
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
                    .Recordset.Fields(6).Value = "Data Agenda ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The Agenda files viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
End Sub

Public Sub cmBaru_Click()
On Error GoTo HancurkanError
If cmJejaringSosial.FontBold = True Then
    With Form_JEJARING_SOSIAL
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            .Caption = "Jejaring Sosial - Data Baru"
            .Command1.Picture = LoadPicture(App.Path & "\image\BannerJejaringSosialID.bmp")
        Else
            .Caption = "New Data - Social Network"
            .Command1.Picture = LoadPicture(App.Path & "\image\BannerJejaringSosialEN.bmp")
        End If
        .Show vbModal, Me
    End With
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
                    .Recordset.Fields(6).Value = "Jendela input Jejaring Sosial ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The input form of Social Network is viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
ElseIf cmElectronicMail.FontBold = True Then
    With Form_ELECTRONIC_MAIL
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            .Caption = "Electronic Mail - Data Baru"
        Else
            .Caption = "New Data - Electronic Mail"
        End If
        .Show vbModal, Me
    End With
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
                    .Recordset.Fields(6).Value = "Jendela input Electronic Mail ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The input form of Electronic Mail is viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
ElseIf cmForumInternet.FontBold = True Then
    With Form_FORUM_INTERNET
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            .Caption = "Forum Internet - Data Baru"
            .Command1.Picture = LoadPicture(App.Path & "\image\BannerForumInternet_ID.bmp")
        Else
            .Caption = "New Data - Internet Forum"
            .Command1.Picture = LoadPicture(App.Path & "\image\BannerForumInternet_EN.bmp")
        End If
        .Show vbModal, Me
    End With
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
                    .Recordset.Fields(6).Value = "Jendela input Forum Internet ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The input form of Internet Forum is viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
ElseIf cmFTP.FontBold = True Then
    With Form_FTP
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            .Caption = "File Transfer Protocol (FTP) - Data Baru"
        Else
            .Caption = "New Data - File Transfer Protocol (FTP)"
        End If
        .Show vbModal, Me
    End With
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
                    .Recordset.Fields(6).Value = "Jendela input FTP ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The input form of FTP is viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
ElseIf cmBlogging.FontBold = True Then
    With Form_BLOGGING_WEBSITE
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            .Caption = "Blogging/Website - Data Baru"
        Else
            .Caption = "New Data - Blogging/Website"
        End If
        .Show vbModal, Me
    End With
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
                    .Recordset.Fields(6).Value = "Jendela input Blogging/Website ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The input form of Blogging/Website is viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
ElseIf cmBukuAlamat.FontBold = True Then
    With Form_BUKU_ALAMAT
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            .Caption = "Buku Alamat - Data Baru"
            .Command1.Picture = LoadPicture(App.Path & "\image\BannerBukuAlamat_ID.bmp")
        Else
            .Caption = "New Data - Address Book"
            .Command1.Picture = LoadPicture(App.Path & "\image\BannerBukuAlamat_EN.bmp")
        End If
        .Show vbModal, Me
    End With
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
                    .Recordset.Fields(6).Value = "Jendela input Buku Alamat ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The input form of Address Book is viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
ElseIf cmUlangTahun.FontBold = True Then
    With Form_ULANG_TAHUN
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            .Caption = "Ulang Tahun - Data Baru"
            .Command1.Picture = LoadPicture(App.Path & "\image\BannerUlangTahun_ID.bmp")
        Else
            .Caption = "New Data - Birthday"
            .Command1.Picture = LoadPicture(App.Path & "\image\BannerUlangTahun_EN.bmp")
        End If
        .Show vbModal, Me
    End With
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
                    .Recordset.Fields(6).Value = "Jendela input Ulang Tahun ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The input form of Birthday is viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
ElseIf cmAgenda.FontBold = True Then
    With Form_AGENDA
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            .Caption = "Agenda - Data Baru"
        Else
            .Caption = "New Data - Agenda"
        End If
        .Show vbModal, Me
    End With
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
                    .Recordset.Fields(6).Value = "Jendela input Agenda ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The input form of Agenda is viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
ElseIf cmRegistrasiSoftware.FontBold = True Then
    With Form_REGISTRASI_SOFTWARE
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            .Caption = "Software Serial - Data Baru"
            .Command1.Picture = LoadPicture(App.Path & "\image\BannerRegistrasiSoftware_ID.bmp")
        Else
            .Caption = "New Data - Software Serial"
            .Command1.Picture = LoadPicture(App.Path & "\image\BannerRegistrasiSoftware_EN.bmp")
        End If
        .Show vbModal, Me
    End With
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
                    .Recordset.Fields(6).Value = "Jendela input Registrasi Software ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The input form of Software Registration is viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
ElseIf cmIdentitasPribadi.FontBold = True Then
    With Form_IDENTITAS_PRIBADI
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            .Caption = "Identitas Pribadi - Data Baru"
            .Command1.Picture = LoadPicture(App.Path & "\image\BannerIdenditasPribadi_ID.bmp")
        Else
            .Caption = "New Data - Personal Biodata"
            .Command1.Picture = LoadPicture(App.Path & "\image\BannerIdenditasPribadi_EN.bmp")
        End If
        .Show vbModal, Me
    End With
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
                    .Recordset.Fields(6).Value = "Jendela input Identitas Pribadi ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The input form of Personal Biodata is viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
End If
Exit Sub
HancurkanError:
    PusatError
End Sub

Public Sub cmBlogging_Click()
    menuTabelFilter.Enabled = False
    cmIdentitasPribadi.FontBold = False
    cmBukuAlamat.FontBold = False
    cmUlangTahun.FontBold = False
    cmAgenda.FontBold = False
    cmRegistrasiSoftware.FontBold = False
    cmJejaringSosial.FontBold = False
    cmElectronicMail.FontBold = False
    cmForumInternet.FontBold = False
    cmFTP.FontBold = False
    cmBlogging.FontBold = True

    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Nama Penyedia Blog", 2500
            .ColumnHeaders.Add , , "Nama Pengguna", 2000, vbCenter
            .ColumnHeaders.Add , , "E-Mail", 2000, vbCenter
            .ColumnHeaders.Add , , "URL", 2000, vbCenter
            .ColumnHeaders.Add , , "Tanggal", 2000, vbCenter
            .ColumnHeaders.Add , , "Keterangan", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    Else
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Blog Providers Name", 2500
            .ColumnHeaders.Add , , "User Name", 2000, vbCenter
            .ColumnHeaders.Add , , "E-Mail", 2000, vbCenter
            .ColumnHeaders.Add , , "URL", 2000, vbCenter
            .ColumnHeaders.Add , , "Date", 2000, vbCenter
            .ColumnHeaders.Add , , "Description", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    End If
        If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
            CN_FormUtama.CursorLocation = adUseClient
            CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & Me.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
        With ADODC_UTAMA
            .ConnectionString = CN_FormUtama.ConnectionString
            .RecordSource = "Select Nama_Penyedia_Blog,Nama_Pengguna,E_Mail,URL,Tanggal,Keterangan From tbBlogging order by Nama_Penyedia_Blog asc;"
            .Refresh
        End With
        LV.ListItems.Clear
        Do Until ADODC_UTAMA.Recordset.EOF
        Set LI = LV.ListItems.Add(, , ADODC_UTAMA.Recordset.Fields(0).Value)
            LI.SubItems(1) = ADODC_UTAMA.Recordset.Fields(1).Value
            LI.SubItems(2) = ADODC_UTAMA.Recordset.Fields(2).Value
            LI.SubItems(3) = ADODC_UTAMA.Recordset.Fields(3).Value
            LI.SubItems(4) = ADODC_UTAMA.Recordset.Fields(4).Value
            LI.SubItems(5) = ADODC_UTAMA.Recordset.Fields(5).Value
            ADODC_UTAMA.Recordset.MoveNext
        Loop
        ADODC_UTAMA.Refresh
        Me.Caption = "Simple Accounts Manager - " & AdodcDataLogin.Recordset.Fields(2).Value & " (" & cmBlogging.Caption & ")"
        AktifkanTombolNavigasi
        With StatusBawah.Panels
            .Item(1).Text = cmBlogging.Caption
            .Item(3).Text = ADODC_UTAMA.Recordset.RecordCount
            .Item(4).Text = ADODC_UTAMA.Recordset.RecordCount * ADODC_UTAMA.Recordset.Fields.Count
        End With
        AturStatusBawah
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            menuTampilkanPassword.Caption = "Tampilkan Password"
        Else
            menuTampilkanPassword.Caption = "Show Passwords"
        End If
            menuTampilkanPassword.Enabled = True
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
                    .Recordset.Fields(6).Value = "Data blogging/website ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The blogging/website files viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
End Sub

Public Sub cmBukuAlamat_Click()
    menuTabelFilter.Enabled = True
    cmIdentitasPribadi.FontBold = False
    cmBukuAlamat.FontBold = True
    cmUlangTahun.FontBold = False
    cmAgenda.FontBold = False
    cmRegistrasiSoftware.FontBold = False
    cmJejaringSosial.FontBold = False
    cmElectronicMail.FontBold = False
    cmForumInternet.FontBold = False
    cmFTP.FontBold = False
    cmBlogging.FontBold = False

    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Nama Kontak", 2000
            .ColumnHeaders.Add , , "Nama Panggilan", 2000, vbCenter
            .ColumnHeaders.Add , , "Nomor Telepon Pribadi", 2500, vbCenter
            .ColumnHeaders.Add , , "Nomor Telepon Rumah", 2500, vbCenter
            .ColumnHeaders.Add , , "Nomor Telepon Kantor", 2500, vbCenter
            .ColumnHeaders.Add , , "Fax", 2000, vbCenter
            .ColumnHeaders.Add , , "Alamat E-Mail", 2000, vbCenter
            .ColumnHeaders.Add , , "Website", 2000, vbCenter
            .ColumnHeaders.Add , , "Kode Pos", 2000, vbCenter
            .ColumnHeaders.Add , , "Alamat Rumah", 2000, vbCenter
            .ColumnHeaders.Add , , "Keterangan", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    Else
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Contact Name", 2000
            .ColumnHeaders.Add , , "Cool Name", 2000, vbCenter
            .ColumnHeaders.Add , , "Private Phone Number", 2500, vbCenter
            .ColumnHeaders.Add , , "House Phone Number", 2500, vbCenter
            .ColumnHeaders.Add , , "Office Phone Number", 2500, vbCenter
            .ColumnHeaders.Add , , "Fax", 2000, vbCenter
            .ColumnHeaders.Add , , "E-Mail Address", 2000, vbCenter
            .ColumnHeaders.Add , , "Website", 2000, vbCenter
            .ColumnHeaders.Add , , "Zip/Postal Code", 2000, vbCenter
            .ColumnHeaders.Add , , "Home Address", 2000, vbCenter
            .ColumnHeaders.Add , , "Description", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    End If
        If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
            CN_FormUtama.CursorLocation = adUseClient
            CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & Me.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
        With ADODC_UTAMA
            .ConnectionString = CN_FormUtama.ConnectionString
            .RecordSource = "Select * From tbBukuAlamat order by Nama_Kontak asc;"
            .Refresh
        End With
        LV.ListItems.Clear
        Do Until ADODC_UTAMA.Recordset.EOF
        Set LI = LV.ListItems.Add(, , ADODC_UTAMA.Recordset.Fields(0).Value)
            LI.SubItems(1) = ADODC_UTAMA.Recordset.Fields(1).Value
            LI.SubItems(2) = ADODC_UTAMA.Recordset.Fields(2).Value
            LI.SubItems(3) = ADODC_UTAMA.Recordset.Fields(3).Value
            LI.SubItems(4) = ADODC_UTAMA.Recordset.Fields(4).Value
            LI.SubItems(5) = ADODC_UTAMA.Recordset.Fields(5).Value
            LI.SubItems(6) = ADODC_UTAMA.Recordset.Fields(6).Value
            LI.SubItems(7) = ADODC_UTAMA.Recordset.Fields(7).Value
            LI.SubItems(8) = ADODC_UTAMA.Recordset.Fields(8).Value
            LI.SubItems(9) = ADODC_UTAMA.Recordset.Fields(9).Value
            LI.SubItems(10) = ADODC_UTAMA.Recordset.Fields(10).Value
            ADODC_UTAMA.Recordset.MoveNext
        Loop
        ADODC_UTAMA.Refresh
        Me.Caption = "Simple Accounts Manager - " & AdodcDataLogin.Recordset.Fields(2).Value & " (" & cmBukuAlamat.Caption & ")"
        AktifkanTombolNavigasi
        With StatusBawah.Panels
            .Item(1).Text = cmBukuAlamat.Caption
            .Item(3).Text = ADODC_UTAMA.Recordset.RecordCount
            .Item(4).Text = ADODC_UTAMA.Recordset.RecordCount * ADODC_UTAMA.Recordset.Fields.Count
        End With
        AturStatusBawah
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            menuTampilkanPassword.Caption = "Tampilkan Password"
        Else
            menuTampilkanPassword.Caption = "Show Passwords"
        End If
            menuTampilkanPassword.Enabled = False
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
                    .Recordset.Fields(6).Value = "Data Buku alamat ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The Address Book files viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
End Sub

Public Sub cmElectronicMail_Click()
    menuTabelFilter.Enabled = False
    cmIdentitasPribadi.FontBold = False
    cmBukuAlamat.FontBold = False
    cmUlangTahun.FontBold = False
    cmAgenda.FontBold = False
    cmRegistrasiSoftware.FontBold = False
    cmJejaringSosial.FontBold = False
    cmElectronicMail.FontBold = True
    cmForumInternet.FontBold = False
    cmFTP.FontBold = False
    cmBlogging.FontBold = False
    
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Nama Server", 2000
            .ColumnHeaders.Add , , "Nama Pengguna", 2000, vbCenter
            .ColumnHeaders.Add , , "Alamat E-Mail", 2000, vbCenter
            .ColumnHeaders.Add , , "Pertanyaan Rahasia", 2000, vbCenter
            .ColumnHeaders.Add , , "Jawaban Pertanyaan", 2000, vbCenter
            .ColumnHeaders.Add , , "URL", 2000, vbCenter
            .ColumnHeaders.Add , , "Pemilik Akun", 2000, vbCenter
            .ColumnHeaders.Add , , "Tanggal", 2000, vbCenter
            .ColumnHeaders.Add , , "Keterangan", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    Else
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Server Name", 2000
            .ColumnHeaders.Add , , "User Name", 2000, vbCenter
            .ColumnHeaders.Add , , "Mail Address", 2000, vbCenter
            .ColumnHeaders.Add , , "Security Question", 2000, vbCenter
            .ColumnHeaders.Add , , "Security Answer", 2000, vbCenter
            .ColumnHeaders.Add , , "URL", 2000, vbCenter
            .ColumnHeaders.Add , , "Account Owner", 2000, vbCenter
            .ColumnHeaders.Add , , "Date", 2000, vbCenter
            .ColumnHeaders.Add , , "Description", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    End If
        If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
            CN_FormUtama.CursorLocation = adUseClient
            CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & Me.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
        With ADODC_UTAMA
            .ConnectionString = CN_FormUtama.ConnectionString
            .RecordSource = "Select Nama_Server,Nama_Pengguna,Alamat_Email,Pertanyaan_Rahasia,Jawaban_Pertanyaan,URL,Pemilik_Akun,Tanggal,Keterangan From tbElectronicMail order by Nama_Server asc;"
            .Refresh
        End With
        LV.ListItems.Clear
        Do Until ADODC_UTAMA.Recordset.EOF
        Set LI = LV.ListItems.Add(, , ADODC_UTAMA.Recordset.Fields(0).Value)
            LI.SubItems(1) = ADODC_UTAMA.Recordset.Fields(1).Value
            LI.SubItems(2) = ADODC_UTAMA.Recordset.Fields(2).Value
            LI.SubItems(3) = ADODC_UTAMA.Recordset.Fields(3).Value
            LI.SubItems(4) = ADODC_UTAMA.Recordset.Fields(4).Value
            LI.SubItems(5) = ADODC_UTAMA.Recordset.Fields(5).Value
            LI.SubItems(6) = ADODC_UTAMA.Recordset.Fields(6).Value
            LI.SubItems(7) = ADODC_UTAMA.Recordset.Fields(7).Value
            LI.SubItems(8) = ADODC_UTAMA.Recordset.Fields(8).Value
            ADODC_UTAMA.Recordset.MoveNext
        Loop
        ADODC_UTAMA.Refresh
        Me.Caption = "Simple Accounts Manager - " & AdodcDataLogin.Recordset.Fields(2).Value & " (" & cmElectronicMail.Caption & ")"
        AktifkanTombolNavigasi
        With StatusBawah.Panels
            .Item(1).Text = cmElectronicMail.Caption
            .Item(3).Text = ADODC_UTAMA.Recordset.RecordCount
            .Item(4).Text = ADODC_UTAMA.Recordset.RecordCount * ADODC_UTAMA.Recordset.Fields.Count
        End With
        AturStatusBawah
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            menuTampilkanPassword.Caption = "Tampilkan Password"
        Else
            menuTampilkanPassword.Caption = "Show Passwords"
        End If
            menuTampilkanPassword.Enabled = True
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
                    .Recordset.Fields(6).Value = "Data Electronic Mail ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The Electronic Mail files viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
End Sub

Public Sub cmForumInternet_Click()
    menuTabelFilter.Enabled = False
    cmIdentitasPribadi.FontBold = False
    cmBukuAlamat.FontBold = False
    cmUlangTahun.FontBold = False
    cmAgenda.FontBold = False
    cmRegistrasiSoftware.FontBold = False
    cmJejaringSosial.FontBold = False
    cmElectronicMail.FontBold = False
    cmForumInternet.FontBold = True
    cmFTP.FontBold = False
    cmBlogging.FontBold = False

    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Nama Forum", 2000
            .ColumnHeaders.Add , , "Nama Pengguna", 2000, vbCenter
            .ColumnHeaders.Add , , "Alamat E-Mail", 2000, vbCenter
            .ColumnHeaders.Add , , "Jabatan", 2000, vbCenter
            .ColumnHeaders.Add , , "NickName", 2000, vbCenter
            .ColumnHeaders.Add , , "URL", 2000, vbCenter
            .ColumnHeaders.Add , , "Tanggal", 2000, vbCenter
            .ColumnHeaders.Add , , "Keterangan", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    Else
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Forum Name", 2000
            .ColumnHeaders.Add , , "User Name", 2000, vbCenter
            .ColumnHeaders.Add , , "Mail Address", 2000, vbCenter
            .ColumnHeaders.Add , , "Position", 2000, vbCenter
            .ColumnHeaders.Add , , "NickName", 2000, vbCenter
            .ColumnHeaders.Add , , "URL", 2000, vbCenter
            .ColumnHeaders.Add , , "Date", 2000, vbCenter
            .ColumnHeaders.Add , , "Description", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    End If
        If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
            CN_FormUtama.CursorLocation = adUseClient
            CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & Me.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
        With ADODC_UTAMA
            .ConnectionString = CN_FormUtama.ConnectionString
            .RecordSource = "Select Nama_Forum,Nama_Pengguna,Alamat_Email,Posisi,NickName,URL,Tanggal,Keterangan From tbForumInternet order by Nama_Forum asc;"
            .Refresh
        End With
        LV.ListItems.Clear
        Do Until ADODC_UTAMA.Recordset.EOF
        Set LI = LV.ListItems.Add(, , ADODC_UTAMA.Recordset.Fields(0).Value)
            LI.SubItems(1) = ADODC_UTAMA.Recordset.Fields(1).Value
            LI.SubItems(2) = ADODC_UTAMA.Recordset.Fields(2).Value
            LI.SubItems(3) = ADODC_UTAMA.Recordset.Fields(3).Value
            LI.SubItems(4) = ADODC_UTAMA.Recordset.Fields(4).Value
            LI.SubItems(5) = ADODC_UTAMA.Recordset.Fields(5).Value
            LI.SubItems(6) = ADODC_UTAMA.Recordset.Fields(6).Value
            LI.SubItems(7) = ADODC_UTAMA.Recordset.Fields(7).Value
            ADODC_UTAMA.Recordset.MoveNext
        Loop
        ADODC_UTAMA.Refresh
        Me.Caption = "Simple Accounts Manager - " & AdodcDataLogin.Recordset.Fields(2).Value & " (" & cmForumInternet.Caption & ")"
        AktifkanTombolNavigasi
        With StatusBawah.Panels
            .Item(1).Text = cmForumInternet.Caption
            .Item(3).Text = ADODC_UTAMA.Recordset.RecordCount
            .Item(4).Text = ADODC_UTAMA.Recordset.RecordCount * ADODC_UTAMA.Recordset.Fields.Count
        End With
        AturStatusBawah
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            menuTampilkanPassword.Caption = "Tampilkan Password"
        Else
            menuTampilkanPassword.Caption = "Show Passwords"
        End If
            menuTampilkanPassword.Enabled = True
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
                    .Recordset.Fields(6).Value = "Data forum internet ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The internet forum files viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
End Sub

Public Sub cmFTP_Click()
    menuTabelFilter.Enabled = False
    cmIdentitasPribadi.FontBold = False
    cmBukuAlamat.FontBold = False
    cmUlangTahun.FontBold = False
    cmAgenda.FontBold = False
    cmRegistrasiSoftware.FontBold = False
    cmJejaringSosial.FontBold = False
    cmElectronicMail.FontBold = False
    cmForumInternet.FontBold = False
    cmFTP.FontBold = True
    cmBlogging.FontBold = False

    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Nama Host", 2000
            .ColumnHeaders.Add , , "Nomor Port", 2000, vbCenter
            .ColumnHeaders.Add , , "Nama Server", 2000, vbCenter
            .ColumnHeaders.Add , , "Nama Pengguna", 2000, vbCenter
            .ColumnHeaders.Add , , "E-Mail", 2000, vbCenter
            .ColumnHeaders.Add , , "Tanggal", 2000, vbCenter
            .ColumnHeaders.Add , , "Keterangan", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    Else
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Host Name", 2000
            .ColumnHeaders.Add , , "Port Number", 2000, vbCenter
            .ColumnHeaders.Add , , "Server Name", 2000, vbCenter
            .ColumnHeaders.Add , , "User Name", 2000, vbCenter
            .ColumnHeaders.Add , , "E-Mail", 2000, vbCenter
            .ColumnHeaders.Add , , "Date", 2000, vbCenter
            .ColumnHeaders.Add , , "Description", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    End If
        If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
            CN_FormUtama.CursorLocation = adUseClient
            CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & Me.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
        With ADODC_UTAMA
            .ConnectionString = CN_FormUtama.ConnectionString
            .RecordSource = "Select Nama_Host,Port,Nama_Server,Nama_Pengguna,Alamat_Email,Tanggal,Keterangan From tbFTP order by Nama_Host asc;"
            .Refresh
        End With
        LV.ListItems.Clear
        Do Until ADODC_UTAMA.Recordset.EOF
        Set LI = LV.ListItems.Add(, , ADODC_UTAMA.Recordset.Fields(0).Value)
            LI.SubItems(1) = ADODC_UTAMA.Recordset.Fields(1).Value
            LI.SubItems(2) = ADODC_UTAMA.Recordset.Fields(2).Value
            LI.SubItems(3) = ADODC_UTAMA.Recordset.Fields(3).Value
            LI.SubItems(4) = ADODC_UTAMA.Recordset.Fields(4).Value
            LI.SubItems(5) = ADODC_UTAMA.Recordset.Fields(5).Value
            LI.SubItems(6) = ADODC_UTAMA.Recordset.Fields(6).Value
            ADODC_UTAMA.Recordset.MoveNext
        Loop
        ADODC_UTAMA.Refresh
        Me.Caption = "Simple Accounts Manager - " & AdodcDataLogin.Recordset.Fields(2).Value & " (" & cmFTP.Caption & ")"
        AktifkanTombolNavigasi
        With StatusBawah.Panels
            .Item(1).Text = cmFTP.Caption
            .Item(3).Text = ADODC_UTAMA.Recordset.RecordCount
            .Item(4).Text = ADODC_UTAMA.Recordset.RecordCount * ADODC_UTAMA.Recordset.Fields.Count
        End With
        AturStatusBawah
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            menuTampilkanPassword.Caption = "Tampilkan Password"
        Else
            menuTampilkanPassword.Caption = "Show Passwords"
        End If
            menuTampilkanPassword.Enabled = True
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
                    .Recordset.Fields(6).Value = "Data file transfer protocol ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The FTP files viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
End Sub

Public Sub cmIdentitasPribadi_Click()
    menuTabelFilter.Enabled = True
    cmIdentitasPribadi.FontBold = True
    cmBukuAlamat.FontBold = False
    cmUlangTahun.FontBold = False
    cmAgenda.FontBold = False
    cmRegistrasiSoftware.FontBold = False
    cmJejaringSosial.FontBold = False
    cmElectronicMail.FontBold = False
    cmForumInternet.FontBold = False
    cmFTP.FontBold = False
    cmBlogging.FontBold = False

    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Nama Lengkap", 2000
            .ColumnHeaders.Add , , "Nama Panggilan", 2000, vbCenter
            .ColumnHeaders.Add , , "TTL", 2500, vbCenter
            .ColumnHeaders.Add , , "Jenis Kelamin", 2000, vbCenter
            .ColumnHeaders.Add , , "Agama", 2000, vbCenter
            .ColumnHeaders.Add , , "Golongan Darah", 2000, vbCenter
            .ColumnHeaders.Add , , "Pekerjaan", 2000, vbCenter
            .ColumnHeaders.Add , , "Alamat Rumah", 2000, vbCenter
            .ColumnHeaders.Add , , "E-Mail", 2000, vbCenter
            .ColumnHeaders.Add , , "Website", 2000, vbCenter
            .ColumnHeaders.Add , , "Nomor Telepon", 2000, vbCenter
            .ColumnHeaders.Add , , "Kota Asal", 2000, vbCenter
            .ColumnHeaders.Add , , "Kota Sekarang", 2000, vbCenter
            .ColumnHeaders.Add , , "Kode Pos", 2000, vbCenter
            .ColumnHeaders.Add , , "Provinsi", 2000, vbCenter
            .ColumnHeaders.Add , , "KewargaNegaraan", 2000, vbCenter
            .ColumnHeaders.Add , , "Status Pendidikan", 2000, vbCenter
            .ColumnHeaders.Add , , "Status Hubungan", 2000, vbCenter
            .ColumnHeaders.Add , , "Hobby", 2000, vbCenter
            .ColumnHeaders.Add , , "Keterangan", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    Else
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Original Name", 2000
            .ColumnHeaders.Add , , "Cool Name", 2000, vbCenter
            .ColumnHeaders.Add , , "Place/Born Date", 2500, vbCenter
            .ColumnHeaders.Add , , "Gender", 2000, vbCenter
            .ColumnHeaders.Add , , "Religion", 2000, vbCenter
            .ColumnHeaders.Add , , "Blood Type", 2000, vbCenter
            .ColumnHeaders.Add , , "Jobs", 2000, vbCenter
            .ColumnHeaders.Add , , "Home Address", 2000, vbCenter
            .ColumnHeaders.Add , , "E-Mail", 2000, vbCenter
            .ColumnHeaders.Add , , "Website", 2000, vbCenter
            .ColumnHeaders.Add , , "Phone Number", 2000, vbCenter
            .ColumnHeaders.Add , , "Home Town", 2000, vbCenter
            .ColumnHeaders.Add , , "City Now", 2000, vbCenter
            .ColumnHeaders.Add , , "Zip/Postal Code", 2000, vbCenter
            .ColumnHeaders.Add , , "State", 2000, vbCenter
            .ColumnHeaders.Add , , "Citizenship", 2000, vbCenter
            .ColumnHeaders.Add , , "Educational Status", 2000, vbCenter
            .ColumnHeaders.Add , , "Relationship Status", 2000, vbCenter
            .ColumnHeaders.Add , , "Hobbies", 2000, vbCenter
            .ColumnHeaders.Add , , "Description", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    End If
        If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
            CN_FormUtama.CursorLocation = adUseClient
            CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & Me.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
        With ADODC_UTAMA
            .ConnectionString = CN_FormUtama.ConnectionString
            .RecordSource = "Select * From tbIdentitasPribadi order by Nama_Lengkap asc;"
            .Refresh
        End With
        LV.ListItems.Clear
        Do Until ADODC_UTAMA.Recordset.EOF
        Set LI = LV.ListItems.Add(, , ADODC_UTAMA.Recordset.Fields(0).Value)
            LI.SubItems(1) = ADODC_UTAMA.Recordset.Fields(1).Value
            LI.SubItems(2) = ADODC_UTAMA.Recordset.Fields(2).Value & ", " & ADODC_UTAMA.Recordset.Fields(3).Value & "-" & ADODC_UTAMA.Recordset.Fields(4).Value & "-" & ADODC_UTAMA.Recordset.Fields(5).Value
            LI.SubItems(3) = ADODC_UTAMA.Recordset.Fields(6).Value
            LI.SubItems(4) = ADODC_UTAMA.Recordset.Fields(7).Value
            LI.SubItems(5) = ADODC_UTAMA.Recordset.Fields(8).Value
            LI.SubItems(6) = ADODC_UTAMA.Recordset.Fields(9).Value
            LI.SubItems(7) = ADODC_UTAMA.Recordset.Fields(10).Value
            LI.SubItems(8) = ADODC_UTAMA.Recordset.Fields(11).Value
            LI.SubItems(9) = ADODC_UTAMA.Recordset.Fields(12).Value
            LI.SubItems(10) = ADODC_UTAMA.Recordset.Fields(13).Value
            LI.SubItems(11) = ADODC_UTAMA.Recordset.Fields(14).Value
            LI.SubItems(12) = ADODC_UTAMA.Recordset.Fields(15).Value
            LI.SubItems(13) = ADODC_UTAMA.Recordset.Fields(16).Value
            LI.SubItems(14) = ADODC_UTAMA.Recordset.Fields(17).Value
            LI.SubItems(15) = ADODC_UTAMA.Recordset.Fields(18).Value
            LI.SubItems(16) = ADODC_UTAMA.Recordset.Fields(19).Value
            LI.SubItems(17) = ADODC_UTAMA.Recordset.Fields(20).Value
            LI.SubItems(18) = ADODC_UTAMA.Recordset.Fields(21).Value
            LI.SubItems(19) = ADODC_UTAMA.Recordset.Fields(22).Value
            ADODC_UTAMA.Recordset.MoveNext
        Loop
        ADODC_UTAMA.Refresh
        Me.Caption = "Simple Accounts Manager - " & AdodcDataLogin.Recordset.Fields(2).Value & " (" & cmIdentitasPribadi.Caption & ")"
        AktifkanTombolNavigasi
        With StatusBawah.Panels
            .Item(1).Text = cmIdentitasPribadi.Caption
            .Item(3).Text = ADODC_UTAMA.Recordset.RecordCount
            .Item(4).Text = ADODC_UTAMA.Recordset.RecordCount * ADODC_UTAMA.Recordset.Fields.Count
        End With
        AturStatusBawah
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            menuTampilkanPassword.Caption = "Tampilkan Password"
        Else
            menuTampilkanPassword.Caption = "Show Passwords"
        End If
            menuTampilkanPassword.Enabled = False
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
                    .Recordset.Fields(6).Value = "Data identitas pribadi ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The personal identity files viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
End Sub

Public Sub cmJejaringSosial_Click()
    menuTabelFilter.Enabled = False
    cmIdentitasPribadi.FontBold = False
    cmBukuAlamat.FontBold = False
    cmUlangTahun.FontBold = False
    cmAgenda.FontBold = False
    cmRegistrasiSoftware.FontBold = False
    cmJejaringSosial.FontBold = True
    cmElectronicMail.FontBold = False
    cmForumInternet.FontBold = False
    cmFTP.FontBold = False
    cmBlogging.FontBold = False
    
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Nama Jejaring", 2000
            .ColumnHeaders.Add , , "Nama Pengguna", 2000, vbCenter
            .ColumnHeaders.Add , , "Alamat E-Mail", 2000, vbCenter
            .ColumnHeaders.Add , , "URL", 2000, vbCenter
            .ColumnHeaders.Add , , "Pemilik Akun", 2000, vbCenter
            .ColumnHeaders.Add , , "Tanggal", 2000, vbCenter
            .ColumnHeaders.Add , , "Keterangan", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    Else
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Social Name", 2000
            .ColumnHeaders.Add , , "User Name", 2000, vbCenter
            .ColumnHeaders.Add , , "E-Mail Address", 2000, vbCenter
            .ColumnHeaders.Add , , "URL", 2000, vbCenter
            .ColumnHeaders.Add , , "Account Owner", 2000, vbCenter
            .ColumnHeaders.Add , , "Date", 2000, vbCenter
            .ColumnHeaders.Add , , "Description", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    End If
        If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
            CN_FormUtama.CursorLocation = adUseClient
            CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & Me.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
        With ADODC_UTAMA
            .ConnectionString = CN_FormUtama.ConnectionString
            .RecordSource = "Select Nama_Jejaring,Nama_Pengguna,Alamat_Email,URL,Pemilik_Akun,Tanggal,Keterangan From tbJejaringSosial;"
            .Refresh
        End With
        LV.ListItems.Clear
        Do Until ADODC_UTAMA.Recordset.EOF
        Set LI = LV.ListItems.Add(, , ADODC_UTAMA.Recordset.Fields(0).Value)
            LI.SubItems(1) = ADODC_UTAMA.Recordset.Fields(1).Value
            LI.SubItems(2) = ADODC_UTAMA.Recordset.Fields(2).Value
            LI.SubItems(3) = ADODC_UTAMA.Recordset.Fields(3).Value
            LI.SubItems(4) = ADODC_UTAMA.Recordset.Fields(4).Value
            LI.SubItems(5) = ADODC_UTAMA.Recordset.Fields(5).Value
            LI.SubItems(6) = ADODC_UTAMA.Recordset.Fields(6).Value
            ADODC_UTAMA.Recordset.MoveNext
        Loop
        ADODC_UTAMA.Refresh
        Me.Caption = "Simple Accounts Manager - " & AdodcDataLogin.Recordset.Fields(2).Value & " (" & cmJejaringSosial.Caption & ")"
        AktifkanTombolNavigasi
        With StatusBawah.Panels
            .Item(1).Text = cmJejaringSosial.Caption
            .Item(3).Text = ADODC_UTAMA.Recordset.RecordCount
            .Item(4).Text = ADODC_UTAMA.Recordset.RecordCount * ADODC_UTAMA.Recordset.Fields.Count
        End With
        AturStatusBawah
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            menuTampilkanPassword.Caption = "Tampilkan Password"
        Else
            menuTampilkanPassword.Caption = "Show Passwords"
        End If
            menuTampilkanPassword.Enabled = True
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
                    .Recordset.Fields(6).Value = "Data Jejaring Sosial ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The Social Networking files viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
End Sub

Sub UntukFormManage()
    With FormManage
        .Show vbModal, Me
    End With
End Sub


Public Sub cmManage_Click()
If ADODC_UTAMA.Recordset.RecordCount = 0 Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Maaf, data masih kosong.", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Sorry, data is empty", vbExclamation + vbOKOnly, ""
    End If
Else
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        If cmIdentitasPribadi.FontBold = True Or cmBukuAlamat.FontBold = True Or cmUlangTahun.FontBold = True Or cmAgenda.FontBold = True Or cmRegistrasiSoftware.FontBold = True Then
            UntukFormManage
        Else
            If Me.menuTampilkanPassword.Caption = "Tampilkan Password" Then
                Pesan = MsgBox("Untuk me-manage data, password harus ditampilkan." & vbCrLf & _
                                "Lanjutkan?", vbQuestion + vbYesNo, "Manage Data?")
                If Pesan = vbYes Then
                    UntukFormManage
                End If
            Else
                UntukFormManage
            End If
        End If
    Else
        If cmIdentitasPribadi.FontBold = True Or cmBukuAlamat.FontBold = True Or cmUlangTahun.FontBold = True Or cmAgenda.FontBold = True Or cmRegistrasiSoftware.FontBold = True Then
            UntukFormManage
        Else
            If Me.menuTampilkanPassword.Caption = "Show Passwords" Then
                Pesan = MsgBox("For to manage data, passwords should be displayed." & vbCrLf & _
                                "Continued?", vbQuestion + vbYesNo, "Manage Data?")
                If Pesan = vbYes Then
                    UntukFormManage
                End If
            Else
                UntukFormManage
            End If
        End If
    End If
End If
End Sub

Public Sub cmRefresh_Click()
    AturKontrol
    AturStatusBawah
    AktifkanTombolNavigasi
    If cmIdentitasPribadi.Font.Bold = True Then cmIdentitasPribadi_Click
    If cmBukuAlamat.Font.Bold = True Then cmBukuAlamat_Click
    If cmUlangTahun.Font.Bold = True Then cmUlangTahun_Click
    If cmAgenda.Font.Bold = True Then cmAgenda_Click
    If cmRegistrasiSoftware.Font.Bold = True Then cmRegistrasiSoftware_Click
    If cmJejaringSosial.Font.Bold = True Then cmJejaringSosial_Click
    If cmElectronicMail.Font.Bold = True Then cmElectronicMail_Click
    If cmForumInternet.Font.Bold = True Then cmForumInternet_Click
    If cmFTP.Font.Bold = True Then cmFTP_Click
    If cmBlogging.Font.Bold = True Then cmBlogging_Click
End Sub

Public Sub cmRegistrasiSoftware_Click()
    menuTabelFilter.Enabled = True
    cmIdentitasPribadi.FontBold = False
    cmBukuAlamat.FontBold = False
    cmUlangTahun.FontBold = False
    cmAgenda.FontBold = False
    cmRegistrasiSoftware.FontBold = True
    cmJejaringSosial.FontBold = False
    cmElectronicMail.FontBold = False
    cmForumInternet.FontBold = False
    cmFTP.FontBold = False
    cmBlogging.FontBold = False

    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Nama Software", 2000
            .ColumnHeaders.Add , , "Kategori", 2000, vbCenter
            .ColumnHeaders.Add , , "Developer/Programmer", 3000, vbCenter
            .ColumnHeaders.Add , , "Nama User/Grup/Office", 3000, vbCenter
            .ColumnHeaders.Add , , "Serial/Key/Code", 3000, vbCenter
            .ColumnHeaders.Add , , "Jenis Lisensi", 2000, vbCenter
            .ColumnHeaders.Add , , "Keterangan", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    Else
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Software Name", 2000
            .ColumnHeaders.Add , , "Category", 2000, vbCenter
            .ColumnHeaders.Add , , "Developer/Programmer", 3000, vbCenter
            .ColumnHeaders.Add , , "User/Group/Office Name", 3000, vbCenter
            .ColumnHeaders.Add , , "Serial/Key/Code", 3000, vbCenter
            .ColumnHeaders.Add , , "License", 2000, vbCenter
            .ColumnHeaders.Add , , "Description", 2000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    End If
        If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
            CN_FormUtama.CursorLocation = adUseClient
            CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & Me.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
        With ADODC_UTAMA
            .ConnectionString = CN_FormUtama.ConnectionString
            .RecordSource = "Select * From tbRegistrasiSoftware order by Nama_Software asc;"
            .Refresh
        End With
        LV.ListItems.Clear
        Do Until ADODC_UTAMA.Recordset.EOF
        Set LI = LV.ListItems.Add(, , ADODC_UTAMA.Recordset.Fields(0).Value)
            LI.SubItems(1) = ADODC_UTAMA.Recordset.Fields(1).Value
            LI.SubItems(2) = ADODC_UTAMA.Recordset.Fields(2).Value
            LI.SubItems(3) = ADODC_UTAMA.Recordset.Fields(3).Value
            LI.SubItems(4) = ADODC_UTAMA.Recordset.Fields(4).Value
            LI.SubItems(5) = ADODC_UTAMA.Recordset.Fields(5).Value
            LI.SubItems(6) = ADODC_UTAMA.Recordset.Fields(6).Value
            ADODC_UTAMA.Recordset.MoveNext
        Loop
        ADODC_UTAMA.Refresh
        Me.Caption = "Simple Accounts Manager - " & AdodcDataLogin.Recordset.Fields(2).Value & " (" & cmRegistrasiSoftware.Caption & ")"
        AktifkanTombolNavigasi
        With StatusBawah.Panels
            .Item(1).Text = cmRegistrasiSoftware.Caption
            .Item(3).Text = ADODC_UTAMA.Recordset.RecordCount
            .Item(4).Text = ADODC_UTAMA.Recordset.RecordCount * ADODC_UTAMA.Recordset.Fields.Count
        End With
        AturStatusBawah
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            menuTampilkanPassword.Caption = "Tampilkan Password"
        Else
            menuTampilkanPassword.Caption = "Show Passwords"
        End If
            menuTampilkanPassword.Enabled = False
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
                    .Recordset.Fields(6).Value = "Data Registrasi Software ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The Software Registration files viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
End Sub


Public Sub cmUlangTahun_Click()
    menuTabelFilter.Enabled = True
    cmIdentitasPribadi.FontBold = False
    cmBukuAlamat.FontBold = False
    cmUlangTahun.FontBold = True
    cmAgenda.FontBold = False
    cmRegistrasiSoftware.FontBold = False
    cmJejaringSosial.FontBold = False
    cmElectronicMail.FontBold = False
    cmForumInternet.FontBold = False
    cmFTP.FontBold = False
    cmBlogging.FontBold = False

    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Nama", 3500
            .ColumnHeaders.Add , , "Tempat/Tanggal Lahir", 3000, vbCenter
            .ColumnHeaders.Add , , "Keterangan", 4000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    Else
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Name", 3500
            .ColumnHeaders.Add , , "Place/Born Day", 3000, vbCenter
            .ColumnHeaders.Add , , "Description", 4000, vbCenter
            .View = lvwReport
            .Sorted = True
        End With
    End If
        If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
            CN_FormUtama.CursorLocation = adUseClient
            CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & Me.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
        With ADODC_UTAMA
            .ConnectionString = CN_FormUtama.ConnectionString
            .RecordSource = "Select * From tbUlangTahun order by Nama asc;"
            .Refresh
        End With
        LV.ListItems.Clear
        Do Until ADODC_UTAMA.Recordset.EOF
        Set LI = LV.ListItems.Add(, , ADODC_UTAMA.Recordset.Fields(0).Value)
            LI.SubItems(1) = ADODC_UTAMA.Recordset.Fields(1).Value
            LI.SubItems(2) = ADODC_UTAMA.Recordset.Fields(2).Value
            ADODC_UTAMA.Recordset.MoveNext
        Loop
        ADODC_UTAMA.Refresh
        Me.Caption = "Simple Accounts Manager - " & AdodcDataLogin.Recordset.Fields(2).Value & " (" & cmUlangTahun.Caption & ")"
        AktifkanTombolNavigasi
        With StatusBawah.Panels
            .Item(1).Text = cmUlangTahun.Caption
            .Item(3).Text = ADODC_UTAMA.Recordset.RecordCount
            .Item(4).Text = ADODC_UTAMA.Recordset.RecordCount * ADODC_UTAMA.Recordset.Fields.Count
        End With
        AturStatusBawah
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            menuTampilkanPassword.Caption = "Tampilkan Password"
        Else
            menuTampilkanPassword.Caption = "Show Passwords"
        End If
            menuTampilkanPassword.Enabled = False
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
                    .Recordset.Fields(6).Value = "Data Ulang Tahun ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "The Birthday files viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
End Sub
Sub PENGATURAN_WARNA()
    'PENGATURAN WARNA UNTUK FORM_UTAMA
    For Each Objek In FORM_UTAMA
        Select Case FormPengaturan.cmbWarnaTampilan.ListIndex
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
            Select Case FormPengaturan.cmbTemaTampilan.ListIndex
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
End Sub

Public Sub CheckSoftware(X As Form)
On Error GoTo Pesan
    Dim SaveTitle$
    If App.PrevInstance Then
        SaveTitle$ = App.Title
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Program ini sedang dijalankan!", _
               vbCritical, "Sedang Dijalankan"
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Program is running!", _
               vbCritical, "Sedang Dijalankan"
        End If
        App.Title = ""
        X.Caption = ""
        AppActivate SaveTitle$
        SendKeys "%{ENTER}", True
        End
    End If
    Exit Sub
Pesan:
    End
    Exit Sub
End Sub

Private Sub Form_Load()
    AturKontrol
    PENGATURAN_FORM
    PENGATURAN_BAHASA
    PENGATURAN_WARNA
    Call CheckSoftware(FORM_UTAMA)
End Sub

Sub SambungkanADODC_UTAMA()
On Error Resume Next
    If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
        CN_FormUtama.CursorLocation = adUseClient
        CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & FormLogin.textPengguna.Text & "\data.rdb;Persist Security Info=False"
End Sub
Sub SambungkanADODC_DataLogin()
On Error Resume Next
    If CN_FormUtamaLogin.State = adStateOpen Then CN_FormUtamaLogin.Close
        CN_FormUtamaLogin.CursorLocation = adUseClient
        CN_FormUtamaLogin.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & FormLogin.textPengguna.Text & "\data.rdb;Persist Security Info=False"
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu MKK
End Sub


Private Sub menuaPengaturan_Click()
    FormPengaturan.Show vbModal, Me
End Sub

Private Sub menuAR_Click()
    Dim Jalankan
    Kalimat = App.Path & "\[program_tambahan]\Adress Register v1.0\Adress Register (PhoneBook) v1.0.exe"
    If Dir$(Kalimat) <> "" Then
        Jalankan = Shell(Kalimat, vbNormalFocus)
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf file 'Adress Register (PhoneBook) v1.0.exe' tidak ditemukan dalam sistem Anda!" & vbCrLf & _
                    "Silahkan hubungi administrator Anda untuk menyelesaikan masalah ini!", vbCritical + vbOKOnly, "ErrorMainSystem - File Tidak Ditemukan."
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry file 'Adress Register (PhoneBook) v1.0.exe' is not found in your system!" & vbCrLf & _
                    "Please contact your administrator to resolve this problem!", vbCritical + vbOKOnly, "ErrorMainSystem - File Not Found."
        End If
    End If
End Sub

Private Sub menuASCG_Click()
    Dim Jalankan
    Kalimat = App.Path & "\[program_tambahan]\Access SQL Code Generator v2.0\Access SQL Code Generator v2.0.exe"
    If Dir$(Kalimat) <> "" Then
        Jalankan = Shell(Kalimat, vbNormalFocus)
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf file 'Access SQL Code Generator v2.0.exe' tidak ditemukan dalam sistem Anda!" & vbCrLf & _
                    "Silahkan hubungi administrator Anda untuk menyelesaikan masalah ini!", vbCritical + vbOKOnly, "ErrorMainSystem - File Tidak Ditemukan."
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry file 'Access SQL Code Generator v2.0.exe' is not found in your system!" & vbCrLf & _
                    "Please contact your administrator to resolve this problem!", vbCritical + vbOKOnly, "ErrorMainSystem - File Not Found."
        End If
    End If
End Sub

Private Sub menuBaruA_Click()
    With Me
        .cmAgenda_Click
        .cmAgenda.SetFocus
        .cmBaru_Click
    End With
End Sub

Private Sub menuBaruBA_Click()
    With Me
        .cmBukuAlamat_Click
        .cmBukuAlamat.SetFocus
        .cmBaru_Click
    End With
End Sub

Private Sub menuBaruBlogging_Click()
    With Me
        .cmBlogging_Click
        .cmBlogging.SetFocus
        .cmBaru_Click
    End With
End Sub

Private Sub menubaruEM_Click()
    With Me
        .cmElectronicMail_Click
        .cmElectronicMail.SetFocus
        .cmBaru_Click
    End With
End Sub

Private Sub menuBaruFTP_Click()
    With Me
        .cmFTP_Click
        .cmFTP.SetFocus
        .cmBaru_Click
    End With
End Sub

Private Sub menuBaruIP_Click()
    With Me
        .cmIdentitasPribadi_Click
        .cmIdentitasPribadi.SetFocus
        .cmBaru_Click
    End With
End Sub

Private Sub menuBaruJS_Click()
    With Me
        .cmJejaringSosial_Click
        .cmJejaringSosial.SetFocus
        .cmBaru_Click
    End With
End Sub

Private Sub menuBaruRS_Click()
    With Me
        .cmRegistrasiSoftware_Click
        .cmRegistrasiSoftware.SetFocus
        .cmBaru_Click
    End With
End Sub

Private Sub menuBaruUT_Click()
    With Me
        .cmUlangTahun_Click
        .cmUlangTahun.SetFocus
        .cmBaru_Click
    End With
End Sub

Private Sub menuBS_Click()
    If menuBS.Checked = False Then
        With Me
            .StatusBawah.Visible = True
            .Height = 8595
            .ProgressLogOut.Top = 7570
        End With
        menuBS.Checked = True
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
                    .Recordset.Fields(6).Value = "Status bar ditampilkan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "Status bar is viewed"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
    ElseIf menuBS.Checked = True Then
        With Me
            .StatusBawah.Visible = False
            .Height = 8265
            .ProgressLogOut.Top = 7200
        End With
        menuBS.Checked = False
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
                    .Recordset.Fields(6).Value = "Status bar disembunyikan"
                ElseIf FormSettingRiwayatLebihLanjut.CmbBahasaPencatatan.ListIndex = 1 Then
                    .Recordset.Fields(6).Value = "Status bar is hide"
                End If
                .Recordset.Fields(7).Value = GetComputerName
                .Recordset.Update
                .Refresh
            End With
        End If
    End If
    SaveSetting "rssamv1.0", "FormUtama", menuBS.Name, menuBS.Checked
End Sub

Private Sub menuES_Click()
    Dim Jalankan
    Kalimat = App.Path & "\[program_tambahan]\Encrypt String v2.0\Encrypt String v2.0.exe"
    If Dir$(Kalimat) <> "" Then
        Jalankan = Shell(Kalimat, vbNormalFocus)
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf file 'Encrypt String v2.0.exe' tidak ditemukan dalam sistem Anda!" & vbCrLf & _
                    "Silahkan hubungi administrator Anda untuk menyelesaikan masalah ini!", vbCritical + vbOKOnly, "ErrorMainSystem - File Tidak Ditemukan."
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry file 'Encrypt String v2.0.exe' is not found in your system!" & vbCrLf & _
                    "Please contact your administrator to resolve this problem!", vbCritical + vbOKOnly, "ErrorMainSystem - File Not Found."
        End If
    End If
End Sub

Private Sub menuFI_Click()
    With Me
        .cmForumInternet_Click
        .cmForumInternet.SetFocus
        .cmBaru_Click
    End With
End Sub

Private Sub menuHubungiDeveloper_Click()
    Kalimat = "rikymetal10@gmail.com"
    EMAIL = ShellExecute(0, vbNullString, "mailto:" & Kalimat, "", "", vbNormalFocus)
End Sub

Private Sub menuKH_Click()
    Kalimat = "http://rikymetalist.blogspot.com/p/software-ku.html"
    SITUS = ShellExecute(0, vbNullString, Kalimat, "", "", vbNormalFocus)
End Sub

Private Sub menuKV_Click()
    Dim Jalankan
    Kalimat = "C:\Windows\system32\osk.exe"
    If Dir$(Kalimat) <> "" Then
        Jalankan = Shell(Kalimat, vbNormalFocus)
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf file 'osk.exe' yang berfungsi sebagai keyboard virtual" & vbCrLf & _
                    "tidak ditemukan dalam sistem operasi Anda!" & vbCrLf & _
                    "Silahkan hubungi administrator Anda untuk menyelesaikan masalah ini!", vbCritical + vbOKOnly, "ErrorMainSystem - File Tidak Ditemukan."
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry file 'osk.exe' which function as virtual keyboard" & vbCrLf & _
                    "is not found in your operating system!" & vbCrLf & _
                    "Please contact your administrator to resolve this problem!", vbCritical + vbOKOnly, "ErrorMainSystem - File Not Found."
        End If
    End If
End Sub

Private Sub menuLogOut_Click()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Pesan = MsgBox("Anda yakin ingin Log Out dari Akun Anda? " & vbCrLf & _
                        "(" & StatusBawah.Panels.Item(2).Text & ")", vbQuestion + vbYesNo, "Log Out?")
    Else
        Pesan = MsgBox("Are You sure to Log Out from your account? " & vbCrLf & _
                        "(" & StatusBawah.Panels.Item(2).Text & ")", vbQuestion + vbYesNo, "Log Out?")
    End If
    If Pesan = vbYes Then
        FormLogOut.Show
        ProgressLogOut.Value = 0
        TimerProgress.Enabled = True
        TimerProgress_Timer
        ProgressLogOut.Visible = True
        Me.Enabled = False
    End If
End Sub


Private Sub menuNotepad_Click()
    Dim Jalankan
    Kalimat = "C:\Windows\system32\Notepad.exe"
    If Dir$(Kalimat) <> "" Then
        Jalankan = Shell(Kalimat, vbNormalFocus)
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf file 'Notepad.exe' tidak ditemukan dalam sistem Anda!" & vbCrLf & _
                    "Silahkan hubungi administrator Anda untuk menyelesaikan masalah ini!", vbCritical + vbOKOnly, "ErrorMainSystem - File Tidak Ditemukan."
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry file 'Notepad.exe' is not found in your system!" & vbCrLf & _
                    "Please contact your administrator to resolve this problem!", vbCritical + vbOKOnly, "ErrorMainSystem - File Not Found."
        End If
    End If
End Sub

Private Sub menuPusatBantuan_Click()
    Kalimat = App.Path & "\bantuan\chm\Simple Account Manajer.chm"
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

Private Sub menuPVT_Click()
    Kalimat = "http://rikymetalist.blogspot.com/p/software-ku.html"
    SITUS = ShellExecute(0, vbNullString, Kalimat, "", "", vbNormalFocus)
End Sub

Private Sub menuRA_Click()
    With FormRiwayatAktivitas
        .Show
        .SetFocus
    End With
End Sub


Private Sub menuSN_Click()
    Dim Jalankan
    Kalimat = "C:\Windows\system32\StikyNot.exe"
    If Dir$(Kalimat) <> "" Then
        Jalankan = Shell(Kalimat, vbNormalFocus)
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf file 'StikyNot.exe' tidak ditemukan dalam sistem Anda!" & vbCrLf & _
                    "Silahkan hubungi administrator Anda untuk menyelesaikan masalah ini!", vbCritical + vbOKOnly, "ErrorMainSystem - File Tidak Ditemukan."
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry file 'StikyNot.exe' is not found in your system!" & vbCrLf & _
                    "Please contact your administrator to resolve this problem!", vbCritical + vbOKOnly, "ErrorMainSystem - File Not Found."
        End If
    End If
End Sub

Private Sub menuTabelFilter_Click()
    If ADODC_UTAMA.Recordset.RecordCount = 0 Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Tidak ada data yang akan di filter!", vbExclamation + vbOKOnly, ""
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "No data to in the filtered!", vbExclamation + vbOKOnly, ""
        End If
    Else
        With FormTabelFilter
            .Show , Me
        End With
    End If
End Sub

Private Sub menuTampilkanPassword_Click()
    If menuTampilkanPassword.Caption = "Tampilkan Password" Or menuTampilkanPassword.Caption = "Show Passwords" Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Pesan = MsgBox("Anda yakin ingin menampilkan password?", vbQuestion + vbYesNo, "Tampilkan Password?")
        Else
            Pesan = MsgBox("Are You sure to show passwords", vbQuestion + vbYesNo, "Show Passwords?")
        End If
            If Pesan = vbYes Then
                If cmJejaringSosial.FontBold = True Then
                    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                        menuTampilkanPassword.Caption = "Sembunyikan Password"
                        With LV
                            .ColumnHeaders.Clear
                            .ColumnHeaders.Add , , "Nama Jejaring", 2000
                            .ColumnHeaders.Add , , "Nama Pengguna", 2000, vbCenter
                            .ColumnHeaders.Add , , "Alamat E-Mail", 2000, vbCenter
                            .ColumnHeaders.Add , , "Password", 2000, vbCenter
                            .ColumnHeaders.Add , , "URL", 2000, vbCenter
                            .ColumnHeaders.Add , , "Pemilik Akun", 2000, vbCenter
                            .ColumnHeaders.Add , , "Tanggal", 2000, vbCenter
                            .ColumnHeaders.Add , , "Keterangan", 2000, vbCenter
                            .View = lvwReport
                            .Sorted = True
                        End With
                    Else
                        menuTampilkanPassword.Caption = "Hide Passwords"
                        With LV
                            .ColumnHeaders.Clear
                            .ColumnHeaders.Add , , "Social Name", 2000
                            .ColumnHeaders.Add , , "User Name", 2000, vbCenter
                            .ColumnHeaders.Add , , "E-Mail Address", 2000, vbCenter
                            .ColumnHeaders.Add , , "Passwords", 2000, vbCenter
                            .ColumnHeaders.Add , , "URL", 2000, vbCenter
                            .ColumnHeaders.Add , , "Account Owner", 2000, vbCenter
                            .ColumnHeaders.Add , , "Date", 2000, vbCenter
                            .ColumnHeaders.Add , , "Description", 2000, vbCenter
                            .View = lvwReport
                            .Sorted = True
                        End With
                    End If
                    If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
                        CN_FormUtama.CursorLocation = adUseClient
                        CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & Me.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
                    With ADODC_UTAMA
                        .ConnectionString = CN_FormUtama.ConnectionString
                        .RecordSource = "Select * From tbJejaringSosial order by Nama_Jejaring asc;"
                        .Refresh
                    End With
                    LV.ListItems.Clear
                    Do Until ADODC_UTAMA.Recordset.EOF
                    Set LI = LV.ListItems.Add(, , ADODC_UTAMA.Recordset.Fields(0).Value)
                        LI.SubItems(1) = ADODC_UTAMA.Recordset.Fields(1).Value
                        LI.SubItems(2) = ADODC_UTAMA.Recordset.Fields(2).Value
                        LI.SubItems(3) = ADODC_UTAMA.Recordset.Fields(3).Value
                        LI.SubItems(4) = ADODC_UTAMA.Recordset.Fields(4).Value
                        LI.SubItems(5) = ADODC_UTAMA.Recordset.Fields(5).Value
                        LI.SubItems(6) = ADODC_UTAMA.Recordset.Fields(6).Value
                        LI.SubItems(7) = ADODC_UTAMA.Recordset.Fields(7).Value
                        ADODC_UTAMA.Recordset.MoveNext
                    Loop
                    ADODC_UTAMA.Refresh
                    StatusBawah.Panels.Item(1).Text = cmJejaringSosial.Caption
                ElseIf cmElectronicMail.FontBold = True Then
                    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    menuTampilkanPassword.Caption = "Sembunyikan Password"
                       With LV
                            .ColumnHeaders.Clear
                            .ColumnHeaders.Add , , "Nama Server", 2000
                            .ColumnHeaders.Add , , "Nama Pengguna", 2000, vbCenter
                            .ColumnHeaders.Add , , "Alamat E-Mail", 2000, vbCenter
                            .ColumnHeaders.Add , , "Password", 2000, vbCenter
                            .ColumnHeaders.Add , , "Pertanyaan Rahasia", 2000, vbCenter
                            .ColumnHeaders.Add , , "Jawaban Pertanyaan", 2000, vbCenter
                            .ColumnHeaders.Add , , "URL", 2000, vbCenter
                            .ColumnHeaders.Add , , "Pemilik Akun", 2000, vbCenter
                            .ColumnHeaders.Add , , "Tanggal", 2000, vbCenter
                            .ColumnHeaders.Add , , "Keterangan", 2000, vbCenter
                            .View = lvwReport
                            .Sorted = True
                        End With
                    Else
                        menuTampilkanPassword.Caption = "Hide Passwords"
                        With LV
                            .ColumnHeaders.Clear
                            .ColumnHeaders.Add , , "Server Name", 2000
                            .ColumnHeaders.Add , , "User Name", 2000, vbCenter
                            .ColumnHeaders.Add , , "Mail Address", 2000, vbCenter
                            .ColumnHeaders.Add , , "Passwords", 2000, vbCenter
                            .ColumnHeaders.Add , , "Security Question", 2000, vbCenter
                            .ColumnHeaders.Add , , "Security Answer", 2000, vbCenter
                            .ColumnHeaders.Add , , "URL", 2000, vbCenter
                            .ColumnHeaders.Add , , "Account Owner", 2000, vbCenter
                            .ColumnHeaders.Add , , "Date", 2000, vbCenter
                            .ColumnHeaders.Add , , "Description", 2000, vbCenter
                            .View = lvwReport
                            .Sorted = True
                        End With
                    End If
                        If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
                            CN_FormUtama.CursorLocation = adUseClient
                            CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & Me.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
                        With ADODC_UTAMA
                            .ConnectionString = CN_FormUtama.ConnectionString
                            .RecordSource = "Select * From tbElectronicMail order by Nama_Server asc;"
                            .Refresh
                        End With
                        LV.ListItems.Clear
                        Do Until ADODC_UTAMA.Recordset.EOF
                        Set LI = LV.ListItems.Add(, , ADODC_UTAMA.Recordset.Fields(0).Value)
                            LI.SubItems(1) = ADODC_UTAMA.Recordset.Fields(1).Value
                            LI.SubItems(2) = ADODC_UTAMA.Recordset.Fields(2).Value
                            LI.SubItems(3) = ADODC_UTAMA.Recordset.Fields(3).Value
                            LI.SubItems(4) = ADODC_UTAMA.Recordset.Fields(4).Value
                            LI.SubItems(5) = ADODC_UTAMA.Recordset.Fields(5).Value
                            LI.SubItems(6) = ADODC_UTAMA.Recordset.Fields(6).Value
                            LI.SubItems(7) = ADODC_UTAMA.Recordset.Fields(7).Value
                            LI.SubItems(8) = ADODC_UTAMA.Recordset.Fields(8).Value
                            LI.SubItems(9) = ADODC_UTAMA.Recordset.Fields(9).Value
                            ADODC_UTAMA.Recordset.MoveNext
                        Loop
                        ADODC_UTAMA.Refresh
                        StatusBawah.Panels.Item(1).Text = cmElectronicMail.Caption
                ElseIf cmForumInternet.FontBold = True Then
                    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                        menuTampilkanPassword.Caption = "Sembunyikan Password"
                        With LV
                            .ColumnHeaders.Clear
                            .ColumnHeaders.Add , , "Nama Forum", 2000
                            .ColumnHeaders.Add , , "Nama Pengguna", 2000, vbCenter
                            .ColumnHeaders.Add , , "Alamat E-Mail", 2000, vbCenter
                            .ColumnHeaders.Add , , "Passwords", 2000, vbCenter
                            .ColumnHeaders.Add , , "Jabatan", 2000, vbCenter
                            .ColumnHeaders.Add , , "NickName", 2000, vbCenter
                            .ColumnHeaders.Add , , "URL", 2000, vbCenter
                            .ColumnHeaders.Add , , "Tanggal", 2000, vbCenter
                            .ColumnHeaders.Add , , "Keterangan", 2000, vbCenter
                            .View = lvwReport
                            .Sorted = True
                        End With
                    Else
                        With LV
                        menuTampilkanPassword.Caption = "Hide Passwords"
                            .ColumnHeaders.Clear
                            .ColumnHeaders.Add , , "Forum Name", 2000
                            .ColumnHeaders.Add , , "User Name", 2000, vbCenter
                            .ColumnHeaders.Add , , "Mail Address", 2000, vbCenter
                            .ColumnHeaders.Add , , "Passwords", 2000, vbCenter
                            .ColumnHeaders.Add , , "Position", 2000, vbCenter
                            .ColumnHeaders.Add , , "NickName", 2000, vbCenter
                            .ColumnHeaders.Add , , "URL", 2000, vbCenter
                            .ColumnHeaders.Add , , "Date", 2000, vbCenter
                            .ColumnHeaders.Add , , "Description", 2000, vbCenter
                            .View = lvwReport
                            .Sorted = True
                        End With
                    End If
                        If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
                            CN_FormUtama.CursorLocation = adUseClient
                            CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & Me.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
                        With ADODC_UTAMA
                            .ConnectionString = CN_FormUtama.ConnectionString
                            .RecordSource = "Select * From tbForumInternet order by Nama_Forum asc;"
                            .Refresh
                        End With
                        LV.ListItems.Clear
                        Do Until ADODC_UTAMA.Recordset.EOF
                        Set LI = LV.ListItems.Add(, , ADODC_UTAMA.Recordset.Fields(0).Value)
                            LI.SubItems(1) = ADODC_UTAMA.Recordset.Fields(1).Value
                            LI.SubItems(2) = ADODC_UTAMA.Recordset.Fields(2).Value
                            LI.SubItems(3) = ADODC_UTAMA.Recordset.Fields(3).Value
                            LI.SubItems(4) = ADODC_UTAMA.Recordset.Fields(4).Value
                            LI.SubItems(5) = ADODC_UTAMA.Recordset.Fields(5).Value
                            LI.SubItems(6) = ADODC_UTAMA.Recordset.Fields(6).Value
                            LI.SubItems(7) = ADODC_UTAMA.Recordset.Fields(7).Value
                            LI.SubItems(8) = ADODC_UTAMA.Recordset.Fields(8).Value
                            ADODC_UTAMA.Recordset.MoveNext
                        Loop
                        ADODC_UTAMA.Refresh
                        StatusBawah.Panels.Item(1).Text = cmForumInternet.Caption
                ElseIf cmFTP.FontBold = True Then
                    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                        With LV
                            .ColumnHeaders.Clear
                            .ColumnHeaders.Add , , "Nama Host", 2000
                            .ColumnHeaders.Add , , "Nomor Port", 2000, vbCenter
                            .ColumnHeaders.Add , , "Nama Server", 2000, vbCenter
                            .ColumnHeaders.Add , , "Nama Pengguna", 2000, vbCenter
                            .ColumnHeaders.Add , , "E-Mail", 2000, vbCenter
                            .ColumnHeaders.Add , , "Passwords", 2000, vbCenter
                            .ColumnHeaders.Add , , "Tanggal", 2000, vbCenter
                            .ColumnHeaders.Add , , "Keterangan", 2000, vbCenter
                            .View = lvwReport
                            .Sorted = True
                        End With
                    Else
                        With LV
                            .ColumnHeaders.Clear
                            .ColumnHeaders.Add , , "Host Name", 2000
                            .ColumnHeaders.Add , , "Port Number", 2000, vbCenter
                            .ColumnHeaders.Add , , "Server Name", 2000, vbCenter
                            .ColumnHeaders.Add , , "User Name", 2000, vbCenter
                            .ColumnHeaders.Add , , "E-Mail", 2000, vbCenter
                            .ColumnHeaders.Add , , "Passwords", 2000, vbCenter
                            .ColumnHeaders.Add , , "Date", 2000, vbCenter
                            .ColumnHeaders.Add , , "Description", 2000, vbCenter
                            .View = lvwReport
                            .Sorted = True
                        End With
                    End If
                        If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
                            CN_FormUtama.CursorLocation = adUseClient
                            CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & Me.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
                        With ADODC_UTAMA
                            .ConnectionString = CN_FormUtama.ConnectionString
                            .RecordSource = "Select * From tbFTP order by Nama_Host asc;"
                            .Refresh
                        End With
                        LV.ListItems.Clear
                        Do Until ADODC_UTAMA.Recordset.EOF
                        Set LI = LV.ListItems.Add(, , ADODC_UTAMA.Recordset.Fields(0).Value)
                            LI.SubItems(1) = ADODC_UTAMA.Recordset.Fields(1).Value
                            LI.SubItems(2) = ADODC_UTAMA.Recordset.Fields(2).Value
                            LI.SubItems(3) = ADODC_UTAMA.Recordset.Fields(3).Value
                            LI.SubItems(4) = ADODC_UTAMA.Recordset.Fields(4).Value
                            LI.SubItems(5) = ADODC_UTAMA.Recordset.Fields(5).Value
                            LI.SubItems(6) = ADODC_UTAMA.Recordset.Fields(6).Value
                            LI.SubItems(7) = ADODC_UTAMA.Recordset.Fields(7).Value
                            ADODC_UTAMA.Recordset.MoveNext
                        Loop
                        ADODC_UTAMA.Refresh
                        StatusBawah.Panels.Item(1).Text = cmFTP.Caption
                ElseIf cmBlogging.FontBold = True Then
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    With LV
                        .ColumnHeaders.Clear
                        .ColumnHeaders.Add , , "Nama Penyedia Blog", 2500
                        .ColumnHeaders.Add , , "Nama Pengguna", 2000, vbCenter
                        .ColumnHeaders.Add , , "E-Mail", 2000, vbCenter
                        .ColumnHeaders.Add , , "Password", 2000, vbCenter
                        .ColumnHeaders.Add , , "URL", 2000, vbCenter
                        .ColumnHeaders.Add , , "Tanggal", 2000, vbCenter
                        .ColumnHeaders.Add , , "Keterangan", 2000, vbCenter
                        .View = lvwReport
                        .Sorted = True
                    End With
                Else
                    With LV
                        .ColumnHeaders.Clear
                        .ColumnHeaders.Add , , "Blog Providers Name", 2500
                        .ColumnHeaders.Add , , "User Name", 2000, vbCenter
                        .ColumnHeaders.Add , , "E-Mail", 2000, vbCenter
                        .ColumnHeaders.Add , , "Password", 2000, vbCenter
                        .ColumnHeaders.Add , , "URL", 2000, vbCenter
                        .ColumnHeaders.Add , , "Date", 2000, vbCenter
                        .ColumnHeaders.Add , , "Description", 2000, vbCenter
                        .View = lvwReport
                        .Sorted = True
                    End With
                End If
                    If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
                        CN_FormUtama.CursorLocation = adUseClient
                        CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & Me.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
                    With ADODC_UTAMA
                        .ConnectionString = CN_FormUtama.ConnectionString
                        .RecordSource = "Select * From tbBlogging order by Nama_Penyedia_Blog asc;"
                        .Refresh
                    End With
                    LV.ListItems.Clear
                    Do Until ADODC_UTAMA.Recordset.EOF
                    Set LI = LV.ListItems.Add(, , ADODC_UTAMA.Recordset.Fields(0).Value)
                        LI.SubItems(1) = ADODC_UTAMA.Recordset.Fields(1).Value
                        LI.SubItems(2) = ADODC_UTAMA.Recordset.Fields(2).Value
                        LI.SubItems(3) = ADODC_UTAMA.Recordset.Fields(3).Value
                        LI.SubItems(4) = ADODC_UTAMA.Recordset.Fields(4).Value
                        LI.SubItems(5) = ADODC_UTAMA.Recordset.Fields(5).Value
                        LI.SubItems(6) = ADODC_UTAMA.Recordset.Fields(6).Value
                        ADODC_UTAMA.Recordset.MoveNext
                    Loop
                    ADODC_UTAMA.Refresh
                    StatusBawah.Panels.Item(1).Text = cmBlogging.Caption
                End If
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    menuTampilkanPassword.Caption = "Sembunyikan Password"
                Else
                    menuTampilkanPassword.Caption = "Hide Passwords"
                End If
            End If
    ElseIf menuTampilkanPassword.Caption = "Sembunyikan Password" Or menuTampilkanPassword.Caption = "Hide Passwords" Then
            If cmJejaringSosial.FontBold = True Then
                cmJejaringSosial_Click
            ElseIf cmElectronicMail.FontBold = True Then
                cmElectronicMail_Click
            ElseIf cmForumInternet.FontBold = True Then
                cmForumInternet_Click
            ElseIf cmFTP.FontBold = True Then
                cmFTP_Click
            ElseIf cmBlogging.FontBold = True Then
                cmBlogging_Click
            End If
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            menuTampilkanPassword.Caption = "Tampilkan Password"
        Else
            menuTampilkanPassword.Caption = "Show Passwords"
        End If
    End If
    Me.Caption = "Simple Accounts Manager - " & AdodcDataLogin.Recordset.Fields(2).Value & " (" & cmJejaringSosial.Caption & ")"
    AktifkanTombolNavigasi
        With StatusBawah.Panels
            .Item(3).Text = ADODC_UTAMA.Recordset.RecordCount
            .Item(4).Text = ADODC_UTAMA.Recordset.RecordCount * ADODC_UTAMA.Recordset.Fields.Count
        End With
    AturStatusBawah
End Sub

Private Sub menuTSAM_Click()
    FormTentang.Show vbModal, Me
End Sub

Private Sub timerDefaultTampilkanData_Timer()
    Select Case FormPengaturan.CmbDefaultTampilkanData.ListIndex
        Case Is = 0
            Me.cmIdentitasPribadi_Click
        Case Is = 1
            Me.cmBukuAlamat_Click
        Case Is = 2
            Me.cmUlangTahun_Click
        Case Is = 3
            Me.cmAgenda_Click
        Case Is = 4
            Me.cmRegistrasiSoftware_Click
        Case Is = 5
            Me.cmJejaringSosial_Click
        Case Is = 6
            Me.cmElectronicMail_Click
        Case Is = 7
            Me.cmForumInternet_Click
        Case Is = 8
            Me.cmFTP_Click
        Case Is = 9
            Me.cmBlogging_Click
    End Select
    timerDefaultTampilkanData.Enabled = False
End Sub

Private Sub TimerProgress_Timer()
ProgressLogOut.Value = ProgressLogOut.Value + 1
If FormPengaturan.cmbBahasa.ListIndex = 0 Then
    FormLogOut.Label2.Caption = " Sedang menyimpan data . . . (" & ProgressLogOut.Value & " %)"
Else
    FormLogOut.Label2.Caption = " Saving all data . . . (" & ProgressLogOut.Value & " %)"
End If
If ProgressLogOut.Value = 100 Then
    Unload FormLogOut
    Unload Me
    FormLogin.Show
End If
End Sub

Private Sub TimerWaktu_Timer()
If FormPengaturan.cmbBahasa.ListIndex = 0 Then
    Select Case Month(Date)
        Case Is = 1
            X = "Januari"
        Case Is = 2
            X = "Februari"
        Case Is = 3
            X = "Maret"
        Case Is = 4
            X = "April"
        Case Is = 5
            X = "Mei"
        Case Is = 6
            X = "Juni"
        Case Is = 7
            X = "Juli"
        Case Is = 8
            X = "Agustus"
        Case Is = 9
            X = "September"
        Case Is = 10
            X = "Oktober"
        Case Is = 11
            X = "November"
        Case Is = 12
            X = "Desember"
    End Select
Else
    Select Case Month(Date)
        Case Is = 1
            X = "January"
        Case Is = 2
            X = "February"
        Case Is = 3
            X = "March"
        Case Is = 4
            X = "April"
        Case Is = 5
            X = "May"
        Case Is = 6
            X = "June"
        Case Is = 7
            X = "July"
        Case Is = 8
            X = "August"
        Case Is = 9
            X = "September"
        Case Is = 10
            X = "October"
        Case Is = 11
            X = "November"
        Case Is = 12
            X = "December"
    End Select
End If
With StatusBawah.Panels
    .Item(5).Text = Time
    .Item(6).Text = Day(Date) & " - " & X & " - " & Year(Date)
End With
End Sub
