VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form Form_IDENTITAS_PRIBADI 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "----------"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_IDENTITAS_PRIBADI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FFFFFF&
      Height          =   7575
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   6855
      Begin VB.ComboBox cmbJenisKelamin 
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1320
         Width           =   2895
      End
      Begin VB.ComboBox cmbGolonganDarah 
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2040
         Width           =   2370
      End
      Begin VB.ComboBox cmbAgama 
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1680
         Width           =   2370
      End
      Begin VB.ComboBox cmbStatusHubungan 
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   6360
         Width           =   2370
      End
      Begin VB.ComboBox cmbTanggal 
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
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cmbBulan 
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
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cmbTahun 
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
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   960
         Width           =   735
      End
      Begin MSAdodcLib.Adodc AdodcStatusHubungan 
         Height          =   330
         Left            =   3120
         Top             =   6360
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
      Begin MSAdodcLib.Adodc AdodcGolonganDarah 
         Height          =   330
         Left            =   3120
         Top             =   2040
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
         Left            =   3120
         Top             =   1680
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
      Begin XPControls.XPText textNamaLengkap 
         Height          =   330
         Left            =   2040
         TabIndex        =   25
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
         TabIndex        =   26
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
         TabIndex        =   27
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
         TabIndex        =   28
         Top             =   960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Salin"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         State           =   3
      End
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   3
         Left            =   4965
         TabIndex        =   29
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
         TabIndex        =   30
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
         TabIndex        =   31
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
         TabIndex        =   32
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
         TabIndex        =   33
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
         TabIndex        =   34
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
         TabIndex        =   35
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
         TabIndex        =   36
         Top             =   960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "&Reset"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         State           =   3
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   3
         Left            =   5610
         TabIndex        =   37
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
         TabIndex        =   38
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
         TabIndex        =   39
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
         TabIndex        =   40
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
         TabIndex        =   41
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
         TabIndex        =   42
         Top             =   3120
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
         TabIndex        =   43
         Top             =   3480
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
         TabIndex        =   44
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
      Begin XPControls.XPText textTempat 
         Height          =   330
         Left            =   2040
         TabIndex        =   45
         Top             =   960
         Width           =   735
         _ExtentX        =   1296
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
      Begin XPControls.XPText textPekerjaan 
         Height          =   330
         Left            =   2040
         TabIndex        =   46
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
      Begin XPControls.XPText textAlamatRumah 
         Height          =   330
         Left            =   2040
         TabIndex        =   47
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
         Left            =   6240
         TabIndex        =   48
         Top             =   960
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
      Begin XPControls.XPText textAlamatEmail 
         Height          =   330
         Left            =   2040
         TabIndex        =   49
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
      Begin XPControls.XPText textAlamatWebsite 
         Height          =   330
         Left            =   2040
         TabIndex        =   50
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
         TabIndex        =   51
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
         TabIndex        =   52
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
         TabIndex        =   53
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
         TabIndex        =   54
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
      Begin XPControls.XPText textNomorTelepon 
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
      Begin Dacara_dcButton.dcButton cmSalin 
         Height          =   330
         Index           =   11
         Left            =   4965
         TabIndex        =   57
         Top             =   4200
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
         Index           =   12
         Left            =   4965
         TabIndex        =   58
         Top             =   4560
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
         Index           =   13
         Left            =   4965
         TabIndex        =   59
         Top             =   4920
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
         Index           =   14
         Left            =   4965
         TabIndex        =   60
         Top             =   5280
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
         Index           =   15
         Left            =   4965
         TabIndex        =   61
         Top             =   5640
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
         Index           =   16
         Left            =   4965
         TabIndex        =   62
         Top             =   6000
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
         Index           =   17
         Left            =   4965
         TabIndex        =   63
         Top             =   6360
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
         TabIndex        =   64
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
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   330
         Index           =   11
         Left            =   5610
         TabIndex        =   65
         Top             =   4200
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
         Index           =   12
         Left            =   5610
         TabIndex        =   66
         Top             =   4560
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
         Index           =   13
         Left            =   5610
         TabIndex        =   67
         Top             =   4920
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
         Index           =   14
         Left            =   5610
         TabIndex        =   68
         Top             =   5280
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
         Index           =   15
         Left            =   5610
         TabIndex        =   69
         Top             =   5640
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
         Index           =   16
         Left            =   5610
         TabIndex        =   70
         Top             =   6000
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
         Index           =   17
         Left            =   5610
         TabIndex        =   71
         Top             =   6360
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
      Begin XPControls.XPText textKotaAsal 
         Height          =   330
         Left            =   2040
         TabIndex        =   72
         Top             =   4200
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
      Begin XPControls.XPText textKotaSekarang 
         Height          =   330
         Left            =   2040
         TabIndex        =   73
         Top             =   4560
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
      Begin XPControls.XPText textKodePos 
         Height          =   330
         Left            =   2040
         TabIndex        =   74
         Top             =   4920
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
      Begin XPControls.XPText textProvinsi 
         Height          =   330
         Left            =   2040
         TabIndex        =   75
         Top             =   5280
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
      Begin XPControls.XPText textKewargaNegaraan 
         Height          =   330
         Left            =   2040
         TabIndex        =   76
         Top             =   5640
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
      Begin XPControls.XPText textStatusPendidikan 
         Height          =   330
         Left            =   2040
         TabIndex        =   77
         Top             =   6000
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
      Begin XPControls.XPText textHobby 
         Height          =   330
         Left            =   2040
         TabIndex        =   78
         Top             =   6720
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
         TabIndex        =   79
         Top             =   7080
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
         Index           =   18
         Left            =   4965
         TabIndex        =   80
         Top             =   6720
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
         Index           =   19
         Left            =   4965
         TabIndex        =   81
         Top             =   7080
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
         Index           =   18
         Left            =   5610
         TabIndex        =   82
         Top             =   6720
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
         Index           =   19
         Left            =   5610
         TabIndex        =   83
         Top             =   7080
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
      Begin Dacara_dcButton.dcButton cmTambahDataGolonganDarah 
         Height          =   330
         Left            =   4440
         TabIndex        =   84
         Top             =   2040
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "+"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmTambahDataAgama 
         Height          =   330
         Left            =   4440
         TabIndex        =   85
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "+"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton cmTambahDataStatusHubungan 
         Height          =   330
         Left            =   4440
         TabIndex        =   86
         Top             =   6360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "+"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label9 
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
         TabIndex        =   126
         Top             =   3120
         Width           =   1995
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Website"
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
         TabIndex        =   125
         Top             =   3480
         Width           =   1995
      End
      Begin VB.Label label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   124
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   123
         Top             =   3480
         Width           =   45
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Lengkap"
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
         TabIndex        =   122
         Top             =   240
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
         TabIndex        =   121
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "T.T.L"
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
         TabIndex        =   120
         Top             =   960
         Width           =   1995
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
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
         TabIndex        =   119
         Top             =   1320
         Width           =   1995
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Agama"
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
         TabIndex        =   118
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Golongan Darah"
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
         TabIndex        =   117
         Top             =   2040
         Width           =   1995
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pekerjaan"
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
         TabIndex        =   116
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label Label8 
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
         TabIndex        =   115
         Top             =   2760
         Width           =   1995
      End
      Begin VB.Label label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   114
         Top             =   240
         Width           =   45
      End
      Begin VB.Label label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   113
         Top             =   600
         Width           =   45
      End
      Begin VB.Label label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   112
         Top             =   960
         Width           =   45
      End
      Begin VB.Label label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   111
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   110
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   109
         Top             =   2040
         Width           =   45
      End
      Begin VB.Label label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   108
         Top             =   2400
         Width           =   45
      End
      Begin VB.Label labl28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   107
         Top             =   2760
         Width           =   45
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   106
         Top             =   6360
         Width           =   45
      End
      Begin VB.Label qqq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   105
         Top             =   6000
         Width           =   45
      End
      Begin VB.Label eee 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   104
         Top             =   5640
         Width           =   45
      End
      Begin VB.Label dsoifhkdf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   103
         Top             =   5280
         Width           =   45
      End
      Begin VB.Label qwewf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   102
         Top             =   4920
         Width           =   45
      End
      Begin VB.Label bfd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   101
         Top             =   4560
         Width           =   45
      End
      Begin VB.Label ewrefds 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   100
         Top             =   4200
         Width           =   45
      End
      Begin VB.Label sdfsdfv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   99
         Top             =   3840
         Width           =   45
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status Hubungan"
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
         TabIndex        =   98
         Top             =   6360
         Width           =   1995
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status Pendidikan"
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
         TabIndex        =   97
         Top             =   6000
         Width           =   1995
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kewarganegaraan"
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
         TabIndex        =   96
         Top             =   5640
         Width           =   1995
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Provinsi"
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
         TabIndex        =   95
         Top             =   5280
         Width           =   1995
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pos"
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
         TabIndex        =   94
         Top             =   4920
         Width           =   1995
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kota Sekarang"
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
         TabIndex        =   93
         Top             =   4560
         Width           =   1995
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kota Asal"
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
         TabIndex        =   92
         Top             =   4200
         Width           =   1995
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Telepon"
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
         TabIndex        =   91
         Top             =   3840
         Width           =   1995
      End
      Begin VB.Label sdfdsfdsf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   90
         Top             =   7080
         Width           =   45
      End
      Begin VB.Label sdfsd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   89
         Top             =   6720
         Width           =   45
      End
      Begin VB.Label Label20 
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
         TabIndex        =   88
         Top             =   7080
         Width           =   1995
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hobby"
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
         TabIndex        =   87
         Top             =   6720
         Width           =   1995
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Height          =   1215
         Left            =   0
         Picture         =   "Form_IDENTITAS_PRIBADI.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   7095
      End
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   8400
      Top             =   9360
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
      TabIndex        =   2
      Top             =   1440
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
      TabIndex        =   3
      Top             =   1800
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
      TabIndex        =   4
      Top             =   3600
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
      TabIndex        =   5
      Top             =   3960
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
      TabIndex        =   6
      Top             =   4320
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
      TabIndex        =   7
      Top             =   4680
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu20 
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
      TabIndex        =   8
      Top             =   8280
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu19 
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
      TabIndex        =   9
      Top             =   7920
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu17 
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
      TabIndex        =   10
      Top             =   7200
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu16 
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
      TabIndex        =   11
      Top             =   6840
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu15 
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
      TabIndex        =   12
      Top             =   6480
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu14 
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
      TabIndex        =   13
      Top             =   6120
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu13 
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
      TabIndex        =   14
      Top             =   5760
      Width           =   2895
   End
   Begin VB.ComboBox cmbDataLalu12 
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
      TabIndex        =   15
      Top             =   5400
      Width           =   2895
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
      TabIndex        =   16
      Top             =   5040
      Width           =   2895
   End
   Begin Dacara_dcButton.dcButton cmSimpan 
      Height          =   375
      Left            =   120
      TabIndex        =   127
      Top             =   8880
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
      PicDown         =   "Form_IDENTITAS_PRIBADI.frx":196E4
      PicHot          =   "Form_IDENTITAS_PRIBADI.frx":19A36
      PicNormal       =   "Form_IDENTITAS_PRIBADI.frx":19D88
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmReset 
      Height          =   375
      Left            =   1440
      TabIndex        =   128
      Top             =   8880
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
      PicDown         =   "Form_IDENTITAS_PRIBADI.frx":1A0DA
      PicHot          =   "Form_IDENTITAS_PRIBADI.frx":1AC24
      PicNormal       =   "Form_IDENTITAS_PRIBADI.frx":1B76E
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmBatal 
      Height          =   375
      Left            =   5760
      TabIndex        =   129
      Top             =   8880
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
      PicDown         =   "Form_IDENTITAS_PRIBADI.frx":1C2B8
      PicHot          =   "Form_IDENTITAS_PRIBADI.frx":1C70A
      PicNormal       =   "Form_IDENTITAS_PRIBADI.frx":1CB5C
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmVerifikasi 
      Height          =   375
      Left            =   2760
      TabIndex        =   130
      Top             =   8880
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
      PicDown         =   "Form_IDENTITAS_PRIBADI.frx":1CFAE
      PicHot          =   "Form_IDENTITAS_PRIBADI.frx":1D400
      PicNormal       =   "Form_IDENTITAS_PRIBADI.frx":1D852
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmBantuan 
      Height          =   375
      Left            =   4080
      TabIndex        =   131
      Top             =   8880
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
      PicDown         =   "Form_IDENTITAS_PRIBADI.frx":1DCA4
      PicHot          =   "Form_IDENTITAS_PRIBADI.frx":1E0F6
      PicNormal       =   "Form_IDENTITAS_PRIBADI.frx":1E548
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
Attribute VB_Name = "Form_IDENTITAS_PRIBADI"
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
            .RecordSource = "Select * From tbIdentitasPribadi order by Nama_Lengkap asc;"
            .Refresh
        End With
End Sub
Sub IsiCMBAgama()
NyambunggUtama
With AdodcAgama
    .ConnectionString = CN.ConnectionString
    .RecordSource = "Select * from tbAgama"
    .Refresh
End With
    With cmbAgama
        .Clear
        Do Until AdodcAgama.Recordset.EOF
            cmbAgama.AddItem AdodcAgama.Recordset.Fields(1).Value
            AdodcAgama.Recordset.MoveNext
        Loop
        AdodcAgama.Refresh
        .Text = AdodcAgama.Recordset.Fields(1).Value
    End With
End Sub
Sub IsiCMBGolonganDarah()
NyambunggUtama
With AdodcGolonganDarah
    .ConnectionString = CN.ConnectionString
    .RecordSource = "Select * from tbGolonganDarah"
    .Refresh
End With
    With cmbGolonganDarah
        .Clear
        Do Until AdodcGolonganDarah.Recordset.EOF
            cmbGolonganDarah.AddItem AdodcGolonganDarah.Recordset.Fields(0).Value
            AdodcGolonganDarah.Recordset.MoveNext
        Loop
        AdodcGolonganDarah.Refresh
        .Text = AdodcGolonganDarah.Recordset.Fields(0).Value
    End With
End Sub
Sub IsiCMBStatusHubungan()
NyambunggUtama
With AdodcStatusHubungan
    .ConnectionString = CN.ConnectionString
    .RecordSource = "Select * from tbStatusHubungan"
    .Refresh
End With
    With cmbStatusHubungan
        .Clear
        Do Until AdodcStatusHubungan.Recordset.EOF
            cmbStatusHubungan.AddItem AdodcStatusHubungan.Recordset.Fields(0).Value
            AdodcStatusHubungan.Recordset.MoveNext
        Loop
        AdodcStatusHubungan.Refresh
        .Text = AdodcStatusHubungan.Recordset.Fields(0).Value
    End With
End Sub
Sub AturKontrol()
    SambungkanKontrolKeADODC_UTAMA
    IsiCMBAgama
    IsiCMBGolonganDarah
    IsiCMBStatusHubungan
    For Each Objek In Me
        If TypeName(Objek) = "XPText" Then
            With Objek
                .ForeColor = Hitam
                .MaxLength = 254
            End With
        End If
    Next
    IsiCMBDataLalu
    With cmbJenisKelamin
        .Clear
        .AddItem "Pria / Male", 0
        .AddItem "Wanita / Female", 1
        .ListIndex = 0
    End With
    cmHapus(3).Enabled = False
    cmHapus(4).Enabled = False
    cmHapus(5).Enabled = False
    cmHapus(17).Enabled = False
    cmSalin(3).Enabled = False
    cmSalin(4).Enabled = False
    cmSalin(5).Enabled = False
    cmSalin(17).Enabled = False
    With cmbBulan
        .Clear
        .AddItem "01", 0
        .AddItem "02", 1
        .AddItem "03", 2
        .AddItem "04", 3
        .AddItem "05", 4
        .AddItem "06", 5
        .AddItem "07", 6
        .AddItem "08", 7
        .AddItem "09", 8
        .AddItem "10", 9
        .AddItem "11", 10
        .AddItem "12", 11
        .ListIndex = 0
    End With
    cmbTahun.Clear
    For Z = 1800 To 3000
        cmbTahun.AddItem Z
    Next
    cmbTahun.Text = Year(Date)
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
    With Me
        .cmbDataLalu1.Clear
        .cmbDataLalu2.Clear
                
        .cmbDataLalu7.Clear
        .cmbDataLalu8.Clear
        .cmbDataLalu9.Clear
        .cmbDataLalu10.Clear
        .cmbDataLalu11.Clear
        .cmbDataLalu12.Clear
        .cmbDataLalu13.Clear
        .cmbDataLalu14.Clear
        .cmbDataLalu15.Clear
        .cmbDataLalu16.Clear
        .cmbDataLalu17.Clear
        
        .cmbDataLalu19.Clear
        .cmbDataLalu20.Clear
        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
            .cmbDataLalu1.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(0).Value
            .cmbDataLalu2.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(1).Value
            
            .cmbDataLalu7.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(9).Value
            .cmbDataLalu8.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(10).Value
            .cmbDataLalu9.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(11).Value
            .cmbDataLalu10.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(12).Value
            .cmbDataLalu11.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(13).Value
            .cmbDataLalu12.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(14).Value
            .cmbDataLalu13.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(15).Value
            .cmbDataLalu14.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(16).Value
            .cmbDataLalu15.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(17).Value
            .cmbDataLalu16.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(18).Value
            .cmbDataLalu17.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(19).Value
            
            .cmbDataLalu19.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(21).Value
            .cmbDataLalu20.AddItem FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(22).Value
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
        .Label1.Caption = "Nama Lengkap"
        .Label2.Caption = "Nama Panggilan"
        .Label3.Caption = "T.T.L"
        .Label4.Caption = "Jenis Kelamin"
        .Label5.Caption = "Agama"
        .Label6.Caption = "Golongan Darah"
        .Label7.Caption = "Pekerjaan"
        .Label8.Caption = "Alamat Rumah"
        .Label9.Caption = "Alamat E-Mail"
        .Label10.Caption = "Alamat Website"
        .Label11.Caption = "Nomor Telepon"
        .Label12.Caption = "Kota Asal"
        .Label13.Caption = "Kota Sekarang"
        .Label14.Caption = "Kode Pos"
        .Label15.Caption = "Provinsi"
        .Label16.Caption = "Kewarganegaraan"
        .Label17.Caption = "Status Pendidikan"
        .Label18.Caption = "Status Hubungan"
        .Label19.Caption = "Hobby"
        .Label20.Caption = "Keterangan"
        'BARU SELESAI SAMPAI DISINI
        For NomorIndex = 0 To 19
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
        .Label1.Caption = "Full Name"
        .Label2.Caption = "Cool Name"
        .Label3.Caption = "Place/Date of birth"
        .Label4.Caption = "Gender"
        .Label5.Caption = "Religion"
        .Label6.Caption = "Blood Type"
        .Label7.Caption = "Jobs"
        .Label8.Caption = "Home Address"
        .Label9.Caption = "E-Mail Address"
        .Label10.Caption = "Web Address"
        .Label11.Caption = "Phone Number"
        .Label12.Caption = "Hometown"
        .Label13.Caption = "City Now"
        .Label14.Caption = "ZIP Code"
        .Label15.Caption = "Province"
        .Label16.Caption = "Citizenship"
        .Label17.Caption = "Educational Status"
        .Label18.Caption = "Relationship Status"
        .Label19.Caption = "Hobby"
        .Label20.Caption = "Description"
        For NomorIndex = 0 To 19
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
                    .Recordset.Fields(0).Value = textNamaLengkap.Text
                    .Recordset.Fields(1).Value = textNamaPanggilan.Text
                    .Recordset.Fields(2).Value = textTempat.Text
                    .Recordset.Fields(3).Value = cmbTanggal.Text
                    .Recordset.Fields(4).Value = cmbBulan.Text
                    .Recordset.Fields(5).Value = cmbTahun.Text
                    .Recordset.Fields(6).Value = cmbJenisKelamin.Text
                    .Recordset.Fields(7).Value = cmbAgama.Text
                    .Recordset.Fields(8).Value = cmbGolonganDarah.Text
                    .Recordset.Fields(9).Value = textPekerjaan.Text
                    .Recordset.Fields(10).Value = textAlamatRumah.Text
                    .Recordset.Fields(11).Value = textAlamatEmail.Text
                    .Recordset.Fields(12).Value = textAlamatWebsite.Text
                    .Recordset.Fields(13).Value = textNomorTelepon.Text
                    .Recordset.Fields(14).Value = textKotaAsal.Text
                    .Recordset.Fields(15).Value = textKotaSekarang.Text
                    .Recordset.Fields(16).Value = textKodePos.Text
                    .Recordset.Fields(17).Value = textProvinsi.Text
                    .Recordset.Fields(18).Value = textKewargaNegaraan.Text
                    .Recordset.Fields(19).Value = textStatusPendidikan.Text
                    .Recordset.Fields(20).Value = cmbStatusHubungan.Text
                    .Recordset.Fields(21).Value = textHobby.Text
                    .Recordset.Fields(22).Value = textKeterangan.Text
                    .Recordset.Update
                    .Refresh
                End With
            ElseIf Me.cmSimpan.Caption = "&Perbarui" Or Me.cmSimpan.Caption = "&Update" Then
                With FormManage.AdodcMain
                    .Recordset.Delete
                    .Recordset.AddNew
                    .Recordset.Fields(0).Value = textNamaLengkap.Text
                    .Recordset.Fields(1).Value = textNamaPanggilan.Text
                    .Recordset.Fields(2).Value = textTempat.Text
                    .Recordset.Fields(3).Value = cmbTanggal.Text
                    .Recordset.Fields(4).Value = cmbBulan.Text
                    .Recordset.Fields(5).Value = cmbTahun.Text
                    .Recordset.Fields(6).Value = cmbJenisKelamin.Text
                    .Recordset.Fields(7).Value = cmbAgama.Text
                    .Recordset.Fields(8).Value = cmbGolonganDarah.Text
                    .Recordset.Fields(9).Value = textPekerjaan.Text
                    .Recordset.Fields(10).Value = textAlamatRumah.Text
                    .Recordset.Fields(11).Value = textAlamatEmail.Text
                    .Recordset.Fields(12).Value = textAlamatWebsite.Text
                    .Recordset.Fields(13).Value = textNomorTelepon.Text
                    .Recordset.Fields(14).Value = textKotaAsal.Text
                    .Recordset.Fields(15).Value = textKotaSekarang.Text
                    .Recordset.Fields(16).Value = textKodePos.Text
                    .Recordset.Fields(17).Value = textProvinsi.Text
                    .Recordset.Fields(18).Value = textKewargaNegaraan.Text
                    .Recordset.Fields(19).Value = textStatusPendidikan.Text
                    .Recordset.Fields(20).Value = cmbStatusHubungan.Text
                    .Recordset.Fields(21).Value = textHobby.Text
                    .Recordset.Fields(22).Value = textKeterangan.Text
                    .Recordset.Update
                    .Refresh
                    FormManage.AturDatabase
                End With
            End If
                With FormPengaturan
                    If .cekAutoRefresh.Value = Checked Then FORM_UTAMA.cmIdentitasPribadi_Click
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
                    .Recordset.Fields(0).Value = textNamaLengkap.Text
                    .Recordset.Fields(1).Value = textNamaPanggilan.Text
                    .Recordset.Fields(2).Value = textTempat.Text
                    .Recordset.Fields(3).Value = cmbTanggal.Text
                    .Recordset.Fields(4).Value = cmbBulan.Text
                    .Recordset.Fields(5).Value = cmbTahun.Text
                    .Recordset.Fields(6).Value = cmbJenisKelamin.Text
                    .Recordset.Fields(7).Value = cmbAgama.Text
                    .Recordset.Fields(8).Value = cmbGolonganDarah.Text
                    .Recordset.Fields(9).Value = textPekerjaan.Text
                    .Recordset.Fields(10).Value = textAlamatRumah.Text
                    .Recordset.Fields(11).Value = textAlamatEmail.Text
                    .Recordset.Fields(12).Value = textAlamatWebsite.Text
                    .Recordset.Fields(13).Value = textNomorTelepon.Text
                    .Recordset.Fields(14).Value = textKotaAsal.Text
                    .Recordset.Fields(15).Value = textKotaSekarang.Text
                    .Recordset.Fields(16).Value = textKodePos.Text
                    .Recordset.Fields(17).Value = textProvinsi.Text
                    .Recordset.Fields(18).Value = textKewargaNegaraan.Text
                    .Recordset.Fields(19).Value = textStatusPendidikan.Text
                    .Recordset.Fields(20).Value = cmbStatusHubungan.Text
                    .Recordset.Fields(21).Value = textHobby.Text
                    .Recordset.Fields(22).Value = textKeterangan.Text
                    .Recordset.Update
                    .Refresh
                End With
            ElseIf Me.cmSimpan.Caption = "&Perbarui" Or Me.cmSimpan.Caption = "&Update" Then
                With FormManage.AdodcMain
                    .Recordset.Delete
                    .Recordset.AddNew
                    .Recordset.Fields(0).Value = textNamaLengkap.Text
                    .Recordset.Fields(1).Value = textNamaPanggilan.Text
                    .Recordset.Fields(2).Value = textTempat.Text
                    .Recordset.Fields(3).Value = cmbTanggal.Text
                    .Recordset.Fields(4).Value = cmbBulan.Text
                    .Recordset.Fields(5).Value = cmbTahun.Text
                    .Recordset.Fields(6).Value = cmbJenisKelamin.Text
                    .Recordset.Fields(7).Value = cmbAgama.Text
                    .Recordset.Fields(8).Value = cmbGolonganDarah.Text
                    .Recordset.Fields(9).Value = textPekerjaan.Text
                    .Recordset.Fields(10).Value = textAlamatRumah.Text
                    .Recordset.Fields(11).Value = textAlamatEmail.Text
                    .Recordset.Fields(12).Value = textAlamatWebsite.Text
                    .Recordset.Fields(13).Value = textNomorTelepon.Text
                    .Recordset.Fields(14).Value = textKotaAsal.Text
                    .Recordset.Fields(15).Value = textKotaSekarang.Text
                    .Recordset.Fields(16).Value = textKodePos.Text
                    .Recordset.Fields(17).Value = textProvinsi.Text
                    .Recordset.Fields(18).Value = textKewargaNegaraan.Text
                    .Recordset.Fields(19).Value = textStatusPendidikan.Text
                    .Recordset.Fields(20).Value = cmbStatusHubungan.Text
                    .Recordset.Fields(21).Value = textHobby.Text
                    .Recordset.Fields(22).Value = textKeterangan.Text
                    .Recordset.Update
                    .Refresh
                End With
                FormManage.AturDatabase
            End If
        With FormPengaturan
            If .cekAutoRefresh.Value = Checked Then FORM_UTAMA.cmIdentitasPribadi_Click
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
            .ForeColor = Hitam
            .Text = ""
        End With
    End If
Next
End Sub

Private Sub cmBantuan_Click()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Kalimat = App.Path & "\bantuan\html\IdentitasPribadi.html"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Kalimat = App.Path & "\bantuan\html\PersonalBiodata.html"
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

Private Sub cmbBulan_Click()
    cmbTanggal.Clear
    Select Case cmbBulan.ListIndex
        Case Is = 0, 2, 4, 6, 7, 9, 11
            For Z = 1 To 31
                With cmbTanggal
                    .AddItem Z
                End With
            Next
        Case Is = 1
            If Val(cmbTahun.Text) Mod 4 Then
                For Z = 1 To 29
                    With cmbTanggal
                        .AddItem Z
                    End With
                Next
            Else
                For Z = 1 To 28
                    With cmbTanggal
                        .AddItem Z
                    End With
                Next
            End If
        Case Is = 3, 5, 8, 10
            For Z = 1 To 30
                With cmbTanggal
                    .AddItem Z
                End With
            Next
    End Select
    cmbTanggal.Text = "1"
End Sub

Private Sub cmbDataLalu1_Click()
    With textNamaLengkap
        .Text = cmbDataLalu1.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu10_Click()
    With textAlamatWebsite
        .Text = cmbDataLalu10.Text
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

Private Sub cmbDataLalu7_Click()
    With textPekerjaan
        .Text = cmbDataLalu7.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu8_Click()
    With textAlamatRumah
        .Text = cmbDataLalu8.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbDataLalu9_Click()
    With textAlamatEmail
        .Text = cmbDataLalu9.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub
Private Sub cmbDataLalu11_Click()
    With textNomorTelepon
        .Text = cmbDataLalu11.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub
Private Sub cmbDataLalu12_Click()
    With textKotaAsal
        .Text = cmbDataLalu12.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub
Private Sub cmbDataLalu13_Click()
    With textKotaSekarang
        .Text = cmbDataLalu13.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub
Private Sub cmbDataLalu14_Click()
    With textKodePos
        .Text = cmbDataLalu14.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub
Private Sub cmbDataLalu15_Click()
    With textProvinsi
        .Text = cmbDataLalu15.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub
Private Sub cmbDataLalu16_Click()
    With textKewargaNegaraan
        .Text = cmbDataLalu16.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub
Private Sub cmbDataLalu17_Click()
    With textStatusPendidikan
        .Text = cmbDataLalu17.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub
Private Sub cmbDataLalu19_Click()
    With textHobby
        .Text = cmbDataLalu19.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub
Private Sub cmbDataLalu20_Click()
    With textKeterangan
        .Text = cmbDataLalu20.Text
        .ForeColor = Hitam
        .SetFocus
    End With
End Sub

Private Sub cmbTahun_Click()
    cmbBulan_Click
End Sub

Private Sub cmHapus_Click(Index As Integer)
Select Case Index
    Case Is = 0
        If textNamaLengkap.Text = IsiTextBoxKosong_ID(0) Or textNamaLengkap.Text = IsiTextBoxKosong_EN(0) Then
            With textNamaLengkap
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(0)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(0)
                End If
            End With
        Else
            With textNamaLengkap
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
        'Ga terjadi apa apa karena objeknya combobox
    Case Is = 3
        'Ga terjadi apa apa karena objeknya combobox
    Case Is = 4
        'Ga terjadi apa apa karena objeknya combobox
    Case Is = 5
        'Ga terjadi apa apa karena objeknya combobox
    Case Is = 6
        If textPekerjaan.Text = IsiTextBoxKosong_ID(6) Or textPekerjaan.Text = IsiTextBoxKosong_EN(6) Then
            With textPekerjaan
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(6)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(6)
                End If
            End With
        Else
            With textPekerjaan
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 7
        If textAlamatRumah.Text = IsiTextBoxKosong_ID(7) Or textAlamatRumah.Text = IsiTextBoxKosong_EN(7) Then
            With textAlamatRumah
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(7)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(7)
                End If
            End With
        Else
            With textAlamatRumah
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 8
        If textAlamatEmail.Text = IsiTextBoxKosong_ID(8) Or textAlamatEmail.Text = IsiTextBoxKosong_EN(8) Then
            With textAlamatEmail
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(8)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(8)
                End If
            End With
        Else
            With textAlamatEmail
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 9
        If textAlamatWebsite.Text = IsiTextBoxKosong_ID(9) Or textAlamatWebsite.Text = IsiTextBoxKosong_EN(9) Then
            With textAlamatWebsite
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(9)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(9)
                End If
            End With
        Else
            With textAlamatWebsite
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 10
        If textNomorTelepon.Text = IsiTextBoxKosong_ID(10) Or textAlamatWebsite.Text = IsiTextBoxKosong_EN(10) Then
            With textNomorTelepon
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(10)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(10)
                End If
            End With
        Else
            With textNomorTelepon
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 11
        If textKotaAsal.Text = IsiTextBoxKosong_ID(11) Or textKotaAsal.Text = IsiTextBoxKosong_EN(11) Then
            With textKotaAsal
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(11)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(11)
                End If
            End With
        Else
            With textKotaAsal
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 12
        If textKotaSekarang.Text = IsiTextBoxKosong_ID(12) Or textKotaSekarang.Text = IsiTextBoxKosong_EN(12) Then
            With textKotaSekarang
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(12)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(12)
                End If
            End With
        Else
            With textKotaSekarang
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 13
        If textKodePos.Text = IsiTextBoxKosong_ID(13) Or textKodePos.Text = IsiTextBoxKosong_EN(13) Then
            With textKodePos
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(13)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(13)
                End If
            End With
        Else
            With textKodePos
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 14
        If textProvinsi.Text = IsiTextBoxKosong_ID(14) Or textProvinsi.Text = IsiTextBoxKosong_EN(14) Then
            With textProvinsi
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(14)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(14)
                End If
            End With
        Else
            With textProvinsi
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 15
        If textKewargaNegaraan.Text = IsiTextBoxKosong_ID(15) Or textKewargaNegaraan.Text = IsiTextBoxKosong_EN(15) Then
            With textKewargaNegaraan
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(15)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(15)
                End If
            End With
        Else
            With textKewargaNegaraan
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 16
        If textStatusPendidikan.Text = IsiTextBoxKosong_ID(16) Or textStatusPendidikan.Text = IsiTextBoxKosong_EN(16) Then
            With textStatusPendidikan
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(16)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(16)
                End If
            End With
        Else
            With textStatusPendidikan
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 17
        'Ga terjadi apa apa karena objeknya combobox
    Case Is = 18
        If textHobby.Text = IsiTextBoxKosong_ID(18) Or textHobby.Text = IsiTextBoxKosong_EN(18) Then
            With textHobby
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(18)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(18)
                End If
            End With
        Else
            With textHobby
                .Text = ""
                .SetFocus
            End With
        End If
    Case Is = 19
        If textKeterangan.Text = IsiTextBoxKosong_ID(19) Or textKeterangan.Text = IsiTextBoxKosong_EN(19) Then
            With textKeterangan
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Text = IsiTextBoxKosong_ID(19)
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Text = IsiTextBoxKosong_EN(19)
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
        If textAlamatEmail.Text = "" Or textAlamatEmail.Text = IsiTextBoxKosong_ID(8) Or textAlamatEmail.Text = IsiTextBoxKosong_EN(8) Then
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
        If textAlamatWebsite.Text = "" Or textAlamatWebsite.Text = IsiTextBoxKosong_ID(9) Or textAlamatWebsite.Text = IsiTextBoxKosong_EN(9) Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                MsgBox "Silahkan isi alamat URL yang ingin di cek!", vbExclamation + vbOKOnly, ""
            Else
                MsgBox "Put in URL Address for checking!", vbExclamation + vbOKOnly, ""
            End If
                textAlamatWebsite.SetFocus
        Else
            AlamatSitus = ShellExecute(0, vbNullString, _
               textAlamatWebsite.Text, "", "", vbNormalFocus)
        End If
    End Select
End Sub

Private Sub cmReset_Click()
    KosongkanTextBox
End Sub

Private Sub cmSalin_Click(Index As Integer)
Select Case Index
    Case Is = 0
        If textNamaLengkap.Text = "" Or textNamaLengkap.Text = IsiTextBoxKosong_ID(0) Or textNamaLengkap.Text = IsiTextBoxKosong_EN(0) Then
            KhususCmSalin
            textNamaLengkap.SetFocus
        Else
            With textNamaLengkap
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textNamaLengkap.Text)
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
        'ga kejadian apa-apa
    Case Is = 3
        'ga kejadian apa-apa
    Case Is = 4
        'ga kejadian apa-apa
    Case Is = 5
        'ga kejadian apa-apa
    Case Is = 6
        If textPekerjaan.Text = "" Or textPekerjaan.Text = IsiTextBoxKosong_ID(6) Or textPekerjaan.Text = IsiTextBoxKosong_EN(6) Then
            KhususCmSalin
            textPekerjaan.SetFocus
        Else
            With textPekerjaan
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textPekerjaan.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 7
        If textAlamatRumah.Text = "" Or textAlamatRumah.Text = IsiTextBoxKosong_ID(7) Or textAlamatRumah.Text = IsiTextBoxKosong_EN(7) Then
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
    Case Is = 8
        If textAlamatEmail.Text = "" Or textAlamatEmail.Text = IsiTextBoxKosong_ID(8) Or textAlamatEmail.Text = IsiTextBoxKosong_EN(8) Then
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
    Case Is = 9
        If textAlamatWebsite.Text = "" Or textAlamatWebsite.Text = IsiTextBoxKosong_ID(9) Or textAlamatWebsite.Text = IsiTextBoxKosong_EN(9) Then
            KhususCmSalin
            textAlamatWebsite.SetFocus
        Else
            With textAlamatWebsite
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textAlamatWebsite.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 10
        If textNomorTelepon.Text = "" Or textNomorTelepon.Text = IsiTextBoxKosong_ID(10) Or textNomorTelepon.Text = IsiTextBoxKosong_EN(10) Then
            KhususCmSalin
            textNomorTelepon.SetFocus
        Else
            With textNomorTelepon
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textNomorTelepon.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 11
        If textKotaAsal.Text = "" Or textKotaAsal.Text = IsiTextBoxKosong_ID(11) Or textKotaAsal.Text = IsiTextBoxKosong_EN(11) Then
            KhususCmSalin
            textKotaAsal.SetFocus
        Else
            With textKotaAsal
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textKotaAsal.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 12
        If textKotaSekarang.Text = "" Or textKotaSekarang.Text = IsiTextBoxKosong_ID(12) Or textKotaSekarang.Text = IsiTextBoxKosong_EN(12) Then
            KhususCmSalin
            textKotaSekarang.SetFocus
        Else
            With textKotaSekarang
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textKotaSekarang.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 13
        If textKodePos.Text = "" Or textKodePos.Text = IsiTextBoxKosong_ID(13) Or textKotaSekarang.Text = IsiTextBoxKosong_EN(13) Then
            KhususCmSalin
            textKodePos.SetFocus
        Else
            With textKodePos
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textKodePos.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 14
        If textProvinsi.Text = "" Or textProvinsi.Text = IsiTextBoxKosong_ID(14) Or textProvinsi.Text = IsiTextBoxKosong_EN(14) Then
            KhususCmSalin
            textProvinsi.SetFocus
        Else
            With textProvinsi
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textProvinsi.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 15
        If textKewargaNegaraan.Text = "" Or textKewargaNegaraan.Text = IsiTextBoxKosong_ID(15) Or textKewargaNegaraan.Text = IsiTextBoxKosong_EN(15) Then
            KhususCmSalin
            textKewargaNegaraan.SetFocus
        Else
            With textKewargaNegaraan
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textKewargaNegaraan.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 16
        If textStatusPendidikan.Text = "" Or textStatusPendidikan.Text = IsiTextBoxKosong_ID(16) Or textStatusPendidikan.Text = IsiTextBoxKosong_EN(16) Then
            KhususCmSalin
            textStatusPendidikan.SetFocus
        Else
            With textStatusPendidikan
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textStatusPendidikan.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 18
        If textHobby.Text = "" Or textHobby.Text = IsiTextBoxKosong_ID(18) Or textHobby.Text = IsiTextBoxKosong_EN(18) Then
            KhususCmSalin
            textHobby.SetFocus
        Else
            With textHobby
                Clipboard.Clear
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(textHobby.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 19
        If textKeterangan.Text = "" Or textKeterangan.Text = IsiTextBoxKosong_ID(19) Or textKeterangan.Text = IsiTextBoxKosong_EN(19) Then
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
If textNamaLengkap.Text = "" Or textNamaLengkap.Text = IsiTextBoxKosong_ID(0) Or textNamaLengkap.Text = IsiTextBoxKosong_EN(0) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Nama Lengkap Anda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the your name!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textNamaLengkap.SetFocus
ElseIf textNamaPanggilan.Text = "" Or textNamaPanggilan.Text = IsiTextBoxKosong_ID(1) Or textNamaPanggilan.Text = IsiTextBoxKosong_EN(1) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Nama Panggilan Anda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the your cool name!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textNamaPanggilan.SetFocus
ElseIf textTempat.Text = "" Or textTempat.Text = IsiTextBoxKosong_ID(2) Or textTempat.Text = IsiTextBoxKosong_EN(2) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi tempat dan tanggal lahir Anda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the place and your born day!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textTempat.SetFocus
ElseIf textPekerjaan.Text = "" Or textPekerjaan.Text = IsiTextBoxKosong_ID(6) Or textPekerjaan.Text = IsiTextBoxKosong_EN(6) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi apa pekerjaan Anda sehari2!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the your jobs!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textPekerjaan.SetFocus
ElseIf textAlamatRumah.Text = "" Or textAlamatRumah.Text = IsiTextBoxKosong_ID(7) Or textAlamatRumah.Text = IsiTextBoxKosong_EN(7) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi pertanyaan Alamat rumah Anda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the your address!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textAlamatRumah.SetFocus
ElseIf textAlamatEmail.Text = "" Or textAlamatEmail.Text = IsiTextBoxKosong_ID(8) Or textAlamatEmail.Text = IsiTextBoxKosong_EN(8) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi alamat email Anda " & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the your email address!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textAlamatEmail.SetFocus
ElseIf textAlamatWebsite.Text = "" Or textAlamatWebsite.Text = IsiTextBoxKosong_ID(9) Or textAlamatWebsite.Text = IsiTextBoxKosong_EN(9) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi alamat website Anda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write your website address!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textAlamatWebsite.SetFocus
ElseIf textNomorTelepon.Text = "" Or textNomorTelepon.Text = IsiTextBoxKosong_ID(10) Or textNomorTelepon.Text = IsiTextBoxKosong_EN(10) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi nomor telepon Anda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write your phone number!" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textNomorTelepon.SetFocus
ElseIf textKotaAsal.Text = "" Or textKotaAsal.Text = IsiTextBoxKosong_ID(11) Or textKotaAsal.Text = IsiTextBoxKosong_EN(11) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan nama kota asal Anda!" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the your home town" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textKotaAsal.SetFocus
ElseIf textKotaSekarang.Text = "" Or textKotaSekarang.Text = IsiTextBoxKosong_ID(12) Or textKotaSekarang.Text = IsiTextBoxKosong_EN(12) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi nama kota tempat tinggal Anda saat ini" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the name of your town at this time" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textKotaSekarang.SetFocus
ElseIf textKodePos.Text = "" Or textKodePos.Text = IsiTextBoxKosong_ID(13) Or textKodePos.Text = IsiTextBoxKosong_EN(13) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi kode pos tempat tinggal Anda" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the your ZIP/postal code" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textKodePos.SetFocus
ElseIf textProvinsi.Text = "" Or textProvinsi.Text = IsiTextBoxKosong_ID(14) Or textProvinsi.Text = IsiTextBoxKosong_EN(14) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi nama provinsi tempat tinggal Anda saat ini" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the name of your state at this time" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textProvinsi.SetFocus
ElseIf textKewargaNegaraan.Text = "" Or textKewargaNegaraan.Text = IsiTextBoxKosong_ID(15) Or textKewargaNegaraan.Text = IsiTextBoxKosong_EN(15) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi nama negara tempat tinggal Anda" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the name of your Citizenship" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textKewargaNegaraan.SetFocus
ElseIf textStatusPendidikan.Text = "" Or textStatusPendidikan.Text = IsiTextBoxKosong_ID(16) Or textStatusPendidikan.Text = IsiTextBoxKosong_EN(16) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Status Pendidikan" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the name of your Educational Status" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textStatusPendidikan.SetFocus
ElseIf textHobby.Text = "" Or textHobby.Text = IsiTextBoxKosong_ID(18) Or textHobby.Text = IsiTextBoxKosong_EN(18) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Hobby Anda" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the name of your your hobby" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textHobby.SetFocus
ElseIf textKeterangan.Text = "" Or textKeterangan.Text = IsiTextBoxKosong_ID(19) Or textKeterangan.Text = IsiTextBoxKosong_EN(19) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Keterangan lain" & vbCrLf & _
                "Jika ingin dikosongkan, tambahkan tanda '-' atau klik 'Verifikasi'", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Please write the other description" & vbCrLf & _
                "If you want be emptied, insert the symbol '-' or click 'Verify", vbExclamation + vbOKOnly, ""
    End If
    textKeterangan.SetFocus
Else
    SIMPAN_KE_DATABASE
    IsiCMBDataLalu
End If
End Sub


Private Sub cmTambahDataAgama_Click()
    With FormTambahAgamaUntukIDP
        .Show vbModal, Me
    End With
End Sub

Private Sub cmTambahDataGolonganDarah_Click()
With FormTambahGolonganDarah
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        .Caption = "Tambah Golongan Darah"
        .Label1.Caption = "Tambah Golongan Darah"
        .cmOK.Caption = "&OK/Simpan"
    Else
        .Caption = "Add New Blood Type"
        .Label1.Caption = "Add New Blood Type"
        .cmOK.Caption = "&OK/Save"
    End If
    .textGolonganDarah.Text = ""
    .Show vbModal, Me
End With
End Sub

Private Sub cmTambahDataStatusHubungan_Click()
With FormTambahStatusHubungan
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        .Caption = "Tambah Status Hubungan"
        .Label1.Caption = "Tambah Status Hubungan"
        .cmOK.Caption = "&OK/Simpan"
    Else
        .Caption = "Add New Relationship"
        .Label1.Caption = "Add New Relationship"
        .cmOK.Caption = "&OK/Save"
    End If
    .TextStatusHubungan.Text = ""
    .Show vbModal, Me
End With
End Sub

Private Sub cmVerifikasi_Click()
If textNamaLengkap.Text = "" Or textNamaLengkap.Text = IsiTextBoxKosong_ID(0) Or textNamaLengkap.Text = IsiTextBoxKosong_EN(0) Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Pesan = MsgBox("Nama Server belum terisi, yakin ingin mem-verifikasi?", vbQuestion + vbYesNo, "Nama Server")
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Pesan = MsgBox("Server Name is Empty!, Are you sure to Verify entry?", vbQuestion + vbYesNo, "Server Name?")
    End If
        If Pesan = vbYes Then
            For Each Objek In Me
                If TypeName(Objek) = "XPText" Then
                    If Objek.Text = "" Or Objek.ForeColor = Hitam Then
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
            If Objek.Text = "" Or Objek.ForeColor = Hitam Then
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
    KosongkanTextBox
    PENGATURAN_BAHASA
    PENGATURAN_WARNA
End Sub

Private Sub TextAlamatEmail_DblClick()
       R = SendMessageLong(cmbDataLalu9.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textAlamatRumah_DblClick()
       R = SendMessageLong(cmbDataLalu8.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textAlamatWebsite_DblClick()
       R = SendMessageLong(cmbDataLalu10.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textHobby_DblClick()
       R = SendMessageLong(cmbDataLalu19.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textKeterangan_DblClick()
       R = SendMessageLong(cmbDataLalu20.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textKewargaNegaraan_DblClick()
       R = SendMessageLong(cmbDataLalu16.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textKodePos_DblClick()
       R = SendMessageLong(cmbDataLalu14.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textKotaAsal_DblClick()
       R = SendMessageLong(cmbDataLalu12.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textKotaSekarang_DblClick()
       R = SendMessageLong(cmbDataLalu13.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textNamaLengkap_DblClick()
       R = SendMessageLong(cmbDataLalu1.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textNamaPanggilan_DblClick()
       R = SendMessageLong(cmbDataLalu2.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textNomorTelepon_DblClick()
       R = SendMessageLong(cmbDataLalu11.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textPekerjaan_DblClick()
       R = SendMessageLong(cmbDataLalu7.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textProvinsi_DblClick()
       R = SendMessageLong(cmbDataLalu15.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textStatusPendidikan_DblClick()
       R = SendMessageLong(cmbDataLalu17.hwnd, CB_SHOWDROPDOWN, True, 0)
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
