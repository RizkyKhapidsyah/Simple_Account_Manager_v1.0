VERSION 5.00
Object = "{02353968-C1C9-4E0A-88D3-18759BDC60FE}#1.0#0"; "AeroSuite.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormManage 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Data"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormManage.frx":0442
   ScaleHeight     =   8115
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin AeroSuite.AeroGroupBox AeroGroupBox1 
      Height          =   6975
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   12303
      BorderColor     =   11908533
      BackColor       =   14737632
      BackColor2      =   13882323
      HeadColor1      =   14737632
      HeadColor2      =   13092807
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Dacara_dcButton.dcButton cmEdit 
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1508
         BackColor       =   12230304
         ButtonStyle     =   3
         Caption         =   "&Edit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   7
         PicDown         =   "FormManage.frx":2CD34
         PicHot          =   "FormManage.frx":2D04E
         PicNormal       =   "FormManage.frx":2D368
         PicSizeH        =   32
         PicSizeW        =   32
      End
      Begin Dacara_dcButton.dcButton cmHapus 
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1508
         BackColor       =   12230304
         ButtonStyle     =   3
         Caption         =   "&Hapus"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   7
         PicDown         =   "FormManage.frx":2D682
         PicHot          =   "FormManage.frx":2DAD4
         PicNormal       =   "FormManage.frx":2DF26
         PicSizeH        =   32
         PicSizeW        =   32
      End
      Begin Dacara_dcButton.dcButton cmCari 
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   4080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1508
         BackColor       =   12230304
         ButtonStyle     =   3
         Caption         =   "&Cari"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   7
         PicDown         =   "FormManage.frx":2E378
         PicHot          =   "FormManage.frx":2E692
         PicNormal       =   "FormManage.frx":2E9AC
         PicSizeH        =   32
         PicSizeW        =   32
      End
      Begin Dacara_dcButton.dcButton cmTutup 
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   6000
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1508
         BackColor       =   12230304
         ButtonStyle     =   3
         Caption         =   "&Tutup"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   7
         PicDown         =   "FormManage.frx":2ECC6
         PicHot          =   "FormManage.frx":2F118
         PicNormal       =   "FormManage.frx":2F56A
         PicSizeH        =   32
         PicSizeW        =   32
      End
      Begin Dacara_dcButton.dcButton cmBantuan 
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   5040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1508
         BackColor       =   12230304
         ButtonStyle     =   3
         Caption         =   "&Bantuan"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   7
         PicDown         =   "FormManage.frx":2F9BC
         PicHot          =   "FormManage.frx":2FE0E
         PicNormal       =   "FormManage.frx":30260
         PicSizeH        =   32
         PicSizeW        =   32
      End
      Begin Dacara_dcButton.dcButton cmSorot 
         Height          =   855
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1508
         BackColor       =   12230304
         ButtonStyle     =   3
         Caption         =   "&Sorot"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   7
         PicDown         =   "FormManage.frx":306B2
         PicHot          =   "FormManage.frx":37BB4
         PicNormal       =   "FormManage.frx":3F0B6
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
      End
      Begin Dacara_dcButton.dcButton cmFilter 
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1508
         BackColor       =   12230304
         ButtonStyle     =   3
         Caption         =   "&Filter"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   7
         PicDown         =   "FormManage.frx":465B8
         PicHot          =   "FormManage.frx":468D2
         PicNormal       =   "FormManage.frx":46BEC
         PicSizeH        =   32
         PicSizeW        =   32
      End
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   7680
      Top             =   4920
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   3
      Bmp:1           =   "FormManage.frx":46F06
      Mask:1          =   16777215
      Key:1           =   "#menuRefresh"
      Bmp:2           =   "FormManage.frx":47348
      Mask:2          =   3096113
      Key:2           =   "#menuBS"
      Bmp:3           =   "FormManage.frx":4769A
      Key:3           =   "#Properties"
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
      Height          =   735
      Left            =   7800
      TabIndex        =   1
      Top             =   6960
      Width           =   2655
      Begin Dacara_dcButton.dcButton cmTambah 
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BackColor       =   10591645
         ButtonStyle     =   2
         Caption         =   "+"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         SmallChange     =   10
         Min             =   10
         Max             =   700
         SelStart        =   10
         TickStyle       =   3
         Value           =   10
      End
      Begin Dacara_dcButton.dcButton cmKurang 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BackColor       =   10591645
         ButtonStyle     =   2
         Caption         =   "-"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc AdodcMain 
      Height          =   375
      Left            =   7320
      Top             =   3600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6375
      Left            =   1440
      TabIndex        =   0
      Top             =   1440
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11245
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox textPersen 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9240
      TabIndex        =   6
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin XPEngine.XPControl XP_Engine 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.Label LabelPersen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   9720
      TabIndex        =   5
      Top             =   7999
      Width           =   315
   End
   Begin VB.Menu menuTersembunyi 
      Caption         =   "Menu Tersembunyi"
      Begin VB.Menu menuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu menuBS 
         Caption         =   "Bersihkan Sorot / Filter"
      End
      Begin VB.Menu Properties 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "FormManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    With Me
        .Slider1.Value = .DataGrid1.RowHeight
        .LabelPersen.Caption = "Zoom : " & Slider1.Value & " %"
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            .LabelPersen.ToolTipText = "Dobel-Klik untuk merubah value"
            .Slider1.ToolTipText = "Geser untuk merubah value atau klik tombol navigasi"
        Else
            .LabelPersen.ToolTipText = "Double-Click to changes value"
            .Slider1.ToolTipText = "Scroll to changes value or click the navigation button"
        End If
    End With
    menuTersembunyi.Visible = False
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
Sub AturDatabase()
        If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
            CN_FormUtama.CursorLocation = adUseClient
            CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
        If FORM_UTAMA.cmJejaringSosial.FontBold = True Then
            With FormManage
                .AdodcMain.ConnectionString = CN_FormUtama.ConnectionString
                .AdodcMain.RecordSource = "Select * From tbJejaringSosial Order by Nama_Jejaring asc;"
                .AdodcMain.Refresh
                Set .DataGrid1.DataSource = .AdodcMain
                .Caption = "Manage Data : @[" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmJejaringSosial.Caption & ")]"
            End With
        ElseIf FORM_UTAMA.cmElectronicMail.FontBold = True Then
            With FormManage
                .AdodcMain.ConnectionString = CN_FormUtama.ConnectionString
                .AdodcMain.RecordSource = "Select * From tbElectronicMail Order by Nama_Server asc;"
                .AdodcMain.Refresh
                Set .DataGrid1.DataSource = .AdodcMain
                .Caption = "Manage Data : @[" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmElectronicMail.Caption & ")]"
            End With
        ElseIf FORM_UTAMA.cmForumInternet.FontBold = True Then
            With FormManage
                .AdodcMain.ConnectionString = CN_FormUtama.ConnectionString
                .AdodcMain.RecordSource = "Select * From tbForumInternet Order by Nama_Forum asc;"
                .AdodcMain.Refresh
                Set .DataGrid1.DataSource = .AdodcMain
                .Caption = "Manage Data : @[" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmForumInternet.Caption & ")]"
            End With
        ElseIf FORM_UTAMA.cmFTP.FontBold = True Then
            With FormManage
                .AdodcMain.ConnectionString = CN_FormUtama.ConnectionString
                .AdodcMain.RecordSource = "Select * From tbFTP Order by Nama_Host asc;"
                .AdodcMain.Refresh
                Set .DataGrid1.DataSource = .AdodcMain
                .Caption = "Manage Data : @[" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmFTP.Caption & ")]"
            End With
        ElseIf FORM_UTAMA.cmBlogging.FontBold = True Then
            With FormManage
                .AdodcMain.ConnectionString = CN_FormUtama.ConnectionString
                .AdodcMain.RecordSource = "Select * From tbBlogging Order by Nama_Penyedia_Blog asc;"
                .AdodcMain.Refresh
                Set .DataGrid1.DataSource = .AdodcMain
                .Caption = "Manage Data : @[" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmBlogging.Caption & ")]"
            End With
        ElseIf FORM_UTAMA.cmIdentitasPribadi.FontBold = True Then
            With FormManage
                .AdodcMain.ConnectionString = CN_FormUtama.ConnectionString
                .AdodcMain.RecordSource = "Select * From tbIdentitasPribadi Order by Nama_Lengkap asc;"
                .AdodcMain.Refresh
                Set .DataGrid1.DataSource = .AdodcMain
                .Caption = "Manage Data : @[" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmIdentitasPribadi.Caption & ")]"
            End With
        ElseIf FORM_UTAMA.cmBukuAlamat.FontBold = True Then
            With FormManage
                .AdodcMain.ConnectionString = CN_FormUtama.ConnectionString
                .AdodcMain.RecordSource = "Select * From tbBukuAlamat Order by Nama_Kontak asc;"
                .AdodcMain.Refresh
                Set .DataGrid1.DataSource = .AdodcMain
                .Caption = "Manage Data : @[" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmBukuAlamat.Caption & ")]"
            End With
        ElseIf FORM_UTAMA.cmUlangTahun.FontBold = True Then
            With FormManage
                .AdodcMain.ConnectionString = CN_FormUtama.ConnectionString
                .AdodcMain.RecordSource = "Select * From tbUlangTahun Order by Nama asc;"
                .AdodcMain.Refresh
                Set .DataGrid1.DataSource = .AdodcMain
                .Caption = "Manage Data : @[" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmUlangTahun.Caption & ")]"
            End With
        ElseIf FORM_UTAMA.cmAgenda.FontBold = True Then
            With FormManage
                .AdodcMain.ConnectionString = CN_FormUtama.ConnectionString
                .AdodcMain.RecordSource = "Select * From tbAgenda Order by Nama_Agenda asc;"
                .AdodcMain.Refresh
                Set .DataGrid1.DataSource = .AdodcMain
                .Caption = "Manage Data : @[" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmAgenda.Caption & ")]"
            End With
        ElseIf FORM_UTAMA.cmRegistrasiSoftware.FontBold = True Then
            With FormManage
                .AdodcMain.ConnectionString = CN_FormUtama.ConnectionString
                .AdodcMain.RecordSource = "Select * From tbRegistrasiSoftware Order by Nama_Software asc;"
                .AdodcMain.Refresh
                Set .DataGrid1.DataSource = .AdodcMain
                .Caption = "Manage Data : @[" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmRegistrasiSoftware.Caption & ")]"
            End With
        End If
    If FormPengaturan.cekKunciTabel.Value = Checked Then
        Me.DataGrid1.AllowUpdate = False
    Else
        Me.DataGrid1.AllowUpdate = True
    End If
End Sub

Private Sub cmBantuan_Click()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Kalimat = App.Path & "\bantuan\html\Manage.html"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Kalimat = App.Path & "\bantuan\html\Manage1.html"
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

Private Sub cmEdit_Click()
If AdodcMain.Recordset.RecordCount = 0 Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Maaf, tidak ada data yang dapat diedit.", vbExclamation + vbOKOnly, ""
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Sorry, there is no data that can be edited.", vbExclamation + vbOKOnly, ""
    End If
Else
    If FORM_UTAMA.cmJejaringSosial.FontBold = True Then
        For Each Objek In Form_JEJARING_SOSIAL
            If TypeName(Objek) = "XPText" Then
                With Objek
                    .Locked = False
                    .Text = ""
                    .ForeColor = vbBlack
                End With
            End If
        Next
        With Form_JEJARING_SOSIAL
            .Caption = "Edit Data - " & FORM_UTAMA.cmJejaringSosial.Caption & " [@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmJejaringSosial.Caption & ")]"
            .textJejaringSosial(0).Text = AdodcMain.Recordset.Fields(0).Value
            .textJejaringSosial(1).Text = AdodcMain.Recordset.Fields(1).Value
            .textJejaringSosial(2).Text = AdodcMain.Recordset.Fields(2).Value
            .textJejaringSosial(3).Text = AdodcMain.Recordset.Fields(3).Value
            .textJejaringSosial(4).Text = AdodcMain.Recordset.Fields(4).Value
            .textJejaringSosial(5).Text = AdodcMain.Recordset.Fields(5).Value
            .textJejaringSosial(6).Text = AdodcMain.Recordset.Fields(6).Value
            .textJejaringSosial(7).Text = AdodcMain.Recordset.Fields(7).Value
            'tambahkan kode lagi
            If FormPengaturan.cmbBahasa.ListIndex = 1 Then
                .cmSimpan.Caption = "&Update"
            Else
                .cmSimpan.Caption = "&Perbarui"
            End If
            .Show vbModal, Me
        End With
    ElseIf FORM_UTAMA.cmElectronicMail.FontBold = True Then
        For Each Objek In Form_ELECTRONIC_MAIL
            If TypeName(Objek) = "XPText" Then
                With Objek
                    .Locked = False
                    .Text = ""
                    .ForeColor = vbBlack
                End With
            End If
        Next
        With Form_ELECTRONIC_MAIL
            .Caption = "Edit Data - " & FORM_UTAMA.cmElectronicMail.Caption & " [@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmElectronicMail.Caption & ")]"
            .textNamaServer.Text = AdodcMain.Recordset.Fields(0).Value
            .textNamaPengguna.Text = AdodcMain.Recordset.Fields(1).Value
            .textAlamatEmail.Text = AdodcMain.Recordset.Fields(2).Value
            .textPassword.Text = AdodcMain.Recordset.Fields(3).Value
            .textPertanyaanRahasia.Text = AdodcMain.Recordset.Fields(4).Value
            .textJawabanPertanyaan.Text = AdodcMain.Recordset.Fields(5).Value
            .TextURL.Text = AdodcMain.Recordset.Fields(6).Value
            .textPemilikAkun.Text = AdodcMain.Recordset.Fields(7).Value
            .textTanggal.Text = AdodcMain.Recordset.Fields(8).Value
            .textKeterangan.Text = AdodcMain.Recordset.Fields(9).Value
            'tambahkan kode lagi
            If FormPengaturan.cmbBahasa.ListIndex = 1 Then
                .cmSimpan.Caption = "&Update"
            Else
                .cmSimpan.Caption = "&Perbarui"
            End If
            .Show vbModal, Me
        End With
    ElseIf FORM_UTAMA.cmForumInternet.FontBold = True Then
        For Each Objek In Form_FORUM_INTERNET
            If TypeName(Objek) = "XPText" Then
                With Objek
                    .Locked = False
                    .Text = ""
                    .ForeColor = vbBlack
                End With
            End If
        Next
        With Form_FORUM_INTERNET
            .Caption = "Edit Data - " & FORM_UTAMA.cmForumInternet.Caption & " [@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmForumInternet.Caption & ")]"
            .TextNamaForum.Text = AdodcMain.Recordset.Fields(0).Value
            .textNamaPengguna.Text = AdodcMain.Recordset.Fields(1).Value
            .textAlamatEmail.Text = AdodcMain.Recordset.Fields(2).Value
            .textPassword.Text = AdodcMain.Recordset.Fields(3).Value
            .TextPosisi.Text = AdodcMain.Recordset.Fields(4).Value
            .TextNickName.Text = AdodcMain.Recordset.Fields(5).Value
            .TextURL.Text = AdodcMain.Recordset.Fields(6).Value
            .textTanggal.Text = AdodcMain.Recordset.Fields(7).Value
            .textKeterangan.Text = AdodcMain.Recordset.Fields(8).Value
            
            If FormPengaturan.cmbBahasa.ListIndex = 1 Then
                .cmSimpan.Caption = "&Update"
            Else
                .cmSimpan.Caption = "&Perbarui"
            End If
            .Show vbModal, Me
        End With
    ElseIf FORM_UTAMA.cmFTP.FontBold = True Then
        For Each Objek In Form_FTP
            If TypeName(Objek) = "XPText" Then
                With Objek
                    .Locked = False
                    .Text = ""
                    .ForeColor = vbBlack
                End With
            End If
        Next
        With Form_FTP
            .Caption = "Edit Data - " & FORM_UTAMA.cmFTP.Caption & " [@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmFTP.Caption & ")]"
            .textNamaHost.Text = AdodcMain.Recordset.Fields(0).Value
            .textPort.Text = AdodcMain.Recordset.Fields(1).Value
            .textNamaServer.Text = AdodcMain.Recordset.Fields(2).Value
            .textNamaPengguna.Text = AdodcMain.Recordset.Fields(3).Value
            .textAlamatEmail.Text = AdodcMain.Recordset.Fields(4).Value
            .textPassword.Text = AdodcMain.Recordset.Fields(5).Value
            .textTanggal.Text = AdodcMain.Recordset.Fields(6).Value
            .textKeterangan.Text = AdodcMain.Recordset.Fields(7).Value
            
            If FormPengaturan.cmbBahasa.ListIndex = 1 Then
                .cmSimpan.Caption = "&Update"
            Else
                .cmSimpan.Caption = "&Perbarui"
            End If
            .Show vbModal, Me
        End With
    ElseIf FORM_UTAMA.cmBlogging.FontBold = True Then
        For Each Objek In Form_BLOGGING_WEBSITE
            If TypeName(Objek) = "XPText" Then
                With Objek
                    .Locked = False
                    .Text = ""
                    .ForeColor = vbBlack
                End With
            End If
        Next
        With Form_BLOGGING_WEBSITE
            .Caption = "Edit Data - " & FORM_UTAMA.cmBlogging.Caption & " [@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmBlogging.Caption & ")]"
            .textNamaPenyediaBlog.Text = AdodcMain.Recordset.Fields(0).Value
            .textNamaPengguna.Text = AdodcMain.Recordset.Fields(1).Value
            .textAlamatEmail.Text = AdodcMain.Recordset.Fields(2).Value
            .textPassword.Text = AdodcMain.Recordset.Fields(3).Value
            .TextURL.Text = AdodcMain.Recordset.Fields(4).Value
            .textTanggal.Text = AdodcMain.Recordset.Fields(5).Value
            .textKeterangan.Text = AdodcMain.Recordset.Fields(6).Value
            
            If FormPengaturan.cmbBahasa.ListIndex = 1 Then
                .cmSimpan.Caption = "&Update"
            Else
                .cmSimpan.Caption = "&Perbarui"
            End If
            .Show vbModal, Me
        End With
    ElseIf FORM_UTAMA.cmRegistrasiSoftware.FontBold = True Then
        For Each Objek In Form_REGISTRASI_SOFTWARE
            If TypeName(Objek) = "XPText" Then
                With Objek
                    .Locked = False
                    .Text = ""
                    .ForeColor = vbBlack
                End With
            End If
        Next
        With Form_REGISTRASI_SOFTWARE
            .Caption = "Edit Data - " & FORM_UTAMA.cmRegistrasiSoftware.Caption & " [@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmRegistrasiSoftware.Caption & ")]"
            .textNamaSoftware.Text = AdodcMain.Recordset.Fields(0).Value
            .cmbKategori.Text = AdodcMain.Recordset.Fields(1).Value
            .TextDeveloper.Text = AdodcMain.Recordset.Fields(2).Value
            .TextUsername.Text = AdodcMain.Recordset.Fields(3).Value
            .TextSerialKey.Text = AdodcMain.Recordset.Fields(4).Value
            .cmbJenisLisensi.Text = AdodcMain.Recordset.Fields(5).Value
            .textKeterangan.Text = AdodcMain.Recordset.Fields(6).Value
            
            If FormPengaturan.cmbBahasa.ListIndex = 1 Then
                .cmSimpan.Caption = "&Update"
            Else
                .cmSimpan.Caption = "&Perbarui"
            End If
            .Show vbModal, Me
        End With
    ElseIf FORM_UTAMA.cmAgenda.FontBold = True Then
        For Each Objek In Form_AGENDA
            If TypeName(Objek) = "XPText" Then
                With Objek
                    .Locked = False
                    .Text = ""
                    .ForeColor = vbBlack
                End With
            End If
        Next
        With Form_AGENDA
            .Caption = "Edit Data - " & FORM_UTAMA.cmAgenda.Caption & " [@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmAgenda.Caption & ")]"
            .textKodeAgenda.Text = AdodcMain.Recordset.Fields(0).Value
            .textNamaAgenda.Text = AdodcMain.Recordset.Fields(1).Value
            .textTema.Text = AdodcMain.Recordset.Fields(2).Value
            .textTanggal.Text = AdodcMain.Recordset.Fields(3).Value
            .textWaktuMulai.Text = AdodcMain.Recordset.Fields(4).Value
            .textWaktuAkhir.Text = AdodcMain.Recordset.Fields(5).Value
            .textTempat.Text = AdodcMain.Recordset.Fields(6).Value
            .textKeterangan.Text = AdodcMain.Recordset.Fields(7).Value
            
            If FormPengaturan.cmbBahasa.ListIndex = 1 Then
                .cmSimpan.Caption = "&Update"
            Else
                .cmSimpan.Caption = "&Perbarui"
            End If
            .Show vbModal, Me
        End With
    ElseIf FORM_UTAMA.cmUlangTahun.FontBold = True Then
        For Each Objek In Form_ULANG_TAHUN
            If TypeName(Objek) = "XPText" Then
                With Objek
                    .Locked = False
                    .Text = ""
                    .ForeColor = vbBlack
                End With
            End If
        Next
        With Form_ULANG_TAHUN
            .Caption = "Edit Data - " & FORM_UTAMA.cmUlangTahun.Caption & " [@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmUlangTahun.Caption & ")]"
            .textNama.Text = AdodcMain.Recordset.Fields(0).Value
            .textTTL.Text = AdodcMain.Recordset.Fields(1).Value
            .textKeterangan.Text = AdodcMain.Recordset.Fields(2).Value
            
            If FormPengaturan.cmbBahasa.ListIndex = 1 Then
                .cmSimpan.Caption = "&Update"
            Else
                .cmSimpan.Caption = "&Perbarui"
            End If
            .Show vbModal, Me
        End With
    ElseIf FORM_UTAMA.cmBukuAlamat.FontBold = True Then
        For Each Objek In Form_BUKU_ALAMAT
            If TypeName(Objek) = "XPText" Then
                With Objek
                    .Locked = False
                    .Text = ""
                    .ForeColor = vbBlack
                End With
            End If
        Next
        With Form_BUKU_ALAMAT
            .Caption = "Edit Data - " & FORM_UTAMA.cmBukuAlamat.Caption & " [@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmBukuAlamat.Caption & ")]"
            .textNamaKontak.Text = AdodcMain.Recordset.Fields(0).Value
            .textNamaPanggilan.Text = AdodcMain.Recordset.Fields(1).Value
            .textNomorTeleponPribadi.Text = AdodcMain.Recordset.Fields(2).Value
            .textNomorTeleponRumah.Text = AdodcMain.Recordset.Fields(3).Value
            .textNomorTeleponKantor.Text = AdodcMain.Recordset.Fields(4).Value
            .textFax.Text = AdodcMain.Recordset.Fields(5).Value
            .textAlamatEmail.Text = AdodcMain.Recordset.Fields(6).Value
            .textWebsite.Text = AdodcMain.Recordset.Fields(7).Value
            .textZIPPostalCode.Text = AdodcMain.Recordset.Fields(8).Value
            .textAlamatRumah.Text = AdodcMain.Recordset.Fields(9).Value
            .textKeterangan.Text = AdodcMain.Recordset.Fields(10).Value
            
            If FormPengaturan.cmbBahasa.ListIndex = 1 Then
                .cmSimpan.Caption = "&Update"
            Else
                .cmSimpan.Caption = "&Perbarui"
            End If
            .Show vbModal, Me
        End With
    ElseIf FORM_UTAMA.cmIdentitasPribadi.FontBold = True Then
        For Each Objek In Form_IDENTITAS_PRIBADI
            If TypeName(Objek) = "XPText" Then
                With Objek
                    .Locked = False
                    .Text = ""
                    .ForeColor = vbBlack
                End With
            End If
        Next
        With Form_IDENTITAS_PRIBADI
            .Caption = "Edit Data - " & FORM_UTAMA.cmIdentitasPribadi.Caption & " [@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmIdentitasPribadi.Caption & ")]"
            .textNamaLengkap.Text = AdodcMain.Recordset.Fields(0).Value
            .textNamaPanggilan.Text = AdodcMain.Recordset.Fields(1).Value
            .textTempat.Text = AdodcMain.Recordset.Fields(2).Value
            .cmbTanggal.Text = AdodcMain.Recordset.Fields(3).Value
            .cmbBulan.Text = AdodcMain.Recordset.Fields(4).Value
            .cmbTahun.Text = AdodcMain.Recordset.Fields(5).Value
            .cmbJenisKelamin.Text = AdodcMain.Recordset.Fields(6).Value
            .cmbAgama.Text = AdodcMain.Recordset.Fields(7).Value
            .cmbGolonganDarah.Text = AdodcMain.Recordset.Fields(8).Value
            .textPekerjaan.Text = AdodcMain.Recordset.Fields(9).Value
            .textAlamatRumah.Text = AdodcMain.Recordset.Fields(10).Value
            .textAlamatEmail.Text = AdodcMain.Recordset.Fields(11).Value
            .textAlamatWebsite.Text = AdodcMain.Recordset.Fields(12).Value
            .textNomorTelepon.Text = AdodcMain.Recordset.Fields(13).Value
            .textKotaAsal.Text = AdodcMain.Recordset.Fields(14).Value
            .textKotaSekarang.Text = AdodcMain.Recordset.Fields(15).Value
            .textKodePos.Text = AdodcMain.Recordset.Fields(16).Value
            .textProvinsi.Text = AdodcMain.Recordset.Fields(17).Value
            .textKewargaNegaraan.Text = AdodcMain.Recordset.Fields(18).Value
            .textStatusPendidikan.Text = AdodcMain.Recordset.Fields(19).Value
            .cmbStatusHubungan.Text = AdodcMain.Recordset.Fields(20).Value
            .textHobby.Text = AdodcMain.Recordset.Fields(21).Value
            .textKeterangan.Text = AdodcMain.Recordset.Fields(22).Value
            
            If FormPengaturan.cmbBahasa.ListIndex = 1 Then
                .cmSimpan.Caption = "&Update"
            Else
                .cmSimpan.Caption = "&Perbarui"
            End If
            .Show vbModal, Me
        End With
    End If
End If
End Sub

Private Sub cmFilter_Click()
If AdodcMain.Recordset.RecordCount = 0 Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Maaf, tidak ada data yang dapat difilter", vbExclamation + vbOKOnly, ""
    Else
        MsgBox "Sory, no data to filtered", vbExclamation + vbOKOnly, ""
    End If
Else
    FormFilter.Show vbModal, Me
End If
End Sub

Private Sub cmHapus_Click()
On Error GoTo PecahkanError
If FORM_UTAMA.cmJejaringSosial.FontBold = True Then
    If AdodcMain.Recordset.RecordCount = 0 Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf, tidak ada data yang dapat dihapus!", vbExclamation + vbOKOnly, "Info"
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry, there is no data that can be deleted!", vbExclamation + vbOKOnly, "Info"
        End If
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini?" & vbCrLf & vbCrLf & _
                    "===KETERANGAN===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "Perintah ini tidak dapat dibatalkan. Yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus data dengan '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            Pesan = MsgBox("Are you sure want to delete this data?" & vbCrLf & vbCrLf & _
                    "===DESCRIPTION===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "This command can't be undone. Sure you want to delete this data?", vbQuestion + vbYesNo, "Delete data with '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        End If
            If Pesan = vbYes Then
                With AdodcMain
                    .Recordset.Delete
                    Set DataGrid1.DataSource = AdodcMain
                    .Refresh
                End With
                FORM_UTAMA.cmJejaringSosial_Click
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Berhasil Dihapus"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data pada akun '" & FORM_UTAMA.cmJejaringSosial.Caption & "' telah dihapus pada " & Time
                        .ToolTipText = .Text
                    End With
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    MsgBox "Data Successfully Removed", vbInformation + vbOKOnly, "Successfully Removed"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data on '" & FORM_UTAMA.cmJejaringSosial.Caption & "' accounts was removed at " & Time
                        .ToolTipText = .Text
                    End With
                End If
            End If
    End If
ElseIf FORM_UTAMA.cmElectronicMail.FontBold = True Then
    If AdodcMain.Recordset.RecordCount = 0 Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf, tidak ada data yang dapat dihapus!", vbExclamation + vbOKOnly, "Info"
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry, there is no data that can be deleted!", vbExclamation + vbOKOnly, "Info"
        End If
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini?" & vbCrLf & vbCrLf & _
                    "===KETERANGAN===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(8).Name & " : " & AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(9).Name & " : " & AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "Perintah ini tidak dapat dibatalkan. Yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus data dengan '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            Pesan = MsgBox("Are you sure want to delete this data?" & vbCrLf & vbCrLf & _
                    "===DESCRIPTION===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(8).Name & " : " & AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(9).Name & " : " & AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "This command can't be undone. Sure you want to delete this data?", vbQuestion + vbYesNo, "Delete data with '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        End If
            If Pesan = vbYes Then
                With AdodcMain
                    .Recordset.Delete
                    Set DataGrid1.DataSource = AdodcMain
                    .Refresh
                End With
                FORM_UTAMA.cmElectronicMail_Click
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Berhasil Dihapus"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data pada akun '" & FORM_UTAMA.cmElectronicMail.Caption & "' telah dihapus pada " & Time
                        .ToolTipText = .Text
                    End With
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    MsgBox "Data Successfully Removed", vbInformation + vbOKOnly, "Successfully Removed"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data on '" & FORM_UTAMA.cmElectronicMail.Caption & "' accounts was removed at " & Time
                        .ToolTipText = .Text
                    End With
                End If
            End If
    End If
ElseIf FORM_UTAMA.cmForumInternet.FontBold = True Then
    If AdodcMain.Recordset.RecordCount = 0 Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf, tidak ada data yang dapat dihapus!", vbExclamation + vbOKOnly, "Info"
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry, there is no data that can be deleted!", vbExclamation + vbOKOnly, "Info"
        End If
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini?" & vbCrLf & vbCrLf & _
                    "===KETERANGAN===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(8).Name & " : " & AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "Perintah ini tidak dapat dibatalkan. Yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus data dengan '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            Pesan = MsgBox("Are you sure want to delete this data?" & vbCrLf & vbCrLf & _
                    "===DESCRIPTION===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(8).Name & " : " & AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "This command can't be undone. Sure you want to delete this data?", vbQuestion + vbYesNo, "Delete data with '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        End If
            If Pesan = vbYes Then
                With AdodcMain
                    .Recordset.Delete
                    Set DataGrid1.DataSource = AdodcMain
                    .Refresh
                End With
                FORM_UTAMA.cmForumInternet_Click
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Berhasil Dihapus"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data pada akun '" & FORM_UTAMA.cmForumInternet.Caption & "' telah dihapus pada " & Time
                        .ToolTipText = .Text
                    End With
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    MsgBox "Data Successfully Removed", vbInformation + vbOKOnly, "Successfully Removed"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data on '" & FORM_UTAMA.cmForumInternet.Caption & "' accounts was removed at " & Time
                        .ToolTipText = .Text
                    End With
                End If
            End If
    End If
ElseIf FORM_UTAMA.cmFTP.FontBold = True Then
    If AdodcMain.Recordset.RecordCount = 0 Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf, tidak ada data yang dapat dihapus!", vbExclamation + vbOKOnly, "Info"
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry, there is no data that can be deleted!", vbExclamation + vbOKOnly, "Info"
        End If
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini?" & vbCrLf & vbCrLf & _
                    "===KETERANGAN===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "Perintah ini tidak dapat dibatalkan. Yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus data dengan '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            Pesan = MsgBox("Are you sure want to delete this data?" & vbCrLf & vbCrLf & _
                    "===DESCRIPTION===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "This command can't be undone. Sure you want to delete this data?", vbQuestion + vbYesNo, "Delete data with '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        End If
            If Pesan = vbYes Then
                With AdodcMain
                    .Recordset.Delete
                    Set DataGrid1.DataSource = AdodcMain
                    .Refresh
                End With
                FORM_UTAMA.cmFTP_Click
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Berhasil Dihapus"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data pada akun '" & FORM_UTAMA.cmFTP.Caption & "' telah dihapus pada " & Time
                        .ToolTipText = .Text
                    End With
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    MsgBox "Data Successfully Removed", vbInformation + vbOKOnly, "Successfully Removed"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data on '" & FORM_UTAMA.cmFTP.Caption & "' accounts was removed at " & Time
                        .ToolTipText = .Text
                    End With
                End If
            End If
    End If
ElseIf FORM_UTAMA.cmBlogging.FontBold = True Then
    If AdodcMain.Recordset.RecordCount = 0 Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf, tidak ada data yang dapat dihapus!", vbExclamation + vbOKOnly, "Info"
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry, there is no data that can be deleted!", vbExclamation + vbOKOnly, "Info"
        End If
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini?" & vbCrLf & vbCrLf & _
                    "===KETERANGAN===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "Perintah ini tidak dapat dibatalkan. Yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus data dengan '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            Pesan = MsgBox("Are you sure want to delete this data?" & vbCrLf & vbCrLf & _
                    "===DESCRIPTION===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "This command can't be undone. Sure you want to delete this data?", vbQuestion + vbYesNo, "Delete data with '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        End If
            If Pesan = vbYes Then
                With AdodcMain
                    .Recordset.Delete
                    Set DataGrid1.DataSource = AdodcMain
                    .Refresh
                End With
                FORM_UTAMA.cmBlogging_Click
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Berhasil Dihapus"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data pada akun '" & FORM_UTAMA.cmBlogging.Caption & "' telah dihapus pada " & Time
                        .ToolTipText = .Text
                    End With
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    MsgBox "Data Successfully Removed", vbInformation + vbOKOnly, "Successfully Removed"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data on '" & FORM_UTAMA.cmBlogging.Caption & "' accounts was removed at " & Time
                        .ToolTipText = .Text
                    End With
                End If
            End If
    End If
ElseIf FORM_UTAMA.cmRegistrasiSoftware.FontBold = True Then
    If AdodcMain.Recordset.RecordCount = 0 Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf, tidak ada data yang dapat dihapus!", vbExclamation + vbOKOnly, "Info"
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry, there is no data that can be deleted!", vbExclamation + vbOKOnly, "Info"
        End If
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini?" & vbCrLf & vbCrLf & _
                    "===KETERANGAN===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "Perintah ini tidak dapat dibatalkan. Yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus data dengan '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            Pesan = MsgBox("Are you sure want to delete this data?" & vbCrLf & vbCrLf & _
                    "===DESCRIPTION===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "This command can't be undone. Sure you want to delete this data?", vbQuestion + vbYesNo, "Delete data with '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        End If
            If Pesan = vbYes Then
                With AdodcMain
                    .Recordset.Delete
                    Set DataGrid1.DataSource = AdodcMain
                    .Refresh
                End With
                FORM_UTAMA.cmRegistrasiSoftware_Click
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Berhasil Dihapus"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data pada akun '" & FORM_UTAMA.cmRegistrasiSoftware.Caption & "' telah dihapus pada " & Time
                        .ToolTipText = .Text
                    End With
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    MsgBox "Data Successfully Removed", vbInformation + vbOKOnly, "Successfully Removed"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data on '" & FORM_UTAMA.cmRegistrasiSoftware.Caption & "' accounts was removed at " & Time
                        .ToolTipText = .Text
                    End With
                End If
            End If
    End If
ElseIf FORM_UTAMA.cmAgenda.FontBold = True Then
    If AdodcMain.Recordset.RecordCount = 0 Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf, tidak ada data yang dapat dihapus!", vbExclamation + vbOKOnly, "Info"
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry, there is no data that can be deleted!", vbExclamation + vbOKOnly, "Info"
        End If
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini?" & vbCrLf & vbCrLf & _
                    "===KETERANGAN===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "Perintah ini tidak dapat dibatalkan. Yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus data dengan '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            Pesan = MsgBox("Are you sure want to delete this data?" & vbCrLf & vbCrLf & _
                    "===DESCRIPTION===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "This command can't be undone. Sure you want to delete this data?", vbQuestion + vbYesNo, "Delete data with '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        End If
            If Pesan = vbYes Then
                With AdodcMain
                    .Recordset.Delete
                    Set DataGrid1.DataSource = AdodcMain
                    .Refresh
                End With
                FORM_UTAMA.cmAgenda_Click
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Berhasil Dihapus"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data pada akun '" & FORM_UTAMA.cmAgenda.Caption & "' telah dihapus pada " & Time
                        .ToolTipText = .Text
                    End With
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    MsgBox "Data Successfully Removed", vbInformation + vbOKOnly, "Successfully Removed"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data on '" & FORM_UTAMA.cmAgenda.Caption & "' accounts was removed at " & Time
                        .ToolTipText = .Text
                    End With
                End If
            End If
    End If
ElseIf FORM_UTAMA.cmUlangTahun.FontBold = True Then
    If AdodcMain.Recordset.RecordCount = 0 Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf, tidak ada data yang dapat dihapus!", vbExclamation + vbOKOnly, "Info"
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry, there is no data that can be deleted!", vbExclamation + vbOKOnly, "Info"
        End If
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini?" & vbCrLf & vbCrLf & _
                    "===KETERANGAN===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "Perintah ini tidak dapat dibatalkan. Yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus data dengan '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            Pesan = MsgBox("Are you sure want to delete this data?" & vbCrLf & vbCrLf & _
                    "===DESCRIPTION===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "This command can't be undone. Sure you want to delete this data?", vbQuestion + vbYesNo, "Delete data with '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        End If
            If Pesan = vbYes Then
                With AdodcMain
                    .Recordset.Delete
                    Set DataGrid1.DataSource = AdodcMain
                    .Refresh
                End With
                FORM_UTAMA.cmUlangTahun_Click
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Berhasil Dihapus"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data pada akun '" & FORM_UTAMA.cmUlangTahun.Caption & "' telah dihapus pada " & Time
                        .ToolTipText = .Text
                    End With
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    MsgBox "Data Successfully Removed", vbInformation + vbOKOnly, "Successfully Removed"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data on '" & FORM_UTAMA.cmUlangTahun.Caption & "' accounts was removed at " & Time
                        .ToolTipText = .Text
                    End With
                End If
            End If
    End If
ElseIf FORM_UTAMA.cmBukuAlamat.FontBold = True Then
    If AdodcMain.Recordset.RecordCount = 0 Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf, tidak ada data yang dapat dihapus!", vbExclamation + vbOKOnly, "Info"
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry, there is no data that can be deleted!", vbExclamation + vbOKOnly, "Info"
        End If
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini?" & vbCrLf & vbCrLf & _
                    "===KETERANGAN===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(8).Name & " : " & AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(9).Name & " : " & AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(10).Name & " : " & AdodcMain.Recordset.Fields(10).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "Perintah ini tidak dapat dibatalkan. Yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus data dengan '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            Pesan = MsgBox("Are you sure want to delete this data?" & vbCrLf & vbCrLf & _
                    "===DESCRIPTION===" & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(8).Name & " : " & AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(9).Name & " : " & AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(10).Name & " : " & AdodcMain.Recordset.Fields(10).Value & vbCrLf & _
                    "================" & vbCrLf & vbCrLf & _
                    "This command can't be undone. Sure you want to delete this data?", vbQuestion + vbYesNo, "Delete data with '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        End If
            If Pesan = vbYes Then
                With AdodcMain
                    .Recordset.Delete
                    Set DataGrid1.DataSource = AdodcMain
                    .Refresh
                End With
                FORM_UTAMA.cmBukuAlamat_Click
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Berhasil Dihapus"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data pada akun '" & FORM_UTAMA.cmBukuAlamat.Caption & "' telah dihapus pada " & Time
                        .ToolTipText = .Text
                    End With
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    MsgBox "Data Successfully Removed", vbInformation + vbOKOnly, "Successfully Removed"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data on '" & FORM_UTAMA.cmBukuAlamat.Caption & "' accounts was removed at " & Time
                        .ToolTipText = .Text
                    End With
                End If
            End If
    End If
ElseIf FORM_UTAMA.cmIdentitasPribadi.FontBold = True Then
    If AdodcMain.Recordset.RecordCount = 0 Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Maaf, tidak ada data yang dapat dihapus!", vbExclamation + vbOKOnly, "Info"
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Sorry, there is no data that can be deleted!", vbExclamation + vbOKOnly, "Info"
        End If
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini?" & vbCrLf & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(8).Name & " : " & AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(9).Name & " : " & AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(10).Name & " : " & AdodcMain.Recordset.Fields(10).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(11).Name & " : " & AdodcMain.Recordset.Fields(11).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(12).Name & " : " & AdodcMain.Recordset.Fields(12).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(13).Name & " : " & AdodcMain.Recordset.Fields(13).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(14).Name & " : " & AdodcMain.Recordset.Fields(14).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(15).Name & " : " & AdodcMain.Recordset.Fields(15).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(16).Name & " : " & AdodcMain.Recordset.Fields(16).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(17).Name & " : " & AdodcMain.Recordset.Fields(17).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(18).Name & " : " & AdodcMain.Recordset.Fields(18).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(19).Name & " : " & AdodcMain.Recordset.Fields(19).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(20).Name & " : " & AdodcMain.Recordset.Fields(20).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(21).Name & " : " & AdodcMain.Recordset.Fields(21).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(22).Name & " : " & AdodcMain.Recordset.Fields(22).Value & vbCrLf & _
                    "Perintah ini tidak dapat dibatalkan. Yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus data dengan '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            Pesan = MsgBox("Are you sure want to delete this data?" & vbCrLf & vbCrLf & _
                    AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(8).Name & " : " & AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(9).Name & " : " & AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(10).Name & " : " & AdodcMain.Recordset.Fields(10).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(11).Name & " : " & AdodcMain.Recordset.Fields(11).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(12).Name & " : " & AdodcMain.Recordset.Fields(12).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(13).Name & " : " & AdodcMain.Recordset.Fields(13).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(14).Name & " : " & AdodcMain.Recordset.Fields(14).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(15).Name & " : " & AdodcMain.Recordset.Fields(15).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(16).Name & " : " & AdodcMain.Recordset.Fields(16).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(17).Name & " : " & AdodcMain.Recordset.Fields(17).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(18).Name & " : " & AdodcMain.Recordset.Fields(18).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(19).Name & " : " & AdodcMain.Recordset.Fields(19).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(20).Name & " : " & AdodcMain.Recordset.Fields(20).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(21).Name & " : " & AdodcMain.Recordset.Fields(21).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(22).Name & " : " & AdodcMain.Recordset.Fields(22).Value & vbCrLf & _
                    "This command can't be undone. Sure you want to delete this data?", vbQuestion + vbYesNo, "Delete data with '" & AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value & "' ?")
        End If
            If Pesan = vbYes Then
                With AdodcMain
                    .Recordset.Delete
                    Set DataGrid1.DataSource = AdodcMain
                    .Refresh
                End With
                FORM_UTAMA.cmIdentitasPribadi_Click
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    MsgBox "Data Berhasil Dihapus", vbInformation + vbOKOnly, "Berhasil Dihapus"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data pada akun '" & FORM_UTAMA.cmIdentitasPribadi.Caption & "' telah dihapus pada " & Time
                        .ToolTipText = .Text
                    End With
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    MsgBox "Data Successfully Removed", vbInformation + vbOKOnly, "Successfully Removed"
                    With FORM_UTAMA.StatusBawah.Panels.Item(1)
                        .Text = "Data on '" & FORM_UTAMA.cmIdentitasPribadi.Caption & "' accounts was removed at " & Time
                        .ToolTipText = .Text
                    End With
                End If
            End If
    End If
End If
Exit Sub
PecahkanError:
    PusatError
    AdodcMain.Refresh
End Sub

Private Sub cmKurang_Click()
    Slider1.Value = Slider1.Value - 1
    Slider1_Scroll
End Sub

Private Sub cmCari_Click()
    With FORM_UTAMA
        If .cmIdentitasPribadi.FontBold = True Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                FormCari.Caption = "Cari Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmIdentitasPribadi.Caption & "))"
            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                FormCari.Caption = "Search Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmIdentitasPribadi.Caption & "))"
            End If
        ElseIf .cmBukuAlamat.FontBold = True Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                FormCari.Caption = "Cari Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmBukuAlamat.Caption & "))"
            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                FormCari.Caption = "Search Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmBukuAlamat.Caption & "))"
            End If
        ElseIf .cmUlangTahun.FontBold = True Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                FormCari.Caption = "Cari Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmUlangTahun.Caption & "))"
            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                FormCari.Caption = "Search Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmUlangTahun.Caption & "))"
            End If
        ElseIf .cmAgenda.FontBold = True Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                FormCari.Caption = "Cari Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmAgenda.Caption & "))"
            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                FormCari.Caption = "Search Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmAgenda.Caption & "))"
            End If
        ElseIf .cmRegistrasiSoftware.FontBold = True Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                FormCari.Caption = "Cari Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmRegistrasiSoftware.Caption & "))"
            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                FormCari.Caption = "Search Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmRegistrasiSoftware.Caption & "))"
            End If
        ElseIf .cmJejaringSosial.FontBold = True Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                FormCari.Caption = "Cari Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmJejaringSosial.Caption & "))"
            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                FormCari.Caption = "Search Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmJejaringSosial.Caption & "))"
            End If
        ElseIf .cmElectronicMail.FontBold = True Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                FormCari.Caption = "Cari Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmElectronicMail.Caption & "))"
            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                FormCari.Caption = "Search Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmElectronicMail.Caption & "))"
            End If
        ElseIf .cmForumInternet.FontBold = True Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                FormCari.Caption = "Cari Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmForumInternet.Caption & "))"
            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                FormCari.Caption = "Search Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmForumInternet.Caption & "))"
            End If
        ElseIf .cmFTP.FontBold = True Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                FormCari.Caption = "Cari Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmFTP.Caption & "))"
            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                FormCari.Caption = "Search Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmFTP.Caption & "))"
            End If
        ElseIf .cmBlogging.FontBold = True Then
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                FormCari.Caption = "Cari Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmBlogging.Caption & "))"
            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                FormCari.Caption = "Search Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & .cmBlogging.Caption & "))"
            End If
        End If
    End With
    FormCari.Show vbModal, Me
End Sub

Private Sub cmSorot_Click()
If AdodcMain.Recordset.RecordCount = 0 Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Maaf, tidak ada data yang dapat disorot", vbExclamation + vbOKOnly, ""
    Else
        MsgBox "Sory, no data to sort", vbExclamation + vbOKOnly, ""
    End If
Else
    FormSorot.Show vbModal, Me
End If
End Sub

Private Sub cmTambah_Click()
    Slider1.Value = Slider1.Value + 1
    Slider1_Scroll
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu menuTersembunyi
End Sub

Private Sub Form_Load()
    AturKontrol
    AturDatabase
    PENGATURAN_WARNA
    PENGATURAN_BAHASA
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu menuTersembunyi
End Sub

Private Sub LabelPersen_DblClick()
    LabelPersen.Visible = False
    With textPersen
        .Visible = True
        .Text = Slider1.Value
        .Alignment = vbCenter
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            .ToolTipText = "Tekan ENTER untuk menerapkan value"
        Else
            .ToolTipText = "Press ENTER to apply the value"
        End If
        .SetFocus
    End With
End Sub

Private Sub menuBS_Click()
    AturDatabase
    With Me
        .cmEdit.Enabled = True
        .cmHapus.Enabled = True
        .cmSorot.Enabled = True
        .cmCari.Enabled = True
    End With
End Sub

Private Sub menuRefresh_Click()
    AturDatabase
    With Me
        .cmEdit.Enabled = True
        .cmHapus.Enabled = True
        .cmSorot.Enabled = True
        .cmCari.Enabled = True
    End With
End Sub

Private Sub Properties_Click()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Nama Pengguna : " & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & vbCrLf & _
                "Jumlah Data  : " & AdodcMain.Recordset.RecordCount & vbCrLf & _
                "Jumlah Cell  : " & Val(AdodcMain.Recordset.RecordCount) * Val(AdodcMain.Recordset.Fields.Count), vbInformation + vbOKOnly, "Propertis"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "User Name : " & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & vbCrLf & _
                "Data Count  : " & AdodcMain.Recordset.RecordCount & vbCrLf & _
                "Cell Count  : " & Val(AdodcMain.Recordset.RecordCount) * Val(AdodcMain.Recordset.Fields.Count), vbInformation + vbOKOnly, "Properties"
    End If
End Sub

Private Sub Slider1_Scroll()
        DataGrid1.RowHeight = Slider1.Value
        LabelPersen.Caption = "Zoom : " & Slider1.Value & " %"
End Sub

Private Sub textPersen_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, 44, 46
        Case Else
            If Val(textPersen.Text) <= 10 Then
                textPersen.Text = "10"
            ElseIf Val(textPersen.Text) >= 700 Then
                textPersen.Text = "700"
            Else
                With textPersen
                    .Visible = False
                    Slider1.Value = Val(textPersen.Text)
                    LabelPersen.Caption = "Zoom : " & Slider1.Value & " %"
                    LabelPersen.Visible = True
                    DataGrid1.RowHeight = Slider1.Value
                    Slider1.SetFocus
                End With
            End If
    End Select
End Sub

Sub PENGATURAN_WARNA()
    'PENGATURAN WARNA UNTUK FORM INI
    For Each Objek In Me
        Select Case FormPengaturan.cmbWarnaTampilan.ListIndex
        Case Is = 0 'Ungu Natural
            Me.BackColor = UnguNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguNatural
            If TypeName(Objek) = "AeroGroupBox" Then Objek.BackColor = UnguNatural
        Case Is = 1 'Merah
            Me.BackColor = Merah
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Merah
            If TypeName(Objek) = "AeroGroupBox" Then Objek.BackColor = Merah
        Case Is = 2 'Pink
            Me.BackColor = Pink
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Pink
            If TypeName(Objek) = "AeroGroupBox" Then Objek.BackColor = Pink
        Case Is = 3 'HijauMuda
            Me.BackColor = HijauMuda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = HijauMuda
            If TypeName(Objek) = "AeroGroupBox" Then Objek.BackColor = HijauMuda
        Case Is = 4 'Hitam
            Me.BackColor = Hitam
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Hitam
            If TypeName(Objek) = "AeroGroupBox" Then Objek.BackColor = Hitam
        Case Is = 5 'Silver
            Me.BackColor = Silver
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Silver
            If TypeName(Objek) = "AeroGroupBox" Then Objek.BackColor = Silver
        Case Is = 6 'SilverNatural
            Me.BackColor = SilverNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = SilverNatural
            If TypeName(Objek) = "AeroGroupBox" Then Objek.BackColor = SilverNatural
        Case Is = 7 'Orange
            Me.BackColor = Orange
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Orange
            If TypeName(Objek) = "AeroGroupBox" Then Objek.BackColor = Orange
        Case Is = 8 'UnguJanda
            Me.BackColor = UnguJanda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguJanda
            If TypeName(Objek) = "AeroGroupBox" Then Objek.BackColor = UnguJanda
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
                Objek.BackColor = &HBA9EA0
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

Sub PENGATURAN_BAHASA()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        cmEdit.Caption = "&Edit"
        cmHapus.Caption = "&Hapus"
        cmCari.Caption = "&Cari"
        cmTutup.Caption = "&Tutup"
        cmBantuan.Caption = "&Bantuan"
        cmSorot.Caption = "&Sorot"
        cmFilter.Caption = "&Filter"
        Me.menuRefresh.Caption = "Segarkan"
        Me.menuBS.Caption = "Bersihkan Sorot"
        Me.Properties.Caption = "Properties"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        cmEdit.Caption = "&Edit"
        cmHapus.Caption = "&Delete"
        cmCari.Caption = "&Search"
        cmTutup.Caption = "&Close"
        cmBantuan.Caption = "&Help"
        cmSorot.Caption = "&Sort"
        cmFilter.Caption = "&Filter"
        Me.menuRefresh.Caption = "Refresh"
        Me.menuBS.Caption = "Clear Highlight"
        Me.Properties.Caption = "Properties"
    End If
End Sub
