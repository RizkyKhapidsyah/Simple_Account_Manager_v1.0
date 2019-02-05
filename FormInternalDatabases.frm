VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormInternalDatabases 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internal Databases"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormInternalDatabases.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   330
      Left            =   0
      Top             =   3360
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
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DatagridUtama 
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
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
            LCID            =   1057
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
            LCID            =   1057
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
   Begin VB.ComboBox cmbNamaTabel 
      Height          =   390
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   4815
   End
   Begin Dacara_dcButton.dcButton cmBaru 
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Baru"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FormInternalDatabases.frx":27A2
      PicHot          =   "FormInternalDatabases.frx":2BF4
      PicNormal       =   "FormInternalDatabases.frx":3046
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmEdit 
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Edit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FormInternalDatabases.frx":3498
      PicHot          =   "FormInternalDatabases.frx":37B2
      PicNormal       =   "FormInternalDatabases.frx":3ACC
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmHapus 
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Hapus"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FormInternalDatabases.frx":3DE6
      PicHot          =   "FormInternalDatabases.frx":4930
      PicNormal       =   "FormInternalDatabases.frx":547A
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmPropertis 
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Propertis"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicDown         =   "FormInternalDatabases.frx":5FC4
      PicHot          =   "FormInternalDatabases.frx":D4C6
      PicNormal       =   "FormInternalDatabases.frx":149C8
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmTutup 
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Tutup"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   5
      PicDown         =   "FormInternalDatabases.frx":1BECA
      PicHot          =   "FormInternalDatabases.frx":1C31C
      PicNormal       =   "FormInternalDatabases.frx":1C76E
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1320
      TabIndex        =   6
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Tabel"
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "FormInternalDatabases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    NyambunggUtama
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With cmbNamaTabel
            .Clear
            .AddItem "Agama", 0
            .AddItem "Golongan Darah", 1
            .AddItem "Jenis Lisensi", 2
            .AddItem "Kategori Software", 3
            .AddItem "Pertanyaan Rahasia", 4
            .AddItem "Status Hubungan", 5
            .ListIndex = 0
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With cmbNamaTabel
            .Clear
            .AddItem "Religion", 0
            .AddItem "Blood Type", 1
            .AddItem "License", 2
            .AddItem "Software Categories", 3
            .AddItem "Security Questions", 4
            .AddItem "Relationship", 5
            .ListIndex = 0
        End With
    End If
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

Private Sub cmBaru_Click()
    With FormInputSetInternalDatabases
        Select Case cmbNamaTabel.ListIndex
            Case Is = 0
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Caption = "Agama"
                    .Label1.Caption = "Silahkan Isi Agama :"
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Caption = "New Religion"
                    .Label1.Caption = "Please input your religion :"
                End If
            Case Is = 1
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Caption = "Golongan Darah Baru"
                    .Label1.Caption = "Silahkan Isi Golongan Darah Baru :"
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Caption = "New Blood Type"
                    .Label1.Caption = "Please input new blood type :"
                End If
            Case Is = 2
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Caption = "Jenis Lisensi Baru"
                    .Label1.Caption = "Silahkan Isi Jenis Lisensi Baru :"
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Caption = "New License"
                    .Label1.Caption = "Please input new License of programs :"
                End If
            Case Is = 3
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Caption = "Kategori Software Baru"
                    .Label1.Caption = "Silahkan Isi Kategori Baru :"
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Caption = "New Software Categories"
                    .Label1.Caption = "Please input new Software Categories :"
                End If
            Case Is = 4
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Caption = "Pertanyaan Rahasia Baru"
                    .Label1.Caption = "Silahkan Isi Pertanyaan Rahasia Baru :"
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Caption = "New Security Questions"
                    .Label1.Caption = "Please input new Security Questions :"
                End If
            Case Is = 5
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Caption = "Status Hubungan Baru"
                    .Label1.Caption = "Silahkan Isi jenis Status Hubungan Baru :"
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Caption = "New Relationship"
                    .Label1.Caption = "Please input new Relationship :"
                End If
            End Select
            .LabelPenanda.Caption = "Baru"
            .textInputSetInternalDatabases.Text = ""
            .Show vbModal, Me
    End With
End Sub

Private Sub cmbNamaTabel_Click()
With AdodcUtama
    .ConnectionString = CN.ConnectionString
        Select Case cmbNamaTabel.ListIndex
        Case Is = 0
            .RecordSource = "Select * from tbAgama"
            With DatagridUtama
                .AllowUpdate = False
                .Columns.Item(0).Width = 20
            End With
        Case Is = 1
            .RecordSource = "Select * from tbGolonganDarah"
            With DatagridUtama
                .AllowUpdate = False
                .Columns.Item(0).Width = 2000
            End With
        Case Is = 2
            .RecordSource = "Select * from tbJenisLisensi"
            With DatagridUtama
                .AllowUpdate = False
                .Columns.Item(0).Width = 3000
            End With
        Case Is = 3
            .RecordSource = "Select * from tbKategoriSoftware"
            With DatagridUtama
                .AllowUpdate = False
                .Columns.Item(0).Width = 2000
            End With
        Case Is = 4
            .RecordSource = "Select * from tbPertanyaanRahasia"
            With DatagridUtama
                .AllowUpdate = False
                .Columns.Item(0).Width = 5000
            End With
        Case Is = 5
            .RecordSource = "Select * from tbStatusHubungan"
            With DatagridUtama
                .AllowUpdate = False
                .Columns.Item(0).Width = 3000
            End With
        End Select
    Set DatagridUtama.DataSource = AdodcUtama
    .Refresh
End With
End Sub

Private Sub cmEdit_Click()
    With FormInputSetInternalDatabases
        Select Case cmbNamaTabel.ListIndex
            Case Is = 0
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Caption = "Edit Agama"
                    .Label1.Caption = "Edit Agama ini menjadi :"
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Caption = "Edit Religion's Name"
                    .Label1.Caption = "Edit this religion's name as :"
                End If
            Case Is = 1
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Caption = "Edit Golongan Darah"
                    .Label1.Caption = "Edit golongan darah ini menjadi :"
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Caption = "Edit Blood Type"
                    .Label1.Caption = "Edit This blood type's name as :"
                End If
            Case Is = 2
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Caption = "Edit Jenis Lisensi"
                    .Label1.Caption = "Edit Jenis Lisensi ini menjadi :"
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Caption = "Edit Of License"
                    .Label1.Caption = "Edit of this license as:"
                End If
            Case Is = 3
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Caption = "Edit Kategori Software"
                    .Label1.Caption = "Edit Kategori ini menjadi :"
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Caption = "Edit Of Software Categories"
                    .Label1.Caption = "Edit this categories software's name as :"
                End If
            Case Is = 4
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Caption = "Edit Pertanyaan Rahasia"
                    .Label1.Caption = "Edit pertanyaan rahasia ini menjadi:"
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Caption = "Edit of Security Questions"
                    .Label1.Caption = "Edit of this security questions's name as : "
                End If
            Case Is = 5
                If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                    .Caption = "Edit Status Hubungan"
                    .Label1.Caption = "Edit nama status hubungan ini menjadi:"
                ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                    .Caption = "Edit of Relationship"
                    .Label1.Caption = "Edit of relationship's name as : "
                End If
            End Select
                If cmbNamaTabel.ListIndex = 0 Then
                    .textInputSetInternalDatabases.Text = AdodcUtama.Recordset.Fields(1).Value
                Else
                    .textInputSetInternalDatabases.Text = AdodcUtama.Recordset.Fields(0).Value
                End If
            .LabelPenanda.Caption = "Edit"
            .Show vbModal, Me
    End With
End Sub

Private Sub cmHapus_Click()
    If cmbNamaTabel.ListIndex = 0 Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini?" & vbCrLf & vbCrLf & _
                            "===========================================" & vbCrLf & _
                            "Keterangan" & vbCrLf & _
                            "===========================================" & vbCrLf & _
                            AdodcUtama.Recordset.Fields(0).Name & " : " & AdodcUtama.Recordset.Fields(0).Value & vbCrLf & _
                            AdodcUtama.Recordset.Fields(1).Name & " : " & AdodcUtama.Recordset.Fields(1).Value & vbCrLf & _
                            "===========================================", vbQuestion + vbYesNo, "Hapus?")
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            Pesan = MsgBox("Are you sure to delete this data?" & vbCrLf & vbCrLf & _
                            "===========================================" & vbCrLf & _
                            "Description" & vbCrLf & _
                            "===========================================" & vbCrLf & _
                            AdodcUtama.Recordset.Fields(0).Name & " : " & AdodcUtama.Recordset.Fields(0).Value & vbCrLf & _
                            AdodcUtama.Recordset.Fields(1).Name & " : " & AdodcUtama.Recordset.Fields(1).Value & vbCrLf & _
                            "===========================================", vbQuestion + vbYesNo, "Hapus?")
        End If
    Else
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini?" & vbCrLf & vbCrLf & _
                            "===========================================" & vbCrLf & _
                            "Keterangan" & vbCrLf & _
                            "===========================================" & vbCrLf & _
                            AdodcUtama.Recordset.Fields(0).Name & " : " & AdodcUtama.Recordset.Fields(0).Value & vbCrLf & _
                            "===========================================", vbQuestion + vbYesNo, "Hapus?")
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            Pesan = MsgBox("Are you sure to delete this data?" & vbCrLf & vbCrLf & _
                            "===========================================" & vbCrLf & _
                            "Description" & vbCrLf & _
                            "===========================================" & vbCrLf & _
                            AdodcUtama.Recordset.Fields(0).Name & " : " & AdodcUtama.Recordset.Fields(0).Value & vbCrLf & _
                            "===========================================", vbQuestion + vbYesNo, "Hapus?")
        End If
    End If
    If Pesan = vbYes Then AdodcUtama.Recordset.Delete
End Sub

Private Sub cmPropertis_Click()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "========================" & vbCrLf & _
                "Jumlah Data    : " & AdodcUtama.Recordset.RecordCount & vbCrLf & _
                "Jumlah Field   : " & AdodcUtama.Recordset.Fields.Count & vbCrLf & _
                "Jumlah Cell    : " & AdodcUtama.Recordset.Fields.Count * AdodcUtama.Recordset.RecordCount & vbCrLf & _
                "=======================", vbInformation + vbOKOnly, "Properties"
    Else
        MsgBox "========================" & vbCrLf & _
                "Data Count    : " & AdodcUtama.Recordset.RecordCount & vbCrLf & _
                "Field Count   : " & AdodcUtama.Recordset.Fields.Count & vbCrLf & _
                "Cell Count    : " & AdodcUtama.Recordset.Fields.Count * AdodcUtama.Recordset.RecordCount & vbCrLf & _
                "=======================", vbInformation + vbOKOnly, "Properties"
    End If
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
    PENGATURAN_WARNA
    PENGATURAN_BAHASA
End Sub

Sub PENGATURAN_WARNA()
    'PENGATURAN WARNA UNTUK FORM INI
    For Each Objek In Me
        Select Case FormPengaturan.cmbWarnaTampilan.ListIndex
        Case Is = 0 'Ungu Natural
            Me.BackColor = UnguNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguNatural
        Case Is = 1 'Merah
            Me.BackColor = Merah
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Merah
        Case Is = 2 'Pink
            Me.BackColor = Pink
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Pink
        Case Is = 3 'HijauMuda
            Me.BackColor = HijauMuda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = HijauMuda
        Case Is = 4 'Hitam
            Me.BackColor = Hitam
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Hitam
        Case Is = 5 'Silver
            Me.BackColor = Silver
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Silver
        Case Is = 6 'SilverNatural
            Me.BackColor = SilverNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = SilverNatural
        Case Is = 7 'Orange
            Me.BackColor = Orange
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Orange
        Case Is = 8 'UnguJanda
            Me.BackColor = UnguJanda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguJanda
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

Sub PENGATURAN_BAHASA()
If FormPengaturan.cmbBahasa.ListIndex = 0 Then
    With Me
        .cmBaru.Caption = "&Baru"
        .cmEdit.Caption = "&Edit"
        .cmHapus.Caption = "&Hapus"
        .cmPropertis.Caption = "&Propertis"
        .cmTutup.Caption = "&Tutup"
        .Label1.Caption = "Nama Tabel"
    End With
ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
    With Me
        .cmBaru.Caption = "&New"
        .cmEdit.Caption = "&Edit"
        .cmHapus.Caption = "&Delete"
        .cmPropertis.Caption = "&Properties"
        .cmTutup.Caption = "&Close"
        .Label1.Caption = "Tabel Name"
    End With
End If
End Sub
