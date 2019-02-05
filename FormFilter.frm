VERSION 5.00
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormFilter 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Data"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CMBFilterBerdasarkan 
      Height          =   375
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   4815
   End
   Begin VB.ComboBox cmbMode 
      Height          =   375
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   4815
   End
   Begin XPControls.XPCheck cekTutupForm 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      Caption         =   "Tutup Setelah Disorot"
      Value           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton cmOK 
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&OK"
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
      PicDown         =   "FormFilter.frx":030A
      PicHot          =   "FormFilter.frx":0624
      PicNormal       =   "FormFilter.frx":093E
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmBatal 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Batal"
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
      PicDown         =   "FormFilter.frx":0C58
      PicHot          =   "FormFilter.frx":10AA
      PicNormal       =   "FormFilter.frx":14FC
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
   Begin VB.Label LabelFilter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "label....."
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   465
   End
   Begin VB.Label LabelMode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "label....."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   465
   End
End
Attribute VB_Name = "FormFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
On Error GoTo BunuhError
If FORM_UTAMA.cmJejaringSosial.FontBold = True Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmJejaringSosial.Caption & " berdasarkan :"
        Me.Caption = "Filter Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(# " & FORM_UTAMA.cmJejaringSosial.Caption & "))"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmJejaringSosial.Caption & " berdasarkan :"
        Me.Caption = "Filter Of Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(# " & FORM_UTAMA.cmJejaringSosial.Caption & "))"
    End If
        With Me.CMBFilterBerdasarkan
            .Clear
            .AddItem FormManage.AdodcMain.Recordset.Fields(0).Name & " / Social Name", 0
            .AddItem FormManage.AdodcMain.Recordset.Fields(1).Name & " / User Name", 1
            .AddItem FormManage.AdodcMain.Recordset.Fields(2).Name & " / E_Mail Address", 2
            .AddItem FormManage.AdodcMain.Recordset.Fields(3).Name & " / Password", 3
            .AddItem FormManage.AdodcMain.Recordset.Fields(4).Name & " / URL", 4
            .AddItem FormManage.AdodcMain.Recordset.Fields(5).Name & " / Account Owner", 5
            .AddItem FormManage.AdodcMain.Recordset.Fields(6).Name & " / Date", 6
            .AddItem FormManage.AdodcMain.Recordset.Fields(7).Name & " / Description", 7
            .ListIndex = 0
        End With
ElseIf FORM_UTAMA.cmElectronicMail.FontBold = True Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmElectronicMail.Caption & " berdasarkan :"
        Me.Caption = "Filter Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(# " & FORM_UTAMA.cmElectronicMail.Caption & "))"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmElectronicMail.Caption & " berdasarkan :"
        Me.Caption = "Filter Of Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(# " & FORM_UTAMA.cmElectronicMail.Caption & "))"
    End If
        With Me.CMBFilterBerdasarkan
            .Clear
            .AddItem FormManage.AdodcMain.Recordset.Fields(0).Name & " / Server Name", 0
            .AddItem FormManage.AdodcMain.Recordset.Fields(1).Name & " / User Name", 1
            .AddItem FormManage.AdodcMain.Recordset.Fields(2).Name & " / E_Mail Address", 2
            .AddItem FormManage.AdodcMain.Recordset.Fields(3).Name & " / Password", 3
            .AddItem FormManage.AdodcMain.Recordset.Fields(4).Name & " / Security Question", 4
            .AddItem FormManage.AdodcMain.Recordset.Fields(5).Name & " / Security Answer", 5
            .AddItem FormManage.AdodcMain.Recordset.Fields(6).Name & " / URL", 6
            .AddItem FormManage.AdodcMain.Recordset.Fields(7).Name & " / Account Owner", 7
            .AddItem FormManage.AdodcMain.Recordset.Fields(8).Name & " / Date", 8
            .AddItem FormManage.AdodcMain.Recordset.Fields(9).Name & " / Description", 9
            .ListIndex = 0
        End With
ElseIf FORM_UTAMA.cmForumInternet.FontBold = True Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmForumInternet.Caption & " berdasarkan :"
        Me.Caption = "Filter Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(# " & FORM_UTAMA.cmForumInternet.Caption & "))"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmForumInternet.Caption & " berdasarkan :"
        Me.Caption = "Filter Of Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(# " & FORM_UTAMA.cmForumInternet.Caption & "))"
    End If
        With Me.CMBFilterBerdasarkan
            .Clear
            .AddItem FormManage.AdodcMain.Recordset.Fields(0).Name & " / Forum Name", 0
            .AddItem FormManage.AdodcMain.Recordset.Fields(1).Name & " / User Name", 1
            .AddItem FormManage.AdodcMain.Recordset.Fields(2).Name & " / E_Mail Address", 2
            .AddItem FormManage.AdodcMain.Recordset.Fields(3).Name & " / Password", 3
            .AddItem FormManage.AdodcMain.Recordset.Fields(4).Name & " / Position", 4
            .AddItem FormManage.AdodcMain.Recordset.Fields(5).Name & " / NickName", 5
            .AddItem FormManage.AdodcMain.Recordset.Fields(6).Name & " / URL", 6
            .AddItem FormManage.AdodcMain.Recordset.Fields(7).Name & " / Date", 7
            .AddItem FormManage.AdodcMain.Recordset.Fields(8).Name & " / Description", 8
            .ListIndex = 0
        End With
ElseIf FORM_UTAMA.cmFTP.FontBold = True Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmFTP.Caption & " berdasarkan :"
        Me.Caption = "Filter Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(# " & FORM_UTAMA.cmFTP.Caption & "))"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmFTP.Caption & " berdasarkan :"
        Me.Caption = "Filter Of Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(# " & FORM_UTAMA.cmFTP.Caption & "))"
    End If
        With Me.CMBFilterBerdasarkan
            .Clear
            .AddItem FormManage.AdodcMain.Recordset.Fields(0).Name & " / Host Name", 0
            .AddItem FormManage.AdodcMain.Recordset.Fields(1).Name & " / Port", 1
            .AddItem FormManage.AdodcMain.Recordset.Fields(2).Name & " / Server Name", 2
            .AddItem FormManage.AdodcMain.Recordset.Fields(3).Name & " / User Name", 3
            .AddItem FormManage.AdodcMain.Recordset.Fields(4).Name & " / E-Mail", 4
            .AddItem FormManage.AdodcMain.Recordset.Fields(5).Name & " / Password", 5
            .AddItem FormManage.AdodcMain.Recordset.Fields(6).Name & " / Date", 6
            .AddItem FormManage.AdodcMain.Recordset.Fields(7).Name & " / Description", 7
            .ListIndex = 0
        End With
ElseIf FORM_UTAMA.cmBlogging.FontBold = True Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmBlogging.Caption & " berdasarkan :"
        Me.Caption = "Filter Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(# " & FORM_UTAMA.cmBlogging.Caption & "))"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmBlogging.Caption & " berdasarkan :"
        Me.Caption = "Filter Of Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(# " & FORM_UTAMA.cmBlogging.Caption & "))"
    End If
        With Me.CMBFilterBerdasarkan
            .Clear
            .AddItem FormManage.AdodcMain.Recordset.Fields(0).Name & " / Blogs Name Provider", 0
            .AddItem FormManage.AdodcMain.Recordset.Fields(1).Name & " / User Name", 1
            .AddItem FormManage.AdodcMain.Recordset.Fields(2).Name & " / E-Mail", 2
            .AddItem FormManage.AdodcMain.Recordset.Fields(3).Name & " / Password", 3
            .AddItem FormManage.AdodcMain.Recordset.Fields(4).Name & " / URL", 4
            .AddItem FormManage.AdodcMain.Recordset.Fields(5).Name & " / Date", 5
            .AddItem FormManage.AdodcMain.Recordset.Fields(6).Name & " / Description", 6
            .ListIndex = 0
        End With
ElseIf FORM_UTAMA.cmIdentitasPribadi.FontBold = True Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmIdentitasPribadi.Caption & " berdasarkan :"
        Me.Caption = "Filter Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmIdentitasPribadi.Caption & "))"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmIdentitasPribadi.Caption & " berdasarkan :"
        Me.Caption = "Filter Of Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmIdentitasPribadi.Caption & "))"
    End If
        With Me.CMBFilterBerdasarkan
            .Clear
            .AddItem FormManage.AdodcMain.Recordset.Fields(0).Name & " / Original Name", 0
            .AddItem FormManage.AdodcMain.Recordset.Fields(1).Name & " / Nick Name", 1
            .AddItem FormManage.AdodcMain.Recordset.Fields(2).Name & " / Place Born", 2
            .AddItem FormManage.AdodcMain.Recordset.Fields(3).Name & " / Date of Born", 3
            .AddItem FormManage.AdodcMain.Recordset.Fields(4).Name & " / Month of Born", 4
            .AddItem FormManage.AdodcMain.Recordset.Fields(5).Name & " / Year of Born", 5
            .AddItem FormManage.AdodcMain.Recordset.Fields(6).Name & " / Gender", 6
            .AddItem FormManage.AdodcMain.Recordset.Fields(7).Name & " / Religion", 7
            .AddItem FormManage.AdodcMain.Recordset.Fields(8).Name & " / Blood Type", 8
            .AddItem FormManage.AdodcMain.Recordset.Fields(9).Name & " / Jobs", 9
            .AddItem FormManage.AdodcMain.Recordset.Fields(10).Name & " / Home Address", 10
            .AddItem FormManage.AdodcMain.Recordset.Fields(11).Name & " / Mail Address", 11
            .AddItem FormManage.AdodcMain.Recordset.Fields(12).Name & " / Website", 12
            .AddItem FormManage.AdodcMain.Recordset.Fields(13).Name & " / Phone Number", 13
            .AddItem FormManage.AdodcMain.Recordset.Fields(14).Name & " / Home Town", 14
            .AddItem FormManage.AdodcMain.Recordset.Fields(15).Name & " / Town is Now", 15
            .AddItem FormManage.AdodcMain.Recordset.Fields(16).Name & " / ZIP Code", 16
            .AddItem FormManage.AdodcMain.Recordset.Fields(17).Name & " / Province", 17
            .AddItem FormManage.AdodcMain.Recordset.Fields(18).Name & " / Citizenship", 18
            .AddItem FormManage.AdodcMain.Recordset.Fields(19).Name & " / Education Status", 19
            .AddItem FormManage.AdodcMain.Recordset.Fields(20).Name & " / Relationship Status", 20
            .AddItem FormManage.AdodcMain.Recordset.Fields(21).Name & " / Hobby", 21
            .AddItem FormManage.AdodcMain.Recordset.Fields(22).Name & " / Description", 22
            .ListIndex = 0
        End With
ElseIf FORM_UTAMA.cmBukuAlamat.FontBold = True Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmBukuAlamat.Caption & " berdasarkan :"
        Me.Caption = "Filter Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmBukuAlamat.Caption & "))"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmBukuAlamat.Caption & " berdasarkan :"
        Me.Caption = "Filter Of Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmBukuAlamat.Caption & "))"
    End If
        With Me.CMBFilterBerdasarkan
            .Clear
            .AddItem FormManage.AdodcMain.Recordset.Fields(0).Name & " / Contact Name", 0
            .AddItem FormManage.AdodcMain.Recordset.Fields(1).Name & " / Nick Name", 1
            .AddItem FormManage.AdodcMain.Recordset.Fields(2).Name & " / Private Phone", 2
            .AddItem FormManage.AdodcMain.Recordset.Fields(3).Name & " / Home Phone", 3
            .AddItem FormManage.AdodcMain.Recordset.Fields(4).Name & " / Office Phone", 4
            .AddItem FormManage.AdodcMain.Recordset.Fields(5).Name & " / Fax", 5
            .AddItem FormManage.AdodcMain.Recordset.Fields(6).Name & " / Mail Address", 6
            .AddItem FormManage.AdodcMain.Recordset.Fields(7).Name & " / Website", 7
            .AddItem FormManage.AdodcMain.Recordset.Fields(8).Name & " / ZIP Postal Code", 8
            .AddItem FormManage.AdodcMain.Recordset.Fields(9).Name & " / Home Address", 9
            .AddItem FormManage.AdodcMain.Recordset.Fields(10).Name & " / Description", 10
            .ListIndex = 0
        End With
ElseIf FORM_UTAMA.cmUlangTahun.FontBold = True Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmUlangTahun.Caption & " berdasarkan :"
        Me.Caption = "Filter Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(# " & FORM_UTAMA.cmUlangTahun.Caption & "))"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmUlangTahun.Caption & " berdasarkan :"
        Me.Caption = "Filter Of Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(# " & FORM_UTAMA.cmUlangTahun.Caption & "))"
    End If
        With Me.CMBFilterBerdasarkan
            .Clear
            .AddItem FormManage.AdodcMain.Recordset.Fields(0).Name & " / Name", 0
            .AddItem FormManage.AdodcMain.Recordset.Fields(1).Name & " / Birthday", 1
            .AddItem FormManage.AdodcMain.Recordset.Fields(2).Name & " / Decsription", 2
            .ListIndex = 0
        End With
ElseIf FORM_UTAMA.cmAgenda.FontBold = True Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmAgenda.Caption & " berdasarkan :"
        Me.Caption = "Filter Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmAgenda.Caption & "))"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmAgenda.Caption & " berdasarkan :"
        Me.Caption = "Filter Of Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmAgenda.Caption & "))"
    End If
        With Me.CMBFilterBerdasarkan
            .Clear
            .AddItem FormManage.AdodcMain.Recordset.Fields(0).Name & " / Agenda's Code", 0
            .AddItem FormManage.AdodcMain.Recordset.Fields(1).Name & " / Agenda's Name", 1
            .AddItem FormManage.AdodcMain.Recordset.Fields(2).Name & " / Thema", 2
            .AddItem FormManage.AdodcMain.Recordset.Fields(3).Name & " / Date", 3
            .AddItem FormManage.AdodcMain.Recordset.Fields(4).Name & " / Starting Time", 4
            .AddItem FormManage.AdodcMain.Recordset.Fields(5).Name & " / Ending Time", 5
            .AddItem FormManage.AdodcMain.Recordset.Fields(6).Name & " / Place", 6
            .AddItem FormManage.AdodcMain.Recordset.Fields(7).Name & " / Other Decsription", 7
            .ListIndex = 0
        End With
ElseIf FORM_UTAMA.cmRegistrasiSoftware.FontBold = True Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmRegistrasiSoftware.Caption & " berdasarkan :"
        Me.Caption = "Filter Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmRegistrasiSoftware.Caption & "))"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        LabelFilter.Caption = "Filter data pada akun " & FORM_UTAMA.cmRegistrasiSoftware.Caption & " berdasarkan :"
        Me.Caption = "Filter Of Data (@" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "(#" & FORM_UTAMA.cmRegistrasiSoftware.Caption & "))"
    End If
        With Me.CMBFilterBerdasarkan
            .Clear
            .AddItem FormManage.AdodcMain.Recordset.Fields(0).Name & " / Software Name", 0
            .AddItem FormManage.AdodcMain.Recordset.Fields(1).Name & " / Categories Name", 1
            .AddItem FormManage.AdodcMain.Recordset.Fields(2).Name & " / Developer", 2
            .AddItem FormManage.AdodcMain.Recordset.Fields(3).Name & " / UserName", 3
            .AddItem FormManage.AdodcMain.Recordset.Fields(4).Name & " / Serial-Key", 4
            .AddItem FormManage.AdodcMain.Recordset.Fields(5).Name & " / License Type", 5
            .AddItem FormManage.AdodcMain.Recordset.Fields(6).Name & " / Description", 6
            .ListIndex = 0
        End With
End If
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        LabelMode.Caption = "Filter Dengan Mode :"
        cekTutupForm.Caption = "Tutup Setelah DiFilter"
        cmBatal.Caption = "&Batal"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        LabelMode.Caption = "Mode of Filter :"
        cekTutupForm.Caption = "Close window after Filtered"
        cmBatal.Caption = "&Cancel"
    End If
    With cmbMode
        .Clear
        .AddItem "Ascending / A to Z / Terurut Dari Awal", 0
        .AddItem "Descending / Z to A / Terurut Dari Akhir", 1
        .AddItem "Normal / No Filter /Tidak Ada Filter", 2
        .ListIndex = 0
    End With
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
Exit Sub
BunuhError:
    HancurkanError
End Sub

Private Sub cmBatal_Click()
    With FormManage
        .AturDatabase
        .cmEdit.Enabled = True
        .cmHapus.Enabled = True
        .cmSorot.Enabled = True
        .cmCari.Enabled = True
    End With
    Unload Me
End Sub

Private Sub cmOK_Click()
If FORM_UTAMA.cmJejaringSosial.FontBold = True Then
    Select Case Me.cmbMode.ListIndex
        Case Is = 0
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Jejaring From tbJejaringSosial Order by Nama_Jejaring asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Pengguna From tbJejaringSosial Order by Nama_Pengguna asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Alamat_Email From tbJejaringSosial Order by Alamat_Email asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Password From tbJejaringSosial Order by Password asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select URL From tbJejaringSosial Order by URL asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Pemilik_Akun From tbJejaringSosial Order by Pemilik_Akun asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Tanggal From tbJejaringSosial Order by Tanggal asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Keterangan From tbJejaringSosial Order by Keterangan asc;"
                End If
                .Refresh
            End With
        Case Is = 1
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Jejaring From tbJejaringSosial Order by Nama_Jejaring desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Pengguna From tbJejaringSosial Order by Nama_Pengguna desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Alamat_Email From tbJejaringSosial Order by Alamat_Email desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Password From tbJejaringSosial Order by Password desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select URL From tbJejaringSosial Order by URL desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Pemilik_Akun From tbJejaringSosial Order by Pemilik_Akun desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Tanggal From tbJejaringSosial Order by Tanggal desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Keterangan From tbJejaringSosial Order by Keterangan desc;"
                End If
                .Refresh
            End With
        Case Is = 2
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Jejaring From tbJejaringSosial"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Pengguna From tbJejaringSosial"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Alamat_Email From tbJejaringSosial"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Password From tbJejaringSosial"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select URL From tbJejaringSosial"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Pemilik_Akun From tbJejaringSosial"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Tanggal From tbJejaringSosial"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Keterangan From tbJejaringSosial"
                End If
                .Refresh
            End With
    End Select
ElseIf FORM_UTAMA.cmElectronicMail.FontBold = True Then
    Select Case Me.cmbMode.ListIndex
        Case Is = 0
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Server From tbElectronicMail Order by Nama_Server asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Pengguna From tbElectronicMail Order by Nama_Pengguna asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Alamat_Email From tbElectronicMail Order by Alamat_Email asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Password From tbElectronicMail Order by Password asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select Pertanyaan_Rahasia From tbElectronicMail Order by Pertanyaan_Rahasia asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Jawaban_Pertanyaan From tbElectronicMail Order by Jawaban_Pertanyaan asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select URL From tbElectronicMail Order by URL asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Pemilik_Akun From tbElectronicMail Order by Pemilik_Akun asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 8 Then
                    .RecordSource = "Select Tanggal From tbElectronicMail Order by Tanggal asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 9 Then
                    .RecordSource = "Select Keterangan From tbElectronicMail Order by Keterangan asc;"
                End If
                .Refresh
            End With
        Case Is = 1
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Server From tbElectronicMail Order by Nama_Server desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Pengguna From tbElectronicMail Order by Nama_Pengguna desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Alamat_Email From tbElectronicMail Order by Alamat_Email desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Password From tbElectronicMail Order by Password desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select Pertanyaan_Rahasia From tbElectronicMail Order by Pertanyaan_Rahasia desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Jawaban_Pertanyaan From tbElectronicMail Order by Jawaban_Pertanyaan desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select URL From tbElectronicMail Order by URL desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Pemilik_Akun From tbElectronicMail Order by Pemilik_Akun desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 8 Then
                    .RecordSource = "Select Tanggal From tbElectronicMail Order by Tanggal desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 9 Then
                    .RecordSource = "Select Keterangan From tbElectronicMail Order by Keterangan desc;"
                End If
                .Refresh
            End With
        Case Is = 2
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Server From tbElectronicMail"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Pengguna From tbElectronicMail"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Alamat_Email From tbElectronicMail"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Password From tbElectronicMail"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select Pertanyaan_Rahasia From tbElectronicMail"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Jawaban_Pertanyaan From tbElectronicMail"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select URL From tbElectronicMail"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Pemilik_Akun From tbElectronicMail"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 8 Then
                    .RecordSource = "Select Tanggal From tbElectronicMail"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 9 Then
                    .RecordSource = "Select Keterangan From tbElectronicMail"
                End If
                .Refresh
            End With
    End Select
ElseIf FORM_UTAMA.cmForumInternet.FontBold = True Then
    Select Case Me.cmbMode.ListIndex
        Case Is = 0
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Forum From tbForumInternet Order by Nama_Forum asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Pengguna From tbForumInternet Order by Nama_Pengguna asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Alamat_Email From tbForumInternet Order by Alamat_Email asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Password From tbForumInternet Order by Password asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select Posisi From tbForumInternet Order by Posisi asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select NickName From tbForumInternet Order by NickName asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select URL From tbForumInternet Order by URL asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Tanggal From tbForumInternet Order by Tanggal asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 8 Then
                    .RecordSource = "Select Keterangan From tbForumInternet Order by Keterangan asc;"
                End If
                .Refresh
            End With
        Case Is = 1
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Forum From tbForumInternet Order by Nama_Forum desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Pengguna From tbForumInternet Order by Nama_Pengguna desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Alamat_Email From tbForumInternet Order by Alamat_Email desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Password From tbForumInternet Order by Password desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select Posisi From tbForumInternet Order by Posisi desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select NickName From tbForumInternet Order by NickName desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select URL From tbForumInternet Order by URL desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Tanggal From tbForumInternet Order by Tanggal desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 8 Then
                    .RecordSource = "Select Keterangan From tbForumInternet Order by Keterangan desc;"
                End If
                .Refresh
            End With
        Case Is = 2
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Forum From tbForumInternet"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Pengguna From tbForumInternet"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Alamat_Email From tbForumInternet"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Password From tbForumInternet"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select Posisi From tbForumInternet"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select NickName From tbForumInternet"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select URL From tbForumInternet"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Tanggal From tbForumInternet"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 8 Then
                    .RecordSource = "Select Keterangan From tbForumInternet"
                End If
                .Refresh
            End With
    End Select
ElseIf FORM_UTAMA.cmFTP.FontBold = True Then
    Select Case Me.cmbMode.ListIndex
        Case Is = 0
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Host From tbFTP Order by Nama_Host asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Port From tbFTP Order by Port asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Nama_Server From tbFTP Order by Nama_Server asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Nama_Pengguna From tbFTP Order by Nama_Pengguna asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select Alamat_Email From tbFTP Order by Alamat_Email asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Password From tbFTP Order by Password asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Tanggal From tbFTP Order by Tanggal asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Keterangan From tbFTP Order by Keterangan asc;"
                End If
                .Refresh
            End With
        Case Is = 1
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Host From tbFTP Order by Nama_Host desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Port From tbFTP Order by Port desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Nama_Server From tbFTP Order by Nama_Server desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Nama_Pengguna From tbFTP Order by Nama_Pengguna desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select Alamat_Email From tbFTP Order by Alamat_Email desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Password From tbFTP Order by Password desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Tanggal From tbFTP Order by Tanggal desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Keterangan From tbFTP Order by Keterangan desc;"
                End If
                .Refresh
            End With
        Case Is = 2
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Host From tbFTP"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Port From tbFTP"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Nama_Server From tbFTP"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Nama_Pengguna From tbFTP"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select Alamat_Email From tbFTP"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Password From tbFTP"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Tanggal From tbFTP"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Keterangan From tbFTP"
                End If
                .Refresh
            End With
    End Select
ElseIf FORM_UTAMA.cmBlogging.FontBold = True Then
    Select Case Me.cmbMode.ListIndex
        Case Is = 0
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Penyedia_Blog From tbBlogging Order by Nama_Penyedia_Blog asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Pengguna From tbBlogging Order by Nama_Pengguna asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select E_Mail From tbBlogging Order by E_Mail asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Password From tbBlogging Order by Password asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select URL From tbBlogging Order by URL asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Tanggal From tbBlogging Order by Tanggal asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Keterangan From tbBlogging Order by Keterangan asc;"
                End If
                .Refresh
            End With
        Case Is = 1
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Penyedia_Blog From tbBlogging Order by Nama_Penyedia_Blog desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Pengguna From tbBlogging Order by Nama_Pengguna desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select E_Mail From tbBlogging Order by E_Mail desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Password From tbBlogging Order by Password desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select URL From tbBlogging Order by URL desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Tanggal From tbBlogging Order by Tanggal desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Keterangan From tbBlogging Order by Keterangan desc;"
                End If
                .Refresh
            End With
        Case Is = 2
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Penyedia_Blog From tbBlogging Order"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Pengguna From tbBlogging Order"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select E_Mail From tbBlogging"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Password From tbBlogging"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select URL From tbBlogging"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Tanggal From tbBlogging"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Keterangan From tbBlogging"
                End If
                .Refresh
            End With
    End Select
ElseIf FORM_UTAMA.cmIdentitasPribadi.FontBold = True Then
    Select Case Me.cmbMode.ListIndex
        Case Is = 0
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Lengkap From tbIdentitasPribadi Order by Nama_Lengkap asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Panggilan From tbIdentitasPribadi Order by Nama_Panggilan asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select TempatLahir From tbIdentitasPribadi Order by TempatLahir asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select TanggalLahir From tbIdentitasPribadi Order by TanggalLahir asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select BulanLahir From tbIdentitasPribadi Order by BulanLahir asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select TahunLahir From tbIdentitasPribadi Order by TahunLahir asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Jenis_Kelamin From tbIdentitasPribadi Order by Jenis_Kelamin asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Agama From tbIdentitasPribadi Order by Agama asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 8 Then
                    .RecordSource = "Select Golongan_Darah From tbIdentitasPribadi Order by Golongan_Darah asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 9 Then
                    .RecordSource = "Select Pekerjaan From tbIdentitasPribadi Order by Pekerjaan asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 10 Then
                    .RecordSource = "Select Alamat_Rumah From tbIdentitasPribadi Order by Alamat_Rumah asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 11 Then
                    .RecordSource = "Select E_Mail From tbIdentitasPribadi Order by E_Mail asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 13 Then
                    .RecordSource = "Select Website From tbIdentitasPribadi Order by Website asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 14 Then
                    .RecordSource = "Select Nomor_Telepon From tbIdentitasPribadi Order by Nomor_Telepon asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 15 Then
                    .RecordSource = "Select Kota_Asal From tbIdentitasPribadi Order by Kota_Asal asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 16 Then
                    .RecordSource = "Select Kota_Sekarang From tbIdentitasPribadi Order by Kota_Sekarang asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 17 Then
                    .RecordSource = "Select Kode_Pos From tbIdentitasPribadi Order by Kode_Pos asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 18 Then
                    .RecordSource = "Select Provinsi From tbIdentitasPribadi Order by Provinsi asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 19 Then
                    .RecordSource = "Select Kewarganegaraan From tbIdentitasPribadi Order by Kewarganegaraan asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 20 Then
                    .RecordSource = "Select Status_Pendidikan From tbIdentitasPribadi Order by Status_Pendidikan asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 21 Then
                    .RecordSource = "Select Status_Hubungan From tbIdentitasPribadi Order by Status_Hubungan asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 22 Then
                    .RecordSource = "Select Hobby From tbIdentitasPribadi Order by Hobby asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 23 Then
                    .RecordSource = "Select Keterangan From tbIdentitasPribadi Order by Keterangan asc;"
                End If
                .Refresh
            End With
        Case Is = 1
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Lengkap From tbIdentitasPribadi Order by Nama_Lengkap desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Panggilan From tbIdentitasPribadi Order by Nama_Panggilan desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select TempatLahir From tbIdentitasPribadi Order by TempatLahir desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select TanggalLahir From tbIdentitasPribadi Order by TanggalLahir desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select BulanLahir From tbIdentitasPribadi Order by BulanLahir desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select TahunLahir From tbIdentitasPribadi Order by TahunLahir desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Jenis_Kelamin From tbIdentitasPribadi Order by Jenis_Kelamin desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Agama From tbIdentitasPribadi Order by Agama desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 8 Then
                    .RecordSource = "Select Golongan_Darah From tbIdentitasPribadi Order by Golongan_Darah desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 9 Then
                    .RecordSource = "Select Pekerjaan From tbIdentitasPribadi Order by Pekerjaan desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 10 Then
                    .RecordSource = "Select Alamat_Rumah From tbIdentitasPribadi Order by Alamat_Rumah desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 11 Then
                    .RecordSource = "Select E_Mail From tbIdentitasPribadi Order by E_Mail desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 13 Then
                    .RecordSource = "Select Website From tbIdentitasPribadi Order by Website desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 14 Then
                    .RecordSource = "Select Nomor_Telepon From tbIdentitasPribadi Order by Nomor_Telepon desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 15 Then
                    .RecordSource = "Select Kota_Asal From tbIdentitasPribadi Order by Kota_Asal desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 16 Then
                    .RecordSource = "Select Kota_Sekarang From tbIdentitasPribadi Order by Kota_Sekarang desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 17 Then
                    .RecordSource = "Select Kode_Pos From tbIdentitasPribadi Order by Kode_Pos desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 18 Then
                    .RecordSource = "Select Provinsi From tbIdentitasPribadi Order by Provinsi desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 19 Then
                    .RecordSource = "Select Kewarganegaraan From tbIdentitasPribadi Order by Kewarganegaraan desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 20 Then
                    .RecordSource = "Select Status_Pendidikan From tbIdentitasPribadi Order by Status_Pendidikan desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 21 Then
                    .RecordSource = "Select Status_Hubungan From tbIdentitasPribadi Order by Status_Hubungan desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 22 Then
                    .RecordSource = "Select Hobby From tbIdentitasPribadi Order by Hobby desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 23 Then
                    .RecordSource = "Select Keterangan From tbIdentitasPribadi Order by Keterangan desc;"
                End If
                .Refresh
            End With
        Case Is = 2
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                .RecordSource = "Select * From tbIdentitasPribadi;"
                .Refresh
            End With
    End Select
ElseIf FORM_UTAMA.cmBukuAlamat.FontBold = True Then
    Select Case Me.cmbMode.ListIndex
        Case Is = 0
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Kontak From tbBukuAlamat Order by Nama_Kontak asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Panggilan From tbBukuAlamat Order by Nama_Panggilan asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Nomor_Telepon_Pribadi From tbBukuAlamat Order by Nomor_Telepon_Pribadi asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Nomor_Telepon_Rumah From tbBukuAlamat Order by Nomor_Telepon_Rumah asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select Nomor_Telepon_Kantor From tbBukuAlamat Order by Nomor_Telepon_Kantor asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Fax From tbBukuAlamat Order by Fax asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Alamat_EMail From tbBukuAlamat Order by Alamat_EMail asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Website From tbBukuAlamat Order by Website asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 8 Then
                    .RecordSource = "Select ZIP_Postal_Code From tbBukuAlamat Order by ZIP_Postal_Code asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 9 Then
                    .RecordSource = "Select Alamat_Rumah From tbBukuAlamat Order by Alamat_Rumah asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 10 Then
                    .RecordSource = "Select Keterangan From tbBukuAlamat Order by Keterangan asc;"
                End If
                .Refresh
            End With
        Case Is = 1
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Kontak From tbBukuAlamat Order by Nama_Kontak desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Panggilan From tbBukuAlamat Order by Nama_Panggilan desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Nomor_Telepon_Pribadi From tbBukuAlamat Order by Nomor_Telepon_Pribadi desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Nomor_Telepon_Rumah From tbBukuAlamat Order by Nomor_Telepon_Rumah desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select Nomor_Telepon_Kantor From tbBukuAlamat Order by Nomor_Telepon_Kantor desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Fax From tbBukuAlamat Order by Fax desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Alamat_EMail From tbBukuAlamat Order by Alamat_EMail desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Website From tbBukuAlamat Order by Website desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 8 Then
                    .RecordSource = "Select ZIP_Postal_Code From tbBukuAlamat Order by ZIP_Postal_Code desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 9 Then
                    .RecordSource = "Select Alamat_Rumah From tbBukuAlamat Order by Alamat_Rumah desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 10 Then
                    .RecordSource = "Select Keterangan From tbBukuAlamat Order by Keterangan desc;"
                End If
                .Refresh
            End With
        Case Is = 2
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                .RecordSource = "Select * From tbBukuAlamat;"
                .Refresh
            End With
    End Select
ElseIf FORM_UTAMA.cmUlangTahun.FontBold = True Then
    Select Case Me.cmbMode.ListIndex
        Case Is = 0
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama From tbUlangTahun Order by Nama asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select TTL From tbUlangTahun Order by TTL asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Keterangan From tbUlangTahun Order by Keterangan asc;"
                End If
                .Refresh
            End With
        Case Is = 1
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama From tbUlangTahun Order by Nama desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select TTL From tbUlangTahun Order by TTL desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Keterangan From tbUlangTahun Order by Keterangan desc;"
                End If
                .Refresh
            End With
        Case Is = 2
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                .RecordSource = "Select * From tbUlangTahun;"
                .Refresh
            End With
    End Select
ElseIf FORM_UTAMA.cmAgenda.FontBold = True Then
    Select Case Me.cmbMode.ListIndex
        Case Is = 0
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Kode_Agenda From tbAgenda Order by Kode_Agenda asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Agenda From tbAgenda Order by Nama_Agenda asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Tema From tbAgenda Order by Tema asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Tanggal From tbAgenda Order by Tanggal asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select Waktu_Mulai From tbAgenda Order by Waktu_Mulai asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Waktu_Akhir From tbAgenda Order by Waktu_Akhir asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Tempat From tbAgenda Order by Tempat asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Keterangan_Lain From tbAgenda Order by Keterangan_Lain asc;"
                End If
                .Refresh
            End With
        Case Is = 1
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Kode_Agenda From tbAgenda Order by Kode_Agenda desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Nama_Agenda From tbAgenda Order by Nama_Agenda desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Tema From tbAgenda Order by Tema desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Tanggal From tbAgenda Order by Tanggal desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select Waktu_Mulai From tbAgenda Order by Waktu_Mulai desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Waktu_Akhir From tbAgenda Order by Waktu_Akhir desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Tempat From tbAgenda Order by Tempat desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 7 Then
                    .RecordSource = "Select Keterangan_Lain From tbAgenda Order by Keterangan_Lain desc;"
                End If
                .Refresh
            End With
        Case Is = 2
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                .RecordSource = "Select * From tbAgenda;"
                .Refresh
            End With
    End Select
ElseIf FORM_UTAMA.cmRegistrasiSoftware.FontBold = True Then
    Select Case Me.cmbMode.ListIndex
        Case Is = 0
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Software From tbRegistrasiSoftware Order by Nama_Software asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Kategori From tbRegistrasiSoftware Order by Kategori asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Developer From tbRegistrasiSoftware Order by Developer asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Username From tbRegistrasiSoftware Order by Username asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select Serial_Key From tbRegistrasiSoftware Order by Serial_Key asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Jenis_Lisensi From tbRegistrasiSoftware Order by Jenis_Lisensi asc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Keterangan From tbRegistrasiSoftware Order by Keterangan asc;"
                End If
                .Refresh
            End With
        Case Is = 1
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                If Me.CMBFilterBerdasarkan.ListIndex = 0 Then
                    .RecordSource = "Select Nama_Software From tbRegistrasiSoftware Order by Nama_Software desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 1 Then
                    .RecordSource = "Select Kategori From tbRegistrasiSoftware Order by Kategori desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 2 Then
                    .RecordSource = "Select Developer From tbRegistrasiSoftware Order by Developer desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 3 Then
                    .RecordSource = "Select Username From tbRegistrasiSoftware Order by Username desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 4 Then
                    .RecordSource = "Select Serial_Key From tbRegistrasiSoftware Order by Serial_Key desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 5 Then
                    .RecordSource = "Select Jenis_Lisensi From tbRegistrasiSoftware Order by Jenis_Lisensi desc;"
                ElseIf Me.CMBFilterBerdasarkan.ListIndex = 6 Then
                    .RecordSource = "Select Keterangan From tbRegistrasiSoftware Order by Keterangan desc;"
                End If
                .Refresh
            End With
        Case Is = 2
            With FormManage.AdodcMain
                .ConnectionString = CN_FormUtama.ConnectionString
                .RecordSource = "Select * From tbRegistrasiSoftware;"
                .Refresh
            End With
    End Select
End If
With FormManage
    .cmEdit.Enabled = False
    .cmHapus.Enabled = False
    .cmSorot.Enabled = False
    .cmCari.Enabled = False
End With
If Me.cekTutupForm.Value = Checked Then Me.Hide
End Sub

Private Sub Form_Load()
    AturKontrol
    PENGATURAN_WARNA
End Sub

Sub PENGATURAN_WARNA()
    'PENGATURAN WARNA UNTUK FORM INI
    For Each Objek In Me
        Select Case FormPengaturan.cmbWarnaTampilan.ListIndex
        Case Is = 0 'Ungu Natural
            Me.BackColor = UnguNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguNatural
            If TypeName(Objek) = "Frame" Then Objek.BackColor = UnguNatural
            If TypeName(Objek) = "XPCheck" Then Objek.BackColor = UnguNatural
        Case Is = 1 'Merah
            Me.BackColor = Merah
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Merah
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Merah
            If TypeName(Objek) = "XPCheck" Then Objek.BackColor = Merah
        Case Is = 2 'Pink
            Me.BackColor = Pink
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Pink
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Pink
            If TypeName(Objek) = "XPCheck" Then Objek.BackColor = Pink
        Case Is = 3 'HijauMuda
            Me.BackColor = HijauMuda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = HijauMuda
            If TypeName(Objek) = "Frame" Then Objek.BackColor = HijauMuda
            If TypeName(Objek) = "XPCheck" Then Objek.BackColor = HijauMuda
        Case Is = 4 'Hitam
            Me.BackColor = Hitam
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Hitam
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Hitam
            If TypeName(Objek) = "XPCheck" Then Objek.BackColor = Hitam
        Case Is = 5 'Silver
            Me.BackColor = Silver
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Silver
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Silver
            If TypeName(Objek) = "XPCheck" Then Objek.BackColor = Silver
        Case Is = 6 'SilverNatural
            Me.BackColor = SilverNatural
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = SilverNatural
            If TypeName(Objek) = "Frame" Then Objek.BackColor = SilverNatural
            If TypeName(Objek) = "XPCheck" Then Objek.BackColor = SilverNatural
        Case Is = 7 'Orange
            Me.BackColor = Orange
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = Orange
            If TypeName(Objek) = "Frame" Then Objek.BackColor = Orange
            If TypeName(Objek) = "XPCheck" Then Objek.BackColor = Orange
        Case Is = 8 'UnguJanda
            Me.BackColor = UnguJanda
            If TypeName(Objek) = "dcButton" Then Objek.BackColor = UnguJanda
            If TypeName(Objek) = "Frame" Then Objek.BackColor = UnguJanda
            If TypeName(Objek) = "XPCheck" Then Objek.BackColor = UnguJanda
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
