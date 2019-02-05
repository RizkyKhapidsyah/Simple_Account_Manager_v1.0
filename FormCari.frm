VERSION 5.00
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormCari 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "---"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormCari.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Dacara_dcButton.dcButton cmCari 
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   1250
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Cari"
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
      PicDown         =   "FormCari.frx":030A
      PicHot          =   "FormCari.frx":0624
      PicNormal       =   "FormCari.frx":093E
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1150
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox cmbCariBerdasarkan 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
      Begin XPControls.XPText TextKriteria 
         Height          =   330
         Left            =   1920
         TabIndex        =   3
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         Text            =   ""
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   6
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dengan Kriteria"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cari Berdasarkan"
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1155
      End
   End
   Begin Dacara_dcButton.dcButton cmBatal 
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   1250
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
      PicDown         =   "FormCari.frx":0C58
      PicHot          =   "FormCari.frx":10AA
      PicNormal       =   "FormCari.frx":14FC
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmBantuan 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1250
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Bantuan"
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
      PicDown         =   "FormCari.frx":194E
      PicHot          =   "FormCari.frx":1C68
      PicNormal       =   "FormCari.frx":1F82
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
Attribute VB_Name = "FormCari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    cmbCariBerdasarkan.Clear
        If FORM_UTAMA.cmJejaringSosial.FontBold = True Then
            With cmbCariBerdasarkan
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
            With cmbCariBerdasarkan
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
            With cmbCariBerdasarkan
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
            With cmbCariBerdasarkan
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
            With cmbCariBerdasarkan
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
            With cmbCariBerdasarkan
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
            With cmbCariBerdasarkan
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
            With cmbCariBerdasarkan
                .AddItem FormManage.AdodcMain.Recordset.Fields(0).Name & " / Name", 0
                .AddItem FormManage.AdodcMain.Recordset.Fields(1).Name & " / Birthday", 1
                .AddItem FormManage.AdodcMain.Recordset.Fields(2).Name & " / Decsription", 2
                .ListIndex = 0
            End With
        ElseIf FORM_UTAMA.cmAgenda.FontBold = True Then
            With cmbCariBerdasarkan
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
            With cmbCariBerdasarkan
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

Private Sub cmBantuan_Click()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Kalimat = App.Path & "\bantuan\html\Cari.html"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Kalimat = App.Path & "\bantuan\html\Search.html"
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

Private Sub cmCari_Click()
    If TextKriteria.Text = "" Then
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            MsgBox "Silahkan isi kriteria yang ingin Anda cari!", vbExclamation + vbOKOnly, ""
        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
            MsgBox "Please fill in the criteria you want to search!", vbExclamation + vbOKOnly, ""
        End If
        TextKriteria.SetFocus
    Else
        If FORM_UTAMA.cmJejaringSosial.FontBold = True Then
            FormManage.AdodcMain.Refresh
                With FormManage.AdodcMain.Recordset
                    Select Case Me.cmbCariBerdasarkan.ListIndex
                    Case Is = 0
                        .Find "Nama_Jejaring = '" & TextKriteria.Text & "'"
                    Case Is = 1
                        .Find "Nama_Pengguna = '" & TextKriteria.Text & "'"
                    Case Is = 2
                        .Find "Alamat_Email = '" & TextKriteria.Text & "'"
                    Case Is = 3
                        .Find "Password = '" & TextKriteria.Text & "'"
                    Case Is = 4
                        .Find "URL = '" & TextKriteria.Text & "'"
                    Case Is = 5
                        .Find "Pemilik_Akun = '" & TextKriteria.Text & "'"
                    Case Is = 6
                        .Find "Tanggal = '" & TextKriteria.Text & "'"
                    Case Is = 7
                        .Find "Keterangan = '" & TextKriteria.Text & "'"
                    End Select
                    '=============================================================================
                    If .EOF Then
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            MsgBox ("File not found.."), vbExclamation + vbOKOnly, ""
                        End If
                    Else
                        Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                        Select Case FormPengaturan.cmbHasilPencarian.ListIndex
                            Case Is = 0
                                With FormHasilPencarian
                                    .TextDitemukan.Text = FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value
                                    .Show vbModal, Me
                                End With
                            Case Is = 1
                                MsgBox FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value, vbInformation + vbOKOnly, "Data Ditemukan!"
                            Case Is = 2
                                If FormPengaturan.CekTutupFormCAri.Value = Checked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                    Unload Me
                                ElseIf FormPengaturan.CekTutupFormCAri.Value = Unchecked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                End If
                            End Select
                            If FormPengaturan.CekTutupFormCAri.Value = Checked Then Unload Me
                                cmBatal.Caption = "&Tutup"
                    End If
                End With
        ElseIf FORM_UTAMA.cmElectronicMail.FontBold = True Then
            FormManage.AdodcMain.Refresh
                With FormManage.AdodcMain.Recordset
                    Select Case Me.cmbCariBerdasarkan.ListIndex
                    Case Is = 0
                        .Find "Nama_Server = '" & TextKriteria.Text & "'"
                    Case Is = 1
                        .Find "Nama_Pengguna = '" & TextKriteria.Text & "'"
                    Case Is = 2
                        .Find "Alamat_Email = '" & TextKriteria.Text & "'"
                    Case Is = 3
                        .Find "Password = '" & TextKriteria.Text & "'"
                    Case Is = 4
                        .Find "Pertanyaan_Rahasia = '" & TextKriteria.Text & "'"
                    Case Is = 5
                        .Find "Jawaban_Pertanyaan = '" & TextKriteria.Text & "'"
                    Case Is = 6
                        .Find "URL = '" & TextKriteria.Text & "'"
                    Case Is = 7
                        .Find "Pemilik_Akun = '" & TextKriteria.Text & "'"
                    Case Is = 8
                        .Find "Tanggal = '" & TextKriteria.Text & "'"
                    Case Is = 9
                        .Find "Keterangan = '" & TextKriteria.Text & "'"
                    End Select
                    '=============================================================================
                    If .EOF Then
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            MsgBox ("File not found.."), vbExclamation + vbOKOnly, ""
                        End If
                    Else
                        Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                        Select Case FormPengaturan.cmbHasilPencarian.ListIndex
                            Case Is = 0
                                With FormHasilPencarian
                                    .TextDitemukan.Text = FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(9).Name & " : " & FormManage.AdodcMain.Recordset.Fields(9).Value
                                    .Show vbModal, Me
                                End With
                            Case Is = 1
                                MsgBox FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(9).Name & " : " & FormManage.AdodcMain.Recordset.Fields(9).Value, vbInformation + vbOKOnly, "Data Ditemukan!"
                            Case Is = 2
                                If FormPengaturan.CekTutupFormCAri.Value = Checked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                    Unload Me
                                ElseIf FormPengaturan.CekTutupFormCAri.Value = Unchecked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                End If
                        End Select
                            If FormPengaturan.CekTutupFormCAri.Value = Checked Then Unload Me
                                cmBatal.Caption = "&Tutup"
                    End If
                End With
        ElseIf FORM_UTAMA.cmForumInternet.FontBold = True Then
            FormManage.AdodcMain.Refresh
                With FormManage.AdodcMain.Recordset
                    Select Case Me.cmbCariBerdasarkan.ListIndex
                    Case Is = 0
                        .Find "Nama_Forum = '" & TextKriteria.Text & "'"
                    Case Is = 1
                        .Find "Nama_Pengguna = '" & TextKriteria.Text & "'"
                    Case Is = 2
                        .Find "Alamat_Email = '" & TextKriteria.Text & "'"
                    Case Is = 3
                        .Find "Password = '" & TextKriteria.Text & "'"
                    Case Is = 4
                        .Find "Posisi = '" & TextKriteria.Text & "'"
                    Case Is = 5
                        .Find "NickName = '" & TextKriteria.Text & "'"
                    Case Is = 6
                        .Find "URL = '" & TextKriteria.Text & "'"
                    Case Is = 7
                        .Find "Tanggal = '" & TextKriteria.Text & "'"
                    Case Is = 8
                        .Find "Keterangan = '" & TextKriteria.Text & "'"
                    End Select
                    '=============================================================================
                    If .EOF Then
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            MsgBox ("File not found.."), vbExclamation + vbOKOnly, ""
                        End If
                    Else
                        Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                        Select Case FormPengaturan.cmbHasilPencarian.ListIndex
                            Case Is = 0
                                With FormHasilPencarian
                                    .TextDitemukan.Text = FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value
                                    .Show vbModal, Me
                                End With
                            Case Is = 1
                                MsgBox FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value, vbInformation + vbOKOnly, "Data Ditemukan!"
                            Case Is = 2
                                If FormPengaturan.CekTutupFormCAri.Value = Checked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                    Unload Me
                                ElseIf FormPengaturan.CekTutupFormCAri.Value = Unchecked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                End If
                        End Select
                            If FormPengaturan.CekTutupFormCAri.Value = Checked Then Unload Me
                                cmBatal.Caption = "&Tutup"
                    End If
                End With
        ElseIf FORM_UTAMA.cmFTP.FontBold = True Then
            FormManage.AdodcMain.Refresh
                With FormManage.AdodcMain.Recordset
                    Select Case Me.cmbCariBerdasarkan.ListIndex
                    Case Is = 0
                        .Find "Nama_Host = '" & TextKriteria.Text & "'"
                    Case Is = 1
                        .Find "Port = '" & TextKriteria.Text & "'"
                    Case Is = 2
                        .Find "Nama_Server = '" & TextKriteria.Text & "'"
                    Case Is = 3
                        .Find "Nama_Pengguna = '" & TextKriteria.Text & "'"
                    Case Is = 4
                        .Find "Alamat_Email = '" & TextKriteria.Text & "'"
                    Case Is = 5
                        .Find "Password = '" & TextKriteria.Text & "'"
                    Case Is = 6
                        .Find "Tanggal = '" & TextKriteria.Text & "'"
                    Case Is = 7
                        .Find "Keterangan = '" & TextKriteria.Text & "'"
                    End Select
                    '=============================================================================
                    If .EOF Then
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            MsgBox ("File not found.."), vbExclamation + vbOKOnly, ""
                        End If
                    Else
                        Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                        Select Case FormPengaturan.cmbHasilPencarian.ListIndex
                            Case Is = 0
                                With FormHasilPencarian
                                    .TextDitemukan.Text = FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value
                                    .Show vbModal, Me
                                End With
                            Case Is = 1
                                MsgBox FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value, vbInformation + vbOKOnly, "Data Ditemukan!"
                            Case Is = 2
                                If FormPengaturan.CekTutupFormCAri.Value = Checked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                    Unload Me
                                ElseIf FormPengaturan.CekTutupFormCAri.Value = Unchecked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                End If
                        End Select
                            If FormPengaturan.CekTutupFormCAri.Value = Checked Then Unload Me
                                cmBatal.Caption = "&Tutup"
                    End If
                End With
        ElseIf FORM_UTAMA.cmBlogging.FontBold = True Then
            FormManage.AdodcMain.Refresh
                With FormManage.AdodcMain.Recordset
                    Select Case Me.cmbCariBerdasarkan.ListIndex
                    Case Is = 0
                        .Find "Nama_Penyedia_Blog = '" & TextKriteria.Text & "'"
                    Case Is = 1
                        .Find "Nama_Pengguna = '" & TextKriteria.Text & "'"
                    Case Is = 2
                        .Find "E_Mail = '" & TextKriteria.Text & "'"
                    Case Is = 3
                        .Find "Password = '" & TextKriteria.Text & "'"
                    Case Is = 4
                        .Find "URL = '" & TextKriteria.Text & "'"
                    Case Is = 5
                        .Find "Tanggal = '" & TextKriteria.Text & "'"
                    Case Is = 6
                        .Find "Keterangan = '" & TextKriteria.Text & "'"
                    End Select
                    '=============================================================================
                    If .EOF Then
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            MsgBox ("File not found.."), vbExclamation + vbOKOnly, ""
                        End If
                    Else
                        Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                        Select Case FormPengaturan.cmbHasilPencarian.ListIndex
                            Case Is = 0
                                With FormHasilPencarian
                                    .TextDitemukan.Text = FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value
                                    .Show vbModal, Me
                                End With
                            Case Is = 1
                                MsgBox FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value, vbInformation + vbOKOnly, "Data Ditemukan!"
                            Case Is = 2
                                If FormPengaturan.CekTutupFormCAri.Value = Checked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                    Unload Me
                                ElseIf FormPengaturan.CekTutupFormCAri.Value = Unchecked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                End If
                        End Select
                            If FormPengaturan.CekTutupFormCAri.Value = Checked Then Unload Me
                                cmBatal.Caption = "&Tutup"
                    End If
                End With
        ElseIf FORM_UTAMA.cmIdentitasPribadi.FontBold = True Then
            FormManage.AdodcMain.Refresh
                With FormManage.AdodcMain.Recordset
                    Select Case Me.cmbCariBerdasarkan.ListIndex
                    Case Is = 0
                        .Find "Nama_Lengkap = '" & TextKriteria.Text & "'"
                    Case Is = 1
                        .Find "Nama_Panggilan = '" & TextKriteria.Text & "'"
                    Case Is = 2
                        .Find "TempatLahir = '" & TextKriteria.Text & "'"
                    Case Is = 3
                        .Find "TanggalLahir = '" & TextKriteria.Text & "'"
                    Case Is = 4
                        .Find "BulanLahir = '" & TextKriteria.Text & "'"
                    Case Is = 5
                        .Find "TahunLahir = '" & TextKriteria.Text & "'"
                    Case Is = 6
                        .Find "Jenis_Kelamin = '" & TextKriteria.Text & "'"
                    Case Is = 7
                        .Find "Agama = '" & TextKriteria.Text & "'"
                    Case Is = 8
                        .Find "Golongan_Darah = '" & TextKriteria.Text & "'"
                    Case Is = 9
                        .Find "Pekerjaan = '" & TextKriteria.Text & "'"
                    Case Is = 10
                        .Find "Alamat_Rumah = '" & TextKriteria.Text & "'"
                    Case Is = 11
                        .Find "E_Mail = '" & TextKriteria.Text & "'"
                    Case Is = 12
                        .Find "Website = '" & TextKriteria.Text & "'"
                    Case Is = 13
                        .Find "Nomor_Telepon = '" & TextKriteria.Text & "'"
                    Case Is = 14
                        .Find "Kota_Asal = '" & TextKriteria.Text & "'"
                    Case Is = 15
                        .Find "Kota_Sekarang = '" & TextKriteria.Text & "'"
                    Case Is = 16
                        .Find "Kode_Pos = '" & TextKriteria.Text & "'"
                    Case Is = 17
                        .Find "Provinsi = '" & TextKriteria.Text & "'"
                    Case Is = 18
                        .Find "Kewarganegaraan = '" & TextKriteria.Text & "'"
                    Case Is = 19
                        .Find "Status_Pendidikan = '" & TextKriteria.Text & "'"
                    Case Is = 20
                        .Find "Status_Hubungan = '" & TextKriteria.Text & "'"
                    Case Is = 21
                        .Find "Hobby = '" & TextKriteria.Text & "'"
                    Case Is = 22
                        .Find "Keterangan = '" & TextKriteria.Text & "'"
                    End Select
                    '=============================================================================
                    If .EOF Then
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            MsgBox ("File not found.."), vbExclamation + vbOKOnly, ""
                        End If
                    Else
                        Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                        Select Case FormPengaturan.cmbHasilPencarian.ListIndex
                            Case Is = 0
                                With FormHasilPencarian
                                    .TextDitemukan.Text = FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(9).Name & " : " & FormManage.AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(10).Name & " : " & FormManage.AdodcMain.Recordset.Fields(10).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(11).Name & " : " & FormManage.AdodcMain.Recordset.Fields(11).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(12).Name & " : " & FormManage.AdodcMain.Recordset.Fields(12).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(13).Name & " : " & FormManage.AdodcMain.Recordset.Fields(13).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(14).Name & " : " & FormManage.AdodcMain.Recordset.Fields(14).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(15).Name & " : " & FormManage.AdodcMain.Recordset.Fields(15).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(16).Name & " : " & FormManage.AdodcMain.Recordset.Fields(16).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(17).Name & " : " & FormManage.AdodcMain.Recordset.Fields(17).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(18).Name & " : " & FormManage.AdodcMain.Recordset.Fields(18).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(19).Name & " : " & FormManage.AdodcMain.Recordset.Fields(19).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(20).Name & " : " & FormManage.AdodcMain.Recordset.Fields(20).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(21).Name & " : " & FormManage.AdodcMain.Recordset.Fields(21).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(22).Name & " : " & FormManage.AdodcMain.Recordset.Fields(22).Value
                                    .Show vbModal, Me
                                End With
                            Case Is = 1
                                MsgBox FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(9).Name & " : " & FormManage.AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(10).Name & " : " & FormManage.AdodcMain.Recordset.Fields(10).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(11).Name & " : " & FormManage.AdodcMain.Recordset.Fields(11).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(12).Name & " : " & FormManage.AdodcMain.Recordset.Fields(12).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(13).Name & " : " & FormManage.AdodcMain.Recordset.Fields(13).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(14).Name & " : " & FormManage.AdodcMain.Recordset.Fields(14).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(15).Name & " : " & FormManage.AdodcMain.Recordset.Fields(15).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(16).Name & " : " & FormManage.AdodcMain.Recordset.Fields(16).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(17).Name & " : " & FormManage.AdodcMain.Recordset.Fields(17).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(18).Name & " : " & FormManage.AdodcMain.Recordset.Fields(18).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(19).Name & " : " & FormManage.AdodcMain.Recordset.Fields(19).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(20).Name & " : " & FormManage.AdodcMain.Recordset.Fields(20).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(21).Name & " : " & FormManage.AdodcMain.Recordset.Fields(21).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(22).Name & " : " & FormManage.AdodcMain.Recordset.Fields(22).Value, vbInformation + vbOKOnly, "Data Ditemukan!"
                            Case Is = 2
                                If FormPengaturan.CekTutupFormCAri.Value = Checked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                    Unload Me
                                ElseIf FormPengaturan.CekTutupFormCAri.Value = Unchecked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                End If
                        End Select
                            If FormPengaturan.CekTutupFormCAri.Value = Checked Then Unload Me
                                cmBatal.Caption = "&Tutup"
                    End If
                End With
        ElseIf FORM_UTAMA.cmAgenda.FontBold = True Then
            FormManage.AdodcMain.Refresh
                With FormManage.AdodcMain.Recordset
                    Select Case Me.cmbCariBerdasarkan.ListIndex
                    Case Is = 0
                        .Find "Kode_Agenda = '" & TextKriteria.Text & "'"
                    Case Is = 1
                        .Find "Nama_Agenda = '" & TextKriteria.Text & "'"
                    Case Is = 2
                        .Find "Tema = '" & TextKriteria.Text & "'"
                    Case Is = 3
                        .Find "Tanggal = '" & TextKriteria.Text & "'"
                    Case Is = 4
                        .Find "Waktu_Mulai = '" & TextKriteria.Text & "'"
                    Case Is = 5
                        .Find "Waktu_Akhir = '" & TextKriteria.Text & "'"
                    Case Is = 6
                        .Find "Tempat = '" & TextKriteria.Text & "'"
                    Case Is = 7
                        .Find "Keterangan_Lain = '" & TextKriteria.Text & "'"
                    End Select
                    '=============================================================================
                    If .EOF Then
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            MsgBox ("File not found.."), vbExclamation + vbOKOnly, ""
                        End If
                    Else
                        Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                        Select Case FormPengaturan.cmbHasilPencarian.ListIndex
                            Case Is = 0
                                With FormHasilPencarian
                                    .TextDitemukan.Text = FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value
                                    .Show vbModal, Me
                                End With
                            Case Is = 1
                                MsgBox FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value, vbInformation + vbOKOnly, "Data Ditemukan!"
                            Case Is = 2
                                If FormPengaturan.CekTutupFormCAri.Value = Checked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                    Unload Me
                                ElseIf FormPengaturan.CekTutupFormCAri.Value = Unchecked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                End If
                        End Select
                            If FormPengaturan.CekTutupFormCAri.Value = Checked Then Unload Me
                                cmBatal.Caption = "&Tutup"
                    End If
                End With
        ElseIf FORM_UTAMA.cmBukuAlamat.FontBold = True Then
            FormManage.AdodcMain.Refresh
                With FormManage.AdodcMain.Recordset
                    Select Case Me.cmbCariBerdasarkan.ListIndex
                    Case Is = 0
                        .Find "Nama_Kontak = '" & TextKriteria.Text & "'"
                    Case Is = 1
                        .Find "Nama_Panggilan = '" & TextKriteria.Text & "'"
                    Case Is = 2
                        .Find "Nomor_Telepon_Pribadi = '" & TextKriteria.Text & "'"
                    Case Is = 3
                        .Find "Nomor_Telepon_Rumah = '" & TextKriteria.Text & "'"
                    Case Is = 4
                        .Find "Nomor_Telepon_Kantor = '" & TextKriteria.Text & "'"
                    Case Is = 5
                        .Find "Fax = '" & TextKriteria.Text & "'"
                    Case Is = 6
                        .Find "Alamat_EMail = '" & TextKriteria.Text & "'"
                    Case Is = 7
                        .Find "Website = '" & TextKriteria.Text & "'"
                    Case Is = 8
                        .Find "ZIP_Postal_Code = '" & TextKriteria.Text & "'"
                    Case Is = 9
                        .Find "Alamat_Rumah = '" & TextKriteria.Text & "'"
                    Case Is = 10
                        .Find "Keterangan = '" & TextKriteria.Text & "'"
                    End Select
                    '=============================================================================
                    If .EOF Then
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            MsgBox ("File not found.."), vbExclamation + vbOKOnly, ""
                        End If
                    Else
                        Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                        Select Case FormPengaturan.cmbHasilPencarian.ListIndex
                            Case Is = 0
                                With FormHasilPencarian
                                    .TextDitemukan.Text = FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(9).Name & " : " & FormManage.AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(10).Name & " : " & FormManage.AdodcMain.Recordset.Fields(10).Value
                                    .Show vbModal, Me
                                End With
                            Case Is = 1
                                MsgBox FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(9).Name & " : " & FormManage.AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(10).Name & " : " & FormManage.AdodcMain.Recordset.Fields(10).Value, vbInformation + vbOKOnly, "Data Ditemukan!"
                            Case Is = 2
                                If FormPengaturan.CekTutupFormCAri.Value = Checked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                    Unload Me
                                ElseIf FormPengaturan.CekTutupFormCAri.Value = Unchecked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                End If
                        End Select
                            If FormPengaturan.CekTutupFormCAri.Value = Checked Then Unload Me
                                cmBatal.Caption = "&Tutup"
                    End If
                End With
        ElseIf FORM_UTAMA.cmUlangTahun.FontBold = True Then
            FormManage.AdodcMain.Refresh
                With FormManage.AdodcMain.Recordset
                    Select Case Me.cmbCariBerdasarkan.ListIndex
                    Case Is = 0
                        .Find "Nama = '" & TextKriteria.Text & "'"
                    Case Is = 1
                        .Find "TTL = '" & TextKriteria.Text & "'"
                    Case Is = 2
                        .Find "Keterangan = '" & TextKriteria.Text & "'"
                    End Select
                    '=============================================================================
                    If .EOF Then
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            MsgBox ("File not found.."), vbExclamation + vbOKOnly, ""
                        End If
                    Else
                        Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                        Select Case FormPengaturan.cmbHasilPencarian.ListIndex
                            Case Is = 0
                                With FormHasilPencarian
                                    .TextDitemukan.Text = FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value
                                    .Show vbModal, Me
                                End With
                            Case Is = 1
                                MsgBox FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value, vbInformation + vbOKOnly, "Data Ditemukan!"
                            Case Is = 2
                                If FormPengaturan.CekTutupFormCAri.Value = Checked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                    Unload Me
                                ElseIf FormPengaturan.CekTutupFormCAri.Value = Unchecked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                End If
                        End Select
                            If FormPengaturan.CekTutupFormCAri.Value = Checked Then Unload Me
                                cmBatal.Caption = "&Tutup"
                    End If
                End With
        ElseIf FORM_UTAMA.cmRegistrasiSoftware.FontBold = True Then
            FormManage.AdodcMain.Refresh
                With FormManage.AdodcMain.Recordset
                    Select Case Me.cmbCariBerdasarkan.ListIndex
                    Case Is = 0
                        .Find "Nama_Software = '" & TextKriteria.Text & "'"
                    Case Is = 1
                        .Find "Kategori = '" & TextKriteria.Text & "'"
                    Case Is = 2
                        .Find "Developer = '" & TextKriteria.Text & "'"
                    Case Is = 3
                        .Find "Username = '" & TextKriteria.Text & "'"
                    Case Is = 4
                        .Find "Serial_Key = '" & TextKriteria.Text & "'"
                    Case Is = 5
                        .Find "Jenis_Lisensi = '" & TextKriteria.Text & "'"
                    Case Is = 6
                        .Find "Keterangan = '" & TextKriteria.Text & "'"
                    End Select
                    '=============================================================================
                    If .EOF Then
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            MsgBox ("File not found.."), vbExclamation + vbOKOnly, ""
                        End If
                    Else
                        Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                        Select Case FormPengaturan.cmbHasilPencarian.ListIndex
                            Case Is = 0
                                With FormHasilPencarian
                                    .TextDitemukan.Text = FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                                            FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value
                                    .Show vbModal, Me
                                End With
                            Case Is = 1
                                MsgBox FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                        FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value, vbInformation + vbOKOnly, "Data Ditemukan!"
                            Case Is = 2
                                If FormPengaturan.CekTutupFormCAri.Value = Checked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                    Unload Me
                                ElseIf FormPengaturan.CekTutupFormCAri.Value = Unchecked Then
                                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                                End If
                        End Select
                            If FormPengaturan.CekTutupFormCAri.Value = Checked Then Unload Me
                                cmBatal.Caption = "&Tutup"
                    End If
                End With
        End If
    End If
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

Sub PENGATURAN_BAHASA()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        With Me
            .cmCari.Caption = "&Cari"
            .Label3.Caption = "Dengan Kriteria"
            .Label1.Caption = "Cari berdasarkan"
            .cmBatal.Caption = "&Batal"
            .cmBantuan.Caption = "&Bantuan"
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With Me
            .cmCari.Caption = "&Search"
            .Label3.Caption = "with Criteria"
            .Label1.Caption = "Search by"
            .cmBatal.Caption = "&Cancel"
            .cmBantuan.Caption = "&Help"
        End With
    End If
End Sub
