VERSION 5.00
Object = "{02353968-C1C9-4E0A-88D3-18759BDC60FE}#1.0#0"; "AeroSuite.ocx"
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormEkstrakDataKeText 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ekstrak Data Ke Text"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormEkstrakDataKeText.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer TimerProgress 
      Interval        =   70
      Left            =   3840
      Top             =   3000
   End
   Begin AeroSuite.AeroProgressBar Progress 
      Height          =   270
      Left            =   1440
      Top             =   2160
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   476
   End
   Begin VB.TextBox textExtract 
      Appearance      =   0  'Flat
      Height          =   4815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FormEkstrakDataKeText.frx":06C2
      Top             =   120
      Width           =   6015
   End
   Begin Dacara_dcButton.dcButton cmSimpan 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   10591645
      ButtonStyle     =   2
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
      PicDown         =   "FormEkstrakDataKeText.frx":06C8
      PicHot          =   "FormEkstrakDataKeText.frx":0A1A
      PicNormal       =   "FormEkstrakDataKeText.frx":0D6C
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmSalin 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   10591645
      ButtonStyle     =   2
      Caption         =   "&Salin"
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
      PicDown         =   "FormEkstrakDataKeText.frx":10BE
      PicHot          =   "FormEkstrakDataKeText.frx":3870
      PicNormal       =   "FormEkstrakDataKeText.frx":6022
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmTutup 
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      PicAlign        =   5
      PicDown         =   "FormEkstrakDataKeText.frx":87D4
      PicHot          =   "FormEkstrakDataKeText.frx":8C26
      PicNormal       =   "FormEkstrakDataKeText.frx":9078
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmPrint 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   10591645
      ButtonStyle     =   2
      Caption         =   "&Cetak"
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
      PicDown         =   "FormEkstrakDataKeText.frx":94CA
      PicHot          =   "FormEkstrakDataKeText.frx":109CC
      PicNormal       =   "FormEkstrakDataKeText.frx":17ECE
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmRefresh 
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   10591645
      ButtonStyle     =   2
      Caption         =   "&Refresh"
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
      PicDown         =   "FormEkstrakDataKeText.frx":1F3D0
      PicHot          =   "FormEkstrakDataKeText.frx":2B437
      PicNormal       =   "FormEkstrakDataKeText.frx":3749E
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
Attribute VB_Name = "FormEkstrakDataKeText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
If FormPengaturan.cmbBahasa.ListIndex = 0 Then
    Me.Caption = "Extract Data ke Text"
    cmSimpan.Caption = "&Simpan"
    cmSalin.Caption = "&Salin"
    cmPrint.Caption = "&Cetak"
    cmTutup.Caption = "&Batal"
Else
    Me.Caption = "Extract Data to Text"
    cmSimpan.Caption = "&Save"
    cmSalin.Caption = "&Copy"
    cmPrint.Caption = "&Print"
    cmTutup.Caption = "&Cancel"
End If
With textExtract
    .Enabled = False
    .BackColor = Silver
End With
    For Each Objek In Me
        If TypeName(Objek) = "dcButton" Then Objek.Enabled = False
    Next
    cmTutup.Enabled = True
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

Private Sub cmPrint_Click()
On Error GoTo ErrHandler
Dim BeginPage, EndPage, NumCopies, i
    CommonDialog1.CancelError = True
    CommonDialog1.ShowPrinter
    BeginPage = CommonDialog1.FromPage
    EndPage = CommonDialog1.ToPage
    NumCopies = CommonDialog1.Copies
    For i = 1 To NumCopies
        Printer.Print textExtract.Text
    Next i
Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub cmRefresh_Click()
    Form_Load
    Progress.Visible = True
    TimerProgress.Enabled = True
    TimerProgress_Timer
End Sub

Private Sub cmSalin_Click()
Clipboard.Clear
With textExtract
    .SetFocus
    .SelStart = 0
    .SelLength = Len(textExtract)
    Clipboard.SetText textExtract.Text
End With
End Sub

Private Sub cmSimpan_Click()
On Error GoTo ErrorHandler
If textExtract.Text = "" Then
    MsgBox "Stop!", vbCritical + vbOKOnly, "Stop!"
Else
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
    Print #iFile, textExtract.Text

    SaveFileFromTB = True
ErrorHandler:
    Close #iFile
End If
End Sub

Private Sub cmTutup_Click()
TimerProgress.Enabled = False
If FormPengaturan.cmbBahasa.ListIndex = 0 Then
    If textExtract.Enabled = False Then
        Pesan = MsgBox("Data sedang diekstrak. Yakin ingin dibatalkan?", vbQuestion + vbYesNo, "Batal?")
        If Pesan = vbYes Then
            Unload Me
        Else
            TimerProgress.Enabled = True
        End If
    Else
        Pesan = MsgBox("Data telah diekstrak. Yakin ingin ditutup?", vbQuestion + vbYesNo, "Batal?")
        If Pesan = vbYes Then
            Unload Me
        End If
    End If
Else
    If textExtract.Enabled = False Then
        Pesan = MsgBox("Data is extracting. Are you sure to cancelled?", vbQuestion + vbYesNo, "Batal?")
        If Pesan = vbYes Then
            Unload Me
        Else
            TimerProgress.Enabled = True
        End If
    Else
            Pesan = MsgBox("Data is extracted. Are you sure to close window?", vbQuestion + vbYesNo, "Batal?")
        If Pesan = vbYes Then
            Unload Me
        End If
    End If
End If
End Sub


Private Sub Form_Load()
    AturKontrol
    Progress.Value = 0
    PENGATURAN_WARNA
End Sub

Private Sub TimerProgress_Timer()
With FormBuatAkunBaru
    Progress.Value = Progress.Value + 1
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            Me.Caption = "Extrak Data ke Text (" & Progress.Value & " %)"
        Else
            Me.Caption = "Extrak Data to Text (" & Progress.Value & " %)"
        End If
    If Progress.Value = 3 Then
        textExtract.Text = .Label24.Caption
    ElseIf Progress.Value = 5 Then
        textExtract.Text = .Label24.Caption & " : "
    ElseIf Progress.Value = 10 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text
    ElseIf Progress.Value = 11 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption
    ElseIf Progress.Value = 13 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text
    ElseIf Progress.Value = 17 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption
    ElseIf Progress.Value = 21 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : "
    ElseIf Progress.Value = 23 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text
    ElseIf Progress.Value = 26 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption
    ElseIf Progress.Value = 27 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : "
    ElseIf Progress.Value = 35 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text
    ElseIf Progress.Value = 42 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption
    ElseIf Progress.Value = 45 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text
    ElseIf Progress.Value = 49 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text
    ElseIf Progress.Value = 51 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text
    ElseIf Progress.Value = 55 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text
    ElseIf Progress.Value = 58 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : "
    ElseIf Progress.Value = 59 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text
    ElseIf Progress.Value = 60 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption
    ElseIf Progress.Value = 68 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text
    ElseIf Progress.Value = 70 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : "
    ElseIf Progress.Value = 73 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text
    ElseIf Progress.Value = 74 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text & vbCrLf & _
                            .Label12.Caption
    ElseIf Progress.Value = 76 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text & vbCrLf & _
                            .Label12.Caption & " : "
    ElseIf Progress.Value = 77 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text & vbCrLf & _
                            .Label12.Caption & " : " & .textAlamat.Text
    ElseIf Progress.Value = 80 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text & vbCrLf & _
                            .Label12.Caption & " : " & .textAlamat.Text & vbCrLf & _
                            .Label15.Caption & " : "
    ElseIf Progress.Value = 81 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text & vbCrLf & _
                            .Label12.Caption & " : " & .textAlamat.Text & vbCrLf & _
                            .Label15.Caption & " : " & .textNomorTelepon.Text
    ElseIf Progress.Value = 82 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text & vbCrLf & _
                            .Label12.Caption & " : " & .textAlamat.Text & vbCrLf & _
                            .Label15.Caption & " : " & .textNomorTelepon.Text & vbCrLf & _
                            .Label17.Caption
    ElseIf Progress.Value = 84 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text & vbCrLf & _
                            .Label12.Caption & " : " & .textAlamat.Text & vbCrLf & _
                            .Label15.Caption & " : " & .textNomorTelepon.Text & vbCrLf & _
                            .Label17.Caption & " : " & .textAlamatEmail.Text
    ElseIf Progress.Value = 86 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text & vbCrLf & _
                            .Label12.Caption & " : " & .textAlamat.Text & vbCrLf & _
                            .Label15.Caption & " : " & .textNomorTelepon.Text & vbCrLf & _
                            .Label17.Caption & " : " & .textAlamatEmail.Text & vbCrLf & _
                            .Label19.Caption & " : "
    ElseIf Progress.Value = 88 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text & vbCrLf & _
                            .Label12.Caption & " : " & .textAlamat.Text & vbCrLf & _
                            .Label15.Caption & " : " & .textNomorTelepon.Text & vbCrLf & _
                            .Label17.Caption & " : " & .textAlamatEmail.Text & vbCrLf & _
                            .Label19.Caption & " : " & .cmbAlamatWebsite.Text
    ElseIf Progress.Value = 90 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text & vbCrLf & _
                            .Label12.Caption & " : " & .textAlamat.Text & vbCrLf & _
                            .Label15.Caption & " : " & .textNomorTelepon.Text & vbCrLf & _
                            .Label17.Caption & " : " & .textAlamatEmail.Text & vbCrLf & _
                            .Label19.Caption & " : " & .cmbAlamatWebsite.Text & .textAlamatWebsite.Text
    ElseIf Progress.Value = 91 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text & vbCrLf & _
                            .Label12.Caption & " : " & .textAlamat.Text & vbCrLf & _
                            .Label15.Caption & " : " & .textNomorTelepon.Text & vbCrLf & _
                            .Label17.Caption & " : " & .textAlamatEmail.Text & vbCrLf & _
                            .Label19.Caption & " : " & .cmbAlamatWebsite.Text & .textAlamatWebsite.Text & vbCrLf & _
                            .Label20.Caption & " : "
    ElseIf Progress.Value = 92 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text & vbCrLf & _
                            .Label12.Caption & " : " & .textAlamat.Text & vbCrLf & _
                            .Label15.Caption & " : " & .textNomorTelepon.Text & vbCrLf & _
                            .Label17.Caption & " : " & .textAlamatEmail.Text & vbCrLf & _
                            .Label19.Caption & " : " & .cmbAlamatWebsite.Text & .textAlamatWebsite.Text & vbCrLf & _
                            .Label20.Caption & " : " & .textStatusAktivitas.Text
    ElseIf Progress.Value = 93 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text & vbCrLf & _
                            .Label12.Caption & " : " & .textAlamat.Text & vbCrLf & _
                            .Label15.Caption & " : " & .textNomorTelepon.Text & vbCrLf & _
                            .Label17.Caption & " : " & .textAlamatEmail.Text & vbCrLf & _
                            .Label19.Caption & " : " & .cmbAlamatWebsite.Text & .textAlamatWebsite.Text & vbCrLf & _
                            .Label20.Caption & " : " & .textStatusAktivitas.Text & vbCrLf & _
                            .Label22.Caption & " : " & .textStatusHubungan.Text
    ElseIf Progress.Value = 96 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text & vbCrLf & _
                            .Label12.Caption & " : " & .textAlamat.Text & vbCrLf & _
                            .Label15.Caption & " : " & .textNomorTelepon.Text & vbCrLf & _
                            .Label17.Caption & " : " & .textAlamatEmail.Text & vbCrLf & _
                            .Label19.Caption & " : " & .cmbAlamatWebsite.Text & .textAlamatWebsite.Text & vbCrLf & _
                            .Label20.Caption & " : " & .textStatusAktivitas.Text & vbCrLf & _
                            .Label22.Caption & " : " & .textStatusHubungan.Text & vbCrLf & _
                            .Label35.Caption & " : " & .cmbPertanyaanRahasia.Text
    ElseIf Progress.Value = 99 Then
        textExtract.Text = .Label24.Caption & " : " & .textNamaPengguna.Text & vbCrLf & _
                            .Label26.Caption & " : " & .textPasswordBaru.Text & vbCrLf & _
                            .Label29.Caption & " : " & .textKonfirmasiPassword.Text & vbCrLf & _
                            .Label1.Caption & " : " & .textNamaAsli.Text & vbCrLf & _
                            .Label3.Caption & " : " & .textTempat.Text & ", " & .textTanggal.Text & " - " & .textBulan.Text & " - " & .textTahun.Text & vbCrLf & _
                            .Label7.Caption & " : " & .cmbJenisKelamin.Text & vbCrLf & _
                            .Label10.Caption & " : " & .cmbAgama.Text & vbCrLf & _
                            .Label9.Caption & " : " & .textHobby.Text & vbCrLf & _
                            .Label12.Caption & " : " & .textAlamat.Text & vbCrLf & _
                            .Label15.Caption & " : " & .textNomorTelepon.Text & vbCrLf & _
                            .Label17.Caption & " : " & .textAlamatEmail.Text & vbCrLf & _
                            .Label19.Caption & " : " & .cmbAlamatWebsite.Text & .textAlamatWebsite.Text & vbCrLf & _
                            .Label20.Caption & " : " & .textStatusAktivitas.Text & vbCrLf & _
                            .Label22.Caption & " : " & .textStatusHubungan.Text & vbCrLf & _
                            .Label35.Caption & " : " & .cmbPertanyaanRahasia.Text & vbCrLf & _
                            .Label33.Caption & " : " & .textJawaban.Text
    ElseIf Progress.Value = 100 Then
        TimerProgress.Enabled = False
        Progress.Visible = False
        With textExtract
            .Enabled = True
            .BackColor = vbWhite
        End With
        For Each Objek In Me
            If TypeName(Objek) = "dcButton" Then Objek.Enabled = True
        Next
        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
            cmTutup.Caption = "&Tutup"
            Me.Caption = "Extrak Data ke Text"
        Else
            cmTutup.Caption = "&Close"
            Me.Caption = "Extrak Data to Text"
        End If
    End If
End With
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
