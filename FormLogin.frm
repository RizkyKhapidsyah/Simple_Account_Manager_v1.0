VERSION 5.00
Object = "{02353968-C1C9-4E0A-88D3-18759BDC60FE}#1.0#0"; "AeroSuite.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   3945
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPText textPassword 
      Height          =   330
      Left            =   1320
      TabIndex        =   6
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      Text            =   "XPText1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PasswordChar    =   "*"
   End
   Begin XPControls.XPText textPengguna 
      Height          =   330
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      Text            =   "XPText1"
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
   Begin AeroSuite.AeroProgressBar Progress 
      Height          =   270
      Left            =   120
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   476
   End
   Begin VB.Timer TimerProgress 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2040
      Top             =   960
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1800
      Top             =   2400
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
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   720
      Top             =   720
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   7
      Bmp:1           =   "FormLogin.frx":058A
      Mask:1          =   16777215
      Key:1           =   "#menuBuatAkunBaru"
      Bmp:2           =   "FormLogin.frx":08DC
      Key:2           =   "#menuIngatSaya"
      Bmp:3           =   "FormLogin.frx":1644
      Mask:3          =   16777215
      Key:3           =   "#menuKV"
      Bmp:4           =   "FormLogin.frx":1996
      Key:4           =   "#menuKeluar"
      Bmp:5           =   "FormLogin.frx":1DBE
      Mask:5          =   16777215
      Key:5           =   "#menuPusatBantuan"
      Bmp:6           =   "FormLogin.frx":2A10
      Key:6           =   "#menuFAQ"
      Bmp:7           =   "FormLogin.frx":3778
      Key:7           =   "#menuKW"
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
   Begin Dacara_dcButton.dcButton cmLogin 
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   12230304
      ButtonStyle     =   3
      Caption         =   "&Login"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   5
      PicDown         =   "FormLogin.frx":3BA0
      PicHot          =   "FormLogin.frx":3FF2
      PicNormal       =   "FormLogin.frx":4444
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
   End
   Begin XPEngine.XPControl XP_Engine 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "--RikySoft--"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   1320
      TabIndex        =   8
      Top             =   780
      Width           =   585
   End
   Begin VB.Label LabelLogin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loging..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1440
      TabIndex        =   7
      Top             =   1095
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1200
      TabIndex        =   3
      Top             =   525
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   525
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pengguna"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   795
   End
   Begin VB.Menu MenuMenu 
      Caption         =   "Menu"
      Begin VB.Menu menuBuatAkunBaru 
         Caption         =   "Buat Akun Baru"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu menuIngatSaya 
         Caption         =   "Ingat Saya !"
         Begin VB.Menu menuIHNP 
            Caption         =   "Ingat Hanya Nama Pengguna"
         End
         Begin VB.Menu menuIHP 
            Caption         =   "Ingat Hanya Password"
         End
         Begin VB.Menu menuIK 
            Caption         =   "Ingat Keduanya"
         End
         Begin VB.Menu sep1 
            Caption         =   "-"
         End
         Begin VB.Menu menuJIA 
            Caption         =   "Jangan Ingat Apapun!"
         End
      End
      Begin VB.Menu menuKV 
         Caption         =   "Keyboard Virtual"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu menuKB 
         Caption         =   "Karakter Bintang"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu menuKeluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu menuBantuan 
      Caption         =   "Bantuan"
      Begin VB.Menu menuPusatBantuan 
         Caption         =   "Pusat Bantuan"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu menuKW 
         Caption         =   "Kunjungi Website"
      End
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
NyambunggUtama
    With Adodc1
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tbLogin"
        .Refresh
    End With
    textPengguna.Text = ""
    textPassword.Text = ""
    If Adodc1.Recordset.RecordCount = 0 Then FormBuatAkunBaru.Show vbModal, Me
    Progress.Value = 0
    DisableCloseBtn Me
    AmbilPengaturanDariREGISTRY
    If menuKB.Checked = True Then
        textPassword.PasswordChar = "*"
    ElseIf menuKB.Checked = False Then
        textPassword.PasswordChar = ""
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
Sub LOGIN()
    If RS.State = 1 Then RS.Close
        X = "select * from tbLogin where Nama_Pengguna= '" & textPengguna.Text & "' And Password = '" & textPassword.Text & "'"
        RS.Open X, CN, 3, 3
        If Not RS.EOF Then
            Dim CN_FormUtama As New ADODB.Connection
            If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
                CN_FormUtama.CursorLocation = adUseClient
                CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & textPengguna.Text & "\data.rdb;Persist Security Info=False"
                With FORM_UTAMA
                    .ADODC_UTAMA.ConnectionString = CN_FormUtama.ConnectionString
                    .ADODC_UTAMA.RecordSource = "Select * From tbJejaringSosial Order by Nama_Jejaring Asc;"
                    .ADODC_UTAMA.Refresh
                    .AdodcDataLogin.ConnectionString = CN_FormUtama.ConnectionString
                    .AdodcDataLogin.RecordSource = "Select * From tbDataLogin"
                    .AdodcDataLogin.Refresh
                    .Caption = "Simple Accounts Manager - " & .AdodcDataLogin.Recordset.Fields(2).Value
                    .StatusBawah.Panels.Item(2).Text = textPengguna.Text
                    .Show
                End With
                SimpanPengaturanKeREGISTRY
            Unload Me
        Else
            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                MsgBox "Maaf, Nama Pengguna atau Password Anda Salah atau Anda Tidak Terdaftar!" & vbCrLf & _
                        "Silahkan periksa kembali!", vbCritical + vbOKOnly, "Salah!"
            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                MsgBox "Sorry, User Name or Password is Wrong or You're Not Registered!" & vbCrLf & _
                        "Please check it!", vbCritical + vbOKOnly, "Wrong!"
            End If
                textPengguna.SetFocus
        End If
End Sub
Sub SimpanPengaturanKeREGISTRY()
    SaveSetting "rssamv1.0", "login", menuIHNP.Name, menuIHNP.Checked
    SaveSetting "rssamv1.0", "login", menuIHP.Name, menuIHP.Checked
    SaveSetting "rssamv1.0", "login", menuIK.Name, menuIK.Checked
    SaveSetting "rssamv1.0", "login", menuJIA.Name, menuJIA.Checked
    
    If menuIHNP.Checked = True Then
        SaveSetting "rssamv1.0", "login", textPengguna.Name, textPengguna.Text
    ElseIf menuIHNP.Checked = False Then
        SaveSetting "rssamv1.0", "login", textPengguna.Name, ""
    End If
    If menuIHP.Checked = True Then
        SaveSetting "rssamv1.0", "login", textPassword.Name, textPassword.Text
    ElseIf menuIHP.Checked = False Then
        SaveSetting "rssamv1.0", "login", textPassword.Name, ""
    End If
    If menuIK.Checked = True Then
        SaveSetting "rssamv1.0", "login", textPengguna.Name, textPengguna.Text
        SaveSetting "rssamv1.0", "login", textPassword.Name, textPassword.Text
    ElseIf menuIK.Checked = False Then
        SaveSetting "rssamv1.0", "login", textPengguna.Name, ""
        SaveSetting "rssamv1.0", "login", textPassword.Name, ""
    End If
    If menuJIA.Checked = True Then
        SaveSetting "rssamv1.0", "login", textPengguna.Name, ""
        SaveSetting "rssamv1.0", "login", textPassword.Name, ""
    End If
    SaveSetting "rssamv1.0", "login", menuKB.Name, menuKB.Checked
End Sub
Sub AmbilPengaturanDariREGISTRY()
    menuIHNP.Checked = GetSetting("rssamv1.0", "login", menuIHNP.Name, menuIHNP.Checked)
    menuIHP.Checked = GetSetting("rssamv1.0", "login", menuIHP.Name, menuIHP.Checked)
    menuIK.Checked = GetSetting("rssamv1.0", "login", menuIK.Name, menuIK.Checked)
    menuJIA.Checked = GetSetting("rssamv1.0", "login", menuJIA.Name, menuJIA.Checked)
    textPengguna.Text = GetSetting("rssamv1.0", "login", textPengguna.Name, textPengguna.Text)
    textPassword.Text = GetSetting("rssamv1.0", "login", textPassword.Name, textPassword.Text)
    menuKB.Checked = GetSetting("rssamv1.0", "login", menuKB.Name, menuKB.Checked)
End Sub

Private Sub cmLogin_Click()
If textPengguna.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Nama Pengguna!", vbExclamation + vbOKOnly, "Nama Pengguna?"
        textPengguna.SetFocus
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Put in UserName please!", vbExclamation + vbOKOnly, "User Name?"
        textPengguna.SetFocus
    End If
ElseIf textPassword.Text = "" Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Silahkan isi Password Anda!", vbExclamation + vbOKOnly, "Nama Pengguna?"
        textPassword.SetFocus
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Put in Your Password please!", vbExclamation + vbOKOnly, "User Name?"
        textPassword.SetFocus
    End If
Else
    With Progress
        .Visible = True
        .Value = 0
    End With
    LabelLogin.Visible = True
    cmLogin.Enabled = False
    MenuMenu.Enabled = False
    menuBantuan.Enabled = False
    With Me
        .textPengguna.BackColor = Silver
        .textPassword.BackColor = Silver
    End With
    TimerProgress.Enabled = True
    Me.SetFocus
End If
End Sub

Private Sub Form_Load()
    AturKontrol
    PENGATURAN_WARNA
    PENGATURAN_BAHASA
End Sub

Private Sub menuBuatAkunBaru_Click()
    FormBuatAkunBaru.Show vbModal, Me
End Sub


Private Sub menuIHNP_Click()
    menuIHNP.Checked = True
    menuIHP.Checked = False
    menuIK.Checked = False
    menuJIA.Checked = False
End Sub

Private Sub menuIHP_Click()
    menuIHNP.Checked = False
    menuIHP.Checked = True
    menuIK.Checked = False
    menuJIA.Checked = False
End Sub

Private Sub menuIK_Click()
    menuIHNP.Checked = False
    menuIHP.Checked = False
    menuIK.Checked = True
    menuJIA.Checked = False
End Sub

Private Sub menuJIA_Click()
    menuIHNP.Checked = False
    menuIHP.Checked = False
    menuIK.Checked = False
    menuJIA.Checked = True
End Sub

Private Sub menuKB_Click()
    If menuKB.Checked = True Then
        menuKB.Checked = False
        With textPassword
            .PasswordChar = ""
            .SetFocus
        End With
    ElseIf menuKB.Checked = False Then
        menuKB.Checked = True
        With textPassword
            .PasswordChar = "*"
            .SetFocus
        End With
    End If
End Sub
'BAGIAN UNTUK MEMATIKAN APLIKASI TAMBAHAN SAAT PROGRAM UTAMA DIAKHIRI
Sub MatikanProcessTool()
    Call KillApp("Access SQL Code Generator v2.0.exe")
    Call KillApp("Adress Register (PhoneBook) v1.0.exe")
    Call KillApp("Encrypt String v2.0.exe")
End Sub


Private Sub menuKeluar_Click()
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        Pesan = MsgBox("Anda yakin ingin keluar?", vbQuestion + vbYesNo, "Keluar?")
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Pesan = MsgBox("Are you sure to exit?", vbQuestion + vbYesNo, "Exit?")
    End If
    If Pesan = vbYes Then
        MatikanProcessTool
        End
    End If
End Sub

Private Sub menuKV_Click()
    Dim Jalankan
    Kalimat = "C:\Windows\system32\osk.exe"
    If Dir$(Kalimat) <> "" Then
        Jalankan = Shell(Kalimat)
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

Private Sub menuKW_Click()
    Kalimat = "http://rikymetalist.blogspot.com/p/software-ku.html"
    SITUS = ShellExecute(0, vbNullString, Kalimat, "", "", vbNormalFocus)
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

Private Sub TimerProgress_Timer()
Progress.Value = Progress.Value + 1
If Progress.Value = 100 Then
    LOGIN
    TimerProgress.Enabled = False
    Progress.Visible = False
    LabelLogin.Visible = False
    cmLogin.Enabled = True
    MenuMenu.Enabled = True
    menuBantuan.Enabled = True
    With Me
        .textPengguna.BackColor = vbWhite
        .textPassword.BackColor = vbWhite
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
            .Label1.Caption = "Pengguna"
            .menuBuatAkunBaru.Caption = "Buat Akun Baru"
            .menuIngatSaya.Caption = "Ingat Saya !"
            .menuIHNP.Caption = "Ingat Hanya Nama Pengguna"
            .menuIHP.Caption = "Ingat Hanya Password"
            .menuIK.Caption = "Ingat Keduanya"
            .menuJIA.Caption = "Jangan Ingat Apapun!"
            .menuKV.Caption = "Keyboard Virtual"
            .menuKB.Caption = "Karakter Bintang"
            .menuKeluar.Caption = "Keluar"
            .menuBantuan.Caption = "Bantuan"
            .menuPusatBantuan.Caption = "Pusat Bantuan"
            .menuKW.Caption = "Kunjungi Website"
        End With
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        With Me
            .Label1.Caption = "UserName"
            .menuBuatAkunBaru.Caption = "Create New Account"
            .menuIngatSaya.Caption = "Remember me !"
            .menuIHNP.Caption = "Just remember the User Name"
            .menuIHP.Caption = "Just remember the Password"
            .menuIK.Caption = "Remember both"
            .menuJIA.Caption = "Don't remember anything!"
            .menuKV.Caption = "Virtual Keyboard"
            .menuKB.Caption = "Star Character"
            .menuKeluar.Caption = "Exit"
            .menuBantuan.Caption = "Help"
            .menuPusatBantuan.Caption = "Help Center"
            .menuKW.Caption = "Visit the Website"
        End With
    End If
End Sub
