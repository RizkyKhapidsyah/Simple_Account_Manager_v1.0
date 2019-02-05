Attribute VB_Name = "Module1"
Option Explicit
' ENUM STATE UNTUK PILIHAN PADA FORM EXPORT (MEMBUKA FILE DAN FOLDER) >> membuat enum data untuk mode windows state)
Enum State
    SHOWNORMAL = 1
    SHOWMINIMIZED = 2
    SHOWMAXIMIZED = 3
    SHOWMINNOACTIVE = 7
    SHOWDEFAULT = 10
End Enum
'FUNGSI API UNTUK PILIHAN PADA FORM EXPORT (MEMBUKA FILE DAN FOLDER) mendeklarasikan fungsi yang memanggil library dari shell32.dll
'FUNGSI API YANG DIPAKAI JUGA UNTUK MEMBUAT HYPERLINK KE SITUS DAN EMAIL
Public Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hwnd As Long, _
        ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
'FUNGSI YANG DIPAKAI UNTUK BROWSE FOR FOLDER
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'FUNGSI API YANG DPIAKAI UNTUK MEMBUAT BROWSE FOR FOLDER
'FUNGSI API UNTUK MEMBUAT, MEMERIKSA DAN MENGHAPUS FOLDER/DIREKTORY
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
'TYPE UNTUK MEMBUAT, MEMERIKSA DAN MENGHAPUS FOLDER/DIREKTORY
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
'FUNGSI API YANG DIPAKAI UNTUK MENAMPILKAN ISI COMBOBOX TANPA MENGKLIKNYA
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg _
As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const CB_SHOWDROPDOWN = &H14F
'FUNGSI API YANG DIPAKAI UNTUK menonaktifkan tombol close
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'fungsi API untuk mengambil nama komputer
Public Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
'FUNGSI API YANG DIPAKAI UNTUK ALWAYS ON TOP
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long





Public CN As New ADODB.Connection
Public RS As New ADODB.Recordset
Global Objek As Control
Public X As String
Public Y As String
Global Pesan As Integer
Public LST As ListItem
Public Kalimat As String
Global LI As ListItem
Public NomorIndex As Integer
Public IsiTextBoxKosong_ID(100) As String
Public IsiTextBoxKosong_EN(100) As String
Public CN_FormUtama As New ADODB.Connection
Public CN_FormUtamaLogin As New ADODB.Connection
Global R As Long
Public AlamatEmail As Long
Public AlamatSitus As Long
Global ObjekArray(100) As Control
Public DbLokasi As String
Public DbNama As String
Global Z As Integer
'CONSTANTA BUATAN SENDIRI
Public Const UnguNatural = &HFF80FF
Public Const Merah = &HFF&
Public Const Pink = &H8080FF
Public Const HijauMuda = &HFF00&
Public Const Hitam = &H0&
Public Const Silver = &HC0C0C0
Public Const SilverNatural = &HE0E0E0
Public Const Orange = &H80C0FF
Public Const UnguJanda = &HC000C0
Public Const BiruMAC = &HFF9B48
Public Const SilverTua = &HC0C0C0
Public Const SilverFormUtama = &HD2CECF
Public Const NomorErrorUntukRecordYangSama = -2147467259
Public Const TidakTransparan = 500
'KONSTANTA YANG DIPAKAI UNTUK MEMBUAT BROWSE FOR FOLDER
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260
'CONSTANTA YANG DIPAKAI UNTUK MEMBUKA FOLDER DAN MEMBUAT HYPERLINK EMAIL DAN SITUS
Public Const SW_SHOWNORMAL = 1
'VARIABEL YANG DIGUNAKAN UNTUK HYPERLINK SITUS DAN EMAIL
Public EMAIL As Long
Public SITUS As Long
'CONSTANTA YANG DIPAKAI UNTUK ALWAYS ON TOP
Public Const HWND_TOPMOST = -&H1
Public Const HWND_NOTOPMOST = -&H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
'TYPE YANG DIPAKAI UNTUK MEMBUAT BROWSE FOR FOLDER
Public Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
' fungsi untuk membuka file
Public Function OpenLocation(URL As String, _
    WindowsState As State) As Long
    Dim lHWnd As Long
    Dim lAns As Long
    lAns = ShellExecute(lHWnd, "open", URL, vbNullString, _
    vbNullString, WindowsState)
    OpenLocation = lAns
End Function
'bagian untuk menghilangkan tombol close
Public Sub DisableCloseBtn(ByVal Frm As Form)
    Dim H As Long
    H = GetSystemMenu(Frm.hwnd, 0)
    RemoveMenu H, 6, &H400
    RemoveMenu H, 5, &H400
End Sub

'BAGIAN UNTUK ALWAYS ON TOP
Public Function SetOnTop(WndHandle As Long) As Boolean
    If SetWindowPos(WndHandle, HWND_TOPMOST, 0&, 0&, 0&, 0&, (SWP_NOSIZE Or SWP_NOMOVE)) Then
        SetOnTop = True
    Else
        SetOnTop = False
    End If
End Function
'BAGIAN UNTUK TIDAK ALWAYS ON TOP
Public Function NotOnTop(WndHandle As Long) As Boolean
    If SetWindowPos(WndHandle, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, (SWP_NOSIZE Or SWP_NOMOVE)) Then
        NotOnTop = True
    Else
        NotOnTop = False
    End If
End Function


Public Sub NyambunggUtama()
If CN.State = adStateOpen Then CN.Close
    CN.CursorLocation = adUseClient
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\rdln\lgn.rdb;Persist Security Info=False"
End Sub

Public Sub PusatError()
    MsgBox Err.Description & vbCrLf & _
            Err.Number, vbCritical + vbOKOnly, "Error"
End Sub

Public Sub AturDefaultFormat()
Select Case FormPengaturan.cmbDefaultFormat.ListIndex
    Case Is = 0
        FormBuatAkunBaru.CommonDialog1.FilterIndex = 2
        FormEkstrakDataKeText.CommonDialog1.FilterIndex = 2
        FormHasilPencarian.CommonDialog1.FilterIndex = 2
    Case Is = 1
        FormBuatAkunBaru.CommonDialog1.FilterIndex = 3
        FormEkstrakDataKeText.CommonDialog1.FilterIndex = 3
        FormHasilPencarian.CommonDialog1.FilterIndex = 3
    Case Is = 2
        FormBuatAkunBaru.CommonDialog1.FilterIndex = 4
        FormEkstrakDataKeText.CommonDialog1.FilterIndex = 4
        FormHasilPencarian.CommonDialog1.FilterIndex = 4
    Case Is = 3
        FormBuatAkunBaru.CommonDialog1.FilterIndex = 5
        FormEkstrakDataKeText.CommonDialog1.FilterIndex = 5
        FormHasilPencarian.CommonDialog1.FilterIndex = 5
    Case Is = 4
        FormBuatAkunBaru.CommonDialog1.FilterIndex = 6
        FormEkstrakDataKeText.CommonDialog1.FilterIndex = 6
        FormHasilPencarian.CommonDialog1.FilterIndex = 6
    Case Is = 5
        FormBuatAkunBaru.CommonDialog1.FilterIndex = 7
        FormEkstrakDataKeText.CommonDialog1.FilterIndex = 7
        FormHasilPencarian.CommonDialog1.FilterIndex = 7
    End Select
End Sub
'BAGIAN YANG DIPAKAI UNTUK MEMBUKA FOLDER
Public Sub OpenDirectory(Directory As String)
    ShellExecute 0, "Open", Directory, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub
'BAGIAN YANG DIPAKAI UNTUK MENGHAPUS FOLDER/DIREKTORY
Public Sub CreateNewDirectory(NewDirectory As String)
    Dim sDirTest As String
    Dim SecAttrib As SECURITY_ATTRIBUTES
    Dim bSuccess As Boolean
    Dim sPath As String
    Dim iCounter As Integer
    Dim sTempDir As String
    
    sPath = NewDirectory
    
    If Right(sPath, Len(sPath)) <> "\" Then
    sPath = sPath & "\"
    End If
    
    iCounter = 1
    
    Do Until InStr(iCounter, sPath, "\") = 0
        iCounter = InStr(iCounter, sPath, "\")
        sTempDir = Left(sPath, iCounter)
        sDirTest = Dir(sTempDir)
        iCounter = iCounter + 1
        SecAttrib.lpSecurityDescriptor = &O0
        SecAttrib.bInheritHandle = False
        SecAttrib.nLength = Len(SecAttrib)
        bSuccess = CreateDirectory(sTempDir, SecAttrib)
    Loop
End Sub
'BAGIAN UNTUK MENGECEK KEBERADAAN FOLDER/DIREKTORY
Public Function DirectoryExist(DirPath As String) As Boolean
    DirectoryExist = Dir(DirPath, vbDirectory) <> ""
End Function

'function untuk mengambil nama komputer
Public Function GetComputerName() As String
Dim sResult As String * 255
    GetComputerNameA sResult, 255
    GetComputerName = Left$(sResult, InStr(sResult, Chr$(0)) - 1)
End Function

'PENCEGAHAN ERROR
Public Sub HancurkanError()
If Err.Number = NomorErrorUntukRecordYangSama Then
    If FormPengaturan.cmbBahasa.ListIndex = 0 Then
        MsgBox "Maaf, data yang Anda masukkan sudah ada!", vbExclamation + vbOKOnly, "Data Sudah Ada!"
    Else
        MsgBox "Your input is already exist!", vbExclamation + vbOKOnly, "File Already Exist!"
    End If
Else
    If FormPengaturan.cmbBahasa.ListIndex = 1 Then
        MsgBox "Maaf, ada kesalahan internet. Silahkan ulangi perintah Anda!", vbExclamation + vbOKOnly, "MainSystem-Error"
    Else
        MsgBox "Sorry, an internal error. Please restart your commands!", vbExclamation + vbOKOnly, "MainSystem-Error"
    End If
End If
End Sub

Public Function CreateDB() 'BAGIAN UNTUK MEMBUAT DATABASE UTAMA PENGGUNA
    Dim DTB As Database
    Dim tbDataLogin As TableDef
    Dim tbIdentitasPribadi As TableDef
    Dim tbBukuAlamat As TableDef
    Dim tbUlangTahun As TableDef
    Dim tbAgenda As TableDef
    Dim tbRegistrasiSoftware As TableDef
    Dim tbJejaringSosial As TableDef
    Dim tbElectronicMail As TableDef
    Dim tbForumInternet As TableDef
    Dim tbFTP As TableDef
    Dim tbBlogging As TableDef
    Dim tbRiwayat As TableDef
    
    Set DTB = CreateDatabase(DbLokasi & "\" & DbNama, dbLangGeneral)
    
    Set tbDataLogin = DTB.CreateTableDef("tbDataLogin")
        With tbDataLogin
            .Fields.Append .CreateField("Nama_Pengguna", dbText, 254)
            .Fields.Append .CreateField("Password", dbText, 254)
            .Fields.Append .CreateField("Nama_Asli", dbText, 254)
        End With
    Set tbIdentitasPribadi = DTB.CreateTableDef("tbIdentitasPribadi") 'Nama Tabel IdentitasPribadi
    With tbIdentitasPribadi
        .Fields.Append .CreateField("Nama_Lengkap", dbText, 254)
        .Fields.Append .CreateField("Nama_Panggilan", dbText, 254)
        .Fields.Append .CreateField("TempatLahir", dbText, 254)
        .Fields.Append .CreateField("TanggalLahir", dbText, 254)
        .Fields.Append .CreateField("BulanLahir", dbText, 254)
        .Fields.Append .CreateField("TahunLahir", dbText, 254)
        .Fields.Append .CreateField("Jenis_Kelamin", dbText, 254)
        .Fields.Append .CreateField("Agama", dbText, 254)
        .Fields.Append .CreateField("Golongan_Darah", dbText, 254)
        .Fields.Append .CreateField("Pekerjaan", dbText, 254)
        .Fields.Append .CreateField("Alamat_Rumah", dbText, 254)
        .Fields.Append .CreateField("E_Mail", dbText, 254)
        .Fields.Append .CreateField("Website", dbText, 254)
        .Fields.Append .CreateField("Nomor_Telepon", dbText, 254)
        .Fields.Append .CreateField("Kota_Asal", dbText, 254)
        .Fields.Append .CreateField("Kota_Sekarang", dbText, 254)
        .Fields.Append .CreateField("Kode_Pos", dbText, 254)
        .Fields.Append .CreateField("Provinsi", dbText, 254)
        .Fields.Append .CreateField("Kewarganegaraan", dbText, 254)
        .Fields.Append .CreateField("Status_Pendidikan", dbText, 254)
        .Fields.Append .CreateField("Status_Hubungan", dbText, 254)
        .Fields.Append .CreateField("Hobby", dbText, 254)
        .Fields.Append .CreateField("Keterangan", dbText, 254)
    End With
    Set tbBukuAlamat = DTB.CreateTableDef("tbBukuAlamat")
    With tbBukuAlamat
        .Fields.Append .CreateField("Nama_Kontak", dbText, 254)
        .Fields.Append .CreateField("Nama_Panggilan", dbText, 254)
        .Fields.Append .CreateField("Nomor_Telepon_Pribadi", dbText, 254)
        .Fields.Append .CreateField("Nomor_Telepon_Rumah", dbText, 254)
        .Fields.Append .CreateField("Nomor_Telepon_Kantor", dbText, 254)
        .Fields.Append .CreateField("Fax", dbText, 254)
        .Fields.Append .CreateField("Alamat_EMail", dbText, 254)
        .Fields.Append .CreateField("Website", dbText, 254)
        .Fields.Append .CreateField("ZIP_Postal_Code", dbText, 254)
        .Fields.Append .CreateField("Alamat_Rumah", dbText, 254)
        .Fields.Append .CreateField("Keterangan", dbText, 254)
    End With
    Set tbUlangTahun = DTB.CreateTableDef("tbUlangTahun")
    With tbUlangTahun
        .Fields.Append .CreateField("Nama", dbText, 254)
        .Fields.Append .CreateField("TTL", dbText, 254)
        .Fields.Append .CreateField("Keterangan", dbText, 254)
    End With
    Set tbAgenda = DTB.CreateTableDef("tbAgenda")
    With tbAgenda
        .Fields.Append .CreateField("Kode_Agenda", dbText, 254)
        .Fields.Append .CreateField("Nama_Agenda", dbText, 254)
        .Fields.Append .CreateField("Tema", dbText, 254)
        .Fields.Append .CreateField("Tanggal", dbText, 254)
        .Fields.Append .CreateField("Waktu_Mulai", dbText, 254)
        .Fields.Append .CreateField("Waktu_Akhir", dbText, 254)
        .Fields.Append .CreateField("Tempat", dbText, 254)
        .Fields.Append .CreateField("Keterangan_Lain", dbText, 254)
    End With
    Set tbRegistrasiSoftware = DTB.CreateTableDef("tbRegistrasiSoftware")
    With tbRegistrasiSoftware
        .Fields.Append .CreateField("Nama_Software", dbText, 254)
        .Fields.Append .CreateField("Kategori", dbText, 254)
        .Fields.Append .CreateField("Developer", dbText, 254)
        .Fields.Append .CreateField("Username", dbText, 254)
        .Fields.Append .CreateField("Serial_Key", dbText, 254)
        .Fields.Append .CreateField("Jenis_Lisensi", dbText, 254)
        .Fields.Append .CreateField("Keterangan", dbText, 254)
    End With
    Set tbJejaringSosial = DTB.CreateTableDef("tbJejaringSosial")
    With tbJejaringSosial
        .Fields.Append .CreateField("Nama_Jejaring", dbText, 254)
        .Fields.Append .CreateField("Nama_Pengguna", dbText, 254)
        .Fields.Append .CreateField("Alamat_Email", dbText, 254)
        .Fields.Append .CreateField("Password", dbText, 254)
        .Fields.Append .CreateField("URL", dbText, 254)
        .Fields.Append .CreateField("Pemilik_Akun", dbText, 254)
        .Fields.Append .CreateField("Tanggal", dbText, 254)
        .Fields.Append .CreateField("Keterangan", dbText, 254)
    End With
    Set tbElectronicMail = DTB.CreateTableDef("tbElectronicMail")
    With tbElectronicMail
        .Fields.Append .CreateField("Nama_Server", dbText, 254)
        .Fields.Append .CreateField("Nama_Pengguna", dbText, 254)
        .Fields.Append .CreateField("Alamat_Email", dbText, 254)
        .Fields.Append .CreateField("Password", dbText, 254)
        .Fields.Append .CreateField("Pertanyaan_Rahasia", dbText, 254)
        .Fields.Append .CreateField("Jawaban_Pertanyaan", dbText, 254)
        .Fields.Append .CreateField("URL", dbText, 254)
        .Fields.Append .CreateField("Pemilik_Akun", dbText, 254)
        .Fields.Append .CreateField("Tanggal", dbText, 254)
        .Fields.Append .CreateField("Keterangan", dbText, 254)
    End With
    Set tbForumInternet = DTB.CreateTableDef("tbForumInternet")
    With tbForumInternet
        .Fields.Append .CreateField("Nama_Forum", dbText, 254)
        .Fields.Append .CreateField("Nama_Pengguna", dbText, 254)
        .Fields.Append .CreateField("Alamat_Email", dbText, 254)
        .Fields.Append .CreateField("Password", dbText, 254)
        .Fields.Append .CreateField("Posisi", dbText, 254)
        .Fields.Append .CreateField("NickName", dbText, 254)
        .Fields.Append .CreateField("URL", dbText, 254)
        .Fields.Append .CreateField("Tanggal", dbText, 254)
        .Fields.Append .CreateField("Keterangan", dbText, 254)
    End With
    Set tbFTP = DTB.CreateTableDef("tbFTP")
    With tbFTP
        .Fields.Append .CreateField("Nama_Host", dbText, 254)
        .Fields.Append .CreateField("Port", dbText, 254)
        .Fields.Append .CreateField("Nama_Server", dbText, 254)
        .Fields.Append .CreateField("Nama_Pengguna", dbText, 254)
        .Fields.Append .CreateField("Alamat_Email", dbText, 254)
        .Fields.Append .CreateField("Password", dbText, 254)
        .Fields.Append .CreateField("Tanggal", dbText, 254)
        .Fields.Append .CreateField("Keterangan", dbText, 254)
    End With
    Set tbBlogging = DTB.CreateTableDef("tbBlogging")
    With tbBlogging
        .Fields.Append .CreateField("Nama_Penyedia_Blog", dbText, 254)
        .Fields.Append .CreateField("Nama_Pengguna", dbText, 254)
        .Fields.Append .CreateField("E_Mail", dbText, 254)
        .Fields.Append .CreateField("Password", dbText, 254)
        .Fields.Append .CreateField("URL", dbText, 254)
        .Fields.Append .CreateField("Tanggal", dbText, 254)
        .Fields.Append .CreateField("Keterangan", dbText, 254)
    End With
    Set tbRiwayat = DTB.CreateTableDef("tbRiwayat")
    With tbRiwayat
        .Fields.Append .CreateField("Tanggal", dbText, 254)
        .Fields.Append .CreateField("Bulan", dbText, 254)
        .Fields.Append .CreateField("Tahun", dbText, 254)
        .Fields.Append .CreateField("Jam", dbText, 254)
        .Fields.Append .CreateField("Menit", dbText, 254)
        .Fields.Append .CreateField("Detik", dbText, 254)
        .Fields.Append .CreateField("Aktivitas", dbText, 254)
        .Fields.Append .CreateField("Nama_Komputer", dbText, 254)
    End With
    DTB.TableDefs.Append tbDataLogin
    DTB.TableDefs.Append tbIdentitasPribadi
    DTB.TableDefs.Append tbBukuAlamat
    DTB.TableDefs.Append tbUlangTahun
    DTB.TableDefs.Append tbAgenda
    DTB.TableDefs.Append tbRegistrasiSoftware
    DTB.TableDefs.Append tbJejaringSosial
    DTB.TableDefs.Append tbElectronicMail
    DTB.TableDefs.Append tbForumInternet
    DTB.TableDefs.Append tbFTP
    DTB.TableDefs.Append tbBlogging
    DTB.TableDefs.Append tbRiwayat
    
    
   
    Set tbDataLogin = Nothing
    Set tbDataLogin = Nothing
    Set tbIdentitasPribadi = Nothing
    Set tbBukuAlamat = Nothing
    Set tbUlangTahun = Nothing
    Set tbAgenda = Nothing
    Set tbRegistrasiSoftware = Nothing
    Set tbJejaringSosial = Nothing
    Set tbElectronicMail = Nothing
    Set tbForumInternet = Nothing
    Set tbFTP = Nothing
    Set tbBlogging = Nothing
    Set tbRiwayat = Nothing

    DTB.Close
    Screen.MousePointer = vbDefault

End Function


