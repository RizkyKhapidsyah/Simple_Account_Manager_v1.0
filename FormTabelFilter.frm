VERSION 5.00
Object = "{5DC43A6F-8B43-4A60-A977-95A8CDDD093A}#1.0#0"; "dcButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormTabelFilter 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabel Filter"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormTabelFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.ComboBox cmbFilter 
         Height          =   390
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin Dacara_dcButton.dcButton cmLihat 
         Height          =   345
         Left            =   2400
         TabIndex        =   2
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         BackColor       =   12230304
         ButtonStyle     =   3
         Caption         =   "&Lihat"
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
         PicDown         =   "FormTabelFilter.frx":030A
         PicHot          =   "FormTabelFilter.frx":0624
         PicNormal       =   "FormTabelFilter.frx":093E
         PicSize         =   1
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmTutup 
         Height          =   345
         Left            =   3600
         TabIndex        =   3
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         BackColor       =   12230304
         ButtonStyle     =   3
         Caption         =   "&Tutup"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   5
         PicDown         =   "FormTabelFilter.frx":0C58
         PicHot          =   "FormTabelFilter.frx":10AA
         PicNormal       =   "FormTabelFilter.frx":14FC
         PicSize         =   1
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Berdasarkan :"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1290
      End
   End
   Begin XPEngine.XPControl XP_Engine 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
End
Attribute VB_Name = "FormTabelFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    With Me
        If FORM_UTAMA.cmIdentitasPribadi.FontBold = True Then
            .Caption = "Filter : " & FORM_UTAMA.cmIdentitasPribadi.Caption
            .cmbFilter.Clear
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(1).Text, 0
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(2).Text, 1
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(3).Text, 2
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(4).Text, 3
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(5).Text, 4
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(6).Text, 5
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(7).Text, 6
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(8).Text, 7
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(9).Text, 8
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(10).Text, 9
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(11).Text, 10
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(12).Text, 11
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(13).Text, 12
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(14).Text, 13
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(15).Text, 14
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(16).Text, 15
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(17).Text, 16
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(18).Text, 17
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(19).Text, 18
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(20).Text, 19
        ElseIf FORM_UTAMA.cmBukuAlamat.FontBold = True Then
            .Caption = "Filter : " & FORM_UTAMA.cmBukuAlamat.Caption
            .cmbFilter.Clear
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(1).Text, 0
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(2).Text, 1
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(3).Text, 2
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(4).Text, 3
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(5).Text, 4
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(6).Text, 5
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(7).Text, 6
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(8).Text, 7
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(9).Text, 8
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(10).Text, 9
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(11).Text, 10
        ElseIf FORM_UTAMA.cmUlangTahun.FontBold = True Then
            .Caption = "Filter : " & FORM_UTAMA.cmUlangTahun.Caption
            .cmbFilter.Clear
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(1).Text, 0
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(2).Text, 1
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(3).Text, 2
        ElseIf FORM_UTAMA.cmAgenda.FontBold = True Then
            .Caption = "Filter : " & FORM_UTAMA.cmAgenda.Caption
            .cmbFilter.Clear
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(1).Text, 0
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(2).Text, 1
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(3).Text, 2
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(4).Text, 3
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(5).Text, 4
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(6).Text, 5
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(7).Text, 6
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(8).Text, 7
        ElseIf FORM_UTAMA.cmRegistrasiSoftware.FontBold = True Then
            .Caption = "Filter : " & FORM_UTAMA.cmRegistrasiSoftware.Caption
            .cmbFilter.Clear
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(1).Text, 0
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(2).Text, 1
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(3).Text, 2
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(4).Text, 3
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(5).Text, 4
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(6).Text, 5
                .cmbFilter.AddItem FORM_UTAMA.LV.ColumnHeaders.Item(7).Text, 6
        Else
            Me.Hide
        End If
        .cmbFilter.ListIndex = 0
    End With
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

Private Sub cmLihat_Click()
If CN_FormUtama.State = adStateOpen Then CN_FormUtama.Close
    CN_FormUtama.CursorLocation = adUseClient
    CN_FormUtama.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Windows\rssam\inc\pggn\" & FORM_UTAMA.StatusBawah.Panels.Item(2).Text & "\data.rdb;Persist Security Info=False"
    
    If FORM_UTAMA.cmIdentitasPribadi.FontBold = True Then
        FORM_UTAMA.LV.ListItems.Clear
        With FORM_UTAMA.ADODC_UTAMA
            .ConnectionString = CN_FormUtama.ConnectionString
            .RecordSource = "Select * From tbIdentitasPribadi;"
            .Refresh
        End With
        
            Select Case cmbFilter.ListIndex
                Case Is = 0
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Nama Lengkap", 2000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Original Name", 2000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(0).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 1
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Nama Panggilan", 2000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Cool Name", 2000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(1).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 2
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "TTL", 5000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Place/Born Date", 5000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(2).Value & ", " & FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(3).Value & " - " & FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(4).Value & " - " & FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(5).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 3
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Jenis Kelamin", 3000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Gender", 3000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(6).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 4
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Agama", 3000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Religion", 3000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(7).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 5
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Golongan Darah", 3000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Blood Type", 3000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(8).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 6
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Pekerjaan", 3000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Jobs", 3000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(9).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 7
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Alamat Rumah", 3000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Home Address", 3000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(10).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 8
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        .ColumnHeaders.Add , , "E-Mail", 3000
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(11).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 9
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        .ColumnHeaders.Add , , "Website", 3000
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(12).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 10
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Nomor Telepon", 3000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Phone Number", 3000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(13).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 11
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Kota Asal", 3000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Home Town", 3000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(14).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 12
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Kota Sekarang", 3000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "City Now", 3000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(15).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 13
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Kode Pos", 3000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Zip/Postal Code", 3000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(16).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 14
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Provinsi", 3000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "State", 3000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(17).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 15
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Kewarganegaraan", 3000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Citizenship", 3000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(18).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 16
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Status Pendidikan", 3000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Educational Status", 3000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(19).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 17
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Status Hubungan", 3000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Relationship Status", 3000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(20).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 18
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Hobby", 3000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Hobbies", 3000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(21).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
                Case Is = 19
                    With FORM_UTAMA.LV
                        .ColumnHeaders.Clear
                        If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                            .ColumnHeaders.Add , , "Keterangan", 3000
                        ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                            .ColumnHeaders.Add , , "Keterangan", 3000
                        End If
                        .View = lvwReport
                        .Sorted = True
                    End With
                    Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                        Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(22).Value)
                            FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                        Loop
            End Select
    ElseIf FORM_UTAMA.cmBukuAlamat.FontBold = True Then
            FORM_UTAMA.LV.ListItems.Clear
            With FORM_UTAMA.ADODC_UTAMA
                .ConnectionString = CN_FormUtama.ConnectionString
                .RecordSource = "Select * From tbBukuAlamat;"
                .Refresh
            End With
            
                Select Case cmbFilter.ListIndex
                    Case Is = 0
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Nama Kontak", 5000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Contact Name", 5000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(0).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 1
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Nama Panggilan", 2000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Cool Name", 2000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(1).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 2
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Nomor Telepon Pribadi", 10000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Private Phone Number", 10000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(2).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 3
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Nomor Telepon Rumah", 10000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "House Phone Number", 10000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(3).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 4
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Nomor Telepon Kantor", 10000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Office Phone Number", 10000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(4).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 5
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            .ColumnHeaders.Add , , "Fax", 3000
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(5).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 6
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Alamat E-Mail", 3000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Mail Address", 3000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(6).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 7
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            .ColumnHeaders.Add , , "Website", 10000
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(7).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 8
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Kode Pos", 3000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "ZIP/Postal Code", 3000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(8).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 9
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Alamat Rumah", 10000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Home Address", 10000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(9).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 10
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Keterangan", 10000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Description", 10000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(10).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
            End Select
    ElseIf FORM_UTAMA.cmUlangTahun.FontBold = True Then
            FORM_UTAMA.LV.ListItems.Clear
            With FORM_UTAMA.ADODC_UTAMA
                .ConnectionString = CN_FormUtama.ConnectionString
                .RecordSource = "Select * From tbUlangTahun;"
                .Refresh
            End With
            
                Select Case cmbFilter.ListIndex
                    Case Is = 0
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Nama ", 9000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Name", 5000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(0).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 1
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Tempat/Tanggal Lahir", 4000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Place/Born Day", 4000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(1).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 2
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Keterangan", 10000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Description", 10000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(2).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                End Select
    ElseIf FORM_UTAMA.cmAgenda.FontBold = True Then
            FORM_UTAMA.LV.ListItems.Clear
            With FORM_UTAMA.ADODC_UTAMA
                .ConnectionString = CN_FormUtama.ConnectionString
                .RecordSource = "Select * From tbAgenda;"
                .Refresh
            End With
            
                Select Case cmbFilter.ListIndex
                    Case Is = 0
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Kode Agenda", 5000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Code", 5000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(0).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 1
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Nama Agenda", 2000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Agenda Name", 2000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(1).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 2
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Tema", 10000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Thema", 10000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(2).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 3
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Tanggal", 10000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Date", 10000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(3).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 4
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Waktu Mulai", 10000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Begin Time", 10000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(4).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 5
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Waktu Akhir", 4000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "End Time", 4000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(5).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 6
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Tempat", 3000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Place", 3000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(6).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 7
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            .ColumnHeaders.Add , , "Website", 10000
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(7).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 8
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Keterangan Lain", 3000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Other Description", 3000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(8).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                End Select
    ElseIf FORM_UTAMA.cmRegistrasiSoftware.FontBold = True Then
            FORM_UTAMA.LV.ListItems.Clear
            With FORM_UTAMA.ADODC_UTAMA
                .ConnectionString = CN_FormUtama.ConnectionString
                .RecordSource = "Select * From tbRegistrasiSoftware;"
                .Refresh
            End With
            
                Select Case cmbFilter.ListIndex
                    Case Is = 0
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Nama Software", 10000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Software Name", 5000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(0).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 1
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Kategori", 2000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Category", 2000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(1).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 2
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            .ColumnHeaders.Add , , "Developer/Programmer", 10000
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(2).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 3
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Nama User/Group/Office", 10000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "User/Group/Office Name", 10000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(3).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 4
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            .ColumnHeaders.Add , , "Serial/Key/Code", 10000
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(4).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 5
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Jenis Lisensi", 4000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "License", 4000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(5).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                    Case Is = 6
                        With FORM_UTAMA.LV
                            .ColumnHeaders.Clear
                            If FormPengaturan.cmbBahasa.ListIndex = 0 Then
                                .ColumnHeaders.Add , , "Keterangan", 20000
                            ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
                                .ColumnHeaders.Add , , "Description", 20000
                            End If
                            .View = lvwReport
                            .Sorted = True
                        End With
                        Do Until FORM_UTAMA.ADODC_UTAMA.Recordset.EOF
                            Set LI = FORM_UTAMA.LV.ListItems.Add(, , FORM_UTAMA.ADODC_UTAMA.Recordset.Fields(6).Value)
                                FORM_UTAMA.ADODC_UTAMA.Recordset.MoveNext
                            Loop
                End Select
        FORM_UTAMA.ADODC_UTAMA.Refresh
    End If
End Sub

Private Sub cmTutup_Click()
    Unload Me
    FORM_UTAMA.cmRefresh_Click
End Sub

Private Sub Form_Activate()
    AturKontrol
    PENGATURAN_WARNA
    PENGATURAN_BAHASA
End Sub

Private Sub Form_Load()
    AturKontrol
    DisableCloseBtn Me
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
        Me.Caption = "Tabel Filter"
        Label1.Caption = "Filter Berdasarlkan : "
        cmLihat.Caption = "&Lihat"
        cmTutup.Caption = "&Tutup"
    ElseIf FormPengaturan.cmbBahasa.ListIndex = 1 Then
        Me.Caption = "Filter Table"
        Label1.Caption = "Filter As : "
        cmLihat.Caption = "&View"
        cmTutup.Caption = "&Close"
    End If
End Sub
