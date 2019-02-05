VERSION 5.00
Begin VB.Form FormLogOut 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LogOuting"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3420
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormLogOut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerLabel1 
      Interval        =   300
      Left            =   120
      Top             =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sedang menyimpan data . . . (100%)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2610
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Harap Tunggu. Sedang mengeluarkan . . ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "FormLogOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If FormPengaturan.cmbBahasa.ListIndex = 1 Then
    Label1.Caption = "Please wait. Logouting . . ."
Else
    Label1.Caption = "Mohon tunggu. Sedang mengeluarkan . . ."
End If
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
End Sub

Private Sub TimerLabel1_Timer()
If FormPengaturan.cmbBahasa.ListIndex = 1 Then
    Select Case Label1.Caption
    Case Is = "Please wait. Logouting . . ."
        Label1.Caption = "Please wait. Logouting"
    Case Is = "Please wait. Logouting"
        Label1.Caption = "Please wait. Logouting ."
    Case Is = "Please wait. Logouting ."
        Label1.Caption = "Please wait. Logouting . ."
    Case Is = "Please wait. Logouting . ."
        Label1.Caption = "Please wait. Logouting . . ."
    End Select
Else
    Select Case Label1.Caption
    Case Is = "Mohon tunggu. Sedang mengeluarkan . . ."
        Label1.Caption = "Mohon tunggu. Sedang mengeluarkan"
    Case Is = "Mohon tunggu. Sedang mengeluarkan"
        Label1.Caption = "Mohon tunggu. Sedang mengeluarkan ."
    Case Is = "Mohon tunggu. Sedang mengeluarkan ."
        Label1.Caption = "Mohon tunggu. Sedang mengeluarkan . ."
    Case Is = "Mohon tunggu. Sedang mengeluarkan . ."
        Label1.Caption = "Mohon tunggu. Sedang mengeluarkan . . ."
    End Select
End If
Select Case Me.Caption
    Case Is = "LogOuting"
        Me.Caption = "LogOuting ."
    Case Is = "LogOuting ."
        Me.Caption = "LogOuting . ."
    Case Is = "LogOuting . ."
        Me.Caption = "LogOuting . . ."
    Case Is = "LogOuting . . ."
        Me.Caption = "LogOuting . . . ."
    Case Is = "LogOuting . . . ."
        Me.Caption = "LogOuting . . . . ."
    Case Is = "LogOuting . . . . ."
        Me.Caption = "LogOuting"
    End Select
End Sub

