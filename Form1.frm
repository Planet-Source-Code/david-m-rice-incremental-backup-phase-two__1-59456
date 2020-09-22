VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Incremental Backup"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRefresh 
      Caption         =   "&Faster (Minimal screen refresh)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2790
      TabIndex        =   26
      Top             =   6975
      Width           =   3165
   End
   Begin VB.ListBox FindFilesTmpResults 
      Height          =   450
      Left            =   7605
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   585
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox FindFilesTmpDirs 
      Height          =   450
      Left            =   7650
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton btnPause 
      Caption         =   "&Pause"
      Enabled         =   0   'False
      Height          =   450
      Left            =   1395
      TabIndex        =   21
      Top             =   6930
      Width           =   1000
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "E&xit"
      Enabled         =   0   'False
      Height          =   450
      Left            =   135
      TabIndex        =   20
      Top             =   6930
      Width           =   1000
   End
   Begin VB.Frame Frame3 
      Caption         =   "Backup Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1800
      Left            =   120
      TabIndex        =   5
      Top             =   5040
      Width           =   2535
      Begin VB.CommandButton cmdReturn 
         Caption         =   "Return to Selection Screen"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Original files are newer than backup files."
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Original files are not newer than backup files."
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Files Copied"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Files Skipped"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Files to Copy"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Copy Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1800
      Left            =   2760
      TabIndex        =   0
      Top             =   5040
      Width           =   5805
      Begin VB.PictureBox Picture2 
         Height          =   285
         Left            =   2880
         ScaleHeight     =   225
         ScaleWidth      =   2520
         TabIndex        =   19
         Top             =   960
         Width           =   2580
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   240
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "Text8"
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Height          =   285
         Left            =   2880
         ScaleHeight     =   225
         ScaleWidth      =   2655
         TabIndex        =   14
         Top             =   1320
         Width           =   2715
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   4035
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   4125
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "kBytes Remaining"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "kBytes Copied"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "To"
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Copying"
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   24
      Top             =   7485
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11113
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "3/12/05"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "8:29 PM"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4425
      Left            =   45
      TabIndex        =   25
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   7805
      _Version        =   393216
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnPause_Click()
    If btnPause.Caption = "&Pause" Then
        btnPause.Caption = "&Continue"
        btnStop.Enabled = True
        Form1.sbStatusBar.Panels(1).Text = "BACK-UP SESSION *PAUSED*"
        Do
            DoEvents
            If btnPause.Caption = "&Pause" Then Exit Do
        Loop
    Else
        btnStop.Enabled = False
        btnPause.Caption = "&Pause"
    End If
End Sub

Private Sub btnStop_Click()
    Unload Me
    End
End Sub

Private Sub cmdReturn_Click()
    SetDrives.Show vbModal
End Sub

Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 7800)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    Me.chkRefresh.Value = GetSetting(App.Title, "Settings", "cRefresh", 0)
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        If Me.Width < 8500 Then Me.Width = 8500
        If Me.Height < 5500 Then Me.Height = 5500

        Grid1.Width = Me.Width - 150
        Grid1.ColWidth(0) = Grid1.Width - 100
        
        btnStop.Top = (Me.Height - 700) - btnStop.Height
        btnPause.Top = btnStop.Top
        chkRefresh.Top = btnStop.Top
        
        Frame2.Top = btnStop.Top - (1800 + 50)
        Frame3.Top = btnStop.Top - (1800 + 50)
        Frame2.Width = Me.Width - (Frame2.Left + 150)

        Grid1.Height = Frame2.Top - (Grid1.Top + 50)

        Text1.Width = Frame2.Width - (Text1.Left + 100)
        Text2.Width = Frame2.Width - (Text2.Left + 100)
        
        Picture1.Width = Frame2.Width - (Picture1.Left + 100)
        Picture2.Width = Frame2.Width - (Picture2.Left + 100)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next

    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
        SaveSetting App.Title, "Settings", "cRefresh", Me.chkRefresh.Value
    End If
End Sub
