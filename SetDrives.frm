VERSION 5.00
Begin VB.Form SetDrives 
   Caption         =   "Set Source and Destination Directories"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   Icon            =   "SetDrives.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5250
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton btnExit 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   4500
         TabIndex        =   12
         Top             =   4560
         Width           =   1300
      End
      Begin VB.CommandButton cmdBackup 
         Caption         =   "Backup"
         Height          =   400
         Left            =   1500
         TabIndex        =   11
         Top             =   4560
         Width           =   1300
      End
      Begin VB.Frame Frame2 
         Caption         =   "Backup To"
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
         Height          =   3615
         Left            =   3720
         TabIndex        =   4
         Top             =   240
         Width           =   3615
         Begin VB.DriveListBox Drive2 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   3375
         End
         Begin VB.DirListBox Dir2 
            Height          =   2565
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   3375
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Backup From"
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
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3615
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   3375
         End
         Begin VB.DirListBox Dir1 
            Height          =   2565
            Left            =   120
            TabIndex        =   2
            Top             =   840
            Width           =   3375
         End
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   10
         Top             =   4320
         Width           =   6135
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   9
         Top             =   3960
         Width           =   6135
      End
      Begin VB.Label Label2 
         Caption         =   "Backup To:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Backup From:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3960
         Width           =   1095
      End
   End
End
Attribute VB_Name = "SetDrives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private Sub btnExit_Click()
    Unload Me
    End
End Sub

Private Sub cmdBackup_Click()
    Src = SetDrives.Label3(0)
    Dest = SetDrives.Label3(1)
    If Right$(Dest, 1) <> "\" Then Dest = Dest & "\"
    If Right$(Src, 1) <> "\" Then Src = Src & "\"

    Junk = Label3(0)
    If Right(Junk, 1) <> "\" Then
        Label3(0) = Junk & "\"
    End If
    Junk = Label3(1)
    If Right(Junk, 1) <> "\" Then
        Label3(1) = Junk & "\"
    End If
    Label3(0).Refresh
    Label3(1).Refresh

    If Label3(0) = Label3(1) Then
        Beep
        MsgBox "Cannot backup to the same folder."
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub Dir1_Change()
   With Dir1
      .Path = .List(.ListIndex)
      Label3(0).Caption = .Path
   End With
End Sub

Private Sub Dir2_Change()
   With Dir2
      .Path = .List(.ListIndex)
      Label3(1).Caption = .Path
   End With
End Sub

Private Sub Drive1_Change()
    On Error GoTo EHandler
    Dir1.Path = Left(Drive1.Drive, 2)
    Label3(0).Caption = Dir1.Path
    Exit Sub
EHandler:
   Select Case Err.Number
        Case 68     'drive not available
            MsgBox "Drive " & Drive1.Drive & " is not available."
            Drive1.Drive = "C:\"
            Exit Sub
    End Select
End Sub

Private Sub Drive2_Change()
    On Error GoTo EHandler
    Dir2.Path = Left(Drive2.Drive, 2)
    Label3(1).Caption = Dir2.Path
    Exit Sub
EHandler:
   Select Case Err.Number
        Case 68     'drive not available
            MsgBox "Drive " & Drive2.Drive & " is not available."
            Drive2.Drive = "C:\"
            Exit Sub
    End Select
End Sub

Private Sub Form_Load()
    Label3(0).Caption = Dir1.Path
    Label3(1).Caption = Dir2.Path
    Dir1.Path = "C:\temp2"
    Dir2.Path = "C:\read"
End Sub
