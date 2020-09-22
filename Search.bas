Attribute VB_Name = "Search"
Option Explicit

Public DoContinue As Boolean, OrigFiles As Single, CopiedFiles As Single
Public Src As String, Dest As String, Junk As String, Fol As String
Public SkippedFiles As Single, KCopied As Single, KSkipped As Single
Public DoneK As Double, X As Long, Trash As String, TotalK As Double

Dim BarString As String, i As Single, KLeft As Single
Dim FCurrent, FDest, Fs
Dim DateCurrent As Date, DateDest As Date
Dim CurFile As String
Dim DestPath As String
Dim CurDir As String

'Find File declarations and types
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * 260
        cAlternate As String * 14
End Type

Public Function FindAllFiles(Directory As String, Optional SearchFor As String)
    
    Dim Exists As Long, DivBy As Integer
    Dim hFindFile As Long
    Dim FileData As WIN32_FIND_DATA
    Dim TotalFileSize As Double
    
    With Form1

        DivBy = 10
        If .Grid1.Rows > 900 Then DivBy = 100

        'Sets Exists to equal 1
        'You need this so the loop doesn't automatically exit
    
        Exists = 1
    
        If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
        If SearchFor = vbNullString Then SearchFor = "*.*"
    
        'If the search for text doesn't contain any * or ?
        'Add *'s before and after
        
        If InStr(1, SearchFor, "?") = 0 And InStr(1, SearchFor, "*") = 0 Then
            SearchFor = "*" & SearchFor & "*"
        End If
    
        hFindFile = FindFirstFile(Directory & SearchFor, FileData)
    
        Do While hFindFile <> -1 And Exists <> 0
            DoEvents
        
            If (GetAttr(Directory & ClearNull(FileData.cFileName)) And vbDirectory) = vbDirectory Then
                If (ClearNull(FileData.cFileName) <> ".") And (ClearNull(FileData.cFileName) <> "..") Then
                    .FindFilesTmpResults.AddItem "[dir]  " & Directory & ClearNull(FileData.cFileName)
                End If
            ElseIf (GetAttr(Directory & ClearNull(FileData.cFileName)) And vbDirectory) <> vbDirectory Then
                .FindFilesTmpResults.AddItem "[file] " & Directory & ClearNull(FileData.cFileName)
                TotalFileSize = TotalFileSize + (FileLen(Directory & ClearNull(FileData.cFileName)) / 1024)
                'Form1.Text7 = Format(TotalFileSize, "########,##")
                Form1.Text8 = Val(Form1.Text8) + (FileLen(Directory & ClearNull(FileData.cFileName)))
            End If
        
            Exists = FindNextFile(hFindFile, FileData)
        Loop
    
        Do While .FindFilesTmpResults.ListCount
            .Grid1.AddItem .FindFilesTmpResults.List(0)
            If .Grid1.Rows / DivBy = (.Grid1.Rows \ DivBy) Then
                Form1.sbStatusBar.Panels(1).Text = "Building File List: " & .Grid1.Rows & " Files"
                DoEvents
            End If
            .FindFilesTmpResults.RemoveItem 0
        Loop
        Form1.sbStatusBar.Panels(1).Text = "Building File List: " & .Grid1.Rows & " Files"
        DoEvents
    
        Exists = 1
    
        hFindFile = FindFirstFile(Directory & "*", FileData)
    
        Do While hFindFile <> -1 And Exists <> 0
            On Error GoTo skiptonextfile
            If (GetAttr(Directory & ClearNull(FileData.cFileName)) And vbDirectory) = vbDirectory And (ClearNull(FileData.cFileName) <> "." And ClearNull(FileData.cFileName) <> "..") Then
                .FindFilesTmpDirs.AddItem Directory & ClearNull(FileData.cFileName)
                DoEvents
            End If
nextfile:
            On Error GoTo 0
        
            Exists = FindNextFile(hFindFile, FileData)
        Loop

  End With
  
  Exit Function
  
skiptonextfile:
    Err.Clear
    Resume nextfile
End Function

Public Function ClearNull(StringToClear As String) As String
    Dim StartOfNulls As Long
    
    'This function clears all the nulls in the string and
    'Returns it, by using Instr to find the first null
    
    StartOfNulls = InStr(1, StringToClear, Chr(0))
    ClearNull = Left(StringToClear, StartOfNulls - 1)
End Function

Public Sub SearchFilesInDir(ByVal Directory As String, Optional SearchFor As String)

    Dim NextDir As String
    
    With Form1
        .Grid1.Clear
        .FindFilesTmpResults.Clear
        FindAllFiles Directory, SearchFor

        Do While .FindFilesTmpDirs.ListCount
            DoEvents
            NextDir = .FindFilesTmpDirs.List(0)
            .FindFilesTmpDirs.RemoveItem 0
            FindAllFiles NextDir, SearchFor
        Loop
    End With
End Sub

Sub main()
    Form1.Show
    SetDrives.Show vbModal
    Analize
End Sub

Private Sub Analize()
    
    Form1.Grid1.ColWidth(0) = Form1.Grid1.Width - 100
    Form1.btnPause.Enabled = True

    OrigFiles = 0
    CopiedFiles = 0
    SkippedFiles = 0
    KCopied = 0
    KSkipped = 0
    DoneK = 0
    Form1.Picture1.ForeColor = RGB(0, 0, 255)
    Form1.Picture2.ForeColor = RGB(0, 0, 255)
    Form1.cmdReturn.Enabled = False

    Form1.sbStatusBar.Panels(1).Text = "Building File List: 0 Files"

    Form1.Visible = True
    Form1.Refresh
    Screen.MousePointer = vbHourglass

    Form1.FindFilesTmpResults.Clear
    Form1.FindFilesTmpDirs.Clear
    Form1.Grid1.Clear

    SearchFilesInDir Src, "*.*"
    
    Form1.sbStatusBar.Panels(1).Text = "Analysing " & Form1.Grid1.Rows & " Files for Backup"
    
    For X = 4 To Len(Src)
        If Mid(Src, X, 1) = "\" Then
            Junk = "[dir]  " & Left(Src, X)
            Form1.Grid1.AddItem Junk
        End If
    Next

    Form1.Grid1.ColSel = 0
    Form1.Grid1.Sort = 5
    Form1.Grid1.Refresh
    Form1.Text3 = Form1.Grid1.Rows
    Form1.Text3.Refresh

    For X = 0 To Form1.Grid1.Rows - 1

        Form1.Grid1.Row = X
        Junk = Form1.Grid1.Text
        DoEvents

        If Left(Junk, 7) <> "[dir]  " Then Exit For
            
        Junk = Right(Junk, Len(Junk) - 10)
        Junk = Dest & Junk
        Trash = Dir(Junk, vbDirectory)

        If Trash = "" Then
            Form1.sbStatusBar.Panels(1).Text = "Creating Destination Directory " & UCase(Junk)
            MkDir Junk
            CopiedFiles = CopiedFiles + 1
            Form1.Text5 = CopiedFiles
            Form1.Text5.Refresh
        Else
            Form1.sbStatusBar.Panels(1).Text = "Destination Directory " & UCase(Junk) & " Exists."
        End If
    
    Next
    
    TotalK = Val(Form1.Text8)
    
    Copy_Files

    Form1.Text7 = 0

    Screen.MousePointer = vbDefault
    Form1.cmdReturn.Enabled = True
    DoEvents
    Form1.btnStop.Enabled = True
    Form1.btnPause.Enabled = False
End Sub

Public Sub Copy_Files()
    On Error GoTo EHandler
    Form1.Frame2.Visible = True
    Form1.Frame3.Visible = True
    Form1.Refresh

    Form1.Grid1.Col = 0

    For X = 0 To Form1.Grid1.Rows - 1
        BarString = "Files "
        
        If X + 1 <= Form1.Grid1.Rows Then i = ((X + 1) / Form1.Grid1.Rows) * 100
        UpdateProgress Form1.Picture1, i
        
        Form1.Grid1.Row = X
        Junk = Form1.Grid1.Text

        If Form1.chkRefresh.Value <> 1 Then
            Form1.Grid1.TopRow = X
            Form1.Grid1.Refresh
        End If

        DoEvents

        CurFile = Right(Junk, Len(Junk) - 7)
        DestPath = Dest & Right(CurFile, Len(CurFile) - 3)

        If Form1.chkRefresh.Value <> 1 Then
            Form1.sbStatusBar.Panels(1).Text = UCase(CurFile)
        End If

        DoEvents

        If Left(Junk, 7) = "[file] " Then

            Set Fs = CreateObject("scripting.filesystemobject")

            DateCurrent = #1/1/2001#
            Set FCurrent = Fs.getfile(CurFile)
            DateCurrent = FCurrent.datelastmodified
            Junk = Format(DateCurrent, "dd/mm/yyyy")
            DateCurrent = CDate(Junk)

            DateDest = #1/1/1991#
            Junk = Dir(DestPath)
            If Junk = "" Then Junk = Dir(DestPath, vbDirectory)
            
            If Junk <> "" Then
                Set FDest = Fs.getfile(DestPath)
                DateDest = FDest.datelastmodified
                Junk = Format(DateDest, "dd/mm/yyyy")
                DateDest = CDate(Junk)
            End If

            If DateCurrent > DateDest Then
                Call FileCopy(CurFile, DestPath)
                CopiedFiles = CopiedFiles + 1
                Form1.Text5 = CopiedFiles
                Form1.Text1 = CurFile
                Form1.Text2 = DestPath
                Form1.Text1.Refresh
                Form1.Text2.Refresh
                KCopied = KCopied + (FileLen(CurFile) / 1024)
                Form1.Text6 = Format(KCopied, "##########,##")
                KLeft = Val(Form1.Text8) - FileLen(CurFile)
                Form1.Text8 = KLeft
                Form1.Text7 = Format(KLeft / 1024, "##########,##")
                DoneK = DoneK + FileLen(CurFile)
                BarString = "kBytes "
                If DoneK <= TotalK Then i = (DoneK / TotalK) * 100
                UpdateProgress Form1.Picture2, i
                DoEvents
            Else
                SkippedFiles = SkippedFiles + 1
                Form1.Text4 = SkippedFiles
                Form1.Text4.Refresh
                If Form1.chkRefresh.Value <> 1 Then
                    Form1.Grid1.Text = Form1.Grid1.Text & "      **SKIPPED **"
                End If
                KLeft = Val(Form1.Text8) - FileLen(CurFile)
                Form1.Text8 = KLeft
                Form1.Text7 = Format(KLeft / 1024, "##########,##")
                DoneK = DoneK + FileLen(CurFile)
                BarString = "kBytes "
                If DoneK <= TotalK Then i = (DoneK / TotalK) * 100
                UpdateProgress Form1.Picture2, i
            End If
        Else
            SkippedFiles = SkippedFiles + 1
            Form1.Text4 = SkippedFiles
            Form1.Text4.Refresh
        End If
NxtX:
    Next X
    Form1.Frame2.ForeColor = &HFF&
    Form1.Frame2.Caption = "Backup Completed"
    Form1.Text5 = Val(Form1.Text3) - Val(Form1.Text4)
    Exit Sub
EHandler:
    Select Case Err.Number
        Case 70
            Resume Next
        Case Else
            MsgBox "File copy error: " & Err.Description & " " & Err.Number
            Resume Next
    End Select
End Sub

Public Sub UpdateProgress(PB As Control, ByVal percent)
    Dim num$        'use percent
    If Not PB.AutoRedraw Then      'picture in memory ?
        PB.AutoRedraw = -1          'no, make one
    End If
    PB.Cls                      'clear picture in memory
    PB.ScaleWidth = 100         'new scalemodus
    PB.DrawMode = 10            'not XOR Pen Modus
    num$ = BarString & Format$(percent, "###") + "%"
    PB.CurrentX = 50 - PB.TextWidth(num$) / 2
    PB.CurrentY = (PB.ScaleHeight - PB.TextHeight(num$)) / 2
    PB.Print num$               'print percent
    PB.Line (0, 0)-(percent, PB.ScaleHeight), , BF
    PB.Refresh          'show difference
End Sub
