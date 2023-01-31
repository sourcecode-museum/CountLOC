VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5490
   ClientLeft      =   2070
   ClientTop       =   2430
   ClientWidth     =   5505
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5505
   Begin VB.Frame fraMain 
      Height          =   3000
      Left            =   45
      TabIndex        =   1
      Top             =   1080
      Width           =   5415
      Begin VB.Frame Frame1 
         Height          =   1395
         Left            =   135
         TabIndex        =   10
         Top             =   1485
         Width           =   2340
         Begin VB.PictureBox picIncludeBoxes 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   45
            ScaleHeight     =   1095
            ScaleWidth      =   2175
            TabIndex        =   11
            Top             =   135
            Width           =   2175
            Begin VB.CheckBox chkFilterOptions 
               Caption         =   "Auto generated code"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   90
               TabIndex        =   14
               Top             =   360
               Width           =   1995
            End
            Begin VB.CheckBox chkFilterOptions 
               Caption         =   "Blank lines"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   13
               Top             =   600
               Width           =   1740
            End
            Begin VB.CheckBox chkFilterOptions 
               Caption         =   "Comments"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   90
               TabIndex        =   12
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label lblTypeOfCode 
               BackStyle       =   0  'Transparent
               Caption         =   "Include in totals"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   0
               Left            =   405
               TabIndex        =   15
               Top             =   45
               Width           =   1560
            End
         End
      End
      Begin VB.CommandButton cmdFileOpen 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4845
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Browse for file to process"
         Top             =   390
         Width           =   405
      End
      Begin VB.TextBox txtOptTitle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         TabIndex        =   8
         Text            =   "txtOptTitle"
         Top             =   1035
         Width           =   5145
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   3502
         TabIndex        =   7
         ToolTipText     =   "Start processing this report"
         Top             =   2205
         Width           =   1035
      End
      Begin VB.PictureBox picDisplayRpt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   3375
         ScaleHeight     =   510
         ScaleWidth      =   1860
         TabIndex        =   5
         Top             =   1485
         Width           =   1860
         Begin VB.CheckBox chkDisplayRpt 
            Caption         =   "Display report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   135
            TabIndex        =   6
            Top             =   45
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   3502
         TabIndex        =   4
         ToolTipText     =   "Start processing this report"
         Top             =   2205
         Width           =   1035
      End
      Begin VB.CommandButton cmdChoice 
         Height          =   615
         Index           =   3
         Left            =   4635
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Terminate this application"
         Top             =   2205
         Width           =   615
      End
      Begin VB.CommandButton cmdChoice 
         Height          =   615
         Index           =   0
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Review an output file"
         Top             =   2205
         Width           =   615
      End
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Browse for project or file to process"
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
         Index           =   1
         Left            =   150
         TabIndex        =   18
         Top             =   240
         Width           =   4605
      End
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Name on title line in report  (Optional)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   810
         Width           =   4920
      End
      Begin VB.Label lblFilename 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   16
         Top             =   450
         Width           =   4605
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   90
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   210
      Width           =   510
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Count Lines of Code"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   0
      Left            =   690
      TabIndex        =   24
      Top             =   90
      Width           =   4770
   End
   Begin VB.Label lblCurrentFile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   45
      TabIndex        =   23
      Top             =   4365
      Width           =   5415
   End
   Begin VB.Label lblAuthor 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   2085
      TabIndex        =   22
      Top             =   870
      Width           =   1350
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "lblDisclaimer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   105
      TabIndex        =   21
      Top             =   4785
      Width           =   2670
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Current file being processed"
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
      Index           =   3
      Left            =   60
      TabIndex        =   20
      Top             =   4155
      Width           =   4650
   End
   Begin VB.Label lblMsg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "lblMsg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   660
      Index           =   4
      Left            =   3345
      TabIndex        =   19
      Top             =   4785
      Width           =   2085
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Module:        frmMain
'
' Description:   This application will count lines of code for both Visual
'                Basic and C projects.  It will handle VB Group projects and
'                single files from both C and VB.  In VB, a line continuation
'                character of " _" at the end of a line is not counted because
'                it is considered a continuation of the previous line.  Only
'                the first occurance of the continuation is counted.  This is
'                the main module to access Visual Basic and C source
'                code files.  This module will allow the user to select
'                either Visual Basic or C type files and determine the lines
'                of code count.
'
'                Depending of the current standard at the user's location,
'                there are three options to either add to the line count or
'                omit.  Options available are whether or not to include
'                Autogenerated code, blank lines, single braces (C only), or
'                comments in the final total.
'
'                Items NOT counted in VB are TRAILERS.  These are the logical
'                ending statements to a procedural heading.
'
'                   Example:    End Sub      End Function     End Property
'                               End Type     End If           Loop
'                               Wend         Next             End With
'                               End Select
'
'                Items NOT counted in C are BRACES( "{" and "}" ) on a line
'                by themselves or with a comment or a semicolon.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-NOV-2000  Kenneth Ives  kenaso@tx.rr.com
'              Module created
' 25-Nov-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Added flag for displaying final report
'              - Added optional title line to report
'              - Store settings in registry
' ***************************************************************************

Option Explicit

' ***************************************************************************
' Module variables
' ***************************************************************************
  Private mcolProjectNames      As Collection   ' List of project names
  Private mblnUnknown           As Boolean
  Private mblnVisualBasic       As Boolean
  Private mblnPreloadedList     As Boolean
  Private mblnIncludeAutoGen    As Boolean
  Private mblnIncludeComments   As Boolean
  Private mblnIncludeBlankLines As Boolean
  Private mstrStartPath         As String
  Private mstrPathAndFile       As String       ' Hold original project name
  Private mobjKeyEdit           As cKeyEdit

Private Sub chkDisplayRpt_Click()

    ' Verify flag is set
    gblnDisplayRpt = CBool(chkDisplayRpt.Value)
    
End Sub

Private Sub chkFilterOptions_Click(Index As Integer)

    ' update appropriate switches
    Select Case Index
           Case 0: mblnIncludeAutoGen = Not mblnIncludeAutoGen
           Case 1: mblnIncludeBlankLines = Not mblnIncludeBlankLines
           Case 2: mblnIncludeComments = Not mblnIncludeComments
    End Select
    
End Sub

Private Sub cmdChoice_Click(Index As Integer)

    Dim lngIndex      As Long
    Dim lngProjectCnt As Long
    Dim strMsg        As String
    Dim strRptFile    As String
    
    ' Process according to which button was selected
    Select Case Index
             
           Case 0   ' Find and review prior LOC reports
                gblnStopProcessing = False       ' Reset global flag
                strRptFile = BrowseForReports()  ' Select a prior report
                
                If Len(strRptFile) = 0 Then
                    GoTo cmdChoice_Exit               ' Cancel was selected
                Else
                    DisplayFile strRptFile, frmMain   ' Display finished report
                End If
                
           Case 1   ' Create LOC report
                If Len(TrimStr(lblFilename.Caption)) = 0 Then
                    InfoMsg "Cannot identify file to process."
                    Exit Sub
                End If
                
                DoEvents
                DisableControls    ' Enable buttons and checkboxes
                lngProjectCnt = 0
                strMsg = vbNullString
                DoEvents
               
                If Len(TrimStr(mstrPathAndFile)) > 0 Then
                    gstrLastPath = GetPath(mstrPathAndFile)  ' Capture last path used
                Else
                    gstrLastPath = vbNullString
                End If
                
                gstrOptTitle = TrimStr(txtOptTitle.Text)   ' Capture report title used
                AlwaysOnTop True                         ' Set main form to always be on top
                
                ' Process list of files
                If mblnPreloadedList Then
                    
                    lngProjectCnt = mcolProjectNames.Count
                    
                    For lngIndex = 1 To mcolProjectNames.Count
                        
                        ' Capture path and file name to be processed
                        mstrPathAndFile = mcolProjectNames.Item(lngIndex)
                        gstrLastPath = GetPath(mstrPathAndFile)  ' Capture path
                        
                        ' Display path and file name
                        lblFilename.Caption = vbNullString
                        lblFilename.Caption = ShrinkToFit(mstrPathAndFile, 50)
                    
                        ' Process each project one at a time
                        strRptFile = BeginProcessing(mstrPathAndFile, _
                                                     mblnIncludeAutoGen, _
                                                     mblnIncludeBlankLines, _
                                                     mblnIncludeComments, _
                                                     mblnVisualBasic, _
                                                     mblnPreloadedList)
                    
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            Exit For
                        End If
                    
                        DoEvents
                        If gblnDisplayRpt Then
                            DisplayFile strRptFile, frmMain   ' Display finished report
                        End If
                        
                    Next lngIndex
                    
                    EmptyCollection mcolProjectNames
                    
                    ' An error occurred or user opted to STOP processing
                    DoEvents
                    If gblnStopProcessing Then
                        GoTo cmdChoice_Exit
                    End If
                    
                Else
                    
                    ' Start processing selected project
                    strRptFile = BeginProcessing(mstrPathAndFile, _
                                                 mblnIncludeAutoGen, _
                                                 mblnIncludeBlankLines, _
                                                 mblnIncludeComments, _
                                                 mblnVisualBasic, _
                                                 mblnPreloadedList)
                End If
        
                AlwaysOnTop False                ' Release main form from being on top
                
                If mblnPreloadedList Then
                    strMsg = IIf(mblnVisualBasic, " Visual BASIC files.", " C language files.")
                    InfoMsg "Finished processing " & CStr(lngProjectCnt) & strMsg
                Else
                    If gblnDisplayRpt Then
                        DisplayFile strRptFile, frmMain   ' Display finished report
                    Else
                        ' Display a finish message
                        strMsg = IIf(mblnVisualBasic, "Visual BASIC file.", "C language file.")
                        InfoMsg "Finished processing one " & strMsg
                    End If
                End If
                        
           Case 2   ' Stop button was selected
                DoEvents
                gblnStopProcessing = True  ' Set global flag
                DoEvents
                    
           Case Else   ' Stop this application
                DoEvents
                gstrOptTitle = txtOptTitle.Text  ' Save optional title line
                gblnStopProcessing = True        ' Set global flag
                DoEvents
                    
                TerminateProgram   ' Stop and close application
                Exit Sub
    End Select
    
cmdChoice_Exit:
    DoEvents
    lblFilename.Caption = vbNullString
    lblCurrentFile.Caption = vbNullString
    mstrPathAndFile = vbNullString
    mblnPreloadedList = False
    Close
    EnableControls
        
    ' Go back to where we started
    On Error Resume Next
    If Len(TrimStr(mstrStartPath)) > 0 Then
        gstrLastPath = mstrStartPath
        ChDir mstrStartPath
    End If
    On Error GoTo 0
    
End Sub

Private Sub DisableControls()

    DoEvents
    gstrTempPath = vbNullString           ' Empty temp path
    gblnStopProcessing = False
    
    With frmMain
        .cmdChoice(0).Enabled = False
        .cmdChoice(1).Visible = False     ' Hide START button
        .cmdChoice(1).Enabled = False
        .cmdChoice(2).Enabled = True      ' Show STOP button
        .cmdChoice(2).Visible = True
        .cmdFileOpen.Enabled = False
        .picIncludeBoxes.Enabled = False  ' Disable check boxes
        .txtOptTitle.Enabled = False
    End With
    
    DoEvents
    cmdChoice(2).SetFocus   ' Set focus on STOP button
    
End Sub

Private Sub EnableControls()

    DoEvents
    gstrTempPath = vbNullString           ' Empty temp path
    
    With frmMain
        .cmdChoice(0).Enabled = True
        .cmdChoice(1).Enabled = True         ' Show START button
        .cmdChoice(1).Visible = True
        .cmdChoice(2).Visible = False        ' Hide STOP button
        .cmdChoice(2).Enabled = False
        .cmdFileOpen.Enabled = True
        .picIncludeBoxes.Enabled = True      ' Enable check boxes
        .chkFilterOptions(0).Enabled = True  ' Disabled if C language
        .lblCurrentFile.Caption = vbNullString
        .txtOptTitle.Enabled = True
    End With
    DoEvents
    
End Sub

Private Sub cmdFileOpen_Click()

    Dim hFile        As Long
    Dim vntFilter    As Variant
    Dim strRecord    As String
    Dim strFileName  As String
    Dim strExtension As String

    gblnStopProcessing = False   ' Set flags to FALSE
    mblnUnknown = False
    mblnVisualBasic = False
    mblnPreloadedList = False
    
    lblFilename.Caption = vbNullString
    mstrPathAndFile = vbNullString
    mstrStartPath = vbNullString
    
    ' Set file selection filters
    vntFilter = "VB Projects (*.vbg,*.vbp)" & vbNullChar & "*.vbg;*.vbp" & vbNullChar & _
                "C Projects (*.dsw,*.dsp,*.mak)" & vbNullChar & "*.dsw;*.dsp;*.mak" & vbNullChar & _
                "Project lists (*.lst)" & vbNullChar & "*.lst" & vbNullChar & _
                "Misc VB Files (*.bas,*.cls,*.frm,*.ctl,*.dsr)" & vbNullChar & "*.bas;*.cls;*.frm;*.ctl;*.dsr" & vbNullChar & _
                "Misc C Files (*.c,*.cpp,*.h,*.hpp)" & vbNullChar & "*.c;*.cpp;*.h;*.hpp" & vbNullChar & vbNullChar
        
    ' Display dialog box
    frmMain.Hide                                                             ' Hide main form
    SetStartingFolder gstrLastPath                                           ' Optional information
    mstrPathAndFile = ShowFileOpen(frmMain, "Select a project", vntFilter)   ' Display File Open dialog box
    frmMain.Show                                                             ' Show main form
    
    mstrPathAndFile = TrimStr(mstrPathAndFile)   ' Remove unwanted characters
        
    ' See if we have some data
    If Len(mstrPathAndFile) > 0 Then
    
        mstrStartPath = GetFullPath(mstrPathAndFile)   ' Capture path
        strFileName = GetFilename(mstrPathAndFile)     ' Capture name of file
        strExtension = GetFilenameExt(strFileName)     ' Capture file extension
        gstrLastPath = mstrStartPath                   ' Save starting path
        
        ' Is this a list file?
        If StrComp("lst", strExtension, vbTextCompare) = 0 Then
            
            EmptyCollection mcolProjectNames          ' Start with empty collection
            Set mcolProjectNames = New Collection     ' Instantiate collection
                        
            hFile = FreeFile                          ' Get first free file handle
            Open mstrPathAndFile For Input As #hFile  ' Open file for sequential read only
            
            Do While Not EOF(hFile)
            
                Line Input #hFile, strRecord     ' Read a record
                strRecord = TrimStr(strRecord)   ' Remove any leading or trailing blanks
                
                ' Do we have some data?
                If Len(strRecord) > 0 Then
                
                    ' Must have a period within the last
                    ' four positions to be a valid file
                    If InStr(1, Right$(strRecord, 4), ".") > 0 Then
                        
                        ' Verify this is a valid file
                        If IsPathValid(strRecord) Then
                                                        
                            ' Add record to collection
                            mcolProjectNames.Add strRecord
                                                                                       
                            ' Test first valid file only
                            If mcolProjectNames.Count = 1 Then
                                
                                ' Are these VB files?
                                mblnVisualBasic = ValidateFileType(strRecord)
            
                                If mblnVisualBasic Then
                                    ' VB - enable Autogen checkbox
                                    chkFilterOptions(0).Enabled = True
                                Else
                                    ' C - disable Autogen checkbox
                                    chkFilterOptions(0).Enabled = False
                                End If
                            End If
                        End If
                    End If
                Else
                    Exit Do  ' end of file
                End If
            Loop
            
            Close #hFile   ' Close file handle
                            
            If mcolProjectNames.Count > 0 Then
                mblnPreloadedList = True
            End If
                                                    
        Else
            ' Is this a VB file?
            mblnVisualBasic = ValidateFileType(strFileName)
         
            If mblnVisualBasic Then
                ' VB - enable Autogen checkbox
                chkFilterOptions(0).Enabled = True
            Else
                ' C - disable Autogen checkbox
                chkFilterOptions(0).Enabled = False
            End If
        
        End If
    
        ' Display file selection in textbox
        lblFilename.Caption = ShrinkToFit(mstrPathAndFile, 50)
    End If
    
    ' An unknown file type was selected
    ' or user pressed Cancel button
    If mblnUnknown Then
        lblFilename.Caption = vbNullString   ' Empty textbox
        mstrPathAndFile = vbNullString       ' Empty variables
        mblnUnknown = False                  ' reset flags
        mblnVisualBasic = False
        mblnPreloadedList = False
        EmptyCollection mcolProjectNames     ' Empty this collection
    End If
            
    vntFilter = Empty         ' Empty variants when not needed
    
End Sub

' ***************************************************************************
' Routine:       BrowseForReports
'
' Description:   Opens File Open dialog box so user can browse for a
'                former report file.
'
' Returns:       Path and filename
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-NOV-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Function BrowseForReports() As String

    Dim vntFilter   As Variant
    Dim strFileName As String
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
        
    BrowseForReports = vbNullString
    strFileName = vbNullString
    vntFilter = Empty
    
    ' Set selection filters
    vntFilter = "Report files (*_LOC.txt)" & vbNullChar & "*_LOC.txt" & vbNullChar & vbNullChar
    
    ' Display "File Open" dialog box
    frmMain.Hide                     ' Hide main form
    SetStartingFolder gstrLastPath   ' Optional information
    strFileName = ShowFileOpen(frmMain, "Browse for LOC reports", vntFilter)
    vntFilter = Empty                ' Empty variants when not needed
    frmMain.Show                     ' Show main form
    
    strFileName = TrimStr(strFileName)   ' Remove unwanted leading/trailing characters

    BrowseForReports = strFileName    ' Return path\file name
    
    
End Function

Private Sub Form_Load()

    Dim intIndex As Integer
    
    Set mobjKeyEdit = New cKeyEdit   ' Instantiate class object
        
    mblnUnknown = False
    mblnVisualBasic = False
    mblnPreloadedList = False
    
    mblnIncludeAutoGen = True   ' Preset to TRUE will reset to FALSE below
    mblnIncludeComments = True
    mblnIncludeBlankLines = True
    
    EnableControls   ' Enable buttons and checkboxes
    chkDisplayRpt_Click
    
    ' Display the form
    With frmMain
        .Caption = gstrVersion
        
        ' If operating system is Windows 8 or newer
        ' form caption is centered automatically
        If Not gblnCenterCaption Then
            mobjKeyEdit.CenterCaption frmMain   ' Manually center form caption
        End If
    
        .lblAuthor.Caption = AUTHOR_NAME
        .lblMsg(4).Caption = "This application will not generate valid matrix " & _
                             "data for 16-bit or .NET projects"
        .lblDisclaimer.Caption = "This software is provided without any " & _
                                 "warrantees or guarantees implied or intended."
                                         
        ' set all filter switches to FALSE
        For intIndex = 0 To .chkFilterOptions.Count - 1
            chkFilterOptions_Click intIndex
        Next intIndex
      
        .txtOptTitle.Text = gstrOptTitle
        .cmdChoice(1).Caption = "Start"
        .cmdChoice(1).ToolTipText = "Start processing report"
             
        ' Center form on screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Show vbModeless  ' less flicker this way
        .Refresh
    End With
  
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If gblnStopProcessing Then
        gblnStopProcessing = True        ' Set global flag

        DoEvents
        
        TerminateProgram  ' Stop and close this application
    End If
    
End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Public Sub UpdateFileDisplay(ByVal strMsg As String)
    
    ' Display current file being processed
    lblCurrentFile.Caption = strMsg

End Sub
    
Private Sub txtOptTitle_GotFocus()

    ' Highlight all the text in the box
    mobjKeyEdit.TextBoxFocus txtOptTitle
    
End Sub

Private Sub txtOptTitle_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Process any key combinations
    mobjKeyEdit.TextBoxKeyDown txtOptTitle, KeyCode, Shift

End Sub

Private Sub txtOptTitle_KeyPress(KeyAscii As Integer)

    ' Save alphanumeric characters including spaces
    mobjKeyEdit.ProcessAlphaNumeric KeyAscii

End Sub




