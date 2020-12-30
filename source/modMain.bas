Attribute VB_Name = "modMain"
' ***************************************************************************
' Module:        modMain
'
' Description:   This is a generic module I use to start and stop an
'                application
'
' IMPORTANT:     Must have access to modTrimStr.bas
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced FileExists() and PathExists() routines with
'              IsPathValid() routine.
' 26-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Deleted RemoveTrailingNulls() routine from this module.
'              - Changed call to RemoveTrailingNulls() to TrimStr module
'                due to speed and accuracy.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Global constants
' ***************************************************************************
  Public Const AUTHOR_NAME           As String = "Kenneth Ives"
  Public Const SUPPORT_EMAIL         As String = "kenaso@tx.rr.com"
  Public Const PGM_NAME              As String = "Count Lines of Code"
  Public Const TMP_PREFIX            As String = "~ki"
  Public Const MAX_SIZE              As Long = 260

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const MODULE_NAME          As String = "modMain"
  Private Const ERROR_ALREADY_EXISTS As Long = 183&
  Private Const SWP_NOMOVE           As Long = 2     ' Do not move window
  Private Const SWP_NOSIZE           As Long = 1     ' Do not size window
  Private Const HWND_TOPMOST         As Long = -1    ' Bring to top and stay there
  Private Const HWND_NOTOPMOST       As Long = -2    ' Rele    Ase hold on window
  Private Const HWND_FLAGS           As Long = SWP_NOMOVE Or SWP_NOSIZE
  Private Const SW_SHOWMAXIMIZED     As Long = 3

  ' Used in ShellAndWait routine
  Private Const STATUS_PENDING            As Long = &H103&
  Private Const PROCESS_QUERY_INFORMATION As Long = &H400
  
' ***************************************************************************
' API Declares
' ***************************************************************************
  ' This is a rough translation of the GetTickCount API. The
  ' tick count of a PC is only valid for the first 49.7 days
  ' since the last reboot.  When you capture the tick count,
  ' you are capturing the total number of milliseconds elapsed
  ' since the last reboot.  The elapsed time is stored as a
  ' DWORD value. Therefore, the time will wrap around to zero
  ' if the system is run continuously for 49.7 days.
  Private Declare Function GetTickCount Lib "kernel32" () As Long

  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

  ' PathFileExists function determines whether a path to a file system
  ' object such as a file or directory is valid. Returns nonzero if the
  ' file exists.
  Private Declare Function PathFileExists Lib "shlwapi" _
          Alias "PathFileExistsA" (ByVal pszPath As String) As Long

  '  OpenProcess function returns a handle of an existing process object.
  ' If the function succeeds, the return value is an open handle of the
  ' specified process.
  Private Declare Function OpenProcess Lib "kernel32" _
          (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
          ByVal dwProcessId As Long) As Long

  ' The GetCurrentProcess function returns a pseudohandle for the current
  ' process. A pseudohandle is a special constant that is interpreted as
  ' the current process handle. The calling process can use this handle to
  ' specify its own process whenever a process handle is required. The
  ' pseudohandle need not be closed when it is no longer needed.
  Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

  ' The GetExitCodeProcess function retrieves the termination status of the
  ' specified process. If the function succeeds, the return value is nonzero.
  Private Declare Function GetExitCodeProcess Lib "kernel32" _
          (ByVal hProcess As Long, lpExitCode As Long) As Long

  ' ExitProcess function ends a process and all its threads
  ' ex:     ExitProcess GetExitCodeProcess(GetCurrentProcess, 0)
  Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

  ' The CreateMutex function creates a named or unnamed mutex object.  Used
  ' to determine if an application is active.
  Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" _
          (lpMutexAttributes As Any, ByVal bInitialOwner As Long, _
          ByVal lpName As String) As Long

  ' This function releases ownership of the specified mutex object.
  ' Finished with the search.
  Private Declare Function ReleaseMutex Lib "kernel32" _
          (ByVal hMutex As Long) As Long

  ' GetDesktopWindow function retrieves a handle to the desktop window.
  ' The desktop window covers the entire screen. The desktop window is
  ' the area on top of which other windows are painted. The return
  ' value is a handle to the desktop window.
  Private Declare Function GetDesktopWindow Lib "user32" () As Long

  ' The ShellExecute function opens or prints a specified file.  The file
  ' can be an executable file or a document file.
  Private Declare Function ShellExecute Lib "shell32.dll" _
          Alias "ShellExecuteA" (ByVal hWnd As Long, _
          ByVal lpOperation As String, ByVal lpFile As String, _
          ByVal lpParameters As String, ByVal lpDirectory As String, _
          ByVal nShowCmd As Long) As Long

  ' Always close a handle if not being used
  Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hObject As Long) As Long

  ' Truncates a path to fit within a certain number of characters by replacing
  ' path components with ellipses.
  Private Declare Function PathCompactPathEx Lib "shlwapi.dll" Alias "PathCompactPathExA" _
          (ByVal pszOut As String, ByVal pszSrc As String, _
          ByVal cchMax As Long, ByVal dwFlags As Long) As Long

  ' The FindExecutable function retrieves the name and handle to the
  ' executable (.EXE) file associated with the specified filename.
  ' Returns a value greater than 32 if successful.
  Private Declare Function FindExecutable Lib "shell32.dll" _
          Alias "FindExecutableA" (ByVal lpFile As String, _
          ByVal lpDirectory As String, ByVal lpResult As String) As Long

  ' Changes the size, position, and Z order of a child, pop-up, or top-level
  ' window. These windows are ordered according to their appearance on the
  ' screen. The topmost window receives the highest rank and is the first
  ' window in the Z order.
  Private Declare Function SetWindowPos Lib "user32" _
          (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
          ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
          ByVal cy As Long, ByVal wFlags As Long) As Long

' ***************************************************************************
' Global Variables
'                    +-------------- Global level designator
'                    |  +----------- Data type (String)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   g str Version
' Variable name:     gstrVersion
' ***************************************************************************
  Public gblnDisplayRpt    As Boolean
  Public gblnCenterCaption As Boolean
  Public gstrVersion       As String
  Public gstrLastPath      As String
  Public gstrTempPath      As String
  Public gstrOptTitle      As String

' ***************************************************************************
' Module Variables
'                    +-------------- Module level designator
'                    |  +----------- Data type (Boolean)
'                    |  |     |----- Variable subname
'                    - --- ---------------
' Naming standard:   m bln IDE_Environment
' Variable name:     mblnIDE_Environment
' ***************************************************************************
  Private mblnIDE_Environment As Boolean



' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       Main
'
' Description:   This is a generic routine to start an application
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub Main()

    Dim lngMajorVer As Long
    Dim lngMinorVer As Long
    Dim objManifest As cManifest
    Dim objOperSys  As cOperSystem
    
    Const KEY_VALUE    As String = "~ RUNASADMIN"
    Const ROUTINE_NAME As String = "Main"
    
    On Error Resume Next
    ChDrive App.Path
    ChDir App.Path
    On Error GoTo 0
    
    On Error GoTo Main_Error

    ' See if there is another instance of this program
    ' running.  The parameter being passed is the name
    ' of this executable without the EXE extension.
    If Not AlreadyRunning(App.EXEName) Then
        
        Set objOperSys = New cOperSystem   ' Instantiate class object
        With objOperSys
            gblnCenterCaption = .bCenterCaption
            lngMajorVer = CLng(.MajorVersion)
            lngMinorVer = CLng(.MinorVersion)
        End With
        Set objOperSys = Nothing   ' Free class object from memory
        
        ' Activate manifest file
        Set objManifest = New cManifest   ' Instantiate class object
        With objManifest
            .MajorVersion = lngMajorVer   ' Save OS major version
            .MinorVersion = lngMinorVer   ' Save OS minor version
            .InitComctl32                 ' Create manifest file
        End With
        Set objManifest = Nothing         ' Free class object from memory
        
        gstrVersion = PGM_NAME & " v" & App.Major & "." & App.Minor & "." & App.Revision
        gblnStopProcessing = False   ' preset global stop flag

        ' Read registry to get last path visited.
        ' HKEY_CURRENT_USER\Software\VB and VBA Program Settings\CountLinesOfCode
        gstrLastPath = GetSetting(App.EXEName, "Settings", "LastPath", "C:\")
        gstrOptTitle = GetSetting(App.EXEName, "Settings", "OptTiTle", "")
        gblnDisplayRpt = CBool(GetSetting(App.EXEName, "Settings", "DisplayRpt", 1))

        ' NOTE:  GetSetting() "Default" option does
        '        not seem to work in Windows 8
        If gblnCenterCaption Then
        
            If Len(gstrLastPath) = 0 Then
                gstrLastPath = "C:\"
            End If
    
            If Len(gblnDisplayRpt) = 0 Then
                gblnDisplayRpt = True
            End If
        
        End If

        gstrTempPath = vbNullString   ' Empty temp path
        Load frmMain        ' Load main form

    End If

Main_CleanUp:
    On Error GoTo 0
    Exit Sub

Main_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume Main_CleanUp

End Sub

' ***************************************************************************
' Routine:       TerminateProgram
'
' Description:   This routine will perform the shutdown process for this
'                application.  The proper sequence to follow is:
'
'                    1.  Deactivate and free from memory all global objects
'                        or classes
'                    2.  Verify there are no file handles left open
'                    3.  Deactivate and free from memory all form objects
'                    4.  Shut this application down
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub TerminateProgram()

    ' Save data to registry
    ' HKEY_CURRENT_USER\Software\VB and VBA Program Settings\CountLinesOfCode
    SaveSetting App.EXEName, "Settings", "LastPath", gstrLastPath
    SaveSetting App.EXEName, "Settings", "OptTiTle", gstrOptTitle
    SaveSetting App.EXEName, "Settings", "DisplayRpt", Val(gblnDisplayRpt)

    ' Free any global objects from memory.
    ' EXAMPLE:    Set gobjFSO = Nothing

    Close           ' Close all files opened by this application
    UnloadAllForms  ' Unload any forms from memory

    ' While in the VB IDE (VB Integrated Developement Environment),
    ' do not call ExitProcess API.  ExitProcess API will close all
    ' processes associated with this application including the IDE.
    ' No changes will be retained that were not previously saved.
    If mblnIDE_Environment Then
        End    ' Terminate this application while in the VB IDE
    Else
        ' Close running application gracefully
        ExitProcess GetExitCodeProcess(GetCurrentProcess, 0)
    End If

End Sub

' ***************************************************************************
' Routine:       UnloadAllForms
'
' Description:   Unload all active forms associated with this application.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub UnloadAllForms()

    Dim frm As Form
    Dim ctl As Control

    ' Loop thru all active forms
    ' associated with this application
    For Each frm In Forms

        frm.Hide            ' hide selected form

        ' free all controls from memory
        For Each ctl In frm.Controls
            Set ctl = Nothing
        Next ctl

        Unload frm          ' deactivate form object
        Set frm = Nothing   ' free form object from memory
                            ' (prevents memory fragmenting)
    Next frm

End Sub

' ***************************************************************************
' Routine:       FindRequiredFile
'
' Description:   Test to see if a required file is in the application folder
'                or in any of the folders in the PATH and WINDIR environment
'                variables.
'
'                If the required file is a registered file, access registry:
'
'                    Ref:  strFilename = "excel.exe"
'                         "Path" is subkey name with complete path w/o filename
'
'                    HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & strFilename, "Path"
'
' Syntax:        FindRequiredFile "msinfo32.exe", strPathFile
'
' Parameters:    strFilename - name of the file without path information
'                strFullPath - Optional - If found then the fully qualified
'                     path and filename are returned
'
' Returns:       TRUE  - Found the required file
'                FALSE - File could not be found
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 04-Apr-2009  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' 29-Dec-2013  Kenneth Ives  kenaso@tx.rr.com
'              Force a search of Windows and System32 folders if file is not
'              found.  Sometimes "PATH" variable becomes corrupted and
'              Windows is not in the "PATH".
' 08-May-2014  Kenneth Ives  kenaso@tx.rr.com
'              Add search of Windows, System32 and SysWOW64 folders
' 18-Sep-2014  Kenneth Ives  kenaso@tx.rr.com
'              Updated search logic
' ***************************************************************************
Public Function FindRequiredFile(ByVal strFileName As String, _
                        Optional ByRef strFullPath As String) As Boolean

    Dim blnFoundit       As Boolean      ' Flag (TRUE if found file else FALSE)
    Dim lngCount         As Long         ' String pointer position
    Dim lngIndex         As Long         ' Array index pointer
    Dim strPath          As String       ' Fully qualified search path
    Dim strMsgFmt        As String       ' Format each message line
    Dim strDosPath       As String       ' DOS environment variable
    Dim strSearched      As String       ' List of searched folders (will be displayed if not found)
    Dim strWindowsFolder As String       ' Windows folder
    Dim astrPath()       As String       ' List of folders to be searched
    
    On Error GoTo FindRequiredFile_Error

    strFullPath = vbNullString   ' Verify empty variables
    strSearched = vbNullString
    
    strMsgFmt = "!" & String$(70, "@")   ' Left justify data
    blnFoundit = False                   ' Preset flag to FALSE

    strPath = QualifyPath(App.Path)      ' Add trailing backslash to application folder
                        
    ' Add current path, without file name, to list of searched folders
    strSearched = Format$(strPath, strMsgFmt) & vbNewLine

    strPath = strPath & strFileName      ' Append file name to path
    blnFoundit = IsPathValid(strPath)    ' Check selected folder
                            
    ' Capture DOS environment variable 'PATH' and
    ' perform a search of the various folders
    If Not blnFoundit Then
    
        ' Capture DOS environment variable 'PATH' statement
        strDosPath = TrimStr(Environ$("PATH"))

        If Len(strDosPath) > 0 Then
            
            Erase astrPath()   ' Start with empty array
            lngCount = 0       ' Initialize array counter
            
            strDosPath = QualifyPath(strDosPath, ";")   ' Add trailing semi-colon
            astrPath() = Split(strDosPath, ";")         ' Load paths into array
            lngCount = UBound(astrPath)                 ' Number of entries in array
            
            For lngIndex = 0 To (lngCount - 1)
                    
                strPath = astrPath(lngIndex)     ' Capture path
                strPath = GetLongName(strPath)   ' Format long path name

                ' Verify there is some data to work with
                If Len(strPath) > 0 Then

                    strPath = QualifyPath(strPath)   ' Add trailing backslash
                        
                    ' Verify this folder exist
                    If IsPathValid(strPath) Then
                    
                        ' Add current path, without file name, to list of searched folders
                        strSearched = strSearched & Format$(strPath, strMsgFmt) & vbNewLine
    
                        strPath = strPath & strFileName     ' Append file name to path
                        blnFoundit = IsPathValid(strPath)   ' Check selected folder
                            
                        If blnFoundit Then
                            Exit For   ' Exit FOR..NEXT loop
                        End If
                    Else
                        ' Add current path, without file name, to list
                        ' of searched folders and update error message
                        strSearched = strSearched & Format$(strPath & " (Folder does not exist)", strMsgFmt) & vbNewLine
                    End If
                End If
                
            Next lngIndex
        Else
            ' Add current path, without file name, to list
            ' of searched folders and update error message
            strSearched = strSearched & Format$(Chr$(34) & "PATH" & Chr$(34) & _
                          " environment variable does not exists.", strMsgFmt) & vbNewLine
        End If
    End If
    
    ' Capture DOS environment variable 'WINDIR' and
    ' perform a search of Windows, System32 and SysWOW64
    ' folders.  Sometimes these folders are not in the
    ' 'PATH" variable and must be searched manually.
    If Not blnFoundit Then
    
        strPath = vbNullString           ' Verify empty variables
        strWindowsFolder = vbNullString

        ' Capture DOS environment variable "WINDIR" statement.
        ' This is the root folder for Windows.
        strWindowsFolder = TrimStr(Environ$("WINDIR"))
            
        ' See if "WINDIR" variable exist
        If Len(strWindowsFolder) > 0 Then
        
            ' Prepare to search Windows main folder (ex:  C:\Windows\)
            ' Format long folder name and add trailing backslash
            strWindowsFolder = QualifyPath(GetLongName(strWindowsFolder))
            
            ' Add current path, without file name, to list of searched folders
            strSearched = strSearched & Format$(strPath, strMsgFmt) & vbNewLine
                
            strPath = strWindowsFolder & strFileName   ' Append file name to path
            blnFoundit = IsPathValid(strPath)          ' Check Windows folder
                
            If Not blnFoundit Then
                    
                ' Prepare to search System32 folder (ex:  C:\Windows\System32\)
                strPath = QualifyPath(strWindowsFolder & "System32")

                ' Add current path, without file name, to list of searched folders
                strSearched = strSearched & Format$(strPath, strMsgFmt) & vbNewLine
                
                strPath = strPath & strFileName     ' Append file name to path
                blnFoundit = IsPathValid(strPath)   ' Check System32 folder
            End If
            
            If Not blnFoundit Then
                    
                ' Prepare to search SysWOW64 folder (ex:  C:\Windows\SysWOW64\)
                strPath = QualifyPath(strWindowsFolder & "SysWOW64")
                
                ' Add current path, without file name, to list of searched folders
                strSearched = strSearched & Format$(strPath, strMsgFmt) & vbNewLine
                
                strPath = strPath & strFileName     ' Append file name to path
                blnFoundit = IsPathValid(strPath)   ' Check SysWOW64 folder
            End If
        Else
            ' Add current path, without file name, to list
            ' of searched folders and update error message
            strSearched = strSearched & Format$(Chr$(34) & "WINDIR" & Chr$(34) & _
                          " environment variable does not exists.", strMsgFmt) & vbNewLine
        End If
    End If
    
FindRequiredFile_CleanUp:
    If blnFoundit Then
        strFullPath = strPath   ' Return full path/filename
    Else
        InfoMsg Format$("A required file that supports this application cannot be found.", strMsgFmt) & _
                vbNewLine & vbNewLine & _
                Format$(Chr$(34) & UCase$(strFileName) & Chr$(34) & _
                " not in any of these folders:", strMsgFmt) & vbNewLine & vbNewLine & _
                strSearched, "File not found"
    End If

    FindRequiredFile = blnFoundit   ' Set status flag
    strSearched = vbNullString      ' Empty variable
    Erase astrPath()                ' Empty array
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

FindRequiredFile_Error:
    If Err.Number <> 0 Then
        Err.Clear
    End If
    
    Resume FindRequiredFile_CleanUp

End Function

' ***************************************************************************
' Procedure:     GetLongName
'
' Description:   The Dir() function can be used to return a long filename
'                but it does not include path information. By parsing a
'                given short path/filename into its constituent directories,
'                you can use the Dir() function to build a long path/filename.
'
' Example:       Syntax:
'                   GetLongName C:\DOCUME~1\KENASO\LOCALS~1\Temp\~ki6A.tmp
'
'                Returns:
'                   "C:\Documents and Settings\Kenaso\Local Settings\Temp\~ki6A.tmp"
'
' Parameters:    strShortName - Path or file name to be converted.
'
' Returns:       A readable path or file name.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Jul-2004  http://support.microsoft.com/kb/154822
'              "How To Get a Long Filename from a Short Filename"
' 09-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 09-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added removal of all double quotes prior to formatting
' ***************************************************************************
Public Function GetLongName(ByVal strShortName As String) As String

    Dim strTemp     As String
    Dim strLongName As String
    Dim intPosition As Integer

    On Error Resume Next

    GetLongName = vbNullString
    strLongName = vbNullString

    ' Remove all double quotes
    strShortName = Replace(strShortName, Chr$(34), vbNullString)

    ' Add a backslash to short name, if needed,
    ' to prevent Instr() function from failing.
    strShortName = QualifyPath(strShortName)

    ' Start at position 4 so as to ignore
    ' "[Drive Letter]:\" characters.
    intPosition = InStr(4, strShortName, "\")

    ' Pull out each string between
    ' backslash character for conversion.
    Do While intPosition > 0

        strTemp = vbNullString   ' Init variable

        ' Progressively parse path to verify
        ' each portion does exist and
        ' capture its expanded version.
        strTemp = Dir$(Left$(strShortName, intPosition - 1), _
                       vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbDirectory)

        ' If no data then exit this loop
        If Len(TrimStr(strTemp)) = 0 Then
            strShortName = vbNullString
            strLongName = vbNullString
            Exit Do   ' exit DO..LOOP
        End If

        ' Append new elongated portion to output string
        ' after converting it to propercase format.
        strLongName = strLongName & "\" & StrConv(strTemp, vbProperCase)

        ' Find next backslash
        intPosition = InStr(intPosition + 1, strShortName, "\")

    Loop

GetLongName_CleanUp:
    If Len(strShortName & strLongName) > 0 Then
        GetLongName = UCase$(Left$(strShortName, 2)) & strLongName
    Else
        GetLongName = "[Unknown]"
    End If

    On Error GoTo 0   ' Nullify this error trap

End Function

' ***************************************************************************
' Routine:       IsPathValid
'
' Description:   Determines whether a path to a file system object such as
'                a file or directory is valid. This function tests the
'                validity of the path. A path specified by Universal Naming
'                Convention (UNC) is limited to a file only; that is,
'                \\server\share\file is permitted. A UNC path to a server
'                or server share is not permitted; that is, \\server or
'                \\server\share. This function returns FALSE if a mounted
'                remote drive is out of service.
'
'                Requires Version 4.71 and later of Shlwapi.dll
'                Shlwapi.dll first shipped with Internet Explorer 4.0
'
' Reference:     http://msdn.microsoft.com/en-us/library/bb773584(v=vs.85).aspx
'
' Syntax:        IsPathValid("C:\Program Files\Desktop.ini")
'
' Parameters:    strName - Path or filename to be queried.
'
' Returns:       True or False
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function IsPathValid(ByVal strName As String) As Boolean

   IsPathValid = CBool(PathFileExists(strName))

End Function

' ***************************************************************************
' Routine:       AlreadyRunning
'
' Description:   This routine will determine if an application is already
'                active, whether it be hidden, minimized, or displayed.
'
' Parameters:    strTitle - partial/full name of application
'
' Returns:       TRUE  - Currently active
'                FALSE - Inactive
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-DEC-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function AlreadyRunning(ByVal strAppTitle As String) As Boolean

    Dim hMutex As Long

    Const ROUTINE_NAME As String = "AlreadyRunning"

    On Error GoTo AlreadyRunning_Error

    mblnIDE_Environment = False  ' preset flags to FALSE
    AlreadyRunning = False

    ' Are we in VB development environment?
    mblnIDE_Environment = IsVB_IDE

    ' Multiple instances can be run while
    ' in the VB IDE but not as an EXE
    If Not mblnIDE_Environment Then

        ' Try to create a new Mutex handle
        hMutex = CreateMutex(ByVal 0&, 1, strAppTitle)

        ' Did mutex handle already exist?
        If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then

            ReleaseMutex hMutex     ' Release Mutex handle from memory
            CloseHandle hMutex      ' Close the Mutex handle
            Err.Clear               ' Clear any errors
            AlreadyRunning = True   ' prior version already active
        End If
    End If

AlreadyRunning_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

AlreadyRunning_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume AlreadyRunning_CleanUp

End Function

Private Function IsVB_IDE() As Boolean

    ' 09-16-2000  Michael Culley  m_culley@one.net.au
    '             http://forums.devx.com/showthread.php?t=37676
    '
    ' Set DebugMode flag.  Call can only be successful if
    ' in the VB Integrated Development Environment (IDE).
    Debug.Assert SetTrue(IsVB_IDE) Or True

End Function

Private Function SetTrue(ByRef blnValue As Boolean) As Boolean

    ' 09-16-2000  Michael Culley  m_culley@one.net.au
    '             http://forums.devx.com/showthread.php?t=37676
    '
    ' Can only be set to TRUE if Debug.Assert call is
    ' successful.  Call can only be successful if in
    ' the VB Integrated Development Environment (IDE).
    blnValue = True

End Function

' ***************************************************************************
' Routine:       QualifyPath
'
' Description:   Adds a trailing character to the path, if missing.
'
' Parameters:    strPath - Current folder being processed.
'                strChar - Optional - Specific character to append.
'                          Default = "\"
'
' Returns:       Fully qualified path with a specific trailing character
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch
'              http://vbnet.mvps.org/index.html
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function QualifyPath(ByVal strPath As String, _
                   Optional ByVal strChar As String = "\") As String

    strPath = TrimStr(strPath)

    If StrComp(Right$(strPath, 1), strChar, vbTextCompare) = 0 Then
        QualifyPath = strPath
    Else
        QualifyPath = strPath & strChar
    End If

End Function

' ***************************************************************************
' Routine:       UnQualifyPath
'
' Description:   Removes a trailing character from the path
'
' Parameters:    strPath - Current folder being processed.
'                strChar - Optional - Specific character to remove
'                          Default = "\"
'
' Returns:       Fully qualified path without a specific trailing character
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch
'              http://vbnet.mvps.org/index.html
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function UnQualifyPath(ByVal strPath As String, _
                     Optional ByVal strChar As String = "\") As String

    strPath = TrimStr(strPath)

    If StrComp(Right$(strPath, 1), strChar, vbTextCompare) = 0 Then
        UnQualifyPath = Left$(strPath, Len(strPath) - 1)
    Else
        UnQualifyPath = strPath
    End If

End Function

' ***************************************************************************
' Routine:       SendEmail
'
' Description:   When the email hyperlink is clicked, this routine will fire.
'                It will create a new email message with the author's name in
'                the "To:" box and the name and version of the application
'                on the "Subject:" line.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 23-Feb-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 25-Apr-2015  Kenneth Ives  kenaso@tx.rr.com
'              Added reference to GetDesktopWindow() API
' ***************************************************************************
Public Sub SendEmail()

    Dim strMail As String
    
    ' Create email heading for user
    strMail = "mailto:" & SUPPORT_EMAIL & "?subject=" & gstrVersion

    ' Call ShellExecute() API to create an email to the author
    ShellExecute GetDesktopWindow(), "open", strMail, _
                 vbNullString, vbNullString, vbNormalFocus
    
End Sub

' ***************************************************************************
' Routine:       ShrinkToFit
'
' Description:   This routine creates the ellipsed string by specifying
'                the size of the desired string in characters.  Adds
'                ellipses to a file path whose maximum length is specified
'                in characters.
'
' Parameters:    strPath - Path to be resized for display
'                intMaxLength - Maximum length of the return string
'
' Returns:       Resized path
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 20-May-2004  Randy Birch
'              http://vbnet.mvps.org/code/fileapi/pathcompactpathex.htm
' 22-Jun-2004  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function ShrinkToFit(ByVal strPath As String, _
                            ByVal intMaxLength As Integer) As String

    Dim strBuffer As String

    strPath = TrimStr(strPath)

    ' See if ellipses need to be inserted into the path
    If Len(strPath) <= intMaxLength Then
        ShrinkToFit = strPath
        Exit Function
    End If

    ' intMaxLength is the maximum number of characters to be contained in the
    ' new string, **including the terminating NULL character**. For example,
    ' if intMaxLength = 8, the resulting string would contain a maximum of
    ' seven characters plus the terminating null.
    '
    ' Because of this, add 1 to the value passed as intMaxLength to ensure
    ' the resulting string is the size requested.
    intMaxLength = intMaxLength + 1
    strBuffer = Space$(MAX_SIZE)
    PathCompactPathEx strBuffer, strPath, intMaxLength, 0&

    ' Return the readjusted data string
    ShrinkToFit = TrimStr(strBuffer)

End Function

' ***************************************************************************
' Routine:       IsArrayInitialized
'
' Description:   This is an ArrPtr function that determines if the passed
'                array is initialized, and if so will return the pointer
'                to the safearray header. If the array is not initialized,
'                it will return zero. Normally you need to declare a VarPtr
'                alias into msvbvm50.dll or msvbvm60.dll depending on the
'                VB version, but this function will work with vb5 or vb6.
'                It is handy to test if the array is initialized as the
'                return value is non-zero.  Use CBool to convert the return
'                value into a boolean value.
'
'                This function returns a pointer to the SAFEARRAY header of
'                any Visual Basic array, including a Visual Basic string
'                array. Substitutes both ArrPtr and StrArrPtr. This function
'                will work with vb5 or vb6 without modification.
'
'                ex:  If CBool(IsArrayInitialized(array_being_tested)) Then ...
'
' Parameters:    vntData - Data to be evaluated
'
' Returns:       Zero     - Bad data (FALSE)
'                Non-zero - Good data (TRUE)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 30-Mar-2008  RD Edwards
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=69970
' ***************************************************************************
Public Function IsArrayInitialized(ByVal avntData As Variant) As Long

    Dim intDataType As Integer   ' Variable must be a short integer

    On Error GoTo IsArrayInitialized_Exit

    IsArrayInitialized = 0  ' preset to FALSE

    ' Get the real VarType of the argument, this is similar
    ' to VarType(), but returns also the VT_BYREF bit
    CopyMemory intDataType, avntData, 2&

    ' if a valid array was passed
    If (intDataType And vbArray) = vbArray Then

        ' get the address of the SAFEARRAY descriptor
        ' stored in the second half of the Variant
        ' parameter that has received the array.
        ' Thanks to Francesco Balena and Monte Hansen.
        CopyMemory IsArrayInitialized, ByVal VarPtr(avntData) + 8&, 4&

    End If

IsArrayInitialized_Exit:
    On Error GoTo 0   ' Nullify this error trap

End Function

' ***************************************************************************
' Routine:       EmptyCollection
'
' Description:   Properly empty and deactivate a collection
'
' Parameters:    colData - Collection to be processed
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub EmptyCollection(ByRef colData As Collection)

    ' Has collection been deactivated?
    If colData Is Nothing Then
        Exit Sub
    End If

    ' Is the collection empty?
    Do While colData.Count > 0

        ' Parse backwards thru collection and delete data.
        ' Backwards parsing prevents a collection from
        ' having to reindex itself after each data removal.
        colData.Remove colData.Count
    Loop

    ' Free collection object from memory
    Set colData = Nothing

End Sub

Public Sub AlwaysOnTop(ByVal blnOnTop As Boolean)

    ' This routine uses an argument to determine whether
    ' to make specified form always on top or not
    '
    ' Syntax:  AlwaysOnTop form_handle, True   ' On top of all other windows
    '          AlwaysOnTop form_handle, False  ' Not on top
    '
    On Error GoTo AlwaysOnTop_Error

    If blnOnTop Then
        ' stay as topmost window
        SetWindowPos frmMain.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, HWND_FLAGS
    Else
        ' not on top anymore
        SetWindowPos frmMain.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, HWND_FLAGS
    End If

AlwaysOnTop_CleanUp:
    On Error GoTo 0
    Exit Sub

AlwaysOnTop_Error:
    ErrorMsg MODULE_NAME, "AlwaysOnTop", Err.Description
    Resume AlwaysOnTop_CleanUp

End Sub

' ***************************************************************************
' Routine:       DisplayFile
'
' Description:   Display a text based file using default text editor.
'
' Parameters:    strFile - Path and file name to be opened
'                frmName - Calling form
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-Jan-2011  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub DisplayFile(ByVal strFile As String, _
                       ByRef frmName As Form)

    Dim lngRetCode     As Long
    Dim strApplication As String

    Screen.MousePointer = vbHourglass   ' Change mouse pointer to hourglass
    strApplication = Space$(MAX_SIZE)

    ' See if double quotes need to be added
    If InStr(1, strFile, Chr$(32)) > 0 Then
        strFile = Chr$(34) & strFile & Chr$(34)
    End If
    
    ' Retrieve name of executable
    ' associated with this file extension
    lngRetCode = FindExecutable(strFile, vbNullString, strApplication)

    If lngRetCode > 32 Then
        strApplication = TrimStr(strApplication)
    Else
        strApplication = "notepad.exe"
    End If

    ' Open default text file viewer
    If Len(strApplication) > 0 Then
        ShellExecute frmName.hWnd, "open", strApplication, strFile, _
                     vbNullString, SW_SHOWMAXIMIZED
    End If

    Screen.MousePointer = vbNormal   ' Change mouse pointer back to normal
    strFile = vbNullString           ' Empty variable
    
End Sub

' ***************************************************************************
' Routine:       ShellAndWait
'
' Description:   Wait for a shelled application to close before continuing.
'
' Parameters:    strCmdLine - Data to be executed via the Shell process
'                lngDisplayMode- Optional - How to display shelled window
'                    Default = vbNormalFocus = 1
'                lngAttempts - Optional - Number of tries before forcing
'                    an exit from this routine.  Default = 3
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 26-Dec-2006  Randy Birch
'              http://vbnet.mvps.org/code/faq/getexitcprocess.htm
' 10-Sep-2014  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Sub ShellAndWait(ByVal strCmdLine As String, _
               Optional ByVal lngDisplayMode As Long = vbNormalFocus, _
               Optional ByVal lngAttempts As Long = 3)

    Dim lngExitCode  As Long
    Dim lngProcHwnd  As Long
    Dim lngProcessID As Long

    On Error GoTo ShellAndWait_CleanUp
    
    ' Start a shelled process and hide window
    lngProcessID = Shell(strCmdLine, lngDisplayMode)
    
    ' Capture shelled process handle
    lngProcHwnd = OpenProcess(PROCESS_QUERY_INFORMATION, False, lngProcessID)

    ' The DoEvents statement relinquishes your application's control
    ' to allow Windows to process any pending messages or events for
    ' your application or any other running process.  Without this,
    ' your application will appear to lock up as the DO...LOOP
    ' essentially "grabs control" of the system.
    Do
        Call GetExitCodeProcess(lngProcHwnd, lngExitCode)
        DoEvents
        
        ' See if exit code designates FINISHED
        If lngExitCode <> STATUS_PENDING Then
            Exit Do   ' exit DO..LOOP
        Else
            Wait 1000                       ' Pause for one second
            lngAttempts = lngAttempts - 1   ' Decrement attempt counter
            DoEvents

            If lngAttempts < 1 Then
                Exit Do   ' exit DO..LOOP
            End If
        End If
    Loop

ShellAndWait_CleanUp:
    Call CloseHandle(lngProcHwnd)   ' Always release handle when not needed
    DoEvents
    On Error GoTo 0                 ' Nullify this error trap

End Sub

Public Sub Wait(ByVal lngMilliseconds As Long)

    Dim lngPause As Long

    ' Calculate a pause
    lngPause = GetTickCount() + lngMilliseconds

    Do
        DoEvents
    Loop While lngPause > GetTickCount()

End Sub
