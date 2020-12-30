Attribute VB_Name = "modFileFind"
Option Explicit

' ***************************************************************************
' Constants - Miscellaneous
' ***************************************************************************
  Private Const ALL_FILES            As String = "*.*"
  Private Const INVALID_HANDLE_VALUE As Long = -1
  
' ***************************************************************************
' Type structures used for searching folders
' ***************************************************************************
  ' The FILETIME structure is a 64-bit value representing the number of
  ' 100-nanosecond intervals since January 1, 1601.
  Private Type FILETIME
       dwLowDateTime  As Long
       dwHighDateTime As Long
  End Type

  ' The WIN32_FIND_DATA structure describes a file found by the
  ' FindFirstFile or FindNextFile function.
  Private Type WIN32_FIND_DATA
      dwFileAttributes As Long
      ftCreationTime   As FILETIME
      ftLastAccessTime As FILETIME
      ftLastWriteTime  As FILETIME
      nFileSizeHigh    As Long
      nFileSizeLow     As Long
      dwReserved0      As Long
      dwReserved1      As Long
      cFilename        As String * MAX_SIZE
      cAlternate       As String * 14
  End Type

  Private Type FILE_PARAMS
      bRecurse  As Boolean
      nCount    As Long
      nSearched As Long
      sPattern  As String
      sFileRoot As String
      sFileName As String
  End Type

' ***************************************************************************
' API Declares used for searching folders
' ***************************************************************************
  ' The FindClose function closes the specified search handle. The
  ' FindFirstFile and FindNextFile functions use the search handle
  ' to locate files with names that match a given name.
  Private Declare Function FindClose Lib "kernel32" _
          (ByVal hFindFile As Long) As Long
   
  ' The FindFirstFile function searches a directory for a file whose
  ' name matches the specified filename. FindFirstFile examines
  ' subdirectory names as well as filenames.
  Private Declare Function FindFirstFile Lib "kernel32" _
          Alias "FindFirstFileA" (ByVal lpFileName As String, _
          lpFindFileData As WIN32_FIND_DATA) As Long
   
  ' The FindNextFile function continues a file search from a previous
  ' call to the FindFirstFile function.
  Private Declare Function FindNextFile Lib "kernel32" _
          Alias "FindNextFileA" (ByVal hFindFile As Long, _
          lpFindFileData As WIN32_FIND_DATA) As Long

  ' perform pattern matching
  Private Declare Function PathMatchSpec Lib "shlwapi" _
          Alias "PathMatchSpecW" _
          (ByVal pszFileParam As Long, ByVal pszSpec As Long) As Long

  '
  Private Declare Function PathIsDirectory Lib "shlwapi" _
          Alias "PathIsDirectoryA" _
          (ByVal pszPath As String) As Long
  
  ' ZeroMemory function fills a block of memory with zeros.
  Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" _
          (Destination As Any, ByVal Length As Long)

' ***************************************************************************
' Module Variables
'                    +-------------- Module level designator
'                    |  +----------- Data type (Type Structure)
'                    |  |     |----- Variable subname
'                    - --- ---------------
' Naming standard:   m typ FP
' Variable name:     mtypFP
' ***************************************************************************
  Private mtypFP      As FILE_PARAMS      ' holds search parameters
  Private mblnFoundit As Boolean          ' TRUE if file is found
  
' ***************************************************************************
' Routine:       FindFile
'
' Description:   Find a specific file
'
' Parameters:    strFolder - Start the search at this folder
'                strFilename - file name to search for
'
' Returns:       Fully qualified path and file name, if found
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch http://www.mvps.org/vbnet/index.html
'              Original routine
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function FindFile(ByVal strFolder As String, _
                         ByVal strFileName As String) As String

    Dim intCount As Integer
    Dim intStart As Integer
    Dim intIndex As Integer
    Dim strTemp  As String
        
    intStart = 1
    intCount = 1
    strTemp = strFileName
    
    ' Determine the path depth
    Do
        intStart = InStr(1, strTemp, "\")
        If intStart > 0 Then
            intCount = intCount + 1
            strTemp = Mid$(strTemp, intStart + 1)
        Else
            Exit Do   ' exit Do..Loop
        End If
    Loop
    
    ' Adjust the starting folder name
    If intCount > 0 Then
        For intIndex = 1 To intCount
            intStart = InStrRev(strFolder, "\", Len(strFolder))
            strFolder = Left$(strFolder, intStart - 1)
        Next intIndex
    End If
    
    ' Initialize file parameter structure
    With mtypFP
         .sFileRoot = QualifyPath(strFolder)  ' start path
         .sPattern = strTemp                  ' Search item
         .sFileName = vbNullString                      ' returned path & filename
         .bRecurse = True                     ' recursive search
         .nCount = 0                          ' results
         .nSearched = 0                       ' results
    End With
    
    mblnFoundit = False
    
    SearchForFiles mtypFP.sFileRoot  ' begin search process
      
    FindFile = mtypFP.sFileName      ' return what we found
  
End Function

' ***************************************************************************
' Routine:       SearchForFiles
'
' Description:   Search for *.zip and *.exe files.  When found, process their
'                date and time stamp as per option selected.
'
' Parameters:    strPath - Current folder being processed.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch http://www.mvps.org/vbnet/index.html
'              Original routine
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Sub SearchForFiles(ByVal strPath As String)

    Dim typWFD      As WIN32_FIND_DATA  ' Folder or file data structure
    Dim hFile       As Long             ' file handle
    Dim strFileName As String           ' Name of file
    
    ' See if we found it
    DoEvents
    If mblnFoundit Then
        FindClose hFile
        Exit Sub
    End If
    
    ZeroMemory typWFD, Len(typWFD)  ' Empty type structure
    strFileName = vbNullString
    
    hFile = FindFirstFile(strPath & ALL_FILES, typWFD)   ' capture first file
    
    ' If we have a valid handle then continue processing
    If hFile <> INVALID_HANDLE_VALUE Then
        Do
            If mblnFoundit Then
                Exit Do   ' exit Do..Loop
            End If
            
            ' if a folder, and recurse specified, call method again
            If (typWFD.dwFileAttributes And vbDirectory) Then
                ' make sure this is not a VTOC identifier
                If (typWFD.cFilename <> ".") And (typWFD.cFilename <> "..") Then
                    ' Check the sub-folders if requested
                    If mtypFP.bRecurse Then
                        SearchForFiles strPath & TrimStr(typWFD.cFilename) & "\"
                    End If
                End If
            Else
                ' This must be a file.  Now see if it matches what we are looking for.
                strFileName = TrimStr(typWFD.cFilename)
                
                If MatchSpec(strFileName, mtypFP.sPattern) Then
                    mtypFP.sFileName = strPath & strFileName   ' save the full path & filename
                    mblnFoundit = True                         ' set the search flag
                    Exit Do                                    ' exit Do..Loop
                End If
            End If
        
            typWFD.cFilename = vbNullString
            strFileName = vbNullString

        Loop While FindNextFile(hFile, typWFD)
        
    End If
    
    FindClose hFile      ' always close file handles when not in use

End Sub

' ***************************************************************************
' Routine:       MatchSpec
'
' Description:   Determines if this file matches the pattern that was requested
'
' Parameters:    strFilename - name of file to be tested
'                strPattern - File extension pattern to test against
'
' Returns:       TRUE/FALSE based on test results
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch http://www.mvps.org/vbnet/index.html
'              Original routine
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function MatchSpec(ByVal strFileName As String, _
                           ByVal strPattern As String) As Boolean

    ' if searching for all files, then this is a valid filename
    If StrComp(strPattern, ALL_FILES, vbTextCompare) = 0 Then
        MatchSpec = True
    Else
        ' look for a specific pattern
        MatchSpec = CBool(PathMatchSpec(StrPtr(strFileName), StrPtr(strPattern)))
    End If
  
End Function

' ***************************************************************************
' Routine:       AvailableDriveLetters
'
' Description:   Get the list of available drive letters, each separated by
'                a null character.  (i.e.  a:\ c:\ d:\)
'
' Returns:       String of available drive letters
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function AvailableDriveLetters() As String

    Dim objFSO    As Scripting.FileSystemObject
    Dim objDrive  As Drive
    Dim objDrives As Drives

    Set objFSO = New Scripting.FileSystemObject
    Set objDrives = objFSO.Drives
    AvailableDriveLetters = vbNullString
    
    ' Get list of current drive letters
    For Each objDrive In objDrives
        AvailableDriveLetters = AvailableDriveLetters & _
                                objDrive.DriveLetter & ":\" & Chr$(0)
    Next objDrive
    
    Set objDrives = Nothing   ' Free objects from memory when not needed
    Set objDrive = Nothing
    Set objFSO = Nothing

End Function

' ***************************************************************************
' Routine:       GetPath
'
' Description:   Capture complete path up to filename.  Path must end with
'                a backslash.
'
' Parameters:    strPathFile - Path and file name
'
' Returns:       Complete path to last backslash
'
' Example:       "C:\Kens Software" = "C:\Kens Software\Gif89.dll"
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetPath(ByVal strPathFile As String) As String

    Dim objFSO As New Scripting.FileSystemObject
    GetPath = objFSO.GetParentFolderName(strPathFile)
    Set objFSO = Nothing
    
End Function

' ***************************************************************************
' Routine:       GetFilename
'
' Description:   Capture file name
'
' Parameters:    strPathFile - Path and file name
'
' Returns:       Just the file name
'
' Example:       "Gif89.dll" = "C:\Kens Software\Gif89.dll"
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetFilename(ByVal strPathFile As String) As String

    Dim objFSO As New Scripting.FileSystemObject
    GetFilename = objFSO.GetFilename(strPathFile)
    Set objFSO = Nothing
    
End Function

' ***************************************************************************
' Routine:       GetFilenameExt
'
' Description:   Capture file name extension
'
' Parameters:    strPathFile - Path and file name
'
' Returns:       File name extension
'
' Example:       "dll" = "C:\Kens Software\Gif89.dll"
'
' =========          Routine created
' ***************************************************************************
Public Function GetFilenameExt(ByVal strPathFile As String) As String

    Dim objFSO As New Scripting.FileSystemObject
    GetFilenameExt = objFSO.GetExtensionName(strPathFile)
    Set objFSO = Nothing
    
End Function

' ***************************************************************************
' Routine:       GetVersion
'
' Description:   Capture file version information
'
' Parameters:    strPathFile - Path and file name
'
' Returns:       Version information
'
' Example:       "1.0.0.1" = "C:\Kens Software\Gif89.dll"
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetVersion(ByVal strPathFile As String) As String

    Dim objFSO As New Scripting.FileSystemObject
    GetVersion = objFSO.GetFileVersion(strPathFile)
    Set objFSO = Nothing
    
End Function

' ***************************************************************************
' Routine:       IsPathAFolder
'
' Description:   Verifies that a path is a valid directory, and returns True
'                if the path is a valid directory, or False otherwise. The
'                path must exist.  If the path is a directory on the local
'                machine, PathIsDirectory returns 16 (the file attribute
'                for a folder).
'
' Parameters:    strFilePath - fully qualified path and filename
'
' Returns:       TRUE or FALSE
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-MAR-2002  Randy Birch  http://www.mvps.org/vbnet/index.htm
'              Routine created
' 03-JUL-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function IsPathAFolder(ByVal strFilePath As String) As Boolean

    Dim lngRetCode As Long
     
    ' If it is neither PathIsDirectory returns 0. If the path is a directory on
    ' a server share, PathIsDirectory returns 1.
    lngRetCode = PathIsDirectory(strFilePath)
     
    ' TRUE - This is a path    FALSE - This is a file
    IsPathAFolder = CBool((lngRetCode = vbDirectory) Or (lngRetCode = 1))

End Function

' ***************************************************************************
' Routine:       ParseData
'
' Description:   Parses backwards thru the data string searching for a
'                particular delimter.
'
' Parameters:    strInput - Data to be parsed
'                strDelimiter - this is what we are looking for in the data
'                string
'                blnReturnFirstPart - [Optional] Default is TRUE.  If true, return
'                all data prior to the delimiter.  If False, return all data
'                after the delimiter.  Return nothing is delimiter cannot be
'                found.
'
' Returns:       part of the data datring
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-NOV-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function ParseData(ByVal strInput As String, _
                          ByVal strDelimiter As String, _
                          ByVal blnReturnFirstPart As Boolean, _
                          ByVal blnStartAtBeginning As Boolean) As String

    Dim intPosition As Integer
    Dim strTemp     As String
    
    strTemp = vbNullString
    
    If blnStartAtBeginning Then
        intPosition = InStr(1, strInput, strDelimiter)
    Else
        intPosition = InStrRev(strInput, strDelimiter, Len(strInput))
    End If
    
    ' parse the input string for the delimiter
    If blnReturnFirstPart Then
        ' capture the first portion up to the delimiter minus one
        If intPosition > 0 Then
            strTemp = Left$(strInput, intPosition - 1)
        End If
    Else
        ' capture the portion after the delimiter
        If intPosition > 0 Then
            strTemp = Mid$(strInput, intPosition + 1)
        End If
    End If
    
    ' Return the captured portion
    ParseData = strTemp
  
End Function

                                                                                                                                                                                                                                                                