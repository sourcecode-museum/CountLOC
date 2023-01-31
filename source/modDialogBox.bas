Attribute VB_Name = "modDialogBox"
' ***************************************************************************
' Module:        modDialogBox  (modDialogBox.bas)
'
' Description:   This module is used to open dialog boxes via API calls
'
' AddIn tools     Callers Add-in v2.24 by RD Edwards (RDE)
' for VB6:        Fantastic VB6 add-in to indentify if a routine calls another
'                 routine or is called by other routines within a project. A must
'                 have tool for any VB6 programmer.
'                 http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=73617&lngWId=1
'
'                 Keep track of MZTools-3.0 for VB6 and VBA, they are constantly
'                 making updates and improvements on this product even tho the
'                 version number does not change. Read the changelog at their web
'                 site for update information. This is a free and valuable VB6
'                 add-in. Another must have tool for VB programmers.
'                 http://www.mztools.com/index.aspx
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 26-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              Changed call to RemoveTrailingNulls() to TrimStr module due
'              to speed and accuracy.
' 01-Jul-2014  Kenneth Ives  kenaso@tx.rr.com
'              - Rewrote ShowBrowseForFolder() routine
'              - Added ShowFileOpen(), ShowFileSaveAs(), ShowColor()
'                routines
'              - Updated SetPBarForegroundColor() and SetPBarBackgroundColor()
'                routines
' 19-Jan-2015  Kenneth Ives  kenaso@tx.rr.com
'              Added ShowPrinter(), SetPrinterOrigin() routines
' 22-Feb-2015  Kenneth Ives  kenaso@tx.rr.com
'              - Renamed main routines for easier maintenance
'              - Updated CenterDialogBox() routine
'              - Updated routines which call CenterDialogBox()
'              - Added CenterOnScreen(), CenterOnForm() routines
'              - Combined setting progress bar color routines
' 05-Mar-2015  Kenneth Ives  kenaso@tx.rr.com
'              Updated ShowBrowseForFolder(), ShowFileOpen(), ShowFileSaveAs()
'              routines to save current folder for next inquiry unless user
'              sets their own starting folder in the calling application
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "modDialogBox"

  ' Constants for centering dialog boxes
  Private Const GWL_HINSTANCE          As Long = (-6)
  Private Const HCBT_ACTIVATE          As Long = 5
  Private Const SWP_NOSIZE             As Long = &H1
  Private Const SWP_NOZORDER           As Long = &H4
  Private Const SWP_NOACTIVATE         As Long = &H10
  Private Const WH_CBT                 As Long = 5

  ' Constants used for folder browsing
  Private Const MAX_AMT                As Long = 260
  Private Const WM_USER                As Long = &H400&
  Private Const BFFM_INITIALIZED       As Long = 1
  Private Const BFFM_SELCHANGED        As Long = 2
  Private Const BFFM_SETSTATUSTEXT     As Long = (WM_USER + 100)
  Private Const BFFM_SETSELECTION      As Long = (WM_USER + 102)
  Private Const BIF_RETURNONLYFSDIRS   As Long = &H1&       ' only file system directories
  Private Const BIF_DONTGOBELOWDOMAIN  As Long = &H2&
  Private Const BIF_STATUSTEXT         As Long = &H4&
  Private Const BIF_EDITBOX            As Long = &H10&      ' Display edit box at bottom
  Private Const BIF_NEWDIALOGSTYLE     As Long = &H40&      ' Use new style dialog box
  Private Const LMEM_FIXED             As Long = &H0&
  Private Const LMEM_ZEROINIT          As Long = &H40&
  
  Private Const BIF_USENEWUI           As Long = BIF_EDITBOX Or _
                                                 BIF_NEWDIALOGSTYLE
  
  Private Const MY_FINDFOLDER          As Long = BIF_DONTGOBELOWDOMAIN Or _
                                                 BIF_RETURNONLYFSDIRS Or _
                                                 BIF_STATUSTEXT Or _
                                                 BIF_USENEWUI
                                                 
  Private Const MY_POINTER             As Long = LMEM_FIXED Or _
                                                 LMEM_ZEROINIT
    
  ' Constants to open\Save files
  Private Const OFN_OVERWRITEPROMPT    As Long = &H2&       ' Causes the "Save As" dialog box to generate a message
                                                            ' box if the selected file already exists
  Private Const OFN_HIDEREADONLY       As Long = &H4&       ' Hides Read Only check box
  Private Const OFN_ALLOWMULTISELECT   As Long = &H200&
  Private Const OFN_EXTENSIONDIFFERENT As Long = &H400&
  Private Const OFN_CREATEPROMPT       As Long = &H2000&
  Private Const OFN_EXPLORER           As Long = &H80000
  Private Const OFN_LONGNAMES          As Long = &H200000
  Private Const OFN_NODEREFERENCELINKS As Long = &H100000
    
  Private Const MY_FILEOPEN_FLAGS      As Long = OFN_EXPLORER Or _
                                                 OFN_LONGNAMES Or _
                                                 OFN_CREATEPROMPT Or _
                                                 OFN_NODEREFERENCELINKS Or _
                                                 OFN_EXTENSIONDIFFERENT

  Private Const MY_FILESAVE_FLAGS      As Long = OFN_EXPLORER Or _
                                                 OFN_LONGNAMES Or _
                                                 OFN_HIDEREADONLY Or _
                                                 OFN_OVERWRITEPROMPT

  ' ChooseColor structure flag constants
  Private Const CC_RGBINIT             As Long = &H1&
  Private Const CC_FULLOPEN            As Long = &H2&
  Private Const CC_ANYCOLOR            As Long = &H100&

  Private Const MY_COLORSHOW_FLAGS     As Long = CC_ANYCOLOR Or _
                                                 CC_FULLOPEN Or _
                                                 CC_RGBINIT
  
  ' Constants used for coloring Microsoft progress bar
  Private Const CCM_FIRST              As Long = &H2000&
  Private Const CCM_SETBKCOLOR         As Long = (CCM_FIRST + 1)
  Private Const PBM_SETBKCOLOR         As Long = CCM_SETBKCOLOR
  Private Const PBM_SETBARCOLOR        As Long = (WM_USER + 9)

  ' Constants used for printer display
  Private Const CCHDEVICENAME          As Long = 32
  Private Const CCHFORMNAME            As Long = 32
  Private Const DLG_PRINT              As Long = 5    ' Show Print dialog box
  Private Const DLG_PRINTSETUP         As Long = 1    ' Show print setup dialog box
  Private Const DM_DUPLEX              As Long = &H1000&
  Private Const DM_ORIENTATION         As Long = &H1&
  Private Const GMEM_MOVEABLE          As Long = &H2&
  Private Const GMEM_ZEROINIT          As Long = &H40&
  Private Const PD_ALLPAGES            As Long = &H0&
  Private Const PD_NOSELECTION         As Long = &H4&
  Private Const PD_NOPAGENUMS          As Long = &H8&
  Private Const PD_PRINTSETUP          As Long = &H40&
  Private Const PD_HIDEPRINTTOFILE     As Long = &H100000
  Private Const PD_NONETWORKBUTTON     As Long = &H200000
  Private Const PHYSICALOFFSETX        As Long = 112
  Private Const PHYSICALOFFSETY        As Long = 113
  Private Const TRANSPARENT            As Long = 1

  ' Public constants because they may be passed parameters
  Public Const MY_PRINTSHOW_FLAGS     As Long = DLG_PRINT Or PD_ALLPAGES Or _
                                                PD_NOSELECTION Or PD_NOPAGENUMS Or _
                                                PD_NONETWORKBUTTON Or PD_HIDEPRINTTOFILE

  Public Const MY_PRINTSETUP_FLAGS    As Long = DLG_PRINTSETUP Or _
                                                PD_PRINTSETUP Or _
                                                PD_NONETWORKBUTTON


' ***************************************************************************
' Type Structures
' ***************************************************************************
  Private Type BROWSEINFO
      hOwner           As Long       ' Handle to owning window for dialog box.  Can be NULL.
      pidlRoot         As Long       ' Specifies location of the root folder from which to start
                                     ' browsing.  If NULL then Desktop folder is used.
      pszDisplayName   As String     ' Pointer to a buffer to receive the display name of the folder
                                     ' selected by the user.
      lpszTitle        As String     ' Dialog box title supplied by user.
      ulFlags          As Long       ' Flags that specify the options for the dialog box.
      lpfnHook         As Long       ' Pointer to an application-defined function that the dialog box
                                     ' calls when an event occurs. This member can be NULL.
      lParam           As Long       ' An application-defined value that the dialog box passes to the
                                     ' callback function, if one is specified in lpfnHook.
      iImage           As Long       ' An integer value that receives the index of the image associated
                                     ' with the selected folder, stored in the system image list.
  End Type

  ' Holds the parameters needed to open the dialog
  ' box. Also receives the returned filename.
  Private Type OPENFILENAME
      lStructSize       As Long      ' Length, in bytes, of the structure
      hwndOwner         As Long      ' Handle to window that owns the dialog box. Can be NULL if no owner.
      hInstance         As Long      ' If not set it is ignored
      lpstrFilter       As String    ' Buffer containing pairs of null-terminated filter strings. The last
                                     ' string in the buffer must be terminated by two NULL characters.
      lpstrCustomFilter As String    ' Static buffer that contains a pair of null-terminated filter strings
                                     ' for preserving the filter pattern chosen by the user. If this member
                                     ' is NULL, the dialog box does not preserve user-defined filter patterns.
                                     ' If this member is not NULL, the value of the nMaxCustFilter member
                                     ' must specify the size, in characters, of the lpstrCustomFilter buffer.
      nMaxCustFilter    As Long      ' Size, in characters, of the buffer identified by lpstrCustomFilter.
                                     ' This buffer should be at least 40 characters long. This member is
                                     ' ignored if lpstrCustomFilter is NULL or points to a NULL string.
      nFilterIndex      As Long      ' The index of the currently selected filter in the File Types control.
      lpstrFile         As String    ' The file name used to initialize the File Name edit control. The first
                                     ' character of this buffer must be NULL if initialization is not necessary.
      nMaxFile          As Long      ' The size, in characters, of the buffer pointed to by lpstrFile. The
                                     ' buffer must be large enough to store the path and file name string or
                                     ' strings, including the terminating NULL character.
      lpstrFileTitle    As String    ' File name and extension (w\o path) of selected file. This can be NULL.
      nMaxFileTitle     As Long      ' The size, in characters, of the buffer pointed to by lpstrFileTitle.
                                     ' This is ignored if lpstrFileTitle is NULL.
      lpstrInitialDir   As String    ' The initial directory
      lpstrTitle        As String    ' A string to be placed in the title bar of the dialog box.
      ulFlags           As Long      ' A set of bit flags you can use to initialize the dialog box.
      nFileOffset       As Integer   ' The zero-based offset, in characters, from the beginning of the path to
                                     ' the file name in the string pointed to by lpstrFile.
      nFileExtension    As Integer   ' The zero-based offset, in characters, from the beginning of the path to
                                     ' the file name extension in the string pointed to by lpstrFile.
      lpstrDefExt       As String    ' The default extension. GetOpenFileName and GetSaveFileName append this
                                     ' extension to the file name if the user fails to type an extension.
      lCustData         As Long      ' Application-defined data that the system passes to the hook procedure
                                     ' identified by the lpfnHook member.
      lpfnHook          As Long      ' A pointer to a hook procedure. This member is ignored unless the Flags
                                     ' member includes the OFN_ENABLEHOOK flag. If OFN_ENABLEHOOK flag is used
                                     ' the Windows 7 and newer style dialog boxes are not displayed, only XP
                                     ' style boxes will be displayed.
      lpTemplateName    As String    ' The name of the dialog template resource in the module identified by
                                     ' the hInstance member.
      ' New for Windows 2000 and later
      pvReserved        As Long      ' Reserved (not used)
      dwReserved        As Long      ' Reserved (not used)
      FlagsEx           As Long      ' Reserved
  End Type
       
  ' For choosing a color
  Private Type COLORSTRUC
      lStructSize       As Long      ' Length, in bytes, of the structure
      hwndOwner         As Long      ' Handle to window that owns the dialog box. Can be NULL if no owner.
      hInstance         As Long      ' If not set it is ignored
      rgbResult         As Long      ' Selected color, if any
      lpCustColors      As String    ' A pointer to an array of 16 values that contain red, green, blue (RGB)
                                     ' values for the custom color boxes in the dialog box.
      ulFlags           As Long      ' A set of bit flags you can use to initialize the dialog box.
      lCustData         As Long      ' Application-defined data that the system passes to the hook procedure
                                     ' identified by the lpfnHook member.
      lpfnHook          As Long      ' A pointer to a hook procedure. This member is ignored unless the Flags
                                     ' member includes the CC_ENABLEHOOK flag. If CC_ENABLEHOOK flag is used
                                     ' the Windows 7 and newer style dialog boxes are not displayed, only XP
                                     ' style boxes will be displayed.
      lpTemplateName    As String    ' The name of the dialog template resource in the module identified by
                                     ' the hInstance member. This member is ignored unless the CC_ENABLETEMPLATE
                                     ' flag is set in the Flags member.
  End Type

  ' For centering dialog boxes
  Private Type RECT
      Left              As Long      ' Left side of dialog box
      Top               As Long      ' Top side of dialog box
      Right             As Long      ' Right side of dialog box
      Bottom            As Long      ' Bottom side of dialog box
  End Type

  ' Used for print dialog
  Private Type PRINTDLG_TYPE
      lStructSize         As Long      ' Structure size, in bytes
      hwndOwner           As Long      ' Handle to the window that owns the dialog box
      hDevMode            As Long      ' Handle to a movable global memory object that contains a DEVMODE structure
      hDevNames           As Long      ' Handle to a movable global memory object that contains a DEVNAMES structure
      hdc                 As Long      ' handle to a device context or an information context, depending on whether
                                       ' the Flags member specifies the PD_RETURNDC or PC_RETURNIC flag. If neither
                                       ' flag is specified, the value of this member is undefined. If both flags are
                                       ' specified, PD_RETURNDC has priority.
      flags               As Long      ' Initializes the Print dialog box. When the dialog box returns, it sets these
                                       ' flags to indicate the user's input.
      nFromPage           As Integer   ' Initial value for the starting page edit controL
      nToPage             As Integer   ' Initial value for the ending page edit control. When PrintDlg returns, nToPage
                                       ' is the ending page specified by the user.
      nMinPage            As Integer   ' Minimum value for the page range specified in the From and To page edit controls
      nMaxPage            As Integer   ' maximum value for the page range specified in the From and To page edit controls
      nCopies             As Integer   ' Actual number of copies to print
      hInstance           As Long      ' Handle to the application or module instance that contains the dialog box template
      lCustData           As Long      ' Application-defined data that the system passes to the hook procedure identified
                                       ' by the lpfnPrintHook or lpfnSetupHook member.
      lpfnPrintHook       As Long      ' A pointer to a PrintHookProc hook procedure that can process messages intended
                                       ' for the Print dialog box. This member is ignored unless the PD_ENABLEPRINTHOOK
                                       ' flag is set in the Flags member.
      lpfnSetupHook       As Long      ' A pointer to a SetupHookProc hook procedure that can process messages intended
                                       ' for the Print Setup dialog box. This member is ignored unless the
                                       ' PD_ENABLESETUPHOOK flag is set in the Flags member.
      lpPrintTemplateName As String    ' This template replaces the default Print dialog box template. This member is
                                       ' ignored unless the PD_ENABLEPRINTTEMPLATE flag is set in the Flags member.
      lpSetupTemplateName As String    ' This template replaces the default Print Setup dialog box template. This member
                                       ' is ignored unless the PD_ENABLESETUPTEMPLATE flag is set in the Flags member.
      hPrintTemplate      As Long      ' If the PD_ENABLEPRINTTEMPLATEHANDLE flag is set in the Flags member, hPrintTemplate
                                       ' is a handle to a memory object containing a dialog box template. This template
                                       ' replaces the default Print dialog box template.
      hSetupTemplate      As Long      ' If the PD_ENABLESETUPTEMPLATEHANDLE flag is set in the Flags member, hSetupTemplate
                                       ' is a handle to a memory object containing a dialog box template. This template
                                       ' replaces the default Print Setup dialog box template.
  End Type

  ' Used for print dialog
  Private Type DEVNAMES_TYPE
      wDriverOffset As Integer         ' On input, this string is used to determine the printer to display initially in
                                       ' the dialog box.
      wDeviceOffset As Integer         ' The offset, in characters, from the beginning of this structure to the null
                                       ' terminated string that contains the name of the device.
      wOutputOffset As Integer         ' The offset, in characters, from the beginning of this structure to the null
                                       ' terminated string that contains the device name for the physical output medium.
      wDefault      As Integer         ' Indicates whether the strings contained in the DEVNAMES structure identify the
                                       ' default printer
      extra         As String * 200
  End Type

  ' Used for print dialog
  Private Type DEVMODE_TYPE
      dmDeviceName       As String * CCHDEVICENAME   ' A zero-terminated character array that specifies the "friendly"
                                       ' name of the printer
      dmSpecVersion      As Integer    ' The version number of the initialization data specification on which the
                                       ' structure is based.
      dmDriverVersion    As Integer    ' The driver version number assigned by the driver developer.
      dmSize             As Integer    ' Specifies the size, in bytes, of the DEVMODE structure, not including any
                                       ' private driver-specific data that might follow the structure's public members.
      dmDriverExtra      As Integer    ' Contains the number of bytes of private driver-data that follow this structure.
                                       ' If a device driver does not use device-specific information, set this member to
                                       ' zero.
      dmFields           As Long       ' Specifies whether certain members of the DEVMODE structure have been initialized.
                                       ' If a member is initialized, its corresponding bit is set, otherwise the bit is clear.
      dmOrientation      As Integer    ' Orientation of the paper.  Can be either DMORIENT_PORTRAIT (1) or DMORIENT_LANDSCAPE (2).
      dmPaperSize        As Integer    ' Selects the size of the paper to print on. This member can be set to zero if the
                                       ' length and width of the paper are both set by the dmPaperLength and dmPaperWidth
                                       ' members.
      dmPaperLength      As Integer    ' Overrides the length of the paper specified by the dmPaperSize member
      dmPaperWidth       As Integer    ' Overrides the width of the paper specified by the dmPaperSize member
      dmScale            As Integer    ' Specifies the factor by which the printed output is to be scaled
      dmCopies           As Integer    ' Number of copies printed if the device supports multiple-page copies
      dmDefaultSource    As Integer    ' Specifies the paper source
      dmPrintQuality     As Integer    ' Specifies the printer resolution
      dmColor            As Integer    ' Switches between color and monochrome on color printers.
      dmDuplex           As Integer    ' Selects duplex or double-sided printing for printers capable of duplex printing
      dmYResolution      As Integer    ' Specifies the y-resolution, in dots per inch, of the printer
      dmTTOption         As Integer    ' Specifies how TrueType fonts should be printed
      dmCollate          As Integer    ' Specifies whether collation should be used when printing multiple copies
      dmFormName         As String * CCHFORMNAME   ' A zero-terminated character array that specifies the name of the
                                       ' form to use; for example, "Letter" or "Legal"
      dmUnusedPadding    As Integer    '
      dmBitsPerPel       As Integer    ' Specifies the color resolution, in bits per pixel, of the display device
      dmPelsWidth        As Long       ' Specifies the width, in pixels, of the visible device surface
      dmPelsHeight       As Long       ' Specifies the height, in pixels, of the visible device surface
      dmDisplayFlags     As Long       ' Specifies the device's display mode
      dmDisplayFrequency As Long       ' Display device's vertical refresh rate
      dmICMMethod        As Long       '
      dmICMIntent        As Long       '
      dmMediaType        As Long       ' Specifies the type of media being printed on
      dmDitherType       As Long       ' Specifies how dithering is to be done
      dmReserved1        As Long       ' Not used; must be zero
      dmReserved2        As Long       ' Not used; must be zero
      dmPanningWidth     As Long       ' Not used; must be zero
      dmPanningHeight    As Long       ' Not used; must be zero
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

  ' Converts an item identifier list to a file system path.  This function
  ' returns non-zero if successful.
  Private Declare Function SHGetPathFromIDList Lib "shell32" _
          Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
          ByVal pszPath As String) As Long

  ' Displays a dialog box that enables the user to select a Shell folder.
  Private Declare Function SHBrowseForFolder Lib "shell32" _
          Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

  ' Creates an Open dialog box that lets the user specify the drive,
  ' directory, and the name of a file or set of files to be opened.
  ' This function returns 1 if successful else 0 if an error occured
  ' or Cancel was selected.
  Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
          Alias "GetOpenFileNameA" (pOPENFILENAME As OPENFILENAME) As Long
  
  ' Creates a Save dialog box that lets the user specify the drive,
  ' directory, and name of a file to save. It does not actually save
  ' the file.  This function returns 1 if successful else 0 if an
  ' error occured or Cancel was selected.
  Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
          Alias "GetSaveFileNameA" (pOPENFILENAME As OPENFILENAME) As Long

  ' ChooseColor function creates a Color dialog box that enables
  ' the user to select a color.  This function returns 1 if successful
  ' else 0 if an error occured or Cancel was selected.
  Private Declare Function ChooseColor Lib "comdlg32.dll" _
          Alias "ChooseColorA" (pChoosecolor As COLORSTRUC) As Long

  ' GetWindowRect function retrieves the dimensions of the bounding
  ' rectangle of the specified window.  The dimensions are given in
  ' screen coordinates that are relative to the upper-left corner of
  ' the screen.
  Private Declare Function GetWindowRect Lib "user32" _
          (ByVal hWnd As Long, lpRect As RECT) As Long

  ' Truncates a path to fit within a certain number of characters
  ' by replacing path components with ellipses.
  Private Declare Function PathCompactPathEx Lib "shlwapi.dll" _
          Alias "PathCompactPathExA" (ByVal pszOut As String, _
          ByVal pszSrc As String, ByVal cchMax As Long, _
          ByVal dwFlags As Long) As Long

  ' ZeroMemory fills a location in memory with zeros. The function sets
  ' each byte starting at the given memory location to zero. The memory
  ' location is identified by a pointer to the memory address.
  Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" _
          (ByRef Destination As Any, ByVal Length As Long)

  ' The SendMessage function sends the specified message to a window or
  ' windows. The function calls the window procedure for the specified
  ' window and does not return until the window procedure has processed
  ' the message.  If function fails, return value is zero.
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
          (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
          lParam As Any) As Long

  ' CoTaskMemFree function frees a block of task memory previously
  ' allocated through a call. A pointer to the memory block to be
  ' freed. If this parameter is NULL, the function has no effect.
  Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

  ' LocalAlloc function allocates the specified number of bytes
  ' from the heap.
  Private Declare Function LocalAlloc Lib "kernel32" _
          (ByVal uFlags As Long, ByVal uBytes As Long) As Long

  ' GetWindowLong function retrieves information about the specified window.
  ' The function also retrieves the 32-bit (DWORD) value at the specified
  ' offset into the extra window memory. If the function succeeds, the return
  ' value is the requested value.
  Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
          (ByVal hWnd As Long, ByVal nIndex As Long) As Long

  ' GetCurrentThreadId function retrieves the thread identifier of the calling
  ' thread.  The return value is the thread identifier of the calling thread.
  Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

  ' SetWindowsHookEx function installs an application-defined hook procedure
  ' into a hook chain.  If the function succeeds, the return value is the
  ' handle to the hook procedure.
  Private Declare Function SetWindowsHookEx Lib "user32" _
          Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, _
          ByVal hmod As Long, ByVal dwThreadId As Long) As Long

  ' UnhookWindowsHookEx function removes a hook procedure installed in a hook
  ' chain by the SetWindowsHookEx function.  If the function succeeds, the
  ' return value is nonzero.
  Private Declare Function UnhookWindowsHookEx Lib "user32" _
          (ByVal mlngHook As Long) As Long

  ' SetWindowPos function changes the size, position, and Z order of a child,
  ' pop-up, or top-level window.  If the function succeeds, the return value
  ' is nonzero.
  Private Declare Function SetWindowPos Lib "user32" _
          (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
          ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
          ByVal cy As Long, ByVal wFlags As Long) As Long


  '====================================================================
  ' Used for print dialog
  '
  ' GetDeviceCaps function retrieves device-specific information
  ' about a specified device.
  Private Declare Function GetDeviceCaps Lib "gdi32" _
          (ByVal hdc As Long, ByVal nIndex As Long) As Long

  ' PrintDialog function displays a Print dialog box or a Print Setup
  ' dialog box. The Print dialog box enables the user to specify the
  ' properties of a particular print job.
  Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" _
          (pPrintdlg As PRINTDLG_TYPE) As Long

  ' GlobalLock function locks a global memory object and returns a
  ' pointer to the first byte of the object's memory block.
  Private Declare Function GlobalLock Lib "kernel32" _
          (ByVal hMem As Long) As Long

  ' GlobalUnlock function decrements the lock count associated with
  ' a memory object that was allocated with the GMEM_MOVEABLE flag.
  Private Declare Function GlobalUnlock Lib "kernel32" _
          (ByVal hMem As Long) As Long

  ' GlobalAlloc function allocates the specified number of bytes from
  ' the heap. If the function succeeds, the return value is a handle
  ' to the newly allocated memory object.
  Private Declare Function GlobalAlloc Lib "kernel32" _
          (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
  
  ' GlobalFree function frees the specified global memory object and
  ' invalidates its handle. If the function succeeds, the return value
  ' is NULL.
  Private Declare Function GlobalFree Lib "kernel32" _
         (ByVal hMem As Long) As Long

  ' SetBkMode function sets the background mix mode of the specified
  ' device context. The background mix mode is used with text, hatched
  ' brushes, and pen styles that are not solid lines.
  Private Declare Function SetBkMode Lib "gdi32" _
          (ByVal hdc As Long, ByVal nBkMode As Long) As Long
  '====================================================================

' ***************************************************************************
' Module Variables
'                    +-------------- Module level designator
'                    |  +----------- Data type (String)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m str DialogTitle
' Variable name:     mstrDialogTitle
' ***************************************************************************
  Private mlngHook        As Long     ' Used for centering dialog boxes
  Private mlngFormHwnd    As Long     ' Calling form owner handle
  Private mlngLastColor   As Long     ' Previously selected color
  Private mstrDialogTitle As String   ' Title of dialog box
  Private mstrStartFolder As String   ' User defined starting path


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       SetStartingFolder
'
' Description:   User defined starting folder (path) when accessing folder
'                or file dialog boxes.
'
' Parameters:    strFolder - Input - Fully qualified path to starting folder
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 26-Jun-2014  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub SetStartingFolder(ByVal strFolder As String)
    
    mstrStartFolder = TrimStr(strFolder)    ' Remove unwanted characters
    
End Sub

' ***************************************************************************
' Routine:       ShowBrowseForFolder
'
' Description:   This function will open the "Browse for Folder" dialog box.
'
' Parameters:    frmName  - Name of calling form (ex:  frmMain)
'                blnCenterOnScreen - Optional - Flag designating where
'                      dialog box should be centered.
'                      TRUE  = Center on screen (Default)
'                      FALSE = Center over top of calling form
'
' Returns:       Name of folder selected.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Jul-2014  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function ShowBrowseForFolder(ByVal frmName As Form, _
                           Optional ByVal blnCenterOnScreen As Boolean = True) As String

    Dim lngFolderID       As Long
    Dim strPath           As String
    Dim strSelectedFolder As String
    Dim typBrowseInfo     As BROWSEINFO

    ZeroMemory typBrowseInfo, Len(typBrowseInfo)    ' Initialize type structure
    mlngFormHwnd = frmName.hWnd                     ' Save handle of calling form
    
    ' Determine starting folder location
    mstrStartFolder = IIf(Len(mstrStartFolder) = 0, "C:\", mstrStartFolder)
    
    ' Current highlighted path
    ' displayed at top of dialog box
    strPath = ShrinkToFit(mstrStartFolder, 100)
    
    With typBrowseInfo
        .hOwner = mlngFormHwnd      ' Calling form handle
        .pidlRoot = 0               ' Desktop folder will be root folder
        .lpszTitle = strPath        ' Data string denoting starting folder location
        .ulFlags = MY_FINDFOLDER    ' Flags designating how to display folders
        
        ' Create pointer to starting folder
        lngFolderID = LocalAlloc(MY_POINTER, Len(mstrStartFolder) + 1)
        CopyMemory ByVal lngFolderID, ByVal mstrStartFolder, Len(mstrStartFolder) + 1
        .lParam = lngFolderID
    End With
    
    ' Hook to identify starting folder
    typBrowseInfo.lpfnHook = GetFunctionAddress(AddressOf FindStartingFolder)
    
    CenterDialogBox blnCenterOnScreen   ' Center dialog box
    DoEvents
    
    ' Make API call to display Browse for Folder dialog box
    lngFolderID = SHBrowseForFolder(typBrowseInfo)       ' Show folder dialog box
    strSelectedFolder = String$(MAX_AMT, 0)              ' Preload folder name with nulls
    
    ' See if any data was captured
    If SHGetPathFromIDList(ByVal lngFolderID, ByVal strSelectedFolder) <> 0 Then
        strSelectedFolder = TrimStr(strSelectedFolder)   ' Remove unwanted leading\trailing characters
        mstrStartFolder = strSelectedFolder              ' New starting folder
        CoTaskMemFree lngFolderID                        ' Free return code from memory
    Else
        strSelectedFolder = mstrStartFolder        ' User selected CANCEL button,
    End If                                         ' revert to starting folder
    
    ZeroMemory typBrowseInfo, Len(typBrowseInfo)   ' Clear type structure
    mstrStartFolder = vbNullString                 ' Empty variable
    
    ShowBrowseForFolder = strSelectedFolder        ' Return path\name of folder
    
End Function

' ***************************************************************************
' Routine:       ShowFileOpen
'
' Description:   This function will display the "File Open" dialog box.
'
' Parameters:    frmName  - Name of calling form (ex:  frmMain)
'                strTitle - [Optional] Title to be displayed on the dialog
'                      dialog box.  Uses default title if none is provided.
'                strFileExts - [Optional] File extension filters.
'                      Default = all files
'                blnOneFileOnly - [Optional] - Flag designating if one or
'                      more than one file(s) may be selected.
'                      TRUE - One file only (Default)
'                      FALSE - More than one file
'                blnCenterOnScreen - Optional - Flag designating where
'                      dialog box should be centered.
'                      TRUE  = Center on screen (Default)
'                      FALSE = Center over top of calling form
'
' Returns:       Path and name(s) of file selected.
'
'                ex:  One file only    "C:\Temp\Test1.txt"
'
'                     Multiple files   "C:\Temp Test1.txt Test2.txt Test3.txt"
'       (Delimited by NULL characters) |-------+---------+---------+---------|
'                                       Folder   File #1   File #2    etc.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Jul-2014  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function ShowFileOpen(ByVal frmName As Form, _
                    Optional ByVal strTitle As String = vbNullString, _
                    Optional ByVal strFilter As String = vbNullString, _
                    Optional ByVal blnOneFileOnly As Boolean = True, _
                    Optional ByVal blnCenterOnScreen As Boolean = True) As String

    Dim strFile       As String
    Dim lngFlags      As Long
    Dim lngRetCode    As Long
    Dim lngBufferSize As Long
    Dim typOpenFile   As OPENFILENAME

    strFile = vbNullString                      ' Preload return variable
    ZeroMemory typOpenFile, Len(typOpenFile)    ' Initialize type structure
    mlngFormHwnd = frmName.hWnd                 ' Save handle of calling form
    
    ' Determine starting folder location
    mstrStartFolder = IIf(Len(mstrStartFolder) = 0, "C:\", mstrStartFolder)
    
    ' Caption of dialog box
    If Len(strTitle) > 0 Then
        mstrDialogTitle = strTitle              ' User supplied caption
    Else
        mstrDialogTitle = "Browse for a file"   ' Default caption
    End If
    
    ' Default view is to look for all file types
    If Len(strFilter) = 0 Then
        strFilter = "All files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    End If
         
    ' See if application wants to be able
    ' to select one or more files at a time
    If blnOneFileOnly Then
        lngFlags = MY_FILEOPEN_FLAGS          ' Single file selection
        lngBufferSize = ((MAX_AMT * 2) + 1)   ' Allow enough space to receive path\file name
    Else
        lngFlags = MY_FILEOPEN_FLAGS Or OFN_ALLOWMULTISELECT
        
        ' Calc buffer size to allow many files to be selected
        lngBufferSize = ((MAX_AMT * (25 + 1)) + 1)
    End If
    
    ' Preload type structure
    With typOpenFile
        .lStructSize = Len(typOpenFile)                   ' Size of type structure
        .hwndOwner = mlngFormHwnd                         ' Calling form handle
        .hInstance = App.hInstance                        ' Application owns dialog box
        .lpstrFilter = strFilter                          ' Types of file extensions to display
        .nFilterIndex = 1                                 ' Use first filter (file extension) option
        .lpstrFile = String$(lngBufferSize, 0)            ' Receives path and filename of selected file
        .nMaxFile = Len(.lpstrFile)                       ' Size of the path and filename buffer
        .lpstrFileTitle = 0&                              ' Receives filename of selected file
        .nMaxFileTitle = 0&                               ' Size of the filename buffer
        .lpstrInitialDir = mstrStartFolder & vbNullChar   ' Starting folder
        .ulFlags = lngFlags                               ' Designating how many files may be selected
    End With
    
    CenterDialogBox blnCenterOnScreen   ' Center dialog box
    DoEvents
    
    ' Make API call to display File Open dialog box
    lngRetCode = GetOpenFileName(typOpenFile)      ' Make API call to display open file dialog box
    
    ' See if user pressed CANCEL button
    If lngRetCode > 0 Then
        mstrStartFolder = GetFullPath(TrimStr(typOpenFile.lpstrFile))   ' Capture full path
        strFile = TrimStr(typOpenFile.lpstrFile)                        ' Remove unwanted characters
    End If
    
    ZeroMemory typOpenFile, Len(typOpenFile)   ' Clear type structure
    mstrStartFolder = vbNullString             ' Empty variable

    ShowFileOpen = strFile   ' Return path\File name

End Function

' ***************************************************************************
' Routine:       ShowFileSaveAs
'
' Description:   This function will display the "File Save As" dialog box.
'
' Parameters:    frmName  - Name of calling form (ex:  frmMain)
'                strFilename - Name of file to be saved
'                strTitle - [Optional] Title to be displayed on the dialog
'                      dialog box.  Uses default title if none is provided.
'                strFilter - [Optional] File extension filters.
'                      Default = all files.
'                blnCenterOnScreen - Optional - Flag designating where
'                      dialog box should be centered.
'                      TRUE  = Center on screen (Default)
'                      FALSE = Center over top of calling form
'
' Returns:       TRUE if successful
'                FALSE if not
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Jul-2014  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function ShowFileSaveAs(ByVal frmName As Form, _
                               ByRef strFileName As String, _
                      Optional ByVal strTitle As String = vbNullString, _
                      Optional ByVal strFilter As String = vbNullString, _
                      Optional ByVal blnCenterOnScreen As Boolean = True) As Boolean
                  
    Dim lngRetCode  As Long
    Dim typOpenFile As OPENFILENAME

    ZeroMemory typOpenFile, Len(typOpenFile)   ' Initialize type structure
    mlngFormHwnd = frmName.hWnd                ' Save handle of calling form
    
    ' Determine starting folder location
    mstrStartFolder = IIf(Len(mstrStartFolder) = 0, "C:\", mstrStartFolder)
    
    ' Caption of dialog box
    If Len(strTitle) > 0 Then
        mstrDialogTitle = strTitle         ' User supplied caption
    Else
        mstrDialogTitle = "File Save As"   ' Default caption
    End If
    
    ' Default view is to look for all file types
    If Len(strFilter) = 0 Then
        strFilter = "All files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    End If
         
    ' Preload type structure
    With typOpenFile
        .lStructSize = Len(typOpenFile)                   ' Size of type structure
        .hwndOwner = mlngFormHwnd                         ' Calling form handle
        .hInstance = App.hInstance                        ' Application owns dialog box
        .lpstrFilter = strFilter                          ' Types of file extensions to display
        .nFilterIndex = 1                                 ' Use first filter (file extension) option
        .lpstrFile = strFileName & vbNullChar             ' Receives path and filename of selected file
        .lpstrFileTitle = String$(1024, 0)                ' Receives filename of selected file
        .nMaxFileTitle = Len(.lpstrFileTitle)             ' Size of the filename buffer
        .lpstrFile = String$(1024, 0)                     ' Name used to initialize File Name edit control
        .nMaxFile = Len(.lpstrFile)                       ' Size of the path and filename buffer
        .lpstrInitialDir = mstrStartFolder & vbNullChar   ' Starting folder
        .ulFlags = MY_FILESAVE_FLAGS                      ' Designating how "File Save As" dialog is displayed
    End With
    
    CenterDialogBox blnCenterOnScreen   ' Center dialog box
    DoEvents
    
    ' Make API call to display File Save AS dialog box
    lngRetCode = GetSaveFileName(typOpenFile)   ' Make API call to display file "Save As" dialog box
    
    If lngRetCode > 0 Then
        ' Capture full path
        mstrStartFolder = GetFullPath(TrimStr(typOpenFile.lpstrFile))
        
        ' Capture fully qualified path\file name
        strFileName = QualifyPath(mstrStartFolder) & _
                      TrimStr(typOpenFile.lpstrFileTitle)
        ShowFileSaveAs = True   ' Successful finish
    Else
        ShowFileSaveAs = False  ' CANCEL button pressed or file not saved
    End If
    
    ZeroMemory typOpenFile, Len(typOpenFile)   ' Clear type structure
    mstrStartFolder = vbNullString             ' Empty variable
    
End Function

' ***************************************************************************
' Routine:       ShowColor
'
' Description:   This function will display the "Choose Color" dialog box.
'
'                To get a Custom color to 'Take', you must open the dialog
'                box and then Click on the vertical color bar at the far
'                right at least once.  After that, you can click anywhere
'                in the multi-color portion to choose a color.
'
'                Example:  Dim lngColor     As Long
'                          Dim alngColors() As Long
'                          ReDim alngColors(16)
'                          ' Return one selected color and
'                          ' array of all 16 custom colors
'                          ShowColor frmMain, lngColor, , , , alngColors()
'
' Reference:     Randy Birch  29-Mar-2002
'                http://vbnet.mvps.org/code/hooks/choosecolorcustomize.html
'
' Parameters:    frmName    - Name of calling form (Required)
'                lngColor   - Color selected or start with this color
'                intRed     - Optional - Red color
'                intGreen   - Optional - Green color
'                intBlue    - Optional - Blue color
'                avntColors - Optional - Long Integer array of all
'                             sixteen custom colors
'                strTitle   - Optional - Title for dialog box
'                blnCenterOnScreen - Optional - Flag designating where
'                      dialog box should be centered.
'                      TRUE  = Center on screen (Default)
'                      FALSE = Center over top of calling form
'
' Returns:       Long integer designating a selected color,
'                optionally returned are RGB values for the selected color,
'                and a long integer array of custom colors.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Jun-2014  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub ShowColor(ByVal frmName As Form, _
                     ByRef lngColor As Long, _
            Optional ByRef intRed As Integer = -1, _
            Optional ByRef intGreen As Integer = -1, _
            Optional ByRef intBlue As Integer = -1, _
            Optional ByRef avntColors As Variant = Empty, _
            Optional ByVal strTitle As String = vbNullString, _
            Optional ByVal blnCenterOnScreen As Boolean = True)

    Dim lngIdx         As Long
    Dim lngSize        As Long
    Dim lngIndex       As Long
    Dim lngRetCode     As Long
    Dim abytRGB()      As Byte
    Dim typChooseColor As COLORSTRUC

    Const ROUTINE_NAME As String = "ShowColor"
    
    ZeroMemory typChooseColor, Len(typChooseColor)   ' Initialize type structure
    mlngFormHwnd = frmName.hWnd                      ' Save handle of calling form
    lngIdx = 0                                       ' Initialize variables
    lngSize = 0
    
    Erase abytRGB()                  ' Empty array
    ReDim abytRGB(0 To 63) As Byte   ' Size custom color array
    
    If lngColor > 0 Then
        mlngLastColor = lngColor     ' Color passed by user
    Else
        ' If no previous color selected then use
        ' color in upper left corner of dialog box
        If mlngLastColor = 0 Then
            mlngLastColor = RGB(255, 128, 128)   ' Long integer equivalent (8421631)
        End If
    End If
    
    ' Caption of dialog box
    If Len(strTitle) > 0 Then
        mstrDialogTitle = strTitle           ' User supplied caption
    Else
        mstrDialogTitle = "Select a color"   ' Default caption
    End If
    
    ' See if user has passed any custom colors
    If IsEmpty(avntColors) Then
                
        ' Preload custom color array
        ' with varying shades of gray
        For lngIndex = 15 To 240 Step 15
        
            abytRGB(lngIdx) = CByte(lngIndex)       ' Save red
            abytRGB(lngIdx + 1) = CByte(lngIndex)   ' Save green
            abytRGB(lngIdx + 2) = CByte(lngIndex)   ' Save blue
            abytRGB(lngIdx + 3) = CByte(0)          ' Null value (Delimiter)
            
            lngIdx = lngIdx + 4                     ' Increment output index
        
        Next lngIndex
        
    Else
        ' At this point, incoming variant
        ' must be an a array of data
        If CBool(IsArrayInitialized(avntColors)) Then
        
            ' Verify this is a long integer array
            If VarType(avntColors(0)) <> vbLong Then
            
                InfoMsg "Incoming variant data is not a Long Integer array." & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                
                GoTo ShowColor_CleanUp
                
            End If   ' VarType(avntColors(0))
            
            ' Check array size
            If UBound(avntColors) >= 15 Then
            
                lngSize = UBound(avntColors)   ' Save array size
                
                ' If no valid values greater than zero then
                ' preload with varying shades of gray
                If avntColors(0) = 0 And _
                   avntColors(15) = 0 Then
                
                    ' Preload custom color array
                    ' with varying shades of gray
                    For lngIndex = 15 To 240 Step 15
                    
                        abytRGB(lngIdx) = CByte(lngIndex)       ' Save red
                        abytRGB(lngIdx + 1) = CByte(lngIndex)   ' Save green
                        abytRGB(lngIdx + 2) = CByte(lngIndex)   ' Save blue
                        abytRGB(lngIdx + 3) = CByte(0)          ' Null value (Delimiter)
                        
                        lngIdx = lngIdx + 4                     ' Increment output index
                    Next lngIndex
                
                Else
                    ' Populate custom colors with user defined colors
                    For lngIndex = 0 To 15
                    
                        ' Convert long integer color code to Red,
                        ' Green and Blue color codes (0 to 255)
                        LongToRGB avntColors(lngIndex), intRed, intGreen, intBlue
                        
                        abytRGB(lngIdx) = CByte(intRed)         ' Save red
                        abytRGB(lngIdx + 1) = CByte(intGreen)   ' Save green
                        abytRGB(lngIdx + 2) = CByte(intBlue)    ' Save blue
                        abytRGB(lngIdx + 3) = CByte(0)          ' Null value (Delimiter)
                        
                        lngIdx = lngIdx + 4                     ' Increment output index
                        
                    Next lngIndex
                
                End If  ' avntColors(0)
                
            Else
                ' Array size cannot hold sixteen colors
                InfoMsg "Incoming custom colors array is wrong size." & vbNewLine & _
                        "Please make appropriate corrections and try again." & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                
                GoTo ShowColor_CleanUp
                
            End If   ' UBound(avntColors)
            
        Else
            InfoMsg "Incoming variant data has not been" & vbNewLine & _
                    "properly initialized as an array." & _
                    vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
            
            GoTo ShowColor_CleanUp
            
        End If   ' IsArrayInitialized()
        
    End If   ' IsEmpty(avntColors)
    
    With typChooseColor
        .lStructSize = Len(typChooseColor)              ' Size of type structure
        .hwndOwner = mlngFormHwnd                       ' Calling form handle
        .hInstance = App.hInstance                      ' Application owns dialog box
        .rgbResult = mlngLastColor                      ' Starting color
        .ulFlags = MY_COLORSHOW_FLAGS                   ' Flags for displaying dialog box
        .lpCustColors = StrConv(abytRGB(), vbUnicode)   ' Insert custom colors into dialog box
    End With
    
    CenterDialogBox blnCenterOnScreen   ' Center dialog box
    DoEvents
    
    ' Make API call to display Color Selection dialog box
    lngRetCode = ChooseColor(typChooseColor)
        
    ' Evaluate return code (zero = cancel selected)
    If lngRetCode > 0 Then
        
        lngColor = typChooseColor.rgbResult   ' Color selected
        mlngLastColor = lngColor              ' Save selected color
        
        ' Convert long integer color code to Red,
        ' Green and Blue color codes (0 to 255)
        LongToRGB lngColor, intRed, intGreen, intBlue
                
        ' Load return array with custom colors
        abytRGB() = StrConv(typChooseColor.lpCustColors, vbFromUnicode)
        
        If lngSize >= 15 Then
            
            ReDim avntColors(lngSize) As Long   ' Empty and size return array
            lngIdx = 0                          ' Initialize array pointer
            
            ' Transfer custom colors from byte
            ' array to long integer array
            For lngIndex = 0 To UBound(abytRGB) Step 4
            
                ' Data in byte array represents RGB
                ' colors delimited by a zero (null)
                '   ex:  abytRGB(lngIndex)     = 255
                '        abytRGB(lngIndex + 1) = 128
                '        abytRGB(lngIndex + 2) = 128
                '        abytRGB(lngIndex + 3) = 0   Null value
                avntColors(lngIdx) = RGB(abytRGB(lngIndex), _
                                         abytRGB(lngIndex + 1), _
                                         abytRGB(lngIndex + 2))
                                         
                lngIdx = lngIdx + 1   ' Increment return array index
            Next lngIndex
            
        End If   ' lngSize
        
    End If
    
ShowColor_CleanUp:
    ZeroMemory typChooseColor, Len(typChooseColor)   ' Clear type structure
    Erase abytRGB()                                  ' Empty array when not needed
    
End Sub

' ***************************************************************************
' Routine:       LongToRGB
'
' Description:   Calculate Red, Green, Blue (RGB) values based on long
'                integer color code.
'
' Parameters:    lngColor - Input - Long integer color code
'                intRed   - Output - Red color code (0-255)
'                intGreen - Output - Green color code (0-255)
'                intBlue  - Output - Blue color code (0-255)
'
' Returns:       Numeric values for Red, Green, and Blue hues that make
'                up the color (lngColor) that was passed to this routine
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Jul-2014  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub LongToRGB(ByVal lngColor As Long, _
                     ByRef intRed As Integer, _
                     ByRef intGreen As Integer, _
                     ByRef intBlue As Integer)
    
    intRed = -1     ' Initialize color base
    intGreen = -1
    intBlue = -1
    
    intRed = CInt(lngColor And &HFF&)                   ' Calc Red
    intGreen = CInt((lngColor And &HFF00&) \ &H100&)    ' Calc Green
    intBlue = CInt((lngColor And &HFF0000) \ &H10000)   ' Calc Blue
                         
End Sub

' ***************************************************************************
' Routine:       ShowPrinter
'
' Description:   Display printer dialog box allowing the user to select a
'                printer and change print options.  This will not change
'                the default printer.  Dialog box will display in the same
'                space as the calling form.
'
' References:    Mike Williams
'                Print API instead of Common Dialog (VB6)
'                https://groups.google.com/forum/#!topic/microsoft.public.vb.controls/vda3CbgGQUI
'
'                PrintDlg
'                http://www.ex-designz.net/apidetail.asp?api_id=188
'
' Parameters:    frmName - Form calling this routine
'                strRequestedName - [Optional] - Selected printer by user
'                lngPrintFlags - [Optional] - Flags to be used with selected
'                           printer.  Default = MY_PRINTSHOW_FLAGS
'                blnCenterOnScreen - Optional - Flag designating where
'                      dialog box should be centered.
'                      TRUE  = Center on screen (Default)
'                      FALSE = Center over top of calling form
'
' Returns:       TRUE if user selects a printer
'                FALSE if user cancels selection
'
' Sample call:   Specifying Printer.DeviceName in the following
'                line will start the dialog off with the default
'                printer initially highlighted in the selection
'                box, but you can use any other string you wish.
'                For example, using "Epson" will cause the dialog
'                to start with the first printer it finds with
'                the word "Epson" in its device name.
'
'                If ShowPrinter(frmMain, Printer.DeviceName) Then
'
'                    Printer.TrackDefault = True    ' Must be TRUE
'                    Printer.ScaleMode = vbInches   ' Set to measure in inches
'                    DoEvents
'
'                    ' Set printer to start at 1/2 inch from left (x)
'                    ' edge of page and 1/2 inch down from top (y)
'                    SetPrinterOrigin 0.5, 0.5
'
'                    Printer.Print "Hello World"   ' Data to be printed
'                    Printer.EndDoc                ' Release printer control
'                    DoEvents
'                End If
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Jan-2015  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function ShowPrinter(ByRef frmName As Form, _
                   Optional ByRef strRequestedName As String = vbNullString, _
                   Optional ByRef lngPrintFlags As Long = MY_PRINTSHOW_FLAGS, _
                   Optional ByVal blnCenterOnScreen As Boolean = True) As Boolean

    Dim lngRetCode         As Long
    Dim lngDevMode         As Long
    Dim lngDevName         As Long
    Dim strNewPrinterName  As String
    Dim strOriginalPrinter As String
    Dim typPRINTDLG        As PRINTDLG_TYPE
    Dim typDEVMODE         As DEVMODE_TYPE
    Dim typDEVNAME         As DEVNAMES_TYPE
    Dim objPrinter         As Printer

    On Error GoTo ShowPrinter_Cancel
    
    Set objPrinter = Nothing                   ' Verify print object is free from memory
    ZeroMemory typPRINTDLG, Len(typPRINTDLG)   ' Empty type structures
    ZeroMemory typDEVMODE, Len(typDEVMODE)
    ZeroMemory typDEVNAME, Len(typDEVNAME)

    strNewPrinterName = vbNullString    ' Verify variables are empty
    strOriginalPrinter = vbNullString
    
    mlngFormHwnd = frmName.hWnd               ' Save handle of calling form
    strOriginalPrinter = Printer.DeviceName   ' Save current printer name

    ' If no flags values are passed
    ' then use my default values
    If lngPrintFlags < 1 Then
        lngPrintFlags = MY_PRINTSHOW_FLAGS    ' Show print dialog box
       ' lngPrintFlags = MY_PRINTSETUP_FLAGS   ' Show print setup dialog box
    End If
    
    ' Load initialization settings into typPRINTDLG,
    ' which is passed to the function
    With typPRINTDLG
        .lStructSize = Len(typPRINTDLG)
        .hwndOwner = mlngFormHwnd
        .nMinPage = 1
        .flags = lngPrintFlags   ' Define dialog box display
    End With
    
    ' Load default settings into typDEVMODE
    With typDEVMODE
        .dmDeviceName = Printer.DeviceName
        .dmSize = Len(typDEVMODE)
        .dmFields = DM_ORIENTATION Or DM_DUPLEX
        .dmPaperWidth = Printer.Width
        .dmOrientation = Printer.Orientation
        .dmPaperSize = Printer.PaperSize
        .dmDuplex = Printer.Duplex
    End With
    
    ' Allocate memory for the initialization hDevMode structure
    ' and copy the settings gathered above into this memory
    typPRINTDLG.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(typDEVMODE))
    lngDevMode = GlobalLock(typPRINTDLG.hDevMode)
    
    If lngDevMode > 0 Then
        CopyMemory ByVal lngDevMode, typDEVMODE, Len(typDEVMODE)
        lngRetCode = GlobalUnlock(typPRINTDLG.hDevMode)
    End If

    ' Load strings for default printer into typDEVNAME
    With typDEVNAME
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
        .extra = Printer.DriverName & vbNullChar & _
                 Printer.DeviceName & vbNullChar & _
                 Printer.Port & vbNullChar
    End With

    ' See if user is requesting a specific
    ' printer.  Could be a partial name.
    If Len(strRequestedName) > 0 Then
        
        ' Search for requested printer
        For Each objPrinter In Printers
            If InStr(1, objPrinter.DeviceName, strRequestedName, vbTextCompare) > 0 Then
                Set Printer = objPrinter   ' Set printer object to be highlighted
                Exit For                   ' exit For..Next loop
            End If
        Next
    Else
        ' Search for current printer
        For Each objPrinter In Printers
            If UCase$(objPrinter.DeviceName) = UCase$(strOriginalPrinter) Then
                Set Printer = objPrinter   ' Set printer object to be highlighted
                Exit For                   ' exit For..Next loop
            End If
        Next
    End If

    CenterDialogBox blnCenterOnScreen   ' Center dialog box
    DoEvents
    
    ' Make API call to display Printer Selection dialog box
    lngRetCode = PrintDialog(typPRINTDLG)
    
    ' Evaluate return code (zero = cancel selected)
    If lngRetCode <> 0 Then
        
        lngDevMode = GlobalLock(typPRINTDLG.hDevMode)              ' Get pointer to memory block
        CopyMemory typDEVMODE, ByVal lngDevMode, Len(typDEVMODE)   ' Copy structure to memory block
        lngRetCode = GlobalUnlock(typPRINTDLG.hDevMode)            ' Unlock memory block
        GlobalFree typPRINTDLG.hDevMode                            ' Free memory block
        
        ' Copy memory block data back into the structures
        lngDevName = GlobalLock(typPRINTDLG.hDevNames)             ' Get pointer to memory block
        CopyMemory typDEVNAME, ByVal lngDevName, 45                ' Copy structure to memory block
        lngRetCode = GlobalUnlock(typPRINTDLG.hDevNames)           ' Unlock memory block
        GlobalFree typPRINTDLG.hDevNames                           ' Free memory block

        ' Capture newly selected printer name
        CopyMemory typDEVNAME, ByVal lngDevName, Len(typDEVMODE)
        
        ' Capture long printer name
        strNewPrinterName = TrimStr(Mid$(typDEVNAME.extra, InStr(1, typDEVNAME.extra, Chr$(0))))   ' Find beginning of name
        strNewPrinterName = Left$(strNewPrinterName, InStr(strNewPrinterName, Chr$(0)) - 1)        ' Find end of name
        DoEvents                                                                                   ' Allow dialog box to be updated

        ' If new printer is not the same as
        ' current printer then search thru
        ' all printer names and highlight
        ' newly selected printer
        If InStr(1, Printer.DeviceName, strNewPrinterName, vbTextCompare) = 0 Then

            ' Locate selected printer
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = UCase$(strNewPrinterName) Then
                    Set Printer = objPrinter   ' Reset printer object to new printer name
                    Exit For                   ' exit For..Next loop
                End If
            Next
        End If
        
        ' Transfer settings from Devmode structure to
        ' VB printer object (these are just a few of
        ' the options to save)
        With typDEVMODE
            Printer.Copies = .dmCopies
            Printer.Duplex = .dmDuplex
            Printer.Orientation = .dmOrientation
            Printer.PaperSize = .dmPaperSize
            Printer.PrintQuality = .dmPrintQuality
            Printer.ColorMode = .dmColor
            Printer.PaperBin = .dmDefaultSource
        End With
        
        SetBkMode Printer.hdc, TRANSPARENT   ' Set background mix mode of selected printer
        
        frmName.Refresh      ' Refresh calling form
        ShowPrinter = True   ' Set flag to TRUE
        
    Else
                 
        ' User pressed CANCEL
        ShowPrinter = False
        
        ' Reset original default printer
        For Each objPrinter In Printers
            If UCase$(objPrinter.DeviceName) = UCase$(strOriginalPrinter) Then
                Set Printer = objPrinter   ' Set printer object
                Exit For                   ' exit For..Next loop
            End If
        Next

    End If
    
ShowPrinter_Cancel:
    GlobalFree typPRINTDLG.hDevNames           ' Verify memory blocks are freed
    GlobalFree typPRINTDLG.hDevMode
    ZeroMemory typPRINTDLG, Len(typPRINTDLG)   ' Empty type structures
    ZeroMemory typDEVMODE, Len(typDEVMODE)
    ZeroMemory typDEVNAME, Len(typDEVNAME)
    Set objPrinter = Nothing                   ' Free print object from memory
    On Error GoTo 0                            ' Nullify current error trap
    
End Function

' ***************************************************************************
' Routine:       SetPrinterOrigin
'
' Description:   Set printer to point to top of page. Called after printer
'                has been selected.  See ShowPrinter() flowerbox for
'                additional information.
'
' References:    Mike Williams
'                Print API instead of Common Dialog (VB6)
'                https://groups.google.com/forum/#!topic/microsoft.public.vb.controls/vda3CbgGQUI
'
'                PrintDlg
'                http://www.ex-designz.net/apidetail.asp?api_id=188
'
' Parameters:    x - Left position
'                y - Top position
'
' Example:       ' Set printer to start at 1/2 inch from left (x)
'                ' edge of page and 1/2 inch down from top (y)
'                SetPrinterOrigin 0.5, 0.5
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Jan-2015  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Sub SetPrinterOrigin(ByVal sngPosX As Single, _
                            ByVal sngPosY As Single)

    ' The calling form accesses this routine to intialize
    ' printer pointers after a printer has been selected
    
    With Printer
        .ScaleLeft = .ScaleX(GetDeviceCaps(.hdc, PHYSICALOFFSETX), vbPixels, .ScaleMode) - sngPosX
        .ScaleTop = .ScaleY(GetDeviceCaps(.hdc, PHYSICALOFFSETY), vbPixels, .ScaleMode) - sngPosY
        .CurrentX = 0
        .CurrentY = 0
    End With

End Sub

' ***************************************************************************
' Routine:       SetPBarColor
'
' Description:   Set Microsoft progress bar background and progression color
'
' Parameters:    lngPBarHwnd - Handle designating progress bar to be modified
'                lngBackgroundColor - long integer representing background
'                     color desired
'                lngForegroundColor - long integer representing progression
'                     color desired
'
'                              +-------------------------- Calling form w/ pbar name and pbar handle
'                              |             +------------ Background color
'                       |-------------|      |        +--- Progression color
' ex:  SetPBarColor frmName.PBarName.hwnd, vbWhite, vbGreen
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-Oct-2001  Randy Birch
'              http://vbnet.mvps.org/code/comctl/progressbarcolours.htm
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Sub SetPBarColor(ByRef lngPBarHwnd As Long, _
                        ByVal lngBackgroundColor As Long, _
                        ByVal lngForegroundColor As Long)

    ' Change background color
    SendMessage lngPBarHwnd, PBM_SETBKCOLOR, 0&, ByVal lngBackgroundColor
    DoEvents
    
    ' Change foreground (progression) color
    SendMessage lngPBarHwnd, PBM_SETBARCOLOR, 0&, ByVal lngForegroundColor
    DoEvents

End Sub

' ***************************************************************************
' Routine:       GetFullPath
'
' Description:   Capture path
'
' Parameters:    strPathFile - Path and file name
'
' Returns:       Just the path
'                ex:   "C:\Kens Software\Gif89.dll" --> "C:\Kens Software"
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetFullPath(ByVal strPathFile As String) As String

    Dim lngPointer As Long
    
    ' Find last backslash in string
    lngPointer = InStrRev(strPathFile, "\", -1, vbBinaryCompare)
    
    If lngPointer > 0 Then
        GetFullPath = Mid$(strPathFile, 1, lngPointer - 1)
    Else
        GetFullPath = strPathFile
    End If
    
End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

' ***************************************************************************
' Routine:       GetFunctionAddress
'
' Description:   This is a workaround function for the hook routine.
'                Procedure that receives and returns the passed value
'                of the AddressOf operator.  This workaround is needed
'                as you cannot assign AddressOf directly to a member
'                of a user defined type, but you can assign it to another
'                long and use that (as returned here).
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-Mar-2002  Randy Birch
'              Centering and Customizing the ChooseColor Common Dialog
'              http://vbnet.mvps.org/code/hooks/choosecolorcustomize.html
' 18-Jun-2014  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function GetFunctionAddress(ByVal lngFunctionAddress As Long) As Long
    
    ' Called by ShowBrowseForFolder()
    
    GetFunctionAddress = lngFunctionAddress

End Function

' ***************************************************************************
' Routine:       CenterDialogBox
'
' Description:   Center API dialog box on screen using a call-back function.
'
' Parameters:    blnCenterOnScreen - Flag designating if dialog box
'                     should be centered on the screen or over calling form
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Sep-1999  Paul Mather
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=3592&lngWId=1
' 19-Feb-2015  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Sub CenterDialogBox(ByVal blnCenterOnScreen As Boolean)

    ' Called by ShowBrowseForFolder()
    '           ShowFileOpen()
    '           ShowFileSaveAs()
    '           ShowColor()
    '           ShowPrinter()
    
    Dim lngThread   As Long
    Dim lngInstance As Long
    
    mlngHook = 0   ' Preset to FALSE
    
    ' Center dialog box
    lngInstance = GetWindowLong(mlngFormHwnd, GWL_HINSTANCE)
    lngThread = GetCurrentThreadId()
    
    If blnCenterOnScreen Then
        mlngHook = SetWindowsHookEx(WH_CBT, AddressOf CenterOnScreen, lngInstance, lngThread)
    Else
        mlngHook = SetWindowsHookEx(WH_CBT, AddressOf CenterOnForm, lngInstance, lngThread)
    End If
    
    DoEvents
    
End Sub

' ***************************************************************************
' Routine:       CenterOnScreen
'
' Description:   Center dialog box on screen regardless of monitor size
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Sep-1999  Paul Mather
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=3592&lngWId=1
' 19-Feb-2015  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function CenterOnScreen(ByVal lngMsg As Long, _
                                ByVal lngInstance As Long, _
                                ByVal lngThread As Long) As Long
    
    ' Called by CenterDialogBox()
    
    Dim lngPosX As Long   ' Upper left corner position from left side
    Dim lngPosY As Long   ' Upper left corner position from top of screen
    Dim typRECT As RECT   ' Screen coordinates
    
    If lngMsg = HCBT_ACTIVATE Then
        
        GetWindowRect lngInstance, typRECT   ' Capture window dimensions
        
        ' Calculate where to locate dialog box
        lngPosX = Screen.Width / Screen.TwipsPerPixelX / 2 - (typRECT.Right - typRECT.Left) / 2
        lngPosY = Screen.Height / Screen.TwipsPerPixelY / 2 - (typRECT.Bottom - typRECT.Top) / 2
        
        ' Move dialog box to center of screen
        SetWindowPos lngInstance, 0, lngPosX, lngPosY, 0, 0, _
                     SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        
        UnhookWindowsHookEx mlngHook   ' Release CBT hook
    
    End If
    
    CenterOnScreen = 0
    
End Function

' ***************************************************************************
' Routine:       CenterOnForm
'
' Description:   Center dialog box based on calling form location
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Sep-1999  Paul Mather
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=3592&lngWId=1
' 19-Feb-2015  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function CenterOnForm(ByVal lngMsg As Long, _
                              ByVal lngInstance As Long, _
                              ByVal lngThread As Long) As Long
    
    ' Called by CenterDialogBox()
    
    Dim lngPosX As Long   ' Upper left corner position from left side
    Dim lngPosY As Long   ' Upper left corner position from top of screen
    Dim typRECT As RECT   ' Screen coordinates
    Dim typFORM As RECT   ' Form coordinates
    
    ' Show dialog box centered over calling form
    If lngMsg = HCBT_ACTIVATE Then
        
        ' Get coordinates of calling form and
        ' dialog box so the dialog box can be
        ' centered over the calling form
        GetWindowRect mlngFormHwnd, typFORM   ' Capture calling form dimensions
        GetWindowRect lngInstance, typRECT    ' Capture window dimensions
        
        ' Calculate where to locate dialog box based
        ' on calling form location and dimensions
        lngPosX = (typFORM.Left + (typFORM.Right - typFORM.Left) / 2) - ((typRECT.Right - typRECT.Left) / 2)
        lngPosY = (typFORM.Top + (typFORM.Bottom - typFORM.Top) / 2) - ((typRECT.Bottom - typRECT.Top) / 2)
        
        ' Move dialog box to new location
        SetWindowPos lngInstance, 0, lngPosX, lngPosY, 0, 0, _
                     SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        
        UnhookWindowsHookEx mlngHook   ' Release CBT hook
    
    End If
    
    CenterOnForm = 0
    
End Function

' ***************************************************************************
' Routine:       FindStartingFolder
'
' Description:   How to open the Browse For Folder Dialog with the ability
'                to specify a default folder to start in.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 20-Feb-2000  Stephen Fonnesbeck  steev@xmission.com
'              Browse for folder using callback
'              http://www.xmission.com/~steev
' 18-Jun-2014  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function FindStartingFolder(ByVal lngHwnd As Long, _
                                    ByVal lngMsg As Long, _
                                    ByVal wParam As Long, _
                                    ByVal lParam As Long) As Long
   
    ' Called by ShowBrowseForFolder()
    
    Dim lngRetCode As Long
    Dim strBuffer  As String
  
    On Error Resume Next  ' Suggested by Microsoft to prevent an error from
                          ' propagating back into the calling process
     
    Select Case lngMsg
          
           Case BFFM_INITIALIZED
                Call SendMessage(lngHwnd, BFFM_SETSELECTION, True, ByVal mstrStartFolder)
               
           Case BFFM_SELCHANGED
                strBuffer = Space$(MAX_AMT)   ' Preload with spaces
                lngRetCode = SHGetPathFromIDList(wParam, strBuffer)
            
                If lngRetCode = 1 Then
                    Call SendMessage(lngHwnd, BFFM_SETSTATUSTEXT, False, ByVal strBuffer)
                End If
    End Select
   
    FindStartingFolder = 0

    On Error GoTo 0   ' Nullify this error trap
  
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
Private Function IsArrayInitialized(ByVal avntData As Variant) As Long

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
Private Function ShrinkToFit(ByVal strPath As String, _
                             ByVal intMaxLength As Integer) As String

    Dim strBuffer As String

    strPath = TrimStr(strPath)   ' Remove unwanted leading\trailing characters

    ' See if ellipses need to be inserted into the path
    If Len(strPath) <= intMaxLength Then
        ShrinkToFit = strPath
        Exit Function
    End If

    ' intMaxLength is the maximum number of characters to be contained in
    ' the new string, including the terminating NULL character. For example,
    ' if intMaxLength = 8, the resulting string would contain a maximum of
    ' seven data characters plus a terminating null.
    '
    ' Because of this, add 1 to the value passed as intMaxLength to ensure
    ' the resulting string is the size requested.
    intMaxLength = intMaxLength + 1
    strBuffer = Space$(MAX_AMT)
    PathCompactPathEx strBuffer, strPath, intMaxLength, 0&

    ' Return readjusted data string
    ShrinkToFit = TrimStr(strBuffer)

End Function

