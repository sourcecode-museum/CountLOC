VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cManifest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ***************************************************************************
' Module:       cManifest (clsManifest.cls)
'
' Important:     If you should get a Side-by-side error message, after
'                making changes to this module, try this:
'
'                1.  Delete current manifest file (ex: app_name.exe.manifest)
'                2.  Restart this application.  A new manifest file should be
'                    created.
'
' Description:  Names of elements and attributes in a manifest file are case
'               sensitive. The values of elements and attributes are not case
'               sensitive, except for the value of the type attribute.
'
'               A manifest is a XML file that contains settings that informs
'               Windows how to handle a program when it is started. The
'               manifest can be embedded inside the program file (as a resource)
'               or it can be located in a separate external XML file. If the
'               manifest is placed in a separate file, then the file must be
'               located in the same folder as the executable file and it must
'               have same filename as the program file, but with a ".manifest"
'               filename extension added at the end (e.g "MYAPP.EXE.manifest").
'
'               Manifests are also often used to inform Windows which version
'               of a DLL a program is dependent of. By using manifests same
'               DLL can exists in the computer in different versions and the
'               program chooses which version of the DLL to load. This is
'               sometimes referred to as the side-by-side technology in Windows.
'               DLL's that exists in different versions in the computer are
'               sometimes referred to as side-by-side assemblies.
'
'               One common use of manifests and the side-by-side technology
'               is to inform Windows which version of the Windows common
'               controls the program need to use. The Windows common controls
'               are used to draw menus, dialog boxes, buttons, etc. in an
'               application and two different styles can be used: a classic
'               visual style and a modern visual style. The "modern visual
'               style" is the modern 3D style that was introduced with
'               Windows XP and that has evolved even more in Windows Vista
'               or Windows 7, with shadows, light effects and metal effects.
'
'               These files have been around for a while, and the new
'               section added for UAC control is the requestedExecutionLevel
'               element. The most effected by UAC Control are Windows Vista
'               and newer versions. The level attribute may be one of the
'               following values:
'
'               --------------------------------------------------------------
'               Level               Description
'               --------------------------------------------------------------
'               asInvoker             Does not require elevation, runs without
'                                     requesting elevation using privileges of
'                                     its parent process.
'
'               highestAvailable      Requests the highest available privilege
'                                     tokens of its parent process. An admin
'                                     account will attempt to elevate to full
'                                     administrator level, but a standard
'                                     account will only request elevation to
'                                     its own highest set of access tokens.
'
'               requireAdministrator  Requires elevation to full administrator
'                                     privileges.
'               --------------------------------------------------------------
'
'               Most applications should use the asInvoker level, as this
'               will not prompt the user for elevation every time it is run,
'               and will use the standard set of privileges to access the
'               unprotected resources it needs. Only applications that require
'               access to protected resources should use the higher access levels.
'
'               The uiAccess attribute determines whether the application
'               requires access to any protected UI elements, such as system
'               dialog boxes or higher-level processes. Only signed applications
'               may do this as it requires additional security tokens. This value
'               defaults to false.
'
' References:   Application Manifests
'               http://msdn.microsoft.com/en-us/library/aa374191(v=vs.85).aspx
'
'               What is a Manifest (in Windows)?
'               http://www.samlogic.net/articles/manifest.htm
'
'               Making Your Application UAC Aware
'               TheCodeKing,  10 Mar 2007
'               http://www.codeproject.com/Articles/17968/Making-Your-Application-UAC-Aware
'
'               Targeting your application for Windows 8.1
'               http://msdn.microsoft.com/en-us/library/windows/desktop/dn481241(v=vs.85).aspx
'
'               SupportedOS manifest for Windows 10
'               https://social.msdn.microsoft.com/Forums/en-US/07cbfc3a-bced-45b7-80d2-a9d32a7c95d4/supportedos-manifest-for-windows-10?forum=windowsgeneraldevelopmentissues
'
'               Manifest Madness
'               This blog post discusses Application Manifest elements for
'               use in Win32 desktop applications. This is not about the App
'               package manifest (aka AppX) used for Windows Store apps for
'               Windows 8.x.
'               http://blogs.msdn.com/b/chuckw/archive/2013/09/10/manifest-madness.aspx
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added IsPathValid() routine.
' 28-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated logic in NeedManifest() routine
' 31-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated error trap by clearing error number in InitComctl32()
'              routine.
' 28-Oct-2013  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote data for inside of manifest file
' 28-Jan-2014  Kenneth Ives  kenaso@tx.rr.com
'              Updated data for inside of manifest file when application
'              began recieving side-by-side error message because of
'              previous manifest file being present.
' 13-Oct-2014  Kenneth Ives  kenaso@tx.rr.com
'              Evaluate operating system to determine if a manifest file
'              needs to be created
' 12-Jul-2015  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote CreateManifestFile() routine
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  ' Set of bit flags that indicate which common
  ' control classes will be loaded from the DLL.
  ' The dwICC value of INIT_COMMON_CTRLS can be
  ' a combination of the following:
  Private Const ICC_ANIMATE_CLASS      As Long = &H80&     ' Load animate control class
  Private Const ICC_BAR_CLASSES        As Long = &H4&      ' Load toolbar, status bar, trackbar, tooltip control classes
  Private Const ICC_COOL_CLASSES       As Long = &H400&    ' Load rebar control class
  Private Const ICC_DATE_CLASSES       As Long = &H100&    ' Load date and time picker control class
  Private Const ICC_HOTKEY_CLASS       As Long = &H40&     ' Load hot key control class
  Private Const ICC_INTERNET_CLASSES   As Long = &H800&    ' Load IP address class
  Private Const ICC_LINK_CLASS         As Long = &H8000&   ' Load a hyperlink control class. Must have trailing ampersand.
  Private Const ICC_LISTVIEW_CLASSES   As Long = &H1&      ' Load list-view and header control classes
  Private Const ICC_NATIVEFNTCTL_CLASS As Long = &H2000&   ' Load a native font control class
  Private Const ICC_PAGESCROLLER_CLASS As Long = &H1000&   ' Load pager control class
  Private Const ICC_PROGRESS_CLASS     As Long = &H20&     ' Load progress bar control class
  Private Const ICC_STANDARD_CLASSES   As Long = &H4000&   ' Load user controls that include button, edit, static, listbox,
                                                           '      combobox, scrollbar
  Private Const ICC_TREEVIEW_CLASSES   As Long = &H2&      ' Load tree-view and tooltip control classes
  Private Const ICC_TAB_CLASSES        As Long = &H8&      ' Load tab and tooltip control classes
  Private Const ICC_UPDOWN_CLASS       As Long = &H10&     ' Load up-down control class
  Private Const ICC_USEREX_CLASSES     As Long = &H200&    ' Load ComboBoxEx class
  Private Const ICC_WIN95_CLASSES      As Long = &HFF&     ' Load animate control, header, hot key, list-view, progress bar,
                                                           '      status bar, tab, tooltip, toolbar, trackbar, tree-view,
                                                           '      and up-down control classes

  ' All bit flags combined. Total value = &HFFFF& (65535)
  Private Const ALL_FLAGS As Long = ICC_ANIMATE_CLASS Or ICC_BAR_CLASSES Or ICC_COOL_CLASSES Or _
                                    ICC_DATE_CLASSES Or ICC_HOTKEY_CLASS Or ICC_INTERNET_CLASSES Or _
                                    ICC_LINK_CLASS Or ICC_LISTVIEW_CLASSES Or ICC_NATIVEFNTCTL_CLASS Or _
                                    ICC_PAGESCROLLER_CLASS Or ICC_PROGRESS_CLASS Or ICC_STANDARD_CLASSES Or _
                                    ICC_TREEVIEW_CLASSES Or ICC_TAB_CLASSES Or ICC_UPDOWN_CLASS Or _
                                    ICC_USEREX_CLASSES Or ICC_WIN95_CLASSES
                                                    
' ***************************************************************************
' Type structures
' ***************************************************************************
  Private Type INIT_COMMON_CTRLS
      dwSize As Long   ' size of this structure
      dwICC  As Long   ' flags indicating which classes to be initialized
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' Initializes the entire common control dynamic-link library. Exported by
  ' all versions of Comctl32.dll.
  Private Declare Sub InitCommonControls Lib "comctl32" ()
  
  ' Initializes specific common controls classes from the common control
  ' dynamic-link library. Returns TRUE (non-zero) if successful, or FALSE
  ' otherwise. Began being exported with Comctl32.dll version 4.7
  ' (IE3.0 & later).
  Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
          (iccex As INIT_COMMON_CTRLS) As Boolean

' ***************************************************************************
' Module Variables
'                    +-------------- Module level designator
'                    |  +----------- Data type (Long)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m lng MajorVer
' Variable name:     mlngMajorVer
' ***************************************************************************
  Private mlngMajorVer As Long
  Private mlngMinorVer As Long


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let MajorVersion(ByVal lngData As Long)
    mlngMajorVer = lngData
End Property

Public Property Let MinorVersion(ByVal lngData As Long)
    mlngMinorVer = lngData
End Property


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       InitComctl32
'
' Description:   This will create the XP Manifest file and utilize it. You
'                will only see the results when the exe (not in the IDE)
'                is run.  This routine is usually called before any forms
'                are loaded.  (See modMain.bas)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jan-2006  Randy Birch
'              http://vbnet.mvps.org/
' 03-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 31-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated error trap by clearing error number
' ***************************************************************************
Public Sub InitComctl32()

    Dim typICC As INIT_COMMON_CTRLS
    
    If CreateManifestFile() Then
         
        On Error GoTo Use_Old_Version
             
        With typICC
            .dwSize = LenB(typICC)
            .dwICC = ALL_FLAGS
        End With
        
        ' VB will generate error 453 "Specified DLL function not found"
        ' if InitCommonControlsEx can't be located in the library. The
        ' error is trapped and the original InitCommonControls is called
        ' instead below.
        If InitCommonControlsEx(typICC) = 0 Then
            InitCommonControls
        End If
        
    End If
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub
    
Use_Old_Version:
    Err.Clear            ' Clear any error codes
    InitCommonControls
    On Error GoTo 0      ' Nullify this error trap
    
End Sub


' ***************************************************************************
' ****               Internal Procedures and Functions                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       CreateManifestFile
'
' Description:   Create a manifest file based on the Windows major version
'                number.
'
' Important:     If you should get a Side-by-side error message, after
'                making changes to this module, try this:
'
'                1.  Delete current manifest file (ex: app_name.exe.manifest)
'                2.  Restart this application.  A new manifest file should be
'                    created.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jan-2006  Randy Birch
'              http://vbnet.mvps.org/
' 28-Oct-2013  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 13-Oct-2014  Kenneth Ives  kenaso@tx.rr.com
'              Evaluate operating system to determine if a manifest file
'              needs to be created
' 18-Aug-2015  Kenneth Ives  kenaso@tx.rr.com
'              Updated to test for Windows major version six or newer
' ***************************************************************************
Private Function CreateManifestFile() As Boolean
    
    Dim hFile            As Long
    Dim strXML           As String
    Dim strCompany       As String
    Dim strExeName       As String
    Dim strFileName      As String
    Dim blnVer6orNewer   As Boolean
    Dim blnWin7orEarlier As Boolean
    
    On Error Resume Next

    CreateManifestFile = False      ' Preset to FALSE
    strXML = vbNullString           ' Verify output string is empty
    strCompany = "Kens.Software."   ' Enter unique name here (Periods are delimiters)
    strExeName = App.EXEName        ' EXE name without an extension
    strFileName = QualifyPath(App.Path) & strExeName & ".exe.manifest"
    
    ' See if a manifest file is needed
    If NeedManifest(blnVer6orNewer, blnWin7orEarlier) Then
        
        ' Create manifest file for Vista, 7, 8, 8.1, 10
        ' Manifest file contents are mostly case sensitive
        If blnVer6orNewer Then
           
            ' Contents of manifest file if Windows major version is 6
            strXML = strXML & "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & " encoding=" & Chr$(34) & _
                              "UTF-8" & Chr$(34) & " standalone=" & Chr$(34) & "yes" & Chr$(34) & "?>" & vbNewLine
                              
            If blnWin7orEarlier Then
                ' Vista, Win 7
                strXML = strXML & "<assembly xmlns=" & Chr$(34) & "urn:schemas-microsoft-com:asm.v1" & Chr$(34) & _
                                  " manifestVersion=" & Chr$(34) & "1.0" & Chr$(34) & ">" & vbNewLine
            Else
                ' Win 8, 8.1, 10
                strXML = strXML & "<assembly manifestVersion=" & Chr$(34) & "1.0" & Chr$(34) & " xmlns=" & Chr$(34) & _
                                  "urn:schemas-microsoft-com:asm.v1" & Chr$(34) & " xmlns:asmv3=" & Chr$(34) & _
                                  "urn:schemas-microsoft-com:asm.v3" & Chr$(34) & ">" & vbNewLine
            End If
            
            ' Minimum version of this application
            strXML = strXML & "  <assemblyIdentity" & vbNewLine
            strXML = strXML & "    type=" & Chr$(34) & "win32" & Chr$(34) & vbNewLine
            strXML = strXML & "    name=" & Chr$(34) & strExeName & Chr$(34) & vbNewLine
            strXML = strXML & "    version=" & Chr$(34) & "1.0.0.0" & Chr$(34) & vbNewLine
            strXML = strXML & "    processorArchitecture=" & Chr$(34) & "x86" & Chr$(34) & vbNewLine
            strXML = strXML & "  />" & vbNewLine
            
            ' Company and application name
            strXML = strXML & "  <description>" & strCompany & strExeName & "</description>" & vbNewLine
            
            ' Minimum version of Microsoft controls used with this application
            strXML = strXML & "  <dependency>" & vbNewLine
            strXML = strXML & "    <dependentAssembly>" & vbNewLine
            strXML = strXML & "      <assemblyIdentity" & vbNewLine
            strXML = strXML & "        type=" & Chr$(34) & "win32" & Chr$(34) & vbNewLine
            strXML = strXML & "        name=" & Chr$(34) & "Microsoft.Windows.Common-Controls" & Chr$(34) & vbNewLine
            strXML = strXML & "        version=" & Chr$(34) & "6.0.0.0" & Chr$(34) & vbNewLine
            strXML = strXML & "        processorArchitecture=" & Chr$(34) & "x86" & Chr$(34) & vbNewLine
            strXML = strXML & "        publicKeyToken=" & Chr$(34) & "6595b64144ccf1df" & Chr$(34) & vbNewLine
            strXML = strXML & "      />" & vbNewLine
            strXML = strXML & "    </dependentAssembly>" & vbNewLine
            strXML = strXML & "  </dependency>" & vbNewLine

            ' Only recognized by Windows 7 or newer
            strXML = strXML & "  <trustInfo xmlns=" & Chr$(34) & "urn:schemas-microsoft-com:asm.v3" & Chr$(34) & ">" & vbNewLine
            strXML = strXML & "    <security>" & vbNewLine
            strXML = strXML & "      <requestedPrivileges>" & vbNewLine
            strXML = strXML & "        <requestedExecutionLevel level=" & Chr$(34) & "highestAvailable" & Chr$(34) & _
                                                                      " uiAccess=" & Chr$(34) & "false" & Chr$(34) & "/>" & vbNewLine
            strXML = strXML & "      </requestedPrivileges>" & vbNewLine
            strXML = strXML & "    </security>" & vbNewLine
            strXML = strXML & "  </trustInfo>" & vbNewLine
            
            ' Current operating system
            strXML = strXML & "  <compatibility xmlns=" & Chr$(34) & "urn:schemas-microsoft-com:compatibility.v1" & Chr$(34) & ">" & vbNewLine
            strXML = strXML & "    <application>" & vbNewLine
            strXML = strXML & "      <!-- Windows 10 / Server 2016 -->" & vbNewLine
            strXML = strXML & "      <supportedOS Id=" & Chr$(34) & "{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}" & Chr$(34) & "/>" & vbNewLine
            strXML = strXML & "      <!-- Windows 8.1 / Server 2012 R2 -->" & vbNewLine
            strXML = strXML & "      <supportedOS Id=" & Chr$(34) & "{1f676c76-80e1-4239-95bb-83d0f6d0da78}" & Chr$(34) & "/>" & vbNewLine
            strXML = strXML & "      <!-- Windows 8 / Server 2012 -->" & vbNewLine
            strXML = strXML & "      <supportedOS Id=" & Chr$(34) & "{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}" & Chr$(34) & "/>" & vbNewLine
            strXML = strXML & "      <!-- Windows 7 / Server 2008 R2 -->" & vbNewLine
            strXML = strXML & "      <supportedOS Id=" & Chr$(34) & "{35138b9a-5d96-4fbd-8e2d-a2440225f93a}" & Chr$(34) & "/>" & vbNewLine
            strXML = strXML & "      <!-- Windows Vista / Server 2008 -->" & vbNewLine
            strXML = strXML & "      <supportedOS Id=" & Chr$(34) & "{e2011457-1546-43c5-a5fe-008deee3d3f0}" & Chr$(34) & "/>" & vbNewLine
            strXML = strXML & "    </application>" & vbNewLine
            strXML = strXML & "  </compatibility>" & vbNewLine
            
            strXML = strXML & "</assembly>"
        Else
            ' Windows XP manifest file
            strXML = strXML & "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & " encoding=" & Chr$(34) & _
                              "UTF-8" & Chr$(34) & " standalone=" & Chr$(34) & "yes" & Chr$(34) & "?>" & vbNewLine
                              
            strXML = strXML & "<assembly xmlns=" & Chr$(34) & "urn:schemas-microsoft-com:asm.v1" & Chr$(34) & _
                              " manifestVersion=" & Chr$(34) & "1.0" & Chr$(34) & ">" & vbNewLine
                                          
            ' Minimum version of this application
            strXML = strXML & "  <assemblyIdentity" & vbNewLine
            strXML = strXML & "    type=" & Chr$(34) & "win32" & Chr$(34) & vbNewLine
            strXML = strXML & "    name=" & Chr$(34) & strExeName & Chr$(34) & vbNewLine
            strXML = strXML & "    version=" & Chr$(34) & "1.0.0.0" & Chr$(34) & vbNewLine
            strXML = strXML & "    processorArchitecture=" & Chr$(34) & "x86" & Chr$(34) & vbNewLine
            strXML = strXML & "    publicKeyToken=" & Chr$(34) & "0000000000000000" & Chr$(34) & vbNewLine
            strXML = strXML & "  />" & vbNewLine
                        
            ' Company and application name
            strXML = strXML & "  <description>" & strCompany & strExeName & "</description>" & vbNewLine
            
            ' Minimum version of Microsoft controls used with this application
            strXML = strXML & "  <dependency>" & vbNewLine
            strXML = strXML & "    <dependentAssembly>" & vbNewLine
            strXML = strXML & "      <assemblyIdentity" & vbNewLine
            strXML = strXML & "        type=" & Chr$(34) & "win32" & Chr$(34) & vbNewLine
            strXML = strXML & "        name=" & Chr$(34) & "Microsoft.Windows.Common-Controls" & Chr$(34) & vbNewLine
            strXML = strXML & "        version=" & Chr$(34) & "6.0.0.0" & Chr$(34) & vbNewLine
            strXML = strXML & "        processorArchitecture=" & Chr$(34) & "x86" & Chr$(34) & vbNewLine
            strXML = strXML & "        publicKeyToken=" & Chr$(34) & "6595b64144ccf1df" & Chr$(34) & vbNewLine
            strXML = strXML & "      />" & vbNewLine
            strXML = strXML & "    </dependentAssembly>" & vbNewLine
            strXML = strXML & "  </dependency>" & vbNewLine
            strXML = strXML & "</assembly>"
        End If
        
        ' Create a new file or
        ' empty an existing file
        hFile = FreeFile
        Open strFileName For Output As #hFile
        Print #hFile, strXML
        Close #hFile
    
    Else
        
        ' Manifest file is not needed
        If IsPathValid(strFileName) Then
            Kill strFileName   ' Remove old file
            DoEvents
        End If
    
    End If
    
    CreateManifestFile = True   ' Successful finish
    On Error GoTo 0             ' Nullify this error trap
    
End Function

' ***************************************************************************
' Routine:       NeedManifest
'
' Description:   Evaluate current operating system
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Aug-2015  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' **************************************************************************
Private Function NeedManifest(ByRef blnVer6orNewer As Boolean, _
                              ByRef blnWin7orEarlier As Boolean) As Boolean
                              
    ' Called by CreateManifestFile()
    
    ' Preset flags to FALSE
    NeedManifest = False
    blnVer6orNewer = False
    blnWin7orEarlier = False
    
    Select Case mlngMajorVer
           Case 5
                If mlngMinorVer > 0 Then
                    NeedManifest = True
                End If
           
           Case 6
                NeedManifest = True
                blnVer6orNewer = True
                        
                If mlngMinorVer < 2 Then
                    blnWin7orEarlier = True
                End If
    
           Case 10
                blnVer6orNewer = True
                NeedManifest = True
    End Select
       
End Function

