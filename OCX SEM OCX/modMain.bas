Attribute VB_Name = "modMain"
Option Explicit
'
'MMM
'===
'
'"Make My Manifest"
'
'A program that analyzes the project (*.vbp) file of a
'compiled VB 6.0 EXE and produces a Registration Free COM
'XCopy-deployable execution package from it.
'
'Reg-Free COM only works on Windows XP and later operating
'systems such as Windows 2003 and Windows Vista.
'
'An XCopy folder is created inside the project folder.
'The EXE and all project ActiveX OCXs and DLLs are copied
'into the folder.  An application manifest is created for
'the EXE (which also enables XP Styles) along with an
'assembly manifest for the ActiveX libraries.
'
'Programs must still properly call the InitCommonControls
'API in many cases to enable XP Styles.
'
'Use of licensed ActiveX components may require manually
'adding licenses within the EXE via the VB Licenses
'collection.
'
'
'Legal
'=====
'
'This program is free for all to use or modify.  It is
'"as is" software, with no guarantee that it will not
'result in damage to computers or associated data files.
'No promise of support is implied and none should be
'inferred.  It should be considered experimental software.
'
'
'modMain
'=======
'
'This module is the main program, which begins execution
'in Sub Main() below.
'
'A substantial amount of boilerplate text is stored within
'the compiled EXE as internal resources.
'
'This module loads and shows frmStatus non-modally to use
'it for UI functions like progress and status display and
'requesting the name of the VB project to be processed.
'

'====================== Consts ===========================

Private Const EXCLUDED_LIBS As String = "stdole#"
Private Const PACKAGE_FOLDER As String = "XCopy"
Private Const ASSEM_NAME_SUFFIX As String = ".X"

'=================== Project data ========================

Private strProjName As String
Private strProjMajorVer As String
Private strProjMinorVer As String
Private strProjRevisionVer As String
Private strProjVersion As String
Private strProjDescription As String
Private strProjCompany As String
Private strProjEXE As String
Private strProjAssemName As String
Private colLibData As New Collection

'================== Main procedure =======================

Private Sub Main()
    Dim strProjFile As String
    Dim strProjDir As String
    Dim strPackageDir As String
    Dim strProjEXEFile As String
    Dim intResult As Integer
    
    With frmStatus
        .Show vbModeless
        
        .Log App.ProductName & " " _
           & CStr(App.Major) & "." & CStr(App.Minor) _
           & ".0." & CStr(App.Revision)
        
        strProjFile = .GetProjectFile()
        If Len(strProjFile) = 0 Then
            .Log
            .Log "## Manifest generation canceled. ##"
            .Done
            Exit Sub
        End If
        
        strProjDir = PathToFileName(strProjFile)
        
        ProjectProcess strProjFile
        
        If Len(strProjEXE) = 0 Then
            .Log
            .Log "## Can't build XCopy package.  No EXE name ##"
            .Log "## found in project file.  Compile your    ##"
            .Log "## program and try again.                  ##"
            MsgBox "Can't build XCopy package.  No EXE name" & vbNewLine _
                 & "found in project file.  Compile your" & vbNewLine _
                 & "program and try again.", vbOKOnly Or vbCritical, _
                   App.FileDescription
            .Done
            Exit Sub
        End If
        
        strProjEXEFile = strProjDir & strProjEXE
        If FileNotPresent(strProjEXEFile) Then
            .Log
            .Log "## Can't build XCopy package.  No EXE file ##"
            .Log "## found in project folder.  Compile your  ##"
            .Log "## program and try again.                  ##"
            MsgBox "Can't build XCopy package.  No EXE file" & vbNewLine _
                 & "found in project folder.  Compile your" & vbNewLine _
                 & "program and try again.", vbOKOnly Or vbCritical, _
                   App.FileDescription
            .Done
            Exit Sub
        End If
        
        strPackageDir = strProjDir & PACKAGE_FOLDER
        If Len(Dir$(strPackageDir, vbDirectory)) > 0 Then
            .Log
            .Log "** XCopy package folder exists. **"
            intResult = _
                MsgBox("XCopy package folder exists." & vbNewLine & vbNewLine _
                     & "Replace with new package?", vbYesNo Or vbQuestion, _
                       App.FileDescription)
            If intResult = vbNo Then
                .Log
                .Log "** No changes made. **"
                .Done
                Exit Sub
            Else
                .Log
                .Log "** Replacing package folder contents. **"
                Kill strPackageDir & "\*.*"
            End If
        Else
            .Log
            .Log "** Creating XCopy package folder. **"
            MkDir strPackageDir
        End If
        
        strPackageDir = strPackageDir & "\"
        PackageApp strProjDir, strPackageDir
        PackageAssem strPackageDir
        
        .Log
        .Log "--Complete--"
        .Done
    End With
End Sub

'================ Packaging procedures ===================

Private Sub PackageApp(ByVal ProjDir As String, ByVal PackageDir As String)
    Dim bytAppMan() As Byte
    Dim strAppMan As String
    Dim intFAppMan As Integer
    Dim strDescTag As String
    
    With frmStatus
        .Log
        .Log "** Writing application manifest. **"
    End With
    intFAppMan = FreeFile()
    Open PackageDir & strProjEXE & ".manifest" For Output As #intFAppMan
    
    bytAppMan = LoadResData("APPMAN", "TEXT")
    strAppMan = StrConv(bytAppMan, vbUnicode)
    Erase bytAppMan
    strAppMan = Replace$(strAppMan, _
                         "[APPNAME]", _
                         strProjCompany & "." & strProjName)
    strAppMan = Replace$(strAppMan, _
                         "[VERSION]", _
                         strProjVersion)
    strDescTag = strProjDescription
    If Len(strDescTag) > 0 Then
        strDescTag = "<description>" & strDescTag & "</description>"
    End If
    strAppMan = Replace$(strAppMan, _
                         "[APPDESC]", _
                         strDescTag)
    strAppMan = Replace$(strAppMan, _
                         "[ASSEMNAME]", _
                         strProjAssemName)
    Print #intFAppMan, strAppMan;
        
    Close #intFAppMan
        
    With frmStatus
        .Log
        .Log "** Copying project EXE. **"
    End With
    FileCopy ProjDir & strProjEXE, PackageDir & strProjEXE
End Sub

Private Sub PackageAssem(ByVal PackageDir As String)
    Dim bytAssemManTempl() As Byte
    Dim strAssemManTempl As String
    Dim strAssemMan As String
    Dim intFAssemMan As Integer
    Dim ldLib As LibData
    Dim lngClass As Long
    
    With frmStatus
        .Log
        .Log "** Writing assembly manifest and **"
        .Log "** copying component libraries.  **"
    End With
    intFAssemMan = FreeFile()
    Open PackageDir & strProjAssemName & ".manifest" For Output As #intFAssemMan
    
    'Write assembly manifest head.
    bytAssemManTempl = LoadResData("ASSEMHD", "TEXT")
    strAssemManTempl = StrConv(bytAssemManTempl, vbUnicode)
    Erase bytAssemManTempl
    strAssemMan = Replace$(strAssemManTempl, _
                           "[ASSEMNAME]", _
                           strProjAssemName)
    strAssemMan = Replace$(strAssemMan, _
                           "[VERSION]", _
                           strProjVersion)
    Print #intFAssemMan, strAssemMan;
    
    For Each ldLib In colLibData
        'Write assembly manifest component file head.
        bytAssemManTempl = LoadResData("ASSEMFH", "TEXT")
        strAssemManTempl = StrConv(bytAssemManTempl, vbUnicode)
        Erase bytAssemManTempl
        strAssemMan = Replace$(strAssemManTempl, _
                               "[LIBFILE]", _
                               SimpleFileName(ldLib.FileLocation))
        Print #intFAssemMan, strAssemMan;
        
        FileCopy ldLib.FileLocation, PackageDir & SimpleFileName(ldLib.FileLocation)
        
        For lngClass = 0 To ldLib.Count - 1
            'Write assembly manifest component file class instance.
            bytAssemManTempl = LoadResData("ASSEMCL", "TEXT")
            strAssemManTempl = StrConv(bytAssemManTempl, vbUnicode)
            Erase bytAssemManTempl
            strAssemMan = Replace$(strAssemManTempl, _
                                   "[CLSID]", _
                                   ldLib.Class(lngClass).CLSID)
            Print #intFAssemMan, strAssemMan;
        Next
        
        'Write assembly manifest component file foot.
        bytAssemManTempl = LoadResData("ASSEMFF", "TEXT")
        strAssemManTempl = StrConv(bytAssemManTempl, vbUnicode)
        Erase bytAssemManTempl
        strAssemMan = Replace$(strAssemManTempl, _
                               "[LIBID]", _
                               ldLib.LIBID)
        strAssemMan = Replace$(strAssemMan, _
                               "[VERSION]", _
                               ldLib.Version)
        Print #intFAssemMan, strAssemMan;
    Next
    
    'Write assembly manifest foot.
    bytAssemManTempl = LoadResData("ASSEMFT", "TEXT")
    strAssemMan = StrConv(bytAssemManTempl, vbUnicode)
    Erase bytAssemManTempl
    Print #intFAssemMan, strAssemMan;
    
    Close #intFAssemMan
End Sub

'=========== Project processing procedures ===============

Private Sub ProjectProcess(ByVal ProjectFileName As String)
    ProjectProcessFileScan ProjectFileName
    ProjectProcessDump
End Sub

Private Sub ProjectProcessDump()
    Dim ldLib As LibData
    Dim lngClass As Long
    
    With frmStatus
        .Log "--Project--"
        .Log strProjCompany & "." & strProjName
        .Log strProjVersion
        If Len(strProjDescription) > 0 Then
            .Log strProjDescription
        End If
        .Log
        .Log "--Libraries: " & CStr(colLibData.Count) & " --"
        For Each ldLib In colLibData
            .Log
            .Log ldLib.Name
            .Log "  " & ldLib.Description
            .Log "  " & ldLib.FileLocation
            .Log "  " & ldLib.LIBID
            .Log
            .Log "    --Classes: " & CStr(ldLib.Count) & " --"
            For lngClass = 0 To ldLib.Count - 1
                .Log
                .Log "    " & ldLib.Class(lngClass).Name
                .Log "    " & ldLib.Class(lngClass).CLSID
            Next
        Next
    End With
End Sub

Private Sub ProjectProcessFileScan(ByVal ProjectFileName As String)
    Dim intFProj As Integer
    Dim strSource As String
    Dim strSParts() As String
    Dim ldLib As LibData
    
    With frmStatus
        .Log "--Scanning project file--"
        .Log ProjectFileName
        .Log
        intFProj = FreeFile()
        Open ProjectFileName For Input As #intFProj
        Do While Not EOF(intFProj)
            Line Input #intFProj, strSource
            strSParts = Split(strSource, "=")
            If UBound(strSParts) > 0 Then
                Select Case UCase$(strSParts(0))
                    Case "REFERENCE"
                        Set ldLib = New LibData
                        With ldLib
                            .LoadReference strSParts(1)
                            If NotExcludedLib(.Name) Then
                                colLibData.Add ldLib, .Name
                            End If
                        End With
                        
                    Case "OBJECT"
                        Set ldLib = New LibData
                        With ldLib
                            .LoadObject strSParts(1)
                            If NotExcludedLib(.Name) Then
                                colLibData.Add ldLib, .Name
                            End If
                        End With
                        
                    Case "NAME"
                        strProjName = DQ(strSParts(1))
                        
                    Case "MAJORVER"
                        strProjMajorVer = strSParts(1)
                        
                    Case "MINORVER"
                        strProjMinorVer = strSParts(1)
                        
                    Case "REVISIONVER"
                        strProjRevisionVer = strSParts(1)
                        
                    Case "DESCRIPTION"
                        strProjDescription = DQ(strSParts(1))
                        
                    Case "VERSIONCOMPANYNAME"
                        strProjCompany = Replace$(DQ(strSParts(1)), " ", ".")
                        
                    Case "EXENAME32"
                        strProjEXE = DQ(strSParts(1))
                End Select
            End If
        Loop
        Close #intFProj
    End With
    strProjVersion = strProjMajorVer & "." & strProjMinorVer & ".0." & strProjRevisionVer
    strProjAssemName = strProjName & ASSEM_NAME_SUFFIX
End Sub

'================= Helper procedures =====================

Private Function DQ(ByVal QuotedText As String) As String
    DQ = Mid$(QuotedText, 2, Len(QuotedText) - 2)
End Function

Private Function FileNotPresent(ByVal FileName As String) As Boolean
    On Error Resume Next
    FileLen FileName
    FileNotPresent = Err.Number <> 0
End Function

Private Function NotExcludedLib(ByVal LibName As String) As Boolean
    NotExcludedLib = InStr(EXCLUDED_LIBS, LibName & "#") < 1
End Function

Private Function PathToFileName(ByVal FQFileName As String) As String
    Dim lngBack As Long
    
    lngBack = InStrRev(FQFileName, "\")
    PathToFileName = Left$(FQFileName, lngBack)
End Function

Private Function SimpleFileName(ByVal FQFileName As String) As String
    Dim lngBack As Long
    
    lngBack = InStrRev(FQFileName, "\")
    SimpleFileName = Mid$(FQFileName, lngBack + 1)
End Function

