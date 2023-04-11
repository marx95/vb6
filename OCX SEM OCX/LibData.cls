VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'LibData
'=======
'
'This class is used to parse, process, and store Reference
'and Object strings from a VB 6.0 project file.
'
'Usage:
'
'   Create and instance and call the LoadObject() or
'   LoadReference() method.  Then the properties of this
'   class can be interrogated to retrieve the attributes of
'   the subject ActiveX library.
'
'Dependencies:
'
'   o Uses the VB Type Library Information DLL TLBINF32.DLL
'
'   o The internal class ClsData.
'

'==================== Private data =======================

Private m_ClassData() As ClsData
Private m_ClassLast As Long
Private m_Description As String
Private m_FileLocation As String
Private m_LCID As Long
Private m_LIBID As String
Private m_MajorVersion As Integer
Private m_MinorVersion As Integer
Private m_Name As String

'===================== Properties ========================

Public Property Get Class(ByVal Index As Long) As ClsData
    If 0 <= Index And Index <= UBound(m_ClassData) Then
        Set Class = m_ClassData(Index)
    Else
        Err.Raise &H80045A01, "Class index out of range", "LibData"
    End If
End Property

Public Property Get Count() As Long
    Count = m_ClassLast + 1
End Property

Public Property Get FileLocation() As String
    FileLocation = m_FileLocation
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Get LIBID() As String
    LIBID = m_LIBID
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Version() As String
    Version = CStr(m_MajorVersion) & "." & CStr(m_MinorVersion)
End Property

'====================== Methods ==========================

Public Sub LoadObject(ByVal ObjectString As String)
    'String format:
    '
    '   <LIBID>#<Maj>.<Min>#<LCID>; <OCXName>
    
    Dim strParts() As String
    Dim strVersionParts() As String
    
    strVersionParts = Split(ObjectString, ";")
    strParts = Split(strVersionParts(0), "#")
    m_LIBID = strParts(0)
    strVersionParts = Split(strParts(1), ".")
    m_MajorVersion = CInt(strVersionParts(0))
    m_MinorVersion = CInt(strVersionParts(1))
    m_LCID = CLng(strParts(2))
    LookupRegistration
End Sub

Public Sub LoadReference(ByVal ReferenceString As String)
    'String format:
    '
    '   *\G<LIBID>#<Maj>.<Min>#<LCID>#<RelFileLoc>#<Desc>
    
    Dim strParts() As String
    Dim strVersionParts() As String
    
    strParts = Split(Mid$(ReferenceString, 4), "#")
    m_LIBID = strParts(0)
    strVersionParts = Split(strParts(1), ".")
    m_MajorVersion = CInt(strVersionParts(0))
    m_MinorVersion = CInt(strVersionParts(1))
    m_LCID = CLng(strParts(2))
    LookupRegistration
End Sub

'================== Misc procedures ======================

Private Sub LookupRegistration()
    Dim tliLib As TLI.TypeLibInfo
    Dim tiClass As TLI.TypeInfo
    Dim cdClass As ClsData
    
    Set tliLib = TypeLibInfoFromRegistry(m_LIBID, _
                                         m_MajorVersion, _
                                         m_MinorVersion, _
                                         m_LCID)
    With tliLib
        m_Name = .Name
        m_Description = .HelpString
        m_FileLocation = .ContainingFile
        
        m_ClassLast = -1
        For Each tiClass In .CoClasses
            Set cdClass = New ClsData
            With cdClass
                .CLSID = tiClass.Guid
                .Name = tiClass.Name
            End With
            
            m_ClassLast = m_ClassLast + 1
            ReDim Preserve m_ClassData(m_ClassLast)
            Set m_ClassData(m_ClassLast) = cdClass
        Next
    End With
    
    Set tliLib = Nothing
End Sub


