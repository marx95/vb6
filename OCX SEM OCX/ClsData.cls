VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ClsData
'=======
'
'This class is used by LibData to store class attributes
'for the createable classes found in ActiveX libraries
'referenced in the subject VB 6.0 project file.
'
'Usage:
'
'   Instances should only be created by LibData.  Members
'   of this class are meant to be read-only by programs
'   using LibData and ClsData.
'
'Dependencies:
'
'   none
'

'==================== Private data =======================

Private m_CLSID As String
Private m_Name As String

'===================== Properties ========================

Friend Property Let CLSID(ByVal CLSIDString As String)
    m_CLSID = CLSIDString
End Property

Public Property Get CLSID() As String
    CLSID = m_CLSID
End Property

Friend Property Let Name(ByVal NameString As String)
    m_Name = NameString
End Property

Public Property Get Name() As String
    Name = m_Name
End Property
