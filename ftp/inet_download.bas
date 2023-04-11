Attribute VB_Name = "InternetModule"
Const scUserAgent = "API-Guide test program"
Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_OPEN_TYPE_PROXY = 3
Const INTERNET_FLAG_RELOAD = &H80000000

Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Public Function PegarIP(sLink As String) As String
    Dim hOpen As Long
    Dim hFile As Long
    Dim sBuffer As String
    Dim ret As Long
    
    sBuffer = Space(16) 'Create a buffer for the file we're going to download
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0) 'Create an internet connection
    hFile = InternetOpenUrl(hOpen, sLink, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&) 'Open the url
    InternetReadFile hFile, sBuffer, 16, ret 'Read the first 1000 bytes of the file
    InternetCloseHandle hFile 'clean up
    InternetCloseHandle hOpen 'clean up
    sBuffer = Replace(sBuffer, " ", vbNullString)
    PegarIP = sBuffer
End Function

