VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLyrdb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#######################################
'#         clsLYRDB:                   #
'# VB Wrapper for LYRDB's Lyrics       #
'#       Database Service              #
'#                                     #
'#     class and service by            #
'#    Flavio González Vázquez          #
'#          flavio@ya.com              #
'#                                     #
'#    == Function credits ==           #
'#    --> StringCount: Edward Moth     #
'#    --> CrLfTrim (modified):         #
'#                     Dane Jones      #
'#                                     #
'#    - Update #1: Support for         #
'#                 lyrics submissions  #
'#                                     #
'#######################################

'-- LYRDB Server Checking
Private Const FLAG_ICC_FORCE_CONNECTION = &H1
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

'-- URL Constants
Private Const LYRDB_HELLO = "http://www.lyrdb.com/hello.php"
Private Const LYRDB_LOOKUP = "http://www.lyrdb.com/lookup.php"
Private Const LYRDB_GETLYR = "http://www.lyrdb.com/getlyr.php"
Private Const LYRDB_UPLOADLYR_PATH = "/uploadlyr.php"

'-- Internet HTTP Connections
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
   (ByVal hInternetSession As Long, _
    ByVal lpszServerName As String, _
    ByVal nProxyPort As Integer, _
    ByVal lpszUsername As String, _
    ByVal lpszPassword As String, _
    ByVal dwService As Long, _
    ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Long
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
   (ByVal hInternetSession As Long, _
    ByVal lpszVerb As String, _
    ByVal lpszObjectName As String, _
    ByVal lpszVersion As String, _
    ByVal lpszReferer As String, _
    ByVal lpszAcceptTypes As Long, _
    ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Long
Private Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" _
   (ByVal hHttpRequest As Long, _
   ByVal sHeaders As String, _
   ByVal lHeadersLength As Long, _
   ByVal lModifiers As Long) As Integer
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" _
   (ByVal hHttpRequest As Long, _
    ByVal sHeaders As String, _
    ByVal lHeadersLength As Long, _
    ByVal sOptional As String, _
    ByVal lOptionalLength As Long) As Boolean


Const scUserAgent = "clsLyrdb/1.0"
Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_OPEN_TYPE_PROXY = 3
Const INTERNET_FLAG_RELOAD = &H80000000

Const LYRDB_SUCCESS = "success"

'-- Last lookup data
Private Type LyrdbItem
    sID As String
    sTrackname As String
    sArtist As String
End Type

Private lItem() As LyrdbItem

Private mProxy As String
Private mUseProxy As Boolean
Private mUseDefault As Boolean
Private mLastError As String
Private mAgent As String

Public Property Get LastError() As String
    LastError = mLastError
End Property

Public Property Get ServiceAvailable() As Boolean
    If InternetCheckConnection(LYRDB_HELLO, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
        ServiceAvailable = False
    Else
        ServiceAvailable = True
    End If
End Property

Public Property Get LyrdbHello() As String
    Dim buff As String
    GetHTTPFile LYRDB_HELLO, buff
    LyrdbHello = Trim(buff)
End Property

Public Function LyrdbLookup(sField As String, sSearchStr As String) As Boolean
    Dim buff As String, results() As String, n As Integer, sRet() As String, lCrlf As Long
    sSearchStr = Replace(sSearchStr, "'", "`")
    GetHTTPFile LYRDB_LOOKUP & "?q=" & sSearchStr & "&for=" & sField & "&agent=" & mAgent, buff
    
    lCrlf = InStr(buff, vbCrLf)
    If lCrlf > 0 Then
        buff = Left(buff, lCrlf)
    End If
    buff = Trim(buff)
    
    If Left(buff, 6) = "error:" Then
        LyrdbLookup = False
        mLastError = buff
    Else
    
        results = Split(buff, vbLf)
        
        If UBound(results) = -1 Then
            Erase lItem
        Else
        
            ReDim lItem(0 To UBound(results))
            
            For n = 0 To UBound(results)
                sRet = Split(results(n), "\")
                lItem(n).sID = sRet(0)
                If UBound(sRet) >= 1 Then lItem(n).sTrackname = sRet(1)
                If UBound(sRet) >= 2 Then lItem(n).sArtist = sRet(2)
            Next
            
        End If
        
        LyrdbLookup = True
    End If
    
End Function

Public Function LyrdbItem(sID As String) As String
    Dim buff As String, iLflf As Long
    GetHTTPFile LYRDB_GETLYR & "?q=" & sID, buff
    iLflf = InStr(buff, vbLf & vbLf)
    If iLflf > 0 Then
        buff = Left(buff, iLflf)
    End If
    
    buff = Replace(buff, vbCrLf, vbLf)
    buff = Replace(buff, vbCr, vbLf)
    buff = Replace(buff, vbLf, vbCrLf)
    
    buff = CrLfTrim(buff)
    
    If StringCount(buff, vbCrLf) = 0 Then
        LyrdbItem = ""
    Else
        LyrdbItem = buff
    End If
End Function

Private Function StringCount(ByVal Expression As String, _
                            item As String) As Integer
                        
  Dim lPosition As Integer
  Dim lCount As Integer

  Do
    lPosition = InStr(lPosition + 1, Expression, item)
    If lPosition > 0 Then
      lCount = lCount + 1
    End If
  Loop While lPosition > 0
  StringCount = lCount

End Function

Private Function CrLfTrim(ByVal StringToTrim As String) As String

    
    Dim chrToCheck As String
    Dim I As Byte

    For I = 1 To 2
        chrToCheck = Mid(StringToTrim, 1, 1)

        Do While chrToCheck = vbCr Or chrToCheck = vbLf Or chrToCheck = vbTab Or chrToCheck = " "
            StringToTrim = Mid(StringToTrim, 2)
            chrToCheck = Mid(StringToTrim, 1, 1)
        Loop
        StringToTrim = StrReverse(StringToTrim)
    Next I
    
    CrLfTrim = StringToTrim
End Function


Public Property Get Count() As Integer
    On Error Resume Next
    Count = UBound(lItem) + 1
    If Err Then Count = 0
End Property

Public Property Get LyrdbID(Index As Integer) As String
    LyrdbID = lItem(Index).sID
End Property
Public Property Get LyrdbTrackname(Index As Integer) As String
    LyrdbTrackname = lItem(Index).sTrackname
End Property
Public Property Get LyrdbArtist(Index As Integer) As String
    LyrdbArtist = lItem(Index).sArtist
End Property

Public Property Let Proxy(sProxy As String)
    mProxy = sProxy
End Property
Public Property Get Proxy() As String
    Proxy = mProxy
End Property

Public Property Let UseProxy(sUseProxy As Boolean)
    mUseProxy = sUseProxy
End Property
Public Property Get UseProxy() As Boolean
    UseProxy = mUseProxy
End Property

Public Property Let UseDefault(sUseDefault As Boolean)
    mUseDefault = sUseDefault
End Property
Public Property Get UseDefault() As Boolean
    UseDefault = mUseDefault
End Property

Public Property Let Agent(sAgent As String)
    mAgent = sAgent
End Property
Public Property Get Agent() As String
    Agent = mAgent
End Property


Function GetHTTPFile(ByVal URL As String, ByRef sBuffer As String, Optional ByVal sHeaders As String = vbNullString) As Boolean
    Dim hInternetSession As Long
    Dim hUrl As Long
    Dim FileNum As Integer
    Dim ok As Boolean
    Dim NumberOfBytesRead As Long
    Dim Buffer As String
    Dim fileIsOpen As Boolean

    On Error GoTo ErrorHandler

    If Len(URL) = 0 Then
        GetHTTPFile = False
        Exit Function
    End If

    hInternetSession = InternetOpen(scUserAgent, IIf(mUseDefault = True, 0, IIf(mUseProxy = False, INTERNET_OPEN_TYPE_DIRECT, INTERNET_OPEN_TYPE_PROXY)), _
        IIf(mUseProxy = True, mProxy, vbNullString), vbNullString, 0)
    If hInternetSession = 0 Then
        GetHTTPFile = False
        Exit Function
    End If

    hUrl = InternetOpenUrl(hInternetSession, URL, sHeaders, 0, _
        INTERNET_FLAG_EXISTING_CONNECT, 0)
    If hUrl = 0 Then
        GetHTTPFile = False
        Exit Function
    End If


    On Error GoTo ErrorHandler
    
    Buffer = Space(4096)
    sBuffer = vbNullString
    
    Do
        ok = InternetReadFile(hUrl, Buffer, Len(Buffer), NumberOfBytesRead)

        If NumberOfBytesRead = 0 Or Not ok Then Exit Do

        sBuffer = sBuffer & Left$(Buffer, NumberOfBytesRead)
    Loop

ErrorHandler:

    If hUrl Then InternetCloseHandle hUrl
    If hInternetSession Then InternetCloseHandle hInternetSession

    If Err Then
        GetHTTPFile = False
    Else
        GetHTTPFile = True
    End If
End Function

Private Function GetHttpFileByPost(ByVal URL As String, ByVal PostData As String) As String

    Dim hInternetOpen As Long
    Dim hInternetConnect As Long
    Dim hHttpOpenRequest As Long
    Dim bRet As Boolean
    Dim strServer As String
    Dim intPort As Integer
    Dim strPath As String
    
    Const INTERNET_DEFAULT_HTTP_PORT = 80

    strServer = "www.lyrdb.com"
    strPath = URL
    intPort = 80

    hInternetOpen = 0
    hInternetConnect = 0
    hHttpOpenRequest = 0

    Const INTERNET_OPEN_TYPE_PRECONFIG = 0

    hInternetOpen = InternetOpen(scUserAgent, IIf(mUseDefault = True, 0, IIf(mUseProxy = False, INTERNET_OPEN_TYPE_DIRECT, INTERNET_OPEN_TYPE_PROXY)), _
        IIf(mUseProxy = True, mProxy, vbNullString), vbNullString, 0)
    
    If hInternetOpen <> 0 Then

       Const INTERNET_SERVICE_HTTP = 3

       hInternetConnect = InternetConnect(hInternetOpen, _
                          strServer, _
                          intPort, _
                          vbNullString, _
                          "HTTP/1.1", _
                          INTERNET_SERVICE_HTTP, _
                          0, _
                          0)
    
       If hInternetConnect <> 0 Then

         Const INTERNET_FLAG_RELOAD = &H80000000
         hHttpOpenRequest = HttpOpenRequest(hInternetConnect, _
                             "POST", _
                             strPath, _
                             "HTTP/1.0", _
                             vbNullString, _
                             0, _
                             INTERNET_FLAG_RELOAD, _
                             0)
    
          If hHttpOpenRequest <> 0 Then
             Dim sHeader As String
             Const HTTP_ADDREQ_FLAG_ADD = &H20000000
             Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000
             sHeader = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
             bRet = HttpAddRequestHeaders(hHttpOpenRequest, _
                sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE _
                Or HTTP_ADDREQ_FLAG_ADD)
    
             Dim lPostDataLen As Long
    
             lPostDataLen = Len(PostData)
             bRet = HttpSendRequest(hHttpOpenRequest, _
                    vbNullString, _
                    0, _
                    PostData, _
                    lPostDataLen)
    
             Dim bDoLoop             As Boolean
             Dim sReadBuffer         As String * 2048
             Dim lNumberOfBytesRead  As Long
             Dim sBuffer             As String
             bDoLoop = True
             While bDoLoop
                sReadBuffer = vbNullString
                bDoLoop = InternetReadFile(hHttpOpenRequest, _
                   sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
                sBuffer = sBuffer & _
                     Left(sReadBuffer, lNumberOfBytesRead)
                If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
                DoEvents
             Wend
             GetHttpFileByPost = sBuffer
             bRet = InternetCloseHandle(hHttpOpenRequest)
          End If
          bRet = InternetCloseHandle(hInternetConnect)
       End If
       bRet = InternetCloseHandle(hInternetOpen)
    End If

End Function

Private Sub Class_Initialize()
    mAgent = vbNullString
End Sub

Public Function UploadLyrics(sAgent As String, _
                             sArtist As String, _
                             sTrackname As String, _
                             sLyrics As String, _
                             Optional sLanguage As String = "--", _
                             Optional sYourName As String = vbNullString, _
                             Optional sYourEmail As String = vbNullString, _
                             Optional bShowEmail As Boolean = False _
                             ) As Boolean
                             
    Dim ret As String, sHeaders As String
    
    If sAgent = vbNullString Or _
        sArtist = vbNullString Or _
        sTrackname = vbNullString Or _
        sLyrics = vbNullString Then
        
        UploadLyrics = False
    Else
    
        sHeaders = "name=" & UrlEncode(sYourName) & _
                   "&email=" & UrlEncode(sYourEmail) & _
                   "&allowemail=" & IIf(bShowEmail = True, "y", "n") & _
                   "&artist=" & UrlEncode(sArtist) & _
                   "&trackname=" & UrlEncode(sTrackname) & _
                   "&source=0&language=" & sLanguage & _
                   "&lyrics=" & UrlEncode(sLyrics) & _
                   "&agent=" & UrlEncode(sAgent)
        
        
        ret = GetHttpFileByPost(LYRDB_UPLOADLYR_PATH, sHeaders)
        If Trim(ret) <> LYRDB_SUCCESS Then
            UploadLyrics = False
        Else
            UploadLyrics = True
        End If

        
    End If
        
End Function

Private Function UrlEncode(strBefore As String) As String
    Dim strAfter As String
    Dim intLoop As Integer
    
    strBefore = Replace(strBefore, "'", "’")

    If Len(strBefore) > 0 Then
        For intLoop = 1 To Len(strBefore)
            Select Case Asc(Mid(strBefore, intLoop, 1))
                Case 48 To 57, 65 To 90, 97 To 122, 46, 45, 95, 42  '0-9, A-Z, a-z . - _ *
                strAfter = strAfter & Mid(strBefore, intLoop, 1)
                Case 32
                strAfter = strAfter & "+"
                Case Else
                strAfter = strAfter & "%" & Right("0" & Hex(Asc(Mid(strBefore, intLoop, 1))), 2)
            End Select
        Next
    End If

    UrlEncode = strAfter
End Function
