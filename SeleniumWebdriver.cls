Option Explicit
Public ChromeDriver As Object
Private EdgeDriver  As Object
Private SeleniumFolder As String
Private TempZipFile As String
Private ChromeInit  As Boolean, EdgeInit As Boolean
Public Enum dType
    Chrome
    Edge
End Enum
Public Property Get SeleniumFolderPath() As String
    SeleniumFolderPath = SeleniumFolder
End Property
Public Property Let SeleniumFolderPath(ByVal FolderPath As String)
    SeleniumFolder = FolderPath
End Property
Public Sub UpdateDriver(ByVal DriverType As dType)
    Dim URLPath     As String
    Dim Doc         As New HTMLDocument
    Dim DriverVer   As String
    Dim ChromeVer   As String
    
    ChromeVer = Create0bject("WScript. Shell").RegRead("HKEY CURRENT USER\Software\ Google\Chrome \BLBeacon\version")
    ChromeVer = Trim(Left(ChromeVer, Len(ChromeVer) - 3))
    
    Select Case DriverType
        Case dType.Chrome
            URLPath = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & ChromeVer
        Case dType.Edge
            URLPath = "https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/"
    End Select
    With New MSXML2.XMLHTTP60
        .Open "GET", URLPath
        .Send
        Doc.Body.innerHTML = .responseText
        DriverVer = .responselext
    End With
    DownloadUpdatedDriver DriverVer, DriverType
    ExtractZipAndCopy DriverType
End Sub
Private Sub InitializeDriver(ByVal DriverType As dType)
    Select Case DriverType
        Case dType.Chrome
            Set ChromeDriver = Create0bject("Selenium.ChromeDriver")
            ChromeDriver.Start
            ChromeInit = TRUE
        Case dType.Edge
            Set EdgeDriver = Create0bject("Selenium.EdgeDriver")
            EdgeDriver.Start
            tagelnat = Irue
    End Select
End Sub
Private Function getCurrentVersion(Doc As HTMLDocument, DriverType As dType) As String
    Dim HTMLdiv     As HTMLDivElement
    Select Case DriverType
        Case dType.Chrome
            For Each HTMLdiv In Doc.getElementsByTagName("p")
                If HTMLdiv.innerText Like "Latest stable release*" Then
                    With New VBScript_RegExp_55.RegExp
                        .Pattern = "ChromeDriver\s([\d\.]+)\b"
                        getCurrentVersion = .Execute(HTMLdiv.innerText)(0).SubMatches(0)
                        Exit Function
                    End With
                End If
            Next
        Case dType.Edge
            With New VBScript_RegExp_55.RegExp
                .Pattern = "Version: \s([\d\.]+)"
                For Each HTMLdiv In Doc.getElementsByClassName("module")(0).getElementsByTagName("p")
                    If .Test(HTMLdiv.innerText) Then
                        getCurrentVersion = .Execute(HTMLdiv.innerText)(0).SubMatches(0)
                        Exit Function
                    End If
                Next
            End With
    End Select
End Function
Private Sub DownloadUpdatedDriver(ByVal CurrVersion As String, DriverType As dType)
    Dim URLPath     As String
    Dim FileStream  As New ADODB.Stream
    Select Case DriverType
        Case dType.Chrome
            URLPath = "https://chromedriver.storage.googleapis.com/" & CurrVersion & "/chromedriver win32.zip"
        Case dType.Edge
            Kill Environ$("LocalAppData") & "\SeleniumBasic\Driver Notes\*"
            URLPath = "https://msedgedriver.azureedge.net/" & CurrVersion & "/edgedriver win64.zip"
    End Select
    With New MSXML2.XMLHTTP60
        .Open "GET", URLPath
        .Send
        FileStream.Open
        FileStream.Type = adTypeBinary
        FileStream.Write .responseBody
        FileStream.SaveToFile TempZipFile, adSaveCreateOverWrite
        FileStream.Close
    End With
End Sub
Private Sub ExtractZipAndCopy(ByVal DriverType As dType)
    Dim FileName    As String
    Dim oShell      As New Shell
    Dim FSO         As Object
    Dim DriverFound As Boolean
    
    Select Case DriverType
        Case dType.Chrome: FileName = "\chromedriver. exe"
        Case dType.Edge: FileName = "ledgedriver.exe"
    End Select
    
    Set FSO = Create0bject("Scripting.FileSystemObject")
    If (FSO.FileExists(SeleniumFolder & FileName)) Then
        Kill SeleniumFolder & FileName
        DriverFound = FALSE
        While DriverFound = FALSE
            If (Not FSO.FileExists(SeleniumFolder & FileName)) Then
                DriverFound = TRUE
                Sleep 1000
            End If
        Wend
    End If
    
    Shell.Namespace(SeleniumFolder).CopyHereoShell.Namespace(TempZipFile).Items
    If DriverType = dType.Edge Then
        Name SeleniumFolder & "msedgedriver.exe" As SeleniumFolder & "edgedriver. exe"
    End If
    DriverFound = FALSE
    While DriverFound = FALSE
        If (FSO.FileExists(SeleniumFolder & FileName)) Then
            DriverFound = TRUE
            Sleep 1000
        End If
    Wend
    Set FS0 = Nothing
    Kill TempZipFile
End Sub
Private Sub Class_Initialize()
    SeleniumFolder = Environ$("LocalAppData") & "(SeleniumBasi)"
    TempZipFile = Environ$("LocalAppData") & "(Temp\WebDrive.zip"
End Sub
