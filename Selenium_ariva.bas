Attribute VB_Name = "Selenium_ariva"
' you are free to use/modify/sell this table as you wish

Option Explicit

Const consentUUID = "985a0cc2-9534-4bbc-97a3-e265d6ec630d_2"

Private SeleniumDriver_ariva As Variant
Private seleniumStarted As Boolean

Dim keys As New Selenium.keys
Public cookiesAccepted As Boolean

Function GetAriva_Fund(url$, wkn$)
    Dim tmpWebElements As WebElements
    Dim tmpWebElement As WebElement
    Dim tmpString As String
    Dim tmpStrings() As String
    Dim i%
    
    If IsEmpty(SeleniumDriver_ariva) Then
        ' create new selenium driver
        Set SeleniumDriver_ariva = CreateObject2("Selenium.ChromeDriver")
        SeleniumDriver_ariva.SetPreference "download.default_directory", Environ$("USERPROFILE") & "\Downloads\"
        
        seleniumStarted = True
    End If
    
    ' if "url" is empty, then search for fund according to WKN and update url
    Call SeleniumDriver_ariva.Get(url)
    Call Sleep(3000)

    AcceptCookies
    
    ' //div[@id='pageHistoricQuotes']/div[6]/div[4]/form/ul/li[4]/input
    Set tmpWebElements = _
            SeleniumDriver_ariva.FindElementsByXPath("//div[@id='pageHistoricQuotes']/div[6]/div[4]/form/ul/li[4]/input")
    
    Call tmpWebElements(1).SendKeys(" ")
    ' tmpWebElements(1).Click
End Function

Sub CloseSeleniumDriver()
    If Not (SeleniumDriver_ariva Is Nothing) Then
        SeleniumDriver_ariva.Close
        SeleniumDriver_ariva.Quit
    End If
End Sub

Sub AcceptCookies()
    Dim tmpWebElement As WebElement
    Dim cookies
    
    If Not cookiesAccepted Then
        Call SeleniumDriver_ariva.Manage.AddCookie("consentUUID", _
            consentUUID, ".ariva.de")
        
        Call SeleniumDriver_ariva.Get(SeleniumDriver_ariva.url)
        
        cookiesAccepted = True
    End If
End Sub

Function CreateObject2(typeName As String) As Object
    Static domain As mscorlib.AppDomain
    If domain Is Nothing Then
        Dim host As New mscoree.CorRuntimeHost
        host.Start
        host.GetDefaultDomain domain
    End If
    Set CreateObject2 = domain.CreateInstanceFrom(Environ("USERPROFILE") & "\AppData\Local\SeleniumBasic\Selenium.dll", typeName).Unwrap
End Function
