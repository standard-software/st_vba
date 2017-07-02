'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    IE Module
'ObjectName:    st_vba_IE
'--------------------------------------------------
'Version:       2017/03/19
'--------------------------------------------------
'   ・  IEコントロールするためのモジュール
'--------------------------------------------------
Option Explicit


'----------------------------------------
'◆InternetExplorer操作
'----------------------------------------

'----------------------------------------
'・新規IEオブジェクト生成
'----------------------------------------
Function IE_NewObject(Optional ByVal Visible As Boolean = True) As InternetExplorer
    Set IE_NewObject = New InternetExplorer
    IE_NewObject.Visible = Visible
End Function

Sub testIE_NewObject()
    Dim ie As InternetExplorer
    Set ie = IE_NewObject
    Call IE_Navigate(ie, "http://www.yahoo.co.jp/")
    Call MsgBox("finish")
    ie.Quit
End Sub

'----------------------------------------
'・既存の起動IEの取得
'----------------------------------------
'   ・  起動済みIEがあればそれを取得
'       なければ新規にIEを起動する
'----------------------------------------
Function IE_GetObject(Optional ByVal URL As String = "") As InternetExplorer
    Dim ie As InternetExplorer
    Set ie = Nothing

    Dim ShellApp As Object
    Dim Window As Object
    
    Set ShellApp = CreateObject("Shell.Application")
    
    For Each Window In ShellApp.Windows
       If IsIncludeStr(Window.Name, "Internet Explorer") Then
            If (URL = "") Then
                Set ie = Window
                Exit For
            Else
                If (IsIncludeStr(Window.LocationURL, URL)) Then
                    Set ie = Window
                    Exit For
                End If
            End If
        End If
    Next
    
    If ie Is Nothing Then
        Set ie = IE_NewObject(True)
    End If

    Set IE_GetObject = ie
End Function

Sub testIE_GetObject()
    Dim ie As InternetExplorer
    Set ie = IE_GetObject("https://www.google.co.jp/")
    Call IE_Navigate(ie, "http://www.yahoo.co.jp/")
    Call MsgBox("finish")
    Call IE_Quit(ie)
End Sub

'----------------------------------------
'・IE起動とアドレス表示
'----------------------------------------
Sub IE_Navigate(ByVal ie As InternetExplorer, _
ByVal URL As String, _
Optional ByRef NavigateCancelFlag As Boolean = False)
    Call ie.Navigate(URL)
    Call IE_NavigateWait(IE, 30, 60, NavigateCancelFlag)
End Sub

'----------------------------------------
'・IEリフレッシュ
'----------------------------------------
'   ・  リフレッシュ時に非表示が表示になってしまう場合が
'       あるようなので対応
'----------------------------------------
Sub IE_Refresh(ByVal ie As InternetExplorer)
    Dim VisibleBuffer As Boolean
    VisibleBuffer = ie.Visible
    Call ie.Refresh
    ie.Visible = VisibleBuffer
End Sub

'----------------------------------------
'・IEナビゲート待機
'----------------------------------------
Sub IE_NavigateWait(ByVal ie As InternetExplorer, _
ByVal RefreshSecond As Long, _
ByVal TimeOutSecond As Long, _
Optional ByRef NavigateCancelFlag As Boolean = False)

On Error GoTo Error

    'リフレッシュ時間/タイムアウト時間
    Dim RefreshTime As Date
    Dim TimeOutTime As Date
    RefreshTime = Now + TimeSerial(0, 0, RefreshSecond)
    TimeOutTime = Now + TimeSerial(0, 0, TimeOutSecond)

    Do Until (ie.Busy = False) And (ie.ReadyState = READYSTATE_COMPLETE)
        'READYSTATE_COMPLETE=4
        DoEvents
        Call Sleep(1)
        If NavigateCancelFlag Then
            Exit Sub
        End If
        If Now > TimeOutTime Then
            Exit Sub
        End If
        If Now > RefreshTime Then
            'ページの再読み込み(リフレッシュ)
            Call IE_Refresh(ie)
            RefreshTime = Now + TimeSerial(0, 0, RefreshSecond)
        End If
    Loop

    RefreshTime = Now + TimeSerial(0, 0, RefreshSecond)
    TimeOutTime = Now + TimeSerial(0, 0, TimeOutSecond)

    Do Until (ie.Document.ReadyState = "complete")
        DoEvents
        Call Sleep(1)
        If NavigateCancelFlag Then
            Exit Sub
        End If
        If Now > TimeOutTime Then
            Exit Sub
        End If
        If Now > RefreshTime Then
            Call IE_Refresh(ie)
            RefreshTime = Now + TimeSerial(0, 0, RefreshSecond)
        End If
    Loop

Error:

End Sub

Sub testIE_NavigateWait()

    Dim ie As InternetExplorer
    Set ie = IE_NewObject
    Call IE_Navigate(ie, "https://www.google.co.jp/")

    Set ie = IE_GetObject
    Call IE_Navigate(ie, "http://www.yahoo.co.jp/")
    
    Call MsgBox("finish")

    ie.Quit
End Sub

'----------------------------------------
'・IEを閉じる
'----------------------------------------
'   ・  IEオブジェクトが見つからない場合でも
'       エラーを起こさずに終わる
'----------------------------------------
Sub IE_Quit(ByVal ie As InternetExplorer)
On Error Resume Next
    ie.Quit
End Sub

'----------------------------------------
'・IEでJavaScriptを実行する
'----------------------------------------
Sub IE_RunJavaScript(ByVal ie As InternetExplorer, ByVal ScriptCode As String)
    Call ie.Navigate("JavaScript:" + ScriptCode)
End Sub

Sub testIE_RunJavaScript()
    Dim ie As InternetExplorer
    Set ie = IE_GetObject
    Call IE_Navigate(ie, "http://www.yahoo.co.jp/")
    
    '1ページ分スクロール
    Call IE_RunJavaScript(ie, "scrollTo(0," & ie.Document.body.ScrollHeight & ")")

    Call MsgBox("finish")
    Call IE_Quit(ie)
End Sub

'----------------------------------------
'・IEで条件に一致したエレメントを取得する
'----------------------------------------
'   ・  Elementには ie.Document を指定するとよい
'----------------------------------------
Public Function IE_GetElementByTagName(ByVal Element As Object, _
ByVal TagName As String) As Object
    Dim Result As Object: Set Result = Nothing
    Dim E1 As Object
    For Each E1 In Element.GetElementsByTagName(TagName)
        Set Result = E1
        Exit For
    Next
    Set IE_GetElementByTagName = Result
End Function

Public Function IE_GetElementByTagNameClassName(ByVal Element As Object, _
ByVal TagName As String, ByVal ClassNameWildCard As String) As Object
    Dim Result As Object: Set Result = Nothing
    Dim E1 As Object
    For Each E1 In Element.GetElementsByTagName(TagName)
        If E1.ClassName Like ClassNameWildCard Then
            Set Result = E1
            Exit For
        End If
    Next
    Set IE_GetElementByTagNameClassName = Result
End Function

Public Function IE_GetElementByTagNameId(ByVal Element As Object, _
ByVal TagName As String, ByVal IdWildCard As String) As Object
    Dim Result As Object: Set Result = Nothing
    Dim E1 As Object
    For Each E1 In Element.GetElementsByTagName(TagName)
        If E1.ID Like IdWildCard Then
            Set Result = E1
            Exit For
        End If
    Next
    Set IE_GetElementByTagNameId = Result
End Function

Public Function IE_GetElementByTagNameName(ByVal Element As Object, _
ByVal TagName As String, ByVal NameWildCard As String) As Object
    Dim Result As Object: Set Result = Nothing
    Dim E1 As Object
    For Each E1 In Element.GetElementsByTagName(TagName)
        If E1.Name Like NameWildCard Then
            Set Result = E1
            Exit For
        End If
    Next
    Set IE_GetElementByTagNameName = Result
End Function

Public Function IE_GetElementByTagNameInnerHTML(ByVal Element As Object, _
ByVal TagName As String, ByVal InnerHTMLWildCard As String) As Object
    Dim Result As Object: Set Result = Nothing
    Dim E1 As Object
    For Each E1 In Element.GetElementsByTagName(TagName)
        If E1.InnerHTML Like InnerHTMLWildCard Then
            Set Result = E1
            Exit For
        End If
    Next
    Set IE_GetElementByTagNameInnerHTML = Result
End Function


