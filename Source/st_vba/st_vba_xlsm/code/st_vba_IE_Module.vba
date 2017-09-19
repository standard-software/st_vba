'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    IE Module
'ObjectName:    st_vba_IE
'--------------------------------------------------
'Version:       2017/09/19
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
Optional ByRef NavigateCancelFlag As Boolean = False, _
Optional ByVal RefreshSecond As Long = 30, _
Optional ByVal TimeOutSecond As Long = 60)
    Call ie.Navigate(URL)
    Call IE_NavigateWait(ie, _
        RefreshSecond, TimeOutSecond, NavigateCancelFlag)
End Sub

'----------------------------------------
'・IE起動とアドレス表示
'----------------------------------------
'   ・  Basic認証として
'       AuthBasicに
'       user:password
'       この文字列をBASE64エンコードしたキーを入れると
'       認証を通過する
'----------------------------------------
Sub IE_Navigate_AuthBasic(ByVal ie As InternetExplorer, _
ByVal URL As String, _
Optional ByRef NavigateCancelFlag As Boolean = False, _
Optional ByVal AuthBasic As String, _
Optional ByVal RefreshSecond As Long = 30, _
Optional ByVal TimeOutSecond As Long = 60)
    If AuthBasic = "" Then
        Call ie.Navigate(URL)
    Else
        AuthBasic = _
            "Authorization: Basic " & AuthBasic & vbCrLf
        Call ie.Navigate(URL, , , , AuthBasic)
    End If
    Call IE_NavigateWait(ie, _
        RefreshSecond, TimeOutSecond, NavigateCancelFlag)
End Sub

'----------------------------------------
'・IE起動とアドレス表示
'----------------------------------------
'   ・  Basic認証として
'       AuthBasicInputにSendKeysのコードを入れる
'       ID, "{TAB}", PASSWORD, "{ENTER}"
'       あるいは
'       "+({TAB})", ID, "{TAB}", PASSWORD, "{ENTER}"
'----------------------------------------
Sub IE_Navigate_AuthBasicInput(ByVal ie As InternetExplorer, _
ByVal URL As String, _
ByRef AuthBasicInput() As String, _
Optional ByRef NavigateCancelFlag As Boolean = False, _
Optional ByVal RefreshSecond As Long = 30, _
Optional ByVal TimeOutSecond As Long = 60, _
Optional ByVal InputBeforeMiliSecond As Long = 5000)
    Call ie.Navigate(URL)
    
    '5秒停止
    Call Sleep(InputBeforeMiliSecond)
    Dim I As Long
    For I = 0 To ArrayCount(AuthBasicInput) - 1
        Call SendKeys(AuthBasicInput(I))
    Next
    
    Call IE_NavigateWait(ie, RefreshSecond, TimeOutSecond, NavigateCancelFlag)
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
Optional ByVal RefreshSecond As Long = 30, _
Optional ByVal TimeOutSecond As Long = 60, _
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

'ClassNameWildCardに一致するが
'さらに除外条件をつけたい場合に
'NotClassNameWildCardを指定する
Public Function IE_GetElementByTagNameClassName(ByVal Element As Object, _
ByVal TagName As String, ByVal ClassNameWildCard As String, _
Optional ByVal NotClassNameWildCard As String = "") As Object
    Dim Result As Object: Set Result = Nothing
    Dim E1 As Object
    If NotClassNameWildCard = "" Then
        For Each E1 In Element.GetElementsByTagName(TagName)
            If E1.ClassName Like ClassNameWildCard Then
                Set Result = E1
                Exit For
            End If
        Next
    Else
        For Each E1 In Element.GetElementsByTagName(TagName)
            If E1.ClassName Like ClassNameWildCard Then
                If Not (E1.ClassName Like NotClassNameWildCard) Then
                    Set Result = E1
                    Exit For
                End If
            End If
        Next
    End If
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

'すべての条件指定を網羅した関数
'MatchCounterに
'0を設定すると見つかった要素の中の最終要素
'負の値を設定すると見つかった要素の中の最後から二番目
Public Function IE_GetElementByTagNameSearch( _
ByVal Element As Object, _
ByVal TagName As String, _
Optional ByVal IdWildCard As String = "", _
Optional ByVal NotIdWildCard As String = "", _
Optional ByVal NameWildCard As String = "", _
Optional ByVal NotNameWildCard As String = "", _
Optional ByVal ClassWildCard As String = "", _
Optional ByVal NotClassWildCard As String = "", _
Optional ByVal InnerHtmlWildCard As String = "", _
Optional ByVal NotInnerHtmlWildCard As String = "", _
Optional ByVal InnerTextWildCard As String = "", _
Optional ByVal NotInnerTextWildCard As String = "", _
Optional ByVal OuterHtmlWildCard As String = "", _
Optional ByVal NotOuterHtmlWildCard As String = "", _
Optional ByVal MatchCounter As Long = 1) As Object
    Dim Result As Object: Set Result = Nothing
    Dim Results() As Object
    Dim E1 As Object
    Dim Counter As Long
    Counter = 0
    For Each E1 In Element.GetElementsByTagName(TagName)
    Do
        If NotIdWildCard <> "" Then
            If E1.Id Like NotIdWildCard Then
                Exit Do
            End If
        End If
        If NotNameWildCard <> "" Then
            If E1.Name Like NotNameWildCard Then
                Exit Do
            End If
        End If
        If NotClassWildCard <> "" Then
            If E1.ClassName Like NotClassWildCard Then
                Exit Do
            End If
        End If
        If NotInnerHtmlWildCard <> "" Then
            If E1.InnerHTML Like NotInnerHtmlWildCard Then
                Exit Do
            End If
        End If
        If NotInnerTextWildCard <> "" Then
            If E1.InnerHTML Like NotInnerTextWildCard Then
                Exit Do
            End If
        End If
        If NotOuterHtmlWildCard <> "" Then
            If E1.OuterHTML Like NotOuterHtmlWildCard Then
                Exit Do
            End If
        End If
        
        Dim MatchFlag As Boolean    '条件に一致
        Dim UnMatchFlag As Boolean  '条件で否定
        MatchFlag = False
        UnMatchFlag = False
        
        If IdWildCard <> "" Then
            If E1.Id Like IdWildCard Then
                MatchFlag = True
            Else
                UnMatchFlag = True
            End If
        End If
        If NameWildCard <> "" Then
            If E1.Name Like NameWildCard Then
                MatchFlag = True
            Else
                UnMatchFlag = True
            End If
        End If
        If ClassWildCard <> "" Then
            Debug.Print E1.ClassName
            If E1.ClassName Like ClassWildCard Then
                MatchFlag = True
            Else
                UnMatchFlag = True
            End If
        End If
        If InnerHtmlWildCard <> "" Then
            If E1.InnerHTML Like InnerHtmlWildCard Then
                MatchFlag = True
            Else
                UnMatchFlag = True
            End If
        End If
        If InnerTextWildCard <> "" Then
            If E1.InnerText Like InnerTextWildCard Then
                MatchFlag = True
            Else
                UnMatchFlag = True
            End If
        End If
        If OuterHtmlWildCard <> "" Then
            If E1.OuterHTML Like OuterHtmlWildCard Then
                MatchFlag = True
            Else
                UnMatchFlag = True
            End If
        End If
        
        If MatchFlag And (UnMatchFlag = False) Then
            Set Result = E1
            Call ArrayAdd(Results, E1)
            
            '一致回数(Counter)と
            '指定回数(MatchCounter)が
            'イコールなら結果を返す。
            'MatchCounter=0の場合最終値を返す
            Counter = Counter + 1
            If MatchCounter = Counter Then
                Exit For
            End If
        End If
        
    Loop While False
    Next
    
    If 0 <= MatchCounter Then
        Set IE_GetElementByTagNameSearch = Result
    Else
        Set IE_GetElementByTagNameSearch = _
            Results(ArrayCount(Results) - 1 + MatchCounter)
    End If
End Function


