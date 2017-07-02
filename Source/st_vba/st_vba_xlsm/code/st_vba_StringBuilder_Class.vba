'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    StringBuilder Class
'ObjectName:    st_vba_StringBuilder
'--------------------------------------------------
'Version:       2017/03/19
'--------------------------------------------------
'   ・  文字列を高速化して扱うためのモジュール
'   ・  .NET の StringBuilder とメソッドの互換性があるわけではないが
'       結果同じことができるものなので、名前を借りている
'--------------------------------------------------
Option Explicit

Const BufferBlockSize As Long = 1000000

Private Buffer As String
Public LengthIndex As Long

Sub Initialize()
    LengthIndex = 0
    Buffer = VBA.String(BufferBlockSize, " ")
End Sub

Sub Clear()
    Call Initialize
End Sub

Sub Add(Value As String)
    Dim NewLengthIndex As Long
    NewLengthIndex = LengthIndex + Len(Value)
    Do While (Len(Buffer) < NewLengthIndex)
        Buffer = Buffer + VBA.String(BufferBlockSize, " ")
    Loop

    Mid(Buffer, LengthIndex + 1) = Value
    LengthIndex = LengthIndex + Len(Value)
End Sub

Public Property Get Text() As String
    Text = Left(Buffer, LengthIndex)
End Property

'----------
'使い方と､動作確認コード
'Option Explicit
'
'    Sub 文字列連結でMidステートメントとの比較()
'
'        Dim Str1 As String
'        Dim Str2 As String
'        Dim Str3 As String
'        Dim lng As Long
'        Dim strTime As String
'        Dim dblTimer As Double
'
'        Dim LoopCount As Long
'        LoopCount = 200000
'
'        '―――――――――――――――
'        '通常の連結
'        '―――――――――――――――
'    ' dblTimer = Time
'    '
'    ' Str1 = ""
'    ' For lng = 1 To LoopCount
'    ' Str1 = Str1 & "あいうえおかきくけこ"
'    ' Next lng
'    '
'    ' strTime = "通常：" & Time - dblTimer
'
'        '―――――――――――――――
'        'Midステートメント使用
'        '―――――――――――――――
'
'        dblTimer = Time
'
'        'あらかじめ必要な文字数分を確保
'        Str2 = VBA.String(LoopCount * 10, "a")
'        Dim lngPos As Long
'        lngPos = 1 'ループ内で開始する文字位置
'        For lng = 1 To LoopCount
'            Mid(Str2, lngPos, 10) = "あいうえおかきくけこ"
'            lngPos = lngPos + 10
'        Next lng
'        Str2 = Left(Str2, lngPos - 1)
'
'        strTime = strTime & vbCrLf & "Mid：" & Time - dblTimer
'
'        '―――――――――――――――
'        'StringBuilder
'        '―――――――――――――――
'
'        dblTimer = Time
'
'        Dim StrBuilder As StringBuilder
'        Set StrBuilder = New StringBuilder
'
'        For lng = 1 To LoopCount
'            Call StrBuilder.Add("あいうえおかきくけこ")
'        Next
'        Str3 = StrBuilder.Text
'
'        strTime = strTime & vbCrLf & "Mid：" & Time - dblTimer
'
'        '―――――――――――――――
'        '比較
'        '―――――――――――――――
'
'        If (Str1 <> "") And (Str1 <> Str2) Then
'            MsgBox "Str1<>Str2"
'        End If
'
'        If Str2 <> Str3 Then
'            MsgBox "Str1<>Str2"
'        End If
'
'
'        '時間発表
'        MsgBox strTime
'
'    End Sub

