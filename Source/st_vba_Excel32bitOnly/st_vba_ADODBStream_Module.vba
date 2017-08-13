'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    ADODBStream Module
'ObjectName:    st_vba_ADODBStream
'--------------------------------------------------
'Version:       2017/07/02
'--------------------------------------------------
'   ・  Win7 64bit版 Excel2016 64bit版で
'       ADODB.Streamが32bit版にリンクしてしまい
'       正しく動作しない場合があったために分離
'   ・  参照設定を次のように設定しても
'           Microsoft AxtiveX Data Objects 6.1 Library
'           C:\Program Files\Common Files\System\ado\msado15.dll
'       次のように勝手に修正されて動作不具合を起こしていた
'           C:\Program Files (x86)\Common Files\System\ado\msado15.dll
'--------------------------------------------------
Option Explicit

'----------------------------------------
'◆テキストファイル読み書き
'----------------------------------------

''----------------------------------------
''・エンコードタイプの指定
''----------------------------------------
'
'Public Const TextEncodingTypeEnum_Shift_JIS        As String = "SHIFT_JIS"
'Public Const TextEncodingTypeEnum_UTF8_BOM         As String = "UTF-8"
'Public Const TextEncodingTypeEnum_UTF8_BOM_NO      As String = "UTF-8N"
'Public Const TextEncodingTypeEnum_UTF16_LE_BOM     As String = "UNICODEFFFE"
'Public Const TextEncodingTypeEnum_UTF16_LE_BOM_NO  As String = "UTF-16LE"
'Public Const TextEncodingTypeEnum_UTF16_BE_BOM     As String = "UNICODEFEFF"
'Public Const TextEncodingTypeEnum_UTF16_BE_BOM_NO  As String = "UTF-16BE"
'Public Const TextEncodingTypeEnum_ASCII            As String = "ASCII"
'Public Const TextEncodingTypeEnum_JIS              As String = "ISO-2022-JP"
'Public Const TextEncodingTypeEnum_EUC_JP           As String = "EUC-JP"
'Public Const TextEncodingTypeEnum_UTF_7            As String = "UTF-7"
'
'' UTF16_LE_BOMの指定は、
'' [UNICODEFFFE]だけではなく
'' [UNICODE]や[UTF-16]も同じ動作になるが
'' UTF16_BE_BOMとの対比としてわかりやすいので
'' [UNICODEFFFE]を採用する
'
'
'Public Function CheckEncodeName(EncodingName As String) As Boolean
'   CheckEncodeName = OrValue(UCase$(EncodingName), _
'       "SHIFT_JIS", _
'       "UNICODE", "UNICODEFFFE", "UTF-16", _
'       "UTF-16LE", _
'       "UNICODEFEFF", _
'       "UTF-16BE", _
'       "UTF-8", _
'       "UTF-8N", _
'        "ASCII", _
'       "ISO-2022-JP", _
'       "EUC-JP", _
'       "UTF-7")
'End Function
'
''----------------------------------------
''・テキストファイル読込
''----------------------------------------
''   ・  Charsetに指定する文字が
''       UNICODEFFFE/UNICODE/UTF-16/UTF-16LE の場合は
''       UTF-16LEのBOMの有無にかかわらず読み込むことができる
''   ・  UTF-8のBOM無しの指定文字はないので
''       UTF-8を指定するとBOMの有無にかかわらず読み込むことができる
''----------------------------------------
''Const StreamTypeEnum_adTypeBinary = 1
''Const StreamTypeEnum_adTypeText = 2
''Const StreamReadEnum_adReadAll = -1
''Const StreamReadEnum_adReadLine = -2
''Const SaveOptionsEnum_adSaveCreateOverWrite = 2
'
'Public Function LoadTextFile( _
'ByVal FilePath As String, ByVal EncodingName As String) As String
'    Call Assert(CheckEncodeName(EncodingName), "Error:LoadTextFile")
'
'    Dim Stream As ADODB.Stream
'    Set Stream = New ADODB.Stream
'    Stream.Type = adTypeText
'    Select Case UCase(EncodingName)
'    Case TextEncodingTypeEnum_UTF8_BOM_NO
'        Stream.Charset = TextEncodingTypeEnum_UTF8_BOM
'    Case Else
'        Stream.Charset = EncodingName
'    End Select
'    Stream.Open
'    Stream.LoadFromFile (FilePath)
'    LoadTextFile = Stream.ReadText
'    Stream.Close
'End Function
'
''----------------------------------------
''・テキストファイル保存
''----------------------------------------
''   ・  UTF-16LEとUTF-8はそのままだと それぞれの BOM_ON になるので
''       BON無し指定の場合は独自の処理をしている
''----------------------------------------
'Public Sub SaveTextFile(ByVal Text As String, _
'ByVal FilePath As String, ByVal EncodingName As String)
'    Call Assert(CheckEncodeName(EncodingName), "Error:ADOStream_LoadTextFile")
'
'    Dim Stream As New ADODB.Stream
'    Stream.Type = adTypeText
'    Stream.Charset = EncodingName
'    Stream.Open
'    Call Stream.WriteText(Text)
'
'    Dim ByteData() As Byte
'    Select Case UCase$(EncodingName)
'    Case TextEncodingTypeEnum_UTF16_LE_BOM_NO
'        Stream.Position = 0
'        Stream.Type = adTypeBinary
'        Stream.Position = 2
'        ByteData = Stream.Read
'        Stream.Position = 0
'        Call Stream.Write(ByteData)
'        Call Stream.SetEOS
'    Case TextEncodingTypeEnum_UTF8_BOM_NO
'        Stream.Position = 0
'        Stream.Type = adTypeBinary
'        Stream.Position = 3
'        ByteData = ADOStream.Read
'        Stream.Position = 0
'        Call Stream.Write(ByteData)
'        Call Stream.SetEOS
'    End Select
'    Call Stream.SaveToFile(FilePath, adSaveCreateOverWrite)
'    Stream.Close
'End Sub

'----------------------------------------
'◇テキストファイル読み書き Enum指定版
'----------------------------------------
'   ・  ADODB.Stream が許容する文字列は次の通り
'           使用可能文字列                  エンコード
'           SHIFT_JIS
'           UNICODEFFFE/UNICODE/UTF-16      UTF-16LE_BOM_ON
'           UTF-16LE                        UTF-16LE_BOM_OFF
'           UNICODEFEFF                     UTF-16BE_BOM_ON
'           UTF-16BE                        UTF-16BE_BOM_OFF
'           UTF-8
'           ISO-2022-JP
'           EUC-JP
'           UTF-7
'       このうち、変わった挙動をするのは UTF-16LE と UFT-8
'
'       UTF-16LEは、読み込み時に指定すると
'       テキストファイルが UTF-16LE のBOMありなし関わらず読み込み可能
'       これは特に問題にはならないのだが
'       書き込み時に UTF-16LE を指定しても
'       BOMありとして書き込まれてしまう。
'       つまり、UTF-16LEは UNICODEFFFE/UNICODE/UTF-16 と同じ機能になる
'       それでは機能不足なので、String_SaveToFile ではBOMを除外する処理をしている。
'
'       UFT-8は、常にBOMありとして書き込まれる。
'       BOM無しのUTF-8なんて世の中に存在しない方がいいのだが
'       そうはいっても、UTF-8BOM無しで出力したい場合もあるので
'       ADODB.Stream はBOM無しUTF-8を許容しないのだが
'       UFT-8N という文字列によって
'       UTF-8のBOMなしの文字として表現して、
'       String_SaveToFile ではBOMを除外する処理をしている。
'----------------------------------------
'Public Function GetEncodingTypeName( _
'ByVal EncodingTypeName As String) As EncodingTypeJpCharCode
'
'   Dim Result As EncodingTypeJpCharCode
'   Result = EncodingTypeJpCharCode.NONE
'   Select Case UCase(EncodingTypeName)
'   Case "SHIFT_JIS"
'       Result = EncodingTypeJpCharCode.Shift_JIS
'
'   Case "UNICODE", "UNICODEFFFE", "UTF-16"
'       Result = EncodingTypeJpCharCode.UTF16_LE_BOM
'   Case "UTF-16LE"
'       Result = EncodingTypeJpCharCode.UTF16_LE_BOM_NO
'
'   Case "UNICODEFEFF"
'       Result = EncodingTypeJpCharCode.UTF16_BE_BOM
'   Case "UTF-16BE"
'       Result = EncodingTypeJpCharCode.UTF16_BE_BOM_NO
'
'   Case "UTF-8"
'       Result = EncodingTypeJpCharCode.UTF8_BOM
'   Case "UTF-8N"
'       Result = EncodingTypeJpCharCode.UTF8_BOM_NO
'
'   Case "ISO-2022-JP"
'       Result = EncodingTypeJpCharCode.JIS
'
'   Case "EUC-JP"
'       Result = EncodingTypeJpCharCode.EUC_JP
'
'   Case "UTF-7"
'       Result = EncodingTypeJpCharCode.UTF_7
'
'   End Select
'
'End Function
'

Public Function GetEncodingTypeName( _
ByVal EncodingType As TextFileEncoding) As String
   Dim Result As String: Result = ""

   Select Case EncodingType
   Case TextFileEncoding.Shift_JIS
       Result = "SHIFT_JIS"
       
   Case TextFileEncoding.UTF8_BOM
       Result = "UTF-8"
   Case TextFileEncoding.UTF8_BOM_NO
       Result = "UTF-8N"

   Case TextFileEncoding.UTF16_LE_BOM
       Result = "UNICODEFFFE"
   Case TextFileEncoding.UTF16_LE_BOM_NO
       Result = "UTF-16LE"

   Case TextFileEncoding.UTF16_BE_BOM
       Result = "UNICODEFEFF"
   Case TextFileEncoding.UTF16_BE_BOM_NO
       Result = "UTF-16BE"

   Case TextFileEncoding.JIS
       Result = "ISO-2022-JP"

   Case TextFileEncoding.EUC_JP
       Result = "EUC-JP"

   Case TextFileEncoding.UTF_7
       Result = "UTF-7"

   Case TextFileEncoding.ASCII
       Result = "ASCII"
       
   End Select
   GetEncodingTypeName = Result
End Function

Public Function LoadTextFile( _
ByVal FilePath As String, _
ByVal EncodingType As TextFileEncoding) As String

   Dim EncordingName As String
   EncordingName = GetEncodingTypeName(EncodingType)
   Call Assert(EncordingName <> "", "Error:Encoding No Select")

   Dim Stream As New ADODB.Stream
   Stream.Type = adTypeText
   Select Case EncodingType
   Case TextFileEncoding.UTF8_BOM_NO
       Stream.Charset = GetEncodingTypeName(TextFileEncoding.UTF8_BOM)
   Case Else
       Stream.Charset = EncordingName
   End Select
   Stream.Open
   Stream.LoadFromFile (FilePath)
   LoadTextFile = Stream.ReadText
   Stream.Close

End Function

Public Sub testLoadTextFile()
   Dim FolderPath As String
   FolderPath = PathCombine( _
       ThisWorkbook.Path, "Test", "ADOStream")
   Call ForceCreateFolder(FolderPath)

   Call Assert("Shift-JIS ＡＢＣ１２３" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_Shift-JIS.txt"), _
           TextFileEncoding.Shift_JIS))

   Call Assert("UTF-16LE-BOM ＡＢＣ１２３" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16LE-BOM.txt"), _
           TextFileEncoding.UTF16_LE_BOM))
   Call Assert("UTF-16LE-BOM-NO ＡＢＣ１２３" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16LE-BOM-NO.txt"), _
           TextFileEncoding.UTF16_LE_BOM_NO))

   Call Assert("UTF-16BE-BOM ＡＢＣ１２３" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16BE-BOM.txt"), _
           TextFileEncoding.UTF16_BE_BOM))
   Call Assert("UTF-16BE-BOM-NO ＡＢＣ１２３" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16BE-BOM-NO.txt"), _
           TextFileEncoding.UTF16_BE_BOM_NO))

   Call Assert("UTF-8-BOM ＡＢＣ１２３" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-8-BOM.txt"), _
           TextFileEncoding.UTF8_BOM))
   Call Assert("UTF-8-BOM-NO ＡＢＣ１２３" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-8-BOM-NO.txt"), _
           TextFileEncoding.UTF8_BOM_NO))

   Call Assert("JIS ISO-2022-JP ＡＢＣ１２３" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_JIS.txt"), _
           TextFileEncoding.JIS))

   Call Assert("EUC-JP ＡＢＣ１２３" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_EUC-JP.txt"), _
           TextFileEncoding.EUC_JP))

   Call Assert("UTF-7 ＡＢＣ１２３" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-7.txt"), _
           TextFileEncoding.UTF_7))
End Sub

Public Sub SaveTextFile( _
ByVal Text As String, _
ByVal FilePath As String, _
ByVal EncodingType As TextFileEncoding)

   Dim EncordingName As String
   EncordingName = GetEncodingTypeName(EncodingType)
   Call Assert(EncordingName <> "", "Error:Encoding No Select")

   Dim Stream As New ADODB.Stream
   Stream.Type = adTypeText
   Select Case EncodingType
   Case TextFileEncoding.UTF8_BOM_NO
       Stream.Charset = GetEncodingTypeName(TextFileEncoding.UTF8_BOM)
   Case Else
       Stream.Charset = EncordingName
   End Select
   Stream.Open
   Call Stream.WriteText(Text)

   Dim ByteData() As Byte
   Select Case EncodingType
   Case TextFileEncoding.UTF16_LE_BOM_NO
       Stream.Position = 0
       Stream.Type = adTypeBinary
       Stream.Position = 2
       ByteData = Stream.Read
       Stream.Close
       Stream.Open
       Call Stream.Write(ByteData)
   Case TextFileEncoding.UTF8_BOM_NO
       Stream.Position = 0
       Stream.Type = adTypeBinary
       Stream.Position = 3
       ByteData = Stream.Read
       Stream.Close
       Stream.Open
       Call Stream.Write(ByteData)
   End Select
   Call Stream.SaveToFile(FilePath, adSaveCreateOverWrite)
   Stream.Close
End Sub

Public Sub testString_SaveToFile()
   Dim FolderPath As String
   FolderPath = PathCombine( _
       ThisWorkbook.Path, "Test", "ADOStream")
   Call ForceCreateFolder(FolderPath)

   Call SaveTextFile( _
       "Shift-JIS ＡＢＣ１２３", _
       PathCombine(FolderPath, "test_Shift-JIS.txt"), _
       TextFileEncoding.Shift_JIS)

   Call SaveTextFile( _
       "UTF-16LE-BOM ＡＢＣ１２３", _
       PathCombine(FolderPath, "test_UTF-16LE-BOM.txt"), _
       TextFileEncoding.UTF16_LE_BOM)
   Call SaveTextFile( _
       "UTF-16LE-BOM-NO ＡＢＣ１２３", _
       PathCombine(FolderPath, "test_UTF-16LE-BOM-NO.txt"), _
       TextFileEncoding.UTF16_LE_BOM_NO)

   Call SaveTextFile( _
       "UTF-16BE-BOM ＡＢＣ１２３", _
       PathCombine(FolderPath, "test_UTF-16BE-BOM.txt"), _
       TextFileEncoding.UTF16_BE_BOM)
   Call SaveTextFile( _
       "UTF-16BE-BOM-NO ＡＢＣ１２３", _
       PathCombine(FolderPath, "test_UTF-16BE-BOM-NO.txt"), _
       TextFileEncoding.UTF16_BE_BOM_NO)

   Call SaveTextFile( _
       "UTF-8-BOM ＡＢＣ１２３", _
       PathCombine(FolderPath, "test_UTF-8-BOM.txt"), _
       TextFileEncoding.UTF8_BOM)
   Call SaveTextFile( _
       "UTF-8-BOM-NO ＡＢＣ１２３", _
       PathCombine(FolderPath, "test_UTF-8-BOM-NO.txt"), _
       TextFileEncoding.UTF8_BOM_NO)

   Call SaveTextFile( _
       "JIS ISO-2022-JP ＡＢＣ１２３", _
       PathCombine(FolderPath, "test_JIS.txt"), _
       TextFileEncoding.JIS)

   Call SaveTextFile( _
       "EUC-JP ＡＢＣ１２３", _
       PathCombine(FolderPath, "test_EUC-JP.txt"), _
       TextFileEncoding.EUC_JP)

   Call SaveTextFile( _
       "UTF-7 ＡＢＣ１２３", _
       PathCombine(FolderPath, "test_UTF-7.txt"), _
       TextFileEncoding.UTF_7)

End Sub


