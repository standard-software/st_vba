'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    ADODBStream Module
'ObjectName:    st_vba_ADODBStream
'--------------------------------------------------
'Version:       2017/07/02
'--------------------------------------------------
'   E  Win7 64bit”Å Excel2016 64bit”Å‚Å
'       ADODB.Stream‚ª32bit”Å‚ÉƒŠƒ“ƒN‚µ‚Ä‚µ‚Ü‚¢
'       ³‚µ‚­“®ì‚µ‚È‚¢ê‡‚ª‚ ‚Á‚½‚½‚ß‚É•ª—£
'   E  QÆİ’è‚ğŸ‚Ì‚æ‚¤‚Éİ’è‚µ‚Ä‚à
'           Microsoft AxtiveX Data Objects 6.1 Library
'           C:\Program Files\Common Files\System\ado\msado15.dll
'       Ÿ‚Ì‚æ‚¤‚ÉŸè‚ÉC³‚³‚ê‚Ä“®ì•s‹ï‡‚ğ‹N‚±‚µ‚Ä‚¢‚½
'           C:\Program Files (x86)\Common Files\System\ado\msado15.dll
'--------------------------------------------------
Option Explicit

'----------------------------------------
'ŸƒeƒLƒXƒgƒtƒ@ƒCƒ‹“Ç‚İ‘‚«
'----------------------------------------

''----------------------------------------
''EƒGƒ“ƒR[ƒhƒ^ƒCƒv‚Ìw’è
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
'' UTF16_LE_BOM‚Ìw’è‚ÍA
'' [UNICODEFFFE]‚¾‚¯‚Å‚Í‚È‚­
'' [UNICODE]‚â[UTF-16]‚à“¯‚¶“®ì‚É‚È‚é‚ª
'' UTF16_BE_BOM‚Æ‚Ì‘Î”ä‚Æ‚µ‚Ä‚í‚©‚è‚â‚·‚¢‚Ì‚Å
'' [UNICODEFFFE]‚ğÌ—p‚·‚é
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
''EƒeƒLƒXƒgƒtƒ@ƒCƒ‹“Ç
''----------------------------------------
''   E  Charset‚Éw’è‚·‚é•¶š‚ª
''       UNICODEFFFE/UNICODE/UTF-16/UTF-16LE ‚Ìê‡‚Í
''       UTF-16LE‚ÌBOM‚Ì—L–³‚É‚©‚©‚í‚ç‚¸“Ç‚İ‚Ş‚±‚Æ‚ª‚Å‚«‚é
''   E  UTF-8‚ÌBOM–³‚µ‚Ìw’è•¶š‚Í‚È‚¢‚Ì‚Å
''       UTF-8‚ğw’è‚·‚é‚ÆBOM‚Ì—L–³‚É‚©‚©‚í‚ç‚¸“Ç‚İ‚Ş‚±‚Æ‚ª‚Å‚«‚é
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
''EƒeƒLƒXƒgƒtƒ@ƒCƒ‹•Û‘¶
''----------------------------------------
''   E  UTF-16LE‚ÆUTF-8‚Í‚»‚Ì‚Ü‚Ü‚¾‚Æ ‚»‚ê‚¼‚ê‚Ì BOM_ON ‚É‚È‚é‚Ì‚Å
''       BON–³‚µw’è‚Ìê‡‚Í“Æ©‚Ìˆ—‚ğ‚µ‚Ä‚¢‚é
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
'ƒeƒLƒXƒgƒtƒ@ƒCƒ‹“Ç‚İ‘‚« Enumw’è”Å
'----------------------------------------
'   E  ADODB.Stream ‚ª‹–—e‚·‚é•¶š—ñ‚ÍŸ‚Ì’Ê‚è
'           g—p‰Â”\•¶š—ñ                  ƒGƒ“ƒR[ƒh
'           SHIFT_JIS
'           UNICODEFFFE/UNICODE/UTF-16      UTF-16LE_BOM_ON
'           UTF-16LE                        UTF-16LE_BOM_OFF
'           UNICODEFEFF                     UTF-16BE_BOM_ON
'           UTF-16BE                        UTF-16BE_BOM_OFF
'           UTF-8
'           ISO-2022-JP
'           EUC-JP
'           UTF-7
'       ‚±‚Ì‚¤‚¿A•Ï‚í‚Á‚½‹““®‚ğ‚·‚é‚Ì‚Í UTF-16LE ‚Æ UFT-8
'
'       UTF-16LE‚ÍA“Ç‚İ‚İ‚Éw’è‚·‚é‚Æ
'       ƒeƒLƒXƒgƒtƒ@ƒCƒ‹‚ª UTF-16LE ‚ÌBOM‚ ‚è‚È‚µŠÖ‚í‚ç‚¸“Ç‚İ‚İ‰Â”\
'       ‚±‚ê‚Í“Á‚É–â‘è‚É‚Í‚È‚ç‚È‚¢‚Ì‚¾‚ª
'       ‘‚«‚İ‚É UTF-16LE ‚ğw’è‚µ‚Ä‚à
'       BOM‚ ‚è‚Æ‚µ‚Ä‘‚«‚Ü‚ê‚Ä‚µ‚Ü‚¤B
'       ‚Â‚Ü‚èAUTF-16LE‚Í UNICODEFFFE/UNICODE/UTF-16 ‚Æ“¯‚¶‹@”\‚É‚È‚é
'       ‚»‚ê‚Å‚Í‹@”\•s‘«‚È‚Ì‚ÅAString_SaveToFile ‚Å‚ÍBOM‚ğœŠO‚·‚éˆ—‚ğ‚µ‚Ä‚¢‚éB
'
'       UFT-8‚ÍAí‚ÉBOM‚ ‚è‚Æ‚µ‚Ä‘‚«‚Ü‚ê‚éB
'       BOM–³‚µ‚ÌUTF-8‚È‚ñ‚Ä¢‚Ì’†‚É‘¶İ‚µ‚È‚¢•û‚ª‚¢‚¢‚Ì‚¾‚ª
'       ‚»‚¤‚Í‚¢‚Á‚Ä‚àAUTF-8BOM–³‚µ‚Åo—Í‚µ‚½‚¢ê‡‚à‚ ‚é‚Ì‚Å
'       ADODB.Stream ‚ÍBOM–³‚µUTF-8‚ğ‹–—e‚µ‚È‚¢‚Ì‚¾‚ª
'       UFT-8N ‚Æ‚¢‚¤•¶š—ñ‚É‚æ‚Á‚Ä
'       UTF-8‚ÌBOM‚È‚µ‚Ì•¶š‚Æ‚µ‚Ä•\Œ»‚µ‚ÄA
'       String_SaveToFile ‚Å‚ÍBOM‚ğœŠO‚·‚éˆ—‚ğ‚µ‚Ä‚¢‚éB
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

   Call Assert("Shift-JIS ‚`‚a‚b‚P‚Q‚R" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_Shift-JIS.txt"), _
           TextFileEncoding.Shift_JIS))

   Call Assert("UTF-16LE-BOM ‚`‚a‚b‚P‚Q‚R" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16LE-BOM.txt"), _
           TextFileEncoding.UTF16_LE_BOM))
   Call Assert("UTF-16LE-BOM-NO ‚`‚a‚b‚P‚Q‚R" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16LE-BOM-NO.txt"), _
           TextFileEncoding.UTF16_LE_BOM_NO))

   Call Assert("UTF-16BE-BOM ‚`‚a‚b‚P‚Q‚R" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16BE-BOM.txt"), _
           TextFileEncoding.UTF16_BE_BOM))
   Call Assert("UTF-16BE-BOM-NO ‚`‚a‚b‚P‚Q‚R" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16BE-BOM-NO.txt"), _
           TextFileEncoding.UTF16_BE_BOM_NO))

   Call Assert("UTF-8-BOM ‚`‚a‚b‚P‚Q‚R" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-8-BOM.txt"), _
           TextFileEncoding.UTF8_BOM))
   Call Assert("UTF-8-BOM-NO ‚`‚a‚b‚P‚Q‚R" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-8-BOM-NO.txt"), _
           TextFileEncoding.UTF8_BOM_NO))

   Call Assert("JIS ISO-2022-JP ‚`‚a‚b‚P‚Q‚R" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_JIS.txt"), _
           TextFileEncoding.JIS))

   Call Assert("EUC-JP ‚`‚a‚b‚P‚Q‚R" = _
       LoadTextFile( _
           PathCombine(FolderPath, "test_EUC-JP.txt"), _
           TextFileEncoding.EUC_JP))

   Call Assert("UTF-7 ‚`‚a‚b‚P‚Q‚R" = _
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
       "Shift-JIS ‚`‚a‚b‚P‚Q‚R", _
       PathCombine(FolderPath, "test_Shift-JIS.txt"), _
       TextFileEncoding.Shift_JIS)

   Call SaveTextFile( _
       "UTF-16LE-BOM ‚`‚a‚b‚P‚Q‚R", _
       PathCombine(FolderPath, "test_UTF-16LE-BOM.txt"), _
       TextFileEncoding.UTF16_LE_BOM)
   Call SaveTextFile( _
       "UTF-16LE-BOM-NO ‚`‚a‚b‚P‚Q‚R", _
       PathCombine(FolderPath, "test_UTF-16LE-BOM-NO.txt"), _
       TextFileEncoding.UTF16_LE_BOM_NO)

   Call SaveTextFile( _
       "UTF-16BE-BOM ‚`‚a‚b‚P‚Q‚R", _
       PathCombine(FolderPath, "test_UTF-16BE-BOM.txt"), _
       TextFileEncoding.UTF16_BE_BOM)
   Call SaveTextFile( _
       "UTF-16BE-BOM-NO ‚`‚a‚b‚P‚Q‚R", _
       PathCombine(FolderPath, "test_UTF-16BE-BOM-NO.txt"), _
       TextFileEncoding.UTF16_BE_BOM_NO)

   Call SaveTextFile( _
       "UTF-8-BOM ‚`‚a‚b‚P‚Q‚R", _
       PathCombine(FolderPath, "test_UTF-8-BOM.txt"), _
       TextFileEncoding.UTF8_BOM)
   Call SaveTextFile( _
       "UTF-8-BOM-NO ‚`‚a‚b‚P‚Q‚R", _
       PathCombine(FolderPath, "test_UTF-8-BOM-NO.txt"), _
       TextFileEncoding.UTF8_BOM_NO)

   Call SaveTextFile( _
       "JIS ISO-2022-JP ‚`‚a‚b‚P‚Q‚R", _
       PathCombine(FolderPath, "test_JIS.txt"), _
       TextFileEncoding.JIS)

   Call SaveTextFile( _
       "EUC-JP ‚`‚a‚b‚P‚Q‚R", _
       PathCombine(FolderPath, "test_EUC-JP.txt"), _
       TextFileEncoding.EUC_JP)

   Call SaveTextFile( _
       "UTF-7 ‚`‚a‚b‚P‚Q‚R", _
       PathCombine(FolderPath, "test_UTF-7.txt"), _
       TextFileEncoding.UTF_7)

End Sub


