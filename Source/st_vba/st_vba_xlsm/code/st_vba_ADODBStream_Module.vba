'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    ADODBStream Module
'ObjectName:    st_vba_ADODBStream
'--------------------------------------------------
'Version:       2018/03/04
'--------------------------------------------------
'   ÅE  Win7 64bitî≈ Excel2016 64bitî≈Ç≈
'       ADODB.StreamÇ™32bitî≈Ç…ÉäÉìÉNÇµÇƒÇµÇ‹Ç¢
'       ê≥ÇµÇ≠ìÆçÏÇµÇ»Ç¢èÍçáÇ™Ç†Ç¡ÇΩÇΩÇﬂÇ…ï™ó£
'   ÅE  éQè∆ê›íËÇéüÇÃÇÊÇ§Ç…ê›íËÇµÇƒÇ‡
'           Microsoft AxtiveX Data Objects 6.1 Library
'           C:\Program Files\Common Files\System\ado\msado15.dll
'       éüÇÃÇÊÇ§Ç…èüéËÇ…èCê≥Ç≥ÇÍÇƒìÆçÏïsãÔçáÇãNÇ±ÇµÇƒÇ¢ÇΩ
'           C:\Program Files (x86)\Common Files\System\ado\msado15.dll
'--------------------------------------------------
Option Explicit


'----------------------------------------
'ÅûÉeÉLÉXÉgÉtÉ@ÉCÉãì«Ç›èëÇ´ EnuméwíËî≈
'----------------------------------------
'   ÅE  ADODB.Stream Ç™ãñóeÇ∑ÇÈï∂éöóÒÇÕéüÇÃí ÇË
'           égópâ¬î\ï∂éöóÒ                  ÉGÉìÉRÅ[Éh
'           SHIFT_JIS
'           UNICODEFFFE/UNICODE/UTF-16      UTF-16LE_BOM_ON
'           UTF-16LE                        UTF-16LE_BOM_OFF
'           UNICODEFEFF                     UTF-16BE_BOM_ON
'           UTF-16BE                        UTF-16BE_BOM_OFF
'           UTF-8
'           ISO-2022-JP
'           EUC-JP
'           UTF-7
'       Ç±ÇÃÇ§ÇøÅAïœÇÌÇ¡ÇΩãììÆÇÇ∑ÇÈÇÃÇÕ UTF-16LE Ç∆ UFT-8
'

'       ADODB.StreamÇÕ UTF-16LE BOMñ≥Çµ Ç…ëŒâûÇµÇƒÇ¢Ç»Ç¢ÅB
'       UTF-16LE BOMñ≥ÇµÉeÉLÉXÉgÇÕ
'       UNICODEFFFE/UNICODE/UTF-16/UTF-16LE
'       Ç«ÇÍÇéwíËÇµÇƒÇ‡ì«Ç›çûÇ›ÇÕâ¬î\ÅB
'       èëÇ´çûÇ›éûÇ… UTF-16LE ÇéwíËÇµÇƒÇ‡
'       BOMÇ†ÇËÇ∆ÇµÇƒèëÇ´çûÇ‹ÇÍÇƒÇµÇ‹Ç¢ÅA
'       Ç¬Ç‹ÇËÅAUTF-16LEÇÕ UNICODEFFFE/UNICODE/UTF-16 Ç∆ìØÇ∂ã@î\Ç…Ç»ÇÈ
'       ÇªÇÍÇ≈ÇÕã@î\ïsë´Ç»ÇÃÇ≈
'       String_SaveTextFile Ç≈ÇÕBOMÇèúäOÇ∑ÇÈèàóùÇÇµÇƒÇ¢ÇÈÅB
'
'       ADODB.StreamÇÕ UTF-8 BOMñ≥Çµ Ç…ëŒâûÇµÇƒÇ¢Ç»Ç¢ÅB
'       ì«Ç›çûÇ›ÇÕ UTF-8 BOMñ≥Çµ Ç≈Ç‡ì«Ç›çûÇﬂÇÈÅB
'       èëÇ´çûÇ›ÇÕéwíËÇ≈Ç´ÇÈï∂éöÇ™Ç»Ç¢ÇÃÇ≈
'       UTF-8N ÇíËã`ÇµÇƒ
'       String_SaveToFile Ç≈ÇÕBOMÇèúäOÇ∑ÇÈèàóùÇÇµÇƒÇ¢ÇÈÅB
'----------------------------------------
Public Function GetEncodingTypeJpCharCode( _
ByVal EncodingTypeName As String) As EncodingTypeJpCharCode

   Dim Result As EncodingTypeJpCharCode
   Result = EncodingTypeJpCharCode.NONE
   Select Case UCase(EncodingTypeName)
   Case "SHIFT_JIS"
       Result = EncodingTypeJpCharCode.Shift_JIS

   Case "UNICODE", "UNICODEFFFE", "UTF-16"
       Result = EncodingTypeJpCharCode.UTF16_LE_BOM
   Case "UTF-16LE"
       Result = EncodingTypeJpCharCode.UTF16_LE_BOM_NO

   Case "UNICODEFEFF"
       Result = EncodingTypeJpCharCode.UTF16_BE_BOM
   Case "UTF-16BE"
       Result = EncodingTypeJpCharCode.UTF16_BE_BOM_NO

   Case "UTF-8"
       Result = EncodingTypeJpCharCode.UTF8_BOM
   Case "UTF-8N"
       Result = EncodingTypeJpCharCode.UTF8_BOM_NO

   Case "ISO-2022-JP"
       Result = EncodingTypeJpCharCode.JIS

   Case "EUC-JP"
       Result = EncodingTypeJpCharCode.EUC_JP

   Case "UTF-7"
       Result = EncodingTypeJpCharCode.UTF_7

   End Select

End Function


Public Function GetEncodingTypeName( _
ByVal EncodingType As EncodingTypeJpCharCode) As String
   Dim Result As String: Result = ""

   Select Case EncodingType
   Case EncodingTypeJpCharCode.Shift_JIS
       Result = "SHIFT_JIS"

   Case EncodingTypeJpCharCode.UTF16_LE_BOM
       Result = "UNICODEFFFE"
   Case EncodingTypeJpCharCode.UTF16_LE_BOM_NO
       Result = "UTF-16LE"

   Case EncodingTypeJpCharCode.UTF16_BE_BOM
       Result = "UNICODEFEFF"
   Case EncodingTypeJpCharCode.UTF16_BE_BOM_NO
       Result = "UTF-16BE"

   Case EncodingTypeJpCharCode.UTF8_BOM
       Result = "UTF-8"
   Case EncodingTypeJpCharCode.UTF8_BOM_NO
       Result = "UTF-8N"

   Case EncodingTypeJpCharCode.JIS
       Result = "ISO-2022-JP"

   Case EncodingTypeJpCharCode.EUC_JP
       Result = "EUC-JP"

   Case EncodingTypeJpCharCode.UTF_7
       Result = "UTF-7"

   End Select
   GetEncodingTypeName = Result
End Function

Public Function String_LoadTextFile( _
ByVal FilePath As String, _
ByVal EncodingType As EncodingTypeJpCharCode) As String

   Dim EncordingName As String
   EncordingName = GetEncodingTypeName(EncodingType)
   Call Assert(EncordingName <> "", "Error:Encoding No Select")

   Dim Stream As New ADODB.Stream
   Stream.Type = adTypeText
   Select Case EncodingType
   Case EncodingTypeJpCharCode.UTF8_BOM_NO
       Stream.Charset = GetEncodingTypeName(EncodingTypeJpCharCode.UTF8_BOM)
   Case Else
       Stream.Charset = EncordingName
   End Select
   Stream.Open
   Stream.LoadFromFile (FilePath)
   String_LoadTextFile = Stream.ReadText
   Stream.Close

End Function

Public Sub testString_LoadTextFile()
   Dim FolderPath As String
   FolderPath = PathCombine( _
       ThisWorkbook.Path, "Test", "ADOStream")
   Call ForceCreateFolder(FolderPath)

   Call Assert("Shift-JIS Ç`ÇaÇbÇPÇQÇR" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_Shift-JIS.txt"), _
           EncodingTypeJpCharCode.Shift_JIS))

   Call Assert("UTF-16LE-BOM Ç`ÇaÇbÇPÇQÇR" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16LE-BOM.txt"), _
           EncodingTypeJpCharCode.UTF16_LE_BOM))
   Call Assert("UTF-16LE-BOM-NO Ç`ÇaÇbÇPÇQÇR" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16LE-BOM-NO.txt"), _
           EncodingTypeJpCharCode.UTF16_LE_BOM_NO))

   Call Assert("UTF-16BE-BOM Ç`ÇaÇbÇPÇQÇR" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16BE-BOM.txt"), _
           EncodingTypeJpCharCode.UTF16_BE_BOM))
   Call Assert("UTF-16BE-BOM-NO Ç`ÇaÇbÇPÇQÇR" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16BE-BOM-NO.txt"), _
           EncodingTypeJpCharCode.UTF16_BE_BOM_NO))

   Call Assert("UTF-8-BOM Ç`ÇaÇbÇPÇQÇR" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-8-BOM.txt"), _
           EncodingTypeJpCharCode.UTF8_BOM))
   Call Assert("UTF-8-BOM-NO Ç`ÇaÇbÇPÇQÇR" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-8-BOM-NO.txt"), _
           EncodingTypeJpCharCode.UTF8_BOM_NO))

   Call Assert("JIS ISO-2022-JP Ç`ÇaÇbÇPÇQÇR" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_JIS.txt"), _
           EncodingTypeJpCharCode.JIS))

   Call Assert("EUC-JP Ç`ÇaÇbÇPÇQÇR" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_EUC-JP.txt"), _
           EncodingTypeJpCharCode.EUC_JP))

   Call Assert("UTF-7 Ç`ÇaÇbÇPÇQÇR" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-7.txt"), _
           EncodingTypeJpCharCode.UTF_7))
End Sub

Public Sub String_SaveTextFile( _
ByVal Text As String, _
ByVal FilePath As String, _
ByVal EncodingType As EncodingTypeJpCharCode)

   Dim EncordingName As String
   EncordingName = GetEncodingTypeName(EncodingType)
   Call Assert(EncordingName <> "", "Error:Encoding No Select")

   Dim Stream As New ADODB.Stream
   Stream.Type = adTypeText
   Select Case EncodingType
   Case EncodingTypeJpCharCode.UTF8_BOM_NO
       Stream.Charset = GetEncodingTypeName(EncodingTypeJpCharCode.UTF8_BOM)
   Case Else
       Stream.Charset = EncordingName
   End Select
   Stream.Open
   Call Stream.WriteText(Text)

   Dim ByteData() As Byte
   Select Case EncodingType
   Case EncodingTypeJpCharCode.UTF16_LE_BOM_NO
       Stream.Position = 0
       Stream.Type = adTypeBinary
       Stream.Position = 2
       ByteData = Stream.Read
       Stream.Close
       Stream.Open
       Call Stream.Write(ByteData)
   Case EncodingTypeJpCharCode.UTF8_BOM_NO
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

Public Sub testString_SaveTextFile()
   Dim FolderPath As String
   FolderPath = PathCombine( _
       ThisWorkbook.Path, "Test", "ADOStream")
   Call ForceCreateFolder(FolderPath)

   Call String_SaveTextFile( _
       "Shift-JIS Ç`ÇaÇbÇPÇQÇR", _
       PathCombine(FolderPath, "test_Shift-JIS.txt"), _
       EncodingTypeJpCharCode.Shift_JIS)

   Call String_SaveTextFile( _
       "UTF-16LE-BOM Ç`ÇaÇbÇPÇQÇR", _
       PathCombine(FolderPath, "test_UTF-16LE-BOM.txt"), _
       EncodingTypeJpCharCode.UTF16_LE_BOM)
   Call String_SaveTextFile( _
       "UTF-16LE-BOM-NO Ç`ÇaÇbÇPÇQÇR", _
       PathCombine(FolderPath, "test_UTF-16LE-BOM-NO.txt"), _
       EncodingTypeJpCharCode.UTF16_LE_BOM_NO)

   Call String_SaveTextFile( _
       "UTF-16BE-BOM Ç`ÇaÇbÇPÇQÇR", _
       PathCombine(FolderPath, "test_UTF-16BE-BOM.txt"), _
       EncodingTypeJpCharCode.UTF16_BE_BOM)
   Call String_SaveTextFile( _
       "UTF-16BE-BOM-NO Ç`ÇaÇbÇPÇQÇR", _
       PathCombine(FolderPath, "test_UTF-16BE-BOM-NO.txt"), _
       EncodingTypeJpCharCode.UTF16_BE_BOM_NO)

   Call String_SaveTextFile( _
       "UTF-8-BOM Ç`ÇaÇbÇPÇQÇR", _
       PathCombine(FolderPath, "test_UTF-8-BOM.txt"), _
       EncodingTypeJpCharCode.UTF8_BOM)
   Call String_SaveTextFile( _
       "UTF-8-BOM-NO Ç`ÇaÇbÇPÇQÇR", _
       PathCombine(FolderPath, "test_UTF-8-BOM-NO.txt"), _
       EncodingTypeJpCharCode.UTF8_BOM_NO)

   Call String_SaveTextFile( _
       "JIS ISO-2022-JP Ç`ÇaÇbÇPÇQÇR", _
       PathCombine(FolderPath, "test_JIS.txt"), _
       EncodingTypeJpCharCode.JIS)

   Call String_SaveTextFile( _
       "EUC-JP Ç`ÇaÇbÇPÇQÇR", _
       PathCombine(FolderPath, "test_EUC-JP.txt"), _
       EncodingTypeJpCharCode.EUC_JP)

   Call String_SaveTextFile( _
       "UTF-7 Ç`ÇaÇbÇPÇQÇR", _
       PathCombine(FolderPath, "test_UTF-7.txt"), _
       EncodingTypeJpCharCode.UTF_7)

End Sub


'----------------------------------------
'ÅûCSV/TSVÉtÉ@ÉCÉã
'----------------------------------------

Public Sub SheetOpenUTF8CSV( _
ByVal Sheet As Worksheet, _
ByVal FilePath As String, _
ByVal SeparateChar As String)

    Dim ScreenUpdateBuffer As Boolean
    ScreenUpdateBuffer = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim FileText As String
    FileText = String_LoadTextFile(FilePath, _
        EncodingTypeJpCharCode.UTF8_BOM)

    Dim Row As Long
    Dim Col As Long
    Dim FileLines() As String
    FileLines = Split(FileText, vbCrLf)
    For Row = 0 To ArrayCount(FileLines) - 1
        Dim FileLine() As String
        FileLine = Split(FileLines(Row), SeparateChar)
        For Col = 0 To ArrayCount(FileLine) - 1
            Sheet.Cells(Row + 1, Col + 1).Value = FileLine(Col)
        Next
    Next

    Application.ScreenUpdating = ScreenUpdateBuffer

End Sub

Public Sub SheetSaveUTF8CSV( _
ByVal Sheet As Worksheet, _
ByVal FilePath As String, _
ByVal SeparateChar As String)

    Dim LastCell As Range
    Set LastCell = DataLastCell(Sheet)

    Dim Row As Long
    Dim Col As Long

    Dim StrBuilderText As New st_vba_StringBuilder
    Dim StrBuilderLine As New st_vba_StringBuilder

    For Row = 1 To LastCell.Row
        Call StrBuilderLine.Clear
        For Col = 1 To LastCell.Column
            Call StrBuilderLine.Add(Sheet.Cells(Row, Col).Text & SeparateChar)
        Next
        Call StrBuilderText.Add( _
            ExcludeLastStr(StrBuilderLine.Text, SeparateChar) & vbCrLf)
    Next

    Call String_SaveTextFile(StrBuilderText.Text, FilePath, _
        EncodingTypeJpCharCode.UTF8_BOM)

End Sub
