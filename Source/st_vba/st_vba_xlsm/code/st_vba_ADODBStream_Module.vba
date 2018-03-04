'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    ADODBStream Module
'ObjectName:    st_vba_ADODBStream
'--------------------------------------------------
'Version:       2018/03/04
'--------------------------------------------------
'   �E  Win7 64bit�� Excel2016 64bit�ł�
'       ADODB.Stream��32bit�łɃ����N���Ă��܂�
'       ���������삵�Ȃ��ꍇ�����������߂ɕ���
'   �E  �Q�Ɛݒ�����̂悤�ɐݒ肵�Ă�
'           Microsoft AxtiveX Data Objects 6.1 Library
'           C:\Program Files\Common Files\System\ado\msado15.dll
'       ���̂悤�ɏ���ɏC������ē���s����N�����Ă���
'           C:\Program Files (x86)\Common Files\System\ado\msado15.dll
'--------------------------------------------------
Option Explicit


'----------------------------------------
'���e�L�X�g�t�@�C���ǂݏ��� Enum�w���
'----------------------------------------
'   �E  ADODB.Stream �����e���镶����͎��̒ʂ�
'           �g�p�\������                  �G���R�[�h
'           SHIFT_JIS
'           UNICODEFFFE/UNICODE/UTF-16      UTF-16LE_BOM_ON
'           UTF-16LE                        UTF-16LE_BOM_OFF
'           UNICODEFEFF                     UTF-16BE_BOM_ON
'           UTF-16BE                        UTF-16BE_BOM_OFF
'           UTF-8
'           ISO-2022-JP
'           EUC-JP
'           UTF-7
'       ���̂����A�ς��������������̂� UTF-16LE �� UFT-8
'

'       ADODB.Stream�� UTF-16LE BOM���� �ɑΉ����Ă��Ȃ��B
'       UTF-16LE BOM�����e�L�X�g��
'       UNICODEFFFE/UNICODE/UTF-16/UTF-16LE
'       �ǂ���w�肵�Ă��ǂݍ��݂͉\�B
'       �������ݎ��� UTF-16LE ���w�肵�Ă�
'       BOM����Ƃ��ď������܂�Ă��܂��A
'       �܂�AUTF-16LE�� UNICODEFFFE/UNICODE/UTF-16 �Ɠ����@�\�ɂȂ�
'       ����ł͋@�\�s���Ȃ̂�
'       String_SaveTextFile �ł�BOM�����O���鏈�������Ă���B
'
'       ADODB.Stream�� UTF-8 BOM���� �ɑΉ����Ă��Ȃ��B
'       �ǂݍ��݂� UTF-8 BOM���� �ł��ǂݍ��߂�B
'       �������݂͎w��ł��镶�����Ȃ��̂�
'       UTF-8N ���`����
'       String_SaveToFile �ł�BOM�����O���鏈�������Ă���B
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

   Call Assert("Shift-JIS �`�a�b�P�Q�R" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_Shift-JIS.txt"), _
           EncodingTypeJpCharCode.Shift_JIS))

   Call Assert("UTF-16LE-BOM �`�a�b�P�Q�R" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16LE-BOM.txt"), _
           EncodingTypeJpCharCode.UTF16_LE_BOM))
   Call Assert("UTF-16LE-BOM-NO �`�a�b�P�Q�R" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16LE-BOM-NO.txt"), _
           EncodingTypeJpCharCode.UTF16_LE_BOM_NO))

   Call Assert("UTF-16BE-BOM �`�a�b�P�Q�R" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16BE-BOM.txt"), _
           EncodingTypeJpCharCode.UTF16_BE_BOM))
   Call Assert("UTF-16BE-BOM-NO �`�a�b�P�Q�R" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-16BE-BOM-NO.txt"), _
           EncodingTypeJpCharCode.UTF16_BE_BOM_NO))

   Call Assert("UTF-8-BOM �`�a�b�P�Q�R" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-8-BOM.txt"), _
           EncodingTypeJpCharCode.UTF8_BOM))
   Call Assert("UTF-8-BOM-NO �`�a�b�P�Q�R" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_UTF-8-BOM-NO.txt"), _
           EncodingTypeJpCharCode.UTF8_BOM_NO))

   Call Assert("JIS ISO-2022-JP �`�a�b�P�Q�R" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_JIS.txt"), _
           EncodingTypeJpCharCode.JIS))

   Call Assert("EUC-JP �`�a�b�P�Q�R" = _
       String_LoadTextFile( _
           PathCombine(FolderPath, "test_EUC-JP.txt"), _
           EncodingTypeJpCharCode.EUC_JP))

   Call Assert("UTF-7 �`�a�b�P�Q�R" = _
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
       "Shift-JIS �`�a�b�P�Q�R", _
       PathCombine(FolderPath, "test_Shift-JIS.txt"), _
       EncodingTypeJpCharCode.Shift_JIS)

   Call String_SaveTextFile( _
       "UTF-16LE-BOM �`�a�b�P�Q�R", _
       PathCombine(FolderPath, "test_UTF-16LE-BOM.txt"), _
       EncodingTypeJpCharCode.UTF16_LE_BOM)
   Call String_SaveTextFile( _
       "UTF-16LE-BOM-NO �`�a�b�P�Q�R", _
       PathCombine(FolderPath, "test_UTF-16LE-BOM-NO.txt"), _
       EncodingTypeJpCharCode.UTF16_LE_BOM_NO)

   Call String_SaveTextFile( _
       "UTF-16BE-BOM �`�a�b�P�Q�R", _
       PathCombine(FolderPath, "test_UTF-16BE-BOM.txt"), _
       EncodingTypeJpCharCode.UTF16_BE_BOM)
   Call String_SaveTextFile( _
       "UTF-16BE-BOM-NO �`�a�b�P�Q�R", _
       PathCombine(FolderPath, "test_UTF-16BE-BOM-NO.txt"), _
       EncodingTypeJpCharCode.UTF16_BE_BOM_NO)

   Call String_SaveTextFile( _
       "UTF-8-BOM �`�a�b�P�Q�R", _
       PathCombine(FolderPath, "test_UTF-8-BOM.txt"), _
       EncodingTypeJpCharCode.UTF8_BOM)
   Call String_SaveTextFile( _
       "UTF-8-BOM-NO �`�a�b�P�Q�R", _
       PathCombine(FolderPath, "test_UTF-8-BOM-NO.txt"), _
       EncodingTypeJpCharCode.UTF8_BOM_NO)

   Call String_SaveTextFile( _
       "JIS ISO-2022-JP �`�a�b�P�Q�R", _
       PathCombine(FolderPath, "test_JIS.txt"), _
       EncodingTypeJpCharCode.JIS)

   Call String_SaveTextFile( _
       "EUC-JP �`�a�b�P�Q�R", _
       PathCombine(FolderPath, "test_EUC-JP.txt"), _
       EncodingTypeJpCharCode.EUC_JP)

   Call String_SaveTextFile( _
       "UTF-7 �`�a�b�P�Q�R", _
       PathCombine(FolderPath, "test_UTF-7.txt"), _
       EncodingTypeJpCharCode.UTF_7)

End Sub


'----------------------------------------
'��CSV/TSV�t�@�C��
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
