Attribute VB_Name = "Factory"
Option Explicit

'----------------------------------------------------------------------------------------
' (�T�v)
'  Symbols�N���X�̃C���X�^���X�𐶐����܂��B
'
' (�p�����[�^)
'�@maxVer                : �^�Ԃ̏��
'  ecLevel               : ���������x��
'  allowStructuredAppend : �����V���{���ւ̕�����������ɂ� True ���w�肵�܂��B
'  byteModeCharsetName   : �o�C�g���[�h�̕����R�[�h��"Shift_JIS" �܂��� "UTF-8" �Ŏw�肵�܂��B
'----------------------------------------------------------------------------------------
Public Function NewSymbols(Optional ByVal maxVer As Long = Constants.MAX_VERSION, _
                           Optional ByVal ecLevel As ErrorCorrectionLevel = ErrorCorrectionLevel.M, _
                           Optional ByVal allowStructuredAppend As Boolean = False, _
                           Optional ByVal byteModeCharsetName As String = "Shift_JIS") As Symbols
    
    If LCase(byteModeCharsetName) <> "shift_jis" And _
       LCase(byteModeCharsetName) <> "utf-8" Then
        Err.Raise 5
    End If
    
    Dim sbls As New Symbols
    
    Call sbls.Initialize(maxVer, ecLevel, allowStructuredAppend, byteModeCharsetName)
    Set NewSymbols = sbls
    
End Function
