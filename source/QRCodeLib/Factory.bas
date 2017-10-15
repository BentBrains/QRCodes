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
'----------------------------------------------------------------------------------------
Public Function NewSymbols(Optional ByVal maxVer As Long = Constants.MAX_VERSION, _
                           Optional ByVal ecLevel As ErrorCorrectionLevel = ErrorCorrectionLevel.M, _
                           Optional ByVal allowStructuredAppend As Boolean = False) As Symbols
    
    Dim sbls As New Symbols
    
    Call sbls.Initialize(maxVer, ecLevel, allowStructuredAppend)
    Set NewSymbols = sbls
    
End Function
