Attribute VB_Name = "Factory"
Option Explicit

'------------------------------------------------------------------------------
'(Overview)
'       Create an instance of the Symbols class.
'
'(Parameters)
'       ecLevel: Error correction level
'       maxVer: Model number upper limit
'       allowStructuredAppend: Specify True to allow splitting into multiple symbols.
'       ByteModeCharsetName: Specifies character code in byte mode.
'------------------------------------------------------------------------------
Public Function CreateSymbols( _
    Optional ByVal ecLevel As ErrorCorrectionLevel = ErrorCorrectionLevel.M, _
    Optional ByVal maxVer As Long = Constants.MAX_VERSION, _
    Optional ByVal allowStructuredAppend As Boolean = False, _
    Optional ByVal ByteModeCharsetName As String = "Shift_JIS") As Symbols
    
    Dim sbls As New Symbols
    
    Call sbls.Initialize(ecLevel, maxVer, allowStructuredAppend, ByteModeCharsetName)
    Set CreateSymbols = sbls
    
End Function
