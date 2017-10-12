VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Symbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_items  As Collection

Private m_minVersion As Long
Private m_maxVersion              As Long
Private m_errorCorrectionLevel    As ErrorCorrectionLevel
Private m_structuredAppendAllowed As Boolean

Private m_structuredAppendParity As Long

Private m_currSymbol As Symbol

Private m_encNum    As IQRCodeEncoder
Private m_encAlpha  As IQRCodeEncoder
Private m_encByte   As IQRCodeEncoder
Private m_encKanji  As IQRCodeEncoder

'----------------------------------------------------------------------------------------
' (�T�v)
'  �I�u�W�F�N�g�����������܂��B
'----------------------------------------------------------------------------------------
Friend Sub Initialize(ByVal maxVer As Long, _
                      ByVal ecLevel As ErrorCorrectionLevel, _
                      ByVal allowStructuredAppend As Boolean)
    
    If maxVer < Constants.MIN_VERSION Or _
       maxVer > Constants.MAX_VERSION Then
        Err.Raise 5
    End If

    Set m_items = New Collection
    
    Set m_encNum = New NumericEncoder
    Set m_encAlpha = New AlphanumericEncoder
    Set m_encByte = New ByteEncoder
    Set m_encKanji = New KanjiEncoder

    m_minVersion = 1
    m_maxVersion = maxVer
    m_errorCorrectionLevel = ecLevel
    m_structuredAppendAllowed = allowStructuredAppend

    m_structuredAppendParity = 0

    Set m_currSymbol = New Symbol
    Call m_currSymbol.Initialize(Me)
    Call m_items.Add(m_currSymbol)

End Sub

'----------------------------------------------------------------------------------------
' (�T�v)
'  �w�肵���C���f�b�N�X�ԍ��̃V���{����Ԃ��܂��B
'  ���胁���o�Ƃ��� [Attribute Item.VB_UserMemId = 0] ���w�肳��Ă��܂��B
'----------------------------------------------------------------------------------------
Public Property Get Item(ByVal idx As Long) As Symbol
Attribute Item.VB_UserMemId = 0
    Set Item = m_items(idx + 1)

End Property

'----------------------------------------------------------------------------------------
' (�T�v)
'  �V���{�������擾���܂��B
'----------------------------------------------------------------------------------------
Public Property Get Count() As Long
    
    Count = m_items.Count

End Property

'----------------------------------------------------------------------------------------
' (�T�v)
'  �\���I�A�ڃ��[�h�̎g�p�ۂ�\���u�[���l���擾���܂��B
'----------------------------------------------------------------------------------------
Friend Property Get StructuredAppendAllowed() As Boolean
    
    StructuredAppendAllowed = m_structuredAppendAllowed

End Property

'----------------------------------------------------------------------------------------
' (�T�v)
'  �\���I�A�ڂ̃p���e�B���擾���܂��B
'----------------------------------------------------------------------------------------
Friend Property Get StructuredAppendParity() As Long
    
    StructuredAppendParity = m_structuredAppendParity

End Property

'----------------------------------------------------------------------------------------
' (�T�v)
'  �^�Ԃ̉������擾�܂��͐ݒ肵�܂��B
'----------------------------------------------------------------------------------------
Friend Property Get MinVersion() As Long

    MinVersion = m_minVersion
    
End Property
Friend Property Let MinVersion(ByVal Value As Long)

    m_minVersion = Value
    
End Property

'----------------------------------------------------------------------------------------
' (�T�v)
'  �^�Ԃ̏�����擾���܂��B
'----------------------------------------------------------------------------------------
Friend Property Get MaxVersion() As Long

    MaxVersion = m_maxVersion
    
End Property

'----------------------------------------------------------------------------------------
' (�T�v)
'  ���������x�����擾���܂��B
'----------------------------------------------------------------------------------------
Friend Property Get ErrorCorrectionLevel() As ErrorCorrectionLevel

    ErrorCorrectionLevel = m_errorCorrectionLevel

End Property

'----------------------------------------------------------------------------------------
' (�T�v)
'  Collection�Ƃ��Ĉ������߂ɕK�v�ȃR�[�h�ł��B
'  [Attribute NewEnum.VB_UserMemId = -4] ���w�肳��Ă��܂��B
'----------------------------------------------------------------------------------------
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_items.[_NewEnum]

End Property

'----------------------------------------------------------------------------------------
' (�T�v)
'  �V���{����ǉ����܂��B
'
' (�߂�l)
'  �ǉ������V���{����\���I�u�W�F�N�g
'----------------------------------------------------------------------------------------
Private Function Add() As Symbol
    
#If [DEBUG] Then
    Debug.Assert m_structuredAppendAllowed
    Debug.Assert m_items.Count < 16
#End If

    Set m_currSymbol = New Symbol
    Call m_currSymbol.Initialize(Me)
    Call m_items.Add(m_currSymbol)

    Set Add = m_currSymbol
    
End Function

'----------------------------------------------------------------------------------------
' (�T�v)
'  �������ǉ����܂��B
'----------------------------------------------------------------------------------------
Public Sub AppendString(ByVal s As String)
    
    Dim oldMode As EncodingMode
    Dim newMode As EncodingMode
    Dim i       As Long

    If Len(s) = 0 Then Err.Raise 5
    
    For i = 1 To Len(s)
        oldMode = m_currSymbol.CurrentEncodingMode
    
        Select Case oldMode
            Case EncodingMode.UNKNOWN
                newMode = SelectInitialMode(s, i)
                
            Case EncodingMode.NUMERIC
                newMode = SelectModeWhileInNumericMode(s, i)
                
            Case EncodingMode.ALPHA_NUMERIC
                newMode = SelectModeWhileInAlphanumericMode(s, i)
                
            Case EncodingMode.EIGHT_BIT_BYTE
                newMode = SelectModeWhileInByteMode(s, i)
                
            Case EncodingMode.KANJI
                newMode = SelectInitialMode(s, i)
                
            Case Else
                Err.Raise 51
            
        End Select
    
        If newMode <> oldMode Then
            If Not m_currSymbol.TrySetEncodingMode(newMode, Mid$(s, i, 1)) Then
                If Not m_structuredAppendAllowed Or m_items.Count = 16 Then
                    Err.Raise 6
                End If

                Call Add
                newMode = SelectInitialMode(s, i)
                Call m_currSymbol.TrySetEncodingMode(newMode, Mid$(s, i, 1))
                
            End If
        End If
    
        If Not m_currSymbol.TryAppend(Mid$(s, i, 1)) Then
            If Not m_structuredAppendAllowed Or m_items.Count = 16 Then
                Err.Raise 6
            End If
            
            Call Add
            newMode = SelectInitialMode(s, i)
            Call m_currSymbol.TrySetEncodingMode(newMode, Mid$(s, i, 1))
            Call m_currSymbol.TryAppend(Mid$(s, i, 1))
            
        End If
    Next
    
End Sub

'----------------------------------------------------------------------------------------
' (�T�v)
'  �\���I�A�ڂ̃p���e�B���X�V���܂��B
'----------------------------------------------------------------------------------------
Friend Sub UpdateParity(ByVal c As String)
    
#If [DEBUG] Then
    Debug.Assert Len(c) = 1
#End If

    Dim code As Long
    code = Asc(c) And &HFFFF&
    
    Dim msb As Long
    Dim lsb As Long
    
    msb = (code And &HFF00&) \ 2 ^ 8
    lsb = code And &HFF&
    
    If msb > 0 Then
        m_structuredAppendParity = m_structuredAppendParity Xor msb
    End If
    
    m_structuredAppendParity = m_structuredAppendParity Xor lsb

End Sub

'----------------------------------------------------------------------------------------
' (�T�v)
'  �������[�h�����肵�܂��B
'
' (�p�����[�^)
'  s          : �Ώە�����
'  startIndex : �]�����J�n����ʒu
'----------------------------------------------------------------------------------------
Private Function SelectInitialMode( _
    ByRef s As String, ByVal startIndex As Long) As EncodingMode
    
    Dim cnt     As Long
    Dim flg     As Boolean
    Dim flg1    As Boolean
    Dim flg2    As Boolean
    Dim i       As Long
            
    If m_encKanji.IsInSubset(Mid$(s, startIndex, 1)) Then
        SelectInitialMode = EncodingMode.KANJI
        Exit Function

    ElseIf m_encByte.IsInExclusiveSubset(Mid$(s, startIndex, 1)) Then
        SelectInitialMode = EncodingMode.EIGHT_BIT_BYTE
        Exit Function
    
    ElseIf m_encAlpha.IsInExclusiveSubset(Mid$(s, startIndex, 1)) Then
        cnt = 0
        flg = False
        
        For i = startIndex To Len(s)
            If m_encAlpha.IsInExclusiveSubset(Mid$(s, i, 1)) Then
                cnt = cnt + 1
            Else
                Exit For
            End If
        Next

        Select Case m_currSymbol.Version
            Case 1 To 9
                flg = cnt < 6
                
            Case 10 To 26
                flg = cnt < 7
                
            Case 27 To 40
                flg = cnt < 8
                
            Case Else
                Err.Raise 51
            
        End Select
        
        If flg Then
            If (startIndex + cnt) <= Len(s) Then
                If m_encByte.IsInExclusiveSubset(Mid$(s, startIndex + cnt, 1)) Then
                    SelectInitialMode = EncodingMode.EIGHT_BIT_BYTE
                    Exit Function
                Else
                    SelectInitialMode = EncodingMode.ALPHA_NUMERIC
                    Exit Function
                End If
            Else
                SelectInitialMode = EncodingMode.ALPHA_NUMERIC
                Exit Function
            End If
        Else
            SelectInitialMode = EncodingMode.ALPHA_NUMERIC
            Exit Function
        End If

    ElseIf m_encNum.IsInSubset(Mid$(s, startIndex, 1)) Then
        cnt = 0
        flg1 = False
        flg2 = False
        
        For i = startIndex To Len(s)
            If m_encNum.IsInSubset(Mid$(s, i, 1)) Then
                cnt = cnt + 1
            Else
                Exit For
            End If
        Next
        
        Select Case m_currSymbol.Version
            Case 1 To 9
                flg1 = cnt < 4
                flg2 = cnt < 7
            
            Case 10 To 26
                flg1 = cnt < 4
                flg2 = cnt < 8
                
            Case 27 To 40
                flg1 = cnt < 5
                flg2 = cnt < 9
                
            Case Else
                Err.Raise 51
            
        End Select

        If flg1 Then
            If (startIndex + cnt) <= Len(s) Then
                flg1 = m_encByte.IsInExclusiveSubset(Mid$(s, startIndex + cnt, 1))
            Else
                flg1 = False
            End If
        End If

        If flg2 Then
            If (startIndex + cnt) <= Len(s) Then
                flg2 = m_encAlpha.IsInExclusiveSubset(Mid$(s, startIndex + cnt, 1))
            Else
                flg2 = False
            End If
        End If

        If flg1 Then
            SelectInitialMode = EncodingMode.EIGHT_BIT_BYTE
            Exit Function
        ElseIf flg2 Then
            SelectInitialMode = EncodingMode.ALPHA_NUMERIC
            Exit Function
        Else
            SelectInitialMode = EncodingMode.NUMERIC
            Exit Function
        End If
        
    Else
        Err.Raise 51
        
    End If

End Function

'----------------------------------------------------------------------------------------
' (�T�v)
'  �������[�h����؂�ւ��郂�[�h�����肵�܂��B
'
' (�p�����[�^)
'  s          : �Ώە�����
'  startIndex : �]�����J�n����ʒu
'----------------------------------------------------------------------------------------
Private Function SelectModeWhileInNumericMode( _
    ByRef s As String, ByVal startIndex As Long) As EncodingMode
    
    If m_encKanji.IsInSubset(Mid$(s, startIndex, 1)) Then
        SelectModeWhileInNumericMode = EncodingMode.KANJI
        Exit Function
        
    ElseIf m_encByte.IsInExclusiveSubset(Mid$(s, startIndex, 1)) Then
        SelectModeWhileInNumericMode = EncodingMode.EIGHT_BIT_BYTE
        Exit Function
    
    ElseIf m_encAlpha.IsInExclusiveSubset(Mid$(s, startIndex, 1)) Then
        SelectModeWhileInNumericMode = EncodingMode.ALPHA_NUMERIC
        Exit Function
        
    End If
    
    SelectModeWhileInNumericMode = EncodingMode.NUMERIC

End Function

'----------------------------------------------------------------------------------------
' (�T�v)
'  �p�������[�h����؂�ւ��郂�[�h�����肵�܂��B
'
' (�p�����[�^)
'  s          : �Ώە�����
'  startIndex : �]�����J�n����ʒu
'----------------------------------------------------------------------------------------
Private Function SelectModeWhileInAlphanumericMode( _
    ByRef s As String, ByVal startIndex As Long) As EncodingMode
    
    Dim cnt As Long
    Dim flg As Boolean
    Dim i   As Long
    
    If m_encKanji.IsInSubset(Mid$(s, startIndex, 1)) Then
        SelectModeWhileInAlphanumericMode = EncodingMode.KANJI
        Exit Function
    
    ElseIf m_encByte.IsInExclusiveSubset(Mid$(s, startIndex, 1)) Then
        SelectModeWhileInAlphanumericMode = EncodingMode.EIGHT_BIT_BYTE
        Exit Function
        
    End If
    
    cnt = 0
    flg = False

    For i = startIndex To Len(s)
        If Not m_encAlpha.IsInSubset(Mid$(s, i, 1)) Then
            Exit For
        End If
        
        If m_encNum.IsInSubset(Mid$(s, i, 1)) Then
            cnt = cnt + 1
        Else
            flg = True
            Exit For
        End If
    Next
        
    If flg Then
        Select Case m_currSymbol.Version
            Case 1 To 9
                flg = cnt >= 13
                
            Case 10 To 26
                flg = cnt >= 15
                
            Case 27 To 40
                flg = cnt >= 17
                
            Case Else
                Err.Raise 51
            
        End Select

        If flg Then
            SelectModeWhileInAlphanumericMode = EncodingMode.NUMERIC
            Exit Function
        End If
    End If

    SelectModeWhileInAlphanumericMode = EncodingMode.ALPHA_NUMERIC

End Function

'----------------------------------------------------------------------------------------
' (�T�v)
'  �o�C�g���[�h����؂�ւ��郂�[�h�����肵�܂��B
'
' (�p�����[�^)
'  s          : �Ώە�����
'  startIndex : �]�����J�n����ʒu
'----------------------------------------------------------------------------------------
Private Function SelectModeWhileInByteMode( _
    ByRef s As String, ByVal startIndex As Long) As EncodingMode
    
    Dim cnt     As Long
    Dim flg     As Boolean
    Dim i       As Long
    
    If m_encKanji.IsInSubset(Mid$(s, startIndex, 1)) Then
        SelectModeWhileInByteMode = EncodingMode.KANJI
        Exit Function
    End If

    For i = startIndex To Len(s)
        If Not m_encByte.IsInSubset(Mid$(s, i, 1)) Then
            Exit For
        End If
    
        If m_encNum.IsInSubset(Mid$(s, i, 1)) Then
            cnt = cnt + 1
            
        ElseIf m_encByte.IsInExclusiveSubset(Mid$(s, i, 1)) Then
            flg = True
            Exit For
            
        Else
            Exit For
            
        End If
    Next

    If flg Then
        Select Case m_currSymbol.Version
            Case 1 To 9
                flg = cnt >= 6
                
            Case 10 To 26
                flg = cnt >= 8
                
            Case 27 To 40
                flg = cnt >= 9
                
            Case Else
                Err.Raise 51
            
        End Select
        
        If flg Then
            SelectModeWhileInByteMode = EncodingMode.NUMERIC
            Exit Function
        End If
    End If

    cnt = 0
    flg = False

    For i = startIndex To Len(s)
        If Not m_encByte.IsInSubset(Mid$(s, i, 1)) Then
            Exit For
        End If
        
        If m_encAlpha.IsInExclusiveSubset(Mid$(s, i, 1)) Then
            cnt = cnt + 1
            
        ElseIf m_encByte.IsInExclusiveSubset(Mid$(s, i, 1)) Then
            flg = True
            Exit For
            
        Else
            Exit For
            
        End If
        
        i = i + 1
    Next

    If flg Then
        Select Case m_currSymbol.Version
            Case 1 To 9
                flg = cnt >= 11
                
            Case 10 To 26
                flg = cnt >= 15
                
            Case 27 To 40
                flg = cnt >= 16
                
            Case Else
                Err.Raise 51
            
        End Select

        If flg Then
            SelectModeWhileInByteMode = EncodingMode.ALPHA_NUMERIC
            Exit Function
            
        End If
        
    End If

    SelectModeWhileInByteMode = EncodingMode.EIGHT_BIT_BYTE
    
End Function
