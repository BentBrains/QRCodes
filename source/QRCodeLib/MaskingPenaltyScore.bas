Attribute VB_Name = "MaskingPenaltyScore"
'----------------------------------------------------------------------------------------
' �}�X�N���ꂽ�V���{���̎��_�]��
'----------------------------------------------------------------------------------------
Option Private Module
Option Explicit

'----------------------------------------------------------------------------------------
' (�T�v)
'  �}�X�N�p�^�[�����_�̍��v��Ԃ��܂��B
'----------------------------------------------------------------------------------------
Public Function CalcTotal(ByRef moduleMatrix() As Variant) As Long

    Dim total   As Long
    Dim penalty As Long
    
    penalty = CalcAdjacentModulesInSameColor(moduleMatrix)
    total = total + penalty

    penalty = CalcBlockOfModulesInSameColor(moduleMatrix)
    total = total + penalty

    penalty = CalcModuleRatio(moduleMatrix)
    total = total + penalty

    penalty = CalcProportionOfDarkModules(moduleMatrix)
    total = total + penalty

    CalcTotal = total

End Function


'----------------------------------------------------------------------------------------
' (�T�v)
'  �s�^��̓��F�אڃ��W���[���p�^�[���̎��_���v�Z���܂��B
'----------------------------------------------------------------------------------------
Private Function CalcAdjacentModulesInSameColor(ByRef moduleMatrix() As Variant) As Long
    
    Dim penalty As Integer
    penalty = 0

    penalty = penalty + CalcAdjacentModulesInRowInSameColor(moduleMatrix)
    penalty = penalty + CalcAdjacentModulesInRowInSameColor(MatrixRotate90(moduleMatrix))

    CalcAdjacentModulesInSameColor = penalty

End Function

'----------------------------------------------------------------------------------------
' (�T�v)
'  �s�̓��F�אڃ��W���[���p�^�[���̎��_���v�Z���܂��B
'----------------------------------------------------------------------------------------
Private Function CalcAdjacentModulesInRowInSameColor(ByRef moduleMatrix() As Variant) As Long

    Dim penalty As Long
    penalty = 0

    Dim r As Long
    Dim c As Long
    Dim cnt As Long
    
    For r = 0 To UBound(moduleMatrix)
        cnt = 1

        For c = 0 To UBound(moduleMatrix(r)) - 1
            If (moduleMatrix(r)(c) > 0) = (moduleMatrix(r)(c + 1) > 0) Then
                cnt = cnt + 1
            Else
                If cnt >= 5 Then
                    penalty = penalty + (3 + (cnt - 5))
                End If

                cnt = 1
            End If
        Next

        If cnt >= 5 Then
            penalty = penalty + (3 + (cnt - 5))
        End If

    Next

    CalcAdjacentModulesInRowInSameColor = penalty

End Function

'----------------------------------------------------------------------------------------
' (�T�v)
'  2x2�̓��F���W���[���p�^�[���̎��_���v�Z���܂��B
'----------------------------------------------------------------------------------------
Private Function CalcBlockOfModulesInSameColor(ByRef moduleMatrix() As Variant) As Long

    Dim penalty     As Long
    Dim isSameColor As Boolean
    Dim r           As Long
    Dim c           As Long
    Dim tmp         As Boolean

    For r = 0 To UBound(moduleMatrix) - 1
        For c = 0 To UBound(moduleMatrix(r)) - 1
            tmp = moduleMatrix(r)(c) > 0
            isSameColor = True
            
            isSameColor = isSameColor And (moduleMatrix(r + 0)(c + 1) > 0 = tmp)
            isSameColor = isSameColor And (moduleMatrix(r + 1)(c + 0) > 0 = tmp)
            isSameColor = isSameColor And (moduleMatrix(r + 1)(c + 1) > 0 = tmp)
    
            If isSameColor Then
                penalty = penalty + 3
            End If
        Next
    Next
    
    CalcBlockOfModulesInSameColor = penalty

End Function

'----------------------------------------------------------------------------------------
' (�T�v)
'  �s�^��ɂ�����1 : 1 : 3 : 1 : 1 �䗦�p�^�[���̎��_���v�Z���܂��B
'----------------------------------------------------------------------------------------
Private Function CalcModuleRatio(ByRef moduleMatrix() As Variant) As Long
    
    Dim moduleMatrixTemp() As Variant
    moduleMatrixTemp = QuietZone.Place(moduleMatrix)

    Dim penalty As Integer
    penalty = 0
    
    penalty = penalty + CalcModuleRatioInRow(moduleMatrixTemp)
    penalty = penalty + CalcModuleRatioInRow(MatrixRotate90(moduleMatrixTemp))
    
    CalcModuleRatio = penalty

End Function


'----------------------------------------------------------------------------------------
' (�T�v)
'  �s��1 : 1 : 3 : 1 : 1 �䗦�̃p�^�[����]�����A���_��Ԃ��܂��B
'----------------------------------------------------------------------------------------
Private Function CalcModuleRatioInRow(ByRef moduleMatrix() As Variant) As Long

    Dim penalty As Long
    penalty = 0

    Dim r As Long
    Dim c As Long
    Dim cols() As Long
    Dim startIndexes  As Collection
    
    Dim i        As Long
    Dim idx      As Long
    Dim modRatio As ModuleRatio
    
    For r = 0 To UBound(moduleMatrix)
        cols = moduleMatrix(r)
        Set startIndexes = New Collection

        Call startIndexes.Add(0)

        For c = 0 To UBound(cols) - 2
            If cols(c) > 0 And cols(c + 1) <= 0 Then
                Call startIndexes.Add(c + 1)
            End If
        Next

        For i = 1 To startIndexes.Count
            idx = startIndexes(i)
            Set modRatio = New ModuleRatio

            Do While idx <= UBound(cols)
                If cols(idx) > 0 Then Exit Do
                modRatio.PreLightRatio4 = modRatio.PreLightRatio4 + 1
                idx = idx + 1
            Loop

            Do While idx <= UBound(cols)
                If cols(idx) <= 0 Then Exit Do
                modRatio.PreDarkRatio1 = modRatio.PreDarkRatio1 + 1
                idx = idx + 1
            Loop

            Do While idx <= UBound(cols)
                If cols(idx) > 0 Then Exit Do
                modRatio.PreLightRatio1 = modRatio.PreLightRatio1 + 1
                idx = idx + 1
            Loop

            Do While idx <= UBound(cols)
                If cols(idx) <= 0 Then Exit Do
                modRatio.CenterDarkRatio3 = modRatio.CenterDarkRatio3 + 1
                idx = idx + 1
            Loop

            Do While idx <= UBound(cols)
                If cols(idx) > 0 Then Exit Do
                modRatio.FolLightRatio1 = modRatio.FolLightRatio1 + 1
                idx = idx + 1
            Loop

            Do While idx <= UBound(cols)
                If cols(idx) <= 0 Then Exit Do
                modRatio.FolDarkRatio1 = modRatio.FolDarkRatio1 + 1
                idx = idx + 1
            Loop

            Do While idx <= UBound(cols)
                If cols(idx) > 0 Then Exit Do
                modRatio.FolLightRatio4 = modRatio.FolLightRatio4 + 1
                idx = idx + 1
            Loop

            If modRatio.PenaltyImposed() Then
                penalty = penalty + 40
            End If

        Next
    Next

    CalcModuleRatioInRow = penalty
    
End Function

'----------------------------------------------------------------------------------------
' (�T�v)
'  �S�̂ɑ΂���Ã��W���[���̐�߂銄���ɂ��Ď��_���v�Z���܂��B
'----------------------------------------------------------------------------------------
Private Function CalcProportionOfDarkModules(ByRef moduleMatrix() As Variant) As Long

    Dim darkCount As Long

    Dim r As Long
    Dim c As Long
    
    For r = 0 To UBound(moduleMatrix)
        For c = 0 To UBound(moduleMatrix(r))
            If moduleMatrix(r)(c) > 0 Then
                darkCount = darkCount + 1
            End If
        Next
    Next
    
    Dim tmp As Long
    tmp = CLng(Int((darkCount / (UBound(moduleMatrix) + 1) ^ 2) * 100))
    tmp = Abs(tmp - 50)
    tmp = (tmp + 4) \ 5
    
    CalcProportionOfDarkModules = tmp * 10

End Function

'----------------------------------------------------------------------------------------
' (�T�v)
'  ����90�x��]�����z���Ԃ��܂��B
'----------------------------------------------------------------------------------------
Private Function MatrixRotate90(arg() As Variant) As Variant()

    Dim ret() As Variant
    ReDim ret(UBound(arg(0)))

    Dim i As Long
    Dim j As Long
    Dim cols() As Long
    
    For i = 0 To UBound(ret)
        ReDim cols(UBound(arg))
        ret(i) = cols
    Next
    
    Dim k As Long
    k = UBound(ret)
    
    For i = 0 To UBound(ret)
        For j = 0 To UBound(ret(i))
            ret(i)(j) = arg(j)(k - i)
        Next
    Next

    MatrixRotate90 = ret

End Function
