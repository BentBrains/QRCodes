Attribute VB_Name = "Module"
'----------------------------------------------------------------------------------------
' ���W���[��
'----------------------------------------------------------------------------------------
Option Private Module
Option Explicit

'---------------------------------------------------------------------------
' (�T�v)
'  �P�ӂ̃��W���[������Ԃ��܂��B
'---------------------------------------------------------------------------
Public Function GetNumModulesPerSide(ByVal ver As Long) As Long

    GetNumModulesPerSide = 17 + ver * 4
    
End Function
