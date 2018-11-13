VERSION 5.00
Begin VB.UserControl QR 
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   ScaleHeight     =   2235
   ScaleWidth      =   2865
   ToolboxBitmap   =   "QR.ctx":0000
   Begin VB.Image QRImg 
      Appearance      =   0  'Flat
      Height          =   1935
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "QR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'http://forums.codeguru.com/showthread.php?845-OCX-Version-compatibility
'https://stackoverflow.com/questions/53244605/making-qrcode-activex-control-for-ms-access-control-source-property

'Public Enums
Public Enum eAppearance
    Flat
    Sunken
End Enum

Public Enum eBackStyle
    Transparent
    Solid
End Enum

Public Enum eECR                'Error Correction Level
    L
    M
    Q
    H
End Enum

'Default Property Values:
Const m_def_ShowBorder = 0
Const m_def_ErrorCorrectionLevel = 1
Const m_def_ModuleSize = 1
Const m_def_BackRGB = "#FFFFFF"
Const m_def_ForeRGB = "#000000"
Const m_def_ByteModeCharsetName = "UTF-8"
Const m_def_DataString = ""

'Property Variables:
Dim m_ShowBorder As Boolean
Dim m_ErrorCorrectionLevel As eECR
Dim m_ModuleSize As Long
Dim m_BackRGB As String
Dim m_ForeRGB As String
Dim m_ByteModeCharsetName As String
Dim m_DataString As String

'Event Declarations:
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Double Click Event"
Attribute DblClick.VB_UserMemId = -601
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Single Click Event"
Attribute Click.VB_UserMemId = -600
Event AfterRecalc()
Attribute AfterRecalc.VB_Description = "This event is fired after QRCode was Recalculated"

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get DataString() As String
Attribute DataString.VB_Description = "Data to be encoded to QR Code"
Attribute DataString.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute DataString.VB_UserMemId = 0
Attribute DataString.VB_MemberFlags = "224"
    DataString = m_DataString
End Property

Public Property Let DataString(ByVal New_DataString As String)
    If Len(New_DataString) > 4096 Then
        MsgBox "Maximum DataString size is 4096 Bytes" & vbCrLf & _
        "The QRCode will display no data", vbExclamation
    End If
    m_DataString = New_DataString
    PropertyChanged "DataString"
    Recalc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Recalc() As Variant
Attribute Recalc.VB_Description = "Recalculates the QRCode"
    'Generate QR Code and display it
    If m_DataString = "" Or Len(m_DataString) > 4096 Then
        QRImg.Picture = Nothing
        GoTo Finally_
    End If
    
    On Error GoTo Catch_
    Dim sbls As Symbols
    Set sbls = CreateSymbols(m_ErrorCorrectionLevel, 40, False, ByteModeCharsetName)
    sbls.AppendString m_DataString

    Dim Pict As stdole.IPicture
    Set Pict = sbls(0).Get24bppImage(m_ModuleSize, m_ForeRGB, m_BackRGB)
    QRImg.Picture = Pict

Finally_:
    On Error GoTo 0
    RaiseEvent AfterRecalc
    Exit Function
    
Catch_:
    MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical, "QRCodeAX"
    Resume Finally_
End Function

Private Sub QRImg_Click()
    RaiseEvent Click
End Sub

Private Sub QRImg_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_DataString = m_def_DataString
    m_BackRGB = m_def_BackRGB
    m_ForeRGB = m_def_ForeRGB
    m_ByteModeCharsetName = m_def_ByteModeCharsetName
    m_ModuleSize = m_def_ModuleSize
    m_ShowBorder = m_def_ShowBorder
    m_ErrorCorrectionLevel = m_def_ErrorCorrectionLevel
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error GoTo Catch_
    m_DataString = PropBag.ReadProperty("DataString", m_def_DataString)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    QRImg.Appearance = PropBag.ReadProperty("Appearance", 0)
    QRImg.BorderStyle = PropBag.ReadProperty("ShowBorder", 0)
    m_BackRGB = PropBag.ReadProperty("BackRGB", m_def_BackRGB)
    m_ForeRGB = PropBag.ReadProperty("ForeRGB", m_def_ForeRGB)
    m_ByteModeCharsetName = PropBag.ReadProperty("ByteModeCharsetName", m_def_ByteModeCharsetName)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    m_ModuleSize = PropBag.ReadProperty("ModuleSize", m_def_ModuleSize)
    m_ShowBorder = PropBag.ReadProperty("ShowBorder", m_def_ShowBorder)
    m_ErrorCorrectionLevel = PropBag.ReadProperty("ErrorCorrectionLevel", m_def_ErrorCorrectionLevel)
Finally_:
    On Error GoTo 0
    Exit Sub
Catch_:
    MsgBox Err.Number & ": " & Err.Description
    Resume Finally_
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error GoTo Catch_
    Call PropBag.WriteProperty("DataString", m_DataString, m_def_DataString)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Appearance", QRImg.Appearance, 0)
    Call PropBag.WriteProperty("ShowBorder", QRImg.BorderStyle, 0)
    Call PropBag.WriteProperty("BackRGB", m_BackRGB, m_def_BackRGB)
    Call PropBag.WriteProperty("ForeRGB", m_ForeRGB, m_def_ForeRGB)
    Call PropBag.WriteProperty("ByteModeCharsetName", m_ByteModeCharsetName, m_def_ByteModeCharsetName)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("ModuleSize", m_ModuleSize, m_def_ModuleSize)
    Call PropBag.WriteProperty("ShowBorder", m_ShowBorder, m_def_ShowBorder)
    Call PropBag.WriteProperty("ErrorCorrectionLevel", m_ErrorCorrectionLevel, m_def_ErrorCorrectionLevel)
Finally_:
    On Error GoTo 0
    Exit Sub
Catch_:
    MsgBox Err.Number & ": " & Err.Description
    Resume Finally_
End Sub

Private Sub UserControl_Resize()
    'Make sure the image has aspect ratio 1:1 (square)
    UserControl.Width = UserControl.Height
    QRImg.Width = Width
    QRImg.Height = Width
End Sub

Public Sub Cls()
Attribute Cls.VB_Description = "Clear's the QRCode Picture"
    Set QRImg.Picture = Nothing
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=QRImg,QRImg,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "QRCode Picture data"
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Data"
    Set Picture = QRImg.Picture
End Property

Private Property Set Picture(ByVal New_Picture As Picture)  'I don't want this available
    Set QRImg.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=QRImg,QRImg,-1,Appearance
Public Property Get Appearance() As eAppearance
Attribute Appearance.VB_Description = "3D Appearance - Flat or Sunken"
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Appearance.VB_UserMemId = -520
    Appearance = QRImg.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As eAppearance)
    QRImg.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,#FFFFFF
Public Property Get BackRGB() As String
Attribute BackRGB.VB_Description = "QRCode BackColor, web format"
Attribute BackRGB.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackRGB = m_BackRGB
End Property

Public Property Let BackRGB(ByVal New_BackRGB As String)
    m_BackRGB = New_BackRGB
    PropertyChanged "BackRGB"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,#000000
Public Property Get ForeRGB() As String
Attribute ForeRGB.VB_Description = "QRCode ForeColor, web format"
Attribute ForeRGB.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeRGB = m_ForeRGB
End Property

Public Property Let ForeRGB(ByVal New_ForeRGB As String)
    m_ForeRGB = New_ForeRGB
    PropertyChanged "ForeRGB"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,UTF-8
Public Property Get ByteModeCharsetName() As String
Attribute ByteModeCharsetName.VB_Description = "Encoding UTF-8 or Shift-JIS"
Attribute ByteModeCharsetName.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ByteModeCharsetName = m_ByteModeCharsetName
End Property

Public Property Let ByteModeCharsetName(ByVal New_ByteModeCharsetName As String)
    m_ByteModeCharsetName = New_ByteModeCharsetName
    PropertyChanged "ByteModeCharsetName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Control's BackColor"
Attribute BackColor.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As eBackStyle
Attribute BackStyle.VB_Description = "Controls BackStyle - Transparent or Opaque"
Attribute BackStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackStyle.VB_UserMemId = -502
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As eBackStyle)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,5
Public Property Get ModuleSize() As Long
Attribute ModuleSize.VB_Description = "Pixels per Module. Higher value make it slower"
Attribute ModuleSize.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ModuleSize = m_ModuleSize
End Property

Public Property Let ModuleSize(ByVal New_ModuleSize As Long)
    'Limit the module size to 1-20
    If New_ModuleSize > 20 Then
        m_ModuleSize = 20
    ElseIf New_ModuleSize < 1 Then
        m_ModuleSize = 1
    Else
        m_ModuleSize = New_ModuleSize
    End If
    PropertyChanged "ModuleSize"
End Property
Public Property Get ShowBorder() As Boolean
Attribute ShowBorder.VB_Description = "Show border around the QRCode"
Attribute ShowBorder.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ShowBorder = m_ShowBorder
End Property

Public Property Let ShowBorder(ByVal New_ShowBorder As Boolean)
    m_ShowBorder = New_ShowBorder
    If m_ShowBorder Then
        QRImg.BorderStyle = 1
    Else
        QRImg.BorderStyle = 0
    End If
    PropertyChanged "ShowBorder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=26,0,0,1
Public Property Get ErrorCorrectionLevel() As eECR
Attribute ErrorCorrectionLevel.VB_Description = "0=L(7%), 1=M(15%), 2=Q(25%), 3=H(30%)"
Attribute ErrorCorrectionLevel.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ErrorCorrectionLevel = m_ErrorCorrectionLevel
End Property

Public Property Let ErrorCorrectionLevel(ByVal New_ErrorCorrectionLevel As eECR)
    m_ErrorCorrectionLevel = New_ErrorCorrectionLevel
    PropertyChanged "ErrorCorrectionLevel"
End Property

