VERSION 5.00
Object = "{89D94A1E-DB65-4469-AFB5-D54C6F6B7639}#1.1#0"; "QRCodeAX.ocx"
Begin VB.Form Form1 
   Caption         =   "QRCodeAX TestForm"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbECL 
      Height          =   315
      ItemData        =   "QRCodeTestForm.frx":0000
      Left            =   8040
      List            =   "QRCodeTestForm.frx":0010
      TabIndex        =   7
      Top             =   3720
      Width           =   2055
   End
   Begin QRCodeAX.QR QR1 
      Height          =   4335
      Left            =   600
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7646
      DataString      =   "PRDEL"
      BackRGB         =   "#FFAA88"
      ShowBorder      =   -1  'True
   End
   Begin VB.CommandButton cmdRecalc 
      Caption         =   "Recalc"
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CheckBox chkBorder 
      Caption         =   "Border"
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtData 
      Height          =   1815
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   5520
      Width           =   10335
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "CLS"
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy >>>"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "DataString"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   5160
      Width           =   2895
   End
   Begin VB.Image imgA 
      Height          =   1695
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbECL_Click()
    QR1.ErrorCorrectionLevel = cmbECL.ListIndex
    Debug.Print QR1.ErrorCorrectionLevel
    QR1.Recalc
End Sub


Private Sub cmdCls_Click()
    QR1.Cls
End Sub

Private Sub cmdCopy_Click()
    imgA.Picture = QR1.Picture
End Sub

Private Sub cmdRecalc_Click()
    QR1.Recalc
End Sub

Private Sub Form_Load()
    cmbECL.ListIndex = 1
End Sub

Private Sub chkBorder_Click()
    QR1.ShowBorder = chkBorder
End Sub

Private Sub QR1_AfterReCalc()
    Debug.Print "AfterReCalc"
End Sub

Private Sub QR1_Click()
    Debug.Print "Click"
End Sub

Private Sub QR1_DblClick()
    MsgBox "Double Click", vbInformation
End Sub

Private Sub txtData_Change()
    'QR1.Datastring is default property so it does not have to be specified.
    QR1 = txtData
End Sub
