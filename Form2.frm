VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00008000&
   Caption         =   "Form2"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2070
   LinkTopic       =   "Form2"
   ScaleHeight     =   1725
   ScaleWidth      =   2070
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   27
      ImageHeight     =   47
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":02E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picButton 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   840
      Picture         =   "Form2.frx":05B0
      ScaleHeight     =   825
      ScaleWidth      =   450
      TabIndex        =   0
      Top             =   240
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "Press the Triangle"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShapeThePB As clsTransForm

Private Sub Form_Load()
Set ShapeThePB = New clsTransForm

ShapeThePB.ShapeMe RGB(255, 255, 255), True, , picButton

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set ShapeThePB = Nothing
Set Form2 = Nothing

End Sub

Private Sub picButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Set picButton.Picture = ImageList1.ListImages(2).Picture

End Sub

Private Sub picButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Set picButton.Picture = ImageList1.ListImages(1).Picture

MsgBox "You pressed me."

End Sub
