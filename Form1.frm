VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6285
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Text            =   "Hello!"
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "X"
      Height          =   250
      Left            =   6600
      TabIndex        =   0
      Top             =   160
      Width           =   250
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShapeTheForm As clsTransForm

Private Sub cmdExit_Click()
   
Unload Form1

End Sub

Private Sub Form_Load()
Set ShapeTheForm = New clsTransForm

ShapeTheForm.ShapeMe RGB(255, 255, 255), True, Form1 'do the real work

Form2.Visible = True

Form1.Show
Text1.SetFocus

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

ShapeTheForm.DragForm Me, Button 'move the form

End Sub


Private Sub Form_Unload(Cancel As Integer)

Set ShapeTheForm = Nothing
Set Form1 = Nothing

End Sub
