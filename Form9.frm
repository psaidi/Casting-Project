VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   0  'None
   Caption         =   "Form9"
   ClientHeight    =   1850
   ClientLeft      =   4800
   ClientTop       =   4000
   ClientWidth     =   7080
   LinkTopic       =   "Form9"
   ScaleHeight     =   1850
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   370
      Left            =   4120
      TabIndex        =   1
      Top             =   1200
      Width           =   1330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   370
      Left            =   1480
      TabIndex        =   0
      Top             =   1200
      Width           =   1330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Height          =   490
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   6930
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Form9
    Form9.Visible = False
    Form7.Check2.Value = 0
    Unload Form7
End Sub

Private Sub Command2_Click()
    Unload Form9
    Form9.Visible = False
End Sub
