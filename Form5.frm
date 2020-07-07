VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   1850
   ClientLeft      =   4800
   ClientTop       =   4000
   ClientWidth     =   7080
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1850
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   370
      Left            =   4200
      TabIndex        =   1
      Top             =   1200
      Width           =   1330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   370
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   1330
   End
   Begin VB.Shape Shape1 
      Height          =   1830
      Left            =   0
      Top             =   0
      Width           =   7080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Height          =   490
      Left            =   80
      TabIndex        =   2
      Top             =   360
      Width           =   6930
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Form5
    Form5.Visible = False
    Form1.Check2.Value = 1
    Form1.Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Unload Form5
    Form5.Visible = False
    Form1.Text1(4).SetFocus
End Sub

