VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6330
   ClientLeft      =   6000
   ClientTop       =   1000
   ClientWidth     =   4770
   LinkTopic       =   "Form2"
   ScaleHeight     =   6330
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   250
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   250
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   4450
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Research Version"
      Height          =   250
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   4450
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Production Version"
      Height          =   250
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   4450
   End
   Begin VB.Image Image1 
      Height          =   2290
      Left            =   240
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   2290
   End
   Begin VB.Image Image2 
      Height          =   2290
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   2290
   End
   Begin VB.Shape Shape1 
      Height          =   6270
      Left            =   0
      Top             =   0
      Width           =   4760
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      ForeColor       =   &H00404040&
      Height          =   1090
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4570
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form2.Check1.Value = 0
Form1.Visible = True
Form2.Visible = False
End Sub

Private Sub Command2_Click()
Form2.Check1.Value = 1
Form1.Visible = True
Form2.Visible = False
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\Files\logo.jpg")
Image2.Picture = LoadPicture(App.Path & "\Files\NLM.jpg")
Label1.Caption = "Finite Difference-Based Heat Transfer Analysis for Twin Belt Caster"
Label1.FontSize = 16
Label1.FontBold = True
End Sub

