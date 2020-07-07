VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   ClientHeight    =   8160
   ClientLeft      =   10830
   ClientTop       =   10
   ClientWidth     =   5720
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   5720
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   370
      Left            =   960
      TabIndex        =   74
      Top             =   7200
      Width           =   4210
   End
   Begin VB.ComboBox Combo1 
      Height          =   280
      Left            =   960
      TabIndex        =   73
      Top             =   120
      Width           =   1570
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Submit"
      Height          =   370
      Left            =   3840
      TabIndex        =   72
      Top             =   6720
      Width           =   1330
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   1560
      TabIndex        =   71
      Top             =   840
      Width           =   970
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   960
      TabIndex        =   68
      Top             =   480
      Width           =   1570
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   44
      Left            =   3720
      TabIndex        =   67
      Top             =   6150
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   43
      Left            =   2400
      TabIndex        =   66
      Top             =   6150
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   42
      Left            =   1080
      TabIndex        =   65
      Top             =   6150
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   41
      Left            =   3720
      TabIndex        =   64
      Top             =   5865
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   40
      Left            =   2400
      TabIndex        =   63
      Top             =   5865
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   39
      Left            =   1080
      TabIndex        =   62
      Top             =   5865
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   38
      Left            =   3720
      TabIndex        =   61
      Top             =   5580
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   37
      Left            =   2400
      TabIndex        =   60
      Top             =   5580
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   36
      Left            =   1080
      TabIndex        =   59
      Top             =   5580
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   35
      Left            =   3720
      TabIndex        =   58
      Top             =   5295
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   34
      Left            =   2400
      TabIndex        =   57
      Top             =   5295
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   33
      Left            =   1080
      TabIndex        =   56
      Top             =   5295
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   32
      Left            =   3720
      TabIndex        =   55
      Top             =   5010
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   31
      Left            =   2400
      TabIndex        =   54
      Top             =   5010
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   30
      Left            =   1080
      TabIndex        =   53
      Top             =   5010
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   29
      Left            =   3720
      TabIndex        =   52
      Top             =   4725
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   28
      Left            =   2400
      TabIndex        =   51
      Top             =   4725
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   27
      Left            =   1080
      TabIndex        =   50
      Top             =   4725
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   26
      Left            =   3720
      TabIndex        =   49
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   25
      Left            =   2400
      TabIndex        =   48
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   24
      Left            =   1080
      TabIndex        =   47
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   23
      Left            =   3720
      TabIndex        =   46
      Top             =   4155
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   22
      Left            =   2400
      TabIndex        =   45
      Top             =   4155
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   21
      Left            =   1080
      TabIndex        =   44
      Top             =   4155
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   20
      Left            =   3720
      TabIndex        =   43
      Top             =   3870
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   19
      Left            =   2400
      TabIndex        =   42
      Top             =   3870
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   18
      Left            =   1080
      TabIndex        =   41
      Top             =   3870
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   17
      Left            =   3720
      TabIndex        =   40
      Top             =   3585
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   16
      Left            =   2400
      TabIndex        =   39
      Top             =   3585
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   15
      Left            =   1080
      TabIndex        =   38
      Top             =   3585
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   14
      Left            =   3720
      TabIndex        =   37
      Top             =   3300
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   13
      Left            =   2400
      TabIndex        =   36
      Top             =   3300
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   12
      Left            =   1080
      TabIndex        =   35
      Top             =   3300
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   11
      Left            =   3720
      TabIndex        =   34
      Top             =   3015
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   10
      Left            =   2400
      TabIndex        =   33
      Top             =   3015
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   9
      Left            =   1080
      TabIndex        =   32
      Top             =   3015
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   8
      Left            =   3720
      TabIndex        =   31
      Top             =   2730
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   7
      Left            =   2400
      TabIndex        =   30
      Top             =   2730
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   6
      Left            =   1080
      TabIndex        =   29
      Top             =   2730
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   5
      Left            =   3720
      TabIndex        =   28
      Top             =   2445
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   4
      Left            =   2400
      TabIndex        =   27
      Top             =   2445
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   26
      Top             =   2445
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   25
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   24
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      Height          =   370
      Left            =   2400
      TabIndex        =   23
      Top             =   6720
      Width           =   1330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   370
      Left            =   960
      TabIndex        =   22
      Top             =   6720
      Width           =   1330
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   21
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   970
   End
   Begin VB.Image Image2 
      Height          =   1210
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1210
   End
   Begin VB.Shape Shape1 
      Height          =   8160
      Left            =   -2640
      Top             =   0
      Width           =   8350
   End
   Begin VB.Image Image1 
      Height          =   1210
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1210
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Caster Length (m)"
      Height          =   380
      Left            =   120
      TabIndex        =   70
      Top             =   960
      Width           =   1340
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Location"
      Height          =   380
      Left            =   -1080
      TabIndex        =   69
      Top             =   600
      Width           =   1940
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Bot(W/m^2)"
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Top(W/m^2)"
      Height          =   375
      Left            =   2640
      TabIndex        =   19
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Xpos(m)"
      Height          =   375
      Left            =   1320
      TabIndex        =   18
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "3"
      Height          =   255
      Index           =   15
      Left            =   600
      TabIndex        =   17
      Top             =   2740
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "4"
      Height          =   255
      Index           =   14
      Left            =   600
      TabIndex        =   16
      Top             =   3030
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "5"
      Height          =   255
      Index           =   13
      Left            =   600
      TabIndex        =   15
      Top             =   3320
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "6"
      Height          =   255
      Index           =   12
      Left            =   600
      TabIndex        =   14
      Top             =   3610
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "7"
      Height          =   255
      Index           =   11
      Left            =   600
      TabIndex        =   13
      Top             =   3900
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "8"
      Height          =   255
      Index           =   10
      Left            =   600
      TabIndex        =   12
      Top             =   4190
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "9"
      Height          =   255
      Index           =   9
      Left            =   600
      TabIndex        =   11
      Top             =   4480
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "10"
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   10
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "11"
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   9
      Top             =   5060
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "12"
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   8
      Top             =   5350
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "13"
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   7
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "14"
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   6
      Top             =   5930
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "15"
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   5
      Top             =   6225
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Top             =   2450
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "No Flux Data"
      Height          =   380
      Left            =   -480
      TabIndex        =   1
      Top             =   1320
      Width           =   1940
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "File Name"
      Height          =   380
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   860
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Private Sub Command1_Click()
Command3.Enabled = True
Form3.Text2.Enabled = False
Form3.Text4.Enabled = False
Form3.Text5.Enabled = False

For i = 0 To 44
Form3.Text3(i).Enabled = False
Next i



Form3.Text2.Text = ""
Form3.Text4.Text = ""
Form3.Text5.Text = ""

For i = 0 To 44
Form3.Text3(i).Text = ""
Next i


Dim newfile2
Dim iFileNo2 As Integer
Dim pp
iFileNo2 = FreeFile



Open App.Path & "/Locations/" & Form3.Combo1.Text For Input As #iFileNo2
Line Input #iFileNo2, pp
Text2.Text = pp
Line Input #iFileNo2, pp
Text4.Text = pp
Line Input #iFileNo2, pp
Text5.Text = pp


p1 = Text5.Text * 3 - 1

For i = 0 To p1
    Line Input #iFileNo2, pp
    Text3(i).Text = pp
Next i


Close #iFileNo2

If Dir(App.Path & "/Locations/" & Form3.Combo1.Text) <> "" Then
    FileCopy App.Path & "/Locations/" & Form3.Combo1.Text, App.Path & "/Temporary/" & Form3.Combo1.Text
End If






End Sub

Private Sub Command2_Click()
Text2.Enabled = True
Text4.Enabled = True
Text5.Enabled = True

For i = 0 To 44
Text3(i).Enabled = True
Next i
End Sub

Private Sub Command3_Click()
Dim newfile3



Dim iFileNo3 As Integer
iFileNo3 = FreeFile
Open App.Path & "/Temporary/BC.DAT" For Output As #iFileNo3


Print #iFileNo3, Form3.Text2.Text
Print #iFileNo3, Form3.Text4.Text
Print #iFileNo3, Form3.Text5.Text

PP3 = 3 * Form3.Text5.Text
For i = 0 To PP3
Print #iFileNo3, Form3.Text3(i).Text
Next i




Close #iFileNo3
Unload Me
Unload Form4

Form1.Check1.Value = 1



End Sub

Private Sub Command4_Click()
Unload Form3
Unload Form4
End Sub

Private Sub Form_Load()
Text2.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Image1.Picture = LoadPicture(App.Path & "\Files\logo.jpg")
Image2.Picture = LoadPicture(App.Path & "\Files\NLM.jpg")
For i = 0 To 44
Text3(i).Enabled = False
Next i

Command3.Enabled = False
End Sub


