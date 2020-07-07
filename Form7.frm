VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   4670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7770
   LinkTopic       =   "Form7"
   ScaleHeight     =   4670
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo3 
      Height          =   280
      Left            =   1080
      TabIndex        =   96
      Top             =   2280
      Width           =   1330
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
      Height          =   370
      Left            =   5280
      TabIndex        =   95
      Top             =   3360
      Width           =   1810
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Import-Check-Create File"
      Height          =   370
      Left            =   600
      TabIndex        =   94
      Top             =   3360
      Width           =   4330
   End
   Begin VB.ComboBox Combo2 
      Height          =   280
      Left            =   1080
      TabIndex        =   92
      Top             =   1920
      Width           =   1330
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   250
      Left            =   6720
      TabIndex        =   90
      Top             =   2040
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   250
      Left            =   7320
      TabIndex        =   89
      Top             =   2040
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Frame Frame2 
      Caption         =   "Temperature vs Solid Fraction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4570
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   5410
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   0
         Left            =   600
         TabIndex        =   57
         Top             =   1320
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   1
         Left            =   600
         TabIndex        =   56
         Top             =   1590
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   2
         Left            =   600
         TabIndex        =   55
         Top             =   1870
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   3
         Left            =   600
         TabIndex        =   54
         Top             =   2150
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   4
         Left            =   600
         TabIndex        =   53
         Top             =   2430
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   5
         Left            =   600
         TabIndex        =   52
         Top             =   2720
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   6
         Left            =   600
         TabIndex        =   51
         Top             =   3000
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   7
         Left            =   600
         TabIndex        =   50
         Top             =   3280
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   8
         Left            =   600
         TabIndex        =   49
         Top             =   3560
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   9
         Left            =   600
         TabIndex        =   48
         Top             =   3840
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   10
         Left            =   3360
         TabIndex        =   47
         Top             =   1320
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   11
         Left            =   3360
         TabIndex        =   46
         Top             =   1590
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   12
         Left            =   3360
         TabIndex        =   45
         Top             =   1870
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   13
         Left            =   3360
         TabIndex        =   44
         Top             =   2150
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   14
         Left            =   3360
         TabIndex        =   43
         Top             =   2430
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   15
         Left            =   3360
         TabIndex        =   42
         Top             =   2720
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   16
         Left            =   3360
         TabIndex        =   41
         Top             =   3000
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   17
         Left            =   3360
         TabIndex        =   40
         Top             =   3280
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   18
         Left            =   3360
         TabIndex        =   39
         Top             =   3560
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   19
         Left            =   3360
         TabIndex        =   38
         Top             =   3840
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   20
         Left            =   1560
         TabIndex        =   37
         Top             =   1320
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   21
         Left            =   1560
         TabIndex        =   36
         Top             =   1590
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   22
         Left            =   1560
         TabIndex        =   35
         Top             =   1870
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   23
         Left            =   1560
         TabIndex        =   34
         Top             =   2150
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   24
         Left            =   1560
         TabIndex        =   33
         Top             =   2430
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   25
         Left            =   1560
         TabIndex        =   32
         Top             =   2720
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   26
         Left            =   1560
         TabIndex        =   31
         Top             =   3000
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   27
         Left            =   1560
         TabIndex        =   30
         Top             =   3280
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   28
         Left            =   1560
         TabIndex        =   29
         Top             =   3560
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   29
         Left            =   1560
         TabIndex        =   28
         Top             =   3840
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   30
         Left            =   4320
         TabIndex        =   27
         Top             =   1320
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   31
         Left            =   4320
         TabIndex        =   26
         Top             =   1590
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   32
         Left            =   4320
         TabIndex        =   25
         Top             =   1870
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   33
         Left            =   4320
         TabIndex        =   24
         Top             =   2150
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   34
         Left            =   4320
         TabIndex        =   23
         Top             =   2430
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   35
         Left            =   4320
         TabIndex        =   22
         Top             =   2720
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   36
         Left            =   4320
         TabIndex        =   21
         Top             =   3000
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   37
         Left            =   4320
         TabIndex        =   20
         Top             =   3280
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   38
         Left            =   4320
         TabIndex        =   19
         Top             =   3560
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   39
         Left            =   4320
         TabIndex        =   18
         Top             =   3840
         Width           =   900
      End
      Begin VB.ComboBox Combo1 
         Height          =   280
         Left            =   1920
         TabIndex        =   17
         Top             =   360
         Width           =   1450
      End
      Begin VB.Label Label7 
         Caption         =   "f_s"
         Height          =   370
         Left            =   4680
         TabIndex        =   87
         Top             =   960
         Width           =   370
      End
      Begin VB.Label Label6 
         Caption         =   "Temp. (C)"
         Height          =   370
         Left            =   3480
         TabIndex        =   86
         Top             =   960
         Width           =   850
      End
      Begin VB.Label Label5 
         Caption         =   "f_s"
         Height          =   370
         Left            =   1920
         TabIndex        =   85
         Top             =   960
         Width           =   370
      End
      Begin VB.Label Label4 
         Caption         =   "Temp. (C)"
         Height          =   370
         Left            =   720
         TabIndex        =   84
         Top             =   960
         Width           =   850
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "3"
         Height          =   260
         Index           =   15
         Left            =   120
         TabIndex        =   78
         Top             =   1900
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "4"
         Height          =   260
         Index           =   14
         Left            =   120
         TabIndex        =   77
         Top             =   2190
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "5"
         Height          =   260
         Index           =   13
         Left            =   120
         TabIndex        =   76
         Top             =   2480
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "6"
         Height          =   260
         Index           =   12
         Left            =   120
         TabIndex        =   75
         Top             =   2770
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "7"
         Height          =   260
         Index           =   11
         Left            =   120
         TabIndex        =   74
         Top             =   3060
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "8"
         Height          =   260
         Index           =   10
         Left            =   120
         TabIndex        =   73
         Top             =   3350
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "9"
         Height          =   260
         Index           =   9
         Left            =   120
         TabIndex        =   72
         Top             =   3640
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "10"
         Height          =   260
         Index           =   8
         Left            =   120
         TabIndex        =   71
         Top             =   3930
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "2"
         Height          =   260
         Index           =   1
         Left            =   120
         TabIndex        =   70
         Top             =   1610
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   260
         Index           =   0
         Left            =   120
         TabIndex        =   69
         Top             =   1320
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "13"
         Height          =   260
         Index           =   2
         Left            =   2880
         TabIndex        =   68
         Top             =   1900
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "14"
         Height          =   260
         Index           =   3
         Left            =   2880
         TabIndex        =   67
         Top             =   2190
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "15"
         Height          =   260
         Index           =   4
         Left            =   2880
         TabIndex        =   66
         Top             =   2480
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "16"
         Height          =   260
         Index           =   5
         Left            =   2880
         TabIndex        =   65
         Top             =   2770
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "17"
         Height          =   260
         Index           =   6
         Left            =   2880
         TabIndex        =   64
         Top             =   3060
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "18"
         Height          =   260
         Index           =   7
         Left            =   2880
         TabIndex        =   63
         Top             =   3350
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "19"
         Height          =   260
         Index           =   16
         Left            =   2880
         TabIndex        =   62
         Top             =   3640
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "20"
         Height          =   260
         Index           =   17
         Left            =   2880
         TabIndex        =   61
         Top             =   3930
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "12"
         Height          =   260
         Index           =   18
         Left            =   2880
         TabIndex        =   60
         Top             =   1610
         Width           =   380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "11"
         Height          =   260
         Index           =   19
         Left            =   2880
         TabIndex        =   59
         Top             =   1320
         Width           =   380
      End
      Begin VB.Label Label2 
         Caption         =   "Number of Data Points"
         Height          =   370
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   1810
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Alloy Constants"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1570
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   7570
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   5
         Left            =   5760
         TabIndex        =   83
         Top             =   1080
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   4
         Left            =   5760
         TabIndex        =   82
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   3
         Left            =   5760
         TabIndex        =   81
         Top             =   360
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   2
         Left            =   2400
         TabIndex        =   80
         Top             =   1080
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   1
         Left            =   2400
         TabIndex        =   79
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   0
         Left            =   2400
         TabIndex        =   4
         Top             =   360
         Width           =   1000
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alloy Number"
         Height          =   380
         Index           =   0
         Left            =   720
         TabIndex        =   15
         Top             =   360
         Width           =   1570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Latent Heat"
         Height          =   380
         Index           =   1
         Left            =   4080
         TabIndex        =   14
         Top             =   720
         Width           =   1570
      End
      Begin VB.Label Label1 
         Caption         =   "J/Kg"
         Height          =   380
         Index           =   2
         Left            =   6840
         TabIndex        =   13
         Top             =   720
         Width           =   610
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Density"
         Height          =   380
         Index           =   3
         Left            =   4080
         TabIndex        =   12
         Top             =   360
         Width           =   1570
      End
      Begin VB.Label Label1 
         Caption         =   "Kg/m^3"
         Height          =   380
         Index           =   4
         Left            =   6840
         TabIndex        =   11
         Top             =   360
         Width           =   610
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Heat Capacity"
         Height          =   380
         Index           =   6
         Left            =   4080
         TabIndex        =   10
         Top             =   1080
         Width           =   1570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Thermal Conductivity Solid"
         Height          =   380
         Index           =   7
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   2050
      End
      Begin VB.Label Label1 
         Caption         =   "J/Kg.C"
         Height          =   380
         Index           =   8
         Left            =   6840
         TabIndex        =   8
         Top             =   1080
         Width           =   610
      End
      Begin VB.Label Label1 
         Caption         =   "W/m.C"
         Height          =   380
         Index           =   9
         Left            =   3480
         TabIndex        =   7
         Top             =   720
         Width           =   610
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Thermal Conductivity Liquid"
         Height          =   380
         Index           =   10
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   2050
      End
      Begin VB.Label Label1 
         Caption         =   "W/m.C"
         Height          =   380
         Index           =   11
         Left            =   3480
         TabIndex        =   5
         Top             =   1080
         Width           =   610
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create files"
      Enabled         =   0   'False
      Height          =   370
      Left            =   5760
      TabIndex        =   2
      Top             =   3840
      Width           =   1810
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   370
      Left            =   5760
      TabIndex        =   1
      Top             =   3240
      Width           =   1810
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   370
      Left            =   5760
      TabIndex        =   0
      Top             =   4440
      Width           =   1810
   End
   Begin VB.Label Label10 
      Caption         =   "File Name:"
      Height          =   250
      Left            =   120
      TabIndex        =   93
      Top             =   2280
      Width           =   850
   End
   Begin VB.Label Label9 
      Caption         =   "Format:"
      Height          =   250
      Left            =   360
      TabIndex        =   91
      Top             =   1920
      Width           =   610
   End
   Begin VB.Label Label8 
      Caption         =   "note"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   6.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   250
      Left            =   120
      TabIndex        =   88
      Top             =   2640
      Width           =   7570
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_Click()

Text1(0).BackColor = &HFFFFFF
If Text1(0).Text = "" Then
Text1(0).BackColor = &H8080FF
Combo2.Text = ""
MsgBox "Error21: Specify the alloy number "
Exit Sub
End If

If Dir(App.Path & "\" & Text1(0).Text, vbDirectory) <> "" Then
    If (Check1.Value = 0) Then
    Load Form8
    Form8.Visible = True
    Combo2.Text = ""
    Form8.Label1.Caption = "Folder for alloy number " & Text1(0).Text & " already exists. Do you want to overwrite it?"
    Exit Sub
    End If
    
End If




  


If Combo2.ListIndex = 0 Then 'from file

            'loading all files that exist in the fs-Temp directory
            sFilename = Dir(App.Path & "\fs-Temp\")
            i = 0
        Do While sFilename > ""
            'gets rid of the extension (.DAT)
            sFilename = Replace(sFilename, ".DAT", "")
            Debug.Print sFilename
          
            Combo3.AddItem sFilename, i
            i = i + 1
            sFilename = Dir()
        
        Loop
        'If there is not any file in the alloy directory
        If i = 0 Then
            MsgBox "There is no file in the fs-Temp Directory"
        End If
        
        
     Label10.Visible = True
     Combo3.Visible = True
    Frame2.Visible = False
    Command1.Visible = False
    Command2.Visible = False
    Command4.Visible = False
    Command3.Visible = True
    Label8.Visible = False
    Form7.Height = 4000
    
ElseIf Combo2.ListIndex = 1 Then 'manual
    Frame2.Visible = True
    Command1.Visible = True
    Command2.Visible = True
    Command4.Visible = True
    Command3.Visible = False
    Command7.Visible = False
    combo3Visible = False
    Label10.Visible = False
    Label8.Visible = True
    Form7.Height = 8000
    
End If
    
    




End Sub

Private Sub Command1_Click()


Dim num1 As Double
Dim num2 As Double
num1 = -1
num2 = -1

For i = 0 To 5
Text1(i).BackColor = &HFFFFFF
Next i

Combo1.BackColor = &HFFFFFF

For i = 0 To 39
Text3(i).BackColor = &HFFFFFF
Next i

marker = 0
For i = 0 To 5
    If Text1(i).Text = "" Then
        Text1(i).BackColor = &H8080FF
        marker = 1
    End If
 Next i
 
    If marker = 1 Then
        MsgBox "Error12: The Fields that are specified should be determined "
        Exit Sub
    End If
    
    If Combo1.Text = "" Then
        Combo1.BackColor = &H8080FF
        MsgBox "Error13: Specify how many data points are being used "
        Exit Sub
    End If
    
    DataPointNu = Combo1.Text
    
    For i = 0 To DataPointNu - 1
    If Text3(i).Text = "" Then
    marker = 1
    Text3(i).BackColor = &H8080FF
    End If
    Next i
    
    If marker = 1 Then
    MsgBox "Error15: More Temperature points are needed "
    Exit Sub
    End If
    
    For i = 0 To DataPointNu - 2
    num1 = Text3(i).Text
    num2 = Text3(i + 1).Text
    If num1 >= num2 Then
    marker = 1
    Text3(i).BackColor = &H8080FF
    Text3(i + 1).BackColor = &H8080FF
    End If
    Next i
    
    If marker = 1 Then
    MsgBox "Error17: The temperature should be listed in ascending order "
    Exit Sub
    End If
    
    For i = 20 To 19 + DataPointNu
    If Text3(i).Text = "" Then
    marker = 1
    Text3(i).BackColor = &H8080FF
    End If
    Next i
    
    If marker = 1 Then
    MsgBox "Error16: More Solid fraction points are needed "
    Exit Sub
    End If
    
    For i = 20 To 18 + DataPointNu
    num1 = Text3(i).Text
    num2 = Text3(i + 1).Text
    If num1 <= num2 Then
    marker = 1
    Text3(i).BackColor = &H8080FF
    Text3(i + 1).BackColor = &H8080FF
    End If
    Next i
    
    If marker = 1 Then
    MsgBox "Error17: The solid fraction decreases by increase of temperature "
    Exit Sub
    End If
    
        For i = 20 To 19 + DataPointNu
        num1 = Text3(i).Text
    If num1 > 1 Or num1 < 0 Then
    marker = 1
    Text3(i).BackColor = &H8080FF
    End If
    Next i
    
    If marker = 1 Then
    MsgBox "Error18: The solid fraction changes between 0 and 1 "
    Exit Sub
    End If
    
    
    For i = DataPointNu To 19
    If Text3(i).Text <> "" Or Text3(20 + i).Text <> "" Then
    marker = 1
    Text3(i).BackColor = &H8080FF
    Text3(i + 20).BackColor = &H8080FF
    Combo1.BackColor = &H8080FF
    End If
    Next i
    
    If marker = 1 Then
    MsgBox "Error20: The data points are more than the number you specified "
    Exit Sub
    End If
    
    
    
    
    If Text3(20).Text <> 1 Then
    Text3(20).BackColor = &H8080FF
    MsgBox "Error14: First point should represent solidus line and f_s=1 "
    Exit Sub
    End If
    
    Lastpoint = 19 + DataPointNu
    If Text3(Lastpoint).Text <> 0 Then
    Text3(Lastpoint).BackColor = &H8080FF
    MsgBox "Error15: Last point should represent liquidus line and f_s=0 "
    Exit Sub
    End If
    
    
    
    
    
 Command2.Enabled = True
 Check2.Value = 1

End Sub

Private Sub Command2_Click()
Form1.Text2(0).Text = Text1(4).Text
Dim TempMax, TempMin, dH, Tup, Tlow, f_Tup, f_Tlow As Double
Dim Ks, Kl, Rho, Lf, Cp As Double
Dim AlloyNu, NEnth As Integer

AlloyNu = Text1(0).Text
Ks = Val(Text1(1).Text)
Kl = Val(Text1(2).Text)
Rho = Val(Text1(3).Text)
Lf = Val(Text1(4).Text)
Cp = Val(Text1(5).Text)

TempMin = Val(Text3(0).Text)
TempMax = Val(Text3(Combo1.Text - 1))
NEnth = 1001


'Frac0 and Temp0 are the array that determined by user
Dim Temp0(), Frac0(), Enthalpy0(), Enth1(), Enth2(), Enth3(), Enth4(), Enth5(), Enth6() As Double
Dim df(), dT() As Double

ReDim Temp0(Combo1.Text)
ReDim Frac0(Combo1.Text)
ReDim Enthalpy0(Combo1.Text)
ReDim dT(NEnth), df(NEnth)
'Enth1=counter, Enth2=H=temp*Cp+Lf*(1-fs)
'Enth3=Temperature, Enth4=dT/dH
'Enth5=fs,Enth6=df/dT
ReDim Enth1(NEnth), Enth2(NEnth), Enth3(NEnth), Enth4(NEnth), Enth5(NEnth), Enth6(NEnth)


'values from the form7
For i = 0 To Combo1.Text - 1
    Temp0(i) = Text3(i).Text
    Frac0(i) = Text3(i + 20).Text
    Enthalpy0(i) = Cp * Temp0(i) + (1 - Frac0(i)) * Lf
Next i

HMin = Cp * TempMin
HMax = Cp * TempMax + Lf
dH = (HMax - HMin) / (NEnth - 1)

'Enthalpy values
For i = 0 To NEnth - 1
    Enth2(i) = HMin + dH * i
Next i


    Tlow = Val(Text3(0).Text)
    f_Tlow = Val(Text3(20).Text)
    Hlow = Enthalpy0(0)
    Tup = Val(Text3(1).Text)
    f_Tup = Val(Text3(21).Text)
    Hup = Enthalpy0(1)

    c = 1
' T and f_s values
For i = 0 To NEnth - 1
    If Enth2(i) >= Hup Then
        c = c + 1
        Tlow = Tup
        f_Tlow = f_Tup
        Hlow = Hup
        Tup = Val(Text3(c).Text)
        f_Tup = Val(Text3(20 + c).Text)
        Hup = Tup * Cp + (1 - f_Tup) * Lf
    End If
        Enth3(i) = Tlow + (Tup - Tlow) / (Hup - Hlow) * (Enth2(i) - Hlow)
        Enth5(i) = f_Tlow + (f_Tup - f_Tlow) / (Tup - Tlow) * (Enth3(i) - Tlow)
Next i

For i = 0 To NEnth - 2
    dT(i) = Enth3(i + 1) - Enth3(i)
Next i

dT(NEnth - 1) = dT(NEnth - 2)

For i = 0 To NEnth - 2
    df(i) = Enth5(i + 1) - Enth5(i)
Next i

df(NEnth - 1) = df(NEnth - 2)

'dT/dH
For i = 0 To NEnth - 1
    Enth4(i) = dT(i) / dH
Next i

'df/dT
For i = 0 To NEnth - 1
    Enth6(i) = df(i) / dT(i)
Next i


Dim iFileNo As Integer
iFileNo = FreeFile
If Dir(App.Path & "\ENTH.DAT") <> "" Then
Kill (App.Path & "\ENTH.DAT")
End If


Open App.Path & "\ENTH.DAT" For Output As #iFileNo
Print #iFileNo, " "

For i = 0 To NEnth - 1
Print #iFileNo, i + 1, Format$(Enth2(i), "0.000000E+"), Format$(Enth3(i), "0.000000E+"), Format$(Enth4(i), "0.000000E+"), Format$(Enth5(i), "0.000000E+"), Format$(Enth6(i), "0.000000E+")
Next i

Close #iFileNo

' Making Data file
Dim NData As Integer
NData = 901
Dim Data1(), Data2(), Data3(), Data4(), Data5(), Data6(), Data7(), Data8(), Data9(), Data10(), Data11(), Data12() As Double
ReDim Data1(NData), Data2(NData), Data3(NData), Data4(NData), Data5(NData), Data6(NData), Data7(NData), Data8(NData), Data9(NData), Data10(NData), Data11(NData), Data12(NData)

'dT in Data file
dTData = 1

'Temp values
For i = 0 To NData - 1
    Data2(i) = i
Next i

    Tlow = Val(Text3(0).Text)
    f_Tlow = Val(Text3(20).Text)
    Tup = Val(Text3(1).Text)
    f_Tup = Val(Text3(21).Text)
    c = 1
    
    
    

'f_s values
For i = 0 To NData - 1
    If Data2(i) < TempMin Then
        Data3(i) = 1
    ElseIf Data2(i) > TempMax Then
        Data3(i) = 0
    Else
        If Data2(i) >= Tup Then
            c = c + 1
            Tlow = Tup
            f_Tlow = f_Tup
            Tup = Val(Text3(c).Text)
            f_Tup = Val(Text3(20 + c).Text)
        End If
            Data3(i) = f_Tlow + (Data2(i) - Tlow) * (f_Tup - f_Tlow) / (Tup - Tlow)
   End If
        
Next i

'calculation of df
Dim dfData() As Double
ReDim dfData(NData)

For i = 0 To NData - 2
    dfData(i) = Data3(i + 1) - Data3(i)
Next i

dfData(NData - 1) = dfData(NData - 2)

'dfData/dTData
For i = 0 To NData - 1
    Data4(i) = dfData(i) / dTData
Next i

'fData*Ks+(1-fData)*Kl
For i = 0 To NData - 1
    Data5(i) = Data3(i) * Ks + (1 - Data3(i)) * Kl
Next i

'Never uses 50*Data4
For i = 0 To NData - 1
    Data6(i) = Data4(i) * 50
Next i

'CpT
For i = 0 To NData - 1
    Data7(i) = Data2(i) * Cp
Next i

'Cp
For i = 0 To NData - 1
    Data8(i) = Cp
Next i

'CpT+Lf*(1-fs)
For i = 0 To NData - 1
    Data9(i) = Cp * Data2(i) + Lf * (1 - Data3(i))
Next i

'Cp-Lf*df/dT
For i = 0 To NData - 1
    Data10(i) = Cp - Lf * Data4(i)
Next i

'Rho
For i = 0 To NData - 1
    Data11(i) = Rho
Next i

'0
For i = 0 To NData - 1
    Data12(i) = 0
Next i

Dim iFileNo1 As Integer
iFileNo1 = FreeFile
If Dir(App.Path & "\DATA.DAT") <> "" Then
Kill (App.Path & "\DATA.DAT")
End If


Open App.Path & "\DATA.DAT" For Output As #iFileNo1
Print #iFileNo1, "No.               T              F"

For i = 0 To NData - 1
Print #iFileNo1, i + 1, Format$(Data2(i), "0.000000E+"), Format$(Data3(i), "0.000000E+"), Format$(Data4(i), "0.000000E+"), Format$(Data5(i), "0.000000E+"), Format$(Data6(i), "0.000000E+"), Format$(Data7(i), "0.000000E+"), Format$(Data8(i), "0.000000E+"), Format$(Data9(i), "0.000000E+"), Format$(Data10(i), "0.000000E+"), Format$(Data11(i), "0.000000E+"), Format$(Data12(i), "0.000000E+")
Next i

Close #iFileNo1







'making directory with the name of the castID
If Dir(App.Path & "\" & AlloyNu, vbDirectory) = "" Then
    MkDir (App.Path & "\" & AlloyNu)
End If

'copying the Data files to the directory
If Dir(App.Path & "\DATA.DAT") <> "" Then
FileCopy App.Path & "\DATA.DAT", App.Path & "\" & AlloyNu & "\DATA.DAT"
Else
MsgBox "The DATA.DAT file has not produced. Try again"
Exit Sub
End If

'copying the ENTH files to the directory
If Dir(App.Path & "\ENTH.DAT") <> "" Then
FileCopy App.Path & "\ENTH.DAT", App.Path & "\" & AlloyNu & "\ENTH.DAT"
Else
MsgBox "The ENTH.DAT file has not produced. Try again"
Exit Sub
End If


'making the file of the alloy in the alloy directory
If Dir(App.Path & "\Alloys\" & AlloyNu & ".DAT") <> "" Then
Kill (App.Path & "\Alloys\" & AlloyNu & ".DAT")
End If


Dim iFileNo2 As Integer
iFileNo2 = FreeFile
Open App.Path & "\Alloys\" & AlloyNu & ".DAT" For Output As #iFileNo2
Print #iFileNo2, Lf
Print #iFileNo2, Form1.Text2(1).Text


Close #iFileNo2


Check2.Value = 0


End Sub


Private Sub Command3_Click()
Combo3.BackColor = &HFFFFFF
If Combo3.Text = "" Then
Combo3.BackColor = &H8080FF
    MsgBox "Choose the file from the drop down menu"
    Exit Sub
    End If
'import data points
If Dir(App.Path & "/fs-Temp/" & Combo3.Text) <> "" Then
    Dim newfile2
    Dim iFileNo2 As Integer
    Dim PP1, PP2
 Dim Temperature(0 To 1000), SFraction(0 To 1000), Enthalpy(0 To 1000) As Double
 Dim NuDataPoint As Integer
    iFileNo2 = FreeFile
    NuDataPoint = 0
    Open App.Path & "/fs-Temp/" & Combo3.Text For Input As #iFileNo2

    Do While Not EOF(iFileNo2)
        Input #iFileNo2, Temperature(NuDataPoint), SFraction(NuDataPoint)
        NuDataPoint = NuDataPoint + 1
    Loop
Else
    MsgBox "The file " & Combo3.Text & " does not exist"
    Exit Sub
End If



'Checking the format of the data points



Dim num1 As Double
Dim num2 As Double
num1 = -1
num2 = -1

For i = 0 To 5
Text1(i).BackColor = &HFFFFFF
Next i



marker = 0
For i = 0 To 5
    If Text1(i).Text = "" Then
        Text1(i).BackColor = &H8080FF
        marker = 1
    End If
 Next i
 
    If marker = 1 Then
        MsgBox "Error12: The Fields that are specified should be determined "
        Exit Sub
    End If
    
    
    
    For i = 0 To NuDataPoint - 2
    num1 = Temperature(i)
    num2 = Temperature(i + 1)
    If num1 >= num2 Then
    marker = 1
    End If
    Next i
    
    If marker = 1 Then
    MsgBox "Error17:Wrong file format. The temperature should be listed in ascending order "
    Exit Sub
    End If
    
    
    For i = 0 To NuDataPoint - 2
    num1 = SFraction(i)
    num2 = SFraction(i + 1)
    If num1 <= num2 Then
    marker = 1
    End If
    Next i
    
    If marker = 1 Then
    MsgBox "Error17: Wrong file format. The solid fraction should decrease by increase of temperature "
    Exit Sub
    End If
    
        For i = 0 To NuDataPoint - 1
        num1 = SFraction(i)
    If num1 > 1 Or num1 < 0 Then
    marker = 1
    End If
    Next i
    
    If marker = 1 Then
    MsgBox "Error18:Wrong file format. The solid fraction changes between 0 and 1 "
    Exit Sub
    End If
    
    
    
    
    
    
    If SFraction(0) <> 1 Then
    MsgBox "Error14:Wrong file format. First point should represent solidus line and f_s=1 "
    Exit Sub
    End If
    

    If SFraction(NuDataPoint - 1) <> 0 Then
    MsgBox "Error15:Wrong file format. Last point should represent liquidus line and f_s=0 "
    Exit Sub
    End If
    
    
    
    
    
' Command2.Enabled = True
 Check2.Value = 1


'Making Data files

Form1.Text2(0).Text = Text1(4).Text
Form1.Text2(2).Text = Text1(0).Text
Dim TempMax, TempMin, Tup, Tlow, f_Tup, f_Tlow As Double
Dim Ks, Kl, Rho, Lf, Cp As Double
Dim AlloyNu, NEnth As Integer

AlloyNu = Text1(0).Text
Ks = Val(Text1(1).Text)
Kl = Val(Text1(2).Text)
Rho = Val(Text1(3).Text)
Lf = Val(Text1(4).Text)
Cp = Val(Text1(5).Text)

TempMin = Val(Temperature(0))
TempMax = Val(Temperature(NuDataPoint - 1))
HMin = Cp * TempMin
HMax = Cp * TempMax + Lf
NEnth = 1001

For i = 0 To NuDataPoint - 1
    Enthalpy(i) = Cp * Temperature(i) + Lf * (1 - SFraction(i))
Next i

'Frac0 and Temp0 are the array that determined by user
Dim Enth1(), Enth2(), Enth3(), Enth4(), Enth5(), Enth6() As Double
Dim dH, df(), dT() As Double


ReDim df(NEnth)
ReDim dT(NEnth)
'Enth1=counter, Enth2=H=temp*Cp+Lf*(1-fs)
'Enth3=Temperature, Enth4=dT/dH
'Enth5=fs,Enth6=df/dT
ReDim Enth1(NEnth), Enth2(NEnth), Enth3(NEnth), Enth4(NEnth), Enth5(NEnth), Enth6(NEnth)

dH = (HMax - HMin) / (NEnth - 1)


'Enthalpy values
For i = 0 To NEnth - 1
    Enth2(i) = HMin + dH * i
Next i


    Tlow = Temperature(0)
    f_Tlow = SFraction(0)
    Tup = Temperature(1)
    f_Tup = SFraction(1)
    Hlow = Enthalpy(0)
    Hup = Enthalpy(1)
    c = 1
'T and f_s values
For i = 0 To NEnth - 1
    If Enth2(i) >= Hup Then
        c = c + 1
        Hlow = Hup
        Tlow = Tup
        f_Tlow = f_Tup
        Hup = Val(Enthalpy(c))
        Tup = Val(Temperature(c))
        f_Tup = Val(SFraction(c))
    End If
        Enth3(i) = Tlow + (Tup - Tlow) / (Hup - Hlow) * (Enth2(i) - Hlow)
        Enth5(i) = f_Tlow + (f_Tup - f_Tlow) / (Tup - Tlow) * (Enth3(i) - Tlow)
Next i




For i = 0 To NEnth - 2
    dT(i) = Enth3(i + 1) - Enth3(i)
Next i

dT(NEnth - 1) = dT(NEnth - 2)

For i = 0 To NEnth - 2
    df(i) = Enth5(i + 1) - Enth5(i)
Next i

df(NEnth - 1) = df(NEnth - 2)

'dT/dH
For i = 0 To NEnth - 1
    Enth4(i) = dT(i) / dH
Next i

'df/dT
For i = 0 To NEnth - 1
    Enth6(i) = df(i) / dT(i)
Next i


Dim iFileNo As Integer
iFileNo = FreeFile
If Dir(App.Path & "\ENTH.DAT") <> "" Then
Kill (App.Path & "\ENTH.DAT")
End If


Open App.Path & "\ENTH.DAT" For Output As #iFileNo
Print #iFileNo, " "

For i = 0 To NEnth - 1
Print #iFileNo, i + 1, Format$(Enth2(i), "0.000000E+"), Format$(Enth3(i), "0.000000E+"), Format$(Enth4(i), "0.000000E+"), Format$(Enth5(i), "0.000000E+"), Format$(Enth6(i), "0.000000E+")
Next i

Close #iFileNo

' Making Data file
Dim NData As Integer
NData = 901
Dim Data1(), Data2(), Data3(), Data4(), Data5(), Data6(), Data7(), Data8(), Data9(), Data10(), Data11(), Data12() As Double
ReDim Data1(NData), Data2(NData), Data3(NData), Data4(NData), Data5(NData), Data6(NData), Data7(NData), Data8(NData), Data9(NData), Data10(NData), Data11(NData), Data12(NData)

'dT in Data file
dTData = 1

'Temp values
For i = 0 To NData - 1
    Data2(i) = i
Next i

    Tlow = Val(Temperature(0))
    f_Tlow = Val(SFraction(0))
    Tup = Val(Temperature(1))
    f_Tup = Val(SFraction(1))
    c = 1
    
    
    

'f_s values
For i = 0 To NData - 1
    If Data2(i) < TempMin Then
        Data3(i) = 1
    ElseIf Data2(i) > TempMax Then
        Data3(i) = 0
    Else
        If Data2(i) >= Tup Then
            c = c + 1
            Tlow = Tup
            f_Tlow = f_Tup
            Tup = Val(Temperature(c))
            f_Tup = Val(SFraction(c))
        End If
            Data3(i) = f_Tlow + (Data2(i) - Tlow) * (f_Tup - f_Tlow) / (Tup - Tlow)
   End If
        
Next i

'calculation of df
Dim dfData() As Double
ReDim dfData(NData)

For i = 0 To NData - 2
    dfData(i) = Data3(i + 1) - Data3(i)
Next i

dfData(NData - 1) = dfData(NData - 2)

'dfData/dTData
For i = 0 To NData - 1
    Data4(i) = dfData(i) / dTData
Next i

'fData*Ks+(1-fData)*Kl
For i = 0 To NData - 1
    Data5(i) = Data3(i) * Ks + (1 - Data3(i)) * Kl
Next i

'Never uses 50*Data4
For i = 0 To NData - 1
    Data6(i) = Data4(i) * 50
Next i

'CpT
For i = 0 To NData - 1
    Data7(i) = Data2(i) * Cp
Next i

'Cp
For i = 0 To NData - 1
    Data8(i) = Cp
Next i

'CpT+Lf*(1-fs)
For i = 0 To NData - 1
    Data9(i) = Cp * Data2(i) + Lf * (1 - Data3(i))
Next i

'Cp-Lf*df/dT
For i = 0 To NData - 1
    Data10(i) = Cp - Lf * Data4(i)
Next i

'Rho
For i = 0 To NData - 1
    Data11(i) = Rho
Next i

'0
For i = 0 To NData - 1
    Data12(i) = 0
Next i

Dim iFileNo1 As Integer
iFileNo1 = FreeFile
If Dir(App.Path & "\DATA.DAT") <> "" Then
Kill (App.Path & "\DATA.DAT")
End If


Open App.Path & "\DATA.DAT" For Output As #iFileNo1
Print #iFileNo1, "No.               T              F"

For i = 0 To NData - 1
Print #iFileNo1, i + 1, Format$(Data2(i), "0.000000E+"), Format$(Data3(i), "0.000000E+"), Format$(Data4(i), "0.000000E+"), Format$(Data5(i), "0.000000E+"), Format$(Data6(i), "0.000000E+"), Format$(Data7(i), "0.000000E+"), Format$(Data8(i), "0.000000E+"), Format$(Data9(i), "0.000000E+"), Format$(Data10(i), "0.000000E+"), Format$(Data11(i), "0.000000E+"), Format$(Data12(i), "0.000000E+")
Next i

Close #iFileNo1







'making directory with the name of the castID
If Dir(App.Path & "\" & AlloyNu, vbDirectory) = "" Then
    MkDir (App.Path & "\" & AlloyNu)
End If

'copying the Data files to the directory
If Dir(App.Path & "\DATA.DAT") <> "" Then
FileCopy App.Path & "\DATA.DAT", App.Path & "\" & AlloyNu & "\DATA.DAT"
Else
MsgBox "The DATA.DAT file has not produced. Try again"
Exit Sub
End If

'copying the ENTH files to the directory
If Dir(App.Path & "\ENTH.DAT") <> "" Then
FileCopy App.Path & "\ENTH.DAT", App.Path & "\" & AlloyNu & "\ENTH.DAT"
Else
MsgBox "The ENTH.DAT file has not produced. Try again"
Exit Sub
End If


'making the file of the alloy in the alloy directory
If Dir(App.Path & "\Alloys\" & AlloyNu & ".DAT") <> "" Then
Kill (App.Path & "\Alloys\" & AlloyNu & ".DAT")
End If



iFileNo2 = FreeFile
Open App.Path & "\Alloys\" & AlloyNu & ".DAT" For Output As #iFileNo2
Print #iFileNo2, Lf
Print #iFileNo2, Form1.Text2(1).Text


Close #iFileNo2


Check2.Value = 0












End Sub

Private Sub Command4_Click()
If Check2.Value = 0 Then
Unload Form7
Else

    Load Form9
    Form9.Visible = True
    Form9.Label1.Caption = "You have not created the DATA and ENTH files. Do you want to exit?"
    Exit Sub

End If


    








End Sub




Private Sub Command5_Click()

End Sub

Private Sub Command7_Click()
If Check2.Value = 0 Then
Unload Form7
Else

    Load Form9
    Form9.Visible = True
    Form9.Label1.Caption = "You have not created the DATA and ENTH files. Do you want to exit?"
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Frame2.Visible = False
Combo3.Visible = False
Label10.Visible = False
Label8.Visible = False


Form7.Left = Form1.Left + 2 * Form1.Width / 3
For i = 1 To 20
Combo1.AddItem i, i - 1
Next i
Label8.Caption = "List temperatures in ascending order. The corresponding solid fraction varies between 1 and 0"

Combo2.AddItem "From txt file", 0
Combo2.AddItem "Manual", 1



End Sub

