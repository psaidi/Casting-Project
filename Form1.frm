VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   11610
   ClientLeft      =   4000
   ClientTop       =   10
   ClientWidth     =   8570
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11610
   ScaleMode       =   0  'User
   ScaleWidth      =   8570
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command8 
      Caption         =   "Exit"
      Height          =   250
      Left            =   6840
      TabIndex        =   79
      Top             =   10560
      Width           =   1210
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Accept-Plot "
      Height          =   250
      Left            =   3240
      TabIndex        =   11
      Top             =   10560
      Width           =   1210
   End
   Begin VB.Frame Frame5 
      Caption         =   "Simulation Settings"
      Height          =   1930
      Left            =   120
      TabIndex        =   57
      Top             =   8520
      Width           =   7935
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   9
         Left            =   6600
         TabIndex        =   85
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   8
         Left            =   6600
         TabIndex        =   84
         Top             =   360
         Width           =   1000
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   7
         Left            =   3840
         TabIndex        =   83
         Top             =   1440
         Width           =   1000
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   6
         Left            =   3840
         TabIndex        =   82
         Top             =   1080
         Width           =   1000
      End
      Begin VB.CommandButton Command10 
         Caption         =   "ok"
         Height          =   250
         Left            =   6480
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   1440
         Width           =   970
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Edit"
         Height          =   250
         Left            =   5400
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   1440
         Width           =   970
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   5
         Left            =   3840
         TabIndex        =   69
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   4
         Left            =   3840
         TabIndex        =   68
         Top             =   360
         Width           =   1000
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   67
         Top             =   1440
         Width           =   1000
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   66
         Top             =   1080
         Width           =   1000
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   65
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   61
         Top             =   360
         Width           =   1000
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Caption         =   "Max. No. of Sweeps."
         Height          =   380
         Left            =   4920
         TabIndex        =   89
         Top             =   720
         Width           =   1560
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "Exit Temp. Closness"
         Height          =   380
         Left            =   4920
         TabIndex        =   88
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "Inlet Enth. Frac."
         Height          =   380
         Left            =   2520
         TabIndex        =   87
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   " speed Change"
         Height          =   380
         Left            =   2520
         TabIndex        =   86
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "No. of Sweeps."
         Height          =   380
         Left            =   2520
         TabIndex        =   64
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "Sol. Frac. Relax."
         Height          =   380
         Left            =   2520
         TabIndex        =   63
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "Temp. Relax."
         Height          =   380
         Left            =   360
         TabIndex        =   62
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Upwinding Para."
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Time Step"
         Height          =   375
         Left            =   240
         TabIndex        =   59
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Iterations"
         Height          =   375
         Left            =   240
         TabIndex        =   58
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   11200
      LargeChange     =   90
      Left            =   8160
      Max             =   100
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Caption         =   "Calculation Settings"
      Height          =   2650
      Left            =   120
      TabIndex        =   35
      Top             =   5760
      Width           =   7935
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   250
         Left            =   2520
         TabIndex        =   94
         Top             =   1080
         Visible         =   0   'False
         Width           =   250
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Edit"
         Height          =   250
         Left            =   5400
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   2160
         Width           =   970
      End
      Begin VB.CommandButton Command6 
         Caption         =   "ok"
         Height          =   250
         Left            =   6480
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   2160
         Width           =   970
      End
      Begin VB.ComboBox Combo2 
         Height          =   280
         Left            =   4680
         TabIndex        =   70
         Text            =   "Site Data"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   2760
         TabIndex        =   7
         Text            =   "Combo5"
         Top             =   2280
         Width           =   1545
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   2760
         TabIndex        =   6
         Text            =   "Combo4"
         Top             =   1680
         Width           =   1545
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2760
         TabIndex        =   5
         Text            =   "Combo3"
         Top             =   1080
         Width           =   1545
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   2280
         Width           =   1000
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   2040
         Width           =   1000
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   1800
         Width           =   1000
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   1560
         Width           =   1000
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   1320
         Width           =   1000
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   1080
         Width           =   1000
      End
      Begin VB.TextBox Text4 
         Height          =   300
         Index           =   2
         Left            =   4200
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "Text4"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Height          =   300
         Index           =   1
         Left            =   1440
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "Text4"
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text4 
         Height          =   300
         Index           =   0
         Left            =   1440
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "Text4"
         Top             =   360
         Width           =   1000
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "h.t.c coefficients"
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label20 
         Caption         =   "solid fraction"
         Height          =   375
         Left            =   4680
         TabIndex        =   44
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Max h.t.c"
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Max h.t.c applies at"
         Height          =   375
         Left            =   2280
         TabIndex        =   42
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Min h.t.c"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cast Condition"
      Height          =   2170
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   7935
      Begin VB.CommandButton Command5 
         Caption         =   "ok"
         Height          =   250
         Left            =   6480
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   1680
         Width           =   970
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Edit"
         Height          =   250
         Left            =   5400
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   1680
         Width           =   970
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   5
         Left            =   5520
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   2
         Left            =   1440
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1000
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   1
         Left            =   1440
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   6
         Left            =   5520
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   4
         Left            =   1440
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1000
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   3
         Left            =   1440
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1000
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   0
         Left            =   1440
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   360
         Width           =   1000
      End
      Begin VB.Label Label23 
         Caption         =   "(C)"
         Height          =   260
         Index           =   5
         Left            =   2520
         TabIndex        =   93
         Top             =   1440
         Width           =   260
      End
      Begin VB.Label Label23 
         Caption         =   "(C)"
         Height          =   260
         Index           =   4
         Left            =   2520
         TabIndex        =   92
         Top             =   1080
         Width           =   260
      End
      Begin VB.Label Label23 
         Caption         =   "(m/min)"
         Height          =   260
         Index           =   1
         Left            =   2520
         TabIndex        =   91
         Top             =   720
         Width           =   610
      End
      Begin VB.Label Label23 
         Caption         =   "(C)"
         Height          =   260
         Index           =   0
         Left            =   2520
         TabIndex        =   90
         Top             =   360
         Width           =   260
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Taper"
         Height          =   260
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Width           =   1220
      End
      Begin VB.Label Label16 
         Caption         =   "solid fraction"
         Height          =   380
         Left            =   6000
         TabIndex        =   28
         Top             =   720
         Width           =   1100
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Squeeze Princ. applies at"
         Height          =   260
         Left            =   3360
         TabIndex        =   27
         Top             =   720
         Width           =   2060
      End
      Begin VB.Label Label14 
         Caption         =   "solid fraction"
         Height          =   260
         Left            =   6000
         TabIndex        =   26
         Top             =   360
         Width           =   980
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Inlet Temp."
         Height          =   375
         Left            =   480
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Gap forms at"
         Height          =   380
         Left            =   4320
         TabIndex        =   24
         Top             =   360
         Width           =   1100
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Coolant Temp."
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Exit Temp."
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Casting Speed"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Material Properties"
      Height          =   1210
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   7935
      Begin VB.TextBox Text2 
         Height          =   300
         Index           =   2
         Left            =   1560
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   720
         Width           =   1000
      End
      Begin VB.ComboBox Combo6 
         Height          =   280
         Left            =   960
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   720
         Width           =   1600
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Index           =   1
         Left            =   4080
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   360
         Width           =   1000
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   4080
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   720
         Width           =   1000
      End
      Begin VB.ComboBox Combo1 
         Height          =   280
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   1600
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "Alloy Nu."
         Height          =   260
         Left            =   720
         TabIndex        =   74
         Top             =   720
         Width           =   740
      End
      Begin VB.Label Label5 
         Caption         =   "J/Kg"
         Height          =   260
         Left            =   5160
         TabIndex        =   72
         Top             =   720
         Width           =   610
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Vol. Shrinkage"
         Height          =   260
         Left            =   2900
         TabIndex        =   23
         Top             =   360
         Width           =   1100
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Latent Heat"
         Height          =   380
         Left            =   3000
         TabIndex        =   17
         Top             =   720
         Width           =   980
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "System Geometry"
      Height          =   1340
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   5650
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   250
         Left            =   5040
         TabIndex        =   97
         Top             =   240
         Visible         =   0   'False
         Width           =   250
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   4
         Left            =   960
         TabIndex        =   95
         Top             =   240
         Width           =   3880
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   3
         Left            =   3840
         TabIndex        =   3
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   2
         Left            =   3840
         TabIndex        =   2
         Top             =   600
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   1
         Left            =   960
         TabIndex        =   1
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   0
         Left            =   960
         TabIndex        =   0
         Top             =   600
         Width           =   1000
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Cast ID"
         Height          =   380
         Left            =   240
         TabIndex        =   96
         Top             =   240
         Width           =   610
      End
      Begin VB.Label Label23 
         Caption         =   "(m)"
         Height          =   260
         Index           =   3
         Left            =   4920
         TabIndex        =   56
         Top             =   960
         Width           =   260
      End
      Begin VB.Label Label23 
         Caption         =   "(m)"
         Height          =   260
         Index           =   2
         Left            =   4920
         TabIndex        =   55
         Top             =   600
         Width           =   260
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nx"
         Height          =   380
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   740
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ny"
         Height          =   260
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   620
      End
      Begin VB.Label Label3 
         Caption         =   "Cooling Zone Length"
         Height          =   380
         Left            =   2280
         TabIndex        =   13
         Top             =   600
         Width           =   1690
      End
      Begin VB.Label Label4 
         Caption         =   "Slab Height"
         Height          =   260
         Left            =   2880
         TabIndex        =   12
         Top             =   960
         Width           =   1340
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check"
      Height          =   250
      Left            =   120
      TabIndex        =   8
      Top             =   10560
      Width           =   1450
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Height          =   250
      Left            =   1680
      TabIndex        =   10
      Top             =   10560
      Width           =   1450
   End
   Begin VB.Image Image1 
      Height          =   1090
      Left            =   5880
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1090
   End
   Begin VB.Image Image2 
      Height          =   1090
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1090
   End
   Begin VB.Shape Shape1 
      Height          =   11600
      Left            =   0
      Top             =   -120
      Width           =   8530
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "Label22"
      Height          =   500
      Left            =   120
      TabIndex        =   54
      Top             =   120
      Width           =   5650
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
    Combo1.BackColor = &HFFFFFF
    Combo6.Clear
    Combo6.Text = "Alloy number"
    Text2(0).Text = ""
    Text2(1).Text = ""
    Text2(2).Text = ""
    
    
    
    'IF loads an existing alloy file
    If Combo1.ListIndex = 1 Then
                Form7.Visible = False
                 Label5.Visible = True
                 Label6.Visible = True
                 Label13.Visible = True
                 Text2(0).Visible = True
                 Text2(1).Visible = True
            Combo6.Visible = True
            Text2(2).Visible = False
            Label33.Visible = False
            'loading all files that exist in the Alloy directory
            sFilename = Dir(App.Path & "\Alloys\")
            i = 0
        Do While sFilename > ""
            'gets rid of the extension (.DAT)
            sFilename = Replace(sFilename, ".DAT", "")
            Debug.Print sFilename
          
            Combo6.AddItem sFilename, i
            i = i + 1
            sFilename = Dir()
        
        Loop
        'If there is not any file in the alloy directory
        If i = 0 Then
            MsgBox "There is no file in the Alloy Directory"
        End If
        
        'IF making new alloy alloy file
        ElseIf Combo1.ListIndex = 0 Then
            Combo6.Visible = False
   '             Text2(2).Visible = True
   '             Label33.Visible = True
                 Label5.Visible = False
                 Label6.Visible = False
   '              Label13.Visible = False
                 Text2(0).Visible = False
                 Text2(1).Text = 0.07
                 Form7.Visible = True


    End If


End Sub

Private Sub Combo2_Click()
    Form3.Visible = False
    Form4.Visible = False
    If (Combo2.ListIndex = 0) Then
        Form3.Visible = True
        Form3.Combo1.Clear
            'loading all files that exist in the Location directory
            sFilename = Dir(App.Path & "\Locations\")
            i = 0
        Do While sFilename > ""
            Debug.Print sFilename
          
            Form3.Combo1.AddItem sFilename, i
            i = i + 1
            sFilename = Dir()
        
        Loop
        'If there is not any file in the alloy directory
        If i = 0 Then
            MsgBox "There is no file in the Location Directory"
        End If
    ElseIf (Combo2.ListIndex = 1) Then
        Form4.Visible = True
    End If


End Sub

Private Sub Combo3_click()
Combo3.BackColor = &HFFFFFF
Combo2.Visible = False
If (Combo3.ListIndex = 1) Then
Combo2.Visible = True
End If

'this part tries to make sure that the site data has been submitted if the
'flux values are determined

If (Combo3.ListIndex = 1) Then
Check1.Value = 0
End If
If (Combo3.ListIndex = 2) Then
Check1.Value = 1
End If
End Sub

Private Sub Combo4_click()
Combo4.BackColor = &HFFFFFF
End Sub

Private Sub Combo5_click()
Combo5.BackColor = &HFFFFFF
End Sub

Private Sub Combo6_click()
Combo6.BackColor = &HFFFFFF
Dim newfile6
Dim iFileNo6 As Integer
Dim pp6
iFileNo6 = FreeFile


Text2(0).Text = ""
Text2(1).Text = ""
Open App.Path & "\Alloys\" & Combo6.Text & ".DAT" For Input As #iFileNo6
'latent heat
Line Input #iFileNo6, pp6
Text2(0).Text = pp6
'volume shrinkage
Line Input #iFileNo6, pp6
Text2(1).Text = pp6



Close #iFileNo6







End Sub

Private Sub Command1_Click()
If Dir(App.Path & "\INPUT.DAT") <> "" Then
Kill (App.Path & "\INPUT.DAT")
End If




Dim newfile
Dim AlloyNumber
Dim iFileNo As Integer
iFileNo = FreeFile
Open App.Path & "\INPUT.DAT" For Output As #iFileNo

'nx and ny
Print #iFileNo, Text1(0).Text; " "; Text1(1).Text
'alloy number
AlloyNumber = ""
If Combo1.ListIndex = 1 Then
        Print #iFileNo, Combo6.Text
        AlloyNumber = Combo6.Text
    ElseIf Combo1.ListIndex = 0 Then
        Print #iFileNo, Text2(2).Text
        AlloyNumber = Text2(2).Text
End If

'length of caster; thickness of caster
Print #iFileNo, Text1(2).Text; " "; Text1(3).Text
'latent heat
Print #iFileNo, Text2(0).Text
'logicals for heat flux given (T) or hct given (F)
If (Combo3.ListIndex = 1) Then
Print #iFileNo, "T"; " ";
ElseIf (Combo3.ListIndex = 2) Then
Print #iFileNo, "F"; " ";
Else
MsgBox "Determine the boundary condition type"
End If

'logicals for speed given (F) or exit temperature given (T)this is opposit of the explanation on the fortran code
If (Combo4.ListIndex = 1) Then
Print #iFileNo, "T"; " ";
ElseIf (Combo4.ListIndex = 2) Then
Print #iFileNo, "F"; " ";
Else
MsgBox "Determine the Given Value"
End If

'logicals for elliptic solver (T) or parabolic solver (F)
'Simplifying commands
'------------------------------------------
If (Form2.Check1.Value = 0) Then
    Print #iFileNo, "F"
Else
    If (Combo5.ListIndex = 1) Then
        Print #iFileNo, "T"
    ElseIf (Combo5.ListIndex = 2) Then
        Print #iFileNo, "F"
    Else
        MsgBox "Determine the solver"
    End If
End If

Dim txtInletSpeed
txtInletSpeed = CDbl(Text3(1).Text) / 60

'inlet temp, inlet speed, exit temp, coolant temp, taper
Print #iFileNo, Text3(0).Text; " "; txtInletSpeed; " "; Text3(2).Text; " "; Text3(3).Text; " "; Text3(4).Text

'Volume shrinkage (on solid), frac solid when gap forms, frac solid where squeeze principle applies
Print #iFileNo, Text2(1).Text; " "; Text3(5).Text; " "; Text3(6).Text

'max h.t.c , h=hmax when frac sol< fmax, min. h.t.c
Print #iFileNo, Text4(0).Text; " "; Text4(2).Text; " "; Text4(1).Text

'six parameters for use in h.t.c correlation
Print #iFileNo, Text5(0).Text; " "; Text5(1).Text; " "; Text5(2).Text
Print #iFileNo, Text5(3).Text; " "; Text5(4).Text; " "; Text5(5).Text

' no of internal iterations on solver, false time step, upwinding parameter
Print #iFileNo, Text6(0).Text; " "; Text6(1).Text; " "; Text6(2).Text

'linear relaxation on temperature, and fraction solid, number of sweeps before speed update
Print #iFileNo, Text6(3).Text; " "; Text6(4).Text; " "; Text6(5).Text

'amount of change speed by to match exit temperature, fraction of inlet enth for convergance criteria,
'closeness to desired exit temperature; maximum number of sweeps possible
Print #iFileNo, Text6(6).Text; " "; Text6(7).Text; " "; Text6(8).Text; " "; Text6(9).Text

'file is not from resrart file
Print #iFileNo, "F"


'no name for restart file
Print #iFileNo, "NO RESTART FILE"

'no address for restart file
Print #iFileNo, "NO RESTART FILE ADDRESS"


'logicals for heat flux given (T) or hct given (F)
If (Combo3.ListIndex = 1) Then
If Dir(App.Path & "\Temporary\BC.DAT") <> "" Then

Dim newfileBC
Dim iFileNoBC As Integer
Dim ppBC, ppBC1, ppBC2, ppBC3
iFileNoBC = FreeFile

Open App.Path & "\Temporary\BC.DAT" For Input As #iFileNoBC
Line Input #iFileNoBC, ppBC
Line Input #iFileNoBC, ppBC
Line Input #iFileNoBC, ppBC
Print #iFileNo, ppBC
NoOfNodes = ppBC


For k = 1 To NoOfNodes
    Line Input #iFileNoBC, ppBC1
    Line Input #iFileNoBC, ppBC2
    Line Input #iFileNoBC, ppBC3
    Print #iFileNo, ppBC1; " "; ppBC2; " "; ppBC3
Next k



Close #iFileNoBC

End If
End If
Close #iFileNo

'copy Data and ENTH files in the main path

If Dir(App.Path & "\DATA.DAT") <> "" Then
    Kill (App.Path & "\DATA.DAT")
End If

If Dir(App.Path & "\ENTH.DAT") <> "" Then
    Kill (App.Path & "\ENTH.DAT")
End If




If Dir(App.Path & "\" & AlloyNumber, vbDirectory) <> "" Then
    If Dir(App.Path & "\" & AlloyNumber & "\ENTH.DAT") <> "" Then
        FileCopy App.Path & "\" & AlloyNumber & "\ENTH.DAT", App.Path & "\ENTH.DAT"
    Else
        MsgBox "ENTH.DAT file is not in the corresponding folder"
    End If
    
    If Dir(App.Path & "\" & AlloyNumber & "\DATA.DAT") <> "" Then
        FileCopy App.Path & "\" & AlloyNumber & "\DATA.DAT", App.Path & "\DATA.DAT"
    Else
        MsgBox "DATA.DAT file is not in the corresponding folder"
    End If
Else
    
  MsgBox "The folder of the specified alloy number does not exist"
    
End If


'deleting old files

If Dir(App.Path & "\RESTART.DAT") <> "" Then
    Kill (App.Path & "\RESTART.DAT")
End If


'running the executable file
Dim Res
Dim Filenameexe

    Filenameexe = (App.Path & "/a.exe") 'Check file is here first

If Dir(Filenameexe) = "" Then
    MsgBox Filenameexe & " not found", vbInformation
    Exit Sub
Else
    Res = Shell(Filenameexe)
End If



Command3.Enabled = True



End Sub

Private Sub Command10_Click()

For i = 0 To 9
    Text6(i).Enabled = False
Next i


Dim newfileSim2

Dim iFileNoSim2 As Integer
iFileNoSim2 = FreeFile
Open App.Path & "\Temporary\SimConst.DAT" For Output As #iFileNoSim2


For i = 0 To 9
Print #iFileNoSim2, Text6(i).Text
Next i

Close #iFileNoSim2


End Sub

Private Sub Command2_Click()


For i = 0 To 4
Text1(i).BackColor = &HFFFFFF
Next i

For i = 0 To 1
Text2(i).BackColor = &HFFFFFF
Next i

For i = 0 To 6
Text3(i).BackColor = &HFFFFFF
Next i

For i = 0 To 2
Text4(i).BackColor = &HFFFFFF
Next i

For i = 0 To 5
Text5(i).BackColor = &HFFFFFF
Next i

For i = 0 To 9
Text6(i).BackColor = &HFFFFFF
Next i

Combo1.BackColor = &HFFFFFF
Combo2.BackColor = &HFFFFFF
Combo3.BackColor = &HFFFFFF
Combo4.BackColor = &HFFFFFF
Combo5.BackColor = &HFFFFFF
Combo6.BackColor = &HFFFFFF

'checking the CASTID
If (Text1(4).Text = "") Then
Text1(4).BackColor = &H8080FF
MsgBox "The Cast_ID has not been determined"
Text1(4).SetFocus
Exit Sub
End If





'checking if the size of the system, length and thickness of the caster specified in the form
p = 0
For i = 0 To 4
If (Text1(i).Text = "") Then
Text1(i).BackColor = &H8080FF
p = 1
End If
Next i


For i = 0 To 1
If (Text2(i).Text = "") Then
Text2(i).BackColor = &H8080FF
p = 1
End If
Next i

For i = 0 To 6
If (Text3(i).Text = "") Then
Text3(i).BackColor = &H8080FF
p = 1
End If
Next i

For i = 0 To 2
If (Text4(i).Text = "") Then
Text4(i).BackColor = &H8080FF
p = 1
End If
Next i


For i = 0 To 5
If (Text5(i).Text = "") Then
Text5(i).BackColor = &H8080FF
p = 1
End If
Next i


For i = 0 To 9
If (Text6(i).Text = "") Then
Text6(i).BackColor = &H8080FF
p = 1
End If
Next i

'checking if the boundary condition and the setting of the simulation are specified

If (Combo1.Text = "Alloy Data") Then
Combo1.BackColor = &H8080FF
p = 1
End If

If (Combo3.Text = "Boundary Cond.") Then
Combo3.BackColor = &H8080FF
p = 1
End If


If (Combo4.Text = "Given Term") Then
Combo4.BackColor = &H8080FF
p = 1
End If


If (Combo5.Text = "Solver") Then
Combo5.BackColor = &H8080FF
p = 1
End If

If (p = 1) Then
'Form2.Label1.Caption = "please make sure you have filled all the fields specified with star"
'Form2.Visible = True
MsgBox "The Fields that are specified should be determined"
Exit Sub
ElseIf (Check1.Value <> 1) Then
MsgBox "The site infotmation has not been submitted"
Exit Sub

End If



If Dir(App.Path & "\Results\" & Text1(4).Text, vbDirectory) <> "" Then
    If (Check2.Value = 0) Then
    Load Form5
    Form5.Visible = True
    Form5.Label1.Caption = "Cast_ID " & Text1(4).Text & " already exists. Do you want to overwrite it?"
    Exit Sub
    End If
    
End If


Command1.Enabled = True




End Sub

Private Sub Command3_Click()
Dim CastID
CastID = Text1(4).Text


'making directory with the name of the castID
If Dir(App.Path & "\Results\" & CastID, vbDirectory) = "" Then
    MkDir (App.Path & "\Results\" & CastID)
End If

'copying the results files to the directory
If Dir(App.Path & "\FINALREPORT.DAT") <> "" Then
FileCopy App.Path & "\FINALREPORT.DAT", App.Path & "\Results\" & Text1(4).Text & "\FINALREPORT.DAT"
Kill (App.Path & "\FINALREPORT.DAT")
Else
MsgBox "The FINALREPORT.DAT file has not produced. Re-Run the simulation"
Exit Sub
End If

If Dir(App.Path & "\RESTART.DAT") <> "" Then
FileCopy App.Path & "\RESTART.DAT", App.Path & "\Results\" & Text1(4).Text & "\RESTART.DAT"
Else
MsgBox "The RESTART.DAT file has not produced. Re-Run the simulation"
Exit Sub
End If



Form6.Show
End Sub


Private Sub Command4_Click()
For i = 0 To 6
    Text3(i).Enabled = True
Next i
End Sub

Private Sub Command5_Click()
For i = 0 To 6
    Text3(i).Enabled = False
Next i



Dim newfileCast2



Dim iFileNoCast2 As Integer
iFileNoCast2 = FreeFile
Open App.Path & "\Temporary\Cast.DAT" For Output As #iFileNoCast2




For i = 0 To 6
Print #iFileNoCast2, Text3(i).Text
Next i




Close #iFileNoCast2



End Sub

Private Sub Command6_Click()
For i = 0 To 2
    Text4(i).Enabled = False
Next i

For i = 0 To 5
    Text5(i).Enabled = False
Next i



Dim newfileCal2



Dim iFileNoCal2 As Integer
iFileNoCal2 = FreeFile
Open App.Path & "\Temporary\CalConst.DAT" For Output As #iFileNoCal2




For i = 0 To 2
Print #iFileNoCal2, Text4(i).Text
Next i

For i = 0 To 5
Print #iFileNoCal2, Text5(i).Text
Next i


Close #iFileNoCal2
End Sub

Private Sub Command7_Click()
For i = 0 To 2
    Text4(i).Enabled = True
Next i
For i = 0 To 5
    Text5(i).Enabled = True
Next i
End Sub

Private Sub Command8_Click()
If Dir(App.Path & "\Temporary", vbDirectory) <> "" Then
Kill (App.Path & "\Temporary\*.*")
    RmDir (App.Path & "\Temporary")
End If
Unload Me

Form2.Visible = True


End Sub

Private Sub Command9_Click()
For i = 0 To 9
    Text6(i).Enabled = True
Next i
End Sub

Private Sub Form_Load()
'Simplifying commands
'------------------------------------------
If (Form2.Check1.Value = 0) Then
Shape1.Height = 7360
Frame5.Visible = False

For i = 0 To 2
Text4(i).Visible = False
Next i

For i = 0 To 5
Text5(i).Visible = False
Next i

Label17.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False

Command6.Visible = False
Command7.Visible = False

Combo5.Visible = False


Combo3.Left = 480
Combo3.Top = 360
Combo4.Left = 480
Combo4.Top = 720
Combo2.Left = 2160
Combo2.Top = 360
Check1.Visible = False
Check1.Top = 360
Check1.Left = 120

For i = 4 To 6
Text3(i).Visible = False
Next i

Label11.Visible = False
Label12.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False

Frame4.Top = 5060
Frame4.Height = 1210

Command1.Top = 6480
Command2.Top = 6480
Command3.Top = 6480
Command8.Top = 6480

Frame3.Top = 3500
Frame3.Height = 1450

Label9.Left = 3840
Label9.Top = 360
Label10.Left = 3960
Label10.Top = 720
Label23(4).Left = 6240
Label23(4).Top = 360
Label23(5).Left = 6240
Label23(5).Top = 720

Command4.Left = 5280
Command4.Top = 1080
Command5.Left = 6360
Command5.Top = 1080


Text3(2).Left = 5160
Text3(2).Top = 360
Text3(3).Left = 5160
Text3(3).Top = 720

Text1(0).Enabled = False
Text1(1).Enabled = False

Form1.Height = 7000
VScroll1.Visible = False

Form1.Left = 4000
Form1.Top = 1000

Form1.Width = 8200
Shape1.Width = Form1.Width
Shape1.Height = Form1.Height




End If
'-----------------------------------------
Check2.Value = 0

If Dir(App.Path & "\Temporary", vbDirectory) = "" Then
    MkDir (App.Path & "\Temporary")
End If

If Dir(App.Path & "\Results", vbDirectory) = "" Then
    MkDir (App.Path & "\Results")
End If

Image1.Picture = LoadPicture(App.Path & "\Files\logo.jpg")
Image2.Picture = LoadPicture(App.Path & "\Files\NLM.jpg")


Label22.Caption = "FDM Heat Transfer Analysis for Twin Belt Caster"
Label22.FontSize = 11
Label22.FontBold = True



Frame1.FontBold = True
Command1.Enabled = False
Command3.Enabled = False
For i = 0 To 4
Text1(i).Text = ""
Next i

Text1(0).Text = 100
Text1(1).Text = 25

Text1(2).Text = 0.75
Text1(3).Text = 0.01



Frame2.FontBold = True
Combo1.Text = "Alloy Data"
Combo1.AddItem "Define a new Alloy", 0
Combo1.AddItem "Load an Alloy", 1
Combo6.Visible = False
Text2(2).Visible = False
Label33.Visible = False




For i = 0 To 1
Text2(i).Text = ""
Next i



Frame3.FontBold = True

For i = 0 To 6
Text3(i).Text = ""
Next i



'Loading default cast condition values (frame 3)
If Dir(App.Path & "\Files\Cast.DAT") <> "" Then

FileCopy App.Path & "\Files\Cast.DAT", App.Path & "\Temporary\Cast.DAT"
 
Dim newfileCast
Dim iFileNoCast As Integer
Dim ppCast
iFileNoCast = FreeFile


Open App.Path & "\Files\Cast.DAT" For Input As #iFileNoCast

For i = 0 To 6
    Line Input #iFileNoCast, ppCast
    Text3(i).Text = ppCast
    Text3(i).Enabled = False
Next i

Close #iFileNoCast


End If


Frame4.FontBold = True
For i = 0 To 2
Text4(i).Text = ""
Next i

For i = 0 To 5
Text5(i).Text = ""
Next i

'Loading default Calculation conditions (frame 4)
If Dir(App.Path & "\Files\CalConst.DAT") <> "" Then

FileCopy App.Path & "\Files\CalConst.DAT", App.Path & "\Temporary\CalConst.DAT"
Dim newfileCal
Dim iFileNoCal As Integer
Dim ppCal
iFileNoCal = FreeFile


Open App.Path & "\Files\CalConst.DAT" For Input As #iFileNoCal

For i = 0 To 2
    Line Input #iFileNoCal, ppCal
    Text4(i).Text = ppCal
    Text4(i).Enabled = False
Next i


For i = 0 To 5
    Line Input #iFileNoCal, ppCal
    Text5(i).Text = ppCal
    Text5(i).Enabled = False
Next i

Close #iFileNoCal


End If



Combo2.Visible = False
Combo2.AddItem "Load Site Data", 0
Combo2.AddItem "New Site Data", 1



If (Form2.Check1.Value = 0) Then
Combo3.Text = "Boundary Cond."
Combo3.AddItem "Boundary Cond.", 0
Combo3.AddItem "Heat Flux", 1


Combo4.Text = "Given Term"
Combo4.AddItem "Given Term", 0
Combo4.AddItem "Exit Temp.", 1
Combo4.AddItem "Casting Speed", 2

Else
Combo3.Text = "Boundary Cond."
Combo3.AddItem "Boundary Cond.", 0
Combo3.AddItem "Heat Flux", 1
Combo3.AddItem "h.t.c", 2

Combo4.Text = "Given Term"
Combo4.AddItem "Given Term", 0
Combo4.AddItem "Exit Temp.", 1
Combo4.AddItem "Casting Speed", 2

End If


'Simplifying commands
'------------------------------------------
If (Form2.Check1.Value = 0) Then
Combo5.Text = "Parabolic"
Combo5.AddItem "Parabolic", 0
Else
Combo5.Text = "Solver"
Combo5.AddItem "Solver", 0
Combo5.AddItem "Elliptic", 1
Combo5.AddItem "Parabolic", 2
End If


Frame5.FontBold = True


'Loading default Calculation conditions (frame 4)
If Dir(App.Path & "\Files\SimConst.DAT") <> "" Then

FileCopy App.Path & "\Files\SimConst.DAT", App.Path & "\Temporary\SimConst.DAT"
Dim newfileSim
Dim iFileNoSim As Integer
Dim ppSim
iFileNoSim = FreeFile


Open App.Path & "\Files\SimConst.DAT" For Input As #iFileNoSim



For i = 0 To 9
    Line Input #iFileNoSim, ppSim
    Text6(i).Text = ppSim
    Text6(i).Enabled = False
Next i

Close #iFileNoSim


End If



VScroll1.Min = 0
VScroll1.Max = 100
VScroll1.Value = 0







End Sub



Private Sub Frame6_DragDrop(Source As Control, X As Single, Y As Single)

End Sub



Private Sub VScroll1_Change()
'Label22.Top = 480 - VScroll1.Value
'Frame1.Top = 1080 - VScroll1.Value
'Frame2.Top = 2520 - VScroll1.Value
'Frame3.Top = 3960 - VScroll1.Value
'Frame4.Top = 6360 - VScroll1.Value
'Frame5.Top = 9240 - VScroll1.Value
'Command1.Top = 10800 - VScroll1.Value
'Command2.Top = 10800 - VScroll1.Value
Form1.Top = -10 * VScroll1.Value
End Sub
