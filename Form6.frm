VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3420
   ClientLeft      =   110
   ClientTop       =   540
   ClientWidth     =   12920
   LinkTopic       =   "Form6"
   ScaleHeight     =   3420
   ScaleWidth      =   12920
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1570
      Left            =   600
      TabIndex        =   9
      Top             =   1800
      Width           =   4690
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   220
         Index           =   3
         Left            =   1560
         ScaleHeight     =   200.803
         ScaleMode       =   0  'User
         ScaleWidth      =   230
         TabIndex        =   45
         Top             =   960
         Width           =   250
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   220
         Index           =   2
         Left            =   960
         ScaleHeight     =   200.803
         ScaleMode       =   0  'User
         ScaleWidth      =   230
         TabIndex        =   44
         Top             =   600
         Width           =   250
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   220
         Index           =   0
         Left            =   480
         ScaleHeight     =   200.803
         ScaleMode       =   0  'User
         ScaleWidth      =   230
         TabIndex        =   11
         Top             =   240
         Width           =   250
         Begin VB.Line Line4 
            BorderColor     =   &H80000006&
            BorderStyle     =   4  'Dash-Dot
            Index           =   0
            Visible         =   0   'False
            X1              =   1320
            X2              =   1320
            Y1              =   0
            Y2              =   240.964
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   220
         Index           =   1
         Left            =   480
         ScaleHeight     =   200.803
         ScaleMode       =   0  'User
         ScaleWidth      =   230
         TabIndex        =   10
         Top             =   720
         Width           =   250
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         DrawMode        =   1  'Blackness
         Index           =   0
         X1              =   1560
         X2              =   4200
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         DrawMode        =   1  'Blackness
         Index           =   1
         X1              =   1560
         X2              =   4200
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         DrawMode        =   1  'Blackness
         Index           =   2
         X1              =   1560
         X2              =   4200
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         DrawMode        =   1  'Blackness
         Index           =   3
         X1              =   1560
         X2              =   4200
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Height          =   250
         Index           =   0
         Left            =   6240
         TabIndex        =   43
         Top             =   2040
         Width           =   300
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Height          =   250
         Index           =   1
         Left            =   6960
         TabIndex        =   42
         Top             =   2040
         Width           =   300
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Height          =   250
         Index           =   2
         Left            =   7680
         TabIndex        =   41
         Top             =   1920
         Width           =   370
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Height          =   250
         Index           =   3
         Left            =   8160
         TabIndex        =   40
         Top             =   1920
         Width           =   370
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Height          =   250
         Index           =   4
         Left            =   8640
         TabIndex        =   39
         Top             =   1920
         Width           =   370
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Height          =   250
         Index           =   5
         Left            =   9240
         TabIndex        =   38
         Top             =   1920
         Width           =   370
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Height          =   250
         Index           =   6
         Left            =   9840
         TabIndex        =   37
         Top             =   1920
         Width           =   370
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   0
         Left            =   6240
         TabIndex        =   36
         Top             =   2520
         Width           =   300
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   1
         Left            =   6840
         TabIndex        =   35
         Top             =   2520
         Width           =   300
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   2
         Left            =   7680
         TabIndex        =   34
         Top             =   2520
         Width           =   300
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   3
         Left            =   8400
         TabIndex        =   33
         Top             =   2640
         Width           =   300
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   4
         Left            =   8880
         TabIndex        =   32
         Top             =   2640
         Width           =   300
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   5
         Left            =   9360
         TabIndex        =   31
         Top             =   2520
         Width           =   300
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   6
         Left            =   9840
         TabIndex        =   30
         Top             =   2400
         Width           =   300
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Temperature  (C)"
         Height          =   250
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   1690
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Solid Fraction"
         Height          =   250
         Left            =   3840
         TabIndex        =   28
         Top             =   1800
         Width           =   1810
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   0
         Left            =   0
         TabIndex        =   27
         Top             =   2280
         Width           =   370
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   1
         Left            =   600
         TabIndex        =   26
         Top             =   2280
         Width           =   370
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   2
         Left            =   1320
         TabIndex        =   25
         Top             =   2280
         Width           =   370
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   3
         Left            =   1800
         TabIndex        =   24
         Top             =   2280
         Width           =   370
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   4
         Left            =   2280
         TabIndex        =   23
         Top             =   2280
         Width           =   370
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   5
         Left            =   2640
         TabIndex        =   22
         Top             =   2280
         Width           =   370
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   0
         Left            =   0
         TabIndex        =   21
         Top             =   1680
         Width           =   370
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   1
         Left            =   360
         TabIndex        =   20
         Top             =   1680
         Width           =   370
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   2
         Left            =   720
         TabIndex        =   19
         Top             =   1560
         Width           =   370
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   3
         Left            =   1200
         TabIndex        =   18
         Top             =   1680
         Width           =   370
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   4
         Left            =   1680
         TabIndex        =   17
         Top             =   1800
         Width           =   370
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Height          =   370
         Index           =   5
         Left            =   2280
         TabIndex        =   16
         Top             =   1680
         Width           =   370
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   250
         Index           =   0
         Left            =   8520
         TabIndex        =   15
         Top             =   0
         Width           =   610
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Height          =   250
         Index           =   0
         Left            =   7920
         TabIndex        =   14
         Top             =   0
         Width           =   300
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Height          =   250
         Index           =   0
         Left            =   8520
         TabIndex        =   13
         Top             =   480
         Width           =   610
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Height          =   250
         Index           =   0
         Left            =   7920
         TabIndex        =   12
         Top             =   480
         Width           =   300
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   250
      Left            =   3360
      TabIndex        =   8
      Top             =   840
      Width           =   2410
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1930
      Left            =   6840
      TabIndex        =   7
      Top             =   360
      Width           =   250
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Text            =   "690"
      Top             =   480
      Width           =   490
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Text            =   "300"
      Top             =   120
      Width           =   490
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   3480
      TabIndex        =   1
      Text            =   "10"
      Top             =   120
      Width           =   850
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Plot"
      Height          =   250
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Width           =   2410
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   370
      Left            =   720
      TabIndex        =   46
      Top             =   1440
      Visible         =   0   'False
      Width           =   9490
   End
   Begin VB.Image Image2 
      Height          =   1210
      Left            =   8880
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1210
   End
   Begin VB.Image Image1 
      Height          =   1210
      Left            =   240
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1210
   End
   Begin VB.Label Label3 
      Caption         =   "Temp. Max. (C)"
      Height          =   250
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   1090
   End
   Begin VB.Label Label2 
      Caption         =   "Temp. Min. (C)"
      Height          =   250
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   1210
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Contours"
      Height          =   250
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   1930
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If (Text1.Text > Text2.Text) Then
    MsgBox "Error11: Maximum temperature is smaller than minimum temperature "
    Else
    
    ' number of contour levels
nc = Form6.Text3.Text
If (nc < 3 Or nc > 50) Then
    MsgBox "Error10: Chenge the No of contours"
'    Form5.Show
Else



Form6.Picture1(0).Visible = True
Form6.Picture1(1).Visible = True
Form6.Label6.Visible = True
Form6.Label7.Visible = True
Form6.Frame1.Visible = True
Form6.Frame1.Caption = Results
Form6.Command2.Enabled = True

    Form6.Picture1(0).Cls
    Form6.Picture1(1).Cls
    Form6.Picture1(2).Cls
    Form6.Picture1(3).Cls

Dim eps As Double
eps = 0.01
Dim SlabLength As Double
Dim SlabHeight As Double
SlabLength = Form1.Text1(2).Text * 1.2
SlabHeight = Form1.Text1(3).Text

dLength = SlabLength * 1000 / 5
dHeight = SlabHeight * 1000 / 5



    




 Label14.Visible = True
 Label14.Caption = "Note! The computational domain is 20% longer than caster length. All length units are in mm"
    Dim pp As Integer
    pp = 255 / nc - 1

    Dim i, j As Integer
    Dim newfile6
    Dim iFileNo6 As Integer
    iFileNo6 = FreeFile
    Dim Nxy(0 To 1) As Integer

    
    TempMin = 1000
    TempMax = 0
    Form6.VScroll1.Visible = True







Open App.Path & "/RESTART.DAT" For Input As #iFileNo6
Input #iFileNo6, Nxy(0), Nxy(1)


    Dim ScaleX, ScaleY, sizeX, sizeY As Integer
    sizeX = 100
    sizeY = 25
    ScaleX = 1
    ScaleY = 1
    P0Top = 1000
    P0Left = 400
    Form6.Picture1(0).Top = P0Top
    Form6.Picture1(0).Left = P0Left
    Form6.Picture1(2).Visible = False
    Form6.Picture1(3).Visible = False
    Form6.Picture1(0).Height = 100 * sizeY * ScaleY
    Form6.Picture1(0).Width = 100 * sizeX * ScaleX
    
    Form6.Picture1(1).Left = Form6.Picture1(0).Left
    Form6.Picture1(1).Top = Form6.Picture1(0).Top + Form6.Picture1(0).Height + 1500
    Form6.Picture1(1).Height = 100 * sizeY * ScaleY
    Form6.Picture1(1).Width = 100 * sizeX * ScaleX


    Form6.Height = Form6.Picture1(1).Top + 2 * Form6.Picture1(1).Height + 1000
    Form6.Width = Form6.Picture1(0).Left + Form6.Picture1(0).Width + 4000
    Image2.Left = Form6.Width - 3000
    
    Form6.Frame1.Height = Form6.Picture1(1).Top + 2 * Form6.Picture1(1).Height - 2000
    Form6.Frame1.Width = Form6.Picture1(0).Left + Form6.Picture1(0).Width + 2000
    
    Form6.Label6.Left = Form6.Picture1(0).Left + Form6.Picture1(0).Width / 2 - Form6.Label6.Width / 2
    Form6.Label6.Top = Form6.Picture1(0).Top - 400
    Form6.Label6.FontBold = True
    Form6.Label6.FontSize = 10
    
    Form6.Label7.Left = Form6.Picture1(1).Left + Form6.Picture1(1).Width / 2 - Form6.Label7.Width / 2
    Form6.Label7.Top = Form6.Picture1(1).Top - 400
    Form6.Label7.FontBold = True
    Form6.Label7.FontSize = 10
    
    
    
    Form6.Line1(0).x1 = Form6.Picture1(0).Left
    Form6.Line1(0).y1 = Form6.Picture1(0).Top + Form6.Picture1(0).Height
    
    Form6.Line1(0).x2 = Form6.Line1(0).x1 + Form6.Picture1(0).Width + 500
    Form6.Line1(0).y2 = Form6.Line1(0).y1
    
    Form6.Line1(1).x1 = Form6.Line1(0).x1
    Form6.Line1(1).y1 = Form6.Line1(0).y1
    
    Form6.Line1(1).x2 = Form6.Line1(0).x1
    Form6.Line1(1).y2 = Form6.Line1(1).y1 - Form6.Picture1(0).Height - 500
    
    
    Form6.Line1(2).x1 = Form6.Picture1(1).Left
    Form6.Line1(2).y1 = Form6.Picture1(1).Top + Form6.Picture1(1).Height
    
    Form6.Line1(2).x2 = Form6.Line1(1).x1 + Form6.Picture1(1).Width + 500
    Form6.Line1(2).y2 = Form6.Line1(2).y1
    
    Form6.Line1(3).x1 = Form6.Line1(2).x1
    Form6.Line1(3).y1 = Form6.Line1(2).y1
    
    Form6.Line1(3).x2 = Form6.Line1(2).x1
    Form6.Line1(3).y2 = Form6.Line1(3).y1 - Form6.Picture1(1).Height - 500
    
    
    
    
    
    
    
    
    Dim dx, dy As Integer
    dx = Form6.Picture1(0).Width / 5
    dy = Form6.Picture1(0).Height / 5
    
    Lx = sizeX / 5
    Ly = sizeY / 5
    
    For c = 0 To 5
    Form6.Label4(c).Height = 270
    Form6.Label4(c).Width = 270
    Form6.Label4(c).Left = Form6.Picture1(0).Left + c * dx - Form6.Label4(c).Width / 2
    Form6.Label4(c).Top = Form6.Line1(0).y1 + 100
    Form6.Label4(c).Caption = c * dLength
    Next c
    
    For c = 0 To 5
    Form6.Label5(c).Height = 270
    Form6.Label5(c).Width = 250
    Form6.Label5(c).Left = Form6.Picture1(0).Left - 300
    Form6.Label5(c).Top = Form6.Line1(0).y1 - dy * c - Form6.Label5(c).Height / 2
    Form6.Label5(c).Caption = c * dHeight
    Next c
    
    For c = 0 To 5
    Form6.Label8(c).Height = 270
    Form6.Label8(c).Width = 270
    Form6.Label8(c).Left = Form6.Picture1(1).Left + c * dx - Form6.Label8(c).Width / 2
    Form6.Label8(c).Top = Form6.Line1(2).y1 + 100
    Form6.Label8(c).Caption = c * dLength
    Next c
    
    For c = 0 To 5
    Form6.Label9(c).Height = 270
    Form6.Label9(c).Width = 250
    Form6.Label9(c).Left = Form6.Picture1(1).Left - 300
    Form6.Label9(c).Top = Form6.Line1(2).y1 - dy * c - Form6.Label9(c).Height / 2
    Form6.Label9(c).Caption = c * dHeight
    Next c
    

    
    

    Form6.VScroll1.Left = Form6.Left + Form6.Width - 2 * Form6.VScroll1.Width
    Form6.VScroll1.Height = Form6.Height - 200
    Form6.VScroll1.Top = Form6.Top + 100
    
    Form6.Picture1(0).Refresh
    Form6.Picture1(1).Refresh
    
Dim BOTH() As Double
Dim TEMP() As Double
Dim FRAC() As Double


    
ReDim BOTH(2 * Nxy(1), Nxy(0))
ReDim TEMP(Nxy(1), Nxy(0))
ReDim FRAC(Nxy(1), Nxy(0))
For j = 1 To 2 * Nxy(1)
    For i = 1 To Nxy(0)
    Input #iFileNo6, BOTH(j, i)
    Next i
Next j


'dash lines for the first picture box
    For c = 0 To 3
    Form6.Picture1(0).DrawStyle = vbDot
    Form6.Picture1(1).DrawStyle = vbDot
    Form6.Picture1(0).Line (0, (c + 1) * dy)-(Form6.Picture1(0).Width, (c + 1) * dy), RGB(150, 150, 150)
    Form6.Picture1(0).Line ((c + 1) * dx, 0)-((c + 1) * dx, Form6.Picture1(0).Height), RGB(150, 150, 150)
    Form6.Picture1(1).Line (0, (c + 1) * dy)-(Form6.Picture1(1).Width, (c + 1) * dy), RGB(150, 150, 150)
    Form6.Picture1(1).Line ((c + 1) * dx, 0)-((c + 1) * dx, Form6.Picture1(1).Height), RGB(150, 150, 150)
    Next c
    

    

'producing labels for the temp part
Dim L As Integer
'Label10(0).Caption = "0"
Label10(0).FontBold = True
'Label11(0).Caption = "0"
Label11(0).FontBold = True
Label10(0).Left = Form6.Picture1(0).Left + Form6.Picture1(0).Width + 1000
Label11(0).Left = Form6.Picture1(0).Left + Form6.Picture1(0).Width + 600
Label10(0).Top = Form6.Picture1(0).Top + Form6.Picture1(0).Height - Form6.Label10(0).Height
Label11(0).Top = Label10(0).Top
 
 
'Label12(0).Caption = "0"
Label12(0).FontBold = True
'Label13(0).Caption = "0"
Label13(0).FontBold = True
Label12(0).Left = Form6.Picture1(1).Left + Form6.Picture1(1).Width + 1000
Label13(0).Left = Form6.Picture1(1).Left + Form6.Picture1(1).Width + 600
Label12(0).Top = Form6.Picture1(1).Top + Form6.Picture1(1).Height - Form6.Label12(0).Height
Label13(0).Top = Label12(0).Top
'delete the labels if they are already there

  LL = 0
  Dim ctl As Control
    For Each ctl In Controls
        If ctl.Name = "Label10" Then
        LL = LL + 1
        End If
    Next
    
      LL1 = 0
  Dim ctl1 As Control
    For Each ctl1 In Controls
        If ctl1.Name = "Label11" Then
        LL1 = LL1 + 1
        End If
    Next
    
    
      LL2 = 0
  Dim ctl2 As Control
    For Each ctl2 In Controls
        If ctl2.Name = "Label12" Then
        LL2 = LL2 + 1
        End If
    Next
    
      LL3 = 0
  Dim ctl3 As Control
    For Each ctl3 In Controls
        If ctl3.Name = "Label13" Then
        LL3 = LL3 + 1
        End If
    Next
    
    
    
'unloading the existing array of labels
    If LL > 1 Then
    For KK = 1 To LL - 1
    Unload Label10(KK)
    Next KK
    End If
    
    If LL1 > 1 Then
    For KK = 1 To LL1 - 1
    Unload Label11(KK)
    Next KK
    End If
    
    
    If LL2 > 1 Then
    For KK = 1 To LL2 - 1
    Unload Label12(KK)
    Next KK
    End If
    
    If LL3 > 1 Then
    For KK = 1 To LL3 - 1
    Unload Label13(KK)
    Next KK
    End If
    
    
  
'loading as many labels as neede
'temperature labels
  For L = 1 To nc
    Load Label10(L)
    With Label10(L)
'      .Caption = L
      .Visible = True
      .Top = Label10(L - 1).Top - 250
    End With
  Next L
  
  
    For L = 1 To nc
    Load Label11(L)
    With Label11(L)
'      .Caption = L
      .Visible = True
      .Top = Label11(L - 1).Top - 250
    End With
  Next L
  
    For L = 1 To nc
    Load Label12(L)
    With Label12(L)
'      .Caption = L
      .Visible = True
      .Top = Label12(L - 1).Top - 250
    End With
  Next L
  
  
    For L = 1 To nc
    Load Label13(L)
    With Label13(L)
'      .Caption = L
      .Visible = True
      .Top = Label13(L - 1).Top - 250
    End With
  Next L
  
  
  
  


' matrix of data to contour (Temperature)
For j = 1 To Nxy(1)
    For i = 1 To Nxy(0)
    TEMP(j, i) = BOTH(2 * j - 1, i)
        If (TEMP(j, i) < TempMin) Then
         TempMin = TEMP(j, i)
         ElseIf (TEMP(j, i) > TempMax) Then
         TempMax = TEMP(j, i)
        End If
    Next i
Next j

TempMin = Form6.Text1.Text
TempMax = Form6.Text2.Text

For j = 1 To Nxy(1)
    For i = 1 To Nxy(0)
    FRAC(j, i) = BOTH(2 * j, i)
        If (FRAC(j, i) < FracMin) Then
         FracMin = FRAC(j, i)
         ElseIf (FRAC(j, i) > FracMax) Then
         FracMax = FRAC(j, i)
        End If
    Next i
Next j


FracMin = 0
FracMax = 1





 
 
' index bounds of data matrix (x-lower,x-upper,y-lower,y-upper)
 ilb = 0
 iub = Nxy(1)
 jlb = 0
 jub = Nxy(0)

' data matrix column coordinates
Dim X() As Integer
ReDim X(Nxy(1))

For i = 0 To Nxy(1) - 1
    X(i + 1) = i * sizeY / Nxy(1)
Next i

' data matrix row coordinates
Dim Y() As Integer
ReDim Y(Nxy(0))

For i = 0 To Nxy(0) - 1
    Y(i + 1) = i * sizeX / Nxy(0)
Next i

' contour(#)

dT = (TempMax - TempMin) / nc

' contour levels in increasing order
Dim contour() As Double
ReDim contour(nc + 1)
contour(0) = TempMin

For i = 1 To nc
    contour(i) = contour(i - 1) + dT
Next i



Close #iFileNo6

picNum = 0
AAA = conrec(TEMP(), X(), Y(), nc, contour(), ilb, iub, jlb, jub, dT, ScaleX, ScaleY, pp, picNum)


' contour Fraction(#)

df = (FracMax - FracMin) / nc

' contour levels in increasing order
Dim contourF() As Double
ReDim contourF(nc + 1)
contourF(0) = FracMin

For i = 1 To nc
    contourF(i) = contourF(i - 1) + df
Next i

contourF(0) = contourF(0) + eps
contourF(nc) = contourF(nc) - eps


picNum = 1
AAA = conrec(FRAC(), X(), Y(), nc, contourF(), ilb, iub, jlb, jub, df, ScaleX, ScaleY, pp, picNum)

End If
End If

End Sub












Public Function conrec(Z() As Double, X() As Integer, Y() As Integer, ByVal nc As Integer, _
                         contour() As Double, ByVal ilb As Integer, ByVal iub As Integer, _
                         ByVal jlb As Integer, ByVal jub As Integer, ByVal d As Double, _
                         ByVal ScaleX As Integer, ByVal ScaleY As Integer, ByVal pp As Integer, ByVal picNum As Integer)



    Dim m1, m2, m3, case_value As Integer
    Dim dmin, dmax As Double
    Dim x1, x2, y1, y2 As Double
    Dim k, m As Integer
    Dim h(5) As Double
    Dim sh(5) As Integer
    Dim xh(5), yh(5) As Double
    Dim R As Double
    Dim ColorCtrlR, ColorCtrlB, ColorCtrlG As Integer
    
    Dim Nullcode As Double
    Nullcode = 1E+37
    
    
    Dim im(4), jm(4) As Integer
    im(0) = 0
    im(1) = 1
    im(2) = 1
    im(3) = 0
    jm(0) = 0
    jm(1) = 0
    jm(2) = 1
    jm(3) = 1




Dim castab(3, 3, 3) As Integer
castab(0, 0, 0) = 0
castab(0, 0, 1) = 0
castab(0, 0, 2) = 8
castab(0, 1, 0) = 0
castab(0, 1, 1) = 2
castab(0, 1, 2) = 5
castab(0, 2, 0) = 7
castab(0, 2, 1) = 6
castab(0, 2, 2) = 9
castab(1, 0, 0) = 0
castab(1, 0, 1) = 3
castab(1, 0, 2) = 4
castab(1, 1, 0) = 1
castab(1, 1, 1) = 3
castab(1, 1, 2) = 1
castab(1, 2, 0) = 4
castab(1, 2, 1) = 3
castab(1, 2, 2) = 0
castab(2, 0, 0) = 9
castab(2, 0, 1) = 6
castab(2, 0, 2) = 7
castab(2, 1, 0) = 5
castab(2, 1, 1) = 2
castab(2, 1, 2) = 0
castab(2, 2, 0) = 8
castab(2, 2, 1) = 0
castab(2, 2, 2) = 0






If nc <> 0 Then
For j = jub - 1 To jlb Step -1
  For i = ilb To iub - 1
       Dim temp1, temp2 As Double
       temp1 = Min(Z(i, j), Z(i, j + 1))
       temp2 = Min(Z(i + 1, j), Z(i + 1, j + 1))
       dmin = Min(temp1, temp2)
       temp1 = Max(Z(i, j), Z(i, j + 1))
       temp2 = Max(Z(i + 1, j), Z(i + 1, j + 1))
       dmax = Max(temp1, temp2)
      
'-------------------------------------------------------------------------
       'extra conditional added here to insure that large values are not plotted
       'if an area should not be contoured, values above nullcode should be entered in
       'the matrix Z
      
'------------------------------------------------------------------------
       If dmax >= contour(0) And dmin <= contour(nc) And dmax < Nullcode Then
         For k = 0 To nc
           If contour(k) >= dmin And contour(k) < dmax Then
             For m = 4 To 0 Step -1
               If (m > 0) Then
                 h(m) = Z(i + im(m - 1), j + jm(m - 1)) - contour(k)
                 xh(m) = X(i + im(m - 1))
                 yh(m) = Y(j + jm(m - 1))
               Else:
                 h(0) = 0.25 * (h(1) + h(2) + h(3) + h(4))
                 xh(0) = 0.5 * (X(i) + X(i + 1))
                 yh(0) = 0.5 * (Y(j) + Y(j + 1))
               End If
              If (h(m) > 0#) Then
                sh(m) = 1
              ElseIf h(m) < 0# Then
                sh(m) = -1
              Else:
                sh(m) = 0
              End If
            Next m
           
'=================================================================
            '
            ' Note: at this stage the relative heights of the corners and the
            ' centre are in the h array, and the corresponding coordinates are
            ' in the xh and yh arrays. The centre of the box is indexed by 0
            ' and the 4 corners by 1 to 4 as shown below.
            ' Each triangle is then indexed by the parameter m, and the 3
            ' vertices of each triangle are indexed by parameters m1,m2,and m3.
            ' It is assumed that the centre of the box is always vertex 2
            ' though this isimportant only when all 3 vertices lie exactly on
            ' the same contour level, in which case only the side of the box
            ' is drawn.
            '
            '
            '      vertex 4 +-------------------+ vertex 3
            '               | \               / |
            '               |   \    m-3    /   |
            '               |     \       /     |
            '               |       \   /       |
            '               |  m=2    X   m=2   |       the centre is vertex 0
            '               |       /   \       |
            '               |     /       \     |
            '               |   /    m=1    \   |
            '               | /               \ |
            '      vertex 1 +-------------------+ vertex 2
            '
            '
            '
            '               Scan each triangle in the box
            '
           
'=================================================================
             For m = 1 To 4
               m1 = m
               m2 = 0
               If (m <> 4) Then
                 m3 = m + 1
               Else:
                 m3 = 1
               End If
               case_value = castab(sh(m1) + 1, sh(m2) + 1, sh(m3) + 1)
               If case_value <> 0 Then
                 Select Case case_value
                 
'===========================================================
                  '     Case 1 - Line between vertices 1 and 2
                 
'===========================================================
                Case 1
                   x1 = xh(m1)
                   y1 = yh(m1)
                   x2 = xh(m2)
                   y2 = yh(m2)
                 
'===========================================================
                  '     Case 2 - Line between vertices 2 and 3
                 
'===========================================================
                 Case 2
                   x1 = xh(m2)
                   y1 = yh(m2)
                   x2 = xh(m3)
                   y2 = yh(m3)
                 
'===========================================================
                  '     Case 3 - Line between vertices 3 and 1
                 
'===========================================================
                 Case 3
                   x1 = xh(m3)
                   y1 = yh(m3)
                   x2 = xh(m1)
                   y2 = yh(m1)
                 
'===========================================================
                  '     Case 4 - Line between vertex 1 and side 2-3
                 
'===========================================================
                 Case 4
                   x1 = xh(m1)
                   y1 = yh(m1)
                   x2 = (h(m3) * xh(m2) - h(m2) * xh(m3)) / (h(m3) - h(m2))
                   y2 = (h(m3) * yh(m2) - h(m2) * yh(m3)) / (h(m3) - h(m2))
                   
                 
'===========================================================
                  '     Case 5 - Line between vertex 2 and side 3-1
                 
'===========================================================
                 Case 5
                   x1 = xh(m2)
                   y1 = yh(m2)
                   x2 = (h(m1) * xh(m3) - h(m3) * xh(m1)) / (h(m1) - h(m3))
                   y2 = (h(m1) * yh(m3) - h(m3) * yh(m1)) / (h(m1) - h(m3))
                 
'===========================================================
                  '     Case 6 - Line between vertex 3 and side 1-2
                 
'===========================================================
                 Case 6
                   x1 = xh(m3)
                   y1 = yh(m3)
                   x2 = (h(m2) * xh(m1) - h(m1) * xh(m2)) / (h(m2) - h(m1))
                   y2 = (h(m2) * yh(m1) - h(m1) * yh(m2)) / (h(m2) - h(m1))
                 
'===========================================================
                  '     Case 7 - Line between sides 1-2 and 2-3
                 
'===========================================================
                 Case 7
                   x1 = (h(m2) * xh(m1) - h(m1) * xh(m2)) / (h(m2) - h(m1))
                   y1 = (h(m2) * yh(m1) - h(m1) * yh(m2)) / (h(m2) - h(m1))
                   x2 = (h(m3) * xh(m2) - h(m2) * xh(m3)) / (h(m3) - h(m2))
                   y2 = (h(m3) * yh(m2) - h(m2) * yh(m3)) / (h(m3) - h(m2))
                   
                 
'===========================================================
                  '     Case 8 - Line between sides 2-3 and 3-1
                 
'===========================================================
                 Case 8
                   x1 = (h(m3) * xh(m2) - h(m2) * xh(m3)) / (h(m3) - h(m2))
                   y1 = (h(m3) * yh(m2) - h(m2) * yh(m3)) / (h(m3) - h(m2))
                   x2 = (h(m1) * xh(m3) - h(m3) * xh(m1)) / (h(m1) - h(m3))
                   y2 = (h(m1) * yh(m3) - h(m3) * yh(m1)) / (h(m1) - h(m3))
                 
'===========================================================
                  '     Case 9 - Line between sides 3-1 and 1-2
                 
'===========================================================
                 Case 9
                   x1 = (h(m1) * xh(m3) - h(m3) * xh(m1)) / (h(m1) - h(m3))
                   y1 = (h(m1) * yh(m3) - h(m3) * yh(m1)) / (h(m1) - h(m3))
                   x2 = (h(m2) * xh(m1) - h(m1) * xh(m2)) / (h(m2) - h(m1))
                   y2 = (h(m2) * yh(m1) - h(m1) * yh(m2)) / (h(m2) - h(m1))
                End Select
        '--------------------------------------------------------------
                'this is where the program specific drawing routine comes in.
                'This specific command will work well for a properly dimensioned
                'vb picture box or vb form (where "Form6" is the name of the form)
               
'-------------------------------------------------------------------
R = k / nc
If k < nc / 3 Then
ColorCtrlB = 254 * (1 - 3 * R)
ColorCtrlG = 254 * R
ColorCtrlR = 0
ElseIf k > 2 * nc / 3 Then
ColorCtrlB = 0
ColorCtrlG = 254 * (3 - 3 * R)
ColorCtrlR = 254
Else
ColorCtrlB = 0
ColorCtrlG = 254
ColorCtrlR = 254 * (2 - 3 * R)
End If

If ColorCtrlB > 254 Or ColorCtrlG > 254 Or ColorCtrlR > 254 Then
ColorCtrlB = 254
ColorCtrlG = 254
ColorCtrlR = 254
End If

If ColorCtrlB < 0 Or ColorCtrlG < 0 Or ColorCtrlR < 0 Then
ColorCtrlB = 254
ColorCtrlG = 254
ColorCtrlR = 254
End If







'ColorCtrl = 2 * pp * (k + 1)
'If (ColorCtrl < 254) Then
'ColorCtrlR = 0
'ColorCtrlY = ColorCtrl
'ColorCtrlB = ColorCtrl
'Else
'ColorCtrlR = ColorCtrl / 2
'ColorCtrlY = 255 - ColorCtrl / 2
'ColorCtrlB = 0
'End If





     Form6.Picture1(picNum).DrawWidth = 2
     Form6.Picture1(picNum).Line (105 * CSng(y2), 105 * CSng(x2))-(105 * CSng(y1), 105 * CSng(x1)), RGB(ColorCtrlR, ColorCtrlG, ColorCtrlB)
     If picNum = 0 Then
        Label11(k).BackColor = RGB(ColorCtrlR, ColorCtrlG, ColorCtrlB)
        Label11(k).Caption = ""
     End If
     
      If picNum = 1 Then
        Label13(k).BackColor = RGB(ColorCtrlR, ColorCtrlG, ColorCtrlB)
        Label13(k).Caption = ""
     End If
        
'-------------------------------------------------------------------
               End If
             Next m
          End If
          
        If picNum = 0 Then
        Label10(k).Caption = Format(contour(k), "###0.0")
        End If
        
        If picNum = 1 Then
        Label12(k).Caption = Format(contour(k), "###0.00")
        End If


 '       Label10(k).ZOrder 0
 '       Picture1(0).ZOrder 1
'       Set Label10(k).Container = Form6.Picture1(0)

        
        Next k
      End If
     Next i
  
'--------------------------------------------------------------------------------------
   'used to refresh the drawing surface after each row is contoured (for impatient users)
  ' Form6.Refresh
  
'-------------------------------------------------------------------------------------
   Next j

End If



End Function











Public Function Min(ByVal v1 As Double, ByVal v2 As Double) As Double
If (v1 < v2) Then
  Min = v1
Else: Min = v2
End If
End Function

Public Function Max(ByVal v1 As Double, ByVal v2 As Double) As Double
If (v1 < v2) Then
  Max = v2
Else: Max = v1
End If
End Function




Private Sub Command2_Click()
    Form6.Picture1(2).Visible = True
    Form6.Picture1(3).Visible = True

Dim namefile
namefile1 = (App.Path & "\Photos\Temperature.bmp")
namefile2 = (App.Path & "\Photos\SolidFraction.bmp")
SavePicture Picture1(0).Image, namefile1
SavePicture Picture1(1).Image, namefile2


Form6.Picture1(2).Left = Form6.Picture1(0).Left
Form6.Picture1(2).Top = Form6.Picture1(0).Top
Form6.Picture1(2).Width = Form6.Picture1(0).Width
Form6.Picture1(2).Height = Form6.Picture1(0).Height

Picture1(2).Picture = LoadPicture(App.Path & "\Photos\Temperature.bmp")

Form6.Picture1(3).Left = Form6.Picture1(1).Left
Form6.Picture1(3).Top = Form6.Picture1(1).Top
Form6.Picture1(3).Width = Form6.Picture1(1).Width
Form6.Picture1(3).Height = Form6.Picture1(1).Height

Picture1(3).Picture = LoadPicture(App.Path & "\Photos\SolidFraction.bmp")


With Frame1
 .Top = Me.ScaleTop
 .Left = Me.ScaleLeft
 .BorderStyle = 0
 .BackColor = vbWhite
 'for all text boxes
 'textbox1.BorderStyle = 0
 'textbox1.Appearance = 0
 'back style of all lable controls must be set to zero
 '.ZOrder vbBringToFront
 End With
 Me.Width = Me.Frame1.Width
 Me.Height = 2 * Me.Frame1.Height
 Me.PrintForm
 
 
End Sub





Private Sub Form_Load()


Image1.Picture = LoadPicture(App.Path & "\Files\logo.jpg")
Image2.Picture = LoadPicture(App.Path & "\Files\NLM.jpg")
Form6.Picture1(0).Visible = False
Form6.Picture1(1).Visible = False
Form6.VScroll1.Visible = False
Form6.Label6.Visible = False
Form6.Label7.Visible = False
Form6.Command2.Enabled = False

Picture1(0).AutoRedraw = True
Picture1(1).AutoRedraw = True

Form6.Frame1.Visible = False

  
End Sub

Private Sub VScroll1_Change()
Form6.Top = -0.2 * VScroll1.Value
End Sub
