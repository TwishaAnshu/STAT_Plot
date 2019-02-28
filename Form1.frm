VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8568
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   18696
   LinkTopic       =   "Form1"
   ScaleHeight     =   8568
   ScaleWidth      =   18696
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton intersect 
      Caption         =   "intersecting points"
      Height          =   672
      Left            =   15600
      TabIndex        =   50
      Top             =   6300
      Width           =   1092
   End
   Begin VB.CommandButton cmdEXTEND 
      Caption         =   "EXTEND"
      Height          =   912
      Index           =   1
      Left            =   15480
      TabIndex        =   47
      Top             =   3870
      Width           =   1332
   End
   Begin VB.CommandButton cmdPlot 
      Caption         =   "plotFrom"
      Height          =   912
      Left            =   15480
      TabIndex        =   46
      Top             =   2760
      Width           =   1332
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      Height          =   6732
      Left            =   7200
      ScaleHeight     =   6684
      ScaleWidth      =   7524
      TabIndex        =   37
      Top             =   1080
      Width           =   7572
      Begin VB.Label lblp4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblp4"
         Height          =   492
         Left            =   2460
         TabIndex        =   43
         Top             =   1380
         Visible         =   0   'False
         Width           =   792
      End
      Begin VB.Label lblp3 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblp3"
         Height          =   168
         Left            =   2100
         TabIndex        =   42
         Top             =   1080
         Visible         =   0   'False
         Width           =   396
      End
      Begin VB.Label lblp2 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblp2"
         Height          =   240
         Left            =   1260
         TabIndex        =   41
         Top             =   660
         Visible         =   0   'False
         Width           =   396
      End
      Begin VB.Label lblp1 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblp1"
         Height          =   240
         Left            =   720
         TabIndex        =   40
         Top             =   180
         Visible         =   0   'False
         Width           =   396
      End
   End
   Begin VB.TextBox txtY12 
      Height          =   312
      Left            =   5940
      TabIndex        =   33
      Top             =   960
      Width           =   672
   End
   Begin VB.TextBox txtX12 
      Height          =   372
      Left            =   4560
      TabIndex        =   32
      Top             =   900
      Width           =   912
   End
   Begin VB.TextBox txtY22 
      Height          =   408
      Left            =   6180
      TabIndex        =   30
      Top             =   1620
      Width           =   672
   End
   Begin VB.TextBox txtX22 
      Height          =   432
      Left            =   4680
      TabIndex        =   29
      Top             =   1620
      Width           =   792
   End
   Begin VB.TextBox txtY2 
      Height          =   432
      Left            =   2460
      TabIndex        =   19
      Top             =   1740
      Width           =   612
   End
   Begin VB.TextBox txtX2 
      Height          =   372
      Left            =   1380
      TabIndex        =   18
      Top             =   1680
      Width           =   672
   End
   Begin VB.TextBox txtY1 
      Height          =   432
      Left            =   2340
      TabIndex        =   16
      Top             =   960
      Width           =   612
   End
   Begin VB.TextBox txtX1 
      Height          =   372
      Left            =   1260
      TabIndex        =   15
      Top             =   900
      Width           =   672
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   912
      Left            =   15480
      TabIndex        =   2
      Top             =   4980
      Width           =   1332
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "clear"
      Height          =   912
      Left            =   15480
      TabIndex        =   1
      Top             =   1650
      Width           =   1332
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   912
      Left            =   15480
      TabIndex        =   0
      Top             =   540
      Width           =   1332
   End
   Begin VB.Label lblIntersect 
      Height          =   552
      Left            =   13080
      TabIndex        =   49
      Top             =   180
      Width           =   1032
   End
   Begin VB.Label Label 
      Caption         =   "POINT OF INTERSECTION"
      Height          =   432
      Index           =   12
      Left            =   10800
      TabIndex        =   48
      Top             =   240
      Width           =   1932
   End
   Begin VB.Label lblslope2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   3960
      TabIndex        =   45
      Top             =   2820
      Width           =   84
   End
   Begin VB.Label lblYInt2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   3840
      TabIndex        =   44
      Top             =   6660
      Width           =   84
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<-- Position"
      Height          =   432
      Index           =   11
      Left            =   8580
      TabIndex        =   39
      Top             =   240
      Width           =   972
   End
   Begin VB.Label lblPos 
      BorderStyle     =   1  'Fixed Single
      Height          =   552
      Left            =   5520
      TabIndex        =   38
      Top             =   180
      Width           =   2772
   End
   Begin VB.Label lblMidB 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      Height          =   240
      Left            =   6060
      TabIndex        =   36
      Top             =   7860
      Width           =   156
   End
   Begin VB.Label lblMidA 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      Height          =   240
      Left            =   3720
      TabIndex        =   35
      Top             =   7860
      Width           =   156
   End
   Begin VB.Label Label2 
      Caption         =   "Mid-Point"
      Height          =   432
      Left            =   3900
      TabIndex        =   34
      Top             =   7380
      Width           =   732
   End
   Begin VB.Label Label 
      Caption         =   "point D"
      Height          =   312
      Index           =   10
      Left            =   3780
      TabIndex        =   31
      Top             =   1560
      Width           =   792
   End
   Begin VB.Label Label 
      Caption         =   "Point C"
      Height          =   312
      Index           =   9
      Left            =   3720
      TabIndex        =   28
      Top             =   900
      Width           =   672
   End
   Begin VB.Label Label 
      Caption         =   "LINE2"
      Height          =   432
      Index           =   8
      Left            =   4200
      TabIndex        =   27
      Top             =   240
      Width           =   1512
   End
   Begin VB.Label Label 
      Caption         =   "LINE 1"
      Height          =   492
      Index           =   7
      Left            =   480
      TabIndex        =   26
      Top             =   360
      Width           =   2112
   End
   Begin VB.Label Label 
      Caption         =   "y intercept"
      Height          =   612
      Index           =   6
      Left            =   4680
      TabIndex        =   25
      Top             =   6120
      Width           =   1332
   End
   Begin VB.Label lblEquation2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   3780
      TabIndex        =   24
      Top             =   5400
      Width           =   1104
   End
   Begin VB.Label Equation 
      Caption         =   "Equation"
      Height          =   312
      Left            =   3840
      TabIndex        =   23
      Top             =   4980
      Width           =   1332
   End
   Begin VB.Label lblDistance2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   3960
      TabIndex        =   22
      Top             =   4140
      Width           =   84
   End
   Begin VB.Label Label 
      Caption         =   "DIstance"
      Height          =   312
      Index           =   5
      Left            =   4080
      TabIndex        =   21
      Top             =   3540
      Width           =   912
   End
   Begin VB.Label Label 
      Caption         =   "slope 2"
      Height          =   372
      Index           =   4
      Left            =   4260
      TabIndex        =   20
      Top             =   2340
      Width           =   912
   End
   Begin VB.Label Label 
      Caption         =   "Point B"
      Height          =   252
      Index           =   3
      Left            =   360
      TabIndex        =   17
      Top             =   1680
      Width           =   852
   End
   Begin VB.Line Line 
      X1              =   3540
      X2              =   3540
      Y1              =   780
      Y2              =   8520
   End
   Begin VB.Label lblMidX 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "x"
      Height          =   240
      Left            =   540
      TabIndex        =   14
      Top             =   7800
      Width           =   120
   End
   Begin VB.Label lblMidY 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y"
      Height          =   240
      Left            =   2040
      TabIndex        =   13
      Top             =   7680
      Width           =   156
   End
   Begin VB.Label Label 
      Caption         =   "Mid-Point"
      Height          =   372
      Index           =   2
      Left            =   660
      TabIndex        =   12
      Top             =   7260
      Width           =   972
   End
   Begin VB.Label lblYInt 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   360
      TabIndex        =   11
      Top             =   6540
      Width           =   84
   End
   Begin VB.Label Label 
      Caption         =   "y-intercept"
      Height          =   432
      Index           =   1
      Left            =   480
      TabIndex        =   10
      Top             =   6240
      Width           =   1452
   End
   Begin VB.Label lblEquation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Left            =   360
      TabIndex        =   9
      Top             =   5340
      Width           =   1164
   End
   Begin VB.Label Label 
      Caption         =   "Equation"
      Height          =   312
      Index           =   0
      Left            =   540
      TabIndex        =   8
      Top             =   4920
      Width           =   912
   End
   Begin VB.Label lblDistance 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   540
      TabIndex        =   7
      Top             =   3960
      Width           =   84
   End
   Begin VB.Label Label3 
      Caption         =   "Distance"
      Height          =   552
      Left            =   720
      TabIndex        =   6
      Top             =   3540
      Width           =   1152
   End
   Begin VB.Label lblSlope 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   420
      TabIndex        =   5
      Top             =   2880
      Width           =   84
   End
   Begin VB.Label Labelslope 
      Caption         =   "Slope"
      Height          =   492
      Left            =   1020
      TabIndex        =   4
      Top             =   2460
      Width           =   1512
   End
   Begin VB.Label label1 
      Caption         =   "Point A"
      Height          =   312
      Left            =   180
      TabIndex        =   3
      Top             =   1020
      Width           =   1152
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim slope As Single
Dim slope2 As Single

Dim extendedpoint1 As Single


Dim distance, distance2 As Single
Dim i, p  As Integer

Dim X1, X2, Y1, Y2, B2, B1, A2, A1, X, Y, TRY1, TRY2, W, T, Point1, Point2, U, V As Single

Dim g, o As Single




Private Sub cmdCalculate_Click()
 X1 = Val(txtX1)
 X2 = Val(txtX2)
 Y1 = Val(txtY1)
 Y2 = Val(txtY2)
 B2 = Val(txtY22)
 B1 = Val(txtY12)
 A2 = Val(txtX22)
 A1 = Val(txtX12)
' to calculate slope of lINE 1

slope = (Y2 - Y1) / (X2 - X1)

lblSlope.Caption = slope


'SLOPE OF LINE 2


slope2 = (B2 - B1) / (A2 - A1)


lblslope2.Caption = slope2

'distance 1

distance = Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2)

lblDistance.Caption = distance

'distance2

distance2 = Sqr((A2 - A1) ^ 2 + (B2 - B1) ^ 2)

lblDistance2.Caption = distance2


   
' Y intecept 1

lblYInt.Caption = (Y1 - slope * X1)

' ' Y intecept 2

lblYInt2.Caption = (B1 - slope2 * X2)



'

' equation 1
  
    lblEquation.Caption = "y =" + Str(slope) + ("x") + "+" + Str(lblYInt)
  
' equation 2

      lblEquation2.Caption = "y =" + Str(slope2) + ("x") + "+" + Str(lblYInt2)
  
  
  ' mid point x 1
  
     
   lblMidX.Caption = Rnd((X1 + X2) / 2)
   
 lblMidY.Caption = Rnd((Y1 + Y2) / 2)
   

  ' mid point 2
  

  
  
       
   lblMidA.Caption = (A1 + A2) / 2
   
 lblMidB.Caption = (B1 + B2) / 2
   
  
  
  
  
  'extended line
  
  
  ' create a new point using the same slope
  
  ' so you have the slope. you have one point , you need another point on the same slope at the corner of the graph whichis (-10,10)
  ' so all you have to do is the set up an algorithm that finds a point through y = -10 AND y = 10 - extending both sides
  ' how to do that algebraically?
  
  'y = mx+ b
  
  ' y= slope(x) + yintercept
  ' -10 - y intercept / slope = new value of x
  ' use this new value of x to draw a line from point 1
  ' 10 - y intercept / slope = new value of x for top extension
  
  
  
  
  'TRY1 = New Val of x
  
 'extendedpoint1 = (-10 - Val(lblYInt.Caption)) / slope
 'pic1.Line (extendedpoint1,-10)
 
  
  
  
  
  
  

End Sub

Private Sub cmdClear_Click()
pic1.Cls


lblp1.Caption = ""
lblp2.Caption = ""
lblp3.Caption = ""
lblp4.Caption = ""

lblp1.Visible = False
lblp2.Visible = False
lblp3.Visible = False
lblp4.Visible = False

txtX1.Text = ""
txtX2.Text = ""
txtY1.Text = ""
txtY2.Text = ""

txtX12.Text = ""
txtX22.Text = ""
txtY12.Text = ""
txtY22.Text = ""

lblSlope.Caption = ""
lblslope2.Caption = ""
lblDistance.Caption = ""
lblDistance2.Caption = ""
lblSlope.Caption = ""
lblslope2.Caption = ""
lblYInt.Caption = ""
lblYInt2.Caption = 2



lblMidX.Caption = ""
lblMidA.Caption = ""
lblMidY.Caption = ""
lblMidB.Caption = ""



Form_Activate
End Sub

Private Sub cmdEXTEND_Click(Index As Integer)

W = (10 - Val(lblYInt)) / slope

T = (-10 - Val(lblYInt)) / slope

pic1.Line (W, 10)-(T, -10)



U = (10 - Val(lblYInt2)) / slope2

V = (-10 - Val(lblYInt2)) / slope2

pic1.Line (U, 10)-(V, -10)





End Sub

Private Sub cmdPlot_Click()
'Dim X, Y, X1, X2, A1, A2, B1, B2, A, B, Y1, Y2 As Single
p = p + 1
If p = 1 Then

X = Val(txtX1.Text)
Y = Val(txtY1.Text)
X1 = X
Y1 = Y


pic1.Circle (X1, Y1), 0.25, vbGreen
lblp1.Visible = True
lblp1.Left = X1 + 1
lblp1.Top = Y1 - 1
lblp1.Caption = Format(X1, "fixed") + "," + Format(Y1, "fixed")

'If p = 2 Then
'draw circle
'draw circle  LINE CONNECTING the two circles
'lblp2 show up


ElseIf p = 2 Then
X = Val(txtX2.Text)
Y = Val(txtY2.Text)
X2 = X
Y2 = Y



pic1.Circle (X2, Y2), 0.25, vbGreen
pic1.Line (X1, Y1)-(X2, Y2)
lblp2.Visible = True
lblp2.Left = X2 + 1
lblp2.Top = Y2 - 1
lblp2.Caption = Format(X2, "fixed") + "," + Format(Y2, "fixed")

ElseIf p = 3 Then
X = Val(txtX12.Text)
Y = Val(txtY12.Text)
A1 = X
B1 = Y


pic1.Circle (A1, B1), 0.25, vbGreen
lblp3.Visible = True
lblp3.Left = A1 + 1
lblp3.Top = B1 - 1
lblp3.Caption = Format(A1, "fixed") + "," + Format(B1, "fixed")



ElseIf p = 4 Then
X = Val(txtX22.Text)
Y = Val(txtY22.Text)
A2 = X
B2 = Y


pic1.Circle (A2, B2), 0.25, vbGreen
pic1.Line (A1, B1)-(A2, B2)

lblp4.Visible = True
lblp4.Left = A2 + 1
lblp4.Top = B2 - 1
lblp4.Caption = Format(A2, "fixed") + "," + Format(B2, "fixed")
End If
End Sub

Private Sub cmdQuit_Click()
End

End Sub

Private Sub Form_Activate()
pic1.Scale (-10, 10)-(10, -10)
pic1.Line (-10, 0)-(10, 0), vbRed
pic1.Line (0, -10)-(0, 10), vbBlue

For i = -10 To 10
    pic1.Line (i, 0.5)-(i, -0.5), vbBlue
    pic1.Line (0.5, i)-(-0.5, i), vbRed
    
Next i
    
End Sub

Private Sub Form_Load()
p = 0
End Sub



Private Sub intersect_Click()
g = Val((Val(lblYInt2) - Val(lblYInt)) / (slope - slope2))

o = Val(slope * Val(g) + Val(lblYInt))

lblIntersect.Caption = "(" + Format(g, "fixed") + "," + Format(o, "fixed") + ")"

End Sub

Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
p = p + 1
If p = 1 Then
X1 = X
Y1 = Y

txtX1.Text = X
txtY1.Text = Y


pic1.Circle (X1, Y1), 0.25, vbGreen
lblp1.Visible = True
lblp1.Left = X1 + 1
lblp1.Top = Y1 - 1
lblp1.Caption = Format(X1, "fixed") + "," + Format(Y1, "fixed")

'If p = 2 Then
'draw circle
'draw circle  LINE CONNECTING the two circles
'lblp2 show up


ElseIf p = 2 Then
X2 = X
Y2 = Y

txtX2.Text = X
txtY2.Text = Y

pic1.Circle (X2, Y2), 0.25, vbGreen
pic1.Line (X2, Y2)-(X1, Y1)
lblp2.Visible = True
lblp2.Left = X2 + 1
lblp2.Top = Y2 - 1
lblp2.Caption = Format(X2, "fixed") + "," + Format(Y2, "fixed")

ElseIf p = 3 Then
A1 = X
B1 = Y

txtX12.Text = X
txtY12.Text = Y
pic1.Circle (A1, B1), 0.25, vbGreen
lblp3.Visible = True
lblp3.Left = A1 + 1
lblp3.Top = B1 - 1
lblp3.Caption = Format(A1, "fixed") + "," + Format(B1, "fixed")



ElseIf p = 4 Then
A2 = X
B2 = Y

txtX22.Text = X
txtY22.Text = Y
pic1.Circle (A2, B2), 0.25, vbGreen
pic1.Line (A2, B2)-(A1, B1)

lblp4.Visible = True
lblp4.Left = A2 + 1
lblp4.Top = B2 - 1
lblp4.Caption = Format(A2, "fixed") + "," + Format(B2, "fixed")

End If



'to extend line

'ElseIf p = 5 Then
'TRY1 = X
'TRY2 = Y

'TRY1 = txtTRY1
'TRY2 = txtTRY2
'pic1.Circle (TRY1, TRY2), 0.25, vbGreen
'pic1.Line (A2, B2)-(A1, B1)
'lblp4.Visible = True
'lblp4.Left = A2 + 1
'lblp4.Top = B2 - 1
'lblp4.Caption = Format(A2, "fixed") + "," + Format(B2, "fixed")







End Sub

Private Sub pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPos.Caption = Format(X, "fixed") + "," + Format(Y, "fixed")




End Sub

