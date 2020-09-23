VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   Caption         =   "Shaper - unlimited graphical deformator ver 1.01"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form3"
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Stars extra"
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   5280
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   255
      Left            =   2760
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "(*.bmp;*.jpg)|*.bmp;*.jpg"
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H000000FF&
      Caption         =   "Hmm..? what to do ?"
      Height          =   495
      Left            =   1440
      MaskColor       =   &H000000FF&
      TabIndex        =   21
      Top             =   2880
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Simetrical multi side shape"
      Height          =   1095
      Left            =   2880
      TabIndex        =   16
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      Caption         =   "loadpic"
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Force FllodFill work"
      Height          =   615
      Left            =   5880
      TabIndex        =   14
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Stars"
      Height          =   1095
      Left            =   1560
      TabIndex        =   13
      Top             =   4200
      Width           =   1335
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2055
      Left            =   6240
      Max             =   10000
      TabIndex        =   12
      Top             =   240
      Value           =   8000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2760
      Max             =   10000
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Cls pic 2"
      Height          =   615
      Left            =   5880
      TabIndex        =   10
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Vertical draw"
      Height          =   615
      Left            =   7320
      TabIndex        =   9
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Shape3"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Shape1"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "egg and it's derivates"
      Height          =   1095
      Left            =   4080
      TabIndex        =   6
      Top             =   4200
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   3840
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   170
      TabIndex        =   5
      Top             =   360
      Width           =   2550
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Shape2"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Horizontal drav"
      Height          =   615
      Left            =   7320
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   0
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   215
      TabIndex        =   0
      Top             =   240
      Width           =   3225
      Begin VB.Image Image1 
         Height          =   2535
         Left            =   0
         Picture         =   "Shaper.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3345
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Fixed shapes"
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Predefined modifyable mathematical created shapes"
      Height          =   1095
      Left            =   0
      TabIndex        =   19
      Top             =   3480
      Width           =   1335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mbrush As Long
Dim ActiveShape As String
Dim bigdiameter
Dim smalldiameter

Private Sub Command1_Click()
'an array to hold data about found sections in pic2
'I assume that tere can't be more than 20 sections in one row of
'pixels, enlarge this if u like
'2 holds start point of section,it's lenght, later percentage of sections lenghts
Dim DataHolder() As Variant
ReDim DataHolder(Picture2.Height, 3, 20)
DataHolderIndex3 = 0



'get first upper left corner pixel color,
'we will assume there is no shape
'and that this color represents empty space

For j = 0 To (Picture2.Height - 1) 'line by line to the bottom of the pic
    Do
    i = i + 1
        'main if, if black colour found
        If GetPixel(Picture2.hdc, i, j) = 0 Then ' black
            SectionStartPointX = i
            SectionLenght = 1
            Do
            If GetPixel(Picture2.hdc, i + 1, j) <> 0 Then
                'we have now the data for our array, fill it !
                
                'Stop
                
                DataHolder(j, 1, DataHolderIndex3) = SectionStartPointX
                DataHolder(j, 2, DataHolderIndex3) = SectionLenght
                DataHolderIndex3 = DataHolderIndex3 + 1
                'DoEvents
                Debug.Print j & "," & "SectionStartPointX," & SectionStartPointX
                Debug.Print "SectionLenght" & SectionLenght
                
                
                Exit Do
            End If
            SectionLenght = SectionLenght + 1
            i = i + 1
            Loop
        SectionLenght = 0
        
        End If
    
    Loop Until i = Picture2.Width - 1
    i = 0
    DataHolderIndex3 = 0
Next j

Picture2.Cls

'END OF EATING DATA INTO ARRAY
'erase all shapes if any there


'Nov get data from array and produce final picture

For i = 0 To UBound(DataHolder, 1) ' to pic 2 height

'sumarize lenghts of sections in one row
 For j = 0 To UBound(DataHolder, 3)
 If DataHolder(i, 2, j) = Empty Then Exit For
 SumLenght = SumLenght + DataHolder(i, 2, j)
 'DataHolder(Picture2.Height, 3, 20)
 Next j
 
 'calculate percentages
 For j = 0 To UBound(DataHolder, 3)
 If DataHolder(i, 2, j) = Empty Then Exit For
 dd = DataHolder(i, 2, j)
 DataHolder(i, 3, j) = (dd * 100) / SumLenght
 
 Next j
 
SumLenght = 0
Next i
'EUREKA! NOW WE HAVE ALL THE DATA NEEDED IN OUR ARRAY
'TRANSFER THE PIC1 STUFF INTO PIC2 AS WE NEED!!!

Picture2.AutoRedraw = True '!!!!!!!!!!!!!!!!!!!
DoEvents

For i = 0 To UBound(DataHolder, 1) 'pic2 height
    For j = 0 To UBound(DataHolder, 3) ' up to 20
    If DataHolder(i, 2, j) = Empty Then Exit For
    'DataHolder(Picture2.Height, 3, 20)
    Pic1SectionWidth = (Picture1.Width) * DataHolder(i, 3, j) / 100
    
    'na kateri toèki naj zaène brati source
    If j = 0 Then
    StartPointInsideSource = 0
    Else
        Do 'seštej vse procente do sedaj, minus trenutnega pa dobiš toèko
        ProcentSeštevek = ProcentSeštevek + DataHolder(i, 3, z)
        z = z + 1
        StartPointInsideSource = (Picture1.Width * ProcentSeštevek) / 100
        Loop Until z = j
        ProcentSeštevek = 0
        z = 0
    
    End If
    
    
    StretchBlt Picture2.hdc, DataHolder(i, 1, j), i, DataHolder(i, 2, j), 1, Picture1.hdc, StartPointInsideSource, i, Pic1SectionWidth, 1, SRCCOPY
    DoEvents
    Next j

Next i


Picture2.Refresh
End Sub









Private Sub Command10_Click()
ActiveShape = "Starrs" 'used in hscroll
numberofpeaks = 6

Picture2.Width = Picture1.Width
Picture2.Height = Picture1.Height

HScroll1.Width = Picture2.Width
HScroll1.Top = Picture2.Top + Picture2.Height - 1
HScroll1.Left = Picture2.Left
HScroll1.Min = 1
HScroll1.Max = 50
HScroll1.Value = numberofpeaks
HScroll1.Visible = True

HScroll2.Width = Picture2.Width
HScroll2.Top = Picture2.Top + Picture2.Height + HScroll1.Height
HScroll2.Left = Picture2.Left
HScroll2.Min = 0
HScroll2.Max = 100
HScroll2.Value = 0
HScroll2.Visible = True

HScroll3.Width = Picture2.Width
HScroll3.Top = Picture2.Top + Picture2.Height + HScroll1.Height + HScroll2.Height
HScroll3.Left = Picture2.Left
HScroll3.Min = 0
HScroll3.Max = Picture2.Height * 1.5
HScroll3.Visible = True

HScroll4.Width = Picture2.Width
HScroll4.Top = Picture2.Top + Picture2.Height + HScroll1.Height + HScroll2.Height + HScroll3.Height
HScroll4.Left = Picture2.Left
HScroll4.Min = 0
HScroll4.Max = Picture2.Height * 1.5
HScroll4.Visible = True

DoEvents
If Picture2.Width < Picture2.Height Then
bigdiameter = (Picture2.Width) / 2
smalldiameter = ((Picture2.Width) / 2) / 2
'center point of our starr
SW = Picture2.Width / 2
SH = Picture2.Height / 2
Else
bigdiameter = (Picture2.Height) / 2
smalldiameter = ((Picture2.Height) / 2) / 2
'center point of our starr
SW = Picture2.Width / 2
SH = Picture2.Height / 2
End If

HScroll3.Value = bigdiameter
HScroll3.Value = smalldiameter



d = 1
c = 0
numberofpeaks = 6
'bigdiameter = (Picture2.Width) / 2
'smalldiameter = ((Picture2.Width) / 2) / 2 / 2
CircularOffSet1 = 0
CircularOffSet2 = 0
Dim myarray() As Integer
ReDim myarray(numberofpeaks, 2, 2)

Picture2.AutoRedraw = False



'Do
For angle = 0 To 6.3 Step 0.01
      
      'clockwise
      a = Sin(6.3 - angle) * bigdiameter
      b = Cos(6.3 - angle) * bigdiameter
      a1 = Sin(6.3 - angle) * smalldiameter
      b1 = Cos(6.3 - angle) * smalldiameter


Select Case Format(angle, "0.00")

Case Format((6.3 / numberofpeaks * d) - CircularOffSet1, "0.00") '1.57, 3....
    SetPixel Picture2.hdc, a + SW, b + SH, 0
    
    myarray(d, 1, 1) = a + SW
    myarray(d, 1, 2) = b + SH
    d = d + 1
    
Case Format((6.3 / numberofpeaks / 2) + ((6.3 / numberofpeaks) * c) - CircularOffSet2, "0.00") '0.79, 2.37, 3.94, 5.52
    
'Case 0.79, 2.37, 3.94, 5.52
    'If Format(angle, "0.00") = 2.37 Then Stop
g = Format((6.3 / numberofpeaks / 2) + ((6.3 / numberofpeaks) * c), "0.00")
     c = c + 1
    SetPixel Picture2.hdc, a1 + SW, b1 + SH, 0
    myarray(c, 2, 1) = a1 + SW
    myarray(c, 2, 2) = b1 + SH
   
Case Else
End Select
Next angle

'Exit Sub
For h = 1 To UBound(myarray, 1)
    'For z = 0 To 5000000
    'Next z
    Picture2.Line (myarray(h, 2, 1), myarray(h, 2, 2))-(myarray(h, 1, 1), myarray(h, 1, 2)), 0
        If h < UBound(myarray, 1) Then
        Picture2.Line (myarray(h, 1, 1), myarray(h, 1, 2))-(myarray(h + 1, 2, 1), myarray(h + 1, 2, 2)), 0
        Else
        Picture2.Line (myarray(UBound(myarray, 1), 1, 1), myarray(UBound(myarray, 1), 1, 2))-(myarray(1, 2, 1), myarray(1, 2, 2)), 0
        End If
Next h


End Sub

Private Sub Command11_Click()
Picture2.AutoRedraw = False
End Sub

Private Sub Command12_Click()
CommonDialog1.ShowOpen
Dim ReturnedFile As String
ReturnedFile = CommonDialog1.FileName
Image1.Picture = LoadPicture(ReturnedFile)
Beep
End Sub

Private Sub Command13_Click()
ActiveShape = "SimetricalMultiSideShape" 'used in hscroll
numberofpeaks = 4
CircularOffSet = 0

Picture2.Width = Picture1.Width
Picture2.Height = Picture1.Height
Picture2.Top = Picture1.Top


HScroll1.Width = Picture2.Width
HScroll1.Top = Picture2.Top + Picture2.Height - 1
HScroll1.Left = Picture2.Left
HScroll1.Min = 1
HScroll1.Max = 50
HScroll1.Value = numberofpeaks
HScroll1.Visible = True

HScroll2.Width = Picture2.Width
HScroll2.Top = Picture2.Top + Picture2.Height + HScroll1.Height
HScroll2.Left = Picture2.Left
HScroll2.Min = 0
HScroll2.Max = 100
HScroll2.Value = 0
HScroll2.Visible = True


DoEvents
If Picture2.Width < Picture2.Height Then
bigdiameter = (Picture2.Width) / 2
'center point of our starr
SW = Picture2.Width / 2
SH = Picture2.Height / 2
Else
bigdiameter = (Picture2.Height) / 2
'center point of our starr
SW = Picture2.Width / 2
SH = Picture2.Height / 2
End If

d = 1



Dim myarray() As Integer
ReDim myarray(numberofpeaks, 2)

Picture2.AutoRedraw = False



'Do
For angle = 0 To 6.3 Step 0.01
      'clockwise
      a = Sin(6.3 - angle) * bigdiameter
      b = Cos(6.3 - angle) * bigdiameter
    
Select Case Format(angle, "0.00")

Case Format((6.3 / numberofpeaks * d) - CircularOffSet, "0.00") '1.57, 3....
    SetPixel Picture2.hdc, a + SW, b + SH, 0
    
    myarray(d, 1) = a + SW
    myarray(d, 2) = b + SH
    d = d + 1
Case Else
End Select

Next angle

'Exit Sub
For h = 1 To UBound(myarray, 1)
    'For z = 0 To 5000000
    'Next z
    If h < UBound(myarray, 1) Then
    Picture2.Line (myarray(h, 1), myarray(h, 2))-(myarray(h + 1, 1), myarray(h + 1, 2)), 0
    Else
    Picture2.Line (myarray(h, 1), myarray(h, 2))-(myarray(1, 1), myarray(1, 2)), 0
    End If
Next h

End Sub

Private Sub Command14_Click()
msg = "1.Create 'fillable'shape in picture2, either in design environment or by predefined butons at run time" & Chr(10) _
& "2.Fill the shape with black color = clicking inside shape will do." & Chr(10) _
& "3. Click Vertical or Horizontal draw button" & Chr(10) _
& Chr(10) _
& "That is all. You vill allso notice: it is recomended that in case of using vertical drav, shape ocupies whole width of picture and in case of horizontal drav it is the best that shape ocupies whole height of picture. Allso note that yo can floodFill space outside your shape and Floodfill it. You can allso combine many source pictures into one destination picture."

MsgBox msg
End Sub





Private Sub Command2_Click()

numberofpeaks = 10

Picture2.Width = Picture1.Width
Picture2.Height = Picture1.Height




DoEvents
If Picture2.Width < Picture2.Height Then
bigdiameter = (Picture2.Width) / 2
smalldiameter = ((Picture2.Width) / 2) / 2
'center point of our starr
SW = Picture2.Width / 2
SH = Picture2.Height / 2
Else
bigdiameter = (Picture2.Height) / 2
smalldiameter = ((Picture2.Height) / 2) / 2
'center point of our starr
SW = Picture2.Width / 2
SH = Picture2.Height / 2
End If


d = 1
c = 0

'bigdiameter = (Picture2.Width) / 2
'smalldiameter = ((Picture2.Width) / 2) / 2 / 2
CircularOffSet1 = 0
CircularOffSet2 = 0
Dim myarray() As Integer
ReDim myarray(numberofpeaks, 2, 2)

Picture2.AutoRedraw = False

'Do
For angle = 0 To 6.3 Step 0.01
      
      'clockwise
      a = Sin(6.3 - angle) * bigdiameter
      b = Cos(6.3 - angle) * bigdiameter
      a1 = Sin(6.3 - angle) * smalldiameter
      b1 = Cos(6.3 - angle) * smalldiameter


Select Case Format(angle, "0.00")

Case Format((6.3 / numberofpeaks * d) - CircularOffSet1, "0.00") '1.57, 3....
    SetPixel Picture2.hdc, a + SW, b + SH, 0
    
    myarray(d, 1, 1) = a + SW
    myarray(d, 1, 2) = b + SH
    d = d + 1
    
Case Format((6.3 / numberofpeaks / 2) + ((6.3 / numberofpeaks) * c) - CircularOffSet2, "0.00") '0.79, 2.37, 3.94, 5.52
    
'Case 0.79, 2.37, 3.94, 5.52
    'If Format(angle, "0.00") = 2.37 Then Stop
g = Format((6.3 / numberofpeaks / 2) + ((6.3 / numberofpeaks) * c), "0.00")
     c = c + 1
    SetPixel Picture2.hdc, a1 + SW, b1 + SH, 0
    myarray(c, 2, 1) = a1 + SW
    myarray(c, 2, 2) = b1 + SH
   
Case Else
End Select
Next angle

'Exit Sub
For h = 1 To UBound(myarray, 1)
    'For z = 0 To 5000000
    'Next z
    Picture2.Line (myarray(h, 2, 1), myarray(h, 2, 2))-(myarray(h, 1, 1), myarray(h, 1, 2)), 0
        If h < UBound(myarray, 1) Then
        Picture2.Line (myarray(h, 1, 1), myarray(h, 1, 2))-(myarray(h + 1, 2, 1), myarray(h + 1, 2, 2)), 0
        Picture2.Line (SW, SH)-(myarray(h + 1, 2, 1), myarray(h + 1, 2, 2)), 0
        
        Else
        'last line to draw
        Picture2.Line (myarray(UBound(myarray, 1), 1, 1), myarray(UBound(myarray, 1), 1, 2))-(myarray(1, 2, 1), myarray(1, 2, 2)), 0
        Picture2.Line (SW, SH)-(myarray(1, 2, 1), myarray(1, 2, 2)), 0
        End If
Next h
End Sub

Private Sub Command4_Click()
Picture2.AutoRedraw = True
diameter = Picture2.Width / 1.9
MyColor = 0 'Do
'nariše krog - make egg

For kot = 0 To 20 Step 0.001
a = Sin(kot) * diameter * Cos(Sin(kot / 2))
b = Cos(kot) * diameter * 1.5 * Sin(Sin(kot)) 'change 3!!!!! - cool
SetPixel Picture2.hdc, a + (Picture2.Width / 2), b + (Picture2.Height / 2), 0
'PSet (a + 100, b + 100), RGB(0, 0, MyColor)
Next kot
Picture2.Refresh
Picture2.AutoRedraw = False
End Sub

Private Sub Command5_Click()
Picture2.AutoRedraw = True
diameter = Picture2.Width / 1.3
MyColor = 0



For kot = 0 To 10 Step 0.001
a = Sin(kot / 3) * Cos(kot) * diameter / 2 'change 2!!!!! - cool
b = Sin(kot / 3) * Sin(kot) * diameter / 2 'change 6!!!!! - cool
'veè sinusov bolj je ukrivljeno

SetPixel Picture2.hdc, a + (Picture2.Width / 2), b + (Picture2.Width / 2), 0
'PSet (a + 100, b + 100), RGB(0, 0, MyColor)
Next kot

Picture2.Refresh
Picture2.AutoRedraw = False
End Sub

Private Sub Command6_Click()
ActiveShape = "Egg" 'used in hscroll
'Picture2.AutoRedraw = True
Picture2.Cls

Picture2.Width = Picture1.Width
Picture2.Height = Picture1.Height

HScroll1.Width = Picture2.Width - 1
HScroll1.Top = Picture2.Top + Picture2.Height
HScroll1.Left = Picture2.Left


VScroll1.Height = Picture2.Height - 1
VScroll1.Left = Picture2.Left + Picture2.Width

HScroll1.Max = (Picture2.Width / Picture2.Height) * 6330
'VScroll1.Max = (Picture2.Height / Picture2.Width) * 6340



HScroll1.Value = HScroll1.Max / 2
VScroll1.Value = VScroll1.Max / 2

VScroll1.Visible = True
HScroll1.Visible = True

'decision about diameter
'a = Picture2.Height / Picture2.Width
'Select Case Picture2.Height / Picture2.Width
'Case Is > 1.247706422018 ' to narrow
'Case Is = 1.247706422018 'perfect
'diameter = (Picture2.Height - 1) / 2
'Case Is < 1.247706422018 ' to vide for normal egg
'diameter = (Picture2.Height - 1) / 2
'Case Else
'End Select
diameter = (Picture2.Height - 1) / 2

'nariše krog - make egg

For kot = 0 To 20 Step 0.01

a = Sin(kot) * diameter * Cos(Sin(kot / 2))
b = Cos(kot) * diameter '* 0.5 '* 1.5 magnifay height
SetPixel Picture2.hdc, a + (Picture2.Width / 2), b + (Picture2.Height / 2), 0
Next kot
'Picture2.Refresh
'Picture2.AutoRedraw = False
End Sub

Private Sub Command7_Click()
Picture2.AutoRedraw = True
diameter = Picture2.Width / 2

'Do
'nariše krog - make egg

For kot = 0 To 10 Step 0.0001
a = Sin(kot) * diameter * Sin(kot * 16)  'change 6!!!!! - cool
'veè sinusov bolj je ukrivljeno
b = Cos(kot) * diameter * Sin(Sin(Sin(Sin(Sin(Sin(Sin(Sin(Sin((Sin(Sin(Sin(kot * 16) ^ 1)))))))))))) 'change 2!!!!! - cool
SetPixel Picture2.hdc, a + (Picture2.Width / 2), b + (Picture2.Height / 2), 0
'PSet (a + 100, b + 100), RGB(0, 0, MyColor)
Next kot
Picture2.Refresh
Picture2.AutoRedraw = False
End Sub

Private Sub Command8_Click() 'vertical

Dim DataHolder() As Variant
ReDim DataHolder(Picture2.Width, 3, 20)
DataHolderIndex3 = 0



'get first upper left corner pixel color,
'we will assume there is no shape
'and that this color represents empty space

For j = 0 To (Picture2.Width - 1)  'line by line to the right side of the pic
    Do
    i = i + 1 ' y thing
        'main if, if black colour found
        If GetPixel(Picture2.hdc, j, i) = 0 Then ' black
            SectionStartPointY = i
            SectionLenght = 1
            Do
            If GetPixel(Picture2.hdc, j, i + 1) <> 0 Then
                'we have now the data for our array, fill it !
                
                
                
                DataHolder(j, 1, DataHolderIndex3) = SectionStartPointY
                DataHolder(j, 2, DataHolderIndex3) = SectionLenght
                DataHolderIndex3 = DataHolderIndex3 + 1
                'DoEvents
                'j = x coordinate
                'Debug.Print j & "," & "SectionStartPointy," & SectionStartPointY
                'Debug.Print "SectionLenght" & SectionLenght
                
               
                Exit Do
            End If
            SectionLenght = SectionLenght + 1
            i = i + 1
            Loop
        SectionLenght = 0
        
        End If
    
    Loop Until i = Picture2.Height - 1
    i = 0
    DataHolderIndex3 = 0
Next j

Picture2.Cls

'END OF EATING DATA INTO ARRAY
'erase all shapes if any there


'Nov get data from array and produce final picture

For i = 0 To UBound(DataHolder, 1) ' to pic 2 width

    'sumarize lenghts of sections in one column
     For j = 0 To UBound(DataHolder, 3)
     If DataHolder(i, 2, j) = Empty Then Exit For
     SumLenght = SumLenght + DataHolder(i, 2, j)
     'DataHolder(Picture2.width, 3, 20)
     Next j
     
     'calculate percentages
     For j = 0 To UBound(DataHolder, 3)
     If DataHolder(i, 2, j) = Empty Then Exit For
     dd = DataHolder(i, 2, j)
     DataHolder(i, 3, j) = (dd * 100) / SumLenght
     a = DataHolder(i, 3, j)
     Next j
     
    SumLenght = 0
Next i
'EUREKA! NOW WE HAVE ALL THE DATA NEEDED IN OUR ARRAY
'TRANSFER THE PIC1 STUFF INTO PIC2 AS WE NEED!!!

Picture2.AutoRedraw = True '!!!!!!!!!!!!!!!!!!!
DoEvents

For i = 0 To UBound(DataHolder, 1) 'pic2 width
    For j = 0 To UBound(DataHolder, 3) ' up to 20
    If DataHolder(i, 2, j) = Empty Then Exit For
    'DataHolder(Picture2.Height, 3, 20)
    Pic1SectionHeight = (Picture1.Height) * DataHolder(i, 3, j) / 100
    
    'na kateri toèki naj zaène brati source
    If j = 0 Then
    StartPointInsideSource = 0
    Else
        Do 'seštej vse procente do sedaj, minus trenutnega pa dobiš toèko
        ProcentSeštevek = ProcentSeštevek + DataHolder(i, 3, z)
        z = z + 1
        StartPointInsideSource = (Picture1.Height * ProcentSeštevek) / 100
        Loop Until z = j
        ProcentSeštevek = 0
        z = 0
    
    End If
    
    
    StretchBlt Picture2.hdc, i, DataHolder(i, 1, j), 1, DataHolder(i, 2, j), Picture1.hdc, i, StartPointInsideSource, 1, Pic1SectionHeight, SRCCOPY
    Debug.Print i & ","; StartPointInsideSource
    
    DoEvents
    Next j

Next i


Picture2.Refresh
End Sub

Private Sub Command9_Click()
Picture2.AutoRedraw = True
Picture2.Cls
End Sub

Private Sub Form_Load()
Picture2.Width = Picture1.Width
Picture2.Height = Picture1.Height
Picture2.Top = Picture1.Top
Picture2.Left = Picture1.Left + Picture1.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteObject mbrush
End Sub

Private Sub HScroll1_Scroll()
Select Case ActiveShape
Case Is = "SimetricalMultiSideShape"
Picture2.Cls
    If Picture2.Width < Picture2.Height Then
    bigdiameter = (Picture2.Width) / 2
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    Else
    bigdiameter = (Picture2.Height) / 2
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    End If
    
    d = 1

    Dim myarray() As Integer
    ReDim myarray(HScroll1.Value, 2)
    
    Picture2.AutoRedraw = False
    
    
    
    'Do
    For angle = 0 To 6.3 Step 0.01
          'clockwise
          a = Sin(6.3 - angle) * bigdiameter
          b = Cos(6.3 - angle) * bigdiameter
        
    Select Case Format(angle, "0.00")
    
    Case Format((6.3 / HScroll1.Value * d) - CircularOffSet, "0.00") '1.57, 3....
        SetPixel Picture2.hdc, a + SW, b + SH, 0
        
        myarray(d, 1) = a + SW
        myarray(d, 2) = b + SH
        d = d + 1
    Case Else
    End Select
    
    Next angle
    
    'Exit Sub
    For h = 1 To UBound(myarray, 1)
        'For z = 0 To 5000000
        'Next z
        If h < UBound(myarray, 1) Then
        Picture2.Line (myarray(h, 1), myarray(h, 2))-(myarray(h + 1, 1), myarray(h + 1, 2)), 0
        Else
        Picture2.Line (myarray(h, 1), myarray(h, 2))-(myarray(1, 1), myarray(1, 2)), 0
        End If
    Next h

Case "Starrs"
Picture2.Cls
    If Picture2.Width < Picture2.Height Then
    bigdiameter = (Picture2.Width) / 2
    smalldiameter = ((Picture2.Width) / 2) / 2
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    Else
    bigdiameter = (Picture2.Height) / 2
    smalldiameter = ((Picture2.Height) / 2) / 2
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    End If
    
    d = 1
    c = 0
    numberofpeaks = HScroll1.Value
    'bigdiameter = (Picture2.Width) / 2
    'smalldiameter = ((Picture2.Width) / 2) / 2 / 2
    CircularOffSet1 = HScroll2.Value / 100
    CircularOffSet2 = 0
    Dim myarray1() As Integer
    ReDim myarray1(numberofpeaks, 2, 2)
    
    Picture2.AutoRedraw = False
    
    
    
    'Do
    For angle = 0 To 6.3 Step 0.01
          
          'clockwise
          a = Sin(6.3 - angle) * bigdiameter
          b = Cos(6.3 - angle) * bigdiameter
          a1 = Sin(6.3 - angle) * smalldiameter
          b1 = Cos(6.3 - angle) * smalldiameter
    
    
    Select Case Format(angle, "0.00")
    
    Case Format((6.3 / numberofpeaks * d) - CircularOffSet1, "0.00") '1.57, 3....
        SetPixel Picture2.hdc, a + SW, b + SH, 0
        
        myarray1(d, 1, 1) = a + SW
        myarray1(d, 1, 2) = b + SH
        d = d + 1
        
    Case Format((6.3 / numberofpeaks / 2) + ((6.3 / numberofpeaks) * c) - CircularOffSet2, "0.00") '0.79, 2.37, 3.94, 5.52
        
    'Case 0.79, 2.37, 3.94, 5.52
        'If Format(angle, "0.00") = 2.37 Then Stop
    g = Format((6.3 / numberofpeaks / 2) + ((6.3 / numberofpeaks) * c), "0.00")
         c = c + 1
        SetPixel Picture2.hdc, a1 + SW, b1 + SH, 0
        myarray1(c, 2, 1) = a1 + SW
        myarray1(c, 2, 2) = b1 + SH
       
    Case Else
    End Select
    Next angle
    
    'Exit Sub
    For h = 1 To UBound(myarray1, 1)
        'For z = 0 To 5000000
        'Next z
        Picture2.Line (myarray1(h, 2, 1), myarray1(h, 2, 2))-(myarray1(h, 1, 1), myarray1(h, 1, 2)), 0
            If h < UBound(myarray1, 1) Then
            Picture2.Line (myarray1(h, 1, 1), myarray1(h, 1, 2))-(myarray1(h + 1, 2, 1), myarray1(h + 1, 2, 2)), 0
            Else
            Picture2.Line (myarray1(UBound(myarray1, 1), 1, 1), myarray1(UBound(myarray1, 1), 1, 2))-(myarray1(1, 2, 1), myarray1(1, 2, 2)), 0
            End If
    Next h
    
Case "Egg"
Picture2.Cls
DoEvents
HScroll1.Max = 100
Text1.Text = HScroll1.Value
'Picture2.AutoRedraw = True
Picture2.Cls

diameter = (Picture2.Height - 1) / 2

'nariše krog - make egg

For kot = 0 To 20 Step 0.01

a = Sin(kot) * diameter * Cos(Sin(kot / 2)) * (HScroll1.Value) / 50
b = Cos(kot) * diameter '* 0.5 '* 1.5 magnifay height
SetPixel Picture2.hdc, a + (Picture2.Width / 2), b + (Picture2.Height / 2), 0
Next kot
'Picture2.Refresh
'    Picture2.AutoRedraw = False
Case Else
End Select

    


End Sub

Private Sub HScroll2_Scroll()
Select Case ActiveShape
Case Is = "SimetricalMultiSideShape"
HScroll2.Max = 6.3 / HScroll1.Value * 49

Text1.Text = HScroll2.Value
Picture2.Cls

    If Picture2.Width < Picture2.Height Then
    bigdiameter = (Picture2.Width) / 2
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    Else
    bigdiameter = (Picture2.Height) / 2
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    End If
    
    d = 1

    Dim myarray() As Integer
    ReDim myarray(HScroll1.Value, 2)
    
    Picture2.AutoRedraw = False
    
    
    
    'Do
    For angle = 0 To 6.3 Step 0.01
          'clockwise
          a = Sin(6.3 - angle) * bigdiameter
          b = Cos(6.3 - angle) * bigdiameter
        
    Select Case Format(angle, "0.00")
    
    Case Format((6.3 / HScroll1.Value * d) - (HScroll2.Value / 50), "0.00") '1.57, 3....
        SetPixel Picture2.hdc, a + SW, b + SH, 0
        
        myarray(d, 1) = a + SW
        myarray(d, 2) = b + SH
        d = d + 1
    Case Else
    End Select
    
    Next angle
    
    'Exit Sub
    For h = 1 To UBound(myarray, 1)
        'For z = 0 To 5000000
        'Next z
        If h < UBound(myarray, 1) Then
        Picture2.Line (myarray(h, 1), myarray(h, 2))-(myarray(h + 1, 1), myarray(h + 1, 2)), 0
        Else
        Picture2.Line (myarray(h, 1), myarray(h, 2))-(myarray(1, 1), myarray(1, 2)), 0
        End If
    Next h

Case Is = "Starrs"
Text1.Text = HScroll2.Value
Picture2.Cls
numberofpeaks = HScroll1.Value
HScroll2.Max = 6.3 / HScroll1.Value * 45
    If Picture2.Width < Picture2.Height Then
    bigdiameter = (Picture2.Width) / 2
    smalldiameter = ((Picture2.Width) / 2) / 2
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    Else
    bigdiameter = (Picture2.Height) / 2
    smalldiameter = ((Picture2.Height) / 2) / 2
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    End If
    
    d = 1
    c = 0
   
    'bigdiameter = (Picture2.Width) / 2
    'smalldiameter = ((Picture2.Width) / 2) / 2 / 2
    CircularOffSet1 = HScroll2.Value / 100
    CircularOffSet2 = 0
    Dim myarray1() As Integer
    ReDim myarray1(numberofpeaks, 2, 2)
    
    Picture2.AutoRedraw = False
    
    
    
    'Do
    For angle = 0 To 6.3 Step 0.01
          
          'clockwise
          a = Sin(6.3 - angle) * bigdiameter
          b = Cos(6.3 - angle) * bigdiameter
          a1 = Sin(6.3 - angle) * smalldiameter
          b1 = Cos(6.3 - angle) * smalldiameter
    
    
    Select Case Format(angle, "0.00")
    
    Case Format((6.3 / numberofpeaks * d) - CircularOffSet1, "0.00") '1.57, 3....
        SetPixel Picture2.hdc, a + SW, b + SH, 0
        
        myarray1(d, 1, 1) = a + SW
        myarray1(d, 1, 2) = b + SH
        d = d + 1
        
    Case Format((6.3 / numberofpeaks / 2) + ((6.3 / numberofpeaks) * c) - CircularOffSet2, "0.00") '0.79, 2.37, 3.94, 5.52
        
    'Case 0.79, 2.37, 3.94, 5.52
        'If Format(angle, "0.00") = 2.37 Then Stop
    g = Format((6.3 / numberofpeaks / 2) + ((6.3 / numberofpeaks) * c), "0.00")
         c = c + 1
        SetPixel Picture2.hdc, a1 + SW, b1 + SH, 0
        myarray1(c, 2, 1) = a1 + SW
        myarray1(c, 2, 2) = b1 + SH
       
    Case Else
    End Select
    Next angle
    
    'Exit Sub
    For h = 1 To UBound(myarray1, 1)
        'For z = 0 To 5000000
        'Next z
        Picture2.Line (myarray1(h, 2, 1), myarray1(h, 2, 2))-(myarray1(h, 1, 1), myarray1(h, 1, 2)), 0
            If h < UBound(myarray1, 1) Then
            Picture2.Line (myarray1(h, 1, 1), myarray1(h, 1, 2))-(myarray1(h + 1, 2, 1), myarray1(h + 1, 2, 2)), 0
            Else
            Picture2.Line (myarray1(UBound(myarray1, 1), 1, 1), myarray1(UBound(myarray1, 1), 1, 2))-(myarray1(1, 2, 1), myarray1(1, 2, 2)), 0
            End If
    Next h

Case Else
End Select
    

End Sub

Private Sub HScroll3_Scroll()
Select Case ActiveShape
Case Is = "SimetricalMultiSideShape"
HScroll2.Max = 6.3 / HScroll1.Value * 49

Text1.Text = HScroll2.Value
Picture2.Cls

    If Picture2.Width < Picture2.Height Then
    bigdiameter = (Picture2.Width) / 2
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    Else
    bigdiameter = (Picture2.Height) / 2
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    End If
    
    d = 1

    Dim myarray() As Integer
    ReDim myarray(HScroll1.Value, 2)
    
    Picture2.AutoRedraw = False
    
    
    
    'Do
    For angle = 0 To 6.3 Step 0.01
          'clockwise
          a = Sin(6.3 - angle) * bigdiameter
          b = Cos(6.3 - angle) * bigdiameter
        
    Select Case Format(angle, "0.00")
    
    Case Format((6.3 / HScroll1.Value * d) - (HScroll2.Value / 50), "0.00") '1.57, 3....
        SetPixel Picture2.hdc, a + SW, b + SH, 0
        
        myarray(d, 1) = a + SW
        myarray(d, 2) = b + SH
        d = d + 1
    Case Else
    End Select
    
    Next angle
    
    'Exit Sub
    For h = 1 To UBound(myarray, 1)
        'For z = 0 To 5000000
        'Next z
        If h < UBound(myarray, 1) Then
        Picture2.Line (myarray(h, 1), myarray(h, 2))-(myarray(h + 1, 1), myarray(h + 1, 2)), 0
        Else
        Picture2.Line (myarray(h, 1), myarray(h, 2))-(myarray(1, 1), myarray(1, 2)), 0
        End If
    Next h

Case Is = "Starrs"
'Text1.Text = HScroll2.Value
Picture2.Cls
numberofpeaks = HScroll1.Value
HScroll2.Max = 6.3 / HScroll1.Value * 45
    If Picture2.Width < Picture2.Height Then
    bigdiameter = HScroll3.Value
    Text1.Text = HScroll3.Value
    smalldiameter = ((Picture2.Width) / 2) / 2
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    Else
    bigdiameter = HScroll3.Value / 2
    smalldiameter = ((Picture2.Height) / 2) / 2
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    End If
    
    d = 1
    c = 0
   
    'bigdiameter = (Picture2.Width) / 2
    'smalldiameter = ((Picture2.Width) / 2) / 2 / 2
    CircularOffSet1 = HScroll2.Value / 100
    CircularOffSet2 = 0
    Dim myarray1() As Integer
    ReDim myarray1(numberofpeaks, 2, 2)
    
    Picture2.AutoRedraw = False
    
    
    
    'Do
    For angle = 0 To 6.3 Step 0.01
          
          'clockwise
          a = Sin(6.3 - angle) * bigdiameter
          b = Cos(6.3 - angle) * bigdiameter
          a1 = Sin(6.3 - angle) * smalldiameter
          b1 = Cos(6.3 - angle) * smalldiameter
    
    
    Select Case Format(angle, "0.00")
    
    Case Format((6.3 / numberofpeaks * d) - CircularOffSet1, "0.00") '1.57, 3....
        SetPixel Picture2.hdc, a + SW, b + SH, 0
        
        myarray1(d, 1, 1) = a + SW
        myarray1(d, 1, 2) = b + SH
        d = d + 1
        
    Case Format((6.3 / numberofpeaks / 2) + ((6.3 / numberofpeaks) * c) - CircularOffSet2, "0.00") '0.79, 2.37, 3.94, 5.52
        
    'Case 0.79, 2.37, 3.94, 5.52
        'If Format(angle, "0.00") = 2.37 Then Stop
    g = Format((6.3 / numberofpeaks / 2) + ((6.3 / numberofpeaks) * c), "0.00")
         c = c + 1
        SetPixel Picture2.hdc, a1 + SW, b1 + SH, 0
        myarray1(c, 2, 1) = a1 + SW
        myarray1(c, 2, 2) = b1 + SH
       
    Case Else
    End Select
    Next angle
    
    'Exit Sub
    For h = 1 To UBound(myarray1, 1)
        'For z = 0 To 5000000
        'Next z
        Picture2.Line (myarray1(h, 2, 1), myarray1(h, 2, 2))-(myarray1(h, 1, 1), myarray1(h, 1, 2)), 0
            If h < UBound(myarray1, 1) Then
            Picture2.Line (myarray1(h, 1, 1), myarray1(h, 1, 2))-(myarray1(h + 1, 2, 1), myarray1(h + 1, 2, 2)), 0
            Else
            Picture2.Line (myarray1(UBound(myarray1, 1), 1, 1), myarray1(UBound(myarray1, 1), 1, 2))-(myarray1(1, 2, 1), myarray1(1, 2, 2)), 0
            End If
    Next h

Case Else
End Select
End Sub

Private Sub HScroll4_Scroll()
Select Case ActiveShape
Case Is = "SimetricalMultiSideShape"
HScroll2.Max = 6.3 / HScroll1.Value * 49

Text1.Text = HScroll2.Value
Picture2.Cls

    If Picture2.Width < Picture2.Height Then
    bigdiameter = (Picture2.Width) / 2
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    Else
    bigdiameter = (Picture2.Height) / 2
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    End If
    
    d = 1

    Dim myarray() As Integer
    ReDim myarray(HScroll1.Value, 2)
    
    Picture2.AutoRedraw = False
    
    
    
    'Do
    For angle = 0 To 6.3 Step 0.01
          'clockwise
          a = Sin(6.3 - angle) * bigdiameter
          b = Cos(6.3 - angle) * bigdiameter
        
    Select Case Format(angle, "0.00")
    
    Case Format((6.3 / HScroll1.Value * d) - (HScroll2.Value / 50), "0.00") '1.57, 3....
        SetPixel Picture2.hdc, a + SW, b + SH, 0
        
        myarray(d, 1) = a + SW
        myarray(d, 2) = b + SH
        d = d + 1
    Case Else
    End Select
    
    Next angle
    
    'Exit Sub
    For h = 1 To UBound(myarray, 1)
        'For z = 0 To 5000000
        'Next z
        If h < UBound(myarray, 1) Then
        Picture2.Line (myarray(h, 1), myarray(h, 2))-(myarray(h + 1, 1), myarray(h + 1, 2)), 0
        Else
        Picture2.Line (myarray(h, 1), myarray(h, 2))-(myarray(1, 1), myarray(1, 2)), 0
        End If
    Next h

Case Is = "Starrs"
'Text1.Text = HScroll2.Value
Picture2.Cls
numberofpeaks = HScroll1.Value
HScroll2.Max = 6.3 / HScroll1.Value * 45
    If Picture2.Width < Picture2.Height Then
    bigdiameter = HScroll3.Value
    Text1.Text = HScroll3.Value
    smalldiameter = HScroll4.Value
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    Else
    bigdiameter = bigdiameter
    smalldiameter = HScroll4.Value / 2
    'center point of our starr
    SW = Picture2.Width / 2
    SH = Picture2.Height / 2
    End If
    
    d = 1
    c = 0
   
    'bigdiameter = (Picture2.Width) / 2
    'smalldiameter = ((Picture2.Width) / 2) / 2 / 2
    CircularOffSet1 = HScroll2.Value / 100
    CircularOffSet2 = 0
    Dim myarray1() As Integer
    ReDim myarray1(numberofpeaks, 2, 2)
    
    Picture2.AutoRedraw = False
    
    
    
    'Do
    For angle = 0 To 6.3 Step 0.01
          
          'clockwise
          a = Sin(6.3 - angle) * bigdiameter
          b = Cos(6.3 - angle) * bigdiameter
          a1 = Sin(6.3 - angle) * smalldiameter
          b1 = Cos(6.3 - angle) * smalldiameter
    
    
    Select Case Format(angle, "0.00")
    
    Case Format((6.3 / numberofpeaks * d) - CircularOffSet1, "0.00") '1.57, 3....
        SetPixel Picture2.hdc, a + SW, b + SH, 0
        
        myarray1(d, 1, 1) = a + SW
        myarray1(d, 1, 2) = b + SH
        d = d + 1
        
    Case Format((6.3 / numberofpeaks / 2) + ((6.3 / numberofpeaks) * c) - CircularOffSet2, "0.00") '0.79, 2.37, 3.94, 5.52
        
    'Case 0.79, 2.37, 3.94, 5.52
        'If Format(angle, "0.00") = 2.37 Then Stop
    g = Format((6.3 / numberofpeaks / 2) + ((6.3 / numberofpeaks) * c), "0.00")
         c = c + 1
        SetPixel Picture2.hdc, a1 + SW, b1 + SH, 0
        myarray1(c, 2, 1) = a1 + SW
        myarray1(c, 2, 2) = b1 + SH
       
    Case Else
    End Select
    Next angle
    
    'Exit Sub
    For h = 1 To UBound(myarray1, 1)
        'For z = 0 To 5000000
        'Next z
        Picture2.Line (myarray1(h, 2, 1), myarray1(h, 2, 2))-(myarray1(h, 1, 1), myarray1(h, 1, 2)), 0
            If h < UBound(myarray1, 1) Then
            Picture2.Line (myarray1(h, 1, 1), myarray1(h, 1, 2))-(myarray1(h + 1, 2, 1), myarray1(h + 1, 2, 2)), 0
            Else
            Picture2.Line (myarray1(UBound(myarray1, 1), 1, 1), myarray1(UBound(myarray1, 1), 1, 2))-(myarray1(1, 2, 1), myarray1(1, 2, 2)), 0
            End If
    Next h

Case Else
End Select
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  'KPD-Team 1999
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    'Create a solid brush
    mbrush = CreateSolidBrush(&HFF&)       'black
    'Select the brush into the PictureBox' device context
    SelectObject Picture2.hdc, mbrush
  'Floodfill...
    ExtFloodFill Picture2.hdc, x, y, GetPixel(Picture2.hdc, x, y), FLOODFILLSURFACE
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Text1.Text = x & "," & y
Text2.Text = GetPixel(Picture2.hdc, x, y)
End Sub

Private Sub VScroll1_Scroll()
Text1.Text = VScroll1.Value
Picture2.AutoRedraw = True
Picture2.Cls
DoEvents
Picture2.AutoRedraw = True
diameter = (Picture2.Height - 1) / 2

'nariše krog - make egg

For kot = 0 To 20 Step 0.01

a = Sin(kot) * diameter * Cos(Sin(kot / 2)) * (HScroll1.Value / 50)
b = Cos(kot) * diameter * (VScroll1.Value / 10000)
SetPixel Picture2.hdc, a + (Picture2.Width / 2), b + (Picture2.Height / 2), 0
Next kot
Picture2.Refresh
Picture2.AutoRedraw = False

End Sub
