VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Firestorm Palette Generator"
   ClientHeight    =   3975
   ClientLeft      =   2220
   ClientTop       =   1755
   ClientWidth     =   6090
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   406
   Begin VB.CommandButton Command6 
      Caption         =   "Load palette"
      Height          =   495
      Left            =   1800
      TabIndex        =   15
      Top             =   3420
      Width           =   1050
   End
   Begin VB.PictureBox picBMP 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5430
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   14
      Top             =   3465
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Convert palette"
      Height          =   495
      Left            =   4035
      TabIndex        =   13
      Top             =   3420
      Width           =   1410
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New palette"
      Height          =   495
      Left            =   675
      TabIndex        =   12
      Top             =   3420
      Width           =   1050
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save palette"
      Height          =   495
      Left            =   2910
      TabIndex        =   1
      Top             =   3420
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   2625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   1170
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   2595
      Width           =   3840
      Begin VB.Line Line1 
         DrawMode        =   6  'Mask Pen Not
         X1              =   10
         X2              =   10
         Y1              =   0
         Y2              =   50
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2445
      Left            =   90
      TabIndex        =   2
      Top             =   75
      Width           =   5910
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   270
         TabIndex        =   7
         Top             =   450
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add / Update Item"
         Height          =   420
         Left            =   2250
         TabIndex        =   6
         Top             =   465
         Width           =   1905
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         Left            =   2280
         Max             =   255
         TabIndex        =   5
         Top             =   2040
         Value           =   10
         Width           =   3390
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         Height          =   570
         Left            =   2295
         ScaleHeight     =   510
         ScaleWidth      =   540
         TabIndex        =   4
         Top             =   1200
         Width           =   600
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove Last Item"
         Height          =   420
         Left            =   4260
         TabIndex        =   3
         Top             =   465
         Width           =   1500
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3240
         Picture         =   "Form1.frx":1042
         Top             =   1215
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "Colour:"
         Height          =   195
         Left            =   2310
         TabIndex        =   11
         Top             =   975
         Width           =   1560
      End
      Begin VB.Label Label2 
         Caption         =   "Position:"
         Height          =   300
         Left            =   2280
         TabIndex        =   10
         Top             =   1830
         Width           =   2010
      End
      Begin VB.Label Label3 
         Caption         =   "Gradient Stack:"
         Height          =   300
         Left            =   270
         TabIndex        =   9
         Top             =   210
         Width           =   2010
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1080
         TabIndex        =   8
         Top             =   2280
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' FIRESTORM PALETTE GENERATOR
' ---------------------------
' by Michael Pote:
' michaelpote@worldonline.co.za
'
' To use:
'
' Move the position scrollbar and you will see a
' line move across the black picturebox at the bottom
' of the window.

' Move this slider to somewhere then click on the Colour
' picturebox above the slider. This brings up the color dialog.

' When you're happy with the colour, click 'Add/Update item'

' If you make a mistake and want to change the values after
' adding a gradient to the gradient stack, click on an item
' in the stack and change it's values then click the 'Add/Update'
' button. When you want to place new gradient points,
' click on the 'Add new item' item in the gradient stack.

' The button 'Remove last item' removes the last item from
' the stack

' 'Convert palette' converts old firestorm palettes (*.bmp)
' to new firestorm palettes (*.pal)
' The new palette format is 40 times smaller and more robust
' than the old format.
' The convert palette feature can also convert palettes of
' any 256 color bitmap.

' NOTES ON PALETTE FORMATS:
' -------------------------
'
' You can save a palette in either the Firestorm readable format
' (*.pal) or as a Gradient stack (*.grd) which is readable by
' the palette generator Only. The Palette generator cannot load
' Firestorm readable palettes.

Private Type Gradient
Pos As Integer
R As Long
G As Long
B As Long
End Type

Dim Ind As Long
Private G() As Gradient, GC As Integer
Private Pal(0 To 255) As RGBQUAD

Private Sub Command1_Click()
If List1.ListIndex = -1 Then
ReDim Preserve G(0 To GC) As Gradient
G(GC).R = rgbRed(Picture2.BackColor)
G(GC).G = rgbGreen(Picture2.BackColor)
G(GC).B = rgbBlue(Picture2.BackColor)
G(GC).Pos = HScroll1.Value
List1.List(List1.ListCount - 1) = "Step " & GC
List1.AddItem "Add New Item"
GC = GC + 1
Else

With G(List1.ListIndex)
.Pos = HScroll1.Value
.R = rgbRed(Picture2.BackColor)
.G = rgbGreen(Picture2.BackColor)
.B = rgbBlue(Picture2.BackColor)
End With

End If

DrawGradient
End Sub


Public Sub DrawGradient()
Dim Pos As Long, Cr As Single, Cg As Single, Cb As Single, Ri As Single, Gi As Single, Bi As Single
Dim I As Long, J As Long, D As Boolean

Pos = 0
Do
DoEvents

For I = 0 To GC - 2
If G(I).Pos = Pos Then
Ri = CDbl(G(I + 1).R - G(I).R) / (G(I + 1).Pos - Pos)
Gi = CDbl(G(I + 1).G - G(I).G) / (G(I + 1).Pos - Pos)
Bi = CDbl(G(I + 1).B - G(I).B) / (G(I + 1).Pos - Pos)
Exit For
End If
Next


Cb = Abs(Cb + Bi)
Cr = Abs(Cr + Ri)
Cg = Abs(Cg + Gi)

If Cr > 255 Then Cr = 255
If Cg > 255 Then Cg = 255
If Cb > 255 Then Cb = 255

Pal(Pos).rgbBlue = Cb
Pal(Pos).rgbRed = Cr
Pal(Pos).rgbGreen = Cg

Picture1.Line (Pos, 0)-(Pos, 50), RGB(Cr, Cg, Cb)
Pos = Pos + 1
Loop Until Pos >= 256

End Sub

Private Sub Command3_Click()
List1.Clear
ReDim G(0 To 0) As Gradient
GC = 1
List1.AddItem "Background"
List1.AddItem "Add New Item"
DrawGradient
End Sub

Private Sub Command2_Click()
If GC = 1 Then MsgBox "Cannot delete!": Exit Sub
Dim G2() As Gradient, I As Long
ReDim G2(0 To GC - 1) As Gradient

For I = 0 To GC - 1
G2(I).B = G(I).B
G2(I).R = G(I).R
G2(I).G = G(I).G
G2(I).Pos = G(I).Pos
Next

ReDim G(0 To GC - 1) As Gradient
GC = GC - 1

For I = 0 To GC
G(I).B = G2(I).B
G(I).R = G2(I).R
G(I).G = G2(I).G
G(I).Pos = G2(I).Pos
Next

List1.RemoveItem GC
DrawGradient
End Sub

Private Sub Command4_Click()
CD.Filter = "Firestorm Palette|*.pal|Gradient Stack|*.grd"
CD.ShowSave

If CD.Filename = "" Then Exit Sub

DrawGradient

If LCase(Right(CD.Filename, 3)) = "pal" Then

Open CD.Filename For Binary As #1
Put #1, , Pal
Close #1

Else

Open CD.Filename For Binary As #1
Put #1, , GC
Put #1, , G
Close #1


End If

End Sub

Private Sub Command5_Click()
Dim I As Long
CD.Filter = "Bitmap Palettes|*.bmp|Firestorm Palettes|*.pal"
CD.ShowOpen
If CD.Filename = "" Then Exit Sub

If LCase(Right(CD.Filename, 3)) = "pal" Then

Open CD.Filename For Binary As #1
Get #1, , Pal
Close #1

Else

GetPal CD.Filename, Pal()

End If

CD.Filename = ""
CD.Filter = "Firestorm Palettes|*.pal"
CD.ShowSave
If CD.Filename = "" Then Exit Sub

Open CD.Filename For Binary As #1
Put #1, , Pal
Close #1


End Sub

Sub DrawPal()
Command3_Click

ReDim G(0 To 255) As Gradient
For I = 0 To 255 Step 2
G(I).R = Pal(I).rgbRed
G(I).G = Pal(I).rgbGreen
G(I).B = Pal(I).rgbBlue
G(I).Pos = I
List1.List(List1.ListCount - 1) = "Step " & I
List1.AddItem "Add New Item"
Next
GC = 128
DrawGradient
End Sub

Private Sub Command6_Click()
Dim I As Long
CD.Filter = "Gradient Stack|*.grd"
CD.ShowOpen

If CD.Filename = "" Then Exit Sub

Open CD.Filename For Binary As #1
Get #1, , GC
ReDim G(0 To GC - 1) As Gradient
Get #1, , G
Close #1

GC = UBound(G) + 1

List1.Clear
For I = 0 To GC
If I = 0 Then
List1.AddItem "Background"
ElseIf I = GC Then
List1.AddItem "Add new item"
Else
List1.AddItem "Item " & I
End If
Next

DrawGradient

End Sub

Private Sub Form_Load()
ReDim G(0 To 0) As Gradient
GC = 1
List1.AddItem "Background"
List1.AddItem "Add New Item"
G(0).Pos = 0
End Sub

Private Sub HScroll1_Change()
HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
Line1.X1 = HScroll1.Value
Line1.X2 = Line1.X1
End Sub

Private Sub List1_Click()
On Error Resume Next
Ind = List1.ListIndex
If Ind = GC Then List1.ListIndex = -1: Exit Sub
HScroll1.Min = 0
HScroll1.Value = G(Ind).Pos
Picture2.BackColor = RGB(G(Ind).R, G(Ind).G, G(Ind).B)
End Sub

Private Sub Picture2_Click()
CD.Color = Picture2.BackColor
CD.ShowColor
Picture2.BackColor = CD.Color
End Sub
