VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Tanner's transparency demonstration"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4290
   Icon            =   "transparency.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "transparency.frx":0442
   ScaleHeight     =   5985
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear the Picture Box"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Text            =   "50"
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton CmdTransparency 
      Caption         =   "Build Transparency"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   3975
      Left            =   120
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   2
      Top             =   1200
      Width           =   3735
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   1200
      Picture         =   "transparency.frx":0784
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   120
      Width           =   1020
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   120
      Picture         =   "transparency.frx":37C6
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "%"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Any-percent transparency by Tanner Helland

'This pretty simple code demonstrates how to combine two bitmaps to
'create a perfect transparency between the two at any value.  For
'demonstration purposes, use the values of 0,25,50,75,and 100 to see
'how this works.  The code is really simple, and could be used to make
'some very nice effects inside of a game if rebuilt in ASM.  If you
'have any questions or comments, contact me at
'tannerhelland@hotmail.com.  Enjoy the code!


'Windows API - much faster then VB's PSet and Point
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
'variables for position, colors, etc.
Dim x, y As Integer
Dim color1, color2 As Long
Dim r, g, b As Integer
Dim r2, g2, b2 As Integer
Dim percent As Integer
Private Sub CmdTransparency_Click()
'get the percent value from the text box
percent = Text1.Text
'set the 3rd picture box to the appropriate size
Picture3.Width = Picture1.Width
Picture3.Height = Picture1.Height
'run a loop through the picture
For x = 0 To Picture1.ScaleWidth - 1
For y = 0 To Picture1.ScaleHeight - 1
'get the color of the first picture and extract the R,G,B values
color1 = GetPixel(Picture1.hDC, x, y)
r = color1 Mod 256
b = Int(color1 / 65536)
g = (color1 - (b * 65536) - r) / 256
'get the color of the second picture and extract the R,G,B values
color2 = GetPixel(Picture2.hDC, x, y)
r2 = color2 Mod 256
b2 = Int(color2 / 65536)
g2 = (color2 - (b2 * 65536) - r2) / 256
'mix the colors based on the specified percent to create a new color
'that's a perfect combination of the previous two
r = (((100 - percent) * r) + (percent * r2)) / 100
g = (((100 - percent) * g) + (percent * g2)) / 100
b = (((100 - percent) * b) + (percent * b2)) / 100
'set the new color onto the 3rd picture box
SetPixel Picture3.hDC, x, y, RGB(r, g, b)
'continue through the loop
Next y
'refresh the picture box every 10 lines (a nice progress bar effect)
If x Mod 10 = 0 Then Picture3.Refresh
Next x
'refresh the picture
Picture3.Refresh
End Sub

Private Sub CmdClear_Click()
'clear the 3rd picture box
Picture3.Picture = LoadPicture()
Picture3.Cls
End Sub

Private Sub Picture1_Click()
'load a picture into the box
CommonDialog1.ShowOpen
Picture1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub Picture2_Click()
'load picture into the box
CommonDialog1.ShowOpen
Picture2.Picture = LoadPicture(CommonDialog1.FileName)
End Sub
