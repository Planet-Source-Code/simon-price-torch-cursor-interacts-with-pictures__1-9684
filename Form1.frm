VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Torch Cursor - by Simon Price"
   ClientHeight    =   5028
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   7320
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   610
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog ComDialog 
      Left            =   840
      Top             =   3360
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      Filter          =   "*.bmp, *.jpg, *.jpeg"
   End
   Begin VB.CommandButton cmdLoadPic 
      Caption         =   "Load Another Background Picture"
      Height          =   732
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   1932
   End
   Begin VB.HScrollBar BScroll 
      Height          =   372
      Left            =   2880
      Max             =   20
      Min             =   1
      TabIndex        =   5
      Top             =   3600
      Value           =   10
      Width           =   3612
   End
   Begin VB.HScrollBar RScroll 
      Height          =   372
      LargeChange     =   10
      Left            =   2880
      Max             =   200
      Min             =   20
      TabIndex        =   4
      Top             =   4560
      Value           =   50
      Width           =   3612
   End
   Begin VB.PictureBox Display 
      BackColor       =   &H0000FFFF&
      Height          =   3168
      Left            =   120
      MouseIcon       =   "Form1.frx":0152
      MousePointer    =   99  'Custom
      ScaleHeight     =   260
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   590
      TabIndex        =   0
      Top             =   120
      Width           =   7128
      Begin VB.PictureBox PB 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3168
         Left            =   1560
         ScaleHeight     =   264
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   594
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   7128
         Begin VB.PictureBox TorchPic 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
            Height          =   2400
            Left            =   1920
            ScaleHeight     =   200
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   200
            TabIndex        =   3
            Top             =   600
            Visible         =   0   'False
            Width           =   2400
         End
      End
      Begin VB.PictureBox Original 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3168
         Left            =   480
         ScaleHeight     =   264
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   594
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   7128
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Large"
      Height          =   192
      Left            =   6600
      TabIndex        =   11
      Top             =   4680
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Small"
      Height          =   192
      Left            =   2280
      TabIndex        =   10
      Top             =   4680
      Width           =   408
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Light"
      Height          =   192
      Left            =   6600
      TabIndex        =   9
      Top             =   3720
      Width           =   336
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Dark"
      Height          =   192
      Left            =   2280
      TabIndex        =   8
      Top             =   3720
      Width           =   348
   End
   Begin VB.Label BLabel 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Brightness"
      Height          =   240
      Left            =   2280
      TabIndex        =   7
      Top             =   3360
      Width           =   804
   End
   Begin VB.Label RLabel 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Size"
      Height          =   240
      Left            =   2280
      TabIndex        =   6
      Top             =   4320
      Width           =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This program lights up a picture around
' where the mousepointer is. It's a nice
' effect and teaches you about RGB colours,
' some API and backbuffers.

' I recommend that you run this from the exe
' because it runs alot faster

' Please vote for it at planet source code
' if you think it's cool, interesting, well
' done or useful - I haven't found a use for
' it yet but I'm sure someone will. Maybe
' you could use it in a paint program

' By Simon Price
' Email : Si@VBgames.co.uk
' Website : www.VBgames.co.uk


'declarations...
'get pixel allows us to look at pixel colours
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
'set pixel lets us draw pixels
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'bitblt can copy a picture from one place to another
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'types...
'this type holds the red, green and blue
'values of an RGB colour
Private Type tRGBcolor
  R As Byte 'red
  G As Byte 'green
  B As Byte 'blue
End Type

'variables...
'temporary variables to hold current colours
Dim LongCol As Long 'long colour is needed for API's
Dim RGBcol As tRGBcolor 'RGB colour is needed to make the effect
'file path of the picture used
Dim FilePath As String
'this buffer represents the amount of light in a circle
Dim LightBuffer() As Byte
'the size of the effect
Dim Radius As Integer
'the brightness of the effect
Dim Brightness As Single
'if we've loaded or not
Dim Loaded As Boolean

Sub LoadTestPic() 'loads the background picture
On Error GoTo MuffUp
'if there's no filepath then use default one
If FilePath = "" Then FilePath = App.Path & "\the_jump.jpg"
'load the picture
Original = LoadPicture(FilePath)
'show in display
Display = Original
Exit Sub
MuffUp:
MsgBox "There was an error when loading the picture", vbCritical, "Error!"
End Sub

Private Sub BScroll_Change()
Brightness = BScroll.Value / 10
BLabel = Brightness
End Sub

Private Sub cmdLoadPic_Click()
'show open
ComDialog.ShowOpen
FilePath = ComDialog.FileName
LoadTestPic
End Sub

Private Sub Display_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this is where the torch effect happens
Dim x2, y2, L, i As Integer
Dim Col As tRGBcolor
Dim LongCol As Long

On Error Resume Next

'if loading is not finished, do nothing
If Loaded = False Then Exit Sub

'copy to backbuffer
BitBlt PB.hdc, 0, 0, PB.Width, PB.Height, Original.hdc, 0, 0, vbSrcCopy

'loop through each pixel in light buffer
For x2 = X - Radius To X + Radius
For y2 = Y - Radius To Y + Radius
  'get light amount
  L = LightBuffer(x2 - X + Radius, y2 - Y + Radius)
  If L > 0 Then 'if there's some light
     'get the current pixel
        LongCol = GetPixel(Original.hdc, x2, y2)
     'convert it to RGB color
        Col.R = LongCol And 255
        Col.G = (LongCol And 65280) \ 256&
        Col.B = (LongCol And 16711680) \ 65535
     'calculate amount of light
        L = L * Brightness
     'make the colour brighter
        i = Col.R + L
        If i > 255 Then i = 255
        Col.R = i
        i = Col.G + L
        If i > 255 Then i = 255
        Col.G = i
        i = Col.B + L
        If i > 255 Then i = 255
        Col.B = i
     'now draw in with new brighter colour
        SetPixel PB.hdc, x2, y2, RGB(Col.R, Col.G, Col.B)
   End If
Next
Next

'now copy from buckbuffer into view
BitBlt Display.hdc, 0, 0, PB.Width, PB.Height, PB.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub Form_Load()
Show
DoEvents
LoadTestPic
Radius = 50
Brightness = 1
LoadLightBuffer
End Sub

Sub LoadLightBuffer()
'fills light buffer with values
Dim i As Integer
Dim X, Y As Integer
Dim Col As Long
'show that we're loading
Loaded = False
MousePointer = vbHourglass
Caption = "LOADING... Please Wait"
'change to torch size
TorchPic.Move 0, 0, Radius * 2, Radius * 2
'now draw circles on the torch pic
TorchPic.Cls
For i = 14 To 0 Step -1
  TorchPic.FillColor = QBColor(i)
  TorchPic.Circle (Radius, Radius), Radius / 15 * (i + 1), QBColor(i)
Next
'resize the light buffer
ReDim LightBuffer(0 To Radius * 2, 0 To Radius * 2)
'now read these circles into the light buffer
For X = 0 To Radius * 2
   DoEvents
For Y = 0 To Radius * 2
   For i = 0 To 14
     Col = GetPixel(TorchPic.hdc, X, Y)
     If Col = QBColor(i) Then GoTo FillItIn 'if there is a color, put value in light buffer
   Next
   GoTo SkipIt
FillItIn:
   LightBuffer(X, Y) = (15 - i) * 10 'fill in light buffer
SkipIt:
Next
Next
Beep
'change back to normal cursor
MousePointer = vbDefault
Caption = "The Torch Cursor - by Simon Price"
Loaded = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'get rid of the torch effect
BitBlt Display.hdc, 0, 0, PB.Width, PB.Height, Original.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "If you thought that that was a cool effect, then please vote for me on planet source code. Also, look at some of my other submissions, they're better than this one!", vbInformation, "Vote now for the Cool Torch Cursor by Simon Price!"
End Sub

Private Sub RScroll_Change()
Radius = RScroll.Value
RLabel = Radius
LoadLightBuffer
End Sub
