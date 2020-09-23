VERSION 5.00
Begin VB.Form FastLens 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Lens  by Scythe       Press ESC to QUIT"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8970
   ControlBox      =   0   'False
   ForeColor       =   &H000100FF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   598
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox ChkAutomove 
      Caption         =   "Automove"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   7200
      Value           =   1  'Aktiviert
      Width           =   1935
   End
   Begin VB.HScrollBar HscrMagnify 
      Height          =   255
      LargeChange     =   10
      Left            =   1200
      Max             =   5
      Min             =   100
      TabIndex        =   4
      Top             =   7200
      Value           =   30
      Width           =   1215
   End
   Begin VB.HScrollBar HScrSize 
      Height          =   255
      LargeChange     =   10
      Left            =   0
      Max             =   400
      Min             =   20
      TabIndex        =   2
      Top             =   7200
      Value           =   200
      Width           =   1095
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000100FF&
      Height          =   6750
      Left            =   0
      ScaleHeight     =   450
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   1
      Top             =   0
      Width           =   9000
   End
   Begin VB.PictureBox PicBack 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000100FF&
      Height          =   6750
      Left            =   0
      Picture         =   "Lens.frx":0000
      ScaleHeight     =   450
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3960
      Top             =   3120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   3480
   End
   Begin VB.Label Label2 
      Caption         =   "Magnify"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Size"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6960
      Width           =   495
   End
End
Attribute VB_Name = "FastLens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Simple Lens Demo
' by scythe scythe@cablenet.de

' Compile for real speed

'This demo uses a precalculated lens
'The array LookUp hold the difference between
'the point to set and the point to read
'We dont need to calculate the whole thing every cycle
'All we need is to draw the lens

Option Explicit

'To copy our pic real fast
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Use DIB for fast GFX
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Type RGBQUAD
 rgbBlue As Byte
 rgbGreen As Byte
 rgbRed As Byte
 rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
 biSize           As Long
 biWidth          As Long
 biHeight         As Long
 biPlanes         As Integer
 biBitCount       As Integer
 biCompression    As Long
 biSizeImage      As Long
 biXPelsPerMeter  As Long
 biYPelsPerMeter  As Long
 biClrUsed        As Long
 biClrImportant   As Long
End Type

Private Type BITMAPINFO
 bmiHeader As BITMAPINFOHEADER
End Type

Private Const DIB_RGB_COLORS As Long = 0

Private Type PointApi
 X As Long
 Y As Long
End Type

Dim LookUp() As PointApi      'Table for precalculatet Lens

Dim PicNew()  As RGBQUAD      'Hold our New Picture
Dim PicOrg()  As RGBQUAD      'Hold our Original Picute
Dim Binfo     As BITMAPINFO   'The GetDIBits API needs some Infos
Dim OrgLng    As Long         'Holds the Lenght of the Picture
Dim Drawing   As Boolean      'Is the program at work
Dim MoveX     As Long         'Holds X in Automove Mode
Dim MoveY     As Long         'Holds Y in Automove Mode
Dim DirX      As Byte         'Holds DirectionX in Automove Mode
Dim DirY      As Byte         'Holds DirectionY in Automove Mode

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
  Timer1.Enabled = False
  Unload Me
  End
 End If
End Sub

Private Sub Form_Load()

 'Create a buffer that holds our picture
 ReDim PicNew(0 To Pic.ScaleWidth - 1, 0 To Pic.ScaleHeight - 1)
 ReDim PicOrg(0 To Pic.ScaleWidth - 1, 0 To Pic.ScaleHeight - 1)

 'Get the Picturesize in Memory for CopyMemory
 'X*Y*4 (4 for the 4 Bytes of RGBQUAD)
 OrgLng = (UBound(PicOrg, 1) + 1) * (UBound(PicOrg, 2) + 1) * 4

 'Set the infos for our apicall
 With Binfo.bmiHeader
 .biSize = 40
 .biWidth = Pic.ScaleWidth
 .biHeight = Pic.ScaleHeight
 .biPlanes = 1
 .biBitCount = 32
 .biCompression = 0
 .biClrUsed = 0
 .biClrImportant = 0
 .biSizeImage = Pic.ScaleWidth * Pic.ScaleHeight
 End With

 'If we start in ide show a message
 If InIde = True Then
  PicBack.CurrentX = 100
  PicBack.CurrentY = 50
  PicBack.AutoRedraw = True
  PicBack.Print "Please compile to get full SPEED"
  PicBack.AutoRedraw = False
 End If

 'Now get the Original Picture
 GetDIBits PicBack.hdc, PicBack.Image.Handle, 0, Binfo.bmiHeader.biHeight, PicOrg(0, 0), Binfo, DIB_RGB_COLORS

 'Copy the picture from PicBack to Pic
 ShowPic

 'Calculate our Lens
 CreateLens LookUp, HScrSize.Value, HscrMagnify.Value

 'set direction and position for Automove
 DirX = 1
 DirY = 1
 MoveX = Pic.Width / 2
 MoveY = Pic.Height / 2

End Sub

Private Sub DrawLens(ByVal SourceX As Long, ByVal SourceY As Long)
 Dim X As Long
 Dim Y As Long
 Dim StartX As Long
 Dim StartY As Long
 Dim EndX As Long
 Dim EndY As Long

 'Tell the program that we draw
 Drawing = True

 'Center our lens
 SourceX = SourceX - UBound(LookUp) / 2
 SourceY = SourceY - UBound(LookUp) / 2

 'Ok if we move the lens out of the Picture
 'Draw only the vissible part
 StartX = 2
 EndX = UBound(LookUp, 1) - 1
 If SourceX < 0 Then
  StartX = Abs(SourceX)
  ElseIf SourceX > Pic.Width - EndX Then
  EndX = (Pic.Width - SourceX)
 End If
 StartY = 1
 EndY = UBound(LookUp, 1) - 1
 If SourceY < 0 Then
  StartY = Abs(SourceY)
  ElseIf SourceY > Pic.Height - EndY Then
  EndY = (Pic.Height - SourceY)
 End If


On Error Resume Next
'Get the picture to paint on
CopyMemory PicNew(0, 0), PicOrg(0, 0), OrgLng

For X = StartX To EndX
 For Y = StartY To EndY
  'Now we use our Array
  'Set a new point to x,y
  'Get the position on the original picture from our Array
  PicNew(X + SourceX - 1, Y + SourceY) = PicOrg(LookUp(X, Y).X + SourceX, LookUp(X, Y).Y + SourceY)
 Next Y
Next X

'Show our Lens
SetDIBits Pic.hdc, Pic.Image.Handle, 0, Binfo.bmiHeader.biHeight, PicNew(0, 0), Binfo, DIB_RGB_COLORS
Pic.Refresh
Drawing = False
DoEvents
End Sub


Private Sub HscrMagnify_Change()
 Timer2.Enabled = True
End Sub

Private Sub HScrSize_Change()
 HscrMagnify.Min = HScrSize / 2
 Timer2.Enabled = True
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Drawing = False And Timer1.Enabled = False Then
 DrawLens X, Pic.Height - Y
 End If
End Sub

'Show the Startpicture
Private Sub ShowPic()
 'copy the dib we got for our hidden picture to the front and clear the front
 SetDIBits Pic.hdc, Pic.Image.Handle, 0, Binfo.bmiHeader.biHeight, PicOrg(0, 0), Binfo, DIB_RGB_COLORS
 Pic.Refresh
End Sub


Private Sub CreateLens(ByRef LensArray() As PointApi, Diameter As Long, Magnification As Single)
 'Some simple Math :o)
 'to calculate our Lens

 'Thx to
 'Joey ????
 '   and
 'Jeff Lawson
 'for the Infos they posted about Lenscalculations

 Dim Radius As Integer
 Dim Sphere As Single
 Dim X As Long
 Dim Y As Long
 Dim XOld As Long
 Dim YOld As Long
 Dim XNew As Long
 Dim YNew As Long
 Dim Z As Long
 Dim A As Long
 Dim B As Long
 Dim tmp1 As Long
 Dim tmp As Long

 Radius = Diameter / 2

 Sphere = Sqr(Radius * Radius - Magnification * Magnification)

 ReDim LensArray(Diameter, Diameter)


 For X = -Radius To -Radius + Diameter - 1
  For Y = -Radius To -Radius + Diameter - 1
   If X * X + Y * Y >= Sphere * Sphere Then
    A = X
    B = Y
   Else
    Z = Sqr(Radius * Radius - X * X - Y * Y)
    A = Int(X * Magnification / Z + 0.5)
    B = Int(Y * Magnification / Z + 0.5)
   End If
   tmp1 = (1 + (Y + Radius) * Diameter + (X + Radius))
   YOld = CInt(tmp1 / Diameter - 0.5)
   XOld = CInt(tmp1 - YOld * Diameter)
   tmp = (B + Radius) * Diameter + (A + Radius)
   YNew = CInt(tmp / Diameter - 0.5)
   XNew = CInt(tmp - YNew * Diameter)
   If XNew = 200 Then
    X = X
   End If
   LensArray(XOld, YOld).X = XNew
   LensArray(XOld, YOld).Y = YNew
  Next Y
 Next X
End Sub

'Move Our Lens automatic
Private Sub Timer1_Timer()
 Dim Speed As Byte
 'Scrollspeed
 Speed = 5

 If Drawing = False Then
  If DirX = 1 Then
   If MoveX < Pic.Width - HScrSize.Value / 2 Then
    MoveX = MoveX + Speed
   Else
    DirX = 0
    MoveX = MoveX - Speed
   End If
  Else
   If MoveX > HScrSize.Value / 2 Then
    MoveX = MoveX - Speed
   Else
    DirX = 1
    MoveX = MoveX + Speed
   End If
  End If
  If DirY = 1 Then
   If MoveY < Pic.Height - HScrSize.Value / 2 Then
    MoveY = MoveY + Speed
   Else
    DirY = 0
    MoveY = MoveY - Speed
   End If
  Else
   If MoveY > HScrSize.Value / 2 Then
    MoveY = MoveY - Speed
   Else
    DirY = 1
    MoveY = MoveY + Speed
   End If
  End If
  DrawLens MoveX, MoveY
 End If
End Sub

Private Sub Timer2_Timer()
Dim tmp As Boolean
 'if we change the Lenssize while the code draws then
 'dont calculate a new
 If Drawing = False Then
  tmp = Timer1.Enabled
  'Turn Createtimer off
  Timer2.Enabled = False
  'Turn Movetimer off
  Timer1.Enabled = False
  'Create new Lens
  CreateLens LookUp, HScrSize.Value, HscrMagnify.Value
  'Turn Movetimer on
  Timer1.Enabled = tmp
 End If
End Sub

'Test if we are in ide or compiled mode
Private Function InIde() As Boolean
 On Error GoTo DivideError
 Debug.Print 1 / 0
 Exit Function
DivideError:
 InIde = True
End Function

Private Sub ChkAutomove_Click()
 If ChkAutomove.Value = 0 Then Timer1.Enabled = False Else Timer1.Enabled = True
End Sub

