VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.frame frame1 
      Height          =   6495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11456
   End
   Begin VB.PictureBox MaskPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6135
      Index           =   1
      Left            =   0
      Picture         =   "Nature.frx":0000
      ScaleHeight     =   6135
      ScaleWidth      =   8895
      TabIndex        =   4
      Top             =   0
      Width           =   8895
   End
   Begin VB.Timer polygon 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   9240
      Top             =   2160
   End
   Begin VB.PictureBox SourcePic 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   9240
      Picture         =   "Nature.frx":AAF7
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox SourcePic 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   9240
      Picture         =   "Nature.frx":1226D
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox SourcePic 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   9240
      Picture         =   "Nature.frx":1AD11
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox MaskPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6135
      Index           =   2
      Left            =   0
      Picture         =   "Nature.frx":25808
      ScaleHeight     =   6135
      ScaleWidth      =   8895
      TabIndex        =   5
      Top             =   0
      Width           =   8895
   End
   Begin VB.PictureBox MaskPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6135
      Index           =   3
      Left            =   0
      Picture         =   "Nature.frx":2D917
      ScaleHeight     =   6135
      ScaleWidth      =   8895
      TabIndex        =   6
      Top             =   0
      Width           =   8895
   End
   Begin VB.PictureBox MaskPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6135
      Index           =   4
      Left            =   0
      Picture         =   "Nature.frx":37D1E
      ScaleHeight     =   6135
      ScaleWidth      =   8895
      TabIndex        =   7
      Top             =   0
      Width           =   8895
   End
   Begin VB.PictureBox MaskPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6135
      Index           =   5
      Left            =   0
      Picture         =   "Nature.frx":42815
      ScaleHeight     =   6135
      ScaleWidth      =   9015
      TabIndex        =   8
      Top             =   0
      Width           =   9015
   End
   Begin VB.PictureBox MaskPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6135
      Index           =   0
      Left            =   0
      Picture         =   "Nature.frx":4B2B9
      ScaleHeight     =   6135
      ScaleWidth      =   8895
      TabIndex        =   3
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Polygon Pictures v.2    - Chris 2006/2007 -
'
'Did not notice flicker this time around.
'
'The "frame" UC is the topmost mask. UC's with transparent
'backgrounds don't like JPG's as a mask picture, so I used
'a Monochrome bmp instead. Provided is "frame.jpg" you can
'use if you wish. All of the picture boxes are jpg's.
'
'Press any key to quit.
'
'4 enclosed polygons, 4 points each
Dim Points(16) As POINTAPI

Private Sub polygon_Timer()
'we are just cutting holes in the topmost picture box
'to expose the box underneath it, flipping thru the pics
'one-by-one, then repeating the cycle.
Static B As Integer
'
    If B = 0 Then
'initially, at the start of the sequence (B = 0)
        MaskPic(0).ZOrder 0
    Else
        MaskPic(B).ZOrder 0
    End If
'maintain the frame UC ontop
    frame1.ZOrder 0
'timer speed
    polygon.Interval = 1
'the main region
    hHRgn = CreateRoundRectRgn(0, 0, Form1.Width, Form1.Height, 0, 0)
'remember px
Static px As Integer
'start cutting one pixel at a time
    px = px + 1
'create polygonal regions, the first triangle;
  Points(1).X = 608 '<- width
  Points(1).Y = 0

  Points(2).X = 304 '<- half the width
  Points(2).Y = 217 '<- half the height
'***the picture boxes are rectangles, instead of
'squares, so to have the polygons overlap at the
'same time, the speed of the 2 longer sides is
'increased slightly (1.4 vs. 1.0)***
  Points(3).X = 608 - (px * 1.4) '<-
  Points(3).Y = 0
'***********************************
  Points(4).X = 304
  Points(4).Y = 0
'put it to the screen
  hFRgn = CreatePolygonRgn(Points(1), 4, 1)
  CombineRgn hHRgn, hFRgn, hHRgn, RGN_COPY '<- first one is COPY, the rest are OR in this sequence
  'delete the object
  DeleteObject (hFRgn) '<-
  
'the second triangle
  Points(5).X = 0
  Points(5).Y = 438

  Points(6).X = 304
  Points(6).Y = 217

  Points(7).X = 0 + (px * 1.4) '<- the second longer side
  Points(7).Y = 438

  Points(8).X = 0
  Points(8).Y = 438
'put it to the screen
  hFRgn = CreatePolygonRgn(Points(5), 4, 1)
  CombineRgn hHRgn, hFRgn, hHRgn, RGN_OR
  'delete the object
  DeleteObject (hFRgn) '<-
  
'the third triangle
  Points(9).X = 608
  Points(9).Y = 438

  Points(10).X = 304
  Points(10).Y = 217

  Points(11).X = 608
  Points(11).Y = 438 - px

  Points(12).X = 608
  Points(12).Y = 438
'put it to the screen
  hFRgn = CreatePolygonRgn(Points(9), 4, 1)
  CombineRgn hHRgn, hFRgn, hHRgn, RGN_OR
  'delete the object
  DeleteObject (hFRgn) '<-

  'the forth triangle
  Points(13).X = 0
  Points(13).Y = 0

  Points(14).X = 304
  Points(14).Y = 217

  Points(15).X = 0
  Points(15).Y = 0 + px

  Points(16).X = 0
  Points(16).Y = 0
'put it to the screen
  hFRgn = CreatePolygonRgn(Points(13), 4, 1)
  CombineRgn hHRgn, hFRgn, hHRgn, RGN_OR
  'delete the object
  DeleteObject (hFRgn) '<-

'monitor the progress; swap the pics after
'the regions meet + a small delay. (change
'the value of 608 to change the swap time)
  If px = 608 Then
    'reset the counter
    px = 0
    'MaskPic array advance
    B = B + 1
  End If
'combine everything...
  SetWindowRgn MaskPic(B).hWnd, hHRgn, True
  '
  DeleteObject (hHRgn) '<- delete this region or VB will crash
  
  DoEvents
'only done once to prevent flicker
    If px = 0 And Not Begin Then swappictures
'start over from the beginning
    If B = 5 Then B = 0
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'clean-up and quit
    DeleteObject (hFRgn)
    DeleteObject (hHRgn)
    Set Form1 = Nothing
    End
    
End Sub

Private Sub Form_Load()
'make sure the "frame" UC is ontop
    frame1.ZOrder 0
'start things
    polygon.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
'clean-up and quit
    DeleteObject (hFRgn)
    DeleteObject (hHRgn)
    Set Form1 = Nothing
    End

End Sub

Public Sub swappictures()
'Seperating the first two sequencial pictures (hDC's)
'seems to work in preventing flicker, but I don't
'know why...LOL  ;)
    
        'send it to the back of the z-order
            MaskPic(1).ZOrder 1
            DoEvents
        'swap the pics
            MaskPic(0).Picture = SourcePic(1).Picture
            MaskPic(1).Picture = SourcePic(2).Picture

            Begin = True

End Sub
'FYI; You are not restricted in keeping the Form square
'or rectangular, you can make it an irregular shaped
'object as well. Inspired by Ubisoft.corp, any "Myst"
'fans out there?  ;)


