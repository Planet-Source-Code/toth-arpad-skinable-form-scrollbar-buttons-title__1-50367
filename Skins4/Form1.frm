VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider Slider2 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      Max             =   1000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   135
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   120
      Left            =   360
      Picture         =   "Form1.frx":0000
      Top             =   960
      Width           =   420
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000F&
      FillColor       =   &H8000000F&
      Height          =   135
      Left            =   360
      Top             =   960
      Width           =   3735
   End
   Begin VB.Image Img_exit_1 
      Height          =   495
      Left            =   3480
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label_Title_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Skinable Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Pic_Close_1 
      Height          =   255
      Left            =   4200
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Pic_Title_1 
      Height          =   375
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Toth Arpad
'http://www.totharpad.3x.ro
'arpi@arpisoft.net
'totharpad@as.ro


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Dim xm As Integer
Dim Ym As Integer
Dim verm As Boolean
Dim dis As Integer
Dim longs As Integer

Public Function MoveForm(TheForm As Form)
    Dim ret
    ReleaseCapture
    SendMessage TheForm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Function

Private Sub flood(nr As Integer)
ScaleMode = 1
lScaleHeight = ScaleHeight
lScaleWidth = ScaleWidth
DrawStyle = 5
FillStyle = vbFSSolid
Select Case nr
Case 1
For lY = 0 To lScaleHeight
     FillColor = RGB(255 - (lY * 255) \ lScaleHeight, 0, 0)
     Line (-1, lY - 1)-(lScaleWidth, lY + 1), , B
Next lY
Case 2
For lY = 0 To lScaleHeight
     FillColor = RGB(0, 255 - (lY * 255) \ lScaleHeight, 0)
     Line (-1, lY - 1)-(lScaleWidth, lY + 1), , B
Next lY
Case 3
For lY = 0 To lScaleHeight
     FillColor = RGB(0, 0, 255 - (lY * 255) \ lScaleHeight)
     Line (-1, lY - 1)-(lScaleWidth, lY + 1), , B
Next lY
Case 4
For lY = 0 To lScaleHeight
     FillColor = RGB(255 - (lY * 255) \ lScaleHeight, 255 - (lY * 255) \ lScaleHeight, 0)
     Line (-1, lY - 1)-(lScaleWidth, lY + 1), , B
Next lY
Case 5
For lY = 0 To lScaleHeight
     FillColor = RGB(255 - (lY * 255) \ lScaleHeight, 0, 255 - (lY * 255) \ lScaleHeight)
     Line (-1, lY - 1)-(lScaleWidth, lY + 1), , B
Next lY
Case 6
For lY = 0 To lScaleHeight
     FillColor = RGB(0, 255 - (lY * 255) \ lScaleHeight, 255 - (lY * 255) \ lScaleHeight)
     Line (-1, lY - 1)-(lScaleWidth, lY + 1), , B
Next lY


End Select
End Sub

Private Sub Form_paint()
flood (6)
End Sub

Private Sub Form_Load()
Dim i
Image2.Left = Label2.Left
longs = Slider2.Width
    
    flood (6)


Pic_Close_1.Picture = LoadPicture(App.Path & "\images\skin1\but_close_1.gif")
Img_exit_1.Picture = LoadPicture(App.Path & "\images\skin1\But_Exit_1.gif")

Pic_Title_1.Left = 0
Pic_Title_1.Width = Form1.Width
Pic_Title_1.Picture = LoadPicture(App.Path & "\images\skin1\title_1.bmp")
Label_Title_1.Caption = "Skinable Form"
Label_Title_1.Left = (Form1.ScaleWidth / 2) - (Label_Title_1.Width / 2)
Label_Title_1.Top = (Pic_Title_1.Height / 2) - (Label_Title_1.Height / 2)
         
End Sub




Private Sub Img_exit_1_Click()
End
End Sub

Private Sub Img_exit_1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Img_exit_1.Picture = LoadPicture(App.Path & "\images\skin1\But_Exit_2.gif")
End Sub

Private Sub Img_exit_1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Img_exit_1.Picture = LoadPicture(App.Path & "\images\skin1\But_Exit_1.gif")
End Sub

Private Sub Label_Title_1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 MoveForm Me
End Sub

Private Sub Pic_Close_1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Pic_Close_1.Picture = LoadPicture(App.Path & "\images\skin1\but_close_2.gif")
End Sub

Private Sub Pic_Title_1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm Me
End Sub

Private Sub Pic_close_1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Pic_Close_1.Picture = LoadPicture(App.Path & "\images\skin1\but_close_1.gif")
End Sub




Private Sub label2_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (X > Image2.Left) And X < (Image2.Left + Image2.Width) Then
dis = X - Image2.Left
 Else
 dis = (Image2.Width / 2)
End If
Image2.Left = X - dis
Slider2.Value = (Image2.Left * Slider2.Max) / (longs - Image2.Width)
End Sub
Private Sub label2_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Image2.Left = X - dis
Slider2.Value = (Image2.Left * Slider2.Max) / (longs - Image2.Width)
End If
End Sub


Private Sub Slider2_Change()
Image2.Left = ((Slider2.Value / Slider2.Max) * (longs - Image2.Width)) + Label2.Left
Label1.Caption = Slider2.Value
End Sub


'I use a transparent label and when you push the mouse button on label the image that is under
'the transparent label is moving on X coordonate of maouse on label. The Y coordonates is constant.


