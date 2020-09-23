VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0070DF70&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3105
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   180
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   420
      Picture         =   "Form1.frx":1F83
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2001 Spyder Net"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001FB04F&
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   2460
      Width           =   2595
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CombineRgn Example"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001FB04F&
      Height          =   1215
      Left            =   960
      TabIndex        =   0
      Top             =   1140
      Width           =   3675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-+-+-+-+-+-+-+-+-+-+-
'CombineRgn Example
'-+-+-+-+-+-+-+-+-+-+-

'Code by Nick Ridley of
'Spyder Net
'http://www.spyderhackers.co.uk

'Copyright © 2001 All Spyder Net Rights Reserved

'Written after I worked this out on 18/11/2001 1800 hrs GMT

'This is the 1st example on PSC (that i can find anyways)
'that can tell u how to use the API call 'CombineRgn'
'this also includes code to make the form move able

'All modules, forms and code written by N Ridley
'Modules taken from a shell project I am working on

'Variables form form movement
Private X1 As Long, Y1 As Long

Private Sub Form_Load()
'variable to hold main region
Dim hHRgn As Long
'variable to hold regions to be added
Dim hFRgn As Long
'Create Regions and add them to the main region
hHRgn = CreateRoundRectRgn(0, 0, 50, 50, 10, 10)
hFRgn = CreateEllipticRgn(0, 0, 100, 100)
'Combine regions (RGN_OR combines all parts of both regions
CombineRgn hHRgn, hFRgn, hHRgn, RGN_OR
hFRgn = CreateRoundRectRgn(50, 50, 200, 200, 50, 50)
CombineRgn hHRgn, hFRgn, hHRgn, RGN_OR
hFRgn = CreateRoundRectRgn(60, 60, 300, 190, 50, 70)
CombineRgn hHRgn, hFRgn, hHRgn, RGN_OR
'apply region to form
SetWindowRgn Me.hwnd, hHRgn, True 'Comment this line if you want to make the hole

'We could have used different flags
'uncomment these lines to create a hole were the form move image is
'(also make the setwindowrgn... line a comment)

'hFRgn = CreateEllipticRgn(20, 20, 80, 80)
'CombineRgn hHRgn, hFRgn, hHRgn, RGN_XOR
'SetWindowRgn Me.hwnd, hHRgn, True

'the RGN_XOR flag willkeep all parts of the 1st rgn that
'are not overlapped by the second
End Sub

'Code to move form:

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Tell timer wheremouse is in relation to image1
X1 = Image1.Left + x: Y1 = Image1.Top + y
'turn on moving the form
Timer2.Enabled = True
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Stop form movement
Timer2.Enabled = False
End Sub

Private Sub Timer2_Timer()
'Variables to hold left + top values
Dim x As Long, y As Long

'work out (usiong modMouse) were mouse is and were form should be put
x = (GetX * 15) - X1
y = (GetY * 15) - Y1

'move form
Me.Left = x
Me.Top = y

End Sub
