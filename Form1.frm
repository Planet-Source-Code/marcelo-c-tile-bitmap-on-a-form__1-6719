VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Tile Bitmap Example - Using API"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10275
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Bitmap 2"
      Height          =   345
      Left            =   5475
      TabIndex        =   4
      Top             =   480
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bitmap 1"
      Height          =   345
      Left            =   3765
      TabIndex        =   3
      Top             =   465
      Width           =   1605
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   4845
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   1260
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   2790
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   885
      Picture         =   "Form1.frx":265A
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   2790
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Label Label1 
      Caption         =   "As you can see we have here 2 pictureboxes which at runtime will be visible= false, aka, invisible. (So will this little text)"
      Height          =   615
      Left            =   810
      TabIndex        =   2
      Top             =   2055
      Visible         =   0   'False
      Width           =   3240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WhichPictureWasSelected As Integer

Private Sub Command1_Click()
WhichPictureWasSelected = 1
Form_Resize
End Sub

Private Sub Command2_Click()
WhichPictureWasSelected = 2
Form_Resize

End Sub

Private Sub Form_Activate()

TileBitmap Me, Picture2
WhichPictureWasSelected = 2

End Sub

Private Sub Form_Resize()
'// Thats it, add it to the form resize event
'// and there you go.
'//
'// The code could be as simple as
'// ex.  TileBitmap Me, Picture1
'//
'// but for the demonstration it is a bit more complex.

Select Case WhichPictureWasSelected

Case Is = 1
 TileBitmap Me, Picture1

Case Is = 2
 TileBitmap Me, Picture2

End Select


End Sub

