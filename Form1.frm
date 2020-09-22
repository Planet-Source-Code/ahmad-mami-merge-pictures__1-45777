VERSION 5.00
Begin VB.Form Frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Merge"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   11205
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Undo"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9840
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2535
      Left            =   5640
      ScaleHeight     =   2475
      ScaleWidth      =   5355
      TabIndex        =   9
      Top             =   120
      Width           =   5415
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   5235
      TabIndex        =   8
      Top             =   120
      Width           =   5295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restore All"
      Height          =   495
      Left            =   9840
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save New Pic"
      Height          =   495
      Left            =   8520
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "PLease Vote For This Code And Feel Free To Leave Any Comments."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   5640
      Width           =   6735
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   1320
      X2              =   7680
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   8880
      X2              =   9000
      Y1              =   4800
      Y2              =   4680
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   9120
      X2              =   9000
      Y1              =   4800
      Y2              =   4680
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   9000
      X2              =   9000
      Y1              =   5040
      Y2              =   4680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select The Size Then Start To Delete The Old Background Form Here"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   4920
      Width           =   7800
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   480
      Top             =   4680
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   480
      X2              =   495
      Y1              =   6000
      Y2              =   6015
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   480
      Top             =   5640
      Width           =   135
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   480
      Top             =   5160
      Width           =   255
   End
End
Attribute VB_Name = "Frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim v As Long

Private Sub Command1_Click()
Call SavePicture(Picture2.Image, "test.jpg")
MsgBox "The Picture Saved in " & App.Path & "\test.jpg", vbInformation, "Saved"
End Sub

Private Sub Command2_Click()
Picture2.Picture = LoadPicture("lobby111.JPG")
End Sub

Private Sub Command3_Click()
Picture2.Picture = LoadPicture("undo.JPG")
Command3.Enabled = False
End Sub

Private Sub Form_Load()
v = 20
Picture1.Picture = LoadPicture("sunset.jpg")
Picture2.Picture = LoadPicture("lobby111.JPG")
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
v = 30
Case 1
v = 20
Case 2
v = 10
Case 3
v = 5
End Select
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SavePicture(Picture2.Image, "undo.jpg")
Command3.Enabled = True
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
BitBlt Picture2.hdc, X / 15, Y / 15, v, v, Picture1.hdc, X / 15, Y / 15, vbSrcCopy
Picture2.Refresh
End If
End Sub

