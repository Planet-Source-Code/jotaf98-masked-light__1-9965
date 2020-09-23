VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "''Light'' - Part II: Mask - Made by Jotaf98"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "Light.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   6360
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   26
      Top             =   4320
      Width           =   975
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   0
         Picture         =   "Light.frx":08CA
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   27
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6840
      TabIndex        =   22
      Text            =   "5"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   6840
      TabIndex        =   21
      Text            =   "15"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   6840
      TabIndex        =   20
      Text            =   "40"
      Top             =   3240
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Please choose a mask image:"
      Filter          =   "Bitmap (*.bmp)|*.bmp|JPG/JPEG Images (*.jpg;*.jpeg)|*.jpg;*.jpeg"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Randomly get a pre-defined color"
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   4920
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6840
      TabIndex        =   11
      Text            =   "7"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6840
      TabIndex        =   2
      Text            =   "230"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6840
      TabIndex        =   1
      Text            =   "200"
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4140
      Left            =   120
      ScaleHeight     =   276
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   415
      TabIndex        =   0
      Top             =   120
      Width           =   6225
   End
   Begin VB.Label Label11 
      Caption         =   "Red Brightness:"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   6480
      TabIndex        =   25
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Green Brightness:"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   6480
      TabIndex        =   24
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Blue Brightness:"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   6480
      TabIndex        =   23
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "Path to the mask image:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Mask:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   5760
      TabIndex        =   16
      Top             =   4680
      Width           =   495
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   7320
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label15 
      Caption         =   ";)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   6930
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000016&
      X1              =   3240
      X2              =   3240
      Y1              =   5400
      Y2              =   7320
   End
   Begin VB.Label Label14 
      Caption         =   $"Light.frx":390C
      Height          =   855
      Left            =   3480
      TabIndex        =   13
      Top             =   6480
      Width           =   3855
   End
   Begin VB.Label Label10 
      Caption         =   "Radius:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   3720
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   3225
      X2              =   3225
      Y1              =   5400
      Y2              =   7320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000016&
      X1              =   7320
      X2              =   120
      Y1              =   5415
      Y2              =   5415
   End
   Begin VB.Label Label7 
      Caption         =   "and vote me!"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "http://www.planet-source-code.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6720
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "If you liked this effect, please go to"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6480
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "X:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6480
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Y:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   $"Light.frx":39E7
      Height          =   855
      Left            =   3480
      TabIndex        =   5
      Top             =   5520
      Width           =   3855
   End
   Begin VB.Label Label8 
      Caption         =   "Made by Jotaf98 (João F. S. Henriques)"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Label Label9 
      Caption         =   "E-mail me at: jotaf98@hotmail.com"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CoolColors(9, 2) As Byte 'Where the "cool colors" will be stored
Dim RndCoolColor As Byte '"Cool color" got at random

Private Sub RedrawPicture() 'Redraws the picture
    On Error GoTo ErrHandler
    Picture1.Picture = LoadPicture(App.Path & "\Geyser.jpg")
    Exit Sub
    
ErrHandler:
    MsgBox "There was an error displaying the image.", vbExclamation + vbOKOnly
    Picture1.Cls
End Sub

Private Sub RedrawMask() 'Redraws the mask
    On Error GoTo ErrHandler
    Picture2.Picture = LoadPicture(Text7.Text)
    Exit Sub
    
ErrHandler:
    MsgBox "There was an error displaying the mask image.", vbExclamation + vbOKOnly
    Picture2.Cls
End Sub

Private Sub Command1_Click() 'Pick a random predefined color
    RndCoolColor = Rnd * UBound(CoolColors)
    
    MsgBox "Try this one:" & Chr(13) & Chr(13) & _
      "  Red: " & CoolColors(RndCoolColor, 0) & Chr(13) & _
      "  Green: " & CoolColors(RndCoolColor, 1) & Chr(13) & _
      "  Blue: " & CoolColors(RndCoolColor, 2)
End Sub

Private Sub Command2_Click() 'Choose a path to a mask image
    CommonDialog1.FileName = App.Path & "\Mask_1.bmp"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName <> "" Then
        Text7.Text = CommonDialog1.FileName
    Else
        Text7.Text = App.Path & "\Mask_1.bmp"
    End If
    
    RedrawMask
End Sub

Private Sub Form_Load()
    Randomize Timer
    
    Form1.MousePointer = vbHourglass 'Change pointer to hourglass
    Form1.Caption = "''Light'' - Part II: Mask - Initializing colors databank..." 'Change caption
    
    'Initialize the "cool colors" (they're 10)
    InitCoolColor 0, 25, 15, 50
    InitCoolColor 1, 50, 20, 10
    InitCoolColor 2, 15, 40, 5
    InitCoolColor 3, 10, 20, 60
    InitCoolColor 4, 5, 15, 40
    InitCoolColor 5, 60, 15, 35
    InitCoolColor 6, 30, 20, 15
    InitCoolColor 7, 15, 0, 60
    InitCoolColor 8, 60, 15, 0
    InitCoolColor 9, 7, 20, 0
    
    'Redraw the main image and mask
    Form1.Caption = "''Light'' - Part II: Mask - Redrawing image..." 'Change caption
    
    Text7.Text = App.Path & "\Mask_1.bmp"
    RedrawMask
    RedrawPicture
    
    Form1.Caption = "''Light'' - Part II: Mask - Made by Jotaf98" 'Change caption to default
    Form1.MousePointer = vbDefault 'Change pointer to default
End Sub

Private Sub InitCoolColor(Index, RedB, GreenB, BlueB)
    'Initializes a "cool color"
    CoolColors(Index, 0) = RedB
    CoolColors(Index, 1) = GreenB
    CoolColors(Index, 2) = BlueB
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrHandler
    
    Temp = MsgBox("If you liked this effect, please go to Planet Source Code to vote me ; )" & Chr(13) & Chr(13) & "Visit ''http://www.planet-source-code.com'' now?", vbQuestion + vbYesNo)
    
    If Temp = vbYes Then
        'This code snippet opens an .url file.
        'Please give me some credit if you use it, even
        'if it's just a comment in your program's code
        'saying "OpenURL snippet by Jotaf98" ;)
        Shell "rundll32.exe shdocvw.dll,OpenURL " & App.Path & "\Planet Source Code.url"
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Sorry, ''Planet Source Code.url'' was not found.", vbExclamation + vbOKOnly
End Sub

Private Sub Label4_Click()
    On Error GoTo ErrHandler
    
    Temp = MsgBox("Visit ''http://www.planet-source-code.com'' now?", vbQuestion + vbYesNo)
    
    If Temp = vbYes Then
        'This code snippet opens an .url file.
        'Please give me some credit if you use it, even
        'if it's just a comment in your program's code
        'saying "OpenURL snippet by Jotaf98" ;)
        Shell "rundll32.exe shdocvw.dll,OpenURL " & App.Path & "\Planet Source Code.url"
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Sorry, ''Planet Source Code.url'' was not found.", vbExclamation + vbOKOnly
End Sub

Private Sub Picture1_Click()
    Form1.MousePointer = vbHourglass 'Change pointer to hourglass
    Form1.Caption = "''Light'' - Part II: Mask - Redrawing image..." 'Change caption
    RedrawPicture
    Form1.Caption = "''Light'' - Part II: Mask - Drawing light..." 'Change caption
    
    'Draw the light
    DrawMaskedLight Picture1.hdc, Picture2.hdc, _
      Picture2.ScaleWidth, Picture2.ScaleHeight, _
      Int(Text1.Text), Int(Text2.Text), _
      Int(Text4.Text), Int(Text5.Text), _
      Int(Text6.Text), Int(Text3.Text)
    
    Form1.Caption = "''Light'' - Part II: Mask - Made by Jotaf98" 'Change caption to default
    Form1.MousePointer = vbDefault 'Change pointer to default
End Sub
