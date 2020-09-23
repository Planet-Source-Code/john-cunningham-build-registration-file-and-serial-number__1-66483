VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   Caption         =   "Program By:  John P. Cunningham"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   4920
      Top             =   240
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   4695
      TabIndex        =   4
      Top             =   0
      Width           =   4695
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":030A
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Send Email:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   856
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Go to Web Page:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   218
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   " By:  John P. Cunningham"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1230
      TabIndex        =   3
      Top             =   945
      Width           =   3090
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "jpcunningham@cox.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   240
      Left            =   2310
      MouseIcon       =   "frmAbout.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Send Email To John Cunningham"
      Top             =   1920
      Width           =   2400
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "http://members.cox.net/johnpc7/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   2018
      MouseIcon       =   "frmAbout.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Open John Cunningham'sWeb Page"
      Top             =   1440
      Width           =   3315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Build Registration File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   1230
      TabIndex        =   0
      Top             =   360
      Width           =   3090
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub Sleep(Seconds As Double)
   Dim TempTime As Double
   TempTime = Timer
   While Timer - TempTime < Seconds
      DoEvents
      If Timer < TempTime Then
         TempTime = TempTime - 24# * 3600#
      End If
   Wend
End Sub
Private Sub Form_Activate()

TextEffect Me, "", 12, 12, , 128, 0, RGB(&H80, 0, 0)
End Sub

Private Sub Form_Load()
lblEmail = email
lblURL = URL
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'lblURL.ForeColor = &HFFFF&      'yellow
'lblEmail.ForeColor = &H808000  'green
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
        i = Me.Height
    While i >= Picture1.Height
        i = i - 40
        If i > Picture1.Height Then
            Me.Height = i
        Else
            Me.Height = Picture1.Height
        End If
        
        DoEvents
    Wend
    Me.Caption = "Decimal/Binary Converter By: J.Cunningham"
    i = Me.Top
    i = Me.Top
    While i > 0
        Me.Move Me.Left, i, Me.Width, Me.Height
        i = i - 50
        DoEvents
        
    Wend
    i = Me.Left
    While i < Screen.Width
        Me.Move i, 0, Me.Width, Me.Height
        i = i + 105
        DoEvents
       
    Wend
End Sub

Private Sub lblEmail_Click()
sendemail
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 lblEmail.ForeColor = &HFF8080     'hyperlink blue
Label3.ForeColor = &HFF8080
End Sub

Private Sub lblURL_Click()
gotoweb
End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  lblURL.ForeColor = &HFF8080        'hyperlink blue
  Label2.ForeColor = &HFF8080
End Sub

Private Sub Timer1_Timer()
  
    lblEmail.ForeColor = &H80FF80   '&H808000
    Label3.ForeColor = &H80FF80   '&H808000
    lblURL.ForeColor = &HFFFF&
    Label2.ForeColor = &HFFFF&
End Sub
