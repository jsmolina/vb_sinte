VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   ControlBox      =   0   'False
   Icon            =   "AboutMe.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "AboutMe.frx":000C
   ScaleHeight     =   3555
   ScaleWidth      =   3645
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3240
      Top             =   2880
   End
   Begin VB.TextBox Abouter 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   2535
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   3240
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Picture         =   "AboutMe.frx":613E
      Top             =   0
      Width           =   3750
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CurrentPart As Integer
Public Total As Integer
Public To_Print As String

Private Sub Form_Load()
    CurrentPart = 1
    To_Print = "Diseñado para Arte y Tecnologia" & "      Por Jordi Sesmero y Eric Mora " & "                         Visual Basic 6"
    Total = Len(To_Print)
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Dim MyStr As String
    
    MyStr = String(1, " ")
    MyStr = Mid$(To_Print, CurrentPart, 1)
'    If MyStr = Chr$(13) Then MyStr = Chr$(10)
    Abouter.Text = Abouter.Text + MyStr
    CurrentPart = CurrentPart + 1
    ''14
    If CurrentPart > Total Then Timer1.Enabled = False
End Sub
