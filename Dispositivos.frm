VERSION 5.00
Begin VB.Form Dispositivos 
   BackColor       =   &H00808080&
   Caption         =   "Seleccione el dispostivo"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4515
   ControlBox      =   0   'False
   Icon            =   "Dispositivos.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2370
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ComboBox Soundcard 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   720
      Width           =   2055
   End
   Begin VB.ComboBox Direct 
      Height          =   315
      ItemData        =   "Dispositivos.frx":000C
      Left            =   120
      List            =   "Dispositivos.frx":0019
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Tarjeta de sonido:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Emisión de sonidos."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "Dispositivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
' driver >= FSOUND_GetNumDrivers());

'    FSOUND_SetDriver(driver);                   // Select sound card (0 = default)
Select Case Direct.ListIndex
        Case 0:
            FSOUND_SetOutput (FSOUND_OUTPUT_DSOUND)
            Tipo_Salida = 0
        Case 1:
            FSOUND_SetOutput (FSOUND_OUTPUT_WINMM)
            Tipo_Salida = 1
        Case 2:
            FSOUND_SetOutput (FSOUND_OUTPUT_A3D)
            Tipo_Salida = 2
    End Select
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim valor As Integer
Dim Index As Integer
'Soundcard.Clear
Direct.ListIndex = 0
 '   For Index = 0 To FSOUND_GetNumDrivers
  '      Soundcard.AddItem (FSOUND_GetDriverName(Index))
 '   Next Index
'    Soundcard.ListIndex = 0
End Sub

