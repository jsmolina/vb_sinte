VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Virtual Sampler"
   ClientHeight    =   4620
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7425
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   5160
      TabIndex        =   36
      Text            =   "Notas Ritmo2"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   5160
      TabIndex        =   35
      Text            =   "Notas Ritmo1"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5160
      TabIndex        =   34
      Text            =   "Notas Acompañamiento"
      Top             =   720
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5160
      TabIndex        =   33
      Text            =   "Notas parte melodica"
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "Sample.wav"
      Top             =   225
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Elemento a Editar"
      Height          =   4200
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
      Begin VB.OptionButton Option4 
         Caption         =   "Ritmo 2"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2640
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Ritmo 1"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Acompañamiento"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Melodía"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label29 
         Caption         =   "Activo"
         Height          =   210
         Left            =   705
         TabIndex        =   38
         Top             =   3540
         Width           =   525
      End
      Begin VB.Label Activo 
         BackColor       =   &H000000FF&
         Height          =   105
         Left            =   315
         TabIndex        =   37
         Top             =   3585
         Width           =   315
      End
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   4665
      Picture         =   "Main.frx":030A
      ToolTipText     =   "Cambiar Sample"
      Top             =   210
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   4680
      Picture         =   "Main.frx":097C
      Top             =   225
      Width           =   315
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   22
      Left            =   6840
      TabIndex        =   29
      Top             =   3240
      Width           =   225
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   20
      Left            =   6480
      TabIndex        =   28
      Top             =   3240
      Width           =   225
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   18
      Left            =   6120
      TabIndex        =   27
      Top             =   3240
      Width           =   225
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   15
      Left            =   5400
      TabIndex        =   26
      Top             =   3240
      Width           =   225
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   13
      Left            =   5040
      TabIndex        =   25
      Top             =   3240
      Width           =   225
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   10
      Left            =   4320
      TabIndex        =   24
      Top             =   3240
      Width           =   225
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   8
      Left            =   3960
      TabIndex        =   23
      Top             =   3240
      Width           =   225
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   6
      Left            =   3600
      TabIndex        =   22
      Top             =   3240
      Width           =   225
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   3
      Left            =   2910
      TabIndex        =   21
      Top             =   3240
      Width           =   225
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   23
      Left            =   6960
      TabIndex        =   20
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   21
      Left            =   6600
      TabIndex        =   19
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   19
      Left            =   6240
      TabIndex        =   18
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   17
      Left            =   5880
      TabIndex        =   17
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   16
      Left            =   5520
      TabIndex        =   16
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   14
      Left            =   5160
      TabIndex        =   15
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   12
      Left            =   4800
      TabIndex        =   14
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   11
      Left            =   4440
      TabIndex        =   13
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   9
      Left            =   4080
      TabIndex        =   12
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   7
      Left            =   3720
      TabIndex        =   11
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   5
      Left            =   3360
      TabIndex        =   10
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   4
      Left            =   3000
      TabIndex        =   9
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   1
      Left            =   2550
      TabIndex        =   6
      Top             =   3240
      Width           =   225
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   0
      Left            =   2280
      TabIndex        =   5
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   2
      Left            =   2640
      TabIndex        =   8
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2295
      TabIndex        =   7
      Top             =   3240
      Width           =   5040
   End
   Begin VB.Label Label26 
      Caption         =   "Do     Re   Mi    Fa   Sol    La    Si     Do  Re      Mi    Fa  Sol    La    Si"
      Height          =   255
      Left            =   2280
      TabIndex        =   30
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Label Label27 
      Caption         =   "Do# Re#        Fa#  Sol#  La#        Do#  Re#        Fa#  Sol#  La#"
      Height          =   255
      Left            =   2520
      TabIndex        =   31
      Top             =   2790
      Width           =   4815
   End
   Begin VB.Menu File 
      Caption         =   "&Archivo"
      Begin VB.Menu New 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu Open 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu Save 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu SaveAs 
         Caption         =   "Guardar &Como"
      End
      Begin VB.Menu ticka 
         Caption         =   "-"
      End
      Begin VB.Menu PrintA 
         Caption         =   "Imprimir Notas"
      End
      Begin VB.Menu tickb 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Sa&lir"
      End
   End
   Begin VB.Menu Exec 
      Caption         =   "Ejecu&tar"
      Begin VB.Menu Playa 
         Caption         =   "&Reproducir..."
         Begin VB.Menu FullPlay 
            Caption         =   "Canción &Completa"
            Shortcut        =   {F5}
         End
         Begin VB.Menu Selected 
            Caption         =   "Parte &Seleccionada"
            Shortcut        =   {F6}
         End
      End
      Begin VB.Menu Stopa 
         Caption         =   "&Dejar de Reproducir"
         Shortcut        =   {F8}
      End
      Begin VB.Menu av 
         Caption         =   "-"
      End
      Begin VB.Menu TakeSound 
         Caption         =   "&Tomar Disp. de Sonido"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Freeze 
         Caption         =   "&Liberar Disp. de Sonido"
         Shortcut        =   {F3}
      End
      Begin VB.Menu SelectDisp 
         Caption         =   "&Seleccionar Dipsositivo"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu Helpa 
      Caption         =   "Ay&uda"
      Begin VB.Menu Content 
         Caption         =   "&Contenido"
      End
      Begin VB.Menu AboutThis 
         Caption         =   "&Acerca De..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Initialize_FMOD()
    If FSOUND_Init(44100, 32, 0) = 0 Then
'    Error
    MsgBox "Error inicializando FMOD.DLL  " & vbCrLf & FSOUND_GetErrorString(FSOUND_GetError)
    End
End If
Select Case Tipo_Salida
    Case 0:
        FSOUND_SetOutput (FSOUND_OUTPUT_DSOUND)
    Case 1:
        FSOUND_SetOutput (FSOUND_OUTPUT_WINMM)
    Case 2:
        FSOUND_SetOutput (FSOUND_OUTPUT_A3D)
End Select
lp_cutoff = 5000
lp_reson = 1

FMOD_LOADED = True
Call SwitchLabel
End Sub

Private Sub SwitchLabel()
    If (Activo.BackColor = RGB(255, 0, 0)) Then
        Activo.BackColor = RGB(0, 255, 0)
    Else
        Activo.BackColor = RGB(255, 0, 0)
    End If
End Sub


Public Sub Unload_FMOD()
If channel1 <> 0 Then
    FSOUND_Stream_Stop FILEID
    channel1 = 0
End If

If FILEID <> 0 Then
    FSOUND_Stream_Close FILEID
    FILEID = 0
End If


Call FSOUND_Close
FMOD_LOADED = False
Call SwitchLabel
End Sub

Private Sub AboutThis_Click()
    About.Show 0, Me
End Sub

Private Sub Exit_Click()
Call Unload_FMOD
End
End Sub


Private Sub Form_Load()
    FMOD_LOADED = False
    Tipo_Salida = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseUp = True
Image2.Visible = False
Image1.Visible = True
End Sub


Private Sub Form_Resize()
On Error GoTo saltado
    Form1.Width = 7545
    Form1.Height = 5310
saltado:
End Sub

Private Sub Form_Terminate()
Call Unload_FMOD
End Sub

Private Sub Form_Unload(Cancel As Integer)
If FMOD_LOADED = True Then
    Call Unload_FMOD
End If
End
End Sub

Private Sub Freeze_Click()
If FMOD_LOADED = True Then
    Call Unload_FMOD
End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MouseUp = True Then
        Image1.Visible = False
        Image2.Visible = True
        MouseUp = False
    End If
End Sub

Private Sub Image2_Click()
    Dim Archivo As String
    sample1 = FSOUND_Sample_Load(FSOUND_FREE, Archivo, FSOUND_HW3D, 0)
    sample2 = FSOUND_Sample_Load(FSOUND_FREE, Archivo, FSOUND_HW3D, 0)
    sample3 = FSOUND_Sample_Load(FSOUND_FREE, Archivo, FSOUND_HW3D, 0)
    sample4 = FSOUND_Sample_Load(FSOUND_FREE, Archivo, FSOUND_HW3D, 0)
End Sub

Private Sub Label3_Click()

End Sub


Private Sub SelectDisp_Click()
If FMOD_LOADED = False Then
    valor = MsgBox("Fmod no iniciado! ¿Desea iniciar FMOD ahora?", vbYesNo, "Dispositivos")
    If valor = 6 Then
       Call Form1.Initialize_FMOD
    End If
End If
If FMOD_LOADED = True Then
    Call Dispositivos.Show(1, Me)
End If
End Sub

Private Sub TakeSound_Click()
If FMOD_LOADED = False Then
    Call Initialize_FMOD
End If
End Sub

