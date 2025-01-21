VERSION 5.00
Begin VB.Form SUBCAMB 
   Caption         =   "Cambio Subdirectorio"
   ClientHeight    =   3360
   ClientLeft      =   2808
   ClientTop       =   2112
   ClientWidth     =   4680
   Icon            =   "SUBCAM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4680
   Begin VB.CommandButton Command2 
      Caption         =   "Finalizar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      MaskColor       =   &H000000FF&
      TabIndex        =   6
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   1416
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Dispositivo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Archivos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Directorios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   360
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Menu salcam 
      Caption         =   ""
      Index           =   0
   End
End
Attribute VB_Name = "SUBCAMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim salida As Integer

Private Sub Command1_Click()
    salida = 1
    Close
    Open "C:\GconTA\sccontr.sc" For Random As 3 Len = Len(SCont)
        
        SCont.guarda = Trim(Dir1.Path)
        If Left(Trim(SCont.guarda), 1) <> "\" Then
             SCont.guarda = Trim(SCont.guarda) + "\"
        End If
        ChDrive Mid(Dir1.Path, 1, 2)
        ChDir Dir1.Path
    If tipo = 1 Then
            Put 3, 2, SCont
            tipo = 0
            Else
            Put 3, 1, SCont
    End If
    Close 3
    Unload SUBCAMB
End Sub

Private Sub Command2_Click()
   salida = 1
   terminar = 1
   Unload SUBCAMB
   Close: End
End Sub

Private Sub Dir1_Change()
    salida = 1
    Rem File1.Path = RTrim(Dir1.Path)
    
    Label4.Caption = (Dir1.Path)
    If (Mid(Label4.Caption, 1, 1)) <> UCase("c") Then ChDir (UCase(Mid(Label4.Caption, 1, 2)))
    ChDir Label4.Caption
    Close
    terminar = 0
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
       If Dir1.Path <> Dir1.List(Dir1.ListIndex) Then
              Dir1.Path = (Dir1.List(Dir1.ListIndex))
              
       End If
       Label4.Caption = RTrim(Dir1.Path)
      End If
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub
Private Sub Form_Initialize()
    salida = 0
    Dir1.Path = RTrim(Dir1.Path)
    Rem Label4.Caption = RTrim(LTrim(Dir1.Path)) + " " + LTrim(RTrim(File1.FileName))
    Label4.Caption = (Dir1.Path)
    If salida = 1 Then
          Rem Form1.SetFocus
          Else
          Dir1_Change
    End If
End Sub

Private Sub Form_Load()
     Rem Label4.Caption = LTrim(RTrim(Dir1.Path)) + " " + LTrim(RTrim(File1.FileName))
     Label4.Caption = Dir1.Path
     
     
End Sub

