VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "captura nomina a costos"
   ClientHeight    =   5265
   ClientLeft      =   2100
   ClientTop       =   2415
   ClientWidth     =   6570
   Icon            =   "NOMCAP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   6570
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5106
      _Version        =   327680
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Menu arch 
      Caption         =   "&Archivo"
      Index           =   0
      Begin VB.Menu archdir 
         Caption         =   "&Cambio directorio"
         Index           =   2
      End
      Begin VB.Menu archsal 
         Caption         =   "&Salida"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub archdir_Click(Index As Integer)
      CommonDialog1.Action = 1
      Label1.Caption = CommonDialog1.filename
      Label2.Caption = Dir(CommonDialog1.filename, vbDirectory)
      nomdir = Left$(CommonDialog1.filename, Len(CommonDialog1.filename) - (Len(CommonDialog1.FileTitle) + 1))
      ChDir nomdir
      Label1.Caption = nomdir
      Open Label2.Caption For Input As 1
  Do Until EOF(1)
     Input #1, mes, qna$, obra, empl, impo#
     MSFlexGrid1.AddItem mes & Chr$(9) & qna$ & Chr$(9) & obra & Chr$(9) & empl & Chr$(9) & Format(impo#, z1$)
  Loop
     Close 1
     Open "\contr98\nominita\personal.nom" For Random As 2 Len = Len(personal)
     cm = LOF(2) / Len(personal)
     For r = 2 To MSFlexGrid1.Rows - 1
              Get 2, MSFlexGrid1.TextMatrix(r, 3), personal
              MSFlexGrid1.TextMatrix(r, 5) = RTrim$(personal.nom) + " " + RTrim$(personal.ape1)
     Next r
     Close 2
End Sub

Private Sub archsal_Click(Index As Integer)
Close: End
End Sub

Private Sub Form_Load()
  MSFlexGrid1.Cols = 8
  MSFlexGrid1.ColWidth(5) = 2400
  z1$ = "###,##0.00"
End Sub
