VERSION 5.00
Begin VB.Form Information 
   Caption         =   "Información del Programa"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6225
   LinkTopic       =   "Form10"
   ScaleHeight     =   4230
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Propiedades"
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.Label Label14 
         Caption         =   "win 7,8,10."
         Height          =   200
         Left            =   3000
         TabIndex        =   14
         Top             =   2940
         Width           =   1500
      End
      Begin VB.Label Label13 
         Caption         =   "x64 / x86 bits"
         Height          =   200
         Left            =   3000
         TabIndex        =   13
         Top             =   2535
         Width           =   1500
      End
      Begin VB.Label Label12 
         Caption         =   "1.05 MB"
         Height          =   200
         Left            =   3000
         TabIndex        =   12
         Top             =   2145
         Width           =   1500
      End
      Begin VB.Label Label11 
         Caption         =   "SACMAG DE MËXICO S.A. de C.V."
         Height          =   195
         Left            =   3000
         TabIndex        =   11
         Top             =   1740
         Width           =   2580
      End
      Begin VB.Label Label10 
         Caption         =   "04 de diciembre de 2023"
         Height          =   195
         Left            =   3000
         TabIndex        =   10
         Top             =   1320
         Width           =   1860
      End
      Begin VB.Label Label9 
         Caption         =   "23.7.1"
         Height          =   200
         Left            =   3000
         TabIndex        =   9
         Top             =   945
         Width           =   1500
      End
      Begin VB.Label Label8 
         Caption         =   "Nomina"
         Height          =   200
         Left            =   3000
         TabIndex        =   8
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label Label7 
         Caption         =   "Compatible:"
         Height          =   195
         Left            =   105
         TabIndex        =   7
         Top             =   2940
         Width           =   1710
      End
      Begin VB.Label Label6 
         Caption         =   "Sistema basado en:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   2535
         Width           =   2445
      End
      Begin VB.Label Label5 
         Caption         =   "Tamaño total:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   2145
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Editor:"
         Height          =   195
         Left            =   105
         TabIndex        =   4
         Top             =   1740
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "Última modificación:"
         Height          =   195
         Left            =   105
         TabIndex        =   3
         Top             =   1335
         Width           =   1950
      End
      Begin VB.Label Label2 
         Caption         =   "Versión:"
         Height          =   200
         Left            =   105
         TabIndex        =   2
         Top             =   945
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   200
         Left            =   105
         TabIndex        =   1
         Top             =   540
         Width           =   1005
      End
   End
End
Attribute VB_Name = "Information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Label1.FontBold = True
Label2.FontBold = True
Label3.FontBold = True
Label4.FontBold = True
Label5.FontBold = True
Label6.FontBold = True
Label7.FontBold = True
End Sub

