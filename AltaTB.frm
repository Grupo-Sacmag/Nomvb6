VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form9 
   Caption         =   "Alta tarjetas Banamex"
   ClientHeight    =   5970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form9"
   ScaleHeight     =   5970
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   5400
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Ex 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8493
      _Version        =   393216
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Rem Open "bnxcla.dno" For Random As 12 Len = Len(Clbnx)
    Open "quin.txt" For Output As #11
    Close 13: Open "Empcomp.dno" For Random As 13 Len = Len(Dat_ide)
    Get 13, 1, Dat_ide
    
    Slinimpte = 0: Sabono = 0: Siguiente = 0
    For r1 = 1 To limite
      li = r1
      rgtro = ConNom1.TextMatrix(li, 0)
      Get #2, rgtro, personal
      Get #8, rgtro, maestro
      Get #12, rgtro, Clbnx
   Ex.TextMatrix(li, 1) = RTrim$(personal.ape1) + " " + RTrim$(personal.ape2) + " " + RTrim$(personal.nom)
   Next
End Sub
