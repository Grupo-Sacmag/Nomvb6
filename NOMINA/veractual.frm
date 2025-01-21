VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form Form10 
   Caption         =   "Acumulado de sueldos"
   ClientHeight    =   3090
   ClientLeft      =   3285
   ClientTop       =   3660
   ClientWidth     =   6015
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6015
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   2160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Incorporando Archivo"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Archivos Actualizados"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Menu AcuVerl 
      Caption         =   "&Actualizar"
      Index           =   0
      Begin VB.Menu veract 
         Caption         =   "&Actualizar"
         Index           =   2
      End
      Begin VB.Menu AcuVerSal 
         Caption         =   "&Salida"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub improvisa()
        
      Select Case inter_clave
          Case 0
             no_a = 0
             
             For i = 0 To List1.ListCount - 1
               If Label3.Caption = UCase(RTrim(List1.List(i))) Then
                    
                    i = List1.ListCount - 1
                    no_a = 1
               End If
             Next i
             If no_a = 0 Then
                Acumu_lar
                List1.AddItem Form9.CommonDialog1.FileTitle
                Else
                MsgBox "El archivo ya esta actualizado ", vbCritical + vbDefaultButton1, "Actualizacion de Archivos"
             End If
             
          Case 1
             Label3.Caption = "Verificando Archivos Solamente"
          Case 2
             If dm_acu > 0 Then Close 11: Kill "Contr.acu"
             Open "Contr.Acu" For Random As 11 Len = Len(contr_Acu)
             dm_acu = LOF(11) / Len(contr_Acu)
             Label3.Caption = "Actualizando todas las nominas"
             Close 2
             Open "Acumul.nom" For Random As 2 Len = Len(acumulado)
             nq_2 = LOF(2) / Len(acumulado)
             blanc0
             total
             todito
       End Select

End Sub
Sub total()
    nominilla(1) = "ENE11999.NOM": nominilla(2) = "ENE21999.NOM"
    nominilla(3) = "FEB11999.NOM": nominilla(4) = "FEB21999.NOM"
    nominilla(5) = "MAR11999.NOM": nominilla(6) = "MAR21999.NOM"
    nominilla(7) = "ABR11999.NOM": nominilla(8) = "ABR21999.NOM"
    nominilla(9) = "MAY11999.NOM": nominilla(10) = "MAY21999.NOM"
    nominilla(11) = "JUN11999.NOM": nominilla(12) = "JUN21999.NOM"
    nominilla(13) = "JUL11999.NOM": nominilla(14) = "JUL21999.NOM"
    nominilla(15) = "AGO11999.NOM": nominilla(16) = "AGO21999.NOM"
    nominilla(17) = "SEP11999.NOM": nominilla(18) = "SEP21999.NOM"
    nominilla(19) = "OCT11999.NOM": nominilla(20) = "OCT21999.NOM"
    nominilla(21) = "NOV11999.NOM": nominilla(22) = "NOV21999.NOM"
    nominilla(23) = "DIC11999.NOM": nominilla(24) = "DIC21999.NOM"
End Sub
Sub todito()
   List1.Clear
   For t = 1 To 24
      AR1 = nominilla(t)
      Close 1
      Open nominilla(t) For Random As 1 Len = Len(nomina)
      nq_1 = LOF(1) / Len(nomina)
      If nq_1 > 0 Then
          Acumu_lar
          Label3.Caption = nominilla(t)
          List1.AddItem nominilla(t), t - 1
          dm_acu = dm_acu + 1
          contr_Acu.no_arch = nominilla(t)
          contr_Acu.largo = LOF(1) / Len(nomina)
          Put 11, dm_acu, contr_Acu
         Else
         Close 1
         Kill nominilla(t)
       End If
     Next t
     Close 11
End Sub
Private Sub AcuVerSal_Click(Index As Integer)
 Unload Form10
 Load Form9
 Form9.Show
End Sub

Private Sub Form_Load()
    Close 11
    Open "Contr.Acu" For Random As 11 Len = Len(contr_Acu)
    dm_acu = LOF(11) / Len(contr_Acu)
    Label3.Caption = UCase(Form9.CommonDialog1.FileTitle)
    If dm_acu > 0 Then
       For r = 1 To dm_acu: Get 11, r, contr_Acu
          List1.AddItem contr_Acu.no_arch
       Next r
       Close 11
       Else
       List1.AddItem "No existen archivos actualizados"
       Close 11
     End If
    
      
End Sub

Private Sub veract_Click(Index As Integer)
   improvisa
End Sub
