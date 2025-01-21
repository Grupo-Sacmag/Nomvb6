VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Aplicacion 
   ClientHeight    =   5700
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10515
   Icon            =   "Aplicacion.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   5700
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid ObraC1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   8493
      _Version        =   393216
      Rows            =   100
      Cols            =   10
      FixedCols       =   0
      BackColorBkg    =   -2147483644
      GridLines       =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu Obr 
      Caption         =   "&Obra"
      Begin VB.Menu ObAgr 
         Caption         =   "&Agregar"
      End
      Begin VB.Menu Obsep1 
         Caption         =   "-"
      End
      Begin VB.Menu ObSal 
         Caption         =   "&Salida"
      End
   End
End
Attribute VB_Name = "Aplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cont_col As Integer, ren_ante As Long, vale As Integer
Sub Obr_ini()
     For r = 1 To final: Get 8, Form8.MSFlexGrid1.TextMatrix(r, 0), maestro
      
    If maestro.im_1 = 100 Then
        If cont_col = 5 Then
            ObraC1.TextMatrix(0, cont_col) = maestro.O_1
            Stop
            Else
                        vale = 0
                        For i = 5 To cont_col:
                            If ObraC1.TextMatrix(0, i) Then vale = 1: Exit For
                        Next i
                            If vale = 0 Then
                                cont_col = cont_col + 1
                                ObraC1.TextMatrix(0, cont_col) = maestro.O_1
                            End If
        End If
    End If
   Next r
End Sub
Private Sub Form_Load()
   ObraC1.Row = 0
    ObraC1.Col = 0: ObraC1.ColWidth(0) = 600: ObraC1.CellAlignment = 4: ObraC1.Text = "Cuenta"
    ObraC1.Col = 1: ObraC1.ColWidth(1) = 600:  ObraC1.CellAlignment = 4: ObraC1.Text = "SubCta"
    ObraC1.Col = 2: ObraC1.ColWidth(2) = 3350:  ObraC1.CellAlignment = 4: ObraC1.Text = "Nombre"
    ObraC1.Col = 3: ObraC1.ColWidth(3) = 1200:  ObraC1.CellAlignment = 4: ObraC1.Text = "Parcial"
    ObraC1.Col = 4: ObraC1.ColWidth(4) = 1100:  ObraC1.CellAlignment = 4: ObraC1.Text = "Debe"
    ObraC1.Col = 5: ObraC1.ColWidth(5) = 1100:  ObraC1.CellAlignment = 4: ObraC1.Text = "Haber"
    ObraC1.Col = 6: ObraC1.ColWidth(6) = 80
    ObraC1.Col = 7: ObraC1.ColWidth(7) = 80
    ObraC1.Col = 8: ObraC1.ColWidth(8) = 80
    ObraC1.Col = 9: ObraC1.ColWidth(9) = 80
    ObraC1.Row = 1: Aplicacion.Caption = "Aplicacion " + Form8.Label7.Caption
   cont_col = 5
   Obr_ini
   For r = 1 To final
       ObraC1.AddItem Form8.MSFlexGrid1.TextMatrix(r, 0) & Chr(9) & Form8.MSFlexGrid1.TextMatrix(r, 1) & Chr(9) & Form8.MSFlexGrid1.TextMatrix(r, 2) & Chr(9) & "" & Chr(9) & Form8.MSFlexGrid1.TextMatrix(r, 11)
   Next r
   
End Sub

Private Sub Form_Resize()
   ObraC1.Height = Aplicacion.Height * 0.8
   ObraC1.Width = Aplicacion.Width * 0.95
End Sub

Private Sub ObAgr_Click()
   ren_ante = ObraC1.Row
   ObraC1.Row = 0
   ObraC1.Col = cont_col
      
End Sub
Private Sub ObraC1_LeaveCell()
 If ObraC1.Row > 0 Then
  ObraC1.CellBackColor = vbWhite
 End If
End Sub

Private Sub obrac1_EnterCell()
   'If ObraC1.Row = 1 And ObraC1.Col = 2 Then ObraC1.Col = 1: ObraC1_EnterCell
   'If ObraC1.Row = 7 And ObraC1.Col = 2 Then ObraC1.Col = 1: ObraC1_EnterCell
   'If ObraC1.Row = 8 And ObraC1.Col = 2 Then ObraC1.Col = 1: ObraC1_EnterCell
   'If ObraC1.Row > 8 And ObraC1.Col = 1 Then ObraC1.Col = 2: ObraC1_EnterCell
   'If ObraC1.Row > ultimo.renglon + 1 Then ObraC1.Row = ultimo.renglon + 1
   'valcelant = ObraC1.Text
   If ObraC1.Row > 0 Then
    ObraC1.CellBackColor = vbGreen
   End If
End Sub

Private Sub ObraC1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
       Case 13
            If ObraC1.Row = 0 Then
                cont_col = cont_col + 1
                ObraC1.Row = ren_ante
            
            End If
       Case Else
          ObraC1.Text = ObraC1.Text + Chr(KeyAscii)
    End Select
End Sub
