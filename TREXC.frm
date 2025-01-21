VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form TREXC 
   Caption         =   "TRASPASO NOMINA BANAMEX EN EXCEL"
   ClientHeight    =   9915
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9855
   Icon            =   "TREXC.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   9915
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   8160
      MaxLength       =   12
      TabIndex        =   11
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5520
      MaxLength       =   36
      TabIndex        =   9
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3480
      MaxLength       =   12
      TabIndex        =   7
      Top             =   5400
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "30/11/2017"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      MaxLength       =   4
      TabIndex        =   2
      Top             =   5400
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Ex2 
      Height          =   4095
      Left            =   240
      TabIndex        =   1
      Top             =   5880
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   19
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
   Begin VB.Label Label5 
      Caption         =   "# Usuario 12 digitos"
      Height          =   255
      Left            =   8160
      TabIndex        =   10
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre de cliente 36 digitos"
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "# de cliente 12 digitos"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Secuencial 4 digitos"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Menu Arc 
      Caption         =   "&Archivo"
      Begin VB.Menu ArcSale 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu Edt 
      Caption         =   "&Edicion"
      Begin VB.Menu EdSel 
         Caption         =   "&Seleccionar todo"
      End
      Begin VB.Menu EdSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EdCopiar 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu GenBnx 
      Caption         =   "&Generar Bnx"
   End
   Begin VB.Menu GenTxt 
      Caption         =   "&Generar TXT"
   End
End
Attribute VB_Name = "TREXC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Col_Def()
   Ex.Clear: Ex.Rows = 1: Ex.Cols = 10: Ex.FixedCols = 0
   Ex.Text = "No."
   Ex.Col = 0: Ex.CellAlignment = 4: Ex.ColWidth(0) = 1800: Ex.Text = "Num.Cta"
   Ex.Col = 1: Ex.CellAlignment = 4: Ex.ColWidth(1) = 1200: Ex.Text = "Nombre"
   Ex.Col = 2: Ex.CellAlignment = 4: Ex.ColWidth(2) = 1200: Ex.Text = "Apell.Paterno"
   Ex.Col = 3: Ex.CellAlignment = 4: Ex.ColWidth(3) = 1200: Ex.Text = "Apell.Materno"
   Ex.Col = 4: Ex.CellAlignment = 4: Ex.ColWidth(4) = 800: Ex.Text = "Importe"
   Ex.Col = 5: Ex.CellAlignment = 4: Ex.ColWidth(5) = 800: Ex.Text = "Trabajador"
   Ex.Col = 6: Ex.CellAlignment = 4: Ex.ColWidth(6) = 800: Ex.Text = "Ref.Alfanum"
   Ex.Col = 7: Ex.CellAlignment = 4: Ex.ColWidth(7) = 800: Ex.Text = "Concepto Pago"
   Ex.Col = 8: Ex.CellAlignment = 4: Ex.ColWidth(8) = 800: Ex.Text = "MIsmoDia"
   Ex.Col = 9: Ex.CellAlignment = 4: Ex.ColWidth(9) = 600: Ex.Text = "Cons."
End Sub
  
Private Sub Form_Load()
    Col_Def
End Sub

Private Sub EdCopiar_Click()
 Dim Temporal1
 Clipboard.Clear
   
   difer = Ex.RowSel - Ex.Row
   For i = Ex.Row To Ex.RowSel
      
      For f = Ex.Col To Ex.ColSel
            Temporal1 = Temporal1 + Ex.TextMatrix(i, f)
            If f < Ex.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next f
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Next i
    Clipboard.SetText Temporal1

End Sub

Private Sub EdSel_Click()
Clipboard.Clear
    Ex.Row = 1: Ex.Col = 0
    Ex.RowSel = Ex.Rows - 1
    Ex.ColSel = Ex.Cols - 2
End Sub

Private Sub GenBnx_Click()
Ex2.TextMatrix(0, 1) = "M"
Ex2.TextMatrix(0, 2) = "BAN"
Ex2.TextMatrix(0, 3) = "TC"
Ex2.TextMatrix(0, 4) = "PRO"
Ex2.TextMatrix(0, 5) = "IN"
Ex2.TextMatrix(0, 6) = "SUC"
Ex2.TextMatrix(0, 7) = "CUENTA"
Ex2.TextMatrix(0, 8) = "PER"
Ex2.TextMatrix(0, 9) = "NOMBRE"
Ex2.TextMatrix(0, 10) = "ALIAS"
Ex2.TextMatrix(0, 11) = "M"
Ex2.TextMatrix(0, 12) = "IMPORTE"
Ex2.TextMatrix(0, 13) = "P"
Ex2.TextMatrix(0, 14) = "RFC"
Ex2.ColWidth(0) = 0
Ex2.ColWidth(1) = 250
Ex2.ColWidth(2) = 480
Ex2.ColWidth(3) = 300
Ex2.ColWidth(4) = 480
Ex2.ColWidth(5) = 300
Ex2.ColWidth(6) = 480
Ex2.ColWidth(7) = 1550
Ex2.ColWidth(8) = 300
Ex2.ColWidth(9) = 4000
Ex2.ColWidth(10) = 2200
Ex2.ColWidth(11) = 380
Ex2.ColWidth(12) = 1000
Ex2.ColWidth(13) = 250
Ex2.ColWidth(14) = 500

   Ex2.Rows = Ex.Rows
   Dim fin As Integer
    For fin = 1 To Ex2.Rows - 1
        Ex2.TextMatrix(fin, 1) = "A"
        Ex2.TextMatrix(fin, 2) = "0000"
        Ex2.TextMatrix(fin, 3) = "01"
        Ex2.TextMatrix(fin, 4) = "0000"
        Ex2.TextMatrix(fin, 5) = "00"
        Ex2.TextMatrix(fin, 6) = "00000000"
        Ex2.TextMatrix(fin, 7) = Ex.TextMatrix(fin, 0)
        Ex2.TextMatrix(fin, 8) = "01"
        Dim nombre As String * 55
        nombre = Ex.TextMatrix(fin, 1) + " " + Ex.TextMatrix(fin, 2) + " " + Ex.TextMatrix(fin, 3)
        Ex2.TextMatrix(fin, 9) = nombre
        Dim alias As String * 20
        alias = Ex.TextMatrix(fin, 1) + " " + Ex.TextMatrix(fin, 2) + " " + Ex.TextMatrix(fin, 3)
        Ex2.TextMatrix(fin, 10) = alias
        Ex2.TextMatrix(fin, 11) = "001"
        Ex2.TextMatrix(fin, 12) = "00000000099999"
        Ex2.TextMatrix(fin, 13) = "D"
        Dim RFC As String * 18
        RFC = "                  "
        Ex2.TextMatrix(fin, 14) = RFC
        Ex2.TextMatrix(fin, 15) = "04"
        Dim EMAIL As String * 40
        EMAIL = "                                        "
        Ex2.TextMatrix(fin, 16) = EMAIL
        Dim CEL As String * 10
        CEL = "0000000000"
        Ex2.TextMatrix(fin, 17) = CEL
        Ex2.TextMatrix(fin, 18) = "04"
    Next fin
End Sub

Private Sub GenTxt_Click()
Dim ret As Boolean

    ret = Exportar_FlexGrid_txt(Ex2, "archivogbnx.txt", vbTab)
    If ret Then
        MsgBox "Archivo generado como archivogbnx.txt", vbInformation
    End If
End Sub

Public Function Exportar_FlexGrid_txt(FlexGrid As Object, _
                                      Path_Txt As String, _
                                      Delimitador As Variant) As Boolean
    On Error GoTo Err_Funcion
    Dim fila As Integer
    Dim columna As Integer
    Dim Free_File As Integer
    ' Número de  archivo libre para crear el archivo de texto
    Free_File = FreeFile
    ' Abre y crea el  archivo
    Open Path_Txt For Output As #Free_File
    ' Encabezado del archivo
    Dim secuencial As String * 4
    Dim fecha As String * 10
    Dim hora As String * 5
    Dim cliente As String * 12
    Dim nombre As String * 36
    Dim usuario As String * 12
    
    secuencial = Text1.Text
    fecha = Text2.Text
    hora = "     "
    cliente = Text3.Text
    nombre = Text4.Text
    usuario = Text5.Text
    
    cadena = secuencial + fecha + hora + cliente + nombre + usuario
    Print #Free_File, cadena;
    Print #Free_File, ""

    For fila = 1 To _
        FlexGrid.Rows - 1
        FlexGrid.Row = fila
        For columna = 0 To _
            18
            FlexGrid.Col = columna
            Print #Free_File, FlexGrid.Text;
        Next
        Print #Free_File, ""
    Next
    Close
    Exportar_FlexGrid_txt = True
    Exit Function
Err_Funcion:
    Close #Free_File
    MsgBox Err.Description, vbCritical
End Function
