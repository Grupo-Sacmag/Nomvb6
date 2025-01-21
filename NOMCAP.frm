VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form NOMCF2 
   Caption         =   "Generación de CFDI"
   ClientHeight    =   8940
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10560
   LinkTopic       =   "Form9"
   ScaleHeight     =   8940
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog excel1 
      Left            =   8520
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Sua 
      Height          =   1095
      Left            =   840
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1931
      _Version        =   393216
   End
   Begin VB.TextBox LTXT1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   10815
   End
   Begin MSFlexGridLib.MSFlexGrid NOMCF 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6376
      _Version        =   393216
      FixedRows       =   0
      AllowUserResizing=   3
   End
   Begin VB.Menu NCFED 
      Caption         =   "&Edicion"
      Begin VB.Menu NCFEDCOP 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu EdSep1 
         Caption         =   "-"
      End
      Begin VB.Menu NCFEDPEG 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu NCFEDSEP1 
         Caption         =   "&CAPTION"
      End
      Begin VB.Menu NCFEDSEL 
         Caption         =   "&Selecionar Todo"
         Shortcut        =   ^E
      End
      Begin VB.Menu NCFEDPRIM 
         Caption         =   "&Seleccionar y copiar primer Plantilla"
      End
      Begin VB.Menu NCFEDSG 
         Caption         =   "&Seleccionar y copiar Segunda tabla"
      End
      Begin VB.Menu EdSep2 
         Caption         =   "-"
      End
      Begin VB.Menu EdQui 
         Caption         =   "&Quitar Titulos"
         Shortcut        =   ^Q
      End
      Begin VB.Menu EdFij 
         Caption         =   "&FijarTitulos"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu NCFEspr 
      Caption         =   "&Eliminar"
   End
   Begin VB.Menu Excel 
      Caption         =   "Generar Excel"
   End
   Begin VB.Menu SDI 
      Caption         =   "SDI"
      Index           =   0
   End
End
Attribute VB_Name = "NOMCF2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c_p As Integer
Dim EntradaDATOS
Dim NumerodePersonal As Long
Dim Folio As Long, serie As String, Nombrey As String
Dim Sub_total As Currency, Des_cto As Currency, Gravado As Currency, ConNomina As Currency
Dim Ex_to As Currency, T_ingreso As Currency, T_desctos As Currency, T_isr As Currency, T_gravado As Currency
Dim T_exento As Currency, T_imss As Currency, Reg_Patr As String
Dim Met_pago As String, MiFechaAlta, RiesgoImss, VMiFechaAlta, VMiFecha, VMiFecha1, VFal, ViaTico As Currency
Dim Met_pagcon As String
Dim RFC_EXTERNO As String
Dim prestamos As Double
Dim fonacot As Double
Dim pension As Double
Dim infonavit As Double
Dim impuesto As Double
Dim imss As Double
Dim fila As Long
    
    
Sub IniCols()
On Error GoTo ErrorHandler:
    Dim I_7 As Integer
    Close 7
    Open "DatCFDI.dno" For Random As 7 Len = Len(DATcf)
        
    cm = LOF(7) / Len(DATcf)
    NOMCF.Clear
    NOMCF.Rows = 2
    NOMCF.Cols = cm
    NOMCF.FixedCols = 0
    NOMCF.FixedRows = 1
    NOMCF.Rows = Form8.ConNom1.Rows
    NOMCF.Row = 0
            
    For I_7 = 0 To cm - 1
        Get 7, (I_7 + 1), DATcf
        NOMCF.Col = I_7
        NOMCF.CellAlignment = 4
        NOMCF.ColWidth(I7) = 1200
        NOMCF.Text = Trim(DATcf.Concepto)
    Next I_7

Close 7
Exit Sub
ErrorHandler:
    
    MsgBox ("Aun no capturas una nomina!")

End Sub

Private Sub EdFij_Click()
  NOMCF.FixedRows = 1
End Sub

Private Sub EdQui_Click()
    NOMCF.FixedRows = 0
End Sub

Private Sub Excel_Click()
    If Exportar_Excel(App.Path & "\libro1.xls", NOMCF) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If
End Sub

Public Function Exportar_Excel(sOutputPath As String, FlexGrid As Object) As Boolean
 On Error GoTo Error_Handler
  
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim fila        As Long
    Dim columna     As Long
      
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
      
    ' -- Bucle para Exportar los datos
    With FlexGrid
        For fila = 1 To .Rows - 1
            For columna = 0 To .Cols - 1
                o_Hoja.Cells(fila, columna + 1).Value = .TextMatrix(fila, columna)
            Next
        Next
    End With
    o_Libro.Close True, sOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function

Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub

Private Sub Form_Load()

    IniCols
    
    Close 10
    Close 7
    
    Open "EMP_CFDI.DNO" For Random As 10 Len = Len(EmpCFDI)
    Open "Perscfdi.dno" For Random As 7 Len = Len(Empleado_1)
    Open "PerOtre.dno" For Random As 9 Len = Len(Otros_Rgtros)
    
    Get 10, 1, EmpCFDI
    
    Folio = InputBox("Dame el numero de Folio Anterior", "CFDI NOMINA", EmpCFDI.Folio)
    serie = InputBox("Dame la serie", "CFDI NOMINA", EmpCFDI.serie)
    ConNomina = InputBox("Numero de nomina consecutivo ", "CFDI NOMINA", (EmpCFDI.Consecutivo + 1))
    
ERRORDEPAGO:
    Met_pago = InputBox("Metodo de Pago (utiliza la clave numerica)" & Chr(13) & "01" & Chr(9) & "EFECTIVO" & Chr(13) & "02" & Chr(9) & "CHEQUE NOMINATIVO" & Chr(13) & "03" & Chr(9) & "TRANSFERENCIA ELECTRONICA" & Chr(13) & "28" & Chr(9) & "TARJETA DE DEBITO", "CFDI NOMINA", "28")
    
    Select Case Met_pago
        Case "01", "02", "03", "28"
            If Met_pago = "01" Then Met_pagcon = "01 Eectivo"
            If Met_pago = "02" Then Met_pagcon = "02 Cheque nominativo"
            If Met_pago = "03" Then Met_pagcon = "03 TRANSFERENCIA"
            If Met_pago = "28" Then Met_pagcon = "28 Tarjeta de Débito"
        Case Else
            MsgBox "LA CLAVE NO EXISTE"
            GoTo ERRORDEPAGO
    End Select
    
    Reg_Patr = InputBox("Registro Patronal Imss ", "CFDI NOMINA", EmpCFDI.RegPatr)
    EntradaDATOS = "Riesgo Patronal(utiliza la clave numerica)" & Chr(13) & "1" & Chr(9) & "CLASE I" & Chr(13) & "2" & Chr(9) & "CLASE II" & Chr(13) & "3" & Chr(9) & "CLASE III" & Chr(13) & "4" & Chr(9) & "CLASE IV" & Chr(13) & "5" & Chr(9) & "CLASE V" & Chr(13)

DENUEVO:
    RiesgoImss = InputBox(EntradaDATOS, "CFDI NOMINA", 1)
    
    Select Case RiesgoImss
        Case 1 To 5
            If RiesgoImss = "1" Then RiesgoImss = "1-Clase I"
            If RiesgoImss = "2" Then RiesgoImss = "2-Clase II"
            If RiesgoImss = "3" Then RiesgoImss = "3-Clase III"
            If RiesgoImss = "4" Then RiesgoImss = "4-Clase IV"
            If RiesgoImss = "5" Then RiesgoImss = "5-Clase V"
        Case Else
            MsgBox "LA CLAVE NO EXISTE"
            GoTo DENUEVO
    End Select
        
    EmpCFDI.Folio = Folio: EmpCFDI.Consecutivo = ConNomina: EmpCFDI.serie = serie:  EmpCFDI.RegPatr = Reg_Patr: EmpCFDI.RiesgoImss = RiesgoImss

    Put 10, 1, EmpCFDI
    Dim x8 As Long
    Dim x9 As Long
    
    I8 = 0
       
    For I7 = 1 To Form8.ConNom1.Rows - 3
        Folio = Folio + 1
        If IsNumeric(Form8.ConNom1.TextMatrix(I7, 0)) Then
            NumerodePersonal = Form8.ConNom1.TextMatrix(I7, 0)
        End If
        reng
        MdAbr_1
    Next I7
    
    Get 10, 1, EmpCFDI: EmpCFDI.Folio = Folio
    Put 10, 1, EmpCFDI
    Close 7, 9, 10
        
    SDI_Click (Index)
    fila = 0
        
    Do While (NOMCF.Rows - 1)
        fila = fila + 1
        On Error GoTo handler
        If (CDbl(Trim(NOMCF.TextMatrix(fila, 56))) >= CDbl(Trim(NOMCF.TextMatrix(fila, 58)))) Then
            pintarRojo (fila)
        End If
    Loop
    
handler:

End Sub
Private Sub Llenar_FlexGrid(cn As Connection, Rs As Recordset)
    On Error GoTo ErrSub
    Screen.MousePointer = vbHourglass
    ' Deshabilita el repintado del Flexgrid
    ' Mueve el recordset al primer registro
    Rs.MoveFirst
    ' Agrega las filas necesarias en el FlexGRid
    Sua.Rows = Rs.RecordCount + 1
    Sua.Cols = Rs.Fields.Count
    'Selecciona
    Sua.Row = 1
    Sua.Col = 0
    Sua.RowSel = Sua.Rows - 1
    Sua.ColSel = Sua.Cols - 1
    'Devuelve o establece el contenido de las celdas en _
     una región de FlexGrid seleccionada. No está disponible en tiempo de diseño. _
     ( Esta linea de código es la que carga los registros )
    Sua.Clip = Rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
    Sua.Row = 1
    ' Cierra y elimina las referencias ( recordset y la conexión )
    Rs.Close
    'Set Rs = Nothing
    cn.Close
    'Set cn = Nothing
    Screen.MousePointer = vbDefault
Exit Sub
  
ErrSub:
' Mensaje de error
MsgBox Err.Description, vbCritical
Screen.MousePointer = vbDefault
End Sub
Private Sub pintarRojo(fila)
    For i = 1 To NOMCF.Cols - 1
        NOMCF.Row = fila
        NOMCF.Col = i
        NOMCF.CellForeColor = &HFF&
    Next
        
End Sub


Private Sub Form_Resize()
If Me.WindowState <> 1 Then
   NOMCF.Height = ScaleHeight - 1000
   NOMCF.Width = ScaleWidth - 600
 End If
End Sub

Private Sub NCfEdCop_Click()
      Dim Temporal1
 Clipboard.Clear
   
   difer = NOMCF.RowSel - NOMCF.Row
   For i = NOMCF.Row To NOMCF.RowSel
      
      For f = NOMCF.Col To NOMCF.ColSel
            Temporal1 = Temporal1 + NOMCF.TextMatrix(i, f)
            If f < NOMCF.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next f
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Next i
    Clipboard.SetText Temporal1
End Sub
Private Sub LTXT1_Change()
    NOMCF.Text = LTXT1.Text
End Sub


Private Sub NCFEDPRIM_Click()
Clipboard.Clear
    NOMCF.Row = 1
    NOMCF.Col = 0
    NOMCF.RowSel = NOMCF.Rows - 1
    NOMCF.ColSel = 97
    
Dim Temporal1
    Clipboard.Clear
    difer = NOMCF.RowSel - NOMCF.Row
    For i = NOMCF.Row To NOMCF.RowSel
      ' For f = ConNom1.Col To ConNom1.ColSel
      For f = 0 To NOMCF.ColSel ' esta modificacion permite copiar todas las columnas, desde la 0
            Temporal1 = Temporal1 + NOMCF.TextMatrix(i, f)
            If f < NOMCF.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next f
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
    Next i
    Clipboard.SetText Temporal1

End Sub

Private Sub NCfEdSel_Click()
     Dim limite As Long
    Clipboard.Clear
    NOMCF.Row = 1: NOMCF.Col = 0
   For limite = 1 To NOMCF.Rows - 1
       renglon = limite
    If IsNumeric(NOMCF.TextMatrix(renglon, 0)) Then
           NOMCF.RowSel = renglon
    End If
   Next limite
    NOMCF.ColSel = NOMCF.Cols - 1
End Sub

Private Sub NCFEDSG_Click()
Clipboard.Clear
    NOMCF.Row = 1
    NOMCF.Col = 98
    NOMCF.RowSel = NOMCF.Rows - 1
    NOMCF.ColSel = NOMCF.Cols - 1
Dim Temporal1
    Clipboard.Clear
    difer = NOMCF.RowSel - NOMCF.Row
    For i = NOMCF.Row To NOMCF.RowSel
      ' For f = ConNom1.Col To ConNom1.ColSel
      For f = 98 To NOMCF.ColSel ' esta modificacion permite copiar todas las columnas, desde la 0
            Temporal1 = Temporal1 + NOMCF.TextMatrix(i, f)
            If f < NOMCF.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next f
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
    Next i
    Clipboard.SetText Temporal1
End Sub

Private Sub NCFEspr_Click()
Dim conteocolumnas  As Integer
Dim totaldecolumnas As Integer
    totaldecolumnas = NOMCF.Rows - 1
    conteocolumnas = 1
    While conteocolumnas < totaldecolumnas
        If (NOMCF.TextMatrix(conteocolumnas, 95) = "") And (NOMCF.TextMatrix(conteocolumnas, 70) = "") Then
                    NOMCF.RowSel = conteocolumnas
                    NOMCF.RemoveItem conteocolumnas
                    conteocolumnas = 0
                    totaldecolumnas = (totaldecolumnas - 1)
        Else
                    conteocolumnas = (conteocolumnas + 1)
        End If
    Wend
End Sub


Private Sub NOMCF_EnterCell()
  LTXT1.Text = NOMCF.Text
End Sub

Private Sub NomCf_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            Case vbKeyDelete
               For Q = NOMCF.Row To NOMCF.RowSel
                 For W = NOMCF.Col To NOMCF.ColSel
                    NOMCF.TextMatrix(Q, W) = ""
                 Next W
               Next Q
                LTXT1.Text = NOMCF.Text
            Case vbKeyF2
                If NOMCF.Text <> "" Then valcelant = NOMCF.Text
                LTXT1.Text = LTrim(RTrim(NOMCF.Text))
                LTXT1.SetFocus
       End Select

End Sub

Private Sub NomCf_KeyPress(KeyAscii As Integer)
    valcelant = NOMCF.Text
    LTXT1.Text = Chr(KeyAscii)
    LTXT1.SetFocus
End Sub

Private Sub NomCf_LeaveCell()
  If NOMCF.Row > 0 Then
   NOMCF.CellBackColor = vbWhite
  End If
End Sub

Sub MdAbr_1()
On Error GoTo ErrorHandler:
'**************************************************************************************************
'**                                   Modulo CFDI                                                **
'**************************************************************************************************
    '1 2 FOLIO
        NOMCF2.NOMCF.TextMatrix(I7, 0) = Folio
    '2 3 SERIE
        NOMCF2.NOMCF.TextMatrix(I7, 1) = serie:
    '3 4 NOMBRE
        Get 2, NumerodePersonal, personal:
        MiFechaAlta = Trim(personal.fal)
        Nombrey = Trim(personal.nom) + " " + Trim(personal.ape1) + " " + Trim(personal.ape2)
        NOMCF2.NOMCF.TextMatrix(I7, 2) = Nombrey:
        Get 7, NumerodePersonal, Empleado_1
    '4 5 DIRECCION
        NOMCF2.NOMCF.TextMatrix(I7, 3) = Trim(Empleado_1.Direccion):
    '5 6 COLONIA
        NOMCF2.NOMCF.TextMatrix(I7, 4) = Trim(Empleado_1.Colonia):
    '6 7 CIUDAD
        NOMCF2.NOMCF.TextMatrix(I7, 5) = Trim(Empleado_1.Ciudad):
    '7 8 ESTADO
        NOMCF2.NOMCF.TextMatrix(I7, 6) = Trim(Empleado_1.Estado):
    '8 9 DELEGACION
        NOMCF2.NOMCF.TextMatrix(I7, 7) = Trim(Empleado_1.Delegacion):
    '9 10 CP
        NOMCF2.NOMCF.TextMatrix(I7, 8) = Trim(Empleado_1.Cpostal):
    '10 11 RFC
        NOMCF2.NOMCF.TextMatrix(I7, 9) = Trim(personal.RFC):
    '11 12 PAIS
        NOMCF2.NOMCF.TextMatrix(I7, 10) = "MEXICO":
    '12 13 CORREO
        NOMCF2.NOMCF.TextMatrix(I7, 11) = Trim(Empleado_1.correo):
    '13 14 OBSERVACIONES
        NOMCF2.NOMCF.TextMatrix(I7, 12) = ""
    '14 15 MONEDA
        NOMCF2.NOMCF.TextMatrix(I7, 13) = "PESOS"
    '15 16 TIPOCAMBIO
        NOMCF2.NOMCF.TextMatrix(I7, 14) = 1
'******************************************LLAMADO AL MODULO ***********************************************
    '16 17 TOTAL
        NOMCF2.NOMCF.TextMatrix(I7, 15) = Format(T_neto, "#,##0.00")
    '17 18 SUBTOTAL
        NOMCF2.NOMCF.TextMatrix(I7, 16) = Format(t_per, "#,##0.00")
    '18 19 DESCUENTO
        NOMCF2.NOMCF.TextMatrix(I7, 17) = Format(t_ded, "#,##0.00")
    '19 20 TOTALGRAVADOPERCEPCIONES
        NOMCF2.NOMCF.TextMatrix(I7, 18) = Format(t_grav, "#,##0.00")
    '20 21 TOTALEXENTOPERCEPCIONES
        NOMCF2.NOMCF.TextMatrix(I7, 19) = Format(t_ext, "#,##0.00")
    '21 22 TOTALPERCEPCIONES
        NOMCF2.NOMCF.TextMatrix(I7, 20) = Format(t_per - sub13, "#,##0.00")
    '22 23 TOTALDEDUCCIONES
        NOMCF2.NOMCF.TextMatrix(I7, 21) = Format(t_ded, "#,##0.00")
    '23 24 TOTALOTROSPAGOS
        NOMCF2.NOMCF.TextMatrix(I7, 22) = Format(sub13, "#,##0.00")
    
    '24 25 TOTALSUELDOS
     NOMCF2.NOMCF.TextMatrix(I7, 23) = Format(t_per - sub13, "#,##0.00")
    
    '25 26 TOTALSEPARACIONINDEMNIZACION
        NOMCF2.NOMCF.TextMatrix(I7, 24) = Format(0, "#,##0.00")
    '26 27 TOTALJUBILACIONPENSIONRETIRO
        NOMCF2.NOMCF.TextMatrix(I7, 25) = Format(0, "#,##0.00")
    '27 28 TOTALOTRASDEDUCCIONES
        NOMCF2.NOMCF.TextMatrix(I7, 26) = Format(t_oded, "#,##0.00")
    '28 29 TOTALIMPUESTOSRETENIDOS
        NOMCF2.NOMCF.TextMatrix(I7, 27) = Format(isr12, "#,##0.00")
    '29 30 VALORUNITARIO
        NOMCF2.NOMCF.TextMatrix(I7, 28) = Format(t_per, "#,##0.00")
    '30 31 IMPORTE
        NOMCF2.NOMCF.TextMatrix(I7, 29) = Format(t_per, "#,##0.00")
    Aletra (T_neto)
    '31 32 TOTALCONLETRA
        NOMCF2.NOMCF.TextMatrix(I7, 30) = LetrVal
    '32 33 TIPODENOMINA
      If N_ormal = 0 Then
        NOMCF2.NOMCF.TextMatrix(I7, 31) = "ORDINARIA"
        Else
        NOMCF2.NOMCF.TextMatrix(I7, 31) = "EXTRAORDINARIA"
      End If
    '33 34 SINDICALIZADO
        NOMCF2.NOMCF.TextMatrix(I7, 32) = "No"
    '34 35 METODOPAGO
        NOMCF2.NOMCF.TextMatrix(I7, 33) = Met_pagcon
    '35 36 LUGAREXPEDICION
        NOMCF2.NOMCF.TextMatrix(I7, 34) = "DIF"
    '36 37 REGIMEN
        NOMCF2.NOMCF.TextMatrix(I7, 35) = "601"
        
        '-*********************************************************************
    '36
        NOMCF2.NOMCF.TextMatrix(I7, 36) = "G03 Gastos en general"
        '-*********************************************************************
        
    '37 38 NUMCTAPAG
        'NOMCF2.NOMCF.TextMatrix(I7, 37) = CStr(Form8.ConNom1.TextMatrix(I7, 24))
        NOMCF2.NOMCF.TextMatrix(I7, 37) = Right(Trim(Form8.ConNom1.TextMatrix(I7, 24)), 4)
    '38 39 REGISTROPATRONAL
        NOMCF2.NOMCF.TextMatrix(I7, 38) = Reg_Patr
    '39 40 NUMEMPLEADO
        NOMCF2.NOMCF.TextMatrix(I7, 39) = NumerodePersonal
    '40 41 CURP
        Get 9, NumerodePersonal, Otros_Rgtros
        NOMCF2.NOMCF.TextMatrix(I7, 40) = Trim(Otros_Rgtros.curp)
    '41 42 TIPOREGIMEN
        NOMCF2.NOMCF.TextMatrix(I7, 41) = "02 Sueldos"
    '42 43 NUMSEGURIDADSOCIAL
        NOMCF2.NOMCF.TextMatrix(I7, 42) = Trim(personal.imss)
        
        DDINI = Left(MiFecha, 2)
             If DDINI < 16 Then
                  DDINI = 1
                  MiFecha1 = "01" + Mid(MiFecha, 3)
                  Else
                  DDINI = 16
                  MiFecha1 = "16" + Mid(MiFecha, 3)
             End If
        VMiFecha = Mid(Trim(MiFecha), 7, 4) + "-" + Mid(Trim(MiFecha), 4, 2) + "-" + Mid(Trim(MiFecha), 1, 2)
        VMiFecha1 = Mid(Trim(MiFecha1), 7, 4) + "-" + Mid(Trim(MiFecha1), 4, 2) + "-" + Mid(Trim(MiFecha1), 1, 2)
        
    '43 44 FECHAPAGO
        NOMCF2.NOMCF.TextMatrix(I7, 43) = VMiFecha
    '44 45 FECHAINICIALPAGO
        NOMCF2.NOMCF.TextMatrix(I7, 44) = VMiFecha1
    '45 46 FECHAFINALPAGO
        NOMCF2.NOMCF.TextMatrix(I7, 45) = VMiFecha
    Rem *****************   MODIFICADO CON PARCHE DEL 6/6/17
    '46 47 NUMDIASPAGADOS
       If N_ormal = 1 Then
        NOMCF2.NOMCF.TextMatrix(I7, 46) = 1
        Else
        NOMCF2.NOMCF.TextMatrix(I7, 46) = Form8.ConNom1.TextMatrix(I7, 2)
       End If
    '47 48 DEPARTAMENTO
        NOMCF2.NOMCF.TextMatrix(I7, 47) = "ADMINISTRACION"
    '48 49 CLABE
        NOMCF2.NOMCF.TextMatrix(I7, 48) = ""
    '49 50 BANCO
        NOMCF2.NOMCF.TextMatrix(I7, 49) = "002"
    '50 51 FECHAINICIORELLABORAL
        VFal = Mid(Trim(personal.fal), 7, 4) + "-" + Mid(Trim(personal.fal), 4, 2) + "-" + Mid(Trim(personal.fal), 1, 2)
        NOMCF2.NOMCF.TextMatrix(I7, 50) = VFal
    '51 52 ANTIGUEDAD
        NOMCF2.NOMCF.TextMatrix(I7, 51) = Year(MiFecha) - Year(MiFechaAlta)
    '52 53 PUESTO
        NOMCF2.NOMCF.TextMatrix(I7, 52) = "Administracion"
    '53 54 TIPOCONTRATO
        NOMCF2.NOMCF.TextMatrix(I7, 53) = "01 Contrato de trabajo por tiempo indeterminado"
    '54 55 TIPOJORNADA
        NOMCF2.NOMCF.TextMatrix(I7, 54) = "01 Diurna"
    '55 56 PERIODICIDADPAGO
        If N_ormal = 1 Then
                    NOMCF2.NOMCF.TextMatrix(I7, 55) = "99 OTRA PERIODICIDAD"
                    Else
                    NOMCF2.NOMCF.TextMatrix(I7, 55) = "04 QUINCENAL":
        End If
    '56 57 SALARIOBASECOTAPOR
        NOMCF2.NOMCF.TextMatrix(I7, 56) = Format(personal.ingr, "#,##0.00")
    '57 58 RIESGOPUESTO
        NOMCF2.NOMCF.TextMatrix(I7, 57) = RiesgoImss
    '58 59 SALARIODIARIOINTEGRADO
        'NOMCF2.NOMCF.TextMatrix(I7, 58) = Format(personal.integrado, "#,##0.00")
    '59 60 ENTIDADFEDERATIVA
        NOMCF2.NOMCF.TextMatrix(I7, 59) = "CMX"
    '60 61 RFCLABORA
    '61 62 PORCENTAJETIEMPO
        'If RFC_EXTERNO <> "" Then
            NOMCF2.NOMCF.TextMatrix(I7, 60) = RFC_EXTERNO
            NOMCF2.NOMCF.TextMatrix(I7, 61) = 100
        'Else
            'NOMCF2.NOMCF.TextMatrix(I7, 61) = Format(0, "#,##0.00")
        'End If
        
        '********************Modificacion de la plantilla 1 18/02/2020 ************
    '62 63 DINCAPACIDADR
        NOMCF2.NOMCF.TextMatrix(I7, 62) = Format(0, "#,##0.00")
    '63 64 DINCAPACIDADI
        NOMCF2.NOMCF.TextMatrix(I7, 63) = Format(0, "#,##0.00")
    '64 65 DINCAPACIDADM
        NOMCF2.NOMCF.TextMatrix(I7, 64) = Format(0, "#,##0.00")
    '65 66 DINCAPACIDADC
        NOMCF2.NOMCF.TextMatrix(I7, 65) = Format(0, "#,##0.00")
    '66 67 P014DINCAR
        NOMCF2.NOMCF.TextMatrix(I7, 66) = Format(0, "#,##0.00")
    '67 68 P014DINCA
        NOMCF2.NOMCF.TextMatrix(I7, 67) = Format(0, "#,##0.00")
    '68 69 P014DINCAM
        NOMCF2.NOMCF.TextMatrix(I7, 68) = Format(0, "#,##0.00")
    '69 70 P014DINCAC
        NOMCF2.NOMCF.TextMatrix(I7, 69) = Format(0, "#,##0.00")
    '70 71 DIASDOM
        NOMCF2.NOMCF.TextMatrix(I7, 70) = Format(0, "#,##0.00")
    '71 72 DIASSHE
        NOMCF2.NOMCF.TextMatrix(I7, 71) = Format(0, "#,##0.00")
    '72 73 HORASDOBLES
        NOMCF2.NOMCF.TextMatrix(I7, 72) = Format(0, "#,##0.00")
    '73 74 DIASHETRIPLES
        NOMCF2.NOMCF.TextMatrix(I7, 73) = Format(0, "#,##0.00")
    '74 75 HORASTRIPLES
        NOMCF2.NOMCF.TextMatrix(I7, 74) = Format(0, "#,##0.00")
    '75 76 DIASHESIMPLES
        NOMCF2.NOMCF.TextMatrix(I7, 75) = Format(0, "#,##0.00")
    '76 77 HORASSIMPLES
        NOMCF2.NOMCF.TextMatrix(I7, 76) = Format(0, "#,##0.00")
    '76 77 AT_VMERCADO
        NOMCF2.NOMCF.TextMatrix(I7, 77) = Format(0, "#,##0.00")
    '77 78 AT_POTORGARSE
        NOMCF2.NOMCF.TextMatrix(I7, 78) = Format(0, "#,##0.00")
    '78 79 JB_TEXTHIBICION
        NOMCF2.NOMCF.TextMatrix(I7, 79) = Format(0, "#,##0.00")
    '79 80 JB_TPARCIAL
        NOMCF2.NOMCF.TextMatrix(I7, 80) = Format(0, "#,##0.00")
    '80 81 JB_MDIARIO
        NOMCF2.NOMCF.TextMatrix(I7, 81) = Format(0, "#,##0.00")
    '81 82 JB_IACUM
        NOMCF2.NOMCF.TextMatrix(I7, 82) = Format(0, "#,##0.00")
    '82 83 JB_INOACUM
        NOMCF2.NOMCF.TextMatrix(I7, 83) = Format(0, "#,##0.00")
    '83 84 SI_TPAGADO
        NOMCF2.NOMCF.TextMatrix(I7, 84) = Format(0, "#,##0.00")
    '84 85 SI_ASERVICIO
        NOMCF2.NOMCF.TextMatrix(I7, 85) = Format(0, "#,##0.00")
    '85 86 SI_USUELDO
        NOMCF2.NOMCF.TextMatrix(I7, 86) = Format(0, "#,##0.00")
    '86 87 SI_IACUM
        NOMCF2.NOMCF.TextMatrix(I7, 87) = Format(0, "#,##0.00")
    '87 88 SI_INOACUM
        NOMCF2.NOMCF.TextMatrix(I7, 88) = Format(0, "#,##0.00")
    '88 89 CSF_SALFAV
        NOMCF2.NOMCF.TextMatrix(I7, 89) = Format(0, "#,##0.00")
    '89 90 CSF_ANIO
        NOMCF2.NOMCF.TextMatrix(I7, 90) = Format(0, "#,##0.00")
    '90 91 CSF_RSFAV
        NOMCF2.NOMCF.TextMatrix(I7, 91) = Format(0, "#,##0.00")
    '91 92 SE_SCAUSADO
        NOMCF2.NOMCF.TextMatrix(I7, 92) = Format(subc13, "#,##0.00")
    '92 93 UUID_DOCREL
        NOMCF2.NOMCF.TextMatrix(I7, 93) = ""
    '93 94 H2TOT
        NOMCF2.NOMCF.TextMatrix(I7, 94) = Format(0, "#,##0.00")
    '94 95 H3TOT
        NOMCF2.NOMCF.TextMatrix(I7, 95) = Format(0, "#,##0.00")
    '95 96 HSTOT
        NOMCF2.NOMCF.TextMatrix(I7, 96) = Format(0, "#,##0.00")
    '96 97 BANDERA
        NOMCF2.NOMCF.TextMatrix(I7, 97) = 1
    '*****************************SEGUNDA PLANTILLA**********************************************
        '***********PERCEPCIONES*************
    '97 98 Rfc
        NOMCF2.NOMCF.TextMatrix(I7, 98) = Trim(personal.RFC):
    '98 99 P001G
        NOMCF2.NOMCF.TextMatrix(I7, 99) = Format(sue3, "#,##0.00")
    '99 100 P002G
    '100 101 P002E
        If N_ormal = 1 Then
            If IsNumeric(Form8.ConNom1.TextMatrix(I7, 5)) Then
                NOMCF2.NOMCF.TextMatrix(I7, 100) = Format(Form8.ConNom1.TextMatrix(I7, 5), "###0.00"):  Rem 64 P002 GGratificación Anual (Aguinaldo)
            If IsNumeric(Form8.ConNom1.TextMatrix(I7, 6)) Then
                NOMCF2.NOMCF.TextMatrix(I7, 101) = Format(0, "###0.00")
                Else
                NOMCF2.NOMCF.TextMatrix(I7, 101) = Format(Form8.ConNom1.TextMatrix(I7, 10), "###0.00"): Rem 65 P002 GGratificación Anual (Aguinaldo exento)
           End If
            ElseIf Form8.ConNom1.TextMatrix(I7, 5) = "" Then
                NOMCF2.NOMCF.TextMatrix(I7, 100) = Format(0, "#,##0.00")
                NOMCF2.NOMCF.TextMatrix(I7, 101) = Format(0, "#,##0.00")
            End If
        Else
                NOMCF2.NOMCF.TextMatrix(I7, 100) = Format(0, "#,##0.00")
                NOMCF2.NOMCF.TextMatrix(I7, 101) = Format(0, "#,##0.00")
        End If
    '101 102
    '102 103
    'P003G PTU
    'P003E PTU
    Rem **************** MODIFICADO CON EL PARCHE DEL 6/6/17 *************************************
    If N_ormal = 1 Then
         
        If IsNumeric(Form8.ConNom1.TextMatrix(I7, 6)) Then
            NOMCF2.NOMCF.TextMatrix(I7, 102) = Format(ptu_1, "###0.00"):          Rem 66 P003E PTU
            NOMCF2.NOMCF.TextMatrix(I7, 103) = Format(ptu_2, "###0.00"):          Rem 67 P003G PTU
        ElseIf Form8.ConNom1.TextMatrix(I7, 6) = "" Then
            NOMCF2.NOMCF.TextMatrix(I7, 102) = Format(0, "#,##0.00")
            NOMCF2.NOMCF.TextMatrix(I7, 103) = Format(0, "#,##0.00")
        End If
    Else
            NOMCF2.NOMCF.TextMatrix(I7, 102) = Format(0, "#,##0.00")
            NOMCF2.NOMCF.TextMatrix(I7, 103) = Format(0, "#,##0.00")
    End If
    
    '103 104 P004GGTOS.MEDICOS
        NOMCF2.NOMCF.TextMatrix(I7, 104) = Format(0, "#,##0.00")
    '105 P004E
        NOMCF2.NOMCF.TextMatrix(I7, 105) = Format(0, "#,##0.00")
    '106 P005G
        NOMCF2.NOMCF.TextMatrix(I7, 106) = Format(0, "#,##0.00")
    '107 P005E
        NOMCF2.NOMCF.TextMatrix(I7, 107) = Format(0, "#,##0.00")
    '108 P006G
        NOMCF2.NOMCF.TextMatrix(I7, 108) = Format(0, "#,##0.00")
    '109 P006E
        NOMCF2.NOMCF.TextMatrix(I7, 109) = Format(0, "#,##0.00")
    '110 P009G
        NOMCF2.NOMCF.TextMatrix(I7, 110) = Format(0, "#,##0.00")
    '111 P010G PREMIO DE PUNTUALIDAD
    
    If N_ormal = 1 Then
        If IsNumeric(Form8.ConNom1.TextMatrix(I7, 7)) Then
            NOMCF2.NOMCF.TextMatrix(I7, 104) = Format(Form8.ConNom1.TextMatrix(I7, 7) + Form8.ConNom1.TextMatrix(I7, 10), "###0.00"):           Rem 74 P003GPREMIO DE PUNTUALIDAD
        End If
        
    Else
        NOMCF2.NOMCF.TextMatrix(I7, 104) = Format(0, "#,##0.00")
    End If
        NOMCF2.NOMCF.TextMatrix(I7, 111) = Format(0, "#,##0.00")
    '112 P011G
        NOMCF2.NOMCF.TextMatrix(I7, 112) = Format(0, "#,##0.00")
    '113 P011E
        NOMCF2.NOMCF.TextMatrix(I7, 113) = Format(0, "#,##0.00")
    '114 P012G
        NOMCF2.NOMCF.TextMatrix(I7, 114) = Format(0, "#,##0.00")
    '115 P012E
        NOMCF2.NOMCF.TextMatrix(I7, 115) = Format(0, "#,##0.00")
    '116 P013G
        NOMCF2.NOMCF.TextMatrix(I7, 116) = Format(0, "#,##0.00")
    '117 P014R
        NOMCF2.NOMCF.TextMatrix(I7, 117) = Format(0, "#,##0.00")
    '118 P014I
        NOMCF2.NOMCF.TextMatrix(I7, 118) = Format(0, "#,##0.00")
    '119 P014M
        NOMCF2.NOMCF.TextMatrix(I7, 119) = Format(0, "#,##0.00")
    '120 P014C
        NOMCF2.NOMCF.TextMatrix(I7, 120) = Format(0, "#,##0.00")
    '121 P015E
        NOMCF2.NOMCF.TextMatrix(I7, 121) = Format(0, "#,##0.00")
    '122 P019G
        NOMCF2.NOMCF.TextMatrix(I7, 122) = Format(0, "#,##0.00")
    '123 P019E
        NOMCF2.NOMCF.TextMatrix(I7, 123) = Format(0, "#,##0.00")
    '124 P020G
        NOMCF2.NOMCF.TextMatrix(I7, 124) = Format(0, "#,##0.00")
    '125 P020E
        NOMCF2.NOMCF.TextMatrix(I7, 125) = Format(0, "#,##0.00")
    '126 P021G
       NOMCF2.NOMCF.TextMatrix(I7, 126) = Format(pva8, "#,##0.00")
    '127 P021E
    If N_ormal = 0 Then
        NOMCF2.NOMCF.TextMatrix(I7, 127) = Format(pee10, "#,##0.00")
    Else
        NOMCF2.NOMCF.TextMatrix(I7, 127) = Format(0, "#,##0.00")
    End If
    '128 P022G
        NOMCF2.NOMCF.TextMatrix(I7, 128) = Format(0, "#,##0.00")
    '129 P022E
        NOMCF2.NOMCF.TextMatrix(I7, 129) = Format(0, "#,##0.00")
    '130 P023G
        NOMCF2.NOMCF.TextMatrix(I7, 130) = Format(0, "#,##0.00")
    '131 P024G
        NOMCF2.NOMCF.TextMatrix(I7, 131) = Format(0, "#,##0.00")
    '132 P024E
        NOMCF2.NOMCF.TextMatrix(I7, 132) = Format(0, "#,##0.00")
    '133 P025G
        NOMCF2.NOMCF.TextMatrix(I7, 133) = Format(0, "#,##0.00")
    '134 P025E
        NOMCF2.NOMCF.TextMatrix(I7, 134) = Format(0, "#,##0.00")
    '135 P026G
        NOMCF2.NOMCF.TextMatrix(I7, 135) = Format(0, "#,##0.00")
    '135 P026E
        NOMCF2.NOMCF.TextMatrix(I7, 136) = Format(0, "#,##0.00")
    '137 P027E
        NOMCF2.NOMCF.TextMatrix(I7, 137) = Format(0, "#,##0.00")
    '138 P028G
        NOMCF2.NOMCF.TextMatrix(I7, 138) = Format(0, "#,##0.00")
    '139 P029E
        NOMCF2.NOMCF.TextMatrix(I7, 139) = Format(0, "#,##0.00")
    '140 P030E
        NOMCF2.NOMCF.TextMatrix(I7, 140) = Format(0, "#,##0.00")
    '141 P031E
        NOMCF2.NOMCF.TextMatrix(I7, 141) = Format(0, "#,##0.00")
    '142 P032E
        NOMCF2.NOMCF.TextMatrix(I7, 142) = Format(0, "#,##0.00")
    '143 P033E
        NOMCF2.NOMCF.TextMatrix(I7, 143) = Format(0, "#,##0.00")
    '144 P034E
        NOMCF2.NOMCF.TextMatrix(I7, 144) = Format(0, "#,##0.00")
    '145 P035E
        NOMCF2.NOMCF.TextMatrix(I7, 145) = Format(0, "#,##0.00")
    '146 P036E
        NOMCF2.NOMCF.TextMatrix(I7, 146) = Format(0, "#,##0.00")
    '147 P037E
        NOMCF2.NOMCF.TextMatrix(I7, 147) = Format(0, "#,##0.00")
    '148 P038G Otros ingresos por salarios
        NOMCF2.NOMCF.TextMatrix(I7, 148) = Format(otr9 + via7, "#,##0.00")
    '149 P038E
        NOMCF2.NOMCF.TextMatrix(I7, 149) = Format(0, "#,##0.00")
    '150 P039E
        NOMCF2.NOMCF.TextMatrix(I7, 150) = Format(0, "#,##0.00")
    '151 P044E
        NOMCF2.NOMCF.TextMatrix(I7, 151) = Format(0, "#,##0.00")
    '152 P045E
        NOMCF2.NOMCF.TextMatrix(I7, 152) = Format(0, "#,##0.00")
    '153 P046G
        NOMCF2.NOMCF.TextMatrix(I7, 153) = Format(0, "#,##0.00")
    '154 P047G
        NOMCF2.NOMCF.TextMatrix(I7, 154) = Format(0, "#,##0.00")
    '155 P047E
        NOMCF2.NOMCF.TextMatrix(I7, 155) = Format(0, "#,##0.00")
    '156 P048G
        NOMCF2.NOMCF.TextMatrix(I7, 156) = Format(0, "#,##0.00")
    '157 P048E
        NOMCF2.NOMCF.TextMatrix(I7, 157) = Format(0, "#,##0.00")
    '158 P049G
        NOMCF2.NOMCF.TextMatrix(I7, 158) = Format(0, "#,##0.00")
    '159 P050G
        NOMCF2.NOMCF.TextMatrix(I7, 159) = Format(0, "#,##0.00")
    '160 P050E
        NOMCF2.NOMCF.TextMatrix(I7, 160) = Format(0, "#,##0.00")
    '161 P051G
        NOMCF2.NOMCF.TextMatrix(I7, 161) = Format(0, "#,##0.00")
    '162 P051E
        NOMCF2.NOMCF.TextMatrix(I7, 162) = Format(0, "#,##0.00")
    '163 P052G
        NOMCF2.NOMCF.TextMatrix(I7, 163) = Format(0, "#,##0.00")
    '164 P052E
        NOMCF2.NOMCF.TextMatrix(I7, 164) = Format(0, "#,##0.00")
    '165 P053G
        NOMCF2.NOMCF.TextMatrix(I7, 165) = Format(0, "#,##0.00")
    '166 P053E
        NOMCF2.NOMCF.TextMatrix(I7, 166) = Format(0, "#,##0.00")
    '*************************DEDUCCIONES******************************************************
    '167 D001
        NOMCF2.NOMCF.TextMatrix(I7, 167) = Format(ims14, "#,##0.00")
    '168 D002
        NOMCF2.NOMCF.TextMatrix(I7, 168) = Format(isr12, "#,##0.00")
    '169 D003
        NOMCF2.NOMCF.TextMatrix(I7, 169) = Format(0, "#,##0.00")
    '170 D004
         NOMCF2.NOMCF.TextMatrix(I7, 170) = Format(pre15, "#,##0.00")
    '171 D005
        NOMCF2.NOMCF.TextMatrix(I7, 171) = Format(0, "#,##0.00")
    '172 D006R
        NOMCF2.NOMCF.TextMatrix(I7, 172) = Format(0, "#,##0.00")
    '173 D006I
        NOMCF2.NOMCF.TextMatrix(I7, 173) = Format(0, "#,##0.00")
    '174 D006M
        NOMCF2.NOMCF.TextMatrix(I7, 174) = Format(0, "#,##0.00")
    '175 D006C
        NOMCF2.NOMCF.TextMatrix(I7, 175) = Format(0, "#,##0.00")
    '176 D007
        NOMCF2.NOMCF.TextMatrix(I7, 176) = Format(pea17, "#,##0.00")
    '177 D008
        NOMCF2.NOMCF.TextMatrix(I7, 177) = Format(0, "#,##0.00")
    '178 D009
        NOMCF2.NOMCF.TextMatrix(I7, 178) = Format(0, "#,##0.00")
    '179 D010
        NOMCF2.NOMCF.TextMatrix(I7, 179) = Format(ifv18, "#,##0.00")
    '180 D011
        NOMCF2.NOMCF.TextMatrix(I7, 180) = Format(fon16, "#,##0.00")
    
    '178 D012
        NOMCF2.NOMCF.TextMatrix(I7, 181) = Format(0, "#,##0.00")
    '178 D013
        NOMCF2.NOMCF.TextMatrix(I7, 182) = Format(0, "#,##0.00")
    '178 D014
        NOMCF2.NOMCF.TextMatrix(I7, 183) = Format(0, "#,##0.00")
    '178 D015
        NOMCF2.NOMCF.TextMatrix(I7, 184) = Format(0, "#,##0.00")
    '178 D016
        NOMCF2.NOMCF.TextMatrix(I7, 185) = Format(0, "#,##0.00")
    '178 D017
        NOMCF2.NOMCF.TextMatrix(I7, 186) = Format(0, "#,##0.00")
    '178 D018
        NOMCF2.NOMCF.TextMatrix(I7, 187) = Format(0, "#,##0.00")
    '178 D019
        NOMCF2.NOMCF.TextMatrix(I7, 188) = Format(0, "#,##0.00")
    '178 D020
        NOMCF2.NOMCF.TextMatrix(I7, 189) = Format(0, "#,##0.00")
    '178 D021
        NOMCF2.NOMCF.TextMatrix(I7, 190) = Format(0, "#,##0.00")
    '178 D022
        NOMCF2.NOMCF.TextMatrix(I7, 191) = Format(0, "#,##0.00")
    '178 D023
        NOMCF2.NOMCF.TextMatrix(I7, 192) = Format(0, "#,##0.00")
    '178 D024
        NOMCF2.NOMCF.TextMatrix(I7, 193) = Format(0, "#,##0.00")
    '178 D025
        NOMCF2.NOMCF.TextMatrix(I7, 194) = Format(0, "#,##0.00")
    '178 D026
        NOMCF2.NOMCF.TextMatrix(I7, 195) = Format(0, "#,##0.00")
    '178 D027
        NOMCF2.NOMCF.TextMatrix(I7, 196) = Format(0, "#,##0.00")
    '178 D028
        NOMCF2.NOMCF.TextMatrix(I7, 197) = Format(0, "#,##0.00")
    '178 D029
        NOMCF2.NOMCF.TextMatrix(I7, 198) = Format(0, "#,##0.00")
    '178 D030
        NOMCF2.NOMCF.TextMatrix(I7, 199) = Format(0, "#,##0.00")
    '178 D031
        NOMCF2.NOMCF.TextMatrix(I7, 200) = Format(0, "#,##0.00")
    '178 D032
        NOMCF2.NOMCF.TextMatrix(I7, 201) = Format(0, "#,##0.00")
    '178 D033
        NOMCF2.NOMCF.TextMatrix(I7, 202) = Format(0, "#,##0.00")
    '178 D034
        NOMCF2.NOMCF.TextMatrix(I7, 203) = Format(0, "#,##0.00")
    '178 D035
        NOMCF2.NOMCF.TextMatrix(I7, 204) = Format(0, "#,##0.00")
    '178 D036
        NOMCF2.NOMCF.TextMatrix(I7, 205) = Format(0, "#,##0.00")
    '178 D037
        NOMCF2.NOMCF.TextMatrix(I7, 206) = Format(0, "#,##0.00")
    '178 D038
        NOMCF2.NOMCF.TextMatrix(I7, 207) = Format(0, "#,##0.00")
    '178 D039
        NOMCF2.NOMCF.TextMatrix(I7, 208) = Format(0, "#,##0.00")
    '178 D040
        NOMCF2.NOMCF.TextMatrix(I7, 209) = Format(0, "#,##0.00")
    '178 D041
        NOMCF2.NOMCF.TextMatrix(I7, 210) = Format(0, "#,##0.00")
    '178 D042
        NOMCF2.NOMCF.TextMatrix(I7, 211) = Format(0, "#,##0.00")
    '178 D043
        NOMCF2.NOMCF.TextMatrix(I7, 212) = Format(0, "#,##0.00")
    '178 D044
        NOMCF2.NOMCF.TextMatrix(I7, 213) = Format(0, "#,##0.00")
    '178 D045
        NOMCF2.NOMCF.TextMatrix(I7, 214) = Format(0, "#,##0.00")
    '178 D046
        NOMCF2.NOMCF.TextMatrix(I7, 215) = Format(0, "#,##0.00")
    '178 D047
        NOMCF2.NOMCF.TextMatrix(I7, 216) = Format(0, "#,##0.00")
    '178 D048
        NOMCF2.NOMCF.TextMatrix(I7, 217) = Format(0, "#,##0.00")
    '178 D049
        NOMCF2.NOMCF.TextMatrix(I7, 218) = Format(0, "#,##0.00")
    '178 D050
        NOMCF2.NOMCF.TextMatrix(I7, 219) = Format(0, "#,##0.00")
    '178 D051
        NOMCF2.NOMCF.TextMatrix(I7, 220) = Format(0, "#,##0.00")
    '178 D052
        NOMCF2.NOMCF.TextMatrix(I7, 221) = Format(0, "#,##0.00")
    '178 D053
        NOMCF2.NOMCF.TextMatrix(I7, 222) = Format(0, "#,##0.00")
    '178 D054
        NOMCF2.NOMCF.TextMatrix(I7, 223) = Format(0, "#,##0.00")
    '178 D055
        NOMCF2.NOMCF.TextMatrix(I7, 224) = Format(0, "#,##0.00")
    '178 D056
        NOMCF2.NOMCF.TextMatrix(I7, 225) = Format(0, "#,##0.00")
    '178 D057
        NOMCF2.NOMCF.TextMatrix(I7, 226) = Format(0, "#,##0.00")
    '178 D058
        NOMCF2.NOMCF.TextMatrix(I7, 227) = Format(0, "#,##0.00")
    '178 D059
        NOMCF2.NOMCF.TextMatrix(I7, 228) = Format(0, "#,##0.00")
    '178 D060
        NOMCF2.NOMCF.TextMatrix(I7, 229) = Format(0, "#,##0.00")
    '178 D061
        NOMCF2.NOMCF.TextMatrix(I7, 230) = Format(0, "#,##0.00")
    '178 D062
        NOMCF2.NOMCF.TextMatrix(I7, 231) = Format(0, "#,##0.00")
    '178 D063
        NOMCF2.NOMCF.TextMatrix(I7, 232) = Format(0, "#,##0.00")
    '178 D064
        NOMCF2.NOMCF.TextMatrix(I7, 233) = Format(0, "#,##0.00")
    '178 D065
        NOMCF2.NOMCF.TextMatrix(I7, 234) = Format(0, "#,##0.00")
    '178 D066
        NOMCF2.NOMCF.TextMatrix(I7, 235) = Format(0, "#,##0.00")
    '178 D067
        NOMCF2.NOMCF.TextMatrix(I7, 236) = Format(0, "#,##0.00")
    '178 D068
        NOMCF2.NOMCF.TextMatrix(I7, 237) = Format(0, "#,##0.00")
    '178 D069
        NOMCF2.NOMCF.TextMatrix(I7, 238) = Format(0, "#,##0.00")
    '178 D070
        NOMCF2.NOMCF.TextMatrix(I7, 239) = Format(0, "#,##0.00")
    '178 D071
        NOMCF2.NOMCF.TextMatrix(I7, 240) = Format(0, "#,##0.00")
    '178 D072
        NOMCF2.NOMCF.TextMatrix(I7, 241) = Format(0, "#,##0.00")
    '178 D073
        NOMCF2.NOMCF.TextMatrix(I7, 242) = Format(0, "#,##0.00")
    '178 D074
        NOMCF2.NOMCF.TextMatrix(I7, 243) = Format(0, "#,##0.00")
    '178 D075
        NOMCF2.NOMCF.TextMatrix(I7, 244) = Format(0, "#,##0.00")
    '178 D076
        NOMCF2.NOMCF.TextMatrix(I7, 245) = Format(0, "#,##0.00")
    '178 D077
        NOMCF2.NOMCF.TextMatrix(I7, 246) = Format(0, "#,##0.00")
    '178 D078
        NOMCF2.NOMCF.TextMatrix(I7, 247) = Format(0, "#,##0.00")
    '178 D079
        NOMCF2.NOMCF.TextMatrix(I7, 248) = Format(0, "#,##0.00")
    '178 D080
        NOMCF2.NOMCF.TextMatrix(I7, 249) = Format(0, "#,##0.00")
    '178 D081
        NOMCF2.NOMCF.TextMatrix(I7, 250) = Format(0, "#,##0.00")
    '178 D082
        NOMCF2.NOMCF.TextMatrix(I7, 251) = Format(0, "#,##0.00")
    '178 D083
        NOMCF2.NOMCF.TextMatrix(I7, 252) = Format(0, "#,##0.00")
    '178 D084
        NOMCF2.NOMCF.TextMatrix(I7, 253) = Format(0, "#,##0.00")
    '178 D085
        NOMCF2.NOMCF.TextMatrix(I7, 254) = Format(0, "#,##0.00")
    '178 D086
        NOMCF2.NOMCF.TextMatrix(I7, 255) = Format(0, "#,##0.00")
    '178 D087
        NOMCF2.NOMCF.TextMatrix(I7, 256) = Format(0, "#,##0.00")
    '178 D088
        NOMCF2.NOMCF.TextMatrix(I7, 257) = Format(0, "#,##0.00")
    '178 D089
        NOMCF2.NOMCF.TextMatrix(I7, 258) = Format(0, "#,##0.00")
    '178 D090
        NOMCF2.NOMCF.TextMatrix(I7, 259) = Format(0, "#,##0.00")
    '178 D091
        NOMCF2.NOMCF.TextMatrix(I7, 260) = Format(0, "#,##0.00")
    '178 D092
        NOMCF2.NOMCF.TextMatrix(I7, 261) = Format(0, "#,##0.00")
    '178 D093
        NOMCF2.NOMCF.TextMatrix(I7, 262) = Format(0, "#,##0.00")
    '178 D094
        NOMCF2.NOMCF.TextMatrix(I7, 263) = Format(0, "#,##0.00")
    '178 D095
        NOMCF2.NOMCF.TextMatrix(I7, 264) = Format(0, "#,##0.00")
    '178 D096
        NOMCF2.NOMCF.TextMatrix(I7, 265) = Format(0, "#,##0.00")
    '178 D097
        NOMCF2.NOMCF.TextMatrix(I7, 266) = Format(0, "#,##0.00")
    '178 D098
        NOMCF2.NOMCF.TextMatrix(I7, 267) = Format(0, "#,##0.00")
    '178 D099
        NOMCF2.NOMCF.TextMatrix(I7, 268) = Format(0, "#,##0.00")
    '178 D100
        NOMCF2.NOMCF.TextMatrix(I7, 269) = Format(0, "#,##0.00")
    '178 D101
        NOMCF2.NOMCF.TextMatrix(I7, 270) = Format(0, "#,##0.00")
    '178 D102
        NOMCF2.NOMCF.TextMatrix(I7, 271) = Format(0, "#,##0.00")
    '178 D103
        NOMCF2.NOMCF.TextMatrix(I7, 272) = Format(0, "#,##0.00")
    '178 D104
        NOMCF2.NOMCF.TextMatrix(I7, 273) = Format(0, "#,##0.00")
    '178 D105
        NOMCF2.NOMCF.TextMatrix(I7, 274) = Format(0, "#,##0.00")
    '178 D106
        NOMCF2.NOMCF.TextMatrix(I7, 275) = Format(0, "#,##0.00")
    '178 D107
        NOMCF2.NOMCF.TextMatrix(I7, 276) = Format(0, "#,##0.00")
    '277 OP001
        NOMCF2.NOMCF.TextMatrix(I7, 277) = Format(0, "#,##0.00")
    '278 OP002
        NOMCF2.NOMCF.TextMatrix(I7, 278) = Format(sub13, "#,##0.00")
    '277 OP003
        NOMCF2.NOMCF.TextMatrix(I7, 279) = Format(0, "#,##0.00")
    '277 OP004
        NOMCF2.NOMCF.TextMatrix(I7, 280) = Format(0, "#,##0.00")
    '277 OP005
        NOMCF2.NOMCF.TextMatrix(I7, 281) = Format(0, "#,##0.00")
    '277 OP006
        NOMCF2.NOMCF.TextMatrix(I7, 282) = Format(0, "#,##0.00")
    '277 OP007
        NOMCF2.NOMCF.TextMatrix(I7, 283) = Format(0, "#,##0.00")
    '277 OP008
        NOMCF2.NOMCF.TextMatrix(I7, 284) = Format(0, "#,##0.00")
    '277 OP009
        NOMCF2.NOMCF.TextMatrix(I7, 285) = Format(0, "#,##0.00")
    '277 OP999
        NOMCF2.NOMCF.TextMatrix(I7, 286) = Format(0, "#,##0.00")
    '277 BANDERA
        NOMCF2.NOMCF.TextMatrix(I7, 287) = 1

Exit Sub

ErrorHandler:

        If Form8.ConNom1.TextMatrix(1, 1) = "" Then 'Pregunta si la celda de nombre esta vacia
            Exit Sub
        Else
            MsgBox ("Ups!, Parece que algo paso. Revisa los datos de: " & vbCrLf & vbCrLf & Nombrey)
        End If
    
End Sub


Private Sub SDI_Click(Index As Integer)

On Error GoTo Gestionaerror

Close 19
Open "C:\salarioDI\Archivo.cg" For Random As 19 Len = Len(salarioDI)
For f = 1 To NOMCF.Rows - 3
    Get 19, f, salarioDI
    NOMCF2.NOMCF.TextMatrix(f, 58) = Format(salarioDI.Salario, "#,##0.00")
Next f
Close 19

Gestionaerror:
        If Err.Number <> 0 Then
            MsgBox "Se ha producido un error. Tipo de error = " & Err.Number & ". Descripción: " & Err.Description & Chr(13) & "Debes verificar que tengas el directorio C:\salarioDI\Archivo.cg", vbOKOnly, "C:\salarioDI\Archivo.cg"
        End If
End Sub



