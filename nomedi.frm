VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form4 
   BackColor       =   &H8000000A&
   Caption         =   "NOMINA: EDICION DE PERSONAL"
   ClientHeight    =   8640
   ClientLeft      =   465
   ClientTop       =   885
   ClientWidth     =   17895
   Icon            =   "nomedi.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   8640
   ScaleWidth      =   17895
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton BtnCancelarCambios 
      Caption         =   "Cancelar Cambios"
      Height          =   375
      Left            =   16200
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton BtnGuardarCambios 
      Caption         =   "Guardar Cambios"
      Height          =   375
      Left            =   14520
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton BtnHabilitarEdicion 
      Caption         =   "Habilitar Edicion"
      Height          =   375
      Left            =   12840
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   11760
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   120
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid dat 
      Height          =   4935
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8705
      _Version        =   393216
   End
   Begin VB.TextBox TexEdicion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7335
   End
   Begin MSFlexGridLib.MSFlexGrid ListPer 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   13361
      _Version        =   393216
      Cols            =   6
      FixedCols       =   2
   End
   Begin VB.Menu editarch 
      Caption         =   "&Archivo"
      Index           =   0
      Begin VB.Menu ArOrd 
         Caption         =   "&Ordenar"
         Begin VB.Menu ArOrAlf 
            Caption         =   "&Alfabeticamente"
         End
         Begin VB.Menu ArOrNum 
            Caption         =   "&Numericamente"
         End
      End
      Begin VB.Menu saldosSIA 
         Caption         =   "Actualizar SIA"
      End
      Begin VB.Menu ArSep1 
         Caption         =   "-"
      End
      Begin VB.Menu editsale 
         Caption         =   "&Salida"
         Index           =   1
      End
   End
   Begin VB.Menu EditAr 
      Caption         =   "&Edicion"
      Begin VB.Menu EdCopiar 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu EdiSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EdPegar 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu EdiSep2 
         Caption         =   "-"
      End
      Begin VB.Menu EditSelTot 
         Caption         =   "&Seleccionar Todo"
      End
      Begin VB.Menu EdSep3 
         Caption         =   "-"
      End
      Begin VB.Menu EdSup 
         Caption         =   "&Suprimir titulos"
      End
   End
   Begin VB.Menu edim 
      Caption         =   "&Imprimir"
      Index           =   0
      Begin VB.Menu edimalf 
         Caption         =   "&Pantalla"
         Index           =   1
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu edimnum 
         Caption         =   "&Total"
         Index           =   2
      End
   End
   Begin VB.Menu helpMe 
      Caption         =   "&Ayuda"
      Begin VB.Menu sinFiltroCarga 
         Caption         =   "Carga General"
      End
      Begin VB.Menu calcularIntegrado 
         Caption         =   "Calcular integrado"
      End
      Begin VB.Menu isrCalculo 
         Caption         =   "Calcular ISR"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ValorAnt
Dim ModoEdicion As Boolean

' ============================================================
' GRID PRINCIPAL / EDICION DE CELDAS
' ============================================================

Sub RecPra()
    ListPer.TextMatrix(ListPer.Row, 1) = Trim(personal.ape1) + " " + Trim(personal.ape2) + " " + Trim(personal.nom)
    ListPer.TextMatrix(ListPer.Row, 2) = Trim(personal.RFC)
    ListPer.TextMatrix(ListPer.Row, 3) = Trim(Otros_Rgtros.curp)
    ListPer.TextMatrix(ListPer.Row, 4) = " " + Replace(Trim(personal.imss), "-", "")
    ListPer.TextMatrix(ListPer.Row, 5) = Trim(personal.fal)
    ListPer.TextMatrix(ListPer.Row, 6) = Trim(personal.fab)
    ListPer.TextMatrix(ListPer.Row, 7) = Format(personal.ingr, z1)
    ListPer.TextMatrix(ListPer.Row, 8) = Format(personal.viat, z1)
    ListPer.TextMatrix(ListPer.Row, 9) = Format(personal.otras, z1)
    ListPer.TextMatrix(ListPer.Row, 10) = Format(personal.integrado, z2)
   Rem ListPer.TextMatrix(ListPer.Row, 11) = Format(Clbnx.Q1, "<")
End Sub

Sub ArchCorr()
  If Not ModoEdicion Then
      MsgBox "Debes pulsar 'Habilitar edición' antes de modificar datos.", vbExclamation
      ListPer.Text = ValorAnt
      Exit Sub
  End If
  Close 1, 3
  rgtro = ListPer.TextMatrix(ListPer.Row, 0)   ' <-- mover esta línea PRIMERO

  Open "personal.dno" For Random As 1 Len = Len(personal)
  Get 1, rgtro, personal

  Open "PerOtre.dno" For Random As 3 Len = Len(Otros_Rgtros)
  Get 3, rgtro, Otros_Rgtros
  
    Select Case ListPer.col
        Case 0, 1
            Rem nada
        Case 2
            personal.RFC = ListPer.Text
            Put 1, rgtro, personal
        Case 3
            Otros_Rgtros.curp = ListPer.Text
            Put 3, rgtro, Otros_Rgtros
        Case 4
            personal.imss = Replace(LTrim(ListPer.Text), "-", "")
            Put 1, rgtro, personal
        Case 5
            If ListPer.Text <> "" Then
                Mes = Mid(ListPer.Text, 4, 2)
                yea = Mid(ListPer.Text, 7, 4)
                dia = Left(ListPer.Text, 2)
                
                If (Mes < 1) Or (Mes > 12) Then
                    MsgBox "Mes invalido "
                    Tex_Edicion = ValorAnt
                    ListPer.Text = ValorAnt
                    GoTo salte5
                End If
                
                If (dia < 1) Or (dia > dd(Mes)) Then
                    MsgBox "dia invalido "
                    Tex_Edicion = ValorAnt
                    ListPer.Text = ValorAnt
                    GoTo salte5
                End If
                
                If (yea < 0) Or (yea > empresa.ao) Then
                    MsgBox "Año invalido "
                    Tex_Edicion = ValorAnt
                    ListPer.Text = ValorAnt
                    GoTo salte5
                End If
            
            End If
            personal.fal = ListPer.Text
            Put 1, rgtro, personal
salte5:
        Case 6
            If ListPer.Text <> "" Then
                Mes = Mid(ListPer.Text, 4, 2)
                yea = Mid(ListPer.Text, 7, 4)
                dia = Left(ListPer.Text, 2)
                
                If (Mes < 1) Or (Mes > 12) Then
                    MsgBox "Mes invalido "
                    Tex_Edicion = ValorAnt
                    ListPer.Text = ValorAnt
                    GoTo salte6
                End If
                
                If (dia < 1) Or (dia > dd(Mes)) Then
                    MsgBox "dia invalido "
                    Tex_Edicion = ValorAnt
                    ListPer.Text = ValorAnt
                    GoTo salte6
                End If
                
                If (yea < 1900) Or (yea > empresa.ao) Then
                    MsgBox "Año invalido "
                    Tex_Edicion = ValorAnt
                    ListPer.Text = ValorAnt
                    GoTo salte6
                End If
            End If
            
            personal.fab = ListPer.Text
            Put 1, rgtro, personal
salte6:
        Case 7
            If ListPer.Text = "" Then ListPer.Text = 0
                personal.ingr = ListPer.Text
                ListPer.Text = Format(ListPer.Text, z2$)
                If personal.ingr > 0 Then
                yea = Mid(personal.fal, 7, 4)
                If yea < 1900 Then
                    MsgBox "La fecha de ingreso no es correcta, no es posible calcular el salario integrado"
                Else
                    antig = CalcularAntiguedad(personal.fal)
                    facto = 0
                    factor antig, facto
                    personal.integrado = (personal.ingr * facto) + personal.viat + personal.otras
                    If personal.integrado > (empresa.sm * 25) Then personal.integrado = (empresa.sm * 25)
                        ListPer.TextMatrix(ListPer.Row, 10) = Format(personal.integrado, z2$)
                    End If
            End If
            Put 1, rgtro, personal
        Case 8
            If ListPer.Text = "" Then ListPer.Text = 0
                personal.viat = ListPer.Text
                ListPer.Text = Format(ListPer.Text, z2$)
                yea = Mid(personal.fal, 7, 4)
            
                If yea < 1900 Then
                    MsgBox "La fecha de ingreso no es correcta, no es posible calcular el salario integrado"
                Else
                    antig = CalcularAntiguedad(personal.fal)
                    facto = 0
                    factor antig, facto
                    personal.integrado = (personal.ingr * facto) + personal.viat + personal.otras
                    If personal.integrado > (empresa.sm * 25) Then personal.integrado = (empresa.sm * 25)
                    ListPer.TextMatrix(ListPer.Row, 10) = Format(personal.integrado, z2$)
                End If
                
            Put 1, rgtro, personal
            
        Case 9
            If ListPer.Text = "" Then ListPer.Text = 0
            personal.otras = ListPer.Text
            ListPer.Text = Format(ListPer.Text, z2$)
            yea = Mid(personal.fal, 7, 4)
            
            If yea < 1900 Then
                MsgBox "La fecha de ingreso no es correcta, no es posible calcular el salario integrado"
            Else
                antig = CalcularAntiguedad(personal.fal)
                facto = 0
                factor antig, facto
                personal.integrado = (personal.ingr * facto) + personal.viat + personal.otras
                If personal.integrado > (empresa.sm * 25) Then personal.integrado = (empresa.sm * 25)
                ListPer.TextMatrix(ListPer.Row, 10) = Format(personal.integrado, z2$)
            End If
            
            Put 1, rgtro, personal
        Case 10
            MsgBox "No es posible modificar el salario Integrado "
            Tex_Edicion = ValorAnt
            ListPer.Text = ValorAnt
        End Select
End Sub

' ============================================================
' IMPRESION DE LISTADOS
' ============================================================

Sub pertit()
    fontviejo = Printer.FontSize
    Printer.FontSize = 10
    Printer.FontBold = True
    ancho1 = Int(Printer.TextWidth(empresa.name) / 2)
    Printer.CurrentX = (65 * 120) - ancho1
    Printer.Print empresa.name;
    Printer.Print
    titulo$ = "LISTADO DE PERSONAL " + LTrim$(Str$(empresa.ao))
    ancho1 = Int(Printer.TextWidth(titulo$) / 2)
    ancho2 = (55 * 120) - ancho1
    Printer.CurrentX = ancho2
    Printer.Print titulo$
    Printer.Print
    Printer.Line (0, Printer.CurrentY)-((130 * 120), Printer.CurrentY + 50), , BF
    Printer.FontSize = fontviejo
    Printer.Print
    Printer.CurrentX = (120 * 3)
    Printer.Print "No.";
    Printer.CurrentX = (120 * 10)
    Printer.Print " N o m b r e";
    Printer.CurrentX = (120 * 32)
    Printer.Print "R.F.C.";
    Printer.CurrentX = (120 * 42)
    Printer.Print " I.M.S.S.";
    Printer.CurrentX = (120 * 50)
    Printer.Print "F.Alta";
    Printer.CurrentX = (120 * 59)
    Printer.Print "F.baja";
    Printer.CurrentX = (120 * 68)
    Printer.Print "Diario";
    Printer.CurrentX = (120 * 76)
    Printer.Print "O.F.";
    Printer.CurrentX = (120 * 84)
    Printer.Print "Otros";
    Printer.CurrentX = (120 * 90)
    Printer.Print "Integrado"
    Printer.Print
    Printer.Line (0, Printer.CurrentY)-((130 * 120), Printer.CurrentY + 50), , BF
    Printer.FontBold = False
    Printer.Print
    Printer.Print
    Printer.Print
End Sub
Sub linea()
        apelativo1$ = RTrim$(personal.ape1) + " " + RTrim$(personal.ape2) + " " + RTrim$(personal.nom)
        Printer.CurrentX = (120 * 7)
        Printer.Print apelativo1$;
        Printer.CurrentX = (120 * 30)
        Printer.Print personal.RFC;
        Printer.CurrentX = (120 * 40)
        Printer.Print personal.imss;
        Printer.CurrentX = (120 * 50)
        Printer.Print personal.fal;
        Printer.CurrentX = (120 * 58)
        Printer.Print personal.fab;

        valor$ = Format(personal.ingr, z2$): uso$ = z2$
        pone = 0: colocar pone, valor$, uso$
        Printer.CurrentX = (120 * 68) + pone
        Printer.Print valor$;

        If personal.viat > o Then
            valor$ = Format(personal.viat, z2$)
            pone = 0: colocar pone, valor$, uso$
            Printer.CurrentX = (120 * 76) + pone
            Printer.Print valor$;
        End If
        
        If personal.otras > 0 Then
            valor$ = Format(personal.otras, z2$)
            pone = 0: colocar pone, valor$, uso$
            Printer.CurrentX = (120 * 84) + pone
            Printer.Print valor$;
        End If
        
        valor$ = Format(personal.integrado, z2$)
        pone = 0: colocar pone, valor$, uso$
        Printer.CurrentX = (120 * 90) + pone
        Printer.Print valor$
proximo:
End Sub


' ============================================================
' ORDENAMIENTO DEL GRID
' ============================================================

Private Sub ArOrAlf_Click()
    colanti = ListPer.col
    renati = ListPer.Row
    ListPer.Row = 1
    ListPer.col = 1
    ListPer.RowSel = ListPer.Rows - 1
    ListPer.Sort = 1
    ListPer.col = colanti
    ListPer.Row = renati

End Sub

Private Sub ArOrNum_Click()
    colanti = ListPer.col
    renati = ListPer.Row
    ListPer.Row = 1
    ListPer.col = 0
    ListPer.RowSel = ListPer.Rows - 1
    ListPer.Sort = 3
    ListPer.col = colanti
    ListPer.Row = renati
    ListPer.SetFocus
End Sub


' ============================================================
' MODO EDICION / RESPALDO Y CANCELACION
' ============================================================

Private Sub BtnCancelarCambios_Click()
    If MsgBox("Esto descartará TODOS los cambios hechos desde que habilitaste edición. ¿Continuar?", vbYesNo + vbExclamation) = vbYes Then
        RestaurarArchivos
        ModoEdicion = False
        ActualizarBotones
        CargarGrid   ' releer y repintar desde los archivos restaurados
    End If
End Sub

Private Sub BtnGuardarCambios_Click()
    If MsgBox("¿Confirmar y salir del modo edición?", vbYesNo + vbQuestion) = vbYes Then
        ModoEdicion = False
        ActualizarBotones
        MsgBox "Cambios confirmados.", vbInformation
    End If
End Sub

Private Sub BtnHabilitarEdicion_Click()
    RespaldarArchivos
    ModoEdicion = True
    ActualizarBotones
    MsgBox "Edición habilitada. Los cambios se guardan al vuelo por celda, pero ahora puedes cancelar todo con 'Cancelar cambios'.", vbInformation
End Sub


' ============================================================
' CALCULO MASIVO DE SALARIO INTEGRADO
' ============================================================

Private Sub calcularIntegrado_Click()
Dim i As Integer
Dim idEmp As Integer
Dim lAntig As Integer
Dim dFacto As Double
Dim cSdi As Currency
Dim empFile As String
Dim empRecord As Integer
Dim fNum As Integer
Dim file2 As Integer
Dim dbFile As String

On Error GoTo ErrHandler

    ' 1. Asegurar la lectura del UMA actual
    empFile = "empresa.dno"
    If Dir(empFile) <> "" Then
        fNum = FreeFile
        Open empFile For Random As #fNum Len = Len(empresa)
        empRecord = LOF(fNum) / Len(empresa)
        If empRecord > 0 Then
            Get #fNum, empRecord, empresa
        End If
        Close #fNum
    End If

    ' 2. Abrir archivo de personal
    dbFile = "personal.dno"
    If Dir(dbFile) = "" Then
        MsgBox "No se encontrÃ³ el archivo de personal.", vbCritical, "Error"
        Exit Sub
    End If
    
    file2 = FreeFile
    Open dbFile For Random As #file2 Len = Len(personal)
    
    ' 3. Procesar filas de la cuadrÃ­cula
    For i = 1 To ListPer.Rows - 1
        If IsNumeric(ListPer.TextMatrix(i, 0)) Then
            idEmp = Val(ListPer.TextMatrix(i, 0))
            If idEmp > 0 Then
                Get #file2, idEmp, personal
                
                lAntig = CalcularAntiguedad(personal.fal)
                dFacto = 0
                factor lAntig, dFacto
                
                cSdi = (personal.ingr + personal.viat + personal.otras) * dFacto
                If empresa.sm > 0 Then
                    If cSdi > (empresa.sm * 25) Then cSdi = (empresa.sm * 25)
                End If
                
                personal.integrado = cSdi
                Put #file2, idEmp, personal
                
                ListPer.TextMatrix(i, 10) = Format(personal.integrado, z2)
            End If
        End If
    Next i
    
    Close #file2
    MsgBox "CÃ¡lculo de Salario Integrado completado con Ã©xito para todos los registros.", vbInformation, "Proceso Terminado"
    Exit Sub

ErrHandler:
    If file2 > 0 Then Close #file2
    MsgBox "Error en calcularIntegrado: " & Err.Description, vbCritical, "Error"
End Sub


' ============================================================
' COMANDOS GENERALES / COPIAR / PEGAR / SALIR
' ============================================================

Private Sub Command1_Click()
Form_Load
End Sub

Private Sub EdCopiar_Click()
   Dim Temporal1
   Clipboard.Clear
   
    For i = ListPer.Row To ListPer.RowSel
        For f = ListPer.col To ListPer.ColSel
            Temporal1 = Temporal1 + ListPer.TextMatrix(i, f)
            If f < ListPer.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
        Next f
        Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
    Next i
    
    Clipboard.SetText Temporal1
    ListPer.FixedCols = 2
    ListPer.FixedRows = 1
    ListPer.Row = 1: ListPer.col = 2
    TexEdicion.Text = ListPer.Text

End Sub

Private Sub edimalf_Click(Index As Integer)
    Close 2
    Open "personal.dno" For Random As 2 Len = Len(personal)
    Dm = LOF(2) / Len(personal)
    Close 8
    Open "MAESTRO" For Random As 8 Len = Len(maestro)
    ddm = LOF(8) / Len(maestro)
    
    If Dm > 0 Then
        Printer.FontName = "Courier New"
        Printer.FontSize = 6
        Printer.Orientation = 1
        pertit
        For r = 1 To (ListPer.Rows - 1)
            rgtro = ListPer.TextMatrix(r, 0)
            valor = rgtro: uso$ = "####0.": ancho2 = 0
            colocar ancho2, valor, uso$
            Printer.CurrentX = 0 + ancho2
            Printer.Print rgtro;
            Get 2, rgtro, personal
            linea
            contador = contador + 1
            If contador > 83 Then
                contador = 0
                Printer.NewPage
                pertit
            End If
        Next r
        Printer.EndDoc
    End If
    Close 2
End Sub

Private Sub edimnum_Click(Index As Integer)
     Close 2, 8
     Open "personal.dno" For Random As 2 Len = Len(personal)
     Dm = LOF(2) / Len(personal)
     Open "maestro" For Random As 8 Len = Len(maestro)
     ddm = LOF(8) / Len(maestro)
 If Dm > 0 Then
     Printer.FontName = "Courier New"
     Printer.FontSize = 6
     Printer.Orientation = 1
     pertit
     For r = 1 To Dm
        rgtro = r
        valor = r: uso$ = "####0.": ancho2 = 0
        colocar ancho2, valor, uso$
        Printer.CurrentX = 0 + ancho2
        Printer.Print valor;
        Get 2, r, personal
        linea
        contador = contador + 1
        If contador > 83 Then
           contador = 0
           Printer.NewPage
           pertit
        End If
     Next r
     Printer.EndDoc
  End If
  Close 2
    
End Sub

Private Sub editsale_Click(Index As Integer)
 Unload Form4
 Form1.Show
End Sub

Private Sub EditSelTot_Click()
    Clipboard.Clear
    ListPer.Row = 0
    ListPer.col = 0
    ListPer.RowSel = ListPer.Rows - 1
    ListPer.ColSel = ListPer.Cols - 1
End Sub

Private Sub EdPegar_Click()
    If Not ModoEdicion Then
      MsgBox "Debes pulsar 'Habilitar edición' antes de pegar datos.", vbExclamation
      Exit Sub
  End If
  Dim temporal, DeAqui As Integer, RetornoCarro As Long, InicioCopia As Long
  temporal = Clipboard.GetText(vbCFText)
  RetornoCarro = ListPer.col
  InicioCopia = ListPer.Row
If temporal <> "" Then
  Clipboard.Clear
  DeAqui = 1
For i = 1 To Len(temporal)
    Select Case Mid(temporal, i, 1)
          Case Chr(9)
          TexEdicion.Text = Mid(temporal, DeAqui, (i - DeAqui))
          ListPer.Text = Mid(temporal, DeAqui, (i - DeAqui))
          ListPer.col = ListPer.col + 1
          DeAqui = i + 1
          Case Chr(13)
          TexEdicion.Text = Mid(temporal, DeAqui, (i - DeAqui))
          ListPer.Text = Mid(temporal, DeAqui, (i - DeAqui))
          ArchCorr
          ListPer.Row = ListPer.Row + 1
          DeAqui = i + 1
          Case Chr(10)
          ListPer.col = RetornoCarro
          DeAqui = i + 1
          Case Else
          Rem nada
    End Select
 Next i
 ListPer.Row = InicioCopia: ListPer.col = RetornoCarro
End If
End Sub

Private Sub EdSup_Click()
 ListPer.FixedCols = 0
 ListPer.FixedRows = 0
End Sub


' ============================================================
' CARGA INICIAL DEL FORMULARIO
' ============================================================

Private Sub Form_Load()
    Dim empFile As String
    Dim empRecord As Long
    Dim fNum As Integer
    Dim detalleError As String

    On Error GoTo ErrorCarga

    empFile = "empresa.dno"
    If Dir$(empFile) <> "" Then
        fNum = FreeFile
        Open empFile For Random Access Read As #fNum Len = Len(empresa)
        empRecord = LOF(fNum) \ Len(empresa)
        If empRecord > 0 Then Get #fNum, empRecord, empresa
        Close #fNum
        fNum = 0
    End If

    'NO consultar ni escribir SQL en Form4.
    'Los datos llegaron previamente a los .dno mediante PeCfdi.
    z2$ = "#,###,##0.0000"
    ModoEdicion = False
    CargarGrid
    ActualizarBotones
    Exit Sub

ErrorCarga:
    detalleError = Err.Description
    On Error Resume Next
    If fNum > 0 Then Close #fNum
    On Error GoTo 0
    MsgBox "No fue posible abrir los registros locales: " & _
           detalleError, vbExclamation, "Edicion de personal"
End Sub


' ============================================================
' REPOSICION Y EVENTOS DEL GRID
' ============================================================

Sub REPONE()
Dim Repuse
    Get 2, r, personal
    Get 3, r, Otros_Rgtros
    Get 4, r, Clbnx
    
    Repuse = Format(r, "####0") _
        & Chr(9) & apelativo$ _
        & Chr(9) & RTrim(personal.RFC) _
        & Chr(9) & RTrim(Otros_Rgtros.curp) _
        & Chr(9) & (" " + Replace(RTrim(personal.imss), "-", "")) _
        & Chr(9) & RTrim(personal.fal) _
        & Chr(9) & RTrim(personal.fab) _
        & Chr(9) & Format(personal.ingr, z2) _
        & Chr(9) & Format(personal.viat, z2) _
        & Chr(9) & Format(personal.otras, z2) _
        & Chr(9) & Format(personal.integrado, z2) _
        & Chr(9) & Format(Clbnx.Q1, "<")

End Sub
Private Sub list1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        rgtro = ListPer.TextMatrix(ListPer.Row, 0)
        Load Form2
        Form2.IniciarConsulta CInt(rgtro)   ' reemplaza a Form2.cargarEmpleado (rgtro)
        Form2.Show 1
     End If
End Sub

Private Sub Form_Resize()
On Error GoTo ErrorManejador
    ListPer.Width = ScaleWidth - 200
    ListPer.Height = Form4.Height * 0.85
Exit Sub

ErrorManejador:

End Sub

Private Sub listper_EnterCell()
    If ListPer.col > 1 And ListPer.Row > 0 Then
        ListPer.CellBackColor = &H80FF80
    End If
End Sub

Private Sub listper_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
            Case vbKeyDelete
                Tex_Edicion = ListPer.Text
                ArchCorr
            Case vbKeyF2
                TexEdicion.Text = ListPer.Text
                TexEdicion.SetFocus
       End Select
End Sub

Private Sub ListPer_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
         Case 13:
            rgtro = ListPer.TextMatrix(ListPer.Row, 0)
            Load Form2
            Form2.IniciarConsulta CInt(rgtro)   ' reemplaza a Form2.cargarEmpleado (rgtro)
            Form2.Show 1
            Close 2
            Open "personal.dno" For Random As 2 Len = Len(personal)
            Get 2, rgtro, personal
            Close 3
            Open "PerOtre.dno" For Random As 3 Len = Len(Otros_Rgtros)
            dmper = LOF(3) / Len(personal)
            RecPra
            TexEdicion.Text = ListPer.Text
            Close 2, 3
         Case 27
         Case Else
     End Select
End Sub

Private Sub listper_LeaveCell()
    If ListPer.col > 1 And ListPer.Row > 0 Then
        ListPer.Text = UCase(TexEdicion.Text)
        ListPer.CellBackColor = vbWhite
    End If
End Sub

Private Sub listper_RowColChange()
    ValorAnt = ListPer.Text
    TexEdicion.Text = ListPer.Text
End Sub


Private Sub saldosSIA_Click()
    actualizarSaldos
End Sub

Private Sub sinFiltroCarga_Click()
    cargarEmpleadosSinFiltro
End Sub

Private Sub TexEdicion_Change()
    ListPer.Text = TexEdicion.Text
End Sub

Private Sub TexEdicion_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            TexEdicion.Text = UCase(TexEdicion.Text)
            ListPer.Text = TexEdicion.Text
            ArchCorr
            ListPer.SetFocus
       Case 27
            TexEdicion.Text = ListPer.Text
    End Select
End Sub


' ============================================================
' SINCRONIZACION CON BASE DE DATOS REMOTA
' ============================================================

Public Sub actualizarSaldos()
    'Operacion SQL deshabilitada: la BD es fuente de SOLO LECTURA.
    MsgBox "La actualizacion de saldos en SQL esta deshabilitada." & _
           vbCrLf & "La informacion principal permanece en archivos .dno.", _
           vbInformation, "Solo lectura SQL"
End Sub


Private Sub cargarEmpleadosSinFiltro()
    'Mostrar la informacion local SIN importar ni sobreescribir desde SQL.
    CargarGrid
End Sub


' ============================================================
' CALCULO DE IMPUESTO ISR
' ============================================================

Public Sub calcularYAsignarImpuesto()
    Dim i As Integer
    Dim ingreso As Currency
    Dim viaticos As Currency
    Dim otras As Currency
    Dim ingresoTotal As Currency
    Dim largoArticulo As Double
    Dim filePath As String

    On Error GoTo ErrorHandler

    For i = 1 To ListPer.Rows - 1
        ' Verificar que las celdas no estén vacías antes de intentar convertirlas
        If Trim(ListPer.TextMatrix(i, 7)) <> "" And Trim(ListPer.TextMatrix(i, 8)) <> "" And Trim(ListPer.TextMatrix(i, 9)) <> "" Then
            ' Convertir los valores de las celdas a Currency
            ingreso = CCur(ListPer.TextMatrix(i, 7))
            viaticos = CCur(ListPer.TextMatrix(i, 8))
            otras = CCur(ListPer.TextMatrix(i, 9))

            ' Calcular el ingreso total
            ingresoTotal = ingreso + viaticos + otras
            ingresoTotal = ingresoTotal * 30

            filePath = Trim(Dir_imptos) & "Tab08Mes.ISR"

            ' Abrir archivo de impuestos
            Open filePath For Random As #99 Len = Len(articulo)
            largoArticulo = LOF(99) / Len(articulo)

            ' Calcular el impuesto
            impuestoPersonal = 0 ' Reiniciar impuestoPersonal para cada iteración
            For j = 1 To largoArticulo
                Get #99, j, articulo
                If ingresoTotal >= articulo.liminf And ingresoTotal <= articulo.limsup Then
                    impuestoPersonal = (((ingresoTotal - articulo.liminf) * articulo.porcsl / 100) + articulo.cuotaf)
                    Exit For
                End If
            Next j

            Close #99

            ' Asignar el resultado a la matriz, asegurando que el índice sea correcto
            ListPer.TextMatrix(i, 12) = Format(impuestoPersonal, z1$)
        End If
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "Error al calcular el impuesto: " & Err.Description
End Sub

Private Sub isrCalculo_Click()
    calcularYAsignarImpuesto
End Sub


' ============================================================
' FILTRO DE BUSQUEDA
' ============================================================

Private Sub Text1_KeyPress(KeyAscii As Integer)
    ' 1. Solo se activa si la tecla presionada es Enter
    
    
    If KeyAscii = 13 Then
        
        Dim i As Long
        Dim filtro As String
        Dim NombreEnGrid As String
        
        ' 2. Prepara el texto del filtro
        filtro = UCase(Trim(Text1.Text))
        
        ' 3. Recorre las filas de la tabla (esto es súper rápido)
        For i = 1 To ListPer.Rows - 1 ' Empieza en 1 para no tocar la cabecera
            
            ' Toma el nombre de la columna 1 de la fila actual
            NombreEnGrid = UCase(ListPer.TextMatrix(i, 1))
            
            ' 4. Compara si el nombre CONTIENE el texto del filtro
            If InStr(1, NombreEnGrid, filtro) > 0 Then
                ' Si lo contiene, se asegura de que la fila sea visible
                ListPer.RowHeight(i) = -1 ' -1 es la altura por defecto
            Else
                ' Si NO lo contiene, OCULTA la fila poniendo su altura en 0
                ListPer.RowHeight(i) = 0
            End If
            
        Next i
        
    End If
End Sub


' ============================================================
' ARCHIVOS DE RESPALDO Y ESTADO DE BOTONES
' ============================================================

Private Sub RespaldarArchivos()
    Close 1, 2, 3, 4, 15   ' asegura que nada quede abierto
    FileCopy "personal.dno", "personal.bak"
    FileCopy "PerOtre.dno", "PerOtre.bak"
    If Dir("Bnxcla.dno") <> "" Then FileCopy "Bnxcla.dno", "Bnxcla.bak"
    If Dir("deon.dno") <> "" Then FileCopy "deon.dno", "deon.bak"
End Sub

Private Sub RestaurarArchivos()
    Close 1, 2, 3, 4, 15
    FileCopy "personal.bak", "personal.dno"
    FileCopy "PerOtre.bak", "PerOtre.dno"
    If Dir("Bnxcla.bak") <> "" Then FileCopy "Bnxcla.bak", "Bnxcla.dno"
    If Dir("deon.bak") <> "" Then FileCopy "deon.bak", "deon.dno"
End Sub

Private Sub ActualizarBotones()
    BtnHabilitarEdicion.Enabled = Not ModoEdicion
    BtnGuardarCambios.Enabled = ModoEdicion
    BtnCancelarCambios.Enabled = ModoEdicion
    Form4.Caption = IIf(ModoEdicion, "Edición de personal - MODO EDICIÓN", "Edición de personal - Solo lectura")
End Sub


' ============================================================
' CARGA DEL GRID DESDE ARCHIVOS LOCALES
' ============================================================

Private Sub CargarGrid()

    Close 2, 3, 4, 15

    Open "personal.dno" For Random As 2 Len = Len(personal)
    Dm = LOF(2) / Len(personal)
    Open "PerOtre.dno" For Random As 3 Len = Len(Otros_Rgtros)
    dmper = LOF(3) / Len(Otros_Rgtros)
    Open "Bnxcla.dno" For Random As 4 Len = Len(Clbnx)   ' <- estaba comentada; sin esto, Get/Put 4 fallan
    Open "deon.dno" For Random As 15 Len = Len(DEON)
    largoDeon = LOF(15) / Len(DEON)

    If dmper < Dm Then
        If dmper < 1 Then dmper = 1
        For r = (dmper + 1) To Dm: Get 3, r, Otros_Rgtros
            Otros_Rgtros.curp = "": Otros_Rgtros.otra = ""
            Otros_Rgtros.yotra = "": Otros_Rgtros.yporsi = ""
            Put 3, r, Otros_Rgtros
        Next r
    End If

    ListPer.Cols = 13: ListPer.Rows = 1: ListPer.Row = 0
    ListPer.col = 0: ListPer.CellAlignment = 4: ListPer.ColWidth(0) = 400: ListPer.Text = "#"
    ListPer.col = 1: ListPer.CellAlignment = 4: ListPer.ColWidth(1) = 3200: ListPer.Text = "Nombre"
    ListPer.col = 2: ListPer.CellAlignment = 4: ListPer.ColWidth(2) = 2200: ListPer.Text = "RFC"
    ListPer.col = 3: ListPer.CellAlignment = 4: ListPer.ColWidth(3) = 2200: ListPer.Text = "CURP"
    ListPer.col = 4: ListPer.CellAlignment = 4: ListPer.ColWidth(4) = 1600: ListPer.Text = "IMSS"
    ListPer.col = 5: ListPer.CellAlignment = 4: ListPer.ColWidth(5) = 1200: ListPer.Text = "Fcha.Alta"
    ListPer.col = 6: ListPer.CellAlignment = 4: ListPer.ColWidth(6) = 1200: ListPer.Text = "Fcha.Baja"
    ListPer.col = 7: ListPer.CellAlignment = 4: ListPer.ColWidth(7) = 1200: ListPer.Text = "Salario dia"
    ListPer.col = 8: ListPer.CellAlignment = 4: ListPer.ColWidth(8) = 1200: ListPer.Text = "O.F. dia"
    ListPer.col = 9: ListPer.CellAlignment = 4: ListPer.ColWidth(9) = 1200: ListPer.Text = "Otros diario"
    ListPer.col = 10: ListPer.CellAlignment = 4: ListPer.ColWidth(10) = 1200: ListPer.Text = "Integrado"
    ListPer.col = 11: ListPer.CellAlignment = 4: ListPer.ColWidth(11) = 2200: ListPer.Text = "Tarjeta Bnx"
    ListPer.col = 12: ListPer.CellAlignment = 4: ListPer.ColWidth(12) = 2200: ListPer.Text = "Impuesto (ISR)"

    If Dm > 0 Then
        For r = 1 To Dm: Get 2, r, personal: Get 3, r, Otros_Rgtros: Get 4, r, Clbnx: Get 15, r, DEON
             apelativo1$ = "": abaja = Val(Mid(personal.fab, 7, 4))
             If (abaja > 0) And (abaja < empresa.ao - 1) Then GoTo Sig_te
             If Len(Trim$(Replace(personal.nom, vbNullChar, ""))) > 0 Or _
                Len(Trim$(Replace(personal.ape1, vbNullChar, ""))) > 0 Or _
                Len(Trim$(Replace(personal.ape2, vbNullChar, ""))) > 0 Then
                apelativo1$ = RTrim$(personal.ape1) + " " + RTrim$(personal.ape2) + " " + RTrim$(personal.nom)
                apelativo$ = Left(apelativo1$, 59) + String$(60 - Len(Left(apelativo1$, 59)), " ")

                Dim lAntig As Integer, dFacto As Double, cSdi As Currency
                'Un empleado recien importado no trae fecha de alta
                'de nomina; no intentar calcular antiguedad en ese caso.
                cSdi = personal.integrado
                If IsDate(Trim$(Replace(personal.fal, vbNullChar, ""))) Then
                    lAntig = CalcularAntiguedad(personal.fal)
                    dFacto = 0
                    factor lAntig, dFacto
                    cSdi = (personal.ingr + personal.viat + personal.otras) * dFacto
                    If empresa.sm > 0 Then
                        If cSdi > (empresa.sm * 25) Then cSdi = (empresa.sm * 25)
                    End If
                End If

                If personal.integrado <> cSdi Then
                    personal.integrado = cSdi   ' solo en memoria, para mostrarlo correcto en la grilla
                End If

                ListPer.AddItem Format(r, "####0") _
                    & Chr(9) & apelativo$ _
                    & Chr(9) & RTrim(personal.RFC) _
                    & Chr(9) & RTrim(Otros_Rgtros.curp) _
                    & Chr(9) & (" " + Replace(RTrim(personal.imss), "-", "")) _
                    & Chr(9) & RTrim(personal.fal) _
                    & Chr(9) & RTrim(personal.fab) _
                    & Chr(9) & Format(personal.ingr, z2) _
                    & Chr(9) & Format(personal.viat, z2) _
                    & Chr(9) & Format(personal.otras, z2) _
                    & Chr(9) & Format(personal.integrado, z2) _
                    & Chr(9) & Format(Clbnx.Q1, "<")
             End If
Sig_te:
        Next r

        Close 2, 3

    Else
        MsgBox "No existe personal para la edicion "
        Close 2, 3
        Load Form1
        Form1.Show
    End If

End Sub
' comentario




