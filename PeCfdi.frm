VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PeCfdi 
   Caption         =   "Datos complementarios de Personal"
   ClientHeight    =   5940
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   9555
   LinkTopic       =   "Form9"
   ScaleHeight     =   5940
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSFlexGridLib.MSFlexGrid dat 
      Height          =   2175
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3836
      _Version        =   393216
   End
   Begin VB.TextBox LTXT 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7092
   End
   Begin MSFlexGridLib.MSFlexGrid PCpCfdi 
      Height          =   4932
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   9012
      _ExtentX        =   15901
      _ExtentY        =   8705
      _Version        =   393216
   End
   Begin VB.Menu ArCf 
      Caption         =   "&Archivo"
      Begin VB.Menu DB 
         Caption         =   "&DB"
      End
      Begin VB.Menu ArCfAbr 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu ArCfGuarda 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu ArCfsep1 
         Caption         =   "-"
      End
      Begin VB.Menu ArCfSale 
         Caption         =   "&Salida"
      End
   End
   Begin VB.Menu EdCf 
      Caption         =   "&Editar"
      Begin VB.Menu EdCfPegar 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu EdCfCopia 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu EdCfSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EdCfInsertar 
         Caption         =   "&Insertar"
      End
      Begin VB.Menu EdCfEliminar 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu EdCfSep2 
         Caption         =   "-"
      End
      Begin VB.Menu EdCfSelT 
         Caption         =   "&Seleccionar Todo"
      End
   End
End
Attribute VB_Name = "PeCfdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Nueva Versión

Dim temporal, Temporal1

' ============================================================
' CONFIGURACION DEL GRID
' ============================================================

Sub Col_Def()
    PCpCfdi.Clear
    PCpCfdi.Rows = 200
    PCpCfdi.Cols = 9
    PCpCfdi.FixedCols = 0
    PCpCfdi.FixedRows = 1
    PCpCfdi.Row = 0

    PCpCfdi.Col = 0
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(0) = 800
    PCpCfdi.Text = "Num"

    PCpCfdi.Col = 1
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(1) = 3600
    PCpCfdi.Text = "Nombre"

    PCpCfdi.Col = 2
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(2) = 2400
    PCpCfdi.Text = "Direccion"

    PCpCfdi.Col = 3
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(3) = 2400
    PCpCfdi.Text = "Colonia"

    PCpCfdi.Col = 4
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(4) = 2400
    PCpCfdi.Text = "Ciudad"

    PCpCfdi.Col = 5
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(5) = 2400
    PCpCfdi.Text = "Estado"

    PCpCfdi.Col = 6
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(6) = 2400
    PCpCfdi.Text = "Delegacion"

    PCpCfdi.Col = 7
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(7) = 800
    PCpCfdi.Text = "Codigo"

    PCpCfdi.Col = 8
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(8) = 2400
    PCpCfdi.Text = "Correo "
End Sub

Private Sub Form_Resize()
    PCpCfdi.Height = PeCfdi.Height * 0.85
    PCpCfdi.Width = PeCfdi.Width * 0.95
End Sub

' ============================================================
' CARGA DEL FORMULARIO / SINCRONIZACION INICIAL
' ============================================================


Private Sub Form_Load()

    Col_Def

    If EsEmpresaLocal(emp) Then
        PeCfdi.Caption = "Datos complementarios de Personal - Modo local"
    Else
        PeCfdi.Caption = "Datos complementarios de Personal - Presione DB para sincronizar"
    End If

End Sub


Private Function ObtenerEmpresaSAT()
    Dim abrEmpresa

    abrEmpresa = Left(Trim(emp), 4)

    If (UCase(abrEmpresa) = "SACM") Then abrEmpresa = "SACMAG"
    If (UCase(abrEmpresa) = "SUPT") Then abrEmpresa = "SUPTER"
    If (UCase(abrEmpresa) = "COOR") Then abrEmpresa = "CORDINA"
    If (UCase(abrEmpresa) = "EPES") Then abrEmpresa = "EPESA"
    If (UCase(abrEmpresa) = "SUPE") Then abrEmpresa = "SUPERVISA"

    ObtenerEmpresaSAT = abrEmpresa
End Function

Private Sub ConfigurarGridDatosSat(ByVal oRS As ADODB.Recordset)
    dat.Rows = oRS.RecordCount + 1
    dat.Cols = oRS.Fields.Count
    dat.Row = 0
    dat.Col = 0

    dat.Col = 0
    dat.CellAlignment = 4
    dat.ColWidth(0) = 2800
    dat.Text = "ID NÓMINA"

    dat.Col = 1
    dat.CellAlignment = 4
    dat.ColWidth(1) = 2800
    dat.Text = "NOMBRE"

    dat.Col = 2
    dat.CellAlignment = 4
    dat.ColWidth(2) = 2800
    dat.Text = "APELLIDO P"

    dat.Col = 3
    dat.CellAlignment = 4
    dat.ColWidth(3) = 2800
    dat.Text = "APELLIDO M"

    dat.Col = 4
    dat.CellAlignment = 4
    dat.ColWidth(4) = 2800
    dat.Text = "DIRECCIÓN"

    dat.Col = 5
    dat.CellAlignment = 4
    dat.ColWidth(5) = 2800
    dat.Text = "COLONIA"

    dat.Col = 6
    dat.CellAlignment = 4
    dat.ColWidth(6) = 2800
    dat.Text = "CIUDAD"

    dat.Col = 7
    dat.CellAlignment = 4
    dat.ColWidth(7) = 2800
    dat.Text = "ESTADO"

    dat.Col = 8
    dat.CellAlignment = 4
    dat.ColWidth(8) = 2800
    dat.Text = "DELEGACIÓN"

    dat.Col = 9
    dat.CellAlignment = 4
    dat.ColWidth(9) = 2800
    dat.Text = "CÓDIGO POSTAL"

    dat.Col = 10
    dat.CellAlignment = 4
    dat.ColWidth(10) = 2800
    dat.Text = "CORREO"
End Sub

Private Sub CargarDatosSatEnGrid(ByVal oRS As ADODB.Recordset)
    vardatarows = oRS.GetRows()

    For i = 1 To dat.Rows - 1
        For h = 0 To dat.Cols - 1
            If (IsNull(vardatarows(h, i - 1))) Then
                dat.TextMatrix(i, h) = "N/A"
            Else
                dat.TextMatrix(i, h) = Trim(vardatarows(h, i - 1))
            End If
        Next h
    Next i
End Sub

Private Sub GuardarDatosSatEnArchivoLocal()
    Close 1, 2

    Open "personal.dno" For Random As 2 Len = Len(personal)
    Dm = LOF(2) / Len(personal)

    Open "Perscfdi.dno" For Random As 1 Len = Len(Empleado_1)

    For i = 1 To dat.Rows - 1
        idNomina = dat.TextMatrix(i, 0)

        Get 1, CInt(idNomina), Empleado_1

        Empleado_1.Direccion = UCase(Trim(dat.TextMatrix(i, 4)))
        Empleado_1.Colonia = UCase(Trim(dat.TextMatrix(i, 5)))
        Empleado_1.Ciudad = UCase(Trim(dat.TextMatrix(i, 6)))
        Empleado_1.Estado = UCase(Trim(dat.TextMatrix(i, 7)))
        Empleado_1.Delegacion = UCase(Trim(dat.TextMatrix(i, 8)))
        Empleado_1.Cpostal = UCase(Trim(dat.TextMatrix(i, 9)))
        Empleado_1.correo = UCase(Trim(dat.TextMatrix(i, 10)))

        Put 1, idNomina, Empleado_1
    Next i
End Sub

' ============================================================
' ARCHIVO LOCAL CFDI
' ============================================================

Private Sub ArCfAbr_Click()
    Dim EmpNum As Long
    Dim nombre As String
    Dim Cadena1

    PCpCfdi.Rows = 1

    Close 1, 2

    Open "Personal.dno" For Random As 2 Len = Len(personal)
    cm = LOF(2) / Len(personal)

    Open "Perscfdi.dno" For Random As 1 Len = Len(Empleado_1)

    For i = 1 To cm
        Get 2, i, personal
        Get 1, i, Empleado_1

        If Trim$(personal.nom) <> "" Then
            nombre = (Trim(personal.nom) + " " + Trim(personal.ape1) + " " + Trim(personal.ape2))
            Cadena1 = Format(i, "####0") & Chr(9) & nombre & Chr(9) & Empleado_1.Direccion _
                      & Chr(9) & Empleado_1.Colonia & Chr(9) & Empleado_1.Ciudad _
                      & Chr(9) & Empleado_1.Estado & Chr(9) & Empleado_1.Delegacion _
                      & Chr(9) & Empleado_1.Cpostal & Chr(9) & Empleado_1.correo

            PCpCfdi.AddItem Cadena1
        End If
    Next i
    
    Close #1, #2
End Sub

Private Sub ArCfGuarda_Click()
    Dim EmpNum As Long

    Close 1

    Open "Perscfdi.dno" For Random As 1 Len = Len(Empleado_1)
    cm = LOF(1) / Len(Empleado_1)

    For renglon = 1 To PCpCfdi.Rows - 1
        If IsNumeric(PCpCfdi.TextMatrix(renglon, 0)) Then
            EmpNum = PCpCfdi.TextMatrix(renglon, 0)
            Empleado_1.Direccion = PCpCfdi.TextMatrix(renglon, 2)
            Empleado_1.Colonia = PCpCfdi.TextMatrix(renglon, 3)
            Empleado_1.Ciudad = PCpCfdi.TextMatrix(renglon, 4)
            Empleado_1.Estado = PCpCfdi.TextMatrix(renglon, 5)
            Empleado_1.Delegacion = PCpCfdi.TextMatrix(renglon, 6)
            Empleado_1.Cpostal = PCpCfdi.TextMatrix(renglon, 7)
            Empleado_1.correo = PCpCfdi.TextMatrix(renglon, 8)

            Put 1, EmpNum, Empleado_1
        End If
    Next renglon

    Close 1
End Sub

Private Sub ArCfSale_Click()
    Close
    Unload PeCfdi
End Sub

' ============================================================
' CONSULTA REMOTA
' ============================================================



Private Sub DB_Click()

    Dim oRS As ADODB.Recordset
    Dim sSQL As String
    Dim abrEmpresa As String

    Dim carpeta As String
    Dim rutaPersonal As String
    Dim rutaCFDI As String
    Dim rutaTemporal As String
    Dim rutaRespaldo As String
    Dim rutaAnterior As String
    Dim sello As String
    Dim sufijo As String

    Dim totalPersonal As Long
    Dim totalSQL As Long
    Dim registrosGuardados As Long
    Dim sinDatosSQL As Long

    Dim id As Long
    Dim idTexto As String
    Dim i As Long
    Dim consecutivo As Long

    Dim vistos() As Boolean
    Dim tamRegistro As Long
    Dim tamPersonal As Long
    Dim tamNecesario As Long

    Dim existeCFDI As Boolean
    Dim archivoMovido As Boolean
    Dim archivoInstalado As Boolean
    Dim personalAbierto As Boolean
    Dim temporalAbierto As Boolean

    Dim hPersonal As Integer
    Dim hTemporal As Integer
    Dim hExtension As Integer

    Dim cero As Byte
    Dim numeroError As Long
    Dim descripcionError As String
    Dim errorRestauracion As String

    Dim nombreCompleto As String
    Dim cadenaGrid As String

    On Error GoTo ErrorDB

    '========================================================
    '1. VALIDAR EMPRESA
    '========================================================

    If EsEmpresaLocal(emp) Then
        MsgBox "Esta empresa trabaja en modo local." & vbCrLf & _
               "No se modificaron los archivos CFDI.", _
               vbInformation, "Sincronizacion"
        Exit Sub
    End If

    abrEmpresa = Trim$(ObtenerEmpresaSAT())

    If abrEmpresa = "" Then
        MsgBox "No fue posible identificar la empresa.", _
               vbExclamation, "Sincronizacion"
        Exit Sub
    End If

    '========================================================
    '2. VERIFICAR CONEXION SQL
    '========================================================

    If con Is Nothing Then
        MsgBox "La conexion SQL no esta inicializada.", _
               vbCritical, "Sincronizacion"
        Exit Sub
    End If

    If con.State <> adStateOpen Then
        MsgBox "No hay conexion con SQL Server." & vbCrLf & _
               "No se modificaron los archivos locales.", _
               vbExclamation, "Sincronizacion"
        Exit Sub
    End If

    '========================================================
    '3. IDENTIFICAR ARCHIVOS LOCALES
    '========================================================

    carpeta = CurDir$

    If Right$(carpeta, 1) <> "\" Then
        carpeta = carpeta & "\"
    End If

    rutaPersonal = carpeta & "Personal.dno"
    rutaCFDI = carpeta & "Perscfdi.dno"

    'Personal.dno debe existir. Nunca se crea aqui.
    If Dir$(rutaPersonal) = "" Then
        Err.Raise vbObjectError + 1001, , _
                  "No existe Personal.dno en: " & carpeta
    End If

    tamPersonal = Len(personal)
    tamRegistro = Len(Empleado_1)

    If tamPersonal <= 0 Or tamRegistro <= 0 Then
        Err.Raise vbObjectError + 1002, , _
                  "La estructura de los archivos no es valida."
    End If

    If (FileLen(rutaPersonal) Mod tamPersonal) <> 0 Then
        Err.Raise vbObjectError + 1003, , _
                  "Personal.dno tiene un tamaño incompatible."
    End If

    hPersonal = FreeFile

    Open rutaPersonal For Random Access Read As #hPersonal _
         Len = tamPersonal

    personalAbierto = True

    totalPersonal = LOF(hPersonal) \ tamPersonal

    If totalPersonal <= 0 Then
        Err.Raise vbObjectError + 1004, , _
                  "Personal.dno no contiene empleados."
    End If

    ReDim vistos(1 To totalPersonal)

    '========================================================
    '4. CONSULTAR SQL SERVER - SOLO LECTURA
    '========================================================

    sSQL = "SELECT idNomina, nombre, apellidoP, apellidoM, " & _
           "nombreDeVialidad, nombreColonia, nombreLocalidad, " & _
           "entidadFederativa, nombreMunicipio, cp, correoElectronico " & _
           "FROM datosSat WHERE empresa = '" & _
           Replace(abrEmpresa, "'", "''") & "' " & _
           "ORDER BY idNomina"

    Set oRS = New ADODB.Recordset

    oRS.CursorLocation = adUseClient

    oRS.Open sSQL, con, adOpenStatic, adLockReadOnly, adCmdText

    If oRS.EOF Then
        MsgBox "SQL Server no devolvio empleados para " & _
               abrEmpresa & "." & vbCrLf & _
               "No se modificaron los archivos.", _
               vbInformation, "Sincronizacion"

        GoTo Salida
    End If

    '========================================================
    '5. VALIDAR TODOS LOS EMPLEADOS ANTES DE ESCRIBIR
    '========================================================

    totalSQL = 0

    oRS.MoveFirst

    Do While Not oRS.EOF

        idTexto = CampoSQLCFDI(oRS.Fields("idNomina").Value)

        'Aceptar solamente numeros enteros positivos.
        If idTexto = "" Or idTexto Like "*[!0-9]*" Then
            Err.Raise vbObjectError + 1005, , _
                      "La BD contiene un idNomina invalido: " & _
                      idTexto
        End If

        id = CLng(idTexto)

        If id < 1 Or id > totalPersonal Then
            Err.Raise vbObjectError + 1006, , _
                      "El empleado SQL con ID " & id & _
                      " no existe en Personal.dno."
        End If

        If vistos(id) Then
            Err.Raise vbObjectError + 1007, , _
                      "SQL devolvio dos veces el empleado " & id
        End If

        vistos(id) = True

        'Leer la persona en la posicion indicada.
        Get #hPersonal, id, personal

        If Trim$(personal.nom) = "" Then
            Err.Raise vbObjectError + 1008, , _
                      "El registro local " & id & _
                      " no contiene nombre."
        End If

        '----------------------------------------------------
        'VALIDAR NOMBRE
        '----------------------------------------------------

        If NombreComparableCFDI(Trim$(personal.nom)) <> _
           NombreComparableCFDI(CampoSQLCFDI(oRS.Fields("nombre").Value)) Then

            Err.Raise vbObjectError + 1009, , _
                      "No coincide el nombre del empleado " & id & "." & _
                      vbCrLf & _
                      "Local: " & Trim$(personal.nom) & vbCrLf & _
                      "SQL: " & CampoSQLCFDI(oRS.Fields("nombre").Value)
        End If

        '----------------------------------------------------
        'VALIDAR APELLIDOS PATERNO Y MATERNO
        '----------------------------------------------------

        If NombreComparableCFDI(Trim$(personal.ape1)) <> _
           NombreComparableCFDI(CampoSQLCFDI(oRS.Fields("apellidoP").Value)) Or _
           NombreComparableCFDI(Trim$(personal.ape2)) <> _
           NombreComparableCFDI(CampoSQLCFDI(oRS.Fields("apellidoM").Value)) Then

            Err.Raise vbObjectError + 1010, , _
                      "No coinciden los apellidos del empleado " & id & "." & _
                      vbCrLf & _
                      "Local: " & Trim$(personal.ape1) & " " & _
                      Trim$(personal.ape2) & vbCrLf & _
                      "SQL: " & _
                      CampoSQLCFDI(oRS.Fields("apellidoP").Value) & " " & _
                      CampoSQLCFDI(oRS.Fields("apellidoM").Value)
        End If

        totalSQL = totalSQL + 1

        oRS.MoveNext
    Loop

    '========================================================
    '6. CONTAR EMPLEADOS LOCALES NO DEVUELTOS POR SQL
    '========================================================

    sinDatosSQL = 0

    For i = 1 To totalPersonal

        Get #hPersonal, i, personal

        If Trim$(personal.nom) <> "" Then
            If Not vistos(i) Then
                sinDatosSQL = sinDatosSQL + 1
            End If
        End If

    Next i

    '========================================================
    '7. VALIDAR Perscfdi.dno
    '========================================================

    existeCFDI = (Dir$(rutaCFDI) <> "")

    If existeCFDI Then

        If (FileLen(rutaCFDI) Mod tamRegistro) <> 0 Then

            Err.Raise vbObjectError + 1011, , _
                      "Perscfdi.dno tiene un tamaño incompatible " & _
                      "con la estructura Empleado_1."
        End If

    End If

    '========================================================
    '8. GENERAR NOMBRES UNICOS PARA RESPALDO Y TEMPORAL
    '========================================================

    sello = Format$(Now, "yyyymmdd_hhnnss")
    consecutivo = 0

    Do

        If consecutivo = 0 Then
            sufijo = ""
        Else
            sufijo = "_" & CStr(consecutivo)
        End If

        rutaRespaldo = carpeta & "Perscfdi_" & sello & sufijo & ".bak"
        rutaTemporal = carpeta & "Perscfdi_" & sello & sufijo & ".tmp"
        rutaAnterior = carpeta & "Perscfdi_" & sello & sufijo & ".old"

        If Dir$(rutaRespaldo) = "" And _
           Dir$(rutaTemporal) = "" And _
           Dir$(rutaAnterior) = "" Then
            Exit Do
        End If

        consecutivo = consecutivo + 1

    Loop

    '========================================================
    '9. RESPALDAR Y PREPARAR COPIA TEMPORAL
    '========================================================

    If existeCFDI Then

        'Copia intacta del archivo original.
        FileCopy rutaCFDI, rutaRespaldo

        'Copia de trabajo que sera modificada.
        FileCopy rutaCFDI, rutaTemporal

    Else

        'Solo se permite crear Perscfdi.dno.
        'Personal.dno nunca se crea ni se modifica.

        hTemporal = FreeFile

        Open rutaTemporal For Binary As #hTemporal
        Close #hTemporal

    End If

    '========================================================
    '10. ASEGURAR CAPACIDAD PARA TODOS LOS IDS
    '========================================================

    tamNecesario = totalPersonal * tamRegistro

    If FileLen(rutaTemporal) < tamNecesario Then

        hExtension = FreeFile

        Open rutaTemporal For Binary As #hExtension

        cero = 0

        'Ampliar la copia temporal sin modificar
        'el Perscfdi.dno original.
        Put #hExtension, tamNecesario, cero

        Close #hExtension

    End If

    '========================================================
    '11. GUARDAR DATOS SQL EN LA COPIA TEMPORAL
    '========================================================

    hTemporal = FreeFile

    Open rutaTemporal For Random Access Read Write As #hTemporal _
         Len = tamRegistro

    temporalAbierto = True

    registrosGuardados = 0

    oRS.MoveFirst

    Do While Not oRS.EOF

        id = CLng(oRS.Fields("idNomina").Value)

        'Leer el registro antes de modificarlo.
        'Asi se conservan los demas campos del tipo.
        Get #hTemporal, id, Empleado_1

        Empleado_1.Direccion = UCase$(CampoSQLCFDI(oRS.Fields("nombreDeVialidad").Value))

        Empleado_1.Colonia = UCase$(CampoSQLCFDI(oRS.Fields("nombreColonia").Value))

        Empleado_1.Ciudad = UCase$(CampoSQLCFDI(oRS.Fields("nombreLocalidad").Value))

        Empleado_1.Estado = UCase$(CampoSQLCFDI(oRS.Fields("entidadFederativa").Value))

        Empleado_1.Delegacion = UCase$(CampoSQLCFDI(oRS.Fields("nombreMunicipio").Value))

        Empleado_1.Cpostal = UCase$(CampoSQLCFDI(oRS.Fields("cp").Value))

        Empleado_1.correo = UCase$(CampoSQLCFDI(oRS.Fields("correoElectronico").Value))

        'Guardar exactamente en la posicion del empleado.
        Put #hTemporal, id, Empleado_1

        registrosGuardados = registrosGuardados + 1

        oRS.MoveNext
    Loop

    '========================================================
    '12. VERIFICAR TODOS LOS REGISTROS ESCRITOS
    '========================================================

    oRS.MoveFirst

    Do While Not oRS.EOF

        id = CLng(oRS.Fields("idNomina").Value)

        'Volver a leer lo escrito.
        Get #hTemporal, id, Empleado_1

        If Not VerificarRegistroCFDI(oRS) Then

            Err.Raise vbObjectError + 1012, , _
                      "La verificacion fallo para el empleado " & _
                      id & "." & vbCrLf & _
                      "Los datos almacenados no coinciden con SQL."
        End If

        oRS.MoveNext
    Loop

    '========================================================
    '13. CERRAR LOS ARCHIVOS ANTES DE REEMPLAZAR
    '========================================================

    Close #hTemporal
    temporalAbierto = False

    Close #hPersonal
    personalAbierto = False

    '========================================================
    '14. REEMPLAZAR EL ARCHIVO ORIGINAL
    '========================================================

    If existeCFDI Then

        'Retirar el archivo original sin destruirlo.
        Name rutaCFDI As rutaAnterior
        archivoMovido = True

    End If

    'Instalar la copia verificada.
    Name rutaTemporal As rutaCFDI
    archivoInstalado = True

    'Eliminar la copia retirada.
    'El respaldo .bak se conserva.
    If archivoMovido Then

        On Error Resume Next
        Kill rutaAnterior
        On Error GoTo ErrorDB

    End If

    '========================================================
    '15. MOSTRAR LOS DATOS SIN REABRIR ARCHIVOS .dno
    '========================================================

    On Error GoTo ErrorPantalla

    Col_Def
    PCpCfdi.Rows = 1

    oRS.MoveFirst

    Do While Not oRS.EOF

        id = CLng(oRS.Fields("idNomina").Value)

        nombreCompleto = _
            CampoSQLCFDI(oRS.Fields("nombre").Value) & " " & _
            CampoSQLCFDI(oRS.Fields("apellidoP").Value) & " " & _
            CampoSQLCFDI(oRS.Fields("apellidoM").Value)

        cadenaGrid = CStr(id) & vbTab & _
                     Trim$(nombreCompleto) & vbTab & _
                     UCase$(CampoSQLCFDI(oRS.Fields("nombreDeVialidad").Value)) & vbTab & _
                     UCase$(CampoSQLCFDI(oRS.Fields("nombreColonia").Value)) & vbTab & _
                     UCase$(CampoSQLCFDI(oRS.Fields("nombreLocalidad").Value)) & vbTab & _
                     UCase$(CampoSQLCFDI(oRS.Fields("entidadFederativa").Value)) & vbTab & _
                     UCase$(CampoSQLCFDI(oRS.Fields("nombreMunicipio").Value)) & vbTab & _
                     UCase$(CampoSQLCFDI(oRS.Fields("cp").Value)) & vbTab & _
                     UCase$(CampoSQLCFDI(oRS.Fields("correoElectronico").Value))

        PCpCfdi.AddItem cadenaGrid

        oRS.MoveNext
    Loop

    PeCfdi.Caption = _
        "Datos complementarios de Personal - Sincronizado"

    MsgBox "Sincronizacion completada correctamente." & _
           vbCrLf & vbCrLf & _
           "Empresa: " & abrEmpresa & vbCrLf & _
           "Registros recibidos de SQL: " & totalSQL & vbCrLf & _
           "Registros guardados y verificados: " & registrosGuardados & vbCrLf & _
           "Registros en Personal.dno: " & totalPersonal & vbCrLf & _
           "Personal local sin datos en la consulta SQL: " & sinDatosSQL & _
           vbCrLf & vbCrLf & _
           "Perscfdi.dno actualizado correctamente.", _
           vbInformation, "Sincronizacion CFDI"

    GoTo Salida

'============================================================
'ERROR AL MOSTRAR LA INFORMACION
'EL ARCHIVO YA FUE GUARDADO CORRECTAMENTE
'============================================================

ErrorPantalla:

    numeroError = Err.Number
    descripcionError = Err.Description

    MsgBox "Los datos fueron guardados y verificados en Perscfdi.dno." & _
           vbCrLf & _
           "Sin embargo, ocurrio un error al mostrarlos." & _
           vbCrLf & _
           "Error " & numeroError & ": " & descripcionError, _
           vbExclamation, "Sincronizacion CFDI"

    GoTo Salida

'============================================================
'MANEJO DE ERRORES
'============================================================

ErrorDB:

    numeroError = Err.Number
    descripcionError = Err.Description

    On Error Resume Next

    If temporalAbierto Then
        Close #hTemporal
        temporalAbierto = False
    End If

    If personalAbierto Then
        Close #hPersonal
        personalAbierto = False
    End If

    'Si se retiro el original, intentar restaurarlo.
    If archivoMovido And Not archivoInstalado Then

        If Dir$(rutaCFDI) = "" And Dir$(rutaAnterior) <> "" Then

            Err.Clear

            Name rutaAnterior As rutaCFDI

            If Err.Number <> 0 Then
                errorRestauracion = Err.Description
            End If

        End If

    End If

    'Descartar solamente la copia de trabajo.
    If Not archivoInstalado Then

        If rutaTemporal <> "" Then
            If Dir$(rutaTemporal) <> "" Then
                Kill rutaTemporal
            End If
        End If

    End If

    If errorRestauracion <> "" Then

        MsgBox "Error durante la sincronizacion." & vbCrLf & _
               "Error " & numeroError & ": " & descripcionError & _
               vbCrLf & vbCrLf & _
               "No fue posible restaurar automaticamente el original." & _
               vbCrLf & errorRestauracion & vbCrLf & _
               "Revise el respaldo y los archivos antes de continuar.", _
               vbCritical, "Sincronizacion CFDI"

    Else

        MsgBox "No fue posible completar la sincronizacion." & _
               vbCrLf & _
               "Error " & numeroError & ": " & descripcionError & _
               vbCrLf & vbCrLf & _
               "El archivo original no fue sustituido.", _
               vbCritical, "Sincronizacion CFDI"

    End If

'============================================================
'SALIDA Y LIBERACION DE RECURSOS
'============================================================

Salida:

    On Error Resume Next

    If temporalAbierto Then Close #hTemporal
    If personalAbierto Then Close #hPersonal

    If Not oRS Is Nothing Then
        If oRS.State = adStateOpen Then oRS.Close
    End If

    Set oRS = Nothing

    On Error GoTo 0

End Sub


' ============================================================
' COPIAR / PEGAR / INSERTAR / ELIMINAR
' ============================================================

Private Sub EdCfCopia_Click()
    Dim Temporal1

    Clipboard.Clear

    difer = PCpCfdi.RowSel - PCpCfdi.Row

    For i = PCpCfdi.Row To PCpCfdi.RowSel
        For f = PCpCfdi.Col To PCpCfdi.ColSel
            Temporal1 = Temporal1 + PCpCfdi.TextMatrix(i, f)

            If f < PCpCfdi.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
        Next f

        Temporal1 = Temporal1 & Chr(13) & Chr(10)
    Next i

    Clipboard.SetText Temporal1
End Sub

Private Sub EdCfEliminar_Click()
    Dim w1 As Long

    For w1 = PCpCfdi.Row To PCpCfdi.RowSel
        PCpCfdi.RemoveItem PCpCfdi.Row
    Next w1
End Sub

Private Sub EdCfInsertar_Click()
    Dim W As Long

    For W = PCpCfdi.Row To PCpCfdi.RowSel
        PCpCfdi.AddItem (""), PCpCfdi.Row
    Next W
End Sub

Private Sub EdCfPegar_Click()
    Dim DeAqui As Integer
    Dim RetornoCarro As Long
    Dim InicioCopia As Long

    Temporal1 = ""
    temporal = Clipboard.GetText(vbCFText)

    RetornoCarro = PCpCfdi.Col
    InicioCopia = PCpCfdi.Row
    DeAqui = InicioCopia
    InIPGr = InicioCopia
    InIPGc = RetornoCarro

    If temporal <> "" Then
        Clipboard.Clear
        DeAqui = 1

        For i = 1 To Len(temporal)
            Select Case Mid(temporal, i, 1)
                Case Chr(9)
                    LTXT.Text = Mid(temporal, DeAqui, (i - DeAqui))
                    Temporal1 = Temporal1 + PCpCfdi.TextMatrix(PCpCfdi.Row, PCpCfdi.Col) & Chr(9)
                    PCpCfdi.Text = Mid(temporal, DeAqui, (i - DeAqui))
                    PCpCfdi.Col = PCpCfdi.Col + 1
                    DeAqui = i + 1

                Case Chr(13)
                    Temporal1 = Temporal1 + PCpCfdi.TextMatrix(PCpCfdi.Row, PCpCfdi.Col) & Chr(13)
                    LTXT.Text = Mid(temporal, DeAqui, (i - DeAqui))
                    PCpCfdi.Text = Mid(temporal, DeAqui, (i - DeAqui))

                    If (PCpCfdi.Rows - 1) <= PCpCfdi.Row Then PCpCfdi.Rows = PCpCfdi.Rows + 1

                    PCpCfdi.Row = PCpCfdi.Row + 1
                    PCpCfdi.TopRow = 1
                    DeAqui = i + 1

                Case Chr(10)
                    Temporal1 = Temporal1 & Chr(10)
                    PCpCfdi.Col = RetornoCarro
                    DeAqui = i + 1

                Case Else
                    Rem nada
            End Select
        Next i

        PCpCfdi.Col = RetornoCarro
    End If

    Clipboard.Clear
    Clipboard.SetText Clipboard.GetText + Temporal1
End Sub

Private Sub EdCfSelT_Click()
    Dim limite As Long

    Clipboard.Clear

    PCpCfdi.Row = 1
    PCpCfdi.Col = 0

    For limite = 1 To PCpCfdi.Rows - 1
        renglon = limite

        If IsNumeric(PCpCfdi.TextMatrix(renglon, 0)) Then
            PCpCfdi.RowSel = PCpCfdi.Row
        End If
    Next limite

    PCpCfdi.ColSel = PCpCfdi.Cols - 1
End Sub

' ============================================================
' EDICION DE CELDAS
' ============================================================

Private Sub LTXT_Change()
    PCpCfdi.Text = LTXT.Text
End Sub

Private Sub pcpcfdi_EnterCell()
    PCpCfdi.CellBackColor = vbYellow
    valcelant = PCpCfdi.Text
    LTXT.Text = valcelant
End Sub

Private Sub pcpcfdi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            For q = PCpCfdi.Row To PCpCfdi.RowSel
                For W = PCpCfdi.Col To PCpCfdi.ColSel
                    PCpCfdi.TextMatrix(q, W) = ""
                Next W
            Next q

            LTXT.Text = PCpCfdi.Text

        Case vbKeyF2
            If PCpCfdi.Text <> "" Then valcelant = PCpCfdi.Text

            LTXT.Text = LTrim(RTrim(PCpCfdi.Text))
            LTXT.SetFocus
    End Select
End Sub

Private Sub pcpcfdi_KeyPress(KeyAscii As Integer)
    valcelant = PCpCfdi.Text
    LTXT.Text = Chr(KeyAscii)
    LTXT.SetFocus
End Sub

Private Sub pcpcfdi_LeaveCell()
    PCpCfdi.CellBackColor = vbWhite
End Sub



Private Function CampoSQLCFDI(ByVal valor As Variant) As String

    If IsNull(valor) Or IsEmpty(valor) Then
        CampoSQLCFDI = ""
    Else
        CampoSQLCFDI = Trim$(CStr(valor))
    End If

End Function

Private Function NombreComparableCFDI(ByVal texto As String) As String

    texto = UCase$(Trim$(texto))

    texto = Replace(texto, "Á", "A")
    texto = Replace(texto, "É", "E")
    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ó", "O")
    texto = Replace(texto, "Ú", "U")

    Do While InStr(texto, "  ") > 0
        texto = Replace(texto, "  ", " ")
    Loop

    NombreComparableCFDI = texto

End Function

Private Function VerificarRegistroCFDI( _
    ByVal oRS As ADODB.Recordset) As Boolean

    VerificarRegistroCFDI = False

    If Trim$(Empleado_1.Direccion) <> _
       UCase$(CampoSQLCFDI(oRS.Fields("nombreDeVialidad").Value)) Then Exit Function

    If Trim$(Empleado_1.Colonia) <> _
       UCase$(CampoSQLCFDI(oRS.Fields("nombreColonia").Value)) Then Exit Function

    If Trim$(Empleado_1.Ciudad) <> _
       UCase$(CampoSQLCFDI(oRS.Fields("nombreLocalidad").Value)) Then Exit Function

    If Trim$(Empleado_1.Estado) <> _
       UCase$(CampoSQLCFDI(oRS.Fields("entidadFederativa").Value)) Then Exit Function

    If Trim$(Empleado_1.Delegacion) <> _
       UCase$(CampoSQLCFDI(oRS.Fields("nombreMunicipio").Value)) Then Exit Function

    If Trim$(Empleado_1.Cpostal) <> _
       UCase$(CampoSQLCFDI(oRS.Fields("cp").Value)) Then Exit Function

    If Trim$(Empleado_1.correo) <> _
       UCase$(CampoSQLCFDI(oRS.Fields("correoElectronico").Value)) Then Exit Function

    VerificarRegistroCFDI = True

End Function



