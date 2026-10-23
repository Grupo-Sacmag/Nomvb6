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
Option Explicit

' Nueva Versión

Private temporal As String
Private Temporal1 As String
Private valcelant As String

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

    PCpCfdi.col = 0
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(0) = 800
    PCpCfdi.Text = "Num"

    PCpCfdi.col = 1
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(1) = 3600
    PCpCfdi.Text = "Nombre"

    PCpCfdi.col = 2
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(2) = 2400
    PCpCfdi.Text = "Direccion"

    PCpCfdi.col = 3
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(3) = 2400
    PCpCfdi.Text = "Colonia"

    PCpCfdi.col = 4
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(4) = 2400
    PCpCfdi.Text = "Ciudad"

    PCpCfdi.col = 5
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(5) = 2400
    PCpCfdi.Text = "Estado"

    PCpCfdi.col = 6
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(6) = 2400
    PCpCfdi.Text = "Delegacion"

    PCpCfdi.col = 7
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(7) = 800
    PCpCfdi.Text = "Codigo"

    PCpCfdi.col = 8
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

Private Function ObtenerEmpresaSAT() As String
    Dim abrEmpresa As String

    abrEmpresa = Left(Trim(emp), 4)

    If (UCase(abrEmpresa) = "SACM") Then abrEmpresa = "SACMAG"
    If (UCase(abrEmpresa) = "SUPT") Then abrEmpresa = "SUPTER"
    If (UCase(abrEmpresa) = "COOR") Then abrEmpresa = "CORDINA"
    If (UCase(abrEmpresa) = "EPES") Then abrEmpresa = "EPESA"
    If (UCase(abrEmpresa) = "SUPE") Then abrEmpresa = "SUPERVISA"

    ObtenerEmpresaSAT = abrEmpresa
End Function

' ============================================================
' ARCHIVO LOCAL CFDI
' ============================================================

Private Sub LimpiarEmpleadoCFDI()
    Empleado_1.Direccion = ""
    Empleado_1.Colonia = ""
    Empleado_1.Ciudad = ""
    Empleado_1.Estado = ""
    Empleado_1.Delegacion = ""
    Empleado_1.Cpostal = ""
    Empleado_1.correo = ""
End Sub

'Quita espacios y los caracteres nulos (Chr(0)) que quedan en registros
'que existen en el archivo pero nunca se han escrito.
Private Function CampoArchivoCFDI(ByVal valor As String) As String
    CampoArchivoCFDI = Trim$(Replace(valor, vbNullChar, ""))
End Function

'True si el empleado cargado en el buffer "personal" tiene fecha de baja.
'Siguen existiendo en Personal.dno, pero no se muestran ni se sincronizan.
Private Function EmpleadoDeBaja() As Boolean
    EmpleadoDeBaja = IsDate(Trim$(Replace(personal.fab, vbNullChar, "")))
End Function

'Llena el grid leyendo Personal.dno y Perscfdi.dno.
'Lo usan el boton Abrir y, al terminar, el boton DB, para que el grid
'siempre muestre lo que realmente quedo guardado en el archivo.
Private Sub CargarGridDesdeArchivos()
    Dim hPers As Integer
    Dim hCfdi As Integer
    Dim totalPers As Long
    Dim totalCfdi As Long
    Dim i As Long
    Dim nombre As String
    Dim cadena As String

    On Error GoTo ErrCarga

    PCpCfdi.Rows = 1

    If Dir$("Personal.dno") = "" Then
        Err.Raise vbObjectError + 1101, , "No existe Personal.dno en la carpeta actual."
    End If

    hPers = FreeFile
    Open "Personal.dno" For Random Access Read As #hPers Len = Len(personal)
    totalPers = LOF(hPers) \ Len(personal)

    'Si Perscfdi.dno aun no existe, se muestran los empleados sin datos complementarios.
    If Dir$("Perscfdi.dno") <> "" Then
        hCfdi = FreeFile
        Open "Perscfdi.dno" For Random Access Read As #hCfdi Len = Len(Empleado_1)
        totalCfdi = LOF(hCfdi) \ Len(Empleado_1)
    End If

    For i = 1 To totalPers
        Get #hPers, i, personal

        'Los empleados de baja no se muestran en el grid.
        If Trim$(personal.nom) <> "" And Not EmpleadoDeBaja() Then

            'Los empleados mas nuevos pueden quedar despues del final de Perscfdi.dno.
            'Get no marca error en ese caso y dejaria los datos del empleado anterior.
            If i <= totalCfdi Then
                Get #hCfdi, i, Empleado_1
            Else
                LimpiarEmpleadoCFDI
            End If

            nombre = Trim$(personal.nom) & " " & Trim$(personal.ape1) & " " & Trim$(personal.ape2)

            cadena = Format$(i, "####0") & vbTab & nombre & vbTab & _
                     CampoArchivoCFDI(Empleado_1.Direccion) & vbTab & _
                     CampoArchivoCFDI(Empleado_1.Colonia) & vbTab & _
                     CampoArchivoCFDI(Empleado_1.Ciudad) & vbTab & _
                     CampoArchivoCFDI(Empleado_1.Estado) & vbTab & _
                     CampoArchivoCFDI(Empleado_1.Delegacion) & vbTab & _
                     CampoArchivoCFDI(Empleado_1.Cpostal) & vbTab & _
                     CampoArchivoCFDI(Empleado_1.correo)

            PCpCfdi.AddItem cadena
        End If
    Next i

Salir:
    On Error Resume Next
    If hCfdi <> 0 Then Close #hCfdi
    If hPers <> 0 Then Close #hPers
    Exit Sub

ErrCarga:
    MsgBox "No fue posible cargar los datos." & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Datos complementarios"
    Resume Salir
End Sub

Private Sub ArCfAbr_Click()
    CargarGridDesdeArchivos
End Sub

Private Sub ArCfGuarda_Click()
    Dim h As Integer
    Dim EmpNum As Long
    Dim renglon As Long

    h = FreeFile
    Open "Perscfdi.dno" For Random As #h Len = Len(Empleado_1)

    For renglon = 1 To PCpCfdi.Rows - 1
        If IsNumeric(PCpCfdi.TextMatrix(renglon, 0)) Then
            EmpNum = CLng(PCpCfdi.TextMatrix(renglon, 0))

            If EmpNum >= 1 Then
                'Leer antes de modificar: conserva cualquier otro campo del tipo
                'y evita arrastrar valores del renglon anterior.
                Get #h, EmpNum, Empleado_1

                Empleado_1.Direccion = PCpCfdi.TextMatrix(renglon, 2)
                Empleado_1.Colonia = PCpCfdi.TextMatrix(renglon, 3)
                Empleado_1.Ciudad = PCpCfdi.TextMatrix(renglon, 4)
                Empleado_1.Estado = PCpCfdi.TextMatrix(renglon, 5)
                Empleado_1.Delegacion = PCpCfdi.TextMatrix(renglon, 6)
                Empleado_1.Cpostal = PCpCfdi.TextMatrix(renglon, 7)
                Empleado_1.correo = PCpCfdi.TextMatrix(renglon, 8)

                Put #h, EmpNum, Empleado_1
            End If
        End If
    Next renglon

    Close #h
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
    Dim bajasLocales As Long
    Dim procesables As Long
    Dim omitidos As Long
    Dim conservados As Long
    Dim duplicados As Long
    Dim sinCorreo As Long

    Dim id As Long
    Dim idTexto As String
    Dim i As Long
    Dim consecutivo As Long

    Dim vistos() As Boolean
    Dim procesar() As Boolean
    Dim ultimaFila() As Long
    Dim tamRegistro As Long
    Dim tamPersonal As Long
    Dim tamNecesario As Long

    Dim existeCFDI As Boolean
    Dim archivoMovido As Boolean
    Dim archivoInstalado As Boolean
    Dim personalAbierto As Boolean
    Dim temporalAbierto As Boolean
    Dim conservar As Boolean

    Dim hPersonal As Integer
    Dim hTemporal As Integer
    Dim hExtension As Integer

    Dim cero As Byte
    Dim numeroError As Long
    Dim descripcionError As String
    Dim errorRestauracion As String

    Dim motivo As String
    Dim listaOmitidos As String
    Dim listaConservados As String
    Dim listaDuplicados As String
    Dim listaSinCorreo As String
    Dim resumen As String
    Dim estilo As Long

    On Error GoTo ErrorDB

    '========================================================
    '1. VALIDAR EMPRESA
    '========================================================

    If EsEmpresaLocal(emp) Then
        MsgBox "Esta empresa trabaja en modo local." & vbCrLf & "No se modificaron los archivos CFDI.", vbInformation, "Sincronizacion"
        Exit Sub
    End If

    abrEmpresa = Trim$(ObtenerEmpresaSAT())

    If abrEmpresa = "" Then
        MsgBox "No fue posible identificar la empresa.", vbExclamation, "Sincronizacion"
        Exit Sub
    End If

    '========================================================
    '2. VERIFICAR CONEXION SQL
    '========================================================

    If con Is Nothing Then
        MsgBox "La conexion SQL no esta inicializada.", vbCritical, "Sincronizacion"
        Exit Sub
    End If

    If con.State <> adStateOpen Then
        MsgBox "No hay conexion con SQL Server." & vbCrLf & "No se modificaron los archivos locales.", vbExclamation, "Sincronizacion"
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
        Err.Raise vbObjectError + 1001, , "No existe Personal.dno en: " & carpeta
    End If

    tamPersonal = Len(personal)
    tamRegistro = Len(Empleado_1)

    If tamPersonal <= 0 Or tamRegistro <= 0 Then
        Err.Raise vbObjectError + 1002, , "La estructura de los archivos no es valida."
    End If

    If (FileLen(rutaPersonal) Mod tamPersonal) <> 0 Then
        Err.Raise vbObjectError + 1003, , "Personal.dno tiene un tamaño incompatible."
    End If

    hPersonal = FreeFile

    Open rutaPersonal For Random Access Read As #hPersonal Len = tamPersonal

    personalAbierto = True

    totalPersonal = LOF(hPersonal) \ tamPersonal

    If totalPersonal <= 0 Then
        Err.Raise vbObjectError + 1004, , "Personal.dno no contiene empleados."
    End If

    ReDim vistos(1 To totalPersonal)
    ReDim procesar(1 To totalPersonal)
    ReDim ultimaFila(1 To totalPersonal)

    '========================================================
    '4. CONSULTAR SQL SERVER - SOLO LECTURA
    '========================================================

    sSQL = "SELECT idNomina, nombre, apellidoP, apellidoM, " & _
           "nombreDeVialidad, nombreColonia, nombreLocalidad, " & _
           "entidadFederativa, nombreMunicipio, cp, correoElectronico " & _
           "FROM datosSat WHERE empresa = '" & Replace(abrEmpresa, "'", "''") & "' " & _
           "ORDER BY idNomina"

    Set oRS = New ADODB.Recordset

    oRS.CursorLocation = adUseClient

    oRS.Open sSQL, con, adOpenStatic, adLockReadOnly, adCmdText

    If oRS.EOF Then
        MsgBox "SQL Server no devolvio empleados para " & abrEmpresa & "." & vbCrLf & "No se modificaron los archivos.", vbInformation, "Sincronizacion"
        GoTo Salida
    End If

    '========================================================
    '4B. DUPLICADOS
    '    Si un idNomina viene mas de una vez, se usa el ULTIMO
    '    registro que devolvio SQL; los anteriores se ignoran.
    '    Aqui solo se anota la posicion del ultimo de cada ID.
    '========================================================

    duplicados = 0

    oRS.MoveFirst

    Do While Not oRS.EOF

        id = ObtenerIdCFDI(CampoSQLCFDI(oRS.Fields("idNomina").Value), totalPersonal)

        If id > 0 Then
            If ultimaFila(id) > 0 Then
                duplicados = duplicados + 1
                If Len(listaDuplicados) < 100 Then listaDuplicados = listaDuplicados & id & " "
            End If

            ultimaFila(id) = oRS.AbsolutePosition
        End If

        oRS.MoveNext
    Loop

    '========================================================
    '5. VALIDAR CADA EMPLEADO ANTES DE ESCRIBIR
    '   - Un empleado con problemas se OMITE y se reporta al final;
    '     no cancela la sincronizacion de los demas.
    '   - Si la BD trae datos incompletos y el archivo YA tiene
    '     datos en esa posicion, no se modifica nada de ese empleado.
    '========================================================

    totalSQL = 0
    procesables = 0
    omitidos = 0
    conservados = 0
    sinCorreo = 0

    oRS.MoveFirst

    Do While Not oRS.EOF

        totalSQL = totalSQL + 1
        motivo = ""
        conservar = False

        idTexto = CampoSQLCFDI(oRS.Fields("idNomina").Value)
        id = ObtenerIdCFDI(idTexto, totalPersonal)

        'Registro repetido que no es el ultimo: se ignora.
        If id > 0 Then
            If ultimaFila(id) <> oRS.AbsolutePosition Then GoTo SiguienteFila
        End If

        If id = 0 Then

            motivo = "idNomina invalido o inexistente en Personal.dno"

        Else

            vistos(id) = True

            'Leer la persona en la posicion indicada.
            Get #hPersonal, id, personal

            'Empleado de baja: sigue en Personal.dno pero se ignora en la sincronizacion.
            If EmpleadoDeBaja() Then GoTo SiguienteFila

            If Trim$(personal.nom) = "" Then

                motivo = "el registro local no tiene nombre"

            ElseIf NombreComparableCFDI(Trim$(personal.nom)) <> NombreComparableCFDI(CampoSQLCFDI(oRS.Fields("nombre").Value)) Then

                motivo = "nombre distinto (Local: " & Trim$(personal.nom) & " / SQL: " & CampoSQLCFDI(oRS.Fields("nombre").Value) & ")"

            ElseIf NombreComparableCFDI(Trim$(personal.ape1)) <> NombreComparableCFDI(CampoSQLCFDI(oRS.Fields("apellidoP").Value)) Or _
                   NombreComparableCFDI(Trim$(personal.ape2)) <> NombreComparableCFDI(CampoSQLCFDI(oRS.Fields("apellidoM").Value)) Then

                motivo = "apellidos distintos (Local: " & Trim$(personal.ape1) & " " & Trim$(personal.ape2) & _
                         " / SQL: " & CampoSQLCFDI(oRS.Fields("apellidoP").Value) & " " & CampoSQLCFDI(oRS.Fields("apellidoM").Value) & ")"

            End If

        End If

        '----------------------------------------------------
        'LONGITUD Y CARACTERES GUARDABLES EN CADA CAMPO
        '----------------------------------------------------

        If motivo = "" Then motivo = RevisarCampoCFDI("Direccion", UCase$(CampoSQLCFDI(oRS.Fields("nombreDeVialidad").Value)), Len(Empleado_1.Direccion))
        If motivo = "" Then motivo = RevisarCampoCFDI("Colonia", UCase$(CampoSQLCFDI(oRS.Fields("nombreColonia").Value)), Len(Empleado_1.Colonia))
        If motivo = "" Then motivo = RevisarCampoCFDI("Ciudad", UCase$(CampoSQLCFDI(oRS.Fields("nombreLocalidad").Value)), Len(Empleado_1.Ciudad))
        If motivo = "" Then motivo = RevisarCampoCFDI("Estado", UCase$(CampoSQLCFDI(oRS.Fields("entidadFederativa").Value)), Len(Empleado_1.Estado))
        If motivo = "" Then motivo = RevisarCampoCFDI("Delegacion", UCase$(CampoSQLCFDI(oRS.Fields("nombreMunicipio").Value)), Len(Empleado_1.Delegacion))
        If motivo = "" Then motivo = RevisarCampoCFDI("Codigo postal", UCase$(CampoSQLCFDI(oRS.Fields("cp").Value)), Len(Empleado_1.Cpostal))
        If motivo = "" Then motivo = RevisarCampoCFDI("Correo", UCase$(CampoSQLCFDI(oRS.Fields("correoElectronico").Value)), Len(Empleado_1.correo))

        '----------------------------------------------------
        'BD INCOMPLETA + YA HAY DATOS EN EL ARCHIVO = NO TOCAR
        '----------------------------------------------------

        If motivo = "" Then
            If SQLIncompleto(oRS) Then
                conservar = LocalCFDITieneDatos(rutaCFDI, id)
            End If
        End If

        If conservar Then

            conservados = conservados + 1
            If Len(listaConservados) < 100 Then listaConservados = listaConservados & id & " "

        ElseIf motivo = "" Then

            procesar(id) = True
            procesables = procesables + 1

            If CampoSQLCFDI(oRS.Fields("correoElectronico").Value) = "" Then
                sinCorreo = sinCorreo + 1
                If Len(listaSinCorreo) < 100 Then listaSinCorreo = listaSinCorreo & id & " "
            End If

        Else

            omitidos = omitidos + 1
            If omitidos <= 5 Then
                listaOmitidos = listaOmitidos & "- ID " & idTexto & ": " & motivo & vbCrLf
            End If

        End If

SiguienteFila:
        oRS.MoveNext
    Loop

    If procesables = 0 Then
        resumen = "No hay empleados para actualizar." & vbCrLf & "No se modificaron los archivos."

        If conservados > 0 Then
            resumen = resumen & vbCrLf & vbCrLf & "Sin cambios (BD incompleta y ya hay datos locales): " & conservados & vbCrLf & "IDs: " & Trim$(listaConservados)
        End If

        If omitidos > 0 Then
            resumen = resumen & vbCrLf & vbCrLf & "OMITIDOS (" & omitidos & "):" & vbCrLf & listaOmitidos
        End If

        MsgBox resumen, vbInformation, "Sincronizacion CFDI"
        GoTo Salida
    End If

    '========================================================
    '6. CONTAR EMPLEADOS LOCALES NO DEVUELTOS POR SQL
    '========================================================

    sinDatosSQL = 0
    bajasLocales = 0

    For i = 1 To totalPersonal

        Get #hPersonal, i, personal

        If Trim$(personal.nom) <> "" Then
            If EmpleadoDeBaja() Then
                bajasLocales = bajasLocales + 1
            ElseIf Not vistos(i) Then
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
            Err.Raise vbObjectError + 1011, , "Perscfdi.dno tiene un tamaño incompatible con la estructura Empleado_1."
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
    '    Solo los empleados marcados en el paso 5 (y solo su
    '    ultimo registro si venian repetidos).
    '========================================================

    hTemporal = FreeFile
    Open rutaTemporal For Random Access Read Write As #hTemporal Len = tamRegistro
    temporalAbierto = True
    registrosGuardados = 0

    oRS.MoveFirst

    Do While Not oRS.EOF

        id = ObtenerIdCFDI(CampoSQLCFDI(oRS.Fields("idNomina").Value), totalPersonal)

        If id > 0 Then
            If procesar(id) And ultimaFila(id) = oRS.AbsolutePosition Then

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

            End If
        End If

        oRS.MoveNext
    Loop

    '========================================================
    '12. VERIFICAR TODOS LOS REGISTROS ESCRITOS
    '========================================================

    oRS.MoveFirst

    Do While Not oRS.EOF

        id = ObtenerIdCFDI(CampoSQLCFDI(oRS.Fields("idNomina").Value), totalPersonal)

        If id > 0 Then
            If procesar(id) And ultimaFila(id) = oRS.AbsolutePosition Then

                'Volver a leer lo escrito.
                Get #hTemporal, id, Empleado_1

                If Not VerificarRegistroCFDI(oRS) Then
                    Err.Raise vbObjectError + 1012, , "La verificacion fallo para el empleado " & id & "." & vbCrLf & "Los datos almacenados no coinciden con SQL."
                End If

            End If
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
    '15. MOSTRAR LO QUE QUEDO GUARDADO EN EL ARCHIVO
    '========================================================

    On Error GoTo ErrorPantalla

    Col_Def
    CargarGridDesdeArchivos

    PeCfdi.Caption = "Datos complementarios de Personal - Sincronizado"

    resumen = "Sincronizacion completada." & vbCrLf & vbCrLf & _
              "Empresa: " & abrEmpresa & vbCrLf & _
              "Registros recibidos de SQL: " & totalSQL & vbCrLf & _
              "Guardados y verificados: " & registrosGuardados & vbCrLf & _
              "Registros en Personal.dno: " & totalPersonal & vbCrLf & _
              "Personal activo sin datos en SQL: " & sinDatosSQL & vbCrLf & _
              "Empleados de baja (ignorados): " & bajasLocales

    If duplicados > 0 Then
        resumen = resumen & vbCrLf & vbCrLf & "Repetidos en SQL: " & duplicados & " (se uso el ultimo)" & vbCrLf & "IDs: " & Trim$(listaDuplicados)
    End If

    If conservados > 0 Then
        resumen = resumen & vbCrLf & vbCrLf & "Sin cambios (BD incompleta y ya hay datos locales): " & conservados & vbCrLf & "IDs: " & Trim$(listaConservados)
    End If

    If sinCorreo > 0 Then
        resumen = resumen & vbCrLf & vbCrLf & "Guardados sin correo (aun no llega a SQL): " & sinCorreo & vbCrLf & "IDs: " & Trim$(listaSinCorreo)
    End If

    estilo = vbInformation

    If omitidos > 0 Then
        estilo = vbExclamation
        resumen = resumen & vbCrLf & vbCrLf & "OMITIDOS (" & omitidos & "), no se modificaron:" & vbCrLf & listaOmitidos
        If omitidos > 5 Then resumen = resumen & "... y " & (omitidos - 5) & " mas."
    End If

    MsgBox resumen, estilo, "Sincronizacion CFDI"

    GoTo Salida

'============================================================
'ERROR AL MOSTRAR LA INFORMACION
'EL ARCHIVO YA FUE GUARDADO CORRECTAMENTE
'============================================================

ErrorPantalla:
    numeroError = Err.Number
    descripcionError = Err.Description

    MsgBox "Los datos fueron guardados y verificados en Perscfdi.dno." & vbCrLf & "Sin embargo, ocurrio un error al mostrarlos." & vbCrLf & "Error " & numeroError & ": " & descripcionError, vbExclamation, "Sincronizacion CFDI"

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
        MsgBox "Error durante la sincronizacion." & vbCrLf & "Error " & numeroError & ": " & descripcionError & vbCrLf & vbCrLf & _
               "No fue posible restaurar automaticamente el original." & vbCrLf & errorRestauracion & vbCrLf & _
               "Revise el respaldo y los archivos antes de continuar.", vbCritical, "Sincronizacion CFDI"
    Else
        MsgBox "No fue posible completar la sincronizacion." & vbCrLf & "Error " & numeroError & ": " & descripcionError & vbCrLf & vbCrLf & _
               "El archivo original no fue sustituido.", vbCritical, "Sincronizacion CFDI"
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
    Dim Temporal1 As String
    Dim i As Long
    Dim f As Long

    Clipboard.Clear

    For i = PCpCfdi.Row To PCpCfdi.RowSel
        For f = PCpCfdi.col To PCpCfdi.ColSel
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
    Dim DeAqui As Long
    Dim RetornoCarro As Long
    Dim InicioCopia As Long
    Dim i As Long
    Dim InIPGr As Long
    Dim InIPGc As Long

    Temporal1 = ""
    temporal = Clipboard.GetText(vbCFText)

    RetornoCarro = PCpCfdi.col
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
                    Temporal1 = Temporal1 + PCpCfdi.TextMatrix(PCpCfdi.Row, PCpCfdi.col) & Chr(9)
                    PCpCfdi.Text = Mid(temporal, DeAqui, (i - DeAqui))
                    PCpCfdi.col = PCpCfdi.col + 1
                    DeAqui = i + 1

                Case Chr(13)
                    Temporal1 = Temporal1 + PCpCfdi.TextMatrix(PCpCfdi.Row, PCpCfdi.col) & Chr(13)
                    LTXT.Text = Mid(temporal, DeAqui, (i - DeAqui))
                    PCpCfdi.Text = Mid(temporal, DeAqui, (i - DeAqui))

                    If (PCpCfdi.Rows - 1) <= PCpCfdi.Row Then PCpCfdi.Rows = PCpCfdi.Rows + 1

                    PCpCfdi.Row = PCpCfdi.Row + 1
                    PCpCfdi.TopRow = 1
                    DeAqui = i + 1

                Case Chr(10)
                    Temporal1 = Temporal1 & Chr(10)
                    PCpCfdi.col = RetornoCarro
                    DeAqui = i + 1

                Case Else
                    Rem nada
            End Select
        Next i

        PCpCfdi.col = RetornoCarro
    End If

    Clipboard.Clear
    Clipboard.SetText Clipboard.GetText + Temporal1
End Sub

Private Sub EdCfSelT_Click()
    Dim limite As Long
    Dim renglon As Long

    Clipboard.Clear

    PCpCfdi.Row = 1
    PCpCfdi.col = 0

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
    Dim q As Long
    Dim W As Long

    Select Case KeyCode
        Case vbKeyDelete
            For q = PCpCfdi.Row To PCpCfdi.RowSel
                For W = PCpCfdi.col To PCpCfdi.ColSel
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

' ============================================================
' FUNCIONES AUXILIARES DE LA SINCRONIZACION
' ============================================================

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
    texto = Replace(texto, "Ü", "U")

    Do While InStr(texto, "  ") > 0
        texto = Replace(texto, "  ", " ")
    Loop

    'Un apellido "N/A" en SQL equivale a un apellido vacio en el archivo local.
    If texto = "N/A" Then texto = ""

    NombreComparableCFDI = texto

End Function

'Devuelve el idNomina como numero (1 a maximo), o 0 si no es valido.
Private Function ObtenerIdCFDI(ByVal idTexto As String, ByVal maximo As Long) As Long

    ObtenerIdCFDI = 0

    If idTexto = "" Then Exit Function
    If idTexto Like "*[!0-9]*" Then Exit Function
    If Len(idTexto) > 9 Then Exit Function

    If CLng(idTexto) < 1 Or CLng(idTexto) > maximo Then Exit Function

    ObtenerIdCFDI = CLng(idTexto)

End Function

'Devuelve "" si el valor cabe y se puede guardar sin perder caracteres;
'si no, devuelve el motivo.
Private Function RevisarCampoCFDI(ByVal etiqueta As String, ByVal valor As String, ByVal largoCampo As Long) As String

    If Len(valor) > largoCampo Then
        RevisarCampoCFDI = etiqueta & " excede " & largoCampo & " caracteres"
    ElseIf StrConv(StrConv(valor, vbFromUnicode), vbUnicode) <> valor Then
        RevisarCampoCFDI = etiqueta & " tiene caracteres que no se pueden guardar"
    Else
        RevisarCampoCFDI = ""
    End If

End Function

'True si SQL trae vacio CUALQUIERA de los 7 datos del registro actual.
'Si algun dato puede quedar vacio legitimamente (por ejemplo la colonia),
'basta con quitar su linea de aqui.
Private Function SQLIncompleto(ByVal oRS As ADODB.Recordset) As Boolean

    SQLIncompleto = True

    If CampoSQLCFDI(oRS.Fields("nombreDeVialidad").Value) = "" Then Exit Function
    If CampoSQLCFDI(oRS.Fields("nombreColonia").Value) = "" Then Exit Function
    If CampoSQLCFDI(oRS.Fields("nombreLocalidad").Value) = "" Then Exit Function
    If CampoSQLCFDI(oRS.Fields("entidadFederativa").Value) = "" Then Exit Function
    If CampoSQLCFDI(oRS.Fields("nombreMunicipio").Value) = "" Then Exit Function
    If CampoSQLCFDI(oRS.Fields("cp").Value) = "" Then Exit Function
    If CampoSQLCFDI(oRS.Fields("correoElectronico").Value) = "" Then Exit Function

    SQLIncompleto = False

End Function

'True si Perscfdi.dno ya tiene algun dato guardado en la posicion id.
'Solo lee; abre y cierra el archivo por su cuenta.
'Si no puede leerlo responde True, para no arriesgarse a sobrescribir.
Private Function LocalCFDITieneDatos(ByVal ruta As String, ByVal id As Long) As Boolean
    Dim h As Integer
    Dim total As Long

    LocalCFDITieneDatos = False

    If Dir$(ruta) = "" Then Exit Function

    On Error GoTo Fallo

    h = FreeFile
    Open ruta For Random Access Read As #h Len = Len(Empleado_1)
    total = LOF(h) \ Len(Empleado_1)

    'Mas alla del final del archivo no hay nada guardado todavia.
    If id <= total Then
        Get #h, id, Empleado_1

        LocalCFDITieneDatos = (CampoArchivoCFDI(Empleado_1.Direccion) <> "") Or _
                              (CampoArchivoCFDI(Empleado_1.Colonia) <> "") Or _
                              (CampoArchivoCFDI(Empleado_1.Ciudad) <> "") Or _
                              (CampoArchivoCFDI(Empleado_1.Estado) <> "") Or _
                              (CampoArchivoCFDI(Empleado_1.Delegacion) <> "") Or _
                              (CampoArchivoCFDI(Empleado_1.Cpostal) <> "") Or _
                              (CampoArchivoCFDI(Empleado_1.correo) <> "")
    End If

Salir:
    On Error Resume Next
    If h <> 0 Then Close #h
    Exit Function

Fallo:
    LocalCFDITieneDatos = True
    Resume Salir
End Function

Private Function VerificarRegistroCFDI(ByVal oRS As ADODB.Recordset) As Boolean

    VerificarRegistroCFDI = False

    If Trim$(Empleado_1.Direccion) <> UCase$(CampoSQLCFDI(oRS.Fields("nombreDeVialidad").Value)) Then Exit Function
    If Trim$(Empleado_1.Colonia) <> UCase$(CampoSQLCFDI(oRS.Fields("nombreColonia").Value)) Then Exit Function
    If Trim$(Empleado_1.Ciudad) <> UCase$(CampoSQLCFDI(oRS.Fields("nombreLocalidad").Value)) Then Exit Function
    If Trim$(Empleado_1.Estado) <> UCase$(CampoSQLCFDI(oRS.Fields("entidadFederativa").Value)) Then Exit Function
    If Trim$(Empleado_1.Delegacion) <> UCase$(CampoSQLCFDI(oRS.Fields("nombreMunicipio").Value)) Then Exit Function
    If Trim$(Empleado_1.Cpostal) <> UCase$(CampoSQLCFDI(oRS.Fields("cp").Value)) Then Exit Function
    If Trim$(Empleado_1.correo) <> UCase$(CampoSQLCFDI(oRS.Fields("correoElectronico").Value)) Then Exit Function

    VerificarRegistroCFDI = True

End Function

