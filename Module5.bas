Attribute VB_Name = "Module5"
Option Explicit

'============================================================
' IMPORTACION SQL -> ARCHIVOS LOCALES. JAMAS ESCRIBE EN SQL.
' Destino por posicion: idNomina -> registro idNomina.
' Los valores locales existentes SIEMPRE prevalecen.
' En conflictos de identidad (nombre/RFC/CURP), omitir ID.
'============================================================

Private Const MAX_ID_IMPORTACION As Long = 32767

Private Function TextoSQL(ByVal rs As ADODB.Recordset, ByVal nombre As String) As String
    Dim v As Variant
    v = rs.Fields(nombre).Value
    If IsNull(v) Or IsEmpty(v) Then
        TextoSQL = ""
    Else
        TextoSQL = Trim$(CStr(v))
    End If
End Function

Private Function TextoLocal(ByVal valor As String) As String
    TextoLocal = Trim$(Replace(valor, vbNullChar, ""))
End Function

Private Function Comparable(ByVal valor As String) As String
    Comparable = UCase$(TextoLocal(valor))
    Comparable = Replace(Comparable, "Á", "A")
    Comparable = Replace(Comparable, "É", "E")
    Comparable = Replace(Comparable, "Í", "I")
    Comparable = Replace(Comparable, "Ó", "O")
    Comparable = Replace(Comparable, "Ú", "U")
    Comparable = Replace(Comparable, "Ü", "U")
End Function

Private Function IdSeguro(ByVal valor As Variant) As Long
    Dim s As String
    If IsNull(valor) Or IsEmpty(valor) Then Exit Function
    s = Trim$(CStr(valor))
    If Len(s) = 0 Or Len(s) > 5 Then Exit Function
    If s Like "*[!0-9]*" Then Exit Function
    If CLng(s) < 1 Or CLng(s) > MAX_ID_IMPORTACION Then Exit Function
    IdSeguro = CLng(s)
End Function

Private Function CoincideIdentidad(ByVal rs As ADODB.Recordset) As Boolean
    Dim local As String, remoto As String
    CoincideIdentidad = False

    local = Comparable(personal.nom): remoto = Comparable(TextoSQL(rs, "nombre"))
    If Len(local) > 0 And Len(remoto) > 0 Then If local <> remoto Then Exit Function

    local = Comparable(personal.ape1): remoto = Comparable(TextoSQL(rs, "apellidoP"))
    If Len(local) > 0 And Len(remoto) > 0 Then If local <> remoto Then Exit Function

    local = Comparable(personal.ape2): remoto = Comparable(TextoSQL(rs, "apellidoM"))
    If Len(local) > 0 And Len(remoto) > 0 Then If local <> remoto Then Exit Function

    local = Comparable(personal.RFC): remoto = Comparable(TextoSQL(rs, "rfc"))
    If Len(local) > 0 And Len(remoto) > 0 Then If local <> remoto Then Exit Function

    local = Comparable(Otros_Rgtros.curp): remoto = Comparable(TextoSQL(rs, "curp"))
    If Len(local) > 0 And Len(remoto) > 0 Then If local <> remoto Then Exit Function

    'Si el registro ya tiene datos de nomina pero ninguna identidad
    'local, no asociar ciegamente la persona recibida de SQL.
    If Len(TextoLocal(personal.nom)) = 0 And _
       Len(TextoLocal(personal.ape1)) = 0 And _
       Len(TextoLocal(personal.ape2)) = 0 And _
       Len(TextoLocal(personal.RFC)) = 0 And _
       Len(TextoLocal(Otros_Rgtros.curp)) = 0 Then
        If Len(TextoLocal(personal.fal)) > 0 Or _
           Len(TextoLocal(personal.imss)) > 0 Or _
           Len(TextoLocal(personal.fab)) > 0 Or _
           personal.ingr <> 0 Or personal.viat <> 0 Or personal.otras <> 0 Then Exit Function
    End If

    CoincideIdentidad = True
End Function

'Completa solo si el valor local esta vacio y el valor SQL cabe.
'Para los UDT de ancho fijo, pasar el largo explicitamente.
Private Function DatoNuevo(ByVal actual As String, ByVal remoto As String, _
                            ByVal ancho As Long, ByVal id As Long, _
                            ByVal campo As String) As String
    DatoNuevo = actual
    If Len(TextoLocal(actual)) <> 0 Then Exit Function
    If Len(Trim$(remoto)) = 0 Then Exit Function
    If Len(remoto) > ancho Then
        Debug.Print "ID " & id & ": omitido campo " & campo & " (excede " & ancho & " caracteres)"
        Exit Function
    End If
    If StrConv(StrConv(remoto, vbFromUnicode), vbUnicode) <> remoto Then
        Debug.Print "ID " & id & ": campo " & campo & " no representable en el archivo local."
        Exit Function
    End If
    DatoNuevo = UCase$(remoto)
End Function

Private Function FirmaCampo(ByVal campo As String) As String
    Dim t As String
    t = TextoLocal(campo)
    FirmaCampo = CStr(Len(t)) & ":" & t
End Function

'Firma de los campos efectivamente administrados por este importador.
'Permite verificar la copia temporal inmediatamente despues del Put.
Private Function FirmaRegistroActual() As String
    FirmaRegistroActual = _
        FirmaCampo(personal.nom) & FirmaCampo(personal.RFC) & _
        FirmaCampo(Otros_Rgtros.curp) & FirmaCampo(Empleado_1.Direccion) & _
        FirmaCampo(Empleado_1.Colonia) & FirmaCampo(Empleado_1.Ciudad) & _
        FirmaCampo(Empleado_1.Estado) & FirmaCampo(Empleado_1.Delegacion) & _
        FirmaCampo(Empleado_1.Cpostal) & FirmaCampo(Empleado_1.correo) & _
        FirmaCampo(personal.ape1) & FirmaCampo(personal.ape2)
End Function

'Agrega bytes cero a un archivo .dno de registros fijos.
'Los registros intermedios quedan VACIOS (nunca se recorren posiciones).
Private Sub AsegurarLongitud(ByVal ruta As String, ByVal longitud As Long, ByVal idMaximo As Long)
    Dim h As Integer, objetivo As Long, cero As Byte
    objetivo = longitud * idMaximo
    If FileLen(ruta) >= objetivo Then Exit Sub
    h = FreeFile
    On Error GoTo Fallo
    Open ruta For Binary Access Read Write As #h
    cero = 0
    Put #h, objetivo, cero
    Close #h
    Exit Sub
Fallo:
    On Error Resume Next
    Close #h
    On Error GoTo 0
    Err.Raise vbObjectError + 2010, , "No se pudo ampliar la copia temporal: " & ruta
End Sub

'Devuelve la cantidad de empleados nuevos, por referencia.
'Errores: no se publica ningun archivo nuevo salvo que todas las
'copias temporales se hayan preparado correctamente.
Public Sub ImportarPersonalDesdeSQL(ByVal empresaSAT As String, _
                                   ByRef nuevos As Long, _
                                   ByRef completados As Long, _
                                   ByRef omitidos As Long)
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim veces(1 To MAX_ID_IMPORTACION) As Integer
    Dim id As Long, maxID As Long
    Dim i As Integer, h As Integer
    Dim hPer As Integer, hOtros As Integer, hCfdi As Integer
    Dim originales(1 To 6) As String
    Dim temporales(1 To 6) As String
    Dim respaldos(1 To 6) As String
    Dim antiguos(1 To 6) As String
    Dim anchos(1 To 6) As Long
    Dim existia(1 To 6) As Boolean
    Dim desplazado(1 To 6) As Boolean
    Dim instalado(1 To 6) As Boolean
    Dim sello As String, prefijo As String
    Dim carpeta As String
    Dim totalAnterior As Long
    Dim datosAntes As String, datosDespues As String
    Dim habiaPersona As Boolean
    Dim errorN As Long, errorTexto As String
    Dim sNombre As String, sApe1 As String, sApe2 As String, sRFC As String

    nuevos = 0: completados = 0: omitidos = 0
    On Error GoTo ErrorImportacion

    If EsEmpresaLocal(emp) Then Exit Sub
    If Trim$(empresaSAT) = "" Then Err.Raise vbObjectError + 2001, , "Empresa SAT no identificada."
    If con Is Nothing Then Err.Raise vbObjectError + 2002, , "No existe la conexion SQL."
    If con.State <> adStateOpen Then Err.Raise vbObjectError + 2003, , "La conexion SQL esta cerrada."

    carpeta = CurDir$
    If Right$(carpeta, 1) <> "\" Then carpeta = carpeta & "\"

    originales(1) = carpeta & "personal.dno"
    originales(2) = carpeta & "PerOtre.dno"
    originales(3) = carpeta & "Perscfdi.dno"
    originales(4) = carpeta & "maestro.dno"
    originales(5) = carpeta & "Bnxcla.dno"
    originales(6) = carpeta & "deon.dno"

    anchos(1) = Len(personal)
    anchos(2) = Len(Otros_Rgtros)
    anchos(3) = Len(Empleado_1)
    anchos(4) = Len(maestro)
    anchos(5) = Len(Clbnx)
    anchos(6) = Len(DEON)

    If Dir$(originales(1)) = "" Then Err.Raise vbObjectError + 2004, , "Falta personal.dno."
    For i = 1 To 6
        If anchos(i) < 1 Then Err.Raise vbObjectError + 2005, , "Longitud de registro no valida."
        existia(i) = (Dir$(originales(i)) <> "")
        If existia(i) Then
            If FileLen(originales(i)) Mod anchos(i) <> 0 Then
                Err.Raise vbObjectError + 2006, , "Archivo desalineado: " & originales(i)
            End If
        End If
    Next i
    totalAnterior = FileLen(originales(1)) \ anchos(1)

    'UNICA operacion SQL: SELECT. No se escribe jamas en la BD.
    sql = "SELECT idNomina, rfc, curp, nombre, apellidoP, apellidoM, " & _
          "nombreDeVialidad, nombreColonia, nombreLocalidad, " & _
          "entidadFederativa, nombreMunicipio, cp, correoElectronico " & _
          "FROM datosSat WHERE empresa = '" & Replace(empresaSAT, "'", "''") & "' " & _
          "ORDER BY idNomina"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, con, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then GoTo Salida

    'No escoger arbitrariamente uno entre dos registros del mismo ID.
    rs.MoveFirst
    Do While Not rs.EOF
        id = IdSeguro(rs.Fields("idNomina").Value)
        If id <> 0 Then
            If veces(id) < 2 Then veces(id) = veces(id) + 1
            If id > maxID Then maxID = id
        Else
            omitidos = omitidos + 1
            Debug.Print "Registro SQL omitido: idNomina invalido o superior a 32767."
        End If
        rs.MoveNext
    Loop
    If maxID = 0 Then GoTo Salida
    If totalAnterior > MAX_ID_IMPORTACION Then
        Err.Raise vbObjectError + 2012, , "Personal.dno supera el limite de 32767 posiciones utilizado por Form2/Form4."
    End If
    If totalAnterior > maxID Then maxID = totalAnterior

    'Se prepara TODO sobre copias temporales; nunca directamente
    'sobre los archivos locales originales.
    sello = Format$(Now, "yyyymmdd_hhnnss")
    prefijo = carpeta & "SYNC_" & sello
    i = 0
    Do While Dir$(prefijo & "_1.tmp") <> "" Or Dir$(prefijo & "_1.bak") <> "" Or _
             Dir$(prefijo & "_1.old") <> ""
        i = i + 1
        prefijo = carpeta & "SYNC_" & sello & "_" & CStr(i)
    Loop

    For i = 1 To 6
        temporales(i) = prefijo & "_" & CStr(i) & ".tmp"
        respaldos(i) = prefijo & "_" & CStr(i) & ".bak"
        antiguos(i) = prefijo & "_" & CStr(i) & ".old"
        If existia(i) Then
            FileCopy originales(i), temporales(i)
        Else
            h = FreeFile
            Open temporales(i) For Binary As #h
            Close #h
        End If
        'Incluso los otros archivos deben tener la posicion del
        'nuevo empleado, para que Form4 no lea registros inexistentes.
        AsegurarLongitud temporales(i), anchos(i), maxID
    Next i

    hPer = FreeFile
    Open temporales(1) For Random Access Read Write As #hPer Len = anchos(1)
    hOtros = FreeFile
    Open temporales(2) For Random Access Read Write As #hOtros Len = anchos(2)
    hCfdi = FreeFile
    Open temporales(3) For Random Access Read Write As #hCfdi Len = anchos(3)

    rs.MoveFirst
    Do While Not rs.EOF
        id = IdSeguro(rs.Fields("idNomina").Value)
        If id = 0 Then GoTo Siguiente
        If veces(id) <> 1 Then
            omitidos = omitidos + 1
            Debug.Print "ID " & id & ": registros SQL duplicados; no se importa."
            GoTo Siguiente
        End If

        Get #hPer, id, personal
        Get #hOtros, id, Otros_Rgtros
        Get #hCfdi, id, Empleado_1

        If Not CoincideIdentidad(rs) Then
            omitidos = omitidos + 1
            Debug.Print "ID " & id & ": conflicto de identidad entre SQL y archivos locales."
            GoTo Siguiente
        End If

        'El nombre es obligatorio para dar de alta al empleado.
        sNombre = TextoSQL(rs, "nombre")
        sApe1 = TextoSQL(rs, "apellidoP")
        sApe2 = TextoSQL(rs, "apellidoM")
        sRFC = TextoSQL(rs, "rfc")
        If Comparable(sApe1) = "N/A" Then sApe1 = ""
        If Comparable(sApe2) = "N/A" Then sApe2 = ""
        If Len(TextoLocal(personal.nom)) = 0 Then
            If sNombre = "" Or (sApe1 = "" And sApe2 = "") Then
                omitidos = omitidos + 1
                Debug.Print "ID " & id & ": faltan nombre/apellidos para alta."
                GoTo Siguiente
            End If
        End If
        If Len(sNombre) > Len(personal.nom) Or _
           Len(sApe1) > Len(personal.ape1) Or _
           Len(sApe2) > Len(personal.ape2) Or _
           Len(sRFC) > Len(personal.RFC) Or _
           Len(TextoSQL(rs, "curp")) > Len(Otros_Rgtros.curp) Then
            omitidos = omitidos + 1
            Debug.Print "ID " & id & ": campo de identidad excede longitud local."
            GoTo Siguiente
        End If

        habiaPersona = (Len(TextoLocal(personal.nom)) > 0 Or _
                        Len(TextoLocal(personal.RFC)) > 0)

        'Detectar si hay algo nuevo (sin tocar campos que existan).
        datosAntes = FirmaRegistroActual()

        personal.nom = DatoNuevo(personal.nom, sNombre, Len(personal.nom), id, "nombre")
        personal.ape1 = DatoNuevo(personal.ape1, sApe1, Len(personal.ape1), id, "apellidoP")
        personal.ape2 = DatoNuevo(personal.ape2, sApe2, Len(personal.ape2), id, "apellidoM")
        personal.RFC = DatoNuevo(personal.RFC, sRFC, Len(personal.RFC), id, "rfc")
        Otros_Rgtros.curp = DatoNuevo(Otros_Rgtros.curp, TextoSQL(rs, "curp"), Len(Otros_Rgtros.curp), id, "curp")

        Empleado_1.Direccion = DatoNuevo(Empleado_1.Direccion, TextoSQL(rs, "nombreDeVialidad"), Len(Empleado_1.Direccion), id, "Direccion")
        Empleado_1.Colonia = DatoNuevo(Empleado_1.Colonia, TextoSQL(rs, "nombreColonia"), Len(Empleado_1.Colonia), id, "Colonia")
        Empleado_1.Ciudad = DatoNuevo(Empleado_1.Ciudad, TextoSQL(rs, "nombreLocalidad"), Len(Empleado_1.Ciudad), id, "Ciudad")
        Empleado_1.Estado = DatoNuevo(Empleado_1.Estado, TextoSQL(rs, "entidadFederativa"), Len(Empleado_1.Estado), id, "Estado")
        Empleado_1.Delegacion = DatoNuevo(Empleado_1.Delegacion, TextoSQL(rs, "nombreMunicipio"), Len(Empleado_1.Delegacion), id, "Delegacion")
        Empleado_1.Cpostal = DatoNuevo(Empleado_1.Cpostal, TextoSQL(rs, "cp"), Len(Empleado_1.Cpostal), id, "Cpostal")
        Empleado_1.correo = DatoNuevo(Empleado_1.correo, TextoSQL(rs, "correoElectronico"), Len(Empleado_1.correo), id, "correo")

        datosDespues = FirmaRegistroActual()

        If datosAntes <> datosDespues Then
            Put #hPer, id, personal
            Put #hOtros, id, Otros_Rgtros
            Put #hCfdi, id, Empleado_1

            'Comprobar lo escrito antes de publicar los archivos.
            Get #hPer, id, personal
            Get #hOtros, id, Otros_Rgtros
            Get #hCfdi, id, Empleado_1
            If FirmaRegistroActual() <> datosDespues Then
                Err.Raise vbObjectError + 2013, , "Verificacion de datos fallida para ID " & CStr(id)
            End If

            If habiaPersona Then
                completados = completados + 1
            Else
                nuevos = nuevos + 1
            End If
            Debug.Print "ID " & id & IIf(habiaPersona, ": datos faltantes agregados.", ": alta local desde SQL.")
        End If
Siguiente:
        rs.MoveNext
    Loop

    Close #hPer: hPer = 0
    Close #hOtros: hOtros = 0
    Close #hCfdi: hCfdi = 0

    If nuevos = 0 And completados = 0 Then GoTo Salida

    'Respaldos PERMANENTES antes de cambiar algun archivo.
    For i = 1 To 6
        If existia(i) Then FileCopy originales(i), respaldos(i)
    Next i

    'Publicar las copias preparadas. Ante error se intenta restaurar
    'el conjunto previo usando los originales renombrados.
    For i = 1 To 6
        If existia(i) Then
            Name originales(i) As antiguos(i)
            desplazado(i) = True
        End If
        Name temporales(i) As originales(i)
        instalado(i) = True
    Next i

    On Error Resume Next
    For i = 1 To 6
        If desplazado(i) And Dir$(antiguos(i)) <> "" Then Kill antiguos(i)
    Next i
    On Error GoTo ErrorImportacion
    Debug.Print "Sincronizacion local: " & nuevos & " nuevos; " & _
                completados & " completados; " & omitidos & " omitidos."
    GoTo Salida

ErrorImportacion:
    errorN = Err.Number
    errorTexto = Err.Description
    On Error Resume Next
    If hPer > 0 Then Close #hPer
    If hOtros > 0 Then Close #hOtros
    If hCfdi > 0 Then Close #hCfdi
    If h > 0 Then Close #h
    'Primero retirar lo publicado en esta ejecucion.
    For i = 1 To 6
        If instalado(i) Then
            If Dir$(originales(i)) <> "" Then Kill originales(i)
        End If
    Next i
    'Reponer desde el nombre previo al reemplazo.
    For i = 1 To 6
        If desplazado(i) Then
            If Dir$(antiguos(i)) <> "" Then Name antiguos(i) As originales(i)
        End If
    Next i
    'No borrar archivos .bak para permitir rescate manual.
    On Error GoTo 0
    Err.Raise errorN, "ImportarPersonalDesdeSQL", errorTexto & _
              ". Revise los respaldos SYNC_*.bak si hubo una falla durante la publicacion."

Salida:
    On Error Resume Next
    If hPer > 0 Then Close #hPer
    If hOtros > 0 Then Close #hOtros
    If hCfdi > 0 Then Close #hCfdi
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    Set rs = Nothing
    'Eliminar temporales si no fueron publicados. Nunca eliminar backups.
    For i = 1 To 6
        If temporales(i) <> "" Then
            If Dir$(temporales(i)) <> "" Then Kill temporales(i)
        End If
    Next i
    On Error GoTo 0
End Sub


