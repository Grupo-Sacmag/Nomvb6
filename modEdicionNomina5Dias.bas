Attribute VB_Name = "modEdicionNomina5Dias"
Option Explicit

'=========================================================
' CONTROL DE EDICION DE NOMINAS GUARDADAS
'
' REGLAS:
'
' 1. El plazo inicia con el PRIMER GUARDADO.
' 2. Se permiten 5 dias habiles posteriores.
' 3. No se utilizan las fechas de NOM ni CMP.
' 4. Una edicion NO reinicia el plazo.
' 5. Una nomina sin registro NO puede editarse.
' 6. Si se detecta vencimiento, el bloqueo persiste.
'
' IMPORTANTE:
' Este control es local y utiliza el Registro de Windows.
' Protege frente al cambio de fechas de NOM/CMP, pero
' no sustituye una autorizacion desde servidor.
'=========================================================

Private Const DIAS_EDICION As Long = 5

'=========================================================
' CONTROL BINARIO DE FECHAS CONTABLES DE NOMINA
'
' Archivo:
' C:\GconTA\FeCoNom.dno
'
' Este archivo almacenara:
'
' - Ruta completa del archivo NOM.
' - Fecha del primer guardado.
' - Fecha de la ultima revision.
' - Fecha limite para editar.
' - Estado de vencimiento.
'
' No utiliza DateCreated ni DateLastModified.
'
' IMPORTANTE:
' Los registros nuevos se agregan al archivo.
' Los registros existentes nunca reinician su fecha.
'=========================================================

Private Const RUTA_FECONOM As String = "C:\GconTA\FeCoNom.dno"

Private Const FIRMA_FECONOM As String = "FECON001"

Private Type RegistroFeCoNom
    Firma As String * 8
    RutaNomina As String * 260
    fechaInicio As Long
    FechaUltimaRevision As Long
    fechaLimite As Long
    Vencido As Integer
End Type

'=========================================================
' CONVERTIR FECHAS GUARDADAS EN FORMATO AAAA-MM-DD
'=========================================================

Private Function LeerFechaControl(ByVal texto As String) As Date

    Dim fechaTemporal As Date

    If Len(texto) <> 10 Then GoTo FechaInvalida
    If mID$(texto, 5, 1) <> "-" Then GoTo FechaInvalida
    If mID$(texto, 8, 1) <> "-" Then GoTo FechaInvalida

    If Not IsNumeric(Left$(texto, 4)) Then
        GoTo FechaInvalida
    End If

    If Not IsNumeric(mID$(texto, 6, 2)) Then
        GoTo FechaInvalida
    End If

    If Not IsNumeric(Right$(texto, 2)) Then
        GoTo FechaInvalida
    End If

    fechaTemporal = DateSerial(CInt(Left$(texto, 4)), CInt(mID$(texto, 6, 2)), CInt(Right$(texto, 2)))

    If Format$(fechaTemporal, "yyyy-mm-dd") <> texto Then
        GoTo FechaInvalida
    End If

    LeerFechaControl = fechaTemporal

    Exit Function
    
FechaInvalida:
    Err.Raise vbObjectError + 7102, "LeerFechaControl", "La fecha del control de nomina es invalida."

End Function

'=========================================================
' DETERMINAR EL PRIMER LUNES DE UN MES
'=========================================================

Private Function PrimerLunesControl(ByVal anio As Long, ByVal mes As Long) As Date

    Dim primero As Date

    primero = DateSerial(anio, mes, 1)

    PrimerLunesControl = primero + _
        ((8 - Weekday(primero, vbMonday)) Mod 7)

End Function

'=========================================================
' IDENTIFICAR DIAS NO HABILES
'
' Sabados, domingos y descansos generales LFT.
'
' Otros dias inhábiles de la empresa o electorales
' requeriran una configuracion adicional.
'=========================================================

Private Function EsDiaInhabilControl(ByVal fecha As Date) As Boolean

    Dim anio As Long
    Dim mes As Long
    Dim dia As Long

    anio = Year(fecha)
    mes = Month(fecha)
    dia = Day(fecha)

    EsDiaInhabilControl = True

    'Sabado o domingo.
    If Weekday(fecha, vbMonday) >= 6 Then Exit Function

    'Fechas fijas.
    If mes = 1 And dia = 1 Then Exit Function
    If mes = 5 And dia = 1 Then Exit Function
    If mes = 9 And dia = 16 Then Exit Function
    If mes = 12 And dia = 25 Then Exit Function

    'Primer lunes de febrero.
    If fecha = PrimerLunesControl(anio, 2) Then
        Exit Function
    End If

    'Tercer lunes de marzo.
    If fecha = PrimerLunesControl(anio, 3) + 14 Then
        Exit Function
    End If

    'Tercer lunes de noviembre.
    If fecha = PrimerLunesControl(anio, 11) + 14 Then
        Exit Function
    End If

    'Transmision del Poder Ejecutivo Federal.
    If mes = 10 And dia = 1 And anio >= 2024 Then

        If (anio - 2024) Mod 6 = 0 Then
            Exit Function
        End If

    End If

    EsDiaInhabilControl = False

End Function

'=========================================================
' OBTENER EL QUINTO DIA HABIL POSTERIOR AL GUARDADO
'=========================================================

Private Function LimiteEdicionNomina(ByVal fechaInicio As Date) As Date

    Dim fechaActual As Date
    Dim diasContados As Long

    fechaActual = DateValue(fechaInicio)
    diasContados = 0

    Do While diasContados < DIAS_EDICION

        fechaActual = DateAdd("d", 1, fechaActual)

        If Not EsDiaInhabilControl(fechaActual) Then
            diasContados = diasContados + 1
        End If

    Loop

    LimiteEdicionNomina = fechaActual

End Function

Public Sub RegistrarPrimerGuardadoNomina(ByVal rutaNom As String)

    Dim rutaNormalizada As String
    Dim rutaCmp As String

    Dim numeroArchivo As Integer
    Dim posicion As Long
    Dim cantidadRegistros As Long
    Dim generaciones As Long

    Dim primeraFecha As Long
    Dim ultimaRevisionHistorica As Long
    Dim vencidoHistorico As Boolean

    Dim registro As RegistroFeCoNom

    Dim fechaActual As Date
    Dim fechaOriginal As Date
    Dim fechaLimiteOriginal As Date

    Dim numeroError As Long
    Dim descripcionError As String

    On Error GoTo ErrorRegistro

    '=========================================================
    ' 1. IDENTIFICAR LA NOMINA
    '=========================================================

    rutaNormalizada = NormalizarRutaFeCoNom(rutaNom)
    rutaCmp = Left$(rutaNormalizada, Len(rutaNormalizada) - 4) & ".CMP"

    '=========================================================
    ' 2. COMPROBAR EXISTENCIA DE NOM Y CMP
    '
    ' La funcion solamente debe ejecutarse despues
    ' del primer guardado de una captura nueva.
    '=========================================================

    If Len(Dir$(rutaNormalizada)) = 0 Then
        Err.Raise vbObjectError + 7210, "RegistrarPrimerGuardadoNomina", "No existe el archivo NOM."
    End If

    If FileLen(rutaNormalizada) <= 0 Then
        Err.Raise vbObjectError + 7211, "RegistrarPrimerGuardadoNomina", "No se puede registrar una nomina vacia."
    End If

    If Len(Dir$(rutaCmp)) = 0 Then
        Err.Raise vbObjectError + 7212, "RegistrarPrimerGuardadoNomina", "No existe el complemento CMP."
    End If

    '=========================================================
    ' 3. COMPROBAR CARPETA DEL HISTORIAL
    '=========================================================

    If Len(Dir$("C:\GconTA", vbDirectory)) = 0 Then
        Err.Raise vbObjectError + 7213, "RegistrarPrimerGuardadoNomina", "No existe la carpeta C:\GconTA."
    End If

    '=========================================================
    ' 4. ABRIR EXCLUSIVAMENTE FECONOM.DNO
    '
    ' No consultar ni escribir el Registro de Windows.
    '
    ' Si el archivo no existe, se crea.
    '=========================================================

    numeroArchivo = FreeFile
    Open RUTA_FECONOM For Random Access Read Write Lock Read Write As #numeroArchivo Len = Len(registro)

    '=========================================================
    ' 5. BUSCAR GENERACIONES ANTERIORES
    '
    ' El resultado permite conocer:
    '
    ' - Numero de generaciones anteriores.
    ' - Fecha de la primera generacion.
    ' - Ultima revision registrada.
    ' - Si ya existio un vencimiento.
    '=========================================================

    posicion = BuscarRegistroFeCoNom(numeroArchivo, rutaNormalizada, registro, generaciones, primeraFecha, ultimaRevisionHistorica, vencidoHistorico)
    fechaActual = Date

    '=========================================================
    ' 6. DETERMINAR LA FECHA ORIGINAL
    '
    ' GENERACION 1:
    ' El primer guardado es la fecha actual.
    '
    ' GENERACION 2 O SUPERIOR:
    ' Conservar el primer guardado historico.
    '
    ' La nueva generacion NO renueva el plazo.
    '=========================================================

    If generaciones = 0 Then
        fechaOriginal = fechaActual
    Else
        fechaOriginal = FechaDesdeClaveFeCoNom(primeraFecha)

        '-----------------------------------------------------
        ' COMPROBAR RETROCESO DE FECHA
        '-----------------------------------------------------

        If fechaActual < FechaDesdeClaveFeCoNom(ultimaRevisionHistorica) Then
            Err.Raise vbObjectError + 7216, "RegistrarPrimerGuardadoNomina", "La fecha actual es anterior a una revision registrada."

        End If

    End If

    fechaLimiteOriginal = LimiteEdicionNomina(fechaOriginal)

    '=========================================================
    ' 7. PREPARAR NUEVA GENERACION
    '
    ' Cada registro contiene la fecha en que se
    ' genero esa version de la nomina.
    '
    ' Ejemplo:
    ' Generacion 1 -> NOV12026
    ' Generacion 2 -> NOV12026 (1)
    ' Generacion 3 -> NOV12026 (2)
    '
    ' El nombre fisico del NOM no cambia.
    '=========================================================

    registro.Firma = FIRMA_FECONOM
    registro.RutaNomina = rutaNormalizada
    registro.fechaInicio = ClaveFechaFeCoNom(fechaActual)
    registro.FechaUltimaRevision = ClaveFechaFeCoNom(fechaActual)
    registro.fechaLimite = ClaveFechaFeCoNom(LimiteEdicionNomina(fechaActual))
    registro.Vencido = 0

    '=========================================================
    ' 8. CONSERVAR BLOQUEO HISTORICO
    '
    ' Si alguna generacion ya vencio o termino
    ' el plazo original, la nueva generacion
    ' quedara igualmente bloqueada para edicion.
    '=========================================================

    If vencidoHistorico Or fechaActual > fechaLimiteOriginal Then
        registro.Vencido = 1
    End If

    '=========================================================
    ' 9. AGREGAR LA NUEVA GENERACION AL FINAL
    '
    ' Nunca sobrescribir registros anteriores.
    '
    ' El contador sera:
    '
    ' generaciones + 1
    '
    ' No modificar NOM ni CMP.
    '=========================================================

    cantidadRegistros = LOF(numeroArchivo) \ Len(registro)
    Put #numeroArchivo, cantidadRegistros + 1, registro

    '=========================================================
    ' 10. CERRAR HISTORIAL
    '=========================================================

    Close #numeroArchivo
    numeroArchivo = 0

    Exit Sub

ErrorRegistro:

    numeroError = Err.Number
    descripcionError = Err.Description

    On Error Resume Next

    If numeroArchivo <> 0 Then
        Close #numeroArchivo
    End If

    On Error GoTo 0

    Err.Raise numeroError, "RegistrarPrimerGuardadoNomina", descripcionError
End Sub

Public Function PuedeModificarNomina5Dias(ByVal rutaNom As String, ByRef detalle As String) As Boolean

    Dim rutaNormalizada As String

    Dim numeroArchivo As Integer
    Dim posicion As Long
    Dim generaciones As Long

    Dim primeraFecha As Long
    Dim ultimaRevisionHistorica As Long
    Dim vencidoHistorico As Boolean

    Dim registro As RegistroFeCoNom

    Dim fechaInicio As Date
    Dim fechaUltima As Date
    Dim fechaLimiteEdicion As Date
    Dim fechaRevisionActual As Date

    Dim numeroError As Long
    Dim descripcionError As String

    On Error GoTo ErrorControl

    PuedeModificarNomina5Dias = False
    detalle = ""

    '=========================================================
    ' 1. IDENTIFICAR LA NOMINA
    '=========================================================

    rutaNormalizada = NormalizarRutaFeCoNom(rutaNom)

    '=========================================================
    ' 2. COMPROBAR HISTORIAL BINARIO
    '
    ' Si el archivo no existe:
    ' SOLO CONSULTA.
    '
    ' No crear registros al abrir una nomina.
    '=========================================================

    If Len(Dir$(RUTA_FECONOM)) = 0 Then
        detalle = "No existe un historial de fechas para autorizar modificaciones."
        Exit Function
    End If

    '=========================================================
    ' 3. ABRIR FECONOM.DNO
    '=========================================================

    numeroArchivo = FreeFile

    Open RUTA_FECONOM For Random Access Read Write Lock Read Write As #numeroArchivo Len = Len(registro)

    '=========================================================
    ' 4. BUSCAR LA NOMINA Y SUS GENERACIONES
    '=========================================================

    posicion = BuscarRegistroFeCoNom(numeroArchivo, rutaNormalizada, registro, generaciones, primeraFecha, ultimaRevisionHistorica, vencidoHistorico)

    If posicion = 0 Then
        detalle = "La nomina no tiene una fecha de primer guardado registrada." & vbCrLf & "Solo se permite su consulta."
        Close #numeroArchivo
        numeroArchivo = 0
        Exit Function
    End If

    '=========================================================
    ' 5. RECUPERAR EL PLAZO ORIGINAL
    '
    ' IMPORTANTE:
    '
    ' Se utiliza SIEMPRE la fecha de la primera
    ' generacion, no la de la ultima.
    '
    ' Eliminar y recrear NOM o CMP no modifica
    ' este plazo.
    '=========================================================

    fechaInicio = FechaDesdeClaveFeCoNom(primeraFecha)
    fechaUltima = FechaDesdeClaveFeCoNom(ultimaRevisionHistorica)
    fechaLimiteEdicion = LimiteEdicionNomina(fechaInicio)

    '=========================================================
    ' 6. CONSERVAR REGLA HISTORICA
    '
    ' Nominas anteriores a la segunda quincena
    ' de noviembre de 2026:
    ' SOLO CONSULTA.
    '=========================================================

    If NominaHistoricaSoloConsulta(rutaNormalizada, fechaInicio) Then
        detalle = "Nomina historica." & vbCrLf & "Solo se permite su consulta."
        Close #numeroArchivo
        numeroArchivo = 0
        Exit Function
    End If

    '=========================================================
    ' 7. COMPROBAR BLOQUEOS ANTERIORES
    '
    ' Una nueva generacion no puede reactivar
    ' una nomina cuyo plazo ya fue cerrado.
    '=========================================================

    If vencidoHistorico Then
        detalle = "La nomina se encuentra fuera del periodo permitido para edicion."
        Close #numeroArchivo
        numeroArchivo = 0
        Exit Function
    End If

    '=========================================================
    ' 8. DETECTAR RETROCESO DEL RELOJ
    '
    ' Se utiliza la ultima fecha registrada
    ' entre todas las generaciones.
    '=========================================================

    If Date < fechaInicio Or Date < fechaUltima Then
        detalle = "Se detecto una inconsistencia en la fecha de revision." & vbCrLf & "No se autoriza editar esta nomina."
        Close #numeroArchivo
        numeroArchivo = 0

        Exit Function
    End If

    '=========================================================
    ' 9. VERIFICAR LOS CINCO DIAS HABILES
    '
    ' Siempre desde la primera generacion.
    '
    ' No importa cuantas veces fue recreada.
    '=========================================================

    If Date > fechaLimiteEdicion Then

        '-------------------------------------------------
        ' MARCAR ULTIMA GENERACION COMO VENCIDA
        '-------------------------------------------------

        registro.Vencido = 1
        registro.FechaUltimaRevision = ClaveFechaFeCoNom(Date)

        Put #numeroArchivo, posicion, registro

        Close #numeroArchivo
        numeroArchivo = 0

        detalle = "El plazo de cinco dias habiles ha vencido." & vbCrLf & "Solo se permite su consulta."
        Exit Function
    End If

    '=========================================================
    ' 10. ACTUALIZAR LA ULTIMA REVISION
    '
    ' Solo se actualiza el registro de la
    ' generacion mas reciente.
    '
    ' No cambiar:
    ' - Primera fecha historica.
    ' - Fecha de generacion.
    ' - Fecha limite.
    ' - Numero de generaciones.
    '=========================================================

    fechaRevisionActual = FechaDesdeClaveFeCoNom(registro.FechaUltimaRevision)
    
    If Date > fechaRevisionActual Then
        registro.FechaUltimaRevision = ClaveFechaFeCoNom(Date)
        Put #numeroArchivo, posicion, registro
    End If

    '=========================================================
    ' 11. CERRAR HISTORIAL
    '=========================================================

    Close #numeroArchivo
    numeroArchivo = 0

    '=========================================================
    ' 12. AUTORIZAR EDICION
    '=========================================================

    PuedeModificarNomina5Dias = True
    detalle = "Edicion permitida dentro de los cinco dias habiles." & vbCrLf & "Fecha limite: " & Format$(fechaLimiteEdicion, "dd/mm/yyyy")
    Exit Function

ErrorControl:

    numeroError = Err.Number
    descripcionError = Err.Description

    On Error Resume Next

    If numeroArchivo <> 0 Then
        Close #numeroArchivo
    End If
    On Error GoTo 0

    PuedeModificarNomina5Dias = False
    detalle = "No fue posible verificar el historial de la nomina." & vbCrLf & "No se autorizan modificaciones." & vbCrLf & vbCrLf & "Error " & CStr(numeroError) & ": " & descripcionError

End Function

'=========================================================
' NORMALIZAR LA RUTA DE UNA NOMINA
'
' La identificacion depende de la ruta completa,
' no de la fecha de creacion del archivo.
'=========================================================

Private Function NormalizarRutaFeCoNom(ByVal rutaNom As String) As String

    Dim ruta As String

    ruta = UCase$(Trim$(rutaNom))
    
    If Len(ruta) = 0 Then
        Err.Raise vbObjectError + 7200, "NormalizarRutaFeCoNom", "La ruta de nomina esta vacia."
    End If

    If InStrRev(ruta, "\") = 0 Then
        Err.Raise vbObjectError + 7201, "NormalizarRutaFeCoNom", "Se requiere la ruta completa del archivo NOM."
    End If

    If Right$(ruta, 4) <> ".NOM" Then
        Err.Raise vbObjectError + 7202, "NormalizarRutaFeCoNom", "El archivo debe tener extension NOM."
    End If

    If Len(ruta) > 260 Then
        Err.Raise vbObjectError + 7203, "NormalizarRutaFeCoNom", "La ruta supera la longitud permitida."
    End If
    NormalizarRutaFeCoNom = ruta

End Function

'=========================================================
' CONVERTIR FECHA A NUMERO AAAAMMDD
'
' Ejemplo:
'
' 26/11/2026 -> 20261126
'
' Este numero es informacion interna del registro.
'=========================================================

Private Function ClaveFechaFeCoNom(ByVal fecha As Date) As Long
    ClaveFechaFeCoNom = CLng(Format$(DateValue(fecha), "yyyymmdd"))
End Function

'=========================================================
' CONVERTIR NUMERO AAAAMMDD A FECHA
'
' Utiliza LeerFechaControl(), que ya existe
' en modEdicionNomina5Dias.bas.
'=========================================================

Private Function FechaDesdeClaveFeCoNom(ByVal numeroFecha As Long) As Date

    Dim texto As String

    If numeroFecha < 19000101 Or _
       numeroFecha > 99991231 Then
        Err.Raise vbObjectError + 7204, "FechaDesdeClaveFeCoNom", "Fecha numerica fuera de rango."
    End If

    texto = Format$(numeroFecha, "00000000")
    FechaDesdeClaveFeCoNom = LeerFechaControl(Left$(texto, 4) & "-" & mID$(texto, 5, 2) & "-" & Right$(texto, 2))
End Function

Private Function BuscarRegistroFeCoNom(ByVal numeroArchivo As Integer, ByVal rutaNormalizada As String, ByRef registro As RegistroFeCoNom, ByRef generaciones As Long, ByRef primeraFecha As Long, ByRef ultimaRevisionHistorica As Long, ByRef vencidoHistorico As Boolean) As Long

    Dim temporal As RegistroFeCoNom
    Dim cantidad As Long
    Dim posicion As Long

    '=========================================================
    ' BUSQUEDA DE HISTORIAL POR GENERACIONES
    '
    ' Una misma ruta puede aparecer varias veces.
    '
    ' Primera aparicion:
    '   Generacion 1
    '
    ' Segunda aparicion:
    '   Generacion 2
    '
    ' Tercera aparicion:
    '   Generacion 3
    '
    ' No se eliminan registros anteriores.
    '=========================================================

    BuscarRegistroFeCoNom = 0

    generaciones = 0
    primeraFecha = 0
    ultimaRevisionHistorica = 0
    vencidoHistorico = False

    '=====================================================
    ' 1. VALIDAR LONGITUD DEL ARCHIVO
    '=====================================================

    If LOF(numeroArchivo) Mod Len(temporal) <> 0 Then
        Err.Raise vbObjectError + 7205, "BuscarRegistroFeCoNom", "El historial contiene registros incompletos."
    End If

    cantidad = LOF(numeroArchivo) \ Len(temporal)

    '=====================================================
    ' 2. RECORRER EL HISTORIAL COMPLETO
    '=====================================================

    For posicion = 1 To cantidad
        Get #numeroArchivo, posicion, temporal

        '-------------------------------------------------
        ' VALIDAR EL REGISTRO BINARIO
        '-------------------------------------------------

        If Not RegistroValidoFeCoNom(temporal) Then
            Err.Raise vbObjectError + 7206, "BuscarRegistroFeCoNom", "Se detecto un registro invalido en el historial de nominas."
        End If

        '-------------------------------------------------
        ' COMPROBAR SI PERTENECE A ESTA NOMINA
        '-------------------------------------------------

        If StrComp(UCase$(Trim$(temporal.RutaNomina)), rutaNormalizada, vbBinaryCompare) = 0 Then
            generaciones = generaciones + 1

            '---------------------------------------------
            ' PRIMERA GENERACION
            '
            ' Conservar la fecha mas antigua.
            '---------------------------------------------

            If generaciones = 1 Then
                primeraFecha = temporal.fechaInicio
            Else
                'Una generacion posterior no puede tener una fecha anterior a la primera.

                If temporal.fechaInicio < primeraFecha Then
                    Err.Raise vbObjectError + 7208, "BuscarRegistroFeCoNom", "El historial tiene fechas de generacion inconsistentes."
                End If
            End If

            '---------------------------------------------
            ' ULTIMA REVISION CONOCIDA
            '
            ' Conservar la fecha maxima observada
            ' entre todas las generaciones.
            '---------------------------------------------

            If temporal.FechaUltimaRevision > ultimaRevisionHistorica Then
                ultimaRevisionHistorica = temporal.FechaUltimaRevision
            End If

            '---------------------------------------------
            ' BLOQUEO HISTORICO
            '
            ' Si alguna generacion fue cerrada,
            ' las posteriores no la reactivan.
            '---------------------------------------------

            If temporal.Vencido = 1 Then
                vencidoHistorico = True
            End If

            '---------------------------------------------
            ' CONSERVAR LA GENERACION MAS RECIENTE
            '---------------------------------------------

            registro = temporal
            BuscarRegistroFeCoNom = posicion
        End If
    Next posicion
End Function

Private Function RegistroValidoFeCoNom(ByRef registro As RegistroFeCoNom) As Boolean

    Dim ruta As String

    Dim fechaInicio As Date
    Dim fechaRevision As Date
    Dim fechaLimite As Date

    On Error GoTo ErrorValidacion

    RegistroValidoFeCoNom = False

    '=========================================================
    ' 1. COMPROBAR FIRMA DEL REGISTRO
    '=========================================================

    If registro.Firma <> FIRMA_FECONOM Then
        Exit Function
    End If

    '=========================================================
    ' 2. COMPROBAR IDENTIFICADOR DE NOMINA
    '=========================================================

    ruta = UCase$(Trim$(registro.RutaNomina))
    
    If Len(ruta) = 0 Then Exit Function
    If Right$(ruta, 4) <> ".NOM" Then
        Exit Function
    End If
    If InStrRev(ruta, "\") = 0 Then
        Exit Function
    End If

    '=========================================================
    ' 3. VALIDAR ESTADO DEL REGISTRO
    '=========================================================

    If registro.Vencido <> 0 And registro.Vencido <> 1 Then
        Exit Function
    End If

    '=========================================================
    ' 4. RECUPERAR FECHAS BINARIAS
    '
    ' No consultar:
    ' - Registro de Windows.
    ' - DateCreated.
    ' - DateLastModified.
    '=========================================================

    fechaInicio = FechaDesdeClaveFeCoNom(registro.fechaInicio)
    fechaRevision = FechaDesdeClaveFeCoNom(registro.FechaUltimaRevision)
    fechaLimite = FechaDesdeClaveFeCoNom(registro.fechaLimite)

    '=========================================================
    ' 5. VALIDAR ORDEN CRONOLOGICO
    '=========================================================

    If fechaRevision < fechaInicio Then
        Exit Function
    End If

    If fechaLimite < fechaInicio Then
        Exit Function
    End If

    '=========================================================
    ' 6. VERIFICAR LOS CINCO DIAS HABILES
    '
    ' El limite de cada registro debe coincidir
    ' con su fecha de generacion.
    '=========================================================

    If fechaLimite <> LimiteEdicionNomina(fechaInicio) Then
        Exit Function
    End If

    '=========================================================
    ' 7. REGISTRO VALIDO
    '=========================================================

    RegistroValidoFeCoNom = True
    Exit Function
ErrorValidacion:
    RegistroValidoFeCoNom = False
End Function

'=========================================================
' CORTE HISTORICO DE EDICION
'
' La segunda quincena de NOVIEMBRE 2026 es
' el primer periodo normal que podra editarse,
' siempre que cumpla el plazo de cinco dias habiles.
'
' Periodos anteriores:
'   SOLO CONSULTA.
'
' La fecha original del primer guardado tambien
' debe ser igual o posterior al 16/11/2026.
'
' No se utilizan las fechas de NOM ni CMP.
'=========================================================

Private Function NominaHistoricaSoloConsulta(ByVal rutaNom As String, ByVal fechaPrimerGuardado As Date) As Boolean

    Dim nombre As String
    Dim mesTexto As String
    Dim numeroQuincena As String

    Dim anio As Integer
    Dim mes As Integer
    Dim diaInicio As Integer

    Dim fechaPeriodo As Date
    Dim fechaCorte As Date

    fechaCorte = DateSerial(2026, 11, 16)

    NominaHistoricaSoloConsulta = False

    '=====================================================
    ' 1. BLOQUEAR REGISTROS ANTERIORES AL CORTE
    '=====================================================

    If DateValue(fechaPrimerGuardado) < fechaCorte Then
        NominaHistoricaSoloConsulta = True
        Exit Function
    End If

    '=====================================================
    ' 2. OBTENER EL NOMBRE DE LA NOMINA
    '=====================================================

    nombre = UCase$(mID$(rutaNom, InStrRev(rutaNom, "\") + 1))

    If Right$(nombre, 4) <> ".NOM" Then
        Exit Function
    End If

    nombre = Left$(nombre, Len(nombre) - 4)

    '=====================================================
    ' 3. IDENTIFICAR NOMINA NORMAL
    '
    ' Formato original:
    ' MMM + QUINCENA + AAAA
    '
    ' Ejemplos:
    ' NOV12026
    ' NOV22026
    ' DIC12026
    '
    ' Las especiales conservan el control
    ' de la fecha de su primer guardado.
    '=====================================================

    If Len(nombre) <> 8 Then
        Exit Function
    End If

    numeroQuincena = mID$(nombre, 4, 1)

    If numeroQuincena <> "1" And numeroQuincena <> "2" Then
        Exit Function
    End If

    If Not IsNumeric(Right$(nombre, 4)) Then
        Exit Function
    End If

    anio = CInt(Right$(nombre, 4))
    mesTexto = Left$(nombre, 3)

    Select Case mesTexto
        Case "ENE"
            mes = 1
        Case "FEB"
            mes = 2
        Case "MAR"
            mes = 3
        Case "ABR"
            mes = 4
        Case "MAY"
            mes = 5
        Case "JUN"
            mes = 6
        Case "JUL"
            mes = 7
        Case "AGO"
            mes = 8
        Case "SEP"
            mes = 9
        Case "OCT"
            mes = 10
        Case "NOV"
            mes = 11
        Case "DIC"
            mes = 12
        Case Else
            Exit Function
    End Select

    If numeroQuincena = "1" Then
        diaInicio = 1
    Else
        diaInicio = 16
    End If

    fechaPeriodo = DateSerial(anio, mes, diaInicio)

    If fechaPeriodo < fechaCorte Then
        NominaHistoricaSoloConsulta = True
    End If
End Function

Public Function GeneracionesRegistradasFeCoNom(ByVal rutaNom As String) As Long

    Dim numeroArchivo As Integer
    Dim rutaNormalizada As String
    Dim registro As RegistroFeCoNom
    Dim posicion As Long
    Dim generaciones As Long
    Dim primeraFecha As Long
    Dim ultimaRevision As Long
    Dim vencidoHistorico As Boolean

    On Error GoTo ErrorConsulta

    GeneracionesRegistradasFeCoNom = 0

    If Len(Dir$(RUTA_FECONOM)) = 0 Then
        Exit Function
    End If
    
    rutaNormalizada = NormalizarRutaFeCoNom(rutaNom)
    numeroArchivo = FreeFile

    Open RUTA_FECONOM For Random Access Read As #numeroArchivo Len = Len(registro)

    posicion = BuscarRegistroFeCoNom(numeroArchivo, rutaNormalizada, registro, generaciones, primeraFecha, ultimaRevision, vencidoHistorico)

    Close #numeroArchivo
    numeroArchivo = 0

    GeneracionesRegistradasFeCoNom = generaciones

    Exit Function
ErrorConsulta:
    Debug.Print "Consulta FeCoNom: "; Err.Description

    On Error Resume Next
    If numeroArchivo <> 0 Then Close #numeroArchivo

End Function

