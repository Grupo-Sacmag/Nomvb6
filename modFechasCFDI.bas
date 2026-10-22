Attribute VB_Name = "modFechasCFDI"
Option Explicit

' Fecha que usara NOMCF2 mientras crea una plantilla 2026.
' Se establece ANTES de NOMCF2.Show, para evitar leer fechas de Form8
' cuando el origen es FormViewer.
Public g_FechaPagoCFDI2026 As String

Private Function FechaCFDI_NombreArchivo(ByVal rutaNomina As String) As String
    Dim p As Long
    p = InStrRev(rutaNomina, "\")
    FechaCFDI_NombreArchivo = Mid$(rutaNomina, p + 1)
End Function

Private Function FechaCFDI_RutaFicha(ByVal rutaNomina As String) As String
    rutaNomina = Trim$(rutaNomina)
    If Len(rutaNomina) < 5 Or UCase$(Right$(rutaNomina, 4)) <> ".NOM" Then
        Err.Raise vbObjectError + 5201, "FechaCFDI_RutaFicha", "La ruta debe apuntar a un archivo .NOM."
    End If
    FechaCFDI_RutaFicha = Left$(rutaNomina, Len(rutaNomina) - 4) & ".FEC"
End Function

' Convierte estrictamente dd/mm/aaaa a Date, sin depender del idioma de Windows.
Public Function FechaCFDI_InterpretarDMY(ByVal texto As String) As Date
    Dim partes() As String
    Dim d As Long, m As Long, a As Long
    Dim f As Date

    texto = Trim$(texto)
    partes = Split(texto, "/")
    If UBound(partes) <> 2 Then GoTo FechaInvalida
    If Not IsNumeric(partes(0)) Then GoTo FechaInvalida
    If Not IsNumeric(partes(1)) Then GoTo FechaInvalida
    If Not IsNumeric(partes(2)) Then GoTo FechaInvalida

    d = CLng(partes(0))
    m = CLng(partes(1))
    a = CLng(partes(2))
    If a < 1900 Or a > 9999 Then GoTo FechaInvalida
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Then GoTo FechaInvalida

    f = DateSerial(a, m, d)
    If Year(f) <> a Or Month(f) <> m Or Day(f) <> d Then GoTo FechaInvalida
    FechaCFDI_InterpretarDMY = f
    Exit Function

FechaInvalida:
    Err.Raise vbObjectError + 5202, "FechaCFDI_InterpretarDMY", _
              "Fecha no valida (se espera dd/mm/aaaa): " & texto
End Function

Public Function FechaCFDI_FormatoDMY(ByVal f As Date) As String
    FechaCFDI_FormatoDMY = Right$("0" & CStr(Day(f)), 2) & "/" & Right$("0" & CStr(Month(f)), 2) & "/" & Right$("0000" & CStr(Year(f)), 4)
End Function

Public Function FechaCFDI_FormatoISO(ByVal f As Date) As String
    FechaCFDI_FormatoISO = CStr(Year(f)) & "-" & Right$("0" & CStr(Month(f)), 2) & "-" & Right$("0" & CStr(Day(f)), 2)
End Function

Private Function FechaCFDI_InterpretarISO(ByVal texto As String) As Date
    Dim a As Long, m As Long, d As Long
    Dim f As Date

    texto = Trim$(texto)
    If Len(texto) <> 10 Then GoTo FechaInvalida
    If Mid$(texto, 5, 1) <> "-" Or Mid$(texto, 8, 1) <> "-" Then GoTo FechaInvalida
    If Not IsNumeric(Left$(texto, 4)) Or Not IsNumeric(Mid$(texto, 6, 2)) Or Not IsNumeric(Right$(texto, 2)) Then GoTo FechaInvalida

    a = CLng(Left$(texto, 4))
    m = CLng(Mid$(texto, 6, 2))
    d = CLng(Right$(texto, 2))
    If a < 1900 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then GoTo FechaInvalida
    f = DateSerial(a, m, d)
    If Year(f) <> a Or Month(f) <> m Or Day(f) <> d Then GoTo FechaInvalida
    FechaCFDI_InterpretarISO = f
    Exit Function

FechaInvalida:
    Err.Raise vbObjectError + 5203, "FechaCFDI_InterpretarISO", "Fecha invalida en ficha de nomina: " & texto
End Function

' Guarda la fecha exacta de captura, SIN alterar el Type nom ni los .cmp.
' La ficha viaja en la misma carpeta, con el mismo nombre base que el .NOM.
Public Sub FechaCFDI_GuardarFicha(ByVal rutaNomina As String, ByVal fechaDMY As String)
    Dim n As Integer
    Dim f As Date
    Dim rutaFicha As String

    If Len(Dir$(rutaNomina)) = 0 Then
        Err.Raise vbObjectError + 5204, "FechaCFDI_GuardarFicha", "No existe el archivo de nomina: " & rutaNomina
    End If

    f = FechaCFDI_InterpretarDMY(fechaDMY)
    rutaFicha = FechaCFDI_RutaFicha(rutaNomina)

    n = FreeFile
    On Error GoTo ErrorFicha
    Open rutaFicha For Output As #n
    Print #n, "SACMAG-FECHA-2026"
    Print #n, FechaCFDI_FormatoISO(f)
    Print #n, UCase$(FechaCFDI_NombreArchivo(rutaNomina))
    Close #n
    Exit Sub

ErrorFicha:
    Dim numero As Long, descripcion As String
    numero = Err.Number
    descripcion = Err.Description
    On Error Resume Next
    Close #n
    On Error GoTo 0
    Err.Raise numero, "FechaCFDI_GuardarFicha", descripcion
End Sub

' La fecha NO puede deducirse fiablemente del nombre de FINIQUITO/LIQUIDACION.
' Por eso, si falta la ficha, se reporta claramente y no se inventa una fecha.
Public Function FechaCFDI_LeerFicha(ByVal rutaNomina As String) As String
    Dim n As Integer
    Dim rutaFicha As String
    Dim cabecera As String, iso As String, nombreGuardado As String
    Dim f As Date

    rutaFicha = FechaCFDI_RutaFicha(rutaNomina)
    If Len(Dir$(rutaFicha)) = 0 Then
        Err.Raise vbObjectError + 5205, "FechaCFDI_LeerFicha", "Falta la fecha original de la nomina: " & vbCrLf & rutaFicha & vbCrLf & "Es un archivo anterior a esta correccion. " & "Debe registrar una sola vez su fecha REAL de captura."
    End If

    n = FreeFile
    On Error GoTo ErrorFicha
    Open rutaFicha For Input As #n
    Line Input #n, cabecera
    Line Input #n, iso
    Line Input #n, nombreGuardado
    Close #n
    If Trim$(cabecera) <> "SACMAG-FECHA-2026" Then
        Err.Raise vbObjectError + 5206, "FechaCFDI_LeerFicha", "La ficha de fecha no tiene el formato esperado."
    End If
    If UCase$(Trim$(nombreGuardado)) <> _
       UCase$(FechaCFDI_NombreArchivo(rutaNomina)) Then
        Err.Raise vbObjectError + 5207, "FechaCFDI_LeerFicha", "La ficha de fecha corresponde a otro archivo .NOM."
    End If

    f = FechaCFDI_InterpretarISO(iso)
    FechaCFDI_LeerFicha = FechaCFDI_FormatoDMY(f)
    Exit Function

ErrorFicha:
    Dim numero As Long, descripcion As String
    numero = Err.Number
    descripcion = Err.Description
    On Error Resume Next
    Close #n
    On Error GoTo 0
    Err.Raise numero, "FechaCFDI_LeerFicha", descripcion
End Function


