Attribute VB_Name = "Module5"
Option Explicit

' True solo si empresa.dno se leyo bien en la carpeta activa.
Public g_EmpresaCargada As Boolean

Public Function NormalizarNombreEmpresa(ByVal s As String) As String
    s = Replace(s, Chr$(0), "")
    s = UCase$(Trim$(s))

    s = Replace(s, Chr$(193), "A")
    s = Replace(s, Chr$(201), "E")
    s = Replace(s, Chr$(205), "I")
    s = Replace(s, Chr$(211), "O")
    s = Replace(s, Chr$(218), "U")
    s = Replace(s, Chr$(220), "U")

    s = Replace(s, ".", "")
    s = Replace(s, ",", "")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizarNombreEmpresa = Trim$(s)
End Function

' Lista de empresas AUTORIZADAS (igualdad exacta, ya normalizada:
' mayusculas, sin acentos, sin puntos ni comas).
Private Function EsEmpresaAutorizada(ByVal nombre As String) As Boolean
    Select Case NormalizarNombreEmpresa(nombre)
        Case "SACMAG DE MEXICO SA DE CV"
            EsEmpresaAutorizada = True
    End Select
End Function

' True = puede sincronizar con la BD remota.
Public Function SincronizacionRemotaPermitida() As Boolean
    If Not g_EmpresaCargada Then Exit Function                  ' sin datos de empresa: NO
    If Not EsEmpresaAutorizada(empresa.name) Then Exit Function ' no esta en la lista: NO
    If EsEmpresaLocal(empresa.name) Then Exit Function          ' bandera / reglas legacy: NO
    SincronizacionRemotaPermitida = True
End Function
