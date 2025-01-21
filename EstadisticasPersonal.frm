VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form EstadisticasPersonal 
   Caption         =   "Estadísticas"
   ClientHeight    =   9120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13725
   LinkTopic       =   "Form10"
   ScaleHeight     =   9120
   ScaleWidth      =   13725
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   8895
      Left            =   120
      OleObjectBlob   =   "EstadisticasPersonal.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   13455
   End
End
Attribute VB_Name = "EstadisticasPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ' Definir la estructura de datos
    Dim registro As OtrasCh
    Dim numPersonas As Integer
    Dim i As Integer
    Dim countHombres As Integer
    Dim countMujeres As Integer

    Open "PerOtre.dno" For Random As #3 Len = Len(registro)
    numPersonas = LOF(3) / Len(registro)
    
    With MSChart1
        .ChartType = VtChChartType2dBar
        .RowCount = 1
        .ColumnCount = 2
        .RowLabel = "Género"
        .ColumnLabel = "Cantidad"
        .Title.Text = ("Cantidad de Hombres y Mujeres")
        
        For i = 1 To numPersonas
            Get #3, i, registro
            Dim genero As String
            genero = Mid(registro.curp, 11, 1)
            
            If genero = "H" Then
                countHombres = countHombres + 1
            ElseIf genero = "M" Then
                countMujeres = countMujeres + 1
            End If
        Next i

        .DataGrid.SetData 1, 1, countHombres, 0
        .ChartType = VtChChartType2dBar
        .DataGrid.SetData 1, 2, countMujeres, 0
        
        .RowLabel = "Hombres: " & CStr(countHombres) & "                                                                       " & "Mujeres: " & CStr(countMujeres)  ' Etiqueta para la barra de mujeres
    End With
    Close #3
End Sub

