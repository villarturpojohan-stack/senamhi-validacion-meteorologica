Attribute VB_Name = "SENAMHI_Herramientas"
Option Explicit

'==========================================================================
'  SENAMHI - Limpieza, validacion y consolidado de datos de estaciones
'==========================================================================
'  COMO USARLO:
'  1. Pega la tabla cruda de la estacion (con o sin la metadata de arriba:
'     Fuente, Estacion, Departamento, etc.) en cualquier hoja, empezando
'     en la celda A1. No importa cuantas filas de metadata tenga.
'  2. Parate en esa hoja (que sea la hoja activa) y ejecuta la macro
'     "SENAMHI_LimpiarYConsolidar" (Alt+F8, o un boton asignado).
'  3. La macro crea (o reutiliza) una hoja llamada "Maestro" y ahi va
'     acumulando todo, evitando duplicados.
'  4. Para la siguiente estacion o mes: pega la nueva tabla cruda (puede
'     ser en la misma hoja, sobrescribiendo, o en una hoja nueva) y vuelve
'     a correr la macro. Todo se va acumulando en "Maestro".
'
'  ESTRUCTURA DE "Maestro": las 4 columnas de datos (Temp Max, Temp Min,
'  Humedad, Precipitacion) van juntas y en el mismo orden que la tabla
'  del SENAMHI, para que sea facil comparar visualmente. Las columnas de
'  "Observacion" van agrupadas despues, una por cada dato, en el mismo
'  orden (Obs_Temp_Max, Obs_Temp_Min, Obs_Humedad, Obs_Precipitacion).
'  Si el valor original era normal, la columna de dato se copia tal cual
'  y su observacion queda vacia. Si era S/D, T, -999, estaba fuera de
'  rango o era un texto raro, la columna de dato se deja VACIA (para no
'  ensuciar tus sumas/promedios) y la observacion dice que paso, por
'  ejemplo: "S/D", "T (traza < 0.1 mm/dia)", "Valor invalido: -999",
'  "Fuera de rango (0-100): -80".
'
'  SUPUESTO: las columnas de datos siempre vienen en el orden en que las
'  entrega el SENAMHI: A=Fecha, B=Temp Max, C=Temp Min, D=Humedad Relativa,
'  E=Precipitacion Total. Si alguna estacion viene en otro orden, avisa
'  para ajustar el codigo.
'==========================================================================

Sub SENAMHI_LimpiarYConsolidar()

    Dim wsOrigen As Worksheet
    Dim wsMaestro As Worksheet
    Dim ultimaFilaOrigen As Long
    Dim ultimaFilaMaestro As Long
    Dim filaInicioDatos As Long
    Dim i As Long
    Dim estacion As String, codigo As String
    Dim celda As Range

    Dim fecha As Date
    Dim vTMax As Variant, vTMin As Variant, vHR As Variant, vPrecip As Variant
    Dim oTMax As String, oTMin As String, oHR As String, oPrecip As String

    Dim contadorAgregados As Long, contadorDuplicados As Long
    Dim contadorSD As Long, contadorTraza As Long, contadorMenos999 As Long
    Dim contadorPrecipNeg As Long, contadorHumedadFueraRango As Long, contadorValorRaro As Long
    Dim fechasFaltantes As String
    Dim fechaAnterior As Date
    Dim primeraFecha As Boolean
    Dim f As Date

    Set wsOrigen = ActiveSheet

    If wsOrigen.Name = "Maestro" Then
        MsgBox "Estas parado en la hoja Maestro. Ponte en la hoja donde " & _
               "pegaste los datos crudos de la estacion y vuelve a ejecutar.", vbExclamation
        Exit Sub
    End If

    ' -----------------------------------------------------------------
    ' 1) Determinar la estacion activa (Estacion + Codigo)
    ' -----------------------------------------------------------------
    ' Orden de prioridad:
    '   a) Si los datos pegados traen las etiquetas "Estacion"/"Codigo"
    '      (la metadata que manda el SENAMHI), se usan esas.
    '   b) Si no, se revisa si ya hay una "estacion activa" guardada de
    '      una corrida anterior (hoja Config) y se confirma con un
    '      Si/No, en vez de tener que volver a escribir el nombre.
    '   c) Si no hay nada guardado todavia, se pide una sola vez y se
    '      guarda para las proximas corridas.
    ' Para cambiar de estacion activa, usa la macro SENAMHI_CambiarEstacionActiva.
    estacion = "": codigo = ""

    Set celda = wsOrigen.Range("A1:H20").Find("Estaci", LookIn:=xlValues, LookAt:=xlPart)
    If Not celda Is Nothing Then
        estacion = Trim(CStr(wsOrigen.Cells(celda.Row, celda.Column + 1).Value))
        If estacion = "" Then
            estacion = Trim(CStr(wsOrigen.Cells(celda.Row, celda.Column + 2).Value))
        End If
    End If

    Set celda = wsOrigen.Range("A1:H20").Find("digo", LookIn:=xlValues, LookAt:=xlPart)
    If Not celda Is Nothing Then
        codigo = Trim(CStr(wsOrigen.Cells(celda.Row, celda.Column + 1).Value))
    End If

    Dim wsConfig As Worksheet
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Config")
    On Error GoTo 0
    If wsConfig Is Nothing Then
        Set wsConfig = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsConfig.Name = "Config"
        wsConfig.Range("A1").Value = "Estacion activa:"
        wsConfig.Range("A2").Value = "Codigo activo:"
    End If

    If estacion <> "" Then
        ' se detecto automaticamente en la metadata: se guarda como activa
        wsConfig.Range("B1").Value = estacion
        wsConfig.Range("B2").Value = codigo
    Else
        Dim estacionGuardada As String, codigoGuardado As String
        estacionGuardada = Trim(CStr(wsConfig.Range("B1").Value))
        codigoGuardado = Trim(CStr(wsConfig.Range("B2").Value))

        If estacionGuardada <> "" Then
            Dim respuesta As VbMsgBoxResult
            respuesta = MsgBox("No se detecto el nombre de la estacion en los datos pegados." & vbCrLf & vbCrLf & _
                                "Sigues trabajando con la estacion activa?" & vbCrLf & _
                                estacionGuardada & " (" & codigoGuardado & ")" & vbCrLf & vbCrLf & _
                                "Si = usar esa misma estacion" & vbCrLf & _
                                "No = escribir el nombre de otra estacion", _
                                vbYesNo + vbQuestion, "Estacion activa")
            If respuesta = vbYes Then
                estacion = estacionGuardada
                codigo = codigoGuardado
            End If
        End If

        If estacion = "" Then
            estacion = Trim(InputBox("Escribe el nombre de la estacion tal cual (ej: SOCSI CAÑETE):", "Nombre de la estacion"))
            If estacion = "" Then estacion = "SIN_NOMBRE_" & Format(Now, "hhmmss")
            codigo = Trim(InputBox("Escribe el codigo de la estacion (o deja en blanco):", "Codigo de la estacion"))
            If codigo = "" Then codigo = "SIN_CODIGO"
            wsConfig.Range("B1").Value = estacion
            wsConfig.Range("B2").Value = codigo
        End If
    End If

    ' -----------------------------------------------------------------
    ' 2) Encontrar donde empiezan los datos: primera fila donde la
    '    columna A sea una fecha valida
    ' -----------------------------------------------------------------
    ultimaFilaOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row
    filaInicioDatos = 0

    For i = 1 To ultimaFilaOrigen
        If IsDate(wsOrigen.Cells(i, 1).Value) Then
            filaInicioDatos = i
            Exit For
        End If
    Next i

    If filaInicioDatos = 0 Then
        MsgBox "No encontre ninguna fila con fecha valida en la columna A." & vbCrLf & _
               "Revisa que los datos esten pegados correctamente empezando en A1.", vbExclamation
        Exit Sub
    End If

    ' -----------------------------------------------------------------
    ' 3) Preparar (o crear) la hoja Maestro
    ' -----------------------------------------------------------------
    On Error Resume Next
    Set wsMaestro = ThisWorkbook.Sheets("Maestro")
    On Error GoTo 0

    If wsMaestro Is Nothing Then
        Set wsMaestro = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsMaestro.Name = "Maestro"
        With wsMaestro
            .Range("A1:O1").Value = Array("Estacion", "Codigo", "Fecha", "Anio", "Mes", "Dia", _
                                           "Temp_Max", "Temp_Min", "Humedad_Relativa", "Precipitacion_Total", _
                                           "", _
                                           "Obs_Temp_Max", "Obs_Temp_Min", "Obs_Humedad", "Obs_Precipitacion")
            .Range("A1:N1").Font.Bold = True
            .Columns("C").NumberFormat = "dd/mm/yyyy"
        End With
    End If

    ultimaFilaMaestro = wsMaestro.Cells(wsMaestro.Rows.Count, "A").End(xlUp).Row

    ' -----------------------------------------------------------------
    ' 4) Recorrer las filas de datos: validar y agregar a Maestro
    ' -----------------------------------------------------------------
    contadorAgregados = 0: contadorDuplicados = 0
    contadorSD = 0: contadorTraza = 0: contadorMenos999 = 0
    contadorPrecipNeg = 0: contadorHumedadFueraRango = 0: contadorValorRaro = 0
    fechasFaltantes = ""
    primeraFecha = True

    For i = filaInicioDatos To ultimaFilaOrigen

        If Not IsDate(wsOrigen.Cells(i, 1).Value) Then Exit For ' fin de los datos

        fecha = CDate(wsOrigen.Cells(i, 1).Value)

        ' Si la fecha retrocede respecto a la anterior, probablemente son datos
        ' viejos que quedaron pegados debajo (por ejemplo, pegaste un mes mas
        ' corto encima de uno mas largo y no se sobrescribieron todas las filas).
        ' Se corta la lectura aqui en vez de mezclar dos periodos distintos.
        If Not primeraFecha Then
            If fecha < fechaAnterior Then Exit For
        End If

        ' --- chequeo de dias faltantes (serie diaria consecutiva) ---
        If primeraFecha Then
            primeraFecha = False
        Else
            If fecha <> fechaAnterior + 1 Then
                For f = fechaAnterior + 1 To fecha - 1
                    fechasFaltantes = fechasFaltantes & Format(f, "dd/mm/yyyy") & ", "
                Next f
            End If
        End If
        fechaAnterior = fecha

        ' --- leer y validar cada columna (valor limpio + observacion por separado) ---
        ValidarValor wsOrigen.Cells(i, 2).Value, False, False, vTMax, oTMax, _
                     contadorSD, contadorTraza, contadorMenos999, contadorPrecipNeg, contadorHumedadFueraRango, contadorValorRaro
        ValidarValor wsOrigen.Cells(i, 3).Value, False, False, vTMin, oTMin, _
                     contadorSD, contadorTraza, contadorMenos999, contadorPrecipNeg, contadorHumedadFueraRango, contadorValorRaro
        ValidarValor wsOrigen.Cells(i, 4).Value, False, True, vHR, oHR, _
                     contadorSD, contadorTraza, contadorMenos999, contadorPrecipNeg, contadorHumedadFueraRango, contadorValorRaro
        ValidarValor wsOrigen.Cells(i, 5).Value, True, False, vPrecip, oPrecip, _
                     contadorSD, contadorTraza, contadorMenos999, contadorPrecipNeg, contadorHumedadFueraRango, contadorValorRaro

        ' --- evitar duplicados (misma estacion + misma fecha ya en Maestro) ---
        If ExisteEnMaestro(wsMaestro, estacion, fecha) Then
            contadorDuplicados = contadorDuplicados + 1
        Else
            ultimaFilaMaestro = ultimaFilaMaestro + 1
            With wsMaestro
                .Cells(ultimaFilaMaestro, 1).Value = estacion
                .Cells(ultimaFilaMaestro, 2).Value = codigo
                .Cells(ultimaFilaMaestro, 3).Value = fecha
                .Cells(ultimaFilaMaestro, 3).NumberFormat = "dd/mm/yyyy"
                .Cells(ultimaFilaMaestro, 4).Value = Year(fecha)
                .Cells(ultimaFilaMaestro, 5).Value = Month(fecha)
                .Cells(ultimaFilaMaestro, 6).Value = Day(fecha)
                .Cells(ultimaFilaMaestro, 7).Value = vTMax
                .Cells(ultimaFilaMaestro, 8).Value = vTMin
                .Cells(ultimaFilaMaestro, 9).Value = vHR
                .Cells(ultimaFilaMaestro, 10).Value = vPrecip
                .Cells(ultimaFilaMaestro, 12).Value = oTMax
                .Cells(ultimaFilaMaestro, 13).Value = oTMin
                .Cells(ultimaFilaMaestro, 14).Value = oHR
                .Cells(ultimaFilaMaestro, 15).Value = oPrecip

                ' Color: rojo si hay algun error real, celeste si solo hay trazas (T),
                ' sin color si todo esta normal.
                Dim hayError As Boolean, hayTraza As Boolean
                hayError = (oTMax Like "S/D*" Or oTMax Like "Valor*" Or oTMax Like "Fuera*" Or oTMax Like "Precip*") Or _
                           (oTMin Like "S/D*" Or oTMin Like "Valor*" Or oTMin Like "Fuera*" Or oTMin Like "Precip*") Or _
                           (oHR Like "S/D*" Or oHR Like "Valor*" Or oHR Like "Fuera*" Or oHR Like "Precip*") Or _
                           (oPrecip Like "S/D*" Or oPrecip Like "Valor*" Or oPrecip Like "Fuera*" Or oPrecip Like "Precip*")
                hayTraza = (oTMax Like "T (*") Or (oTMin Like "T (*") Or (oHR Like "T (*") Or (oPrecip Like "T (*")

                If hayError Then
                    .Range(.Cells(ultimaFilaMaestro, 1), .Cells(ultimaFilaMaestro, 15)).Interior.Color = RGB(255, 199, 206)
                ElseIf hayTraza Then
                    .Range(.Cells(ultimaFilaMaestro, 1), .Cells(ultimaFilaMaestro, 15)).Interior.Color = RGB(198, 224, 255)
                End If
            End With
            contadorAgregados = contadorAgregados + 1
        End If

    Next i

    If Len(fechasFaltantes) > 2 Then fechasFaltantes = Left(fechasFaltantes, Len(fechasFaltantes) - 2)

    wsMaestro.Columns("A:O").AutoFit

    ' -----------------------------------------------------------------
    ' 5) Resumen final
    ' -----------------------------------------------------------------
    Dim resumen As String
    resumen = "Estacion: " & estacion & " (" & codigo & ")" & vbCrLf & vbCrLf & _
              "Filas agregadas a Maestro: " & contadorAgregados & vbCrLf & _
              "Filas duplicadas (ya existian, se omitieron): " & contadorDuplicados & vbCrLf & vbCrLf & _
              "Valores S/D: " & contadorSD & vbCrLf & _
              "Trazas (T): " & contadorTraza & vbCrLf & _
              "Valores invalidos tipo -999: " & contadorMenos999 & vbCrLf & _
              "Precipitacion negativa: " & contadorPrecipNeg & vbCrLf & _
              "Humedad fuera de rango (0-100): " & contadorHumedadFueraRango & vbCrLf & _
              "Valores no numericos raros: " & contadorValorRaro & vbCrLf & vbCrLf

    If fechasFaltantes <> "" Then
        resumen = resumen & "ATENCION - Dias faltantes en este rango:" & vbCrLf & fechasFaltantes
    Else
        resumen = resumen & "Sin dias faltantes en este rango."
    End If

    MsgBox resumen, vbInformation, "Resultado de validacion SENAMHI"

End Sub


'==========================================================================
' Valida un valor crudo de una celda (temperatura, humedad o precipitacion).
' Si el valor es normal, valorLimpio = numero y observacion = "".
' Si el valor tiene algun problema (S/D, T, -999, fuera de rango, texto raro),
' valorLimpio queda VACIO (Empty) y observacion explica que paso con el
' valor original, para que la columna numerica no se ensucie.
'==========================================================================
Private Sub ValidarValor(valorCrudo As Variant, esPrecipitacion As Boolean, esHumedad As Boolean, _
                          ByRef valorLimpio As Variant, ByRef observacion As String, _
                          ByRef contadorSD As Long, ByRef contadorTraza As Long, ByRef contadorMenos999 As Long, _
                          ByRef contadorPrecipNeg As Long, ByRef contadorHumedadFueraRango As Long, _
                          ByRef contadorValorRaro As Long)
    Dim txt As String
    txt = Trim(CStr(valorCrudo))
    observacion = ""
    valorLimpio = Empty

    If UCase(txt) = "S/D" Or txt = "" Then
        contadorSD = contadorSD + 1
        observacion = "S/D"
        Exit Sub
    End If

    If UCase(txt) = "T" Then
        contadorTraza = contadorTraza + 1
        observacion = "T (traza < 0.1 mm/dia)"
        Exit Sub
    End If

    If IsNumeric(txt) Then
        If CDbl(txt) <= -900 Then
            contadorMenos999 = contadorMenos999 + 1
            observacion = "Valor invalido: " & txt
        ElseIf esPrecipitacion And CDbl(txt) < 0 Then
            contadorPrecipNeg = contadorPrecipNeg + 1
            observacion = "Precipitacion negativa: " & txt
        ElseIf esHumedad And (CDbl(txt) < 0 Or CDbl(txt) > 100) Then
            contadorHumedadFueraRango = contadorHumedadFueraRango + 1
            observacion = "Fuera de rango (0-100): " & txt
        Else
            valorLimpio = CDbl(txt)
        End If
    Else
        contadorValorRaro = contadorValorRaro + 1
        observacion = "Valor no numerico: " & txt
    End If
End Sub


'==========================================================================
' Revisa si ya existe una fila en Maestro con la misma estacion y fecha.
'==========================================================================
Private Function ExisteEnMaestro(ws As Worksheet, estacion As String, fecha As Date) As Boolean
    Dim ultimaFila As Long, k As Long
    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ExisteEnMaestro = False
    For k = 2 To ultimaFila
        If Trim(UCase(CStr(ws.Cells(k, 1).Value))) = Trim(UCase(estacion)) And ws.Cells(k, 3).Value = fecha Then
            ExisteEnMaestro = True
            Exit Function
        End If
    Next k
End Function


'==========================================================================
' Utilidad: borra la "estacion activa" guardada. Usala cuando termines
' con una estacion y vayas a empezar con otra (y tus datos no traen
' la metadata con el nombre). Asi la proxima corrida te pregunta de nuevo
' en vez de seguir usando la anterior.
'==========================================================================
Sub SENAMHI_CambiarEstacionActiva()
    Dim wsConfig As Worksheet
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Config")
    On Error GoTo 0
    If wsConfig Is Nothing Then
        MsgBox "Todavia no hay ninguna estacion activa guardada.", vbInformation
        Exit Sub
    End If
    wsConfig.Range("B1:B2").ClearContents
    MsgBox "Listo. La proxima vez que corras la macro (con datos sin metadata) " & _
           "te va a preguntar el nombre de la estacion de nuevo.", vbInformation
End Sub
