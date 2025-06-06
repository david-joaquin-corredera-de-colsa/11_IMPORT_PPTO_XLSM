Attribute VB_Name = "Modulo_012_FUNC_Principales_01"
Option Explicit

Public Function F000_Comprobaciones_Iniciales() As Boolean
    

    '******************************************************************************
    ' M�dulo: F000_Comprobaciones_Iniciales
    ' Fecha y Hora de Creaci�n: 2025-05-26 09:32:08 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Esta funci�n realiza las comprobaciones iniciales necesarias y crea las hojas
    ' requeridas para el proceso de importaci�n.
    '
    ' Pasos:
    ' 1. Inicializaci�n de variables globales
    ' 2. Validaci�n y creaci�n de hojas base (Procesos, Inventario, Log)
    ' 3. Generaci�n de nombres para nuevas hojas de importaci�n
    ' 4. Creaci�n de hojas de importaci�n
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para fechas y nombres de hojas
    Dim strFechaHoraIsoActual As String
    Dim strFechaHoraIsoNuevaHojaImportacion As String
    Dim strPrefijoHojaImportacion As String
    Dim strPrefijoHojaImportacion_Working As String
    Dim strPrefijoHojaImportacion_Envio As String
    
    ' Inicializaci�n
    strFuncion = "F000_Comprobaciones_Iniciales"
    F000_Comprobaciones_Iniciales = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicializar variables globales
    '--------------------------------------------------------------------------
    lngLineaError = 50
    Call InitializeGlobalVariables
    
    '--------------------------------------------------------------------------
    ' 2. Validar/Crear hojas base
    '--------------------------------------------------------------------------
    lngLineaError = 57
    ' Validar/Crear hoja Ejecutar Procesos
    If Not fun802_SheetExists(gstrHoja_EjecutarProcesos) Then
        If Not F002_Crear_Hoja(gstrHoja_EjecutarProcesos) Then
            Err.Raise ERROR_BASE_IMPORT + 1, strFuncion, _
                "Error al crear la hoja " & gstrHoja_EjecutarProcesos
        End If
    End If
    
    ' Validar/Crear hoja Inventario
    If Not fun802_SheetExists(gstrHoja_Inventario) Then
        If Not F002_Crear_Hoja(gstrHoja_Inventario) Then
            Err.Raise ERROR_BASE_IMPORT + 2, strFuncion, _
                "Error al crear la hoja " & gstrHoja_Inventario
        End If
    End If
    ThisWorkbook.Worksheets(gstrHoja_Inventario).Visible = xlSheetHidden
    
    
    ' Validar/Crear hoja Log
    If Not fun802_SheetExists(gstrHoja_Log) Then
        If Not F002_Crear_Hoja(gstrHoja_Log) Then
            Err.Raise ERROR_BASE_IMPORT + 3, strFuncion, _
                "Error al crear la hoja " & gstrHoja_Log
        End If
    End If
    
    ' Validar/Crear hoja Delimitadores Originales
    If Not fun802_SheetExists(gstrHoja_DelimitadoresOriginales) Then
        If Not F002_Crear_Hoja(gstrHoja_DelimitadoresOriginales) Then
            Err.Raise ERROR_BASE_IMPORT + 3, strFuncion, _
                "Error al crear la hoja " & gstrHoja_DelimitadoresOriginales
        End If
    End If
    
    ' Validar/Crear hoja UserName
    If Not fun802_SheetExists(gstrHoja_UserName) Then
        If Not F002_Crear_Hoja(gstrHoja_UserName) Then
            Err.Raise ERROR_BASE_IMPORT + 3, strFuncion, _
                "Error al crear la hoja " & gstrHoja_UserName
        End If
    End If
    ' Verificar si debemos ocultar la hoja UserName (comprobando la constante global CONST_OCULTAR_HOJA_USERNAME)
    If CONST_OCULTAR_HOJA_USERNAME = True Then
        ' Ocultar la hoja de delimitadores
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(gstrHoja_UserName)
        If Not fun809_OcultarHojaDelimitadores(ws) Then
            Debug.Print "ADVERTENCIA: Error al ocultar la hoja " & gstrHoja_UserName & " - Funci�n: F000_Comprobaciones_Iniciales - " & Now()
            ' Nota: No es un error cr�tico, el proceso puede continuar
        End If
    End If
    
    
    ' Proceso completado exitosamente
    F000_Comprobaciones_Iniciales = True
    fun801_LogMessage "Comprobaciones iniciales completadas con �xito"
    Exit Function

GestorErrores:
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    fun801_LogMessage strMensajeError, True
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F000_Comprobaciones_Iniciales = False
End Function


Public Function F001_Crear_hojas_de_Importacion() As Boolean
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para fechas y nombres de hojas
    Dim strFechaHoraIsoActual As String
    Dim strFechaHoraIsoNuevaHojaImportacion As String
    Dim strPrefijoHojaImportacion As String
    Dim strPrefijoHojaImportacion_Working As String
    Dim strPrefijoHojaImportacion_Envio As String
    Dim strPrefijoHojaImportacion_Comprobacion As String
    
    ' Inicializaci�n
    strFuncion = "F001_Crear_hojas_de_Importacion"
    F001_Crear_hojas_de_Importacion = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicializar variables globales
    '--------------------------------------------------------------------------
    lngLineaError = 51
    Call InitializeGlobalVariables
   
    '--------------------------------------------------------------------------
    ' 3. Generar nombres para nuevas hojas
    '--------------------------------------------------------------------------
    lngLineaError = 85
    ' Generar timestamp ISO
    strFechaHoraIsoActual = Format(Now(), "yyyymmdd_hhmmss")
    strFechaHoraIsoNuevaHojaImportacion = strFechaHoraIsoActual
    
    ' Definir prefijos
    strPrefijoHojaImportacion = "Import_"
    strPrefijoHojaImportacion_Working = "Import_Working_"
    strPrefijoHojaImportacion_Envio = "Import_Envio_"
    strPrefijoHojaImportacion_Comprobacion = "Import_Comprob_"
    
    ' Generar nombres completos (variables globales)
    gstrNuevaHojaImportacion = strPrefijoHojaImportacion & strFechaHoraIsoNuevaHojaImportacion
    gstrNuevaHojaImportacion_Working = strPrefijoHojaImportacion_Working & strFechaHoraIsoNuevaHojaImportacion
    gstrNuevaHojaImportacion_Envio = strPrefijoHojaImportacion_Envio & strFechaHoraIsoNuevaHojaImportacion
    gstrNuevaHojaImportacion_Comprobacion = strPrefijoHojaImportacion_Comprobacion & strFechaHoraIsoNuevaHojaImportacion
    
    '--------------------------------------------------------------------------
    ' 4. Crear hojas de importaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 102
    ' Crear hoja de importaci�n
    If Not F002_Crear_Hoja(gstrNuevaHojaImportacion) Then
        Err.Raise ERROR_BASE_IMPORT + 4, strFuncion, _
            "Error al crear la hoja " & gstrNuevaHojaImportacion
    End If
    
    ' Crear hoja de trabajo
    If Not F002_Crear_Hoja(gstrNuevaHojaImportacion_Working) Then
        Err.Raise ERROR_BASE_IMPORT + 5, strFuncion, _
            "Error al crear la hoja " & gstrNuevaHojaImportacion_Working
    End If
    
    ' Crear hoja de env�o
    If Not F002_Crear_Hoja(gstrNuevaHojaImportacion_Envio) Then
        Err.Raise ERROR_BASE_IMPORT + 6, strFuncion, _
            "Error al crear la hoja " & gstrNuevaHojaImportacion_Envio
    End If
    
    ' Crear hoja de comprobaci�n
    If Not F002_Crear_Hoja(gstrNuevaHojaImportacion_Comprobacion) Then
        Err.Raise ERROR_BASE_IMPORT + 7, strFuncion, _
            "Error al crear la hoja " & gstrNuevaHojaImportacion_Comprobacion
    End If
    
    ' Proceso completado exitosamente
    F001_Crear_hojas_de_Importacion = True
    fun801_LogMessage "Creacion de hojas de importacion completada con �xito (4 hojas creadas)"
    Exit Function

GestorErrores:
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    fun801_LogMessage strMensajeError, True
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F001_Crear_hojas_de_Importacion = False
End Function


Public Function F002_Importar_Fichero(ByVal vNuevaHojaImportacion As String, _
                                    ByVal vNuevaHojaImportacion_Working As String, _
                                    ByVal vNuevaHojaImportacion_Envio As String, _
                                    ByVal vNuevaHojaImportacion_Comprobacion As String) As Boolean
    
    '******************************************************************************
    ' M�dulo: F002_Importar_Fichero
    ' Fecha y Hora de Creaci�n: 2025-05-29 03:42:14 UTC
    ' Modificado: 2025-05-30 05:33:13 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Funci�n para importar ficheros de texto a Excel, manteniendo el formato original
    ' en la hoja de importaci�n y procesando los datos en la hoja de trabajo.
    ' Incluye detecci�n avanzada de duplicados basada en concatenaci�n de valores.
    ' MODIFICACI�N: A�adido par�metro vNuevaHojaImportacion_Comprobacion y replicaci�n
    ' de acciones de limpieza para esta hoja adicional.
    '
    ' Pasos:
    ' 1. Limpieza de hojas destino (Importaci�n, Working, Env�o, Comprobaci�n)
    ' 2. Selecci�n de archivo mediante cuadro de di�logo
    ' 3. Importaci�n de datos sin procesar a hoja de importaci�n
    ' 4. Copia de datos a hoja de trabajo
    ' 5. Procesamiento en hoja de trabajo:
    '    - Detecci�n de rango de datos
    '    - Conversi�n de texto a columnas con formatos espec�ficos
    ' 6. Procesamiento adicional de datos:
    '    - Concatenaci�n de valores de columnas con delimitador "|"
    '    - Detecci�n de duplicados basada en la concatenaci�n
    '    - Marcado de l�neas duplicadas
    ' 7. Procesamiento complementario de l�neas duplicadas:
    '    - Identificaci�n de l�neas repetidas no tratadas
    '    - Comparaci�n basada en valores concatenados
    '    - Suma de importes para l�neas duplicadas
    '    - Creaci�n de l�neas resumen con totales consolidados
    ' 8. Ajuste de zoom de la hoja de trabajo al 70%
    '******************************************************************************

    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String

    ' Variables para hojas y rangos
    Dim wsImport As Worksheet
    Dim wsWorking As Worksheet
    Dim wsEnvio As Worksheet
    Dim wsComprobacion As Worksheet
    Dim rngConversion As Range

    ' Variables para importaci�n
    Dim strFilePath As String
    Dim lngCol As Long
    
    ' Variables para bucles
    Dim i As Long                      ' Variable para bucle principal
    Dim j As Long                      ' Variable para bucle anidado
    Dim k As Long                      ' Variable para bucle de procesamiento
    Dim m As Long                      ' Variable para bucle de b�squeda l�neas vac�as

    ' Inicializaci�n
    strFuncion = "F002_Importar_Fichero"
    F002_Importar_Fichero = False
    lngLineaError = 0

    On Error GoTo GestorErrores

    '--------------------------------------------------------------------------
    ' 1. Limpiar hojas destino
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Iniciando proceso de importaci�n", False, "", ""
    
    ' Limpiar hoja de importaci�n
    fun801_LogMessage "Limpiando hoja", False, "", vNuevaHojaImportacion
    If Not fun801_LimpiarHoja(vNuevaHojaImportacion) Then
        fun801_LogMessage "Error al limpiar hoja", True, "", vNuevaHojaImportacion
        Err.Raise ERROR_BASE_IMPORT + 1, strFuncion, _
            "Error al limpiar la hoja " & vNuevaHojaImportacion
    End If
    
    ' Limpiar hoja working
    fun801_LogMessage "Limpiando hoja", False, "", vNuevaHojaImportacion_Working
    If Not fun801_LimpiarHoja(vNuevaHojaImportacion_Working) Then
        fun801_LogMessage "Error al limpiar hoja", True, "", vNuevaHojaImportacion_Working
        Err.Raise ERROR_BASE_IMPORT + 2, strFuncion, _
            "Error al limpiar la hoja " & vNuevaHojaImportacion_Working
    End If
    
    ' Limpiar hoja env�o
    fun801_LogMessage "Limpiando hoja", False, "", vNuevaHojaImportacion_Envio
    If Not fun801_LimpiarHoja(vNuevaHojaImportacion_Envio) Then
        fun801_LogMessage "Error al limpiar hoja", True, "", vNuevaHojaImportacion_Envio
        Err.Raise ERROR_BASE_IMPORT + 3, strFuncion, _
            "Error al limpiar la hoja " & vNuevaHojaImportacion_Envio
    End If
    
    ' Limpiar hoja comprobaci�n
    lngLineaError = 55
    fun801_LogMessage "Limpiando hoja", False, "", vNuevaHojaImportacion_Comprobacion
    If Not fun801_LimpiarHoja(vNuevaHojaImportacion_Comprobacion) Then
        fun801_LogMessage "Error al limpiar hoja", True, "", vNuevaHojaImportacion_Comprobacion
        Err.Raise ERROR_BASE_IMPORT + 4, strFuncion, _
            "Error al limpiar la hoja " & vNuevaHojaImportacion_Comprobacion
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Seleccionar archivo
    '--------------------------------------------------------------------------
    lngLineaError = 71
    fun801_LogMessage "Solicitando selecci�n de archivo al usuario", False, "", ""
    strFilePath = fun802_SeleccionarArchivo("�Qu� fichero desea importar?")
    
    If strFilePath = "" Then
        fun801_LogMessage "No se seleccion� ning�n archivo", True, "", ""
        Err.Raise ERROR_BASE_IMPORT + 5, strFuncion, _
            "No se seleccion� ning�n archivo"
    End If
    
    fun801_LogMessage "Archivo seleccionado para importar", False, strFilePath, vNuevaHojaImportacion
    
    '--------------------------------------------------------------------------
    ' 3. Importar datos sin procesar
    '--------------------------------------------------------------------------
    lngLineaError = 81
    fun801_LogMessage "Iniciando importaci�n de archivo", False, strFilePath, vNuevaHojaImportacion
    Set wsImport = ThisWorkbook.Worksheets(vNuevaHojaImportacion)
    
    If Not fun803_ImportarArchivo(wsImport, strFilePath, _
                               vColumnaInicial_Importacion, _
                               vFilaInicial_Importacion) Then
        fun801_LogMessage "Error en la importaci�n", True, strFilePath, vNuevaHojaImportacion
        Err.Raise ERROR_BASE_IMPORT + 6, strFuncion, _
            "Error al importar el archivo"
    End If
    
    fun801_LogMessage "Archivo importado correctamente", False, strFilePath, vNuevaHojaImportacion
    
    '--------------------------------------------------------------------------
    ' 4. Copiar datos a hoja working
    '--------------------------------------------------------------------------
    lngLineaError = 95
    fun801_LogMessage "Copiando datos a hoja de trabajo", False, strFilePath, vNuevaHojaImportacion_Working
    Set wsWorking = ThisWorkbook.Worksheets(vNuevaHojaImportacion_Working)
    
    ' Copiar datos
    wsImport.UsedRange.Copy wsWorking.Range(vColumnaInicial_Importacion & vFilaInicial_Importacion)
    fun801_LogMessage "Datos copiados correctamente", False, strFilePath, vNuevaHojaImportacion_Working
    
    '--------------------------------------------------------------------------
    ' 5. Procesar datos en hoja working
    '--------------------------------------------------------------------------
    lngLineaError = 104
    ' Detectar rango de datos
    fun801_LogMessage "Detectando rango de datos", False, strFilePath, vNuevaHojaImportacion_Working
    If Not fun804_DetectarRangoDatos(wsWorking, _
                                  vLineaInicial_HojaImportacion, _
                                  vLineaFinal_HojaImportacion) Then
        fun801_LogMessage "Error al detectar rango de datos", True, strFilePath, vNuevaHojaImportacion_Working
        Err.Raise ERROR_BASE_IMPORT + 7, strFuncion, _
            "Error al detectar el rango de datos"
    End If
    
    fun801_LogMessage "Rango detectado: " & vLineaInicial_HojaImportacion & " a " & vLineaFinal_HojaImportacion, _
                      False, strFilePath, vNuevaHojaImportacion_Working
    
    ' Seleccionar rango para conversi�n
    Set rngConversion = wsWorking.Range( _
        vColumnaInicial_Importacion & vLineaInicial_HojaImportacion & ":" & _
        vColumnaInicial_Importacion & vLineaFinal_HojaImportacion)
    
    ' Convertir texto a columnas con formatos espec�ficos
    lngLineaError = 120
    fun801_LogMessage "Iniciando conversi�n texto a columnas", False, strFilePath, vNuevaHojaImportacion_Working
    
    With rngConversion
        .TextToColumns _
            Destination:=.Cells(1), _
            DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, _
            ConsecutiveDelimiter:=False, _
            Tab:=False, _
            Semicolon:=(vDelimitador_Importacion = ";"), _
            Comma:=(vDelimitador_Importacion = ","), _
            Space:=(vDelimitador_Importacion = " "), _
            Other:=True, _
            OtherChar:=IIf(vDelimitador_Importacion <> ";" And _
                          vDelimitador_Importacion <> "," And _
                          vDelimitador_Importacion <> " ", _
                          vDelimitador_Importacion, "")
        
        ' Configurar formato de columnas
        lngCol = Range(vColumnaInicial_Importacion & "1").Column
        
        ' Columnas 1-11 como texto
        wsWorking.Range(wsWorking.Cells(vLineaInicial_HojaImportacion, lngCol), _
                       wsWorking.Cells(vLineaFinal_HojaImportacion, lngCol + 10)).NumberFormat = "@"
        
        ' Columnas 12-23 como General
        wsWorking.Range(wsWorking.Cells(vLineaInicial_HojaImportacion, lngCol + 11), _
                       wsWorking.Cells(vLineaFinal_HojaImportacion, lngCol + 22)).NumberFormat = "General"
    End With
    
    fun801_LogMessage "Conversi�n texto a columnas completada", False, strFilePath, vNuevaHojaImportacion_Working
    

    '--------------------------------------------------------------------------
    ' 6. Procesamiento adicional de datos: Concatenaci�n y detecci�n de duplicados
    '--------------------------------------------------------------------------
    lngLineaError = 150
    fun801_LogMessage "Iniciando procesamiento adicional de datos", False, strFilePath, vNuevaHojaImportacion_Working
    
    ' 6.1. Declarar variables para el procesamiento adicional
    Dim vDelimita As String                 ' Variable para almacenar el delimitador
    Dim vCampos01a11 As String              ' Variable para almacenar concatenaci�n
    Dim vCampos01a11_Verificar As String    ' Variable para verificar duplicados
    'Dim vTagLineaRepetida As String         ' Variable para marcar l�neas duplicadas
    Dim vLineaAncla As Long                 ' Variable para almacenar l�nea de referencia
    Dim vLineaEnCurso As Long               ' Variable para almacenar l�nea en procesamiento
    Dim lngColBase As Long                  ' Columna base de inicio
    
    ' 6.2. Inicializar variables
    vDelimita = "|"                    ' Delimitador pipe para concatenaci�n
    'vTagLineaRepetida = "Linea_Repetida" ' Etiqueta para marcar duplicados
    lngColBase = Range(vColumnaInicial_Importacion & "1").Column ' Obtener n�mero de columna base
    
    ' 6.3. Primer bucle: Concatenar columnas y almacenar en columna+23
    lngLineaError = 160
    fun801_LogMessage "Concatenando valores de columnas", False, strFilePath, vNuevaHojaImportacion_Working
    
    For i = vLineaInicial_HojaImportacion To vLineaFinal_HojaImportacion
        ' Inicializar variable para concatenaci�n
        vCampos01a11 = ""
        
        ' Construir concatenaci�n de valores de columnas desde base hasta base+10
        For j = 0 To 10
            ' Para la primera columna no a�adir el delimitador inicial
            If j = 0 Then
                ' Primera columna sin delimitador previo
                vCampos01a11 = Trim(CStr(wsWorking.Cells(i, lngColBase + j).Value))
            Else
                ' Columnas siguientes con delimitador
                vCampos01a11 = vCampos01a11 & vDelimita & Trim(CStr(wsWorking.Cells(i, lngColBase + j).Value))
            End If
        Next j
        
        ' Convertir toda la concatenaci�n a may�sculas
        vCampos01a11 = UCase(vCampos01a11)
        
        ' Almacenar el resultado en la columna base+23
        wsWorking.Cells(i, lngColBase + 23).Value = vCampos01a11
    Next i
    
    ' 6.4. Segundo bucle: Detectar y marcar duplicados
    lngLineaError = 180
    fun801_LogMessage "Detectando y marcando l�neas duplicadas", False, strFilePath, vNuevaHojaImportacion_Working
    
    For i = vLineaInicial_HojaImportacion To vLineaFinal_HojaImportacion
        ' Almacenar l�nea actual como ancla
        vLineaAncla = i
        
        ' Obtener valor concatenado de la l�nea ancla
        vCampos01a11_Verificar = CStr(wsWorking.Cells(vLineaAncla, lngColBase + 23).Value)
        
        ' Buscar duplicados desde la siguiente l�nea
        For j = vLineaAncla + 1 To vLineaFinal_HojaImportacion
            ' Almacenar l�nea actual del bucle anidado
            vLineaEnCurso = j
            
            ' Comprobar si el valor concatenado coincide con el valor de referencia
            If CStr(wsWorking.Cells(vLineaEnCurso, lngColBase + 23).Value) = vCampos01a11_Verificar Then
                ' Si coincide, marcar ambas l�neas como duplicadas
                'wsWorking.Cells(vLineaAncla, lngColBase + 24).Value = vTagLineaRepetida
                wsWorking.Cells(vLineaAncla, lngColBase + 24).Value = CONST_TAG_LINEA_REPETIDA
                'wsWorking.Cells(vLineaEnCurso, lngColBase + 24).Value = vTagLineaRepetida
                wsWorking.Cells(vLineaEnCurso, lngColBase + 24).Value = CONST_TAG_LINEA_REPETIDA
            End If
            
            ' Si estamos en la �ltima l�nea, limpiar la variable de verificaci�n
            If vLineaEnCurso = vLineaFinal_HojaImportacion Then
                vCampos01a11_Verificar = ""
            End If
        Next j
    Next i
    
    fun801_LogMessage "Procesamiento de duplicados completado", False, strFilePath, vNuevaHojaImportacion_Working
    
    
    '--------------------------------------------------------------------------
    ' 7. Procesamiento complementario de l�neas duplicadas
    '--------------------------------------------------------------------------
    lngLineaError = 200
    fun801_LogMessage "Iniciando procesamiento complementario de l�neas duplicadas", False, strFilePath, vNuevaHojaImportacion_Working
    
    ' 7.1. Declaraci�n de variables para el procesamiento de l�neas duplicadas
    ' 7.1.1. Variables de tipo string para etiquetas
    'Dim vTagLineaTratada As String     ' Etiqueta para l�neas ya procesadas
    'Dim vTagLineaSuma As String        ' Etiqueta para l�neas de suma
    
    ' 7.1.2. Variables de tipo string para almacenar valores de columnas
    Dim vValorColumna_Campos01a11_LineaAncla As String   ' Valor concatenado l�nea ancla
    Dim vValorColumna_Campos01a11_LineaEnCurso As String ' Valor concatenado l�nea en curso
    Dim vValorColumna_TagLineaRepetida As String         ' Valor de tag l�nea repetida
    Dim vValorColumna_TagLineaTratada As String          ' Valor de tag l�nea tratada
    Dim vValorColumna_TagLineaSuma As String            ' Valor de tag l�nea suma
    
    ' 7.1.3. Variables para bucles
    ' Se reutilizan i, j, k, m que ya est�n declaradas al inicio de la funci�n
    
    ' 7.1.7. y 7.1.8. Variables para l�neas de referencia
    ' Se reutilizan vLineaAncla y vLineaEnCurso que ya est�n declaradas
    
    ' Inicializar variables
    'vTagLineaTratada = "Linea_Tratada"
    'vTagLineaSuma = "Linea_Suma"
    
    ' Vaciar los valores de las variables de string
    vValorColumna_Campos01a11_LineaAncla = ""
    vValorColumna_Campos01a11_LineaEnCurso = ""
    vValorColumna_TagLineaRepetida = ""
    vValorColumna_TagLineaTratada = ""
    vValorColumna_TagLineaSuma = ""
    
    ' Vaciar los valores de las variables de bucles
    i = 0
    j = 0
    k = 0
    m = 0
    
    ' Vaciar los valores de las l�neas de referencia
    vLineaAncla = 0
    vLineaEnCurso = 0
    
    ' Declarar arrays para almacenar importes
    Dim vArrayImportes_LineaAncla() As Double    ' Para almacenar importes l�nea ancla
    Dim vArrayImportes_LineaEnCurso() As Double  ' Para almacenar importes l�nea en curso
    Dim vArrayImportes_SumaLineas() As Double    ' Para acumular importes
    
    ' Dimensionar arrays para almacenar importes (12 columnas)
    ReDim vArrayImportes_LineaAncla(1 To 12)
    ReDim vArrayImportes_LineaEnCurso(1 To 12)
    ReDim vArrayImportes_SumaLineas(1 To 12)
    
    ' Variables para c�lculo de n�mero m�ximo de duplicados y b�squeda de l�neas vac�as
    Dim vNumeroMaximoDuplicados As Long
    
    ' 7.1.9. Recorrer todas las l�neas para buscar duplicados no tratados
    lngLineaError = 250
    fun801_LogMessage "Recorriendo l�neas para procesar duplicados", False, strFilePath, vNuevaHojaImportacion_Working
    
    ' Bucle principal para recorrer todas las l�neas
    For i = vLineaInicial_HojaImportacion To vLineaFinal_HojaImportacion
        ' 7.1.9.1. Para cada l�nea, obtenemos los valores de las columnas de tags
        vValorColumna_TagLineaRepetida = CStr(wsWorking.Cells(i, lngColBase + 24).Value)
        vValorColumna_TagLineaTratada = CStr(wsWorking.Cells(i, lngColBase + 25).Value)
        
        ' 7.1.9.1.3. Verificar si es una l�nea repetida no tratada
        'If vValorColumna_TagLineaTratada <> vTagLineaTratada And vValorColumna_TagLineaRepetida = vTagLineaRepetida Then
        If vValorColumna_TagLineaTratada <> CONST_TAG_LINEA_TRATADA And vValorColumna_TagLineaRepetida = CONST_TAG_LINEA_REPETIDA Then
        
            ' 7.1.9.1.3.0. Tomar n�mero de l�nea actual como ancla
            vLineaAncla = i
            
            ' 7.1.9.1.3.1. Obtener valor concatenado de la l�nea ancla
            vValorColumna_Campos01a11_LineaAncla = CStr(wsWorking.Cells(vLineaAncla, lngColBase + 23).Value)
            
            ' 7.1.9.1.3.2. y 7.1.9.1.3.3. Almacenar valores de la l�nea ancla
            ' No necesitamos almacenar los valores de las columnas de identificaci�n
            ' ya que despu�s usamos Copy para copiarlos directamente
            
            ' 7.1.9.1.3.3. Almacenar importes de la l�nea ancla
            For k = 1 To 12
                vArrayImportes_LineaAncla(k) = CDbl(IIf(IsNumeric(wsWorking.Cells(vLineaAncla, lngColBase + 10 + k).Value), _
                                                       wsWorking.Cells(vLineaAncla, lngColBase + 10 + k).Value, 0))
            Next k
            
            ' 7.1.9.1.3.4. Inicializar array de suma con valores de la l�nea ancla
            For k = 1 To 12
                vArrayImportes_SumaLineas(k) = vArrayImportes_LineaAncla(k)
            Next k
            
            ' 7.1.9.1.3.5. Buscar l�neas duplicadas a esta l�nea ancla
            lngLineaError = 280
            For j = vLineaAncla + 1 To vLineaFinal_HojaImportacion
                ' 7.1.9.1.3.5.1. Tomar l�nea actual como l�nea en curso
                vLineaEnCurso = j
                
                ' 7.1.9.1.3.5.2. a 7.1.9.1.3.5.4. Obtener valores de la l�nea en curso
                vValorColumna_Campos01a11_LineaEnCurso = CStr(wsWorking.Cells(vLineaEnCurso, lngColBase + 23).Value)
                vValorColumna_TagLineaRepetida = CStr(wsWorking.Cells(vLineaEnCurso, lngColBase + 24).Value)
                vValorColumna_TagLineaTratada = CStr(wsWorking.Cells(vLineaEnCurso, lngColBase + 25).Value)
                
                ' 7.1.9.1.3.5.5. Verificar si es una l�nea duplicada no tratada
                'If vValorColumna_TagLineaTratada <> vTagLineaTratada And vValorColumna_TagLineaRepetida = vTagLineaRepetida Then
                If vValorColumna_TagLineaTratada <> CONST_TAG_LINEA_TRATADA And vValorColumna_TagLineaRepetida = CONST_TAG_LINEA_REPETIDA Then
                    ' 7.1.9.1.3.5.5.1.1. Verificar si el contenido coincide con la l�nea ancla
                    If vValorColumna_Campos01a11_LineaAncla = vValorColumna_Campos01a11_LineaEnCurso Then
                        ' 7.1.9.1.3.5.5.1.1.1. Almacenar importes de la l�nea en curso
                        For k = 1 To 12
                            vArrayImportes_LineaEnCurso(k) = CDbl(IIf(IsNumeric(wsWorking.Cells(vLineaEnCurso, lngColBase + 10 + k).Value), _
                                                                   wsWorking.Cells(vLineaEnCurso, lngColBase + 10 + k).Value, 0))
                            
                            ' 7.1.9.1.3.5.5.1.1.2. Sumar a los importes acumulados
                            vArrayImportes_SumaLineas(k) = vArrayImportes_SumaLineas(k) + vArrayImportes_LineaEnCurso(k)
                        Next k
                        
                        ' Marcar l�nea en curso como tratada
                        'wsWorking.Cells(vLineaEnCurso, lngColBase + 25).Value = vTagLineaTratada
                        wsWorking.Cells(vLineaEnCurso, lngColBase + 25).Value = CONST_TAG_LINEA_TRATADA
                    End If
                    ' Si no coincide el contenido, no hacemos nada con esta l�nea
                End If
                ' Si no es l�nea repetida no tratada, no hacemos nada con esta l�nea
            Next j
            
            ' 7.1.9.1.3.6. Buscar l�nea vac�a para insertar l�nea de suma
            lngLineaError = 300
            
            ' 7.1.9.1.3.6.0. y 7.1.9.1.3.6.1. Calcular n�mero m�ximo de duplicados
            vNumeroMaximoDuplicados = (vLineaFinal_HojaImportacion - vLineaInicial_HojaImportacion) / 2
            vNumeroMaximoDuplicados = Application.WorksheetFunction.RoundUp(vNumeroMaximoDuplicados, 0)
            
            ' 7.1.9.1.3.6.2. Buscar l�nea vac�a despu�s del rango de datos
            For m = vLineaFinal_HojaImportacion + 1 To vLineaFinal_HojaImportacion + vNumeroMaximoDuplicados
                ' 7.1.9.1.3.6.2.1. Verificar si la l�nea est� disponible
                vValorColumna_TagLineaSuma = CStr(wsWorking.Cells(m, lngColBase + 26).Value)
                
                ' 7.1.9.1.3.6.2.1.3. y 7.1.9.1.3.6.2.1.4. Verificar si podemos usar esta l�nea
                'If vValorColumna_TagLineaSuma = vTagLineaSuma Then
                If vValorColumna_TagLineaSuma = CONST_TAG_LINEA_SUMA Then
                
                    ' Si ya es una l�nea de suma, continuar con la siguiente
                    ' No hacemos nada con esta l�nea
                ElseIf vValorColumna_TagLineaSuma = "" Then
                    ' 7.1.9.1.3.6.2.1.4.1. Copiar valores de identificaci�n
                    wsWorking.Range(wsWorking.Cells(vLineaAncla, lngColBase), _
                                   wsWorking.Cells(vLineaAncla, lngColBase + 10)).Copy _
                                   wsWorking.Cells(m, lngColBase)
                    
                    ' 7.1.9.1.3.6.2.1.4.2. Escribir los importes sumados
                    For k = 1 To 12
                        wsWorking.Cells(m, lngColBase + 10 + k).Value = vArrayImportes_SumaLineas(k)
                    Next k
                    
                    ' 7.1.9.1.3.6.2.1.4.3. Limpiar columnas de tags
                    wsWorking.Cells(m, lngColBase + 23).Value = ""
                    wsWorking.Cells(m, lngColBase + 24).Value = ""
                    wsWorking.Cells(m, lngColBase + 25).Value = ""
                    
                    ' 7.1.9.1.3.6.2.1.4.4. A�adir tag de l�nea suma
                    'wsWorking.Cells(m, lngColBase + 26).Value = vTagLineaSuma
                    wsWorking.Cells(m, lngColBase + 26).Value = CONST_TAG_LINEA_SUMA
                    
                    ' 7.1.9.1.3.6.2.1.4.5. Salir del bucle
                    Exit For
                End If
            Next m
            
            ' 7.1.9.1.3.7. Marcar la l�nea ancla como tratada
            'wsWorking.Cells(vLineaAncla, lngColBase + 25).Value = vTagLineaTratada
            wsWorking.Cells(vLineaAncla, lngColBase + 25).Value = CONST_TAG_LINEA_TRATADA
            
            ' 7.1.9.1.3.7. Limpiar variables y arrays
            For k = 1 To 12
                vArrayImportes_LineaAncla(k) = 0
                vArrayImportes_LineaEnCurso(k) = 0
                vArrayImportes_SumaLineas(k) = 0
            Next k
            
            vLineaAncla = 0
            vValorColumna_Campos01a11_LineaAncla = ""
            vLineaEnCurso = 0
            vValorColumna_Campos01a11_LineaEnCurso = ""
            vValorColumna_TagLineaRepetida = ""
            vValorColumna_TagLineaTratada = ""
            vValorColumna_TagLineaSuma = ""
            j = 0
            k = 0
            m = 0
        End If
        ' Si no es l�nea repetida no tratada, continuamos con la siguiente
    Next i
    
    fun801_LogMessage "Procesamiento complementario de l�neas duplicadas completado", False, strFilePath, vNuevaHojaImportacion_Working
    
    '--------------------------------------------------------------------------
    ' 8. Ajustar zoom de la hoja de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 400
    fun801_LogMessage "Configurando zoom de la hoja de trabajo", False, strFilePath, vNuevaHojaImportacion_Working
    
    ' Definir variable de zoom
    Dim vZoom As Long
    vZoom = 70 ' Establecer zoom al 70%
    
    ' Configurar zoom de la hoja - compatible con Excel 97-365
    On Error Resume Next
    
    ' Activar la hoja para asegurarnos que es la activa
    ThisWorkbook.Worksheets(vNuevaHojaImportacion_Working).Activate
    
    ' Intentar establecer el zoom (m�todo 1)
    ActiveWindow.Zoom = vZoom
    
    ' Si falla, intentar m�todo alternativo
    If Err.Number <> 0 Then
        ' M�todo alternativo para Excel m�s antiguo
        Err.Clear
        With ActiveWindow
            .WindowState = xlNormal
            .Zoom = vZoom
        End With
    End If
    
    On Error GoTo GestorErrores
    
    fun801_LogMessage "Zoom configurado al " & vZoom & "%", False, strFilePath, vNuevaHojaImportacion_Working
        

    
    ' Proceso completado exitosamente
    fun801_LogMessage "Proceso de importaci�n completado con �xito", False, strFilePath, vNuevaHojaImportacion_Working
    F002_Importar_Fichero = True
    Exit Function

GestorErrores:
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, strFilePath, IIf(Len(vNuevaHojaImportacion_Working) > 0, _
                                                              vNuevaHojaImportacion_Working, _
                                                              vNuevaHojaImportacion)
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F002_Importar_Fichero = False
    
End Function

Public Function F004_Detectar_Delimitadores_en_Excel() As Boolean
    
    ' =============================================================================
    ' FUNCI�N PRINCIPAL: F004_Detectar_Delimitadores_en_Excel
    ' =============================================================================
    ' Fecha y hora de creaci�n: 2025-05-26 17:43:59 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    ' Descripci�n: Detecta y almacena los delimitadores de Excel actuales
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializar variables globales con valores por defecto
    ' 2. Verificar si existe la hoja de delimitadores originales
    ' 3. Si no existe, crear la hoja y dejarla visible
    ' 4. Si existe, verificar su visibilidad y hacerla visible si est� oculta
    ' 5. Limpiar el contenido de la hoja una vez visible
    ' 6. Configurar headers en las celdas especificadas (B2, B3, B4)
    ' 7. Detectar configuraci�n actual de delimitadores de Excel:
    '    - Use System Separators (True/False)
    '    - Decimal Separator (car�cter)
    '    - Thousands Separator (car�cter)
    ' 8. Almacenar valores detectados en variables globales
    ' 9. Escribir valores en la hoja de delimitadores (C2, C3, C4)
    ' 10. Verificar constante global CONST_OCULTAR_REPOSITORIO_DELIMITADORES
    ' 11. Si es True, ocultar la hoja creada/actualizada
    ' 12. Manejo exhaustivo de errores con informaci�n detallada
    '
    ' Par�metros: Ninguno
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    ' Control de errores con n�mero de l�nea
    On Error GoTo ErrorHandler
    
    ' Variables locales
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim hojaExiste As Boolean
    Dim i As Integer
    Dim lineaError As Long
    
    ' Inicializar resultado como exitoso
    F004_Detectar_Delimitadores_en_Excel = True
    
    ' ==========================================================================
    ' PASO 1: INICIALIZAR VARIABLES GLOBALES CON VALORES POR DEFECTO
    ' ==========================================================================
    lineaError = 100
    
    ' Nombre de la hoja donde se almacenar�n los delimitadores originales
    vHojaDelimitadoresExcelOriginales = CONST_HOJA_DELIMITADORES_ORIGINALES
    
    ' Celdas para los headers (t�tulos)
    vCelda_Header_Excel_UseSystemSeparators = "B2"
    vCelda_Header_Excel_DecimalSeparator = "B3"
    vCelda_Header_Excel_ThousandsSeparator = "B4"
    
    ' Celdas para los valores detectados
    vCelda_Valor_Excel_UseSystemSeparators = "C2"
    vCelda_Valor_Excel_DecimalSeparator = "C3"
    vCelda_Valor_Excel_ThousandsSeparator = "C4"
    
    ' Variables para almacenar los valores detectados (inicialmente vac�as)
    vExcel_UseSystemSeparators = ""
    vExcel_DecimalSeparator = ""
    vExcel_ThousandsSeparator = ""
    
    lineaError = 110
    
    ' ==========================================================================
    ' PASO 2: OBTENER REFERENCIA AL LIBRO ACTUAL
    ' ==========================================================================
    
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    If wb Is Nothing Then
        F004_Detectar_Delimitadores_en_Excel = False
        Exit Function
    End If
    
    lineaError = 120
    
    ' ==========================================================================
    ' PASO 3: VERIFICAR SI EXISTE LA HOJA DE DELIMITADORES ORIGINALES
    ' ==========================================================================
    
    hojaExiste = fun801_VerificarExistenciaHoja(wb, vHojaDelimitadoresExcelOriginales)
    
    lineaError = 130
    
    ' ==========================================================================
    ' PASO 4: CREAR HOJA O VERIFICAR VISIBILIDAD SEG�N CORRESPONDA
    ' ==========================================================================
    
    If Not hojaExiste Then
        ' La hoja no existe, crearla y dejarla visible
        Set ws = fun802_CrearHojaDelimitadores(wb, vHojaDelimitadoresExcelOriginales)
        If ws Is Nothing Then
            F004_Detectar_Delimitadores_en_Excel = False
            Exit Function
        End If
        ' La hoja reci�n creada ya est� visible por defecto
    Else
        ' La hoja existe, obtener referencia y verificar visibilidad
        Set ws = wb.Worksheets(vHojaDelimitadoresExcelOriginales)
        
        ' Verificar si est� oculta y hacerla visible si es necesario
        Call fun803_HacerHojaVisible(ws)
    End If
    
    lineaError = 140
    
    ' ==========================================================================
    ' PASO 5: LIMPIAR CONTENIDO DE LA HOJA (AHORA QUE EST� VISIBLE)
    ' ==========================================================================
    
    Call fun804_LimpiarContenidoHoja(ws)
    
    lineaError = 150
    
    ' ==========================================================================
    ' PASO 6: CONFIGURAR HEADERS EN LAS CELDAS ESPECIFICADAS
    ' ==========================================================================
    
    ' Header para Use System Separators en B2
    ws.Range(vCelda_Header_Excel_UseSystemSeparators).Value = "Excel Use System Separators"
    
    ' Header para Decimal Separator en B3
    ws.Range(vCelda_Header_Excel_DecimalSeparator).Value = "Excel Decimals"
    
    ' Header para Thousands Separator en B4
    ws.Range(vCelda_Header_Excel_ThousandsSeparator).Value = "Excel Thousands"
    
    lineaError = 160
    
    ' ==========================================================================
    ' PASO 7: DETECTAR CONFIGURACI�N ACTUAL DE DELIMITADORES DE EXCEL
    ' ==========================================================================
    
    ' Detectar Use System Separators
    vExcel_UseSystemSeparators = fun805_DetectarUseSystemSeparators()
    
    ' Detectar Decimal Separator
    vExcel_DecimalSeparator = fun806_DetectarDecimalSeparator()
    
    ' Detectar Thousands Separator
    vExcel_ThousandsSeparator = fun807_DetectarThousandsSeparator()
    
    lineaError = 170
    
    ' ==========================================================================
    ' PASO 8: ALMACENAR VALORES DETECTADOS EN LA HOJA
    ' ==========================================================================
    
    ' Almacenar Use System Separators en C2
    ws.Range(vCelda_Valor_Excel_UseSystemSeparators).Value = vExcel_UseSystemSeparators
    
    ' Almacenar Decimal Separator en C3
    ws.Range(vCelda_Valor_Excel_DecimalSeparator).Value = vExcel_DecimalSeparator
    
    ' Almacenar Thousands Separator en C4
    ws.Range(vCelda_Valor_Excel_ThousandsSeparator).Value = vExcel_ThousandsSeparator
    
    lineaError = 180
    
    ' ==========================================================================
    ' PASO 9: VERIFICAR SI DEBE OCULTAR LA HOJA
    ' ==========================================================================
    
    ' Verificar la variable global CONST_OCULTAR_REPOSITORIO_DELIMITADORES
    If CONST_OCULTAR_REPOSITORIO_DELIMITADORES = True Then
        ' Ocultar la hoja de delimitadores
        If Not fun809_OcultarHojaDelimitadores(ws) Then
            Debug.Print "ADVERTENCIA: Error al ocultar la hoja " & vHojaDelimitadoresExcelOriginales & " - Funci�n: F004_Detectar_Delimitadores_en_Excel - " & Now()
            ' Nota: No es un error cr�tico, el proceso puede continuar
        End If
    End If
    lineaError = 190
    
    ' ==========================================================================
    ' PASO 10: FINALIZACI�N EXITOSA
    ' ==========================================================================
    
    Exit Function
    
ErrorHandler:
    ' ==========================================================================
    ' MANEJO EXHAUSTIVO DE ERRORES
    ' ==========================================================================
    
    F004_Detectar_Delimitadores_en_Excel = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: F004_Detectar_Delimitadores_en_Excel" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now() & vbCrLf & _
                   "USUARIO: david-joaquin-corredera-de-colsa"
    
    ' Mostrar mensaje de error (comentar si no se desea)
    ' MsgBox mensajeError, vbCritical, "Error en Detecci�n de Delimitadores"
    
    ' Log del error para debugging
    Debug.Print mensajeError
    
    ' Limpiar objetos
    Set ws = Nothing
    Set wb = Nothing
    
End Function



