Sub AutomatizarControlSalarial()
    ' ============================================================
    ' MACRO: AutomatizarControlSalarial
    ' PERSPECTIVA: Gestión de Compensaciones y Capital Humano
    ' OBJETIVO: Formatear y calcular el nuevo sueldo de 100 empleados
    ' ============================================================
    
    Dim Hoja As Worksheet
    Dim UltimaFila As Long
    
    ' 1. Definir la hoja de trabajo activa
    Set Hoja = ActiveSheet
    
    ' 2. Encontrar dinámicamente la última fila con datos en la columna A (ID Empleado)
    ' Esto permite que funcione con 100 registros o si se agregan más en el futuro
    UltimaFila = Hoja.Cells(Hoja.Rows.Count, "A").End(xlUp).Row
    
    ' Validar que existan registros además del encabezado
    If UltimaFila < 2 Then
        MsgBox "No se encontraron registros de empleados para procesar.", vbExclamation, "Control Interno"
        Exit Sub
    End If
    
    ' 3. AUTOMATIZACIÓN DE CÁLCULOS Y FORMATOS
    Application.ScreenUpdating = False ' Optimiza la velocidad de ejecución bloqueando el parpadeo
    
    With Hoja
        ' A. Insertar la fórmula de cálculo en la Columna F (Nuevo_Sueldo)
        ' Fórmula: Sueldo_Anterior * (1 + Incremento_Porcentaje)
        .Range("F2:F" & UltimaFila).FormulaLocal = "=D2*(1+E2)"
        
        ' B. Aplicar formato ejecutivo de Moneda ($) con dos decimales a los sueldos
        .Range("D2:D" & UltimaFila).NumberFormat = "$#,##0.00" ' Sueldo Anterior
        .Range("F2:F" & UltimaFila).NumberFormat = "$#,##0.00" ' Nuevo Sueldo
        
        ' C. Aplicar formato de Porcentaje (%) a la columna de incrementos
        .Range("E2:E" & UltimaFila).NumberFormat = "0.10"
        
        ' D. Ajustar automáticamente el ancho de las columnas para evitar el error de visualización (###)
        .Columns("A:F").AutoFit
    End With
    
    Application.ScreenUpdating = True
    
    ' 4. Notificación de cierre de proceso al Administrador
    MsgBox "Proceso de nómina completado exitosamente." & vbCrLf & _
           "Registros procesados: " & (UltimaFila - 1) & vbCrLf & _
           "Cálculos y formatos corporativos aplicados.", vbInformation, "Éxito Administrativo"
           
End Sub
