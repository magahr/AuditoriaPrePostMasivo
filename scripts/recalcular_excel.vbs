' Recalcular_excel.vbs
'
' Este script abre un archivo de Excel, fuerza el recálculo y lo guarda.
'
' Uso: cscript recalcular_excel.vbs "C:\ruta\al\archivo.xlsx"

Dim objExcel, objWorkbook
Dim sFilePath

Set fso = CreateObject("Scripting.FileSystemObject")

' Obtiene la ruta del archivo del primer argumento de la línea de comandos
sFilePath = WScript.Arguments(0)

' Verifica que el archivo exista antes de intentar abrirlo
If fso.FileExists(sFilePath) Then
    ' Crea una instancia de la aplicación de Excel
    Set objExcel = CreateObject("Excel.Application")

    ' Mantiene Excel visible (para depuración) o invisible
    objExcel.Visible = False
    
    ' Abre el libro de trabajo
    Set objWorkbook = objExcel.Workbooks.Open(sFilePath)

    ' Fuerza un recálculo completo de todas las fórmulas
    objExcel.Calculation = -4105 ' xlCalculationManual
    ' 06-08-2025 cambio
    objWorkbook.ForceFullCalculation = True
    objExcel.Calculate
    objExcel.Calculation = -4107 ' xlCalculationAutomatic

    ' Guarda los cambios
    objWorkbook.Save
    
    ' Cierra el libro y la aplicación de Excel
    objWorkbook.Close
    objExcel.Quit
    
    ' Limpia los objetos
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    Set fso = Nothing
Else
    WScript.Echo "Error: El archivo no existe en la ruta especificada."
End If

WScript.Quit 0