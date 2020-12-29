Sub GuardarValida()
'///CREA LIBRO

Application.ScreenUpdating = False
    
    
CONVENIO = Sheets("CONTABILIZADOS").Cells(4, 2)
FECHA = Sheets("CONTABILIZADOS").Cells(6, 2)
Application.StatusBar = "Guardando Valida " & CONVENIO
If Trim(Sheets("CONTABILIZADOS").Range("b2")) = "TODOS" Then
Workbooks.Add
NuevoLibro = ActiveWorkbook.Name



    Sheets("Hoja1").Select
    Sheets("Hoja1").Name = "CONTABILIZADOS."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja2").Select
    Sheets("Hoja2").Name = "VALIDA CANCELADOS."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja3").Select
    Sheets("Hoja3").Name = "ACTIVOS."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja4").Select
    Sheets("Hoja4").Name = "CANCELADOS."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja5").Select
    Sheets("Hoja5").Name = "PAGOS."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja6").Select
    Sheets("Hoja6").Name = "NOVEDADES."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja7").Select
    Sheets("Hoja7").Name = "INSOLVENCIA."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja8").Select
    Sheets("Hoja8").Name = "Convenio."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja9").Select
    Sheets("Hoja9").Name = "VALIDACION."
    
    
    
    Workbooks("MACRO CRA MANUALES.xlsm").Activate
'//COPEA INFORMACION
'/Contabilizados
larg = Sheets("CONTABILIZADOS").Range("A" & Rows.Count).End(xlUp).Row
Sheets("CONTABILIZADOS").Range("A1:AW" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("CONTABILIZADOS.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/Valida Cancelados
Workbooks("MACRO CRA MANUALES.xlsm").Activate
Sheets("VALIDA CANCELADOS").Range("A1:G500").Copy
Workbooks(NuevoLibro).Activate
Sheets("VALIDA CANCELADOS.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/Activos
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("ACTIVOS").Range("A" & Rows.Count).End(xlUp).Row
Sheets("ACTIVOS").Range("A1:BP" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("ACTIVOS.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/Cancelados
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("CANCELADOS").Range("A" & Rows.Count).End(xlUp).Row
Sheets("CANCELADOS").Range("A1:BE" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("CANCELADOS.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/pagos
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("PAGOS").Range("A" & Rows.Count).End(xlUp).Row
Sheets("PAGOS").Range("A1:Y" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("PAGOS.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/Novedades
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("NOVEDADES").Range("A" & Rows.Count).End(xlUp).Row
Sheets("NOVEDADES").Range("A1:X" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("NOVEDADES.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/Insolvencia
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("INSOLVENCIA").Range("A" & Rows.Count).End(xlUp).Row
Sheets("INSOLVENCIA").Range("A1:K" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("INSOLVENCIA.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/Convenio
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("Convenio").Range("A" & Rows.Count).End(xlUp).Row
Sheets("Convenio").Range("A1:K" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("Convenio.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/VALIDACION
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("VALIDACION").Range("A" & Rows.Count).End(xlUp).Row
Sheets("VALIDACION").Range("A1:K" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("VALIDACION.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False

ActiveWindow.DisplayGridlines = False


'///crear carpeta y guardar
NombreConvenio = CONVENIO
FechaConvenio = Format(Now(), "MMMM_YYYY(HH°MM)")

    ruta = "D:\PRUEBAS INFORMES\VALIDAS HISTORICO\" & CONVENIO
    
    If Dir(ruta, vbDirectory) = "" Then
            MkDir (ruta)
        End If
        
 
  ActiveWorkbook.SaveAs Filename:=ruta & "\" & CONVENIO & FechaConvenio
  ActiveWorkbook.Close
End If

'////////////////////////////////  \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
If Trim(Sheets("CONTABILIZADOS").Range("b2")) = "NUEVOS" Then
Workbooks.Add
NuevoLibro = ActiveWorkbook.Name

'VALIDACION

    Sheets("Hoja1").Select
    Sheets("Hoja1").Name = "CONTABILIZADOS."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja2").Select
    Sheets("Hoja2").Name = "VALIDA CANCELADOS."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja3").Select
    Sheets("Hoja3").Name = "ACTIVOS."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja4").Select
    Sheets("Hoja4").Name = "CANCELADOS."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja5").Select
    Sheets("Hoja5").Name = "PAGOS."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja6").Select
    Sheets("Hoja6").Name = "VARIACION."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja7").Select
    Sheets("Hoja7").Name = "NOVEDADES."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja8").Select
    Sheets("Hoja8").Name = "REESTRUCTURADOS."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja9").Select
    Sheets("Hoja9").Name = "INSOLVENCIA."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja10").Select
    Sheets("Hoja10").Name = "Convenio."
    Sheets.Add After:=ActiveSheet
    Sheets("Hoja11").Select
    Sheets("Hoja11").Name = "VALIDACION."
    
    
    Workbooks("MACRO CRA MANUALES.xlsm").Activate
'//COPEA INFORMACION
'/Contabilizados
larg = Sheets("CONTABILIZADOS").Range("A" & Rows.Count).End(xlUp).Row
Sheets("CONTABILIZADOS").Range("A1:AW" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("CONTABILIZADOS.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/Valida Cancelados
Workbooks("MACRO CRA MANUALES.xlsm").Activate
Sheets("VALIDA CANCELADOS").Range("A1:G500").Copy
Workbooks(NuevoLibro).Activate
Sheets("VALIDA CANCELADOS.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/Activos
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("ACTIVOS").Range("A" & Rows.Count).End(xlUp).Row
Sheets("ACTIVOS").Range("A1:BZ" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("ACTIVOS.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/Cancelados
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("CANCELADOS").Range("A" & Rows.Count).End(xlUp).Row
Sheets("CANCELADOS").Range("A1:BE" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("CANCELADOS.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/pagos
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("PAGOS").Range("A" & Rows.Count).End(xlUp).Row
Sheets("PAGOS").Range("A1:Y" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("PAGOS.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/Variacion
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("VARIACION").Range("A" & Rows.Count).End(xlUp).Row
Sheets("VARIACION").Range("A1:Z" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("VARIACION.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/Novedades
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("NOVEDADES").Range("A" & Rows.Count).End(xlUp).Row
Sheets("NOVEDADES").Range("A1:X" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("NOVEDADES.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/Novedades
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("Restructurados").Range("A" & Rows.Count).End(xlUp).Row
Sheets("Restructurados").Range("A1:X" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("REESTRUCTURADOS.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False

'/Insolvencia
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("INSOLVENCIA").Range("A" & Rows.Count).End(xlUp).Row
Sheets("INSOLVENCIA").Range("A1:K" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("INSOLVENCIA.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/Convenio
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("Convenio").Range("A" & Rows.Count).End(xlUp).Row
Sheets("Convenio").Range("A1:K" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("Convenio.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
'/VALIDACION
Workbooks("MACRO CRA MANUALES.xlsm").Activate
larg = Sheets("VALIDACION").Range("A" & Rows.Count).End(xlUp).Row
Sheets("VALIDACION").Range("A1:K" & larg).Copy
Workbooks(NuevoLibro).Activate
Sheets("VALIDACION.").Activate
Range("a1").Select
ActiveSheet.Paste
ActiveWindow.DisplayGridlines = False
ActiveWindow.DisplayGridlines = False



'///crear carpeta y guardar
NombreConvenio = CONVENIO
FechaConvenio = Format(Now(), "MMMM_YYYY(HH°MM)")

    ruta = "D:\PRUEBAS INFORMES\VALIDAS HISTORICO\" & CONVENIO
    
    If Dir(ruta, vbDirectory) = "" Then
            MkDir (ruta)
        End If
        
 
  ActiveWorkbook.SaveAs Filename:=ruta & "\" & CONVENIO & "_" & FechaConvenio
  ActiveWorkbook.Close
End If



Call CONSOLIDADOR

 Application.ScreenUpdating = True
    
Application.DisplayAlerts = False


End Sub
Sub CONSOLIDADOR()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.ScreenUpdating = False
Workbooks.Open "D:\PRUEBAS INFORMES\CONSOLIDADO REPORTES ENVIADOS 2019.xlsx"

Workbooks("MACRO CRA MANUALES.xlsm").Activate

Final = Sheets("ACTIVOS").Range("A" & Rows.Count).End(xlUp).Row

For D = 2 To Final
Workbooks("MACRO CRA MANUALES.xlsm").Activate
If Sheets("ACTIVOS").Cells(D, 46).Value <> "NO REPORTAR" Then
 LargoConvenioNuevo = Workbooks("CONSOLIDADO REPORTES ENVIADOS 2019.xlsx").Sheets("NUEVOS").Range("B" & Rows.Count).End(xlUp).Row
    '**
    Workbooks("CONSOLIDADO REPORTES ENVIADOS 2019.xlsx").Sheets("NUEVOS").Cells(LargoConvenioNuevo + 1, 1) = Workbooks("MACRO CRA MANUALES.xlsm").Sheets("CONTABILIZADOS").Range("B4")
    '**
    Workbooks("CONSOLIDADO REPORTES ENVIADOS 2019.xlsx").Sheets("NUEVOS").Cells(LargoConvenioNuevo + 1, 2) = Workbooks("MACRO CRA MANUALES.xlsm").Sheets("ACTIVOS").Cells(D, 1)
    '***
    Workbooks("CONSOLIDADO REPORTES ENVIADOS 2019.xlsx").Sheets("NUEVOS").Cells(LargoConvenioNuevo + 1, 3) = Workbooks("MACRO CRA MANUALES.xlsm").Sheets("ACTIVOS").Cells(D, 13)
    '***
    Workbooks("CONSOLIDADO REPORTES ENVIADOS 2019.xlsx").Sheets("NUEVOS").Cells(LargoConvenioNuevo + 1, 4) = Workbooks("MACRO CRA MANUALES.xlsm").Sheets("ACTIVOS").Cells(D, 2)

    '***
    Workbooks("CONSOLIDADO REPORTES ENVIADOS 2019.xlsx").Sheets("NUEVOS").Cells(LargoConvenioNuevo + 1, 5) = Workbooks("MACRO CRA MANUALES.xlsm").Sheets("ACTIVOS").Cells(D, 45)
    '***
    Workbooks("CONSOLIDADO REPORTES ENVIADOS 2019.xlsx").Sheets("NUEVOS").Cells(LargoConvenioNuevo + 1, 6) = Workbooks("MACRO CRA MANUALES.xlsm").Sheets("ACTIVOS").Cells(D, 46)
    '***
    Workbooks("CONSOLIDADO REPORTES ENVIADOS 2019.xlsx").Sheets("NUEVOS").Cells(LargoConvenioNuevo + 1, 8) = Workbooks("MACRO CRA MANUALES.xlsm").Sheets("NOVEDADES").Range("XFA1")
    
    
End If

Next

'******************************* CANCELADOS *******************************

Workbooks("MACRO CRA MANUALES.xlsm").Activate

Final = Sheets("CANCELADOS").Range("A" & Rows.Count).End(xlUp).Row

For D = 2 To Final
Workbooks("MACRO CRA MANUALES.xlsm").Activate
If Sheets("CANCELADOS").Cells(D, 49).Value <> "" Then
 LargoConvenioNuevo = Workbooks("CONSOLIDADO REPORTES ENVIADOS 2019.xlsx").Sheets("CANCELAD").Range("B" & Rows.Count).End(xlUp).Row
    '**
    Workbooks("CONSOLIDADO REPORTES ENVIADOS 2019.xlsx").Sheets("CANCELAD").Cells(LargoConvenioNuevo + 1, 1) = Workbooks("MACRO CRA MANUALES.xlsm").Sheets("CONTABILIZADOS").Range("B4")
    '**
    Workbooks("CONSOLIDADO REPORTES ENVIADOS 2019.xlsx").Sheets("CANCELAD").Cells(LargoConvenioNuevo + 1, 2) = Workbooks("MACRO CRA MANUALES.xlsm").Sheets("CANCELADOS").Cells(D, 1)
    '***
    Workbooks("CONSOLIDADO REPORTES ENVIADOS 2019.xlsx").Sheets("CANCELAD").Cells(LargoConvenioNuevo + 1, 3) = Workbooks("MACRO CRA MANUALES.xlsm").Sheets("CANCELADOS").Cells(D, 13)
    '***
    Workbooks("CONSOLIDADO REPORTES ENVIADOS 2019.xlsx").Sheets("CANCELAD").Cells(LargoConvenioNuevo + 1, 4) = Workbooks("MACRO CRA MANUALES.xlsm").Sheets("CANCELADOS").Cells(D, 2)

    '***
    Workbooks("CONSOLIDADO REPORTES ENVIADOS 2019.xlsx").Sheets("CANCELAD").Cells(LargoConvenioNuevo + 1, 6) = Workbooks("MACRO CRA MANUALES.xlsm").Sheets("NOVEDADES").Range("XFA1")

    
End If

Next

Workbooks("CONSOLIDADO REPORTES ENVIADOS 2019.xlsx").Activate
ActiveWorkbook.Save
ActiveWorkbook.Close

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
 Application.ScreenUpdating = True
    
Application.DisplayAlerts = False
End Sub
