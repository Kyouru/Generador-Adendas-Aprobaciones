Private Sub btFechaRenovacion_Click()
    Hoja1.Unprotect pword
    frmCalendario.Show
    If Hoja1.Range("FECHA") <> "" Then
        Hoja1.Range("FECHA") = "Lima, " & txtFecha2(DateSerial(Format(Hoja1.Range("TMPFECHA"), "YYYY"), Format(Hoja1.Range("TMPFECHA"), "MM"), Format(Hoja1.Range("TMPFECHA"), "DD"))) & "."
    Else
        Hoja1.Range("FECHA") = "Lima, " & txtFecha2(Now()) & "."
    End If
    Hoja1.Protect Password:=pword, DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Private Sub btRegresar_Click()
    Hoja3.Activate
End Sub

Public Sub PlantillaAdenda()
    Hoja1.Unprotect pword
    Hoja1.Range("ADENDA1") = "Conste por el presente documento, la ADENDA AL CONTRATO DE CRﾉDITO DE FECHA #FECHAC#, que celebran, de una parte, la COOPERATIVA DE AHORRO Y CRﾉDITO PACIFICO, con RUC No. 20111065013 y domicilio en Calle las Orquideas No. 675, Piso 3, Distrito de San Isidro, Provincia y Departamento de Lima, debidamente representada por su Apoderado del Grupo 釘・ ser Jorge Armando Ouchida Noda, identificado con DNI No. 07912017, facultado seg佖 poderes inscritos en el asiento C00037 de la Partida No. 02021617 del Libro de Cooperativas del Registro de Personas Jur冝icas de Lima,  a quien en adelante se denominar・鏑A COOPERATIVA・  y, de la otra parte, 摘L(LA) SOCIO(A)・ cuyo nombre y generales de ley constan al final de este documento; en los t駻minos y condiciones siguientes:"
    Hoja1.Range("ADENDA1.1") = "Con fecha #FECHAC#, las partes celebraron un Contrato de Cr馘ito (en adelante, el 鼎ONTRATO・, mediante el cual LA COOPERATIVA otorg・a EL(LA) SOCIO(A) el pr駸tamo denominado #CODPRESTAMO# ascendente a la cantidad de #MONTOTEXTO# (en adelante, el 鼎RﾉDITO・."
    Hoja1.Range("ADENDA1.2") = "El CRﾉDITO debe ser devuelto en un plazo de #PLAZOACTUAL# d僘s, contados a partir de la fecha del desembolso, habi駭dose pactado una tasa de inter駸 compensatorio de #TASA# anual (TEA)."
    Hoja1.Range("ADENDA2") = "Por la presente Adenda, las partes convienen en renovar el plazo del CRﾉDITO, el mismo que queda establecido en #AMPLIACION# d僘s adicionales, contados a partir del vencimiento del plazo selado en el numeral 1.2 de la Cl疼sula Primera precedente, con vencimiento al #VENCIMIENTO#."
    Hoja1.Range("FECHA") = "Lima, " & txtFecha2(Now()) & "."
    Hoja1.Protect Password:=pword, DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Public Sub FormatoAdenda()
    Hoja1.Range("ADENDA1").Font.Bold = False
    Hoja1.Range("ADENDA1").Characters(InStr(Hoja1.Range("ADENDA1").Value, "ADENDA AL CONTRATO DE CRﾉDITO DE FECHA"), Len("ADENDA AL CONTRATO DE CRﾉDITO DE FECHA") + 1).Font.Bold = True
    Hoja1.Range("ADENDA1").Characters(InStr(Hoja1.Range("ADENDA1").Value, "COOPERATIVA DE AHORRO Y CRﾉDITO PACIFICO"), Len("COOPERATIVA DE AHORRO Y CRﾉDITO PACIFICO")).Font.Bold = True
    Hoja1.Range("ADENDA1").Characters(InStr(Hoja1.Range("ADENDA1").Value, "鏑A COOPERATIVA・), Len("鏑A COOPERATIVA・)).Font.Bold = True
    Hoja1.Range("ADENDA1").Characters(InStr(Hoja1.Range("ADENDA1").Value, "摘L(LA) SOCIO(A)・), Len("摘L(LA) SOCIO(A)・)).Font.Bold = True
    Hoja1.Range("ADENDA1.1").Font.Bold = False
    Hoja1.Range("ADENDA1.1").Font.Underline = False
    Hoja1.Range("ADENDA1.1").Characters(InStr(Hoja1.Range("ADENDA1.1").Value, "鼎ONTRATO・), Len("鼎ONTRATO・)).Font.Bold = True
    Hoja1.Range("ADENDA1.1").Characters(InStr(Hoja1.Range("ADENDA1.1").Value, "鼎ONTRATO・) + 1, Len("CONTRATO")).Font.Underline = True
    Hoja1.Range("ADENDA1.1").Characters(InStr(Hoja1.Range("ADENDA1.1").Value, "LA COOPERATIVA"), Len("LA COOPERATIVA")).Font.Bold = True
    Hoja1.Range("ADENDA1.1").Characters(InStr(Hoja1.Range("ADENDA1.1").Value, "EL(LA) SOCIO(A)"), Len("EL(LA) SOCIO(A)")).Font.Bold = True
    Hoja1.Range("ADENDA1.1").Characters(InStr(Hoja1.Range("ADENDA1.1").Value, "鼎RﾉDITO・), Len("鼎RﾉDITO・)).Font.Bold = True
    Hoja1.Range("ADENDA1.1").Characters(InStr(Hoja1.Range("ADENDA1.1").Value, "鼎RﾉDITO・) + 1, Len("CRﾉDITO")).Font.Underline = True
    Hoja1.Range("ADENDA1.2").Font.Bold = False
    Hoja1.Range("ADENDA1.2").Characters(InStr(Hoja1.Range("ADENDA1.2").Value, "CRﾉDITO"), Len("CRﾉDITO")).Font.Bold = True
    Hoja1.Range("ADENDA2").Font.Bold = False
    Hoja1.Range("ADENDA2").Characters(InStr(Hoja1.Range("ADENDA2").Value, "CRﾉDITO"), Len("CRﾉDITO")).Font.Bold = True
End Sub

Public Sub LlenarAdenda()
    If Hoja3.tbNomSocio.Text <> "" And Hoja3.tbNSol.Text <> "" And Hoja3.tbCodP.Text <> "" And Hoja3.tbFechaO.Text <> "" And Hoja3.tbMonto.Text <> "" And Hoja3.tbTEA.Text <> "" And Hoja3.Range("PLAZO").Value <> "" And Hoja3.Range("PLAZO_NUEVO").Value <> "" Then
    If IsDate(Hoja3.tbFechaO.Text) Then
    If IsDate(Hoja3.Range("FECHA_CONTRATO").Value) Then
    If IsDate(Hoja3.Range("FECHA_DESEMBOLSO").Value) Then
    If IsDate(Hoja3.Range("VENCIMIENTO").Value) Then
    If IsNumeric(Hoja3.Range("NUEVO_PLAZO_DIAS").Value) Then
    If IsNumeric(Hoja3.tbMonto.Text) Then
    If IsNumeric(Hoja3.tbTEA.Text) Then
    If IsNumeric(Hoja3.Range("PLAZO").Value) Then
    If IsNumeric(Hoja3.Range("PLAZO_NUEVO").Value) Then
    If Int(Hoja3.Range("PLAZO").Value) < Int(Hoja3.Range("PLAZO_NUEVO").Value) Then
        PlantillaAdenda
        Hoja1.Range("ADENDA1") = Replace(Hoja1.Range("ADENDA1"), "#FECHAC#", txtFecha(Hoja3.Range("FECHA_CONTRATO").Value))
        Hoja1.Range("ADENDA1.1") = Replace(Hoja1.Range("ADENDA1.1"), "#FECHAC#", txtFecha(Hoja3.Range("FECHA_CONTRATO").Value))
        Hoja1.Range("ADENDA1.1") = Replace(Hoja1.Range("ADENDA1.1"), "#CODPRESTAMO#", Hoja3.tbCodP.Value)
        If Hoja3.obSoles.Value Then
            Hoja1.Range("ADENDA1.1") = Replace(Hoja1.Range("ADENDA1.1"), "#MONTOTEXTO# ", NumLetras(Hoja3.tbMonto.Value, "Sol", "Soles") & " (S/." & Hoja3.tbMonto.Value & ") ")
        Else
            Hoja1.Range("ADENDA1.1") = Replace(Hoja1.Range("ADENDA1.1"), "#MONTOTEXTO# ", NumLetras(Hoja3.tbMonto.Value, "Dar", "Dares") & " (US$" & Hoja3.tbMonto.Value & ") ")
        End If
        Hoja1.Range("ADENDA1.2") = Replace(Hoja1.Range("ADENDA1.2"), "#PLAZOACTUAL#", Hoja3.Range("PLAZO").Value)
        Hoja1.Range("ADENDA1.2") = Replace(Hoja1.Range("ADENDA1.2"), "#TASA#", Format(Hoja3.tbTEA.Value / 100, "Percent"))
        Hoja1.Range("ADENDA2") = Replace(Hoja1.Range("ADENDA2"), "#AMPLIACION#", Hoja3.Range("PLAZO_NUEVO").Value - Hoja3.Range("PLAZO").Value)
        Hoja1.Range("ADENDA2") = Replace(Hoja1.Range("ADENDA2"), "#VENCIMIENTO#", txtFecha2(FVencimiento(Hoja3.Range("FECHA_DESEMBOLSO").Value, Hoja3.Range("PLAZO").Value, Hoja3.Range("PLAZO_NUEVO").Value - Hoja3.Range("PLAZO").Value, False)))
        Hoja1.Range("FECHA") = "Lima, " & txtFecha2(Now()) & "."
        Hoja1.Range("SOLICITUD").Value = Hoja3.tbNSol.Value
        Hoja1.Range("ADENDA_NOMBRE_S") = Hoja3.tbNomSocio.Text
        Hoja1.Range("ADENDA_DNI_S") = Hoja3.tbDNIS.Text
        Hoja1.Range("ADENDA_DIRECCION_S") = Hoja3.tbDirS.Text
        Hoja1.Range("ADENDA_NOMBRE_C") = Hoja3.tbNombreC.Text
        Hoja1.Range("ADENDA_DNI_C") = Hoja3.tbDNIC.Text
        Hoja1.Range("ADENDA_DIRECCION_C") = Hoja3.tbDirC.Text
        closeRS
        OpenDB
        strSQL = "INSERT INTO [HISTORICO$] VALUES ('" & maxID("HISTORICO", "ID") & "','ADENDA'"
        If Hoja3.obFuncionario Then
            strSQL = strSQL & ",'FUNCIONARIO','"
        Else
            If Hoja3.obComite Then
                strSQL = strSQL & ",'COMITE','"
            Else
                strSQL = strSQL & ",'CONSEJO','"
            End If
        End If
        strSQL = strSQL & Now() & "', '" & Hoja3.tbCSocio.Text & "', '" & Hoja3.tbNomSocio.Text & "', '" & Hoja3.tbNSol.Text & "', '" & Hoja3.tbCodP.Text & "', '" & Hoja3.tbFechaO.Text & "', '" & Hoja3.Range("FECHA_CONTRATO").Value & "', '" & Hoja3.tbMonto.Text & "', '"
        If Hoja3.obSoles Then
            strSQL = strSQL & "S', "
        Else
            strSQL = strSQL & "D', "
        End If
        strSQL = strSQL & "'" & Hoja3.tbTEA.Text & "','" & Hoja3.Range("FECHA_DESEMBOLSO") & "','" & Hoja3.Range("VENCIMIENTO") & "','" & DateAdd("d", Hoja3.Range("NUEVO_PLAZO_DIAS"), CDate(Hoja3.Range("VENCIMIENTO"))) & "','" & Hoja3.tbDNIS.Text & "','" & Hoja3.tbDirS.Text & "','" & Hoja3.tbNombreC.Text & "','" & Hoja3.tbDNIC.Text & "','" & Hoja3.tbDirC.Text & "','" & Hoja3.tbNRenov.Text & "','" & Replace(Hoja3.tbGarantias.Text, "'", "`") & "'"
        If Hoja3.obS Then
            strSQL = strSQL & ",'SOLTERO')"
            Hoja1.Range("ADENDA_ESTADO_S") = "SOLTERO"
            Hoja1.Range("ADENDA_ESTADO_C") = ""
            Hoja1.Range("ADENDA_ESTADO_S").EntireRow.Hidden = False
        ElseIf Hoja3.obC Then
            strSQL = strSQL & ",'CASADO')"
            Hoja1.Range("ADENDA_ESTADO_S") = "CASADO"
            Hoja1.Range("ADENDA_ESTADO_C") = "CASADO"
            Hoja1.Range("ADENDA_ESTADO_S").EntireRow.Hidden = False
        Else
            strSQL = strSQL & ",'EMPRESA')"
            Hoja1.Range("ADENDA_ESTADO_S").EntireRow.Hidden = True
        End If
        cnn.Execute (strSQL)
        FormatoAdenda
        Hoja1.Protect Password:=pword, DrawingObjects:=True, Contents:=True, Scenarios:=True
        Hoja1.Activate
    Else
        MsgBox "El Nuevo Plazo es inferior al Plazo Actual"
    End If
    Else
        MsgBox "Error en Plazo Nuevo"
        Hoja3.Range("PLAZO_NUEVO").Select
    End If
    Else
        MsgBox "Error en Plazo"
        Hoja3.Range("PLAZO").Select
    End If
    Else
        MsgBox "Error en TEA"
        Hoja3.tbTEA.Activate
    End If
    Else
        MsgBox "Error en Monto"
        Hoja3.tbMonto.Activate
    End If
    Else
        MsgBox "Error en Fecha Vencimiento de Tarjeta"
        Hoja3.Range("VENCIMIENTO_TARJETA").Select
    End If
    Else
        MsgBox "Error en Fecha Vencimiento"
        Hoja3.Range("VENCIMIENTO").Select
    End If
    Else
        MsgBox "Error en Fecha Desembolso"
        Hoja3.Range("FECHA_DESEMBOLSO").Select
    End If
    Else
        MsgBox "Error en Fecha Contrato"
        Hoja3.Range("FECHA_CONTRATO").Select
    End If
    Else
        MsgBox "Error en Fecha Otorgamiento"
        Hoja3.tbFechaO.Activate
        Hoja3.tbFechaO.SelStart = 0
        Hoja3.tbFechaO.SelLength = Len(Hoja3.tbFechaO.Text)
    End If
    Else
        MsgBox "Datos Insuficientes"
    End If
End Sub


Function FVencimiento(FOtorga As Date, Plazo As Double, Ampliacion As Double, DLaborables As Boolean) As Date
    Dim nDias As Double
    nDias = (Plazo + Ampliacion)
    FVencimiento = DateAdd("d", nDias, FOtorga)
    If Weekday(FVencimiento, vbMonday) = 7 And DLaborables Then
        FVencimiento = FVencimiento - 2
    End If
    If Weekday(FVencimiento, vbMonday) = 6 And DLaborables Then
        FVencimiento = FVencimiento - 1
    End If
End Function

Function NumLetras(Valor As Currency, Optional MonedaSingular As String = "", Optional MonedaPlural As String = "") As String
Dim lyCantidad As Currency, lyCentavos As Currency, lnDigito As Byte, lnPrimerDigito As Byte, lnSegundoDigito As Byte, lnTercerDigito As Byte, lcBloque As String, lnNumeroBloques As Byte, lnBloqueCero
Dim laUnidades As Variant, laDecenas As Variant, laCentenas As Variant, i As Variant 'Si esta como Option Explicit
Dim ValorEntero As Long
Valor = Round(Valor, 2)
lyCantidad = Int(Valor)
ValorEntero = lyCantidad
lyCentavos = (Valor - lyCantidad) * 100
laUnidades = Array("UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUN", "VEINTIDOS", "VEINTITRES", "VEINTICUATRO", "VEINTICINCO", "VEINTISEIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE")
laDecenas = Array("DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
laCentenas = Array("CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
lnNumeroBloques = 1
 
Do
lnPrimerDigito = 0
lnSegundoDigito = 0
lnTercerDigito = 0
lcBloque = ""
lnBloqueCero = 0
For i = 1 To 3
lnDigito = lyCantidad Mod 10
If lnDigito <> 0 Then
Select Case i
Case 1
lcBloque = " " & laUnidades(lnDigito - 1)
lnPrimerDigito = lnDigito
Case 2
If lnDigito = 2 Then
lcBloque = " " & laUnidades((lnDigito * 10) + lnPrimerDigito - 1)
Else
lcBloque = " " & laDecenas(lnDigito - 1) & IIf(lnPrimerDigito <> 0, " Y", Null) & lcBloque
End If
lnSegundoDigito = lnDigito
Case 3
lcBloque = " " & IIf(lnDigito = 1 And lnPrimerDigito = 0 And lnSegundoDigito = 0, "CIEN", laCentenas(lnDigito - 1)) & lcBloque
lnTercerDigito = lnDigito
End Select
Else
lnBloqueCero = lnBloqueCero + 1
End If
lyCantidad = Int(lyCantidad / 10)
If lyCantidad = 0 Then
Exit For
End If
Next i
Select Case lnNumeroBloques
Case 1
NumLetras = lcBloque
Case 2
NumLetras = lcBloque & IIf(lnBloqueCero = 3, Null, " MIL") & NumLetras
Case 3
NumLetras = lcBloque & IIf(lnPrimerDigito = 1 And lnSegundoDigito = 0 And lnTercerDigito = 0, " MILLON", " MILLONES") & NumLetras
End Select
lnNumeroBloques = lnNumeroBloques + 1
Loop Until lyCantidad = 0
NumLetras = NumLetras & " CON " & Format(Str(lyCentavos), "00") & "/100 " & IIf(ValorEntero = 1, MonedaSingular, MonedaPlural)
End Function

Function txtFecha(dateFecha As Date) As String
    Dim txtMeses As Variant
    txtMeses = Array("ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")
    txtFecha = Day(dateFecha) & " DE " & txtMeses(Month(dateFecha) - 1) & " DEL " & Year(dateFecha)
End Function

Function txtFecha2(dateFecha As Date) As String
    Dim txtMeses As Variant
    txtMeses = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    txtFecha2 = Day(dateFecha) & " de " & txtMeses(Month(dateFecha) - 1) & " del " & Year(dateFecha)
End Function


Function maxID(hoja As String, campo As String) As String
    Dim strSQL2 As String
    Dim strSQL3 As String
    Dim sw As Boolean
    sw = False
    strSQL2 = "SELECT COUNT(*) FROM [" & hoja & "$]"
    strSQL3 = "SELECT MAX(" & campo & ") FROM [" & hoja & "$]"
    OpenDB
    rs.Open strSQL2, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        If rs.Fields(0).Value > 0 Then
            sw = True
        Else
            maxID = "00000001"
        End If
    End If
    closeRS
    If sw Then
        OpenDB
        rs.Open strSQL3, cnn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            Dim strID As String
            strID = rs.Fields(0) + 1
            Do While Len(strID) < 8
                strID = "0" & strID
            Loop
            maxID = strID
        End If
        closeRS
    End If
End Function
