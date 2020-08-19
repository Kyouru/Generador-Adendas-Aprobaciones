Private Sub Workbook_Open()

    'Mostrar Hojas tras habilitar Macros
    ThisWorkbook.Sheets("ADENDA").Visible = xlSheetVisible
    ThisWorkbook.Sheets("APROB FUNCIONARIO").Visible = xlSheetVisible
    ThisWorkbook.Sheets("INICIO").Visible = xlSheetVisible
    ThisWorkbook.Sheets("APROB CONSEJO").Visible = xlSheetVisible
    ThisWorkbook.Sheets("APROB COMITE").Visible = xlSheetVisible
    
    'Ocultar Hoja de Introduccion
    ThisWorkbook.Sheets("CHECK").Visible = xlSheetHidden
    If ThisWorkbook.ReadOnly Then
        MsgBox "Excel en Modo Solo Lectura" & vbNewLine & "Descarge el Excel en una ubicaci con permisos de Escritura" & vbNewLine & vbNewLine & "Cerrando..."
        ThisWorkbook.Close SaveChanges:=False
    Else
        Dim answer As Integer
        answer = MsgBox("Deseas Limpiar el Formulario?", vbYesNo + vbQuestion, "Limpiar")
        If answer = vbYes Then
            Hoja3.tbNomSocio.Text = ""
            Hoja3.tbNSol.Text = ""
            Hoja3.tbCodP.Text = ""
            Hoja3.tbFechaO.Text = ""
            Hoja3.tbMonto.Text = ""
            Hoja3.obSoles.Value = True
            Hoja3.obDolares.Value = False
            Hoja3.tbTEA.Text = ""
            Hoja3.tbDNIS.Text = ""
            Hoja3.tbDirS.Text = ""
            Hoja3.obS.Value = True
            Hoja3.tbCSocio.Text = ""
            Hoja3.tbNRenov.Text = ""
            Hoja3.tbGarantias.Text = ""
            Hoja3.Range("CALC_FECHAS") = ""
            Hoja3.Range("FECHA_CONTRATO").Value = ""
            Hoja3.Range("FECHA_DESEMBOLSO").Value = ""
            Hoja3.Range("VENCIMIENTO").Value = ""
            If cbPlastico Then
                Hoja3.Range("VENCIMIENTO_TARJETA").Value = ""
            Else
                Hoja3.Range("NUEVO_PLAZO_DIAS").Value = ""
            End If
        End If
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'Mostrar Hoja de Introduccion
    ThisWorkbook.Sheets("CHECK").Visible = xlSheetVisible
    ThisWorkbook.Sheets("CHECK").Activate

    'Mostrar Hojas tras habilitar Macros
    ThisWorkbook.Sheets("ADENDA").Visible = xlSheetHidden
    ThisWorkbook.Sheets("APROB FUNCIONARIO").Visible = xlSheetHidden
    ThisWorkbook.Sheets("INICIO").Visible = xlSheetHidden
    ThisWorkbook.Sheets("HISTORICO").Visible = xlSheetHidden
    ThisWorkbook.Sheets("HISTORICO_TEMP").Visible = xlSheetHidden
    ThisWorkbook.Sheets("APROB CONSEJO").Visible = xlSheetHidden
    ThisWorkbook.Sheets("APROB COMITE").Visible = xlSheetHidden
    
    
    ActiveWorkbook.Save
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    
End Sub
