Public Sub btLimpiar_Click()
    Hoja1.Unprotect pword
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
    Hoja5.Range(Hoja5.Range("dataSet"), Hoja5.Range("dataSet").End(xlDown)).ClearContents
    Hoja1.PlantillaAdenda
    Hoja2.PlantillaAprobacion
    Hoja6.PlantillaAprobacion
    Hoja7.PlantillaAprobacion
    Hoja3.Protect Password:=pword, DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Private Sub btnAdenda_Click()
    Hoja1.LlenarAdenda
End Sub

Private Sub btnAprobacion_Click()
    If Hoja3.obFuncionario Then
        Hoja2.LlenarAprobacion
    End If
    If Hoja3.obConsejo Then
        Hoja6.LlenarAprobacion
    End If
    If Hoja3.obComite Then
        Hoja7.LlenarAprobacion
    End If
End Sub

Private Sub cbPlastico_Click()
    If cbPlastico Then
        Hoja3.Unprotect pword
        Hoja3.Range("VENCIMIENTO_TARJETA").Interior.ThemeColor = xlThemeColorAccent6
        Hoja3.Range("VENCIMIENTO_TARJETA").Interior.TintAndShade = 0.399975585192419
        Hoja3.Range("VENCIMIENTO_TARJETA").Offset(-1, 0).Interior.ColorIndex = 37
        Hoja3.Range("VENCIMIENTO_TARJETA").Value = ""
        Hoja3.Range("VENCIMIENTO_TARJETA").Locked = False
        Hoja3.Range("NUEVO_PLAZO_DIAS").Interior.ThemeColor = xlThemeColorDark1
        Hoja3.Range("NUEVO_PLAZO_DIAS").Interior.TintAndShade = -0.149998474074526
        Hoja3.Range("NUEVO_PLAZO_DIAS").FormulaR1C1 = "=VENCIMIENTO_TARJETA-VENCIMIENTO"
        Hoja3.Range("NUEVO_PLAZO_DIAS").Locked = True
        Hoja3.Protect Password:=pword, DrawingObjects:=True, Contents:=True, Scenarios:=True
    Else
        Hoja3.Unprotect pword
        Hoja3.Range("NUEVO_PLAZO_DIAS").Interior.ThemeColor = xlThemeColorAccent6
        Hoja3.Range("NUEVO_PLAZO_DIAS").Interior.TintAndShade = 0.399975585192419
        Hoja3.Range("NUEVO_PLAZO_DIAS").Offset(-1, 0).Interior.ColorIndex = 37
        Hoja3.Range("NUEVO_PLAZO_DIAS").Value = ""
        Hoja3.Range("NUEVO_PLAZO_DIAS").Locked = False
        Hoja3.Range("VENCIMIENTO_TARJETA").Interior.ThemeColor = xlThemeColorDark1
        Hoja3.Range("VENCIMIENTO_TARJETA").Interior.TintAndShade = -0.149998474074526
        Hoja3.Range("VENCIMIENTO_TARJETA").FormulaR1C1 = "=VENCIMIENTO+NUEVO_PLAZO_DIAS"
        Hoja3.Range("VENCIMIENTO_TARJETA").Locked = True
        Hoja3.Protect Password:=pword, DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If
End Sub

Private Sub obHistorico_Click()
    busqHistorico.Show
End Sub

Private Sub obS_Change()
    If obS Then
        Me.tbNombreC.Enabled = False
        Me.tbDNIC.Enabled = False
        Me.tbDirC.Enabled = False
        Me.tbNombreC.BackColor = &H80000000
        Me.tbDNIC.BackColor = &H80000000
        Me.tbDirC.BackColor = &H80000000
        Me.tbNombreC.Text = ""
        Me.tbDNIC.Text = ""
        Me.tbDirC.Text = ""
    Else
        Me.tbNombreC.Enabled = True
        Me.tbDNIC.Enabled = True
        Me.tbDirC.Enabled = True
        Me.tbNombreC.BackColor = &H80000005
        Me.tbDNIC.BackColor = &H80000005
        Me.tbDirC.BackColor = &H80000005
    End If
End Sub

Private Sub tbFechaO_LostFocus()
    If IsDate(tbFechaO.Value) Then
        tbFechaO.Value = CDate(tbFechaO.Value)
    Else
        MsgBox "Introduzca una fecha v疝ida"
    End If
End Sub

Private Sub tbMonto_LostFocus()
    If tbMonto.Text <> "" And IsNumeric(tbMonto.Text) Then
        tbMonto.Text = FormatNumber(tbMonto.Text, 2)
    End If
End Sub

Private Sub tbMontoDisp_LostFocus()
    If tbMontoDisp.Text <> "" And IsNumeric(tbMontoDisp.Text) Then
        tbMontoDisp.Text = FormatNumber(tbMontoDisp.Text, 2)
    End If
End Sub

Private Sub tbNomSocio_Change()
    tbNomSocio.Text = UCase(tbNomSocio.Text)
End Sub
