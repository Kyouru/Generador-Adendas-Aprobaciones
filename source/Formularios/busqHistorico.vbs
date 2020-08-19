Private Sub btBuscar_Click()
    ActualizarHoja
    ActualizarLista
End Sub

Private Sub btCargar_Click()
    If ListBox1.ListIndex <> -1 Then
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
        closeRS
        OpenDB
        Hoja3.Unprotect pword
        strSQL = "SELECT * FROM [HISTORICO$] WHERE [ID] = '" & ListBox1.List(ListBox1.ListIndex) & "'"
        rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            If rs.Fields(2).Value = "FUNCIONARIO" Then
                Hoja3.obFuncionario.Value = True
            Else
                If rs.Fields(2).Value = "COMITE" Then
                    Hoja3.obComite.Value = True
                Else
                    Hoja3.obConsejo.Value = True
                End If
            End If
            If Not IsNull(rs.Fields(4).Value) Then
                Hoja3.tbCSocio.Text = rs.Fields(4).Value
            Else
                Hoja3.tbCSocio.Text = ""
            End If
            Hoja3.tbNomSocio.Text = rs.Fields(5).Value
            Hoja3.tbNSol.Text = rs.Fields(6).Value
            Hoja3.tbCodP.Text = rs.Fields(7).Value
            Hoja3.tbFechaO.Text = Format(rs.Fields(8).Value, "DD") & "/" & Format(rs.Fields(8).Value, "MM") & "/" & Format(rs.Fields(8).Value, "YYYY")
            If Not IsNull(rs.Fields(9).Value) Then
                Hoja3.Range("FECHA_CONTRATO").Value = Format(rs.Fields(9).Value, "YYYY-MM-DD")
            Else
                Hoja3.Range("FECHA_CONTRATO").Value = ""
            End If
            Hoja3.tbMonto.Text = rs.Fields(10).Value
            If rs.Fields(11).Value = "S" Then
                Hoja3.obSoles.Value = True
                Hoja3.obDolares.Value = False
            Else
                Hoja3.obSoles.Value = False
                Hoja3.obDolares.Value = True
            End If
            Hoja3.tbTEA.Text = rs.Fields(12).Value
            Hoja3.Range("FECHA_DESEMBOLSO") = Format(rs.Fields(13).Value, "YYYY-MM-DD")
            Hoja3.Range("VENCIMIENTO") = Format(rs.Fields(14).Value, "YYYY-MM-DD")
            Hoja3.Range("VENCIMIENTO_TARJETA") = Format(rs.Fields(15).Value, "YYYY-MM-DD")
            Hoja3.Range("NUEVO_PLAZO_DIAS") = DateDiff("d", Hoja3.Range("VENCIMIENTO"), Hoja3.Range("VENCIMIENTO_TARJETA"))
            
            If Not IsNull(rs.Fields(16).Value) Then
                Hoja3.tbDNIS.Text = rs.Fields(16).Value
            Else
                Hoja3.tbDNIS.Text = ""
            End If
            If Not IsNull(rs.Fields(17).Value) Then
                Hoja3.tbDirS.Text = rs.Fields(17).Value
            Else
                Hoja3.tbDirS.Text = ""
            End If
            If Not IsNull(rs.Fields(18).Value) Then
                Hoja3.tbNombreC.Text = rs.Fields(18).Value
            Else
                Hoja3.tbNombreC.Text = ""
            End If
            If Not IsNull(rs.Fields(19).Value) Then
                Hoja3.tbDNIC.Text = rs.Fields(19).Value
            Else
                Hoja3.tbDNIC.Text = ""
            End If
            If Not IsNull(rs.Fields(20).Value) Then
                Hoja3.tbDirC.Text = rs.Fields(20).Value
            Else
                Hoja3.tbDirC.Text = ""
            End If
            If Not IsNull(rs.Fields(21).Value) Then
                Hoja3.tbNRenov.Text = rs.Fields(21).Value
            Else
                Hoja3.tbNRenov.Text = ""
            End If
            If Not IsNull(rs.Fields(22).Value) Then
                Hoja3.tbGarantias.Text = rs.Fields(22).Value
            Else
                Hoja3.tbGarantias.Text = ""
            End If
            If rs.Fields(23).Value = "SOLTERO" Then
                Hoja3.obS.Value = True
                Hoja3.obC.Value = False
                Hoja3.obE.Value = False
            ElseIf rs.Fields(23).Value = "CASADO" Then
                Hoja3.obS.Value = False
                Hoja3.obC.Value = True
                Hoja3.obE.Value = False
            Else
                Hoja3.obS.Value = False
                Hoja3.obC.Value = False
                Hoja3.obE.Value = True
            End If
        End If
        Hoja3.Protect Password:=pword, DrawingObjects:=True, Contents:=True, Scenarios:=True
        closeRS
        Unload Me
        Else
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
End Sub

Private Sub btEliminar_Click()
    Dim i As Integer
    Dim rDelete As Range
    i = 0
    If ListBox1.ListIndex <> -1 Then
        For i = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(i) Then
                Set rDelete = Hoja4.Range("A:A").Find(What:=ListBox1.List(i), LookIn:=xlValues)
                Hoja4.Rows(rDelete.Row).Delete
            End If
        Next i
        btBuscar_Click
    Else
        MsgBox "Seleccione una entrada"
    End If
End Sub

Private Sub btSeleccionarTodo_Click()
    Dim i As Integer
    i = 0
    For i = 0 To ListBox1.ListCount - 1
        ListBox1.Selected(i) = True
    Next i
End Sub

Private Sub cbAdenda_Change()
    ActualizarHoja
    ActualizarLista
End Sub

Private Sub cbAprobacion_Change()
    ActualizarHoja
    ActualizarLista
End Sub


Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btCargar_Click
End Sub

Private Sub tbCodSocio_Change()
    ActualizarHoja
    ActualizarLista
End Sub

Private Sub tbNomSocio_Change()
    tbNomSocio.Text = UCase(tbNomSocio.Text)
    ActualizarHoja
    ActualizarLista
End Sub

Private Sub tbRenovacion_Change()
    ActualizarHoja
    ActualizarLista
End Sub

Private Sub tbSolicitud_Change()
    ActualizarHoja
    ActualizarLista
End Sub

Private Sub UserForm_Initialize()
    ActualizarHoja
    ActualizarLista
End Sub

Private Sub ActualizarHoja()
    closeRS
    OpenDB
    strSQL = "SELECT * FROM [HISTORICO$] WHERE 1=1"
    If Not cbAdenda Then
        If Not cbAprobacion Then
            strSQL = strSQL & " AND [TIPO] <> 'ADENDA' AND [TIPO] <> 'APROBACION' "
        Else
            strSQL = strSQL & " AND [TIPO] <> 'ADENDA' "
        End If
    Else
        If Not cbAprobacion Then
            strSQL = strSQL & " AND [TIPO] <> 'APROBACION' "
        End If
    End If
    If tbCodSocio.Text <> "" Then
        strSQL = strSQL & " AND [SOCIO] LIKE '%" & tbCodSocio.Text & "%'"
    End If
    If tbNomSocio.Text <> "" Then
        strSQL = strSQL & " AND [NOMBRE] LIKE '%" & tbNomSocio.Text & "%'"
    End If
    If tbSolicitud.Text <> "" Then
        strSQL = strSQL & " AND [SOLICITUD] LIKE '%" & tbSolicitud.Text & "%'"
    End If
    If tbRenovacion.Text <> "" Then
        strSQL = strSQL & " AND [NRENOVACION] LIKE '%" & tbRenovacion.Text & "%'"
    End If
    Hoja5.Range(Hoja5.Range("dataSet"), Hoja5.Range("dataSet").End(xlDown)).ClearContents
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        Hoja5.Range("dataSet").CopyFromRecordset rs
    End If
    closeRS
End Sub


Private Sub ActualizarLista()

    With ListBox1
        .ColumnWidths = "0;65;55;80;40;180;90;120;100;120;200;200;200;200;200;200;200;200;200"
        .ColumnCount = 19
        .ColumnHeads = True
        .RowSource = "HISTORICO_TEMP!A2:R" & Sheets("HISTORICO_TEMP").Range("A1").End(xlDown).Row
    End With
    
End Sub
