Private Sub CommandButton1_Click()
    Hoja3.Activate
End Sub

Public Sub PlantillaAprobacion()
    Hoja6.Unprotect pword
    Hoja6.Range("APROB_SOL") = ""
    Hoja6.Range("APROB_RENOV") = ""
    Hoja6.Range("APROB_NOMBRE") = ""
    Hoja6.Range("APROB_CSOCIO") = ""
    Hoja6.Range("APROB_PREST") = ""
    Hoja6.Range("APROB_MON") = ""
    Hoja6.Range("APROB_MONTO") = ""
    Hoja6.Range("APROB_PLAZO") = ""
    Hoja6.Range("APROB_TEA") = ""
    Hoja6.Range("APROB_GARANTIA") = ""
    Hoja6.Range("APROB_NPLAZO") = ""
    Hoja6.Range("APROB_PLAZOT") = ""
    Hoja6.Protect Password:=pword, DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Public Sub LlenarAprobacion()
    If Hoja3.tbNomSocio.Text <> "" And Hoja3.tbNSol.Text <> "" And Hoja3.tbCodP.Text <> "" And Hoja3.tbFechaO.Text <> "" And Hoja3.tbMonto.Text <> "" And Hoja3.tbTEA.Text <> "" And Hoja3.Range("PLAZO").Value <> "" And Hoja3.Range("PLAZO_NUEVO").Value <> "" And Hoja3.tbCSocio.Text <> "" And Hoja3.tbNRenov.Text <> "" And Hoja3.tbGarantias.Text <> "" Then
    If IsDate(Hoja3.tbFechaO.Text) Then
    If IsNumeric(Hoja3.tbMonto.Text) Then
    If IsNumeric(Hoja3.tbTEA.Text) Then
    If IsNumeric(Hoja3.Range("PLAZO").Value) Then
    If IsNumeric(Hoja3.Range("PLAZO_NUEVO").Value) Then
    If IsNumeric(Hoja3.tbCSocio.Text) Then
    If IsNumeric(Hoja3.tbNRenov.Text) Then
    Hoja6.Unprotect pword
    Hoja6.Range("APROB_SOL") = Hoja3.tbNSol.Text
    If Len(Hoja3.tbNRenov.Text) <= 2 Then
        Hoja6.Range("APROB_RENOV") = Left("00", 2 - Len(Hoja3.tbNRenov.Text)) & Hoja3.tbNRenov.Text
    Else
        Hoja6.Range("APROB_RENOV") = Hoja3.tbNRenov.Text
    End If
    
    Hoja6.Range("APROB_NOMBRE") = Hoja3.tbNomSocio.Text
    
    If Len(Hoja3.tbCSocio.Text) <= 7 Then
        Hoja6.Range("APROB_CSOCIO") = Left("0000000", 7 - Len(Hoja3.tbCSocio.Text)) & Hoja3.tbCSocio.Text
    Else
        Hoja6.Range("APROB_CSOCIO") = Hoja3.tbCSocio.Text
    End If
    
    Hoja6.Range("APROB_PREST") = Hoja3.tbCodP.Text
    
    If Hoja3.obSoles Then
        Hoja6.Range("APROB_MON") = "SOLES"
        Hoja6.Range("APROB_MONTO") = "S/." & Right(Format(Hoja3.tbMonto.Value, "Currency"), Len(Format(Hoja3.tbMonto.Value, "Currency")) - 3)
    Else
        Hoja6.Range("APROB_MON") = "DOLARES"
        Hoja6.Range("APROB_MONTO") = "US$" & Right(Format(Hoja3.tbMonto.Value, "Currency"), Len(Format(Hoja3.tbMonto.Value, "Currency")) - 3)
    End If
    
    Hoja6.Range("APROB_PLAZO") = Hoja3.Range("PLAZO").Value & " d僘s"
    
    Hoja6.Range("APROB_TEA") = Hoja3.tbTEA.Text & "%"
    
    Hoja6.Range("APROB_GARANTIA") = Hoja3.tbGarantias.Text
    
    Hoja6.Range("APROB_NPLAZO") = Int(Hoja3.Range("PLAZO_NUEVO").Value) - Int(Hoja3.Range("PLAZO").Value) & " d僘s"
    
    Hoja6.Range("APROB_PLAZOT") = Int(Hoja3.Range("PLAZO_NUEVO").Value) & " d僘s"
        closeRS
        OpenDB
        strSQL = "INSERT INTO [HISTORICO$] VALUES ('" & maxID("HISTORICO", "ID") & "','APROBACION'"
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
        strSQL = strSQL & "'" & Hoja3.tbTEA.Text & "','" & Hoja3.Range("FECHA_DESEMBOLSO") & "','" & Hoja3.Range("VENCIMIENTO") & "','" & Hoja3.Range("VENCIMIENTO_TARJETA") & "','" & Hoja3.tbDNIS.Text & "','" & Hoja3.tbDirS.Text & "','" & Hoja3.tbNombreC.Text & "','" & Hoja3.tbDNIC.Text & "','" & Hoja3.tbDirC.Text & "','" & Hoja3.tbNRenov.Text & "','" & Hoja3.tbGarantias.Text & "'"
        If Hoja3.obS Then
            strSQL = strSQL & ",'SOLTERO')"
        ElseIf Hoja3.obC Then
            strSQL = strSQL & ",'CASADO')"
        Else
            strSQL = strSQL & ",'EMPRESA')"
        End If
        rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
        
        closeRS
        Hoja6.Protect Password:=pword, DrawingObjects:=True, Contents:=True, Scenarios:=True
        Hoja6.Activate
    Else
        MsgBox "Error en Numero de Renovaci"
        Hoja3.tbNRenov.Activate
    End If
    Else
        MsgBox "Error en Cigo de Socio"
        Hoja3.tbCSocio.Activate
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
        MsgBox "Error en Fecha Otorgamiento"
        Hoja3.tbFechaO.Activate
        Hoja3.tbFechaO.SelStart = 0
        Hoja3.tbFechaO.SelLength = Len(Hoja3.tbFechaO.Text)
    End If
    Else
        MsgBox "Datos Insuficientes"
    End If
End Sub

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

