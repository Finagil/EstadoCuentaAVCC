Imports System.IO
Imports System.Net.Mail

Public Class frmEstadoCuentaAvio
    Dim FechaQuitaFega As Date = Date.ParseExact("26/09/2016", "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture) '***** SE QUITA FEGA Y GL DEL SEGURO DE VIDA A PARTIR DEL MES DE SEPTIEMBRE, ELISANDER PINEDA #ect 26092015.N
    Dim UltimoCorte As String
    Dim UltimoPAGO As String = "20010101"
    Dim UltimoPAGO_CON As String = "20010101"
    Dim UltimoCORTE_CON As String = "20010101"
    Dim UltimoPAGO_GEN As String = "20010101"
    Dim FechaVen As Date
    Dim FecTope As Date
    Dim Tasa As Double
    Dim TasaAUX As Double
    Dim TasaSegVid As Decimal
    Dim Tipar As String
    Dim Fondeo As String
    Dim AplicaGarantiaLIQ As String
    Dim arg() As String
    Dim ArreSegVid(50, 3) As String
    Dim DatFact(3) As String
    Dim SegVid As Double
    Dim FecVid As Date
    Dim CalcSEGVID As Boolean = True
    Dim CalcSEGVIDcad As String = "SEGURO DE VIDA"
    Dim SumaIni As Double = 0
    Dim SumaFin As Double = 0
    Dim Usuario As String
    Dim SinMoratorios As String
    Dim InteresMensual As String
    Dim FechaAutorizacion As String
    Dim PorcFega As Decimal = 0
    Dim DiasMenos As Integer = -3
    Dim TercioDeTasa As Boolean = False
    Dim AHORA As Date
    Dim HastaVENC As Boolean = False
    Dim FECHA_APLICACION As Date
    Dim AplicaFega As Boolean
    Dim FegaFLAT As Boolean
    Private Sub ButtonCargar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCargar.Click
        Me.DetalleFINAGILTableAdapter.QuitaNulosFactura()
        Me.DetalleFINAGILTableAdapter.NoFacturado()
        Me.DetalleFINAGILTableAdapter.FillByAnexo(Me.ProductionDataSet.DetalleFINAGIL, txtanexo.Text, txtCiclo.Text)
        If Me.ProductionDataSet.DetalleFINAGIL.Rows.Count > 0 Then
            Me.ClientesTableAdapter.Fill(Me.ProductionDataSet.Clientes, LabelCliente.Text)
            LBtotal.Text = Format(Me.DetalleFINAGILTableAdapter.Total(txtanexo.Text, txtCiclo.Text), "N")
            Me.DetalleFINAGILTableAdapter.UpdateUltimoSaldo(LBtotal.Text, txtanexo.Text, txtCiclo.Text)
            Me.DetalleFINAGILTableAdapter1.FillByAnexo(Me.ProductionDataSet1.DetalleFINAGIL, "", "")
            ButtonReCalc.Enabled = True
            ButtonCargar.Enabled = False
            txtanexo.Enabled = False
            txtCiclo.Enabled = False
        Else
            If arg(3) = "FIN" Or UCase(arg(3)) = "ECT" Then
                Console.WriteLine("No existen ministraciones de este contrato")
            Else
                MessageBox.Show("No existen ministraciones de este contrato", "Error de Contrato", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        End If
    End Sub

    Private Sub ButtonRecalc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonReCalc.Click
        ReDim ArreSegVid(50, 3)
        If Val(LBtotal.Text) > 0 Then
            'Console.WriteLine(LBtotal.Text)
            TasaSegVid = SeguroVida(txtanexo.Text, txtCiclo.Text)
            If TasaSegVid <= 0 Then
                CalcSEGVID = False
                CalcSEGVIDcad = "NO SEGURO DE VIDA"
            Else
                CalcSEGVID = True
                CalcSEGVIDcad = "SEGURO DE VIDA"
            End If
            '''TasaSegVid = 0 'para no conbrar seguro de vida
            'Console.WriteLine("recalc2")
            FijaTasa(txtanexo.Text, txtCiclo.Text, AHORA.AddDays(DiasMenos))
            If FechaVen > AHORA.AddDays(DiasMenos) And CheckProyectado.Checked = False Then
                Recalcular2()
            Else
                If CheckProyectado.Checked = False Then
                    If FechaVen >= CadenaFecha(AHORA.AddDays(DiasMenos).ToString("yyyyMM01")).AddMonths(-1) And FechaVen >= CadenaFecha(UltimoPAGO_CON) Then
                        Dim diasMEnosAux As Integer = DiasMenos
                        Dim AhoraAux As Date = AHORA
                        AHORA = FechaVen
                        DiasMenos = 0
                        HastaVENC = True
                        Recalcular2()
                        DiasMenos = diasMEnosAux
                        AHORA = AhoraAux
                        HastaVENC = False
                    End If
                Else
                    If FechaVen <= AHORA And FechaVen >= CadenaFecha(UltimoPAGO_CON) Then
                        Dim diasMEnosAux As Integer = DiasMenos
                        Dim AhoraAux As Date = AHORA
                        AHORA = FechaVen
                        DiasMenos = 0
                        HastaVENC = True
                        Recalcular2()
                        DiasMenos = diasMEnosAux
                        AHORA = AhoraAux
                        HastaVENC = False
                    Else
                        CheckProyectado.Checked = False
                        Recalcular2()
                    End If

                End If

            End If

            'Console.WriteLine("recalc3")
            ButtonReCalc.Enabled = False
            ButtonCargar.Enabled = False
            ButtonSave.Enabled = True
            If UCase(arg(3)) = "ECT" Then
                ButtonSave_Click(Nothing, Nothing)
            End If

        Else
            If arg(3) = "FIN" Then
            Else
                MessageBox.Show("Credito Liquidado " & txtanexo.Text & "/" & txtCiclo.Text, "Contratos", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        End If

    End Sub

    Sub Recalcular2()
        Me.ProductionDataSet1.DetalleFINAGIL.Clear()
        Dim r As Estado_de_Cuenta.ProductionDataSet.DetalleFINAGILRow
        Dim rr As Estado_de_Cuenta.ProductionDataSet.DetalleFINAGILRow
        'Dim rri As Estado_de_Cuenta.ProductionDataSet.DetalleFINAGILRow
        Dim uno As Boolean = True
        Dim dias As Integer = 0
        Dim Saldofin As Double = 0
        Dim Saldoini As Double = 0
        Dim Intereses As Double = 0
        Dim Consec As Integer = 0
        Dim fechaINI As Date
        Dim Fecha As Date
        Dim FechaCorte As Date
        Dim FechaANT As Date
        Dim con As String
        Dim Ta As New Estado_de_Cuenta.ProductionDataSet.DetalleFINAGILDataTable
        'Dim fechaFIN As Date
        Me.DetalleFINAGILTableAdapter.FillByAnexoFecha(Ta, txtanexo.Text, txtCiclo.Text)
        SumaIni = DetalleFINAGILTableAdapter.SumaCargosPag(txtanexo.Text, txtCiclo.Text)
        If Ta.Rows.Count > 0 Then
            'Console.WriteLine("recalc21")
            For Each r In Ta.Rows
                con = Trim(r("Concepto"))
                'Console.WriteLine(con)
                If con <> "INTERESES" And con <> CalcSEGVIDcad Then
                    If con = "PAGO" Then UltimoPAGO = r("FechaFinal")
                    If Mid(con, 1, 2) = "NC" Then
                        UltimoPAGO = r("FechaFinal")
                    End If
                    fechaINI = CadenaFecha(r("fechaFinal"))
                    If uno = False Then
                        If fechaINI <> Fecha Then
                            If fechaINI.Month = Fecha.Month Then
                                dias = DateDiff(DateInterval.Day, Fecha, fechaINI)
                                If dias > 0 Then
                                    GeneraInteresesII(Fecha.ToString("yyyyMMdd"), Consec, r, Saldofin, 0, fechaINI.ToString("yyyyMMdd"))
                                End If
                                'LINEA(NORMAL)
                                r.Consecutivo = Consec * 101
                                Me.ProductionDataSet1.DetalleFINAGIL.ImportRow(r)
                                rr = Me.ProductionDataSet1.DetalleFINAGIL.Rows(Consec)
                                Consec += 1
                                rr("Consecutivo") = Consec
                                rr("Factura") = r.Factura
                                rr("FolioFiscal") = r.FolioFiscal
                                rr("Facturado") = r.Facturado
                                rr("dias") = 0
                                rr("fechaInicial") = r("FechaFinal")
                                Saldoini = Saldofin
                                If con <> "PAGO" And con <> "BONIFICACIÓN GARANTÍA LÍQUIDA" And Mid(con, 1, 2) <> "NC" Then
                                    rr("garantia") = 0
                                    rr("fega") = 0
                                End If

                                FijaTasa(r("anexo"), r("ciclo"), fechaINI.AddMonths(-1))
                                If FechaVen < fechaINI And Tipar = "C" Then
                                    Tasa = Tasa * 2
                                ElseIf FechaVen < fechaINI And (Tipar = "H" Or Tipar = "A") Then
                                    Tasa = Tasa * 3
                                End If

                                rr("tasabp") = Tasa
                                If Tipar = "C" And Fondeo = "03" And con <> "PAGO" And Mid(con, 1, 2) <> "NC" Then
                                    If FechaAutorizacion >= "20160101" Then
                                        rr("fega") = r("importe") * 0.0174
                                        If PorcFega > 0 Then
                                            rr("fega") = r("importe") * PorcFega
                                        End If
                                    Else
                                        rr("fega") = r("importe") * 0.0116
                                    End If

                                    If r("Anexo") = "030140002" And r("ciclo") = "02" And Consec = 1 Then
                                        rr("fega") = 43913.27
                                    End If
                                    If r("Anexo") = "030140002" And r("ciclo") = "03" And Consec = 1 Then
                                        rr("fega") = 33694.46
                                    End If
                                    rr("garantia") = 0
                                End If
                                If r("Anexo") = "030140003" And r("ciclo") = "01" And Consec = 1 Then rr("fega") = 58000
                                If r("Anexo") = "030140003" And r("ciclo") = "01" And Consec = 19 Then rr("fega") = 14210
                                If r("Anexo") = "043980001" And r("ciclo") = "01" And Consec = 1 Then rr("fega") = 454.14
                                If r("Anexo") = "043980001" And r("ciclo") = "01" And Consec = 3 Then rr("fega") = 25360.91

                                If Tipar = "H" And Fondeo = "03" And con <> "PAGO" And con <> "BONIFICACIÓN GARANTÍA LÍQUIDA" And Mid(con, 1, 2) <> "NC" Then
                                    If FechaAutorizacion >= "20160101" Then
                                        rr("fega") = r("importe") * 0.0174
                                        If PorcFega > 0 Then
                                            rr("fega") = r("importe") * PorcFega
                                        End If
                                    Else
                                        rr("fega") = r("importe") * 0.0116
                                    End If

                                    rr("garantia") = r("importe") * 0.1

                                    If UCase(AplicaGarantiaLIQ) = "NO" Then
                                        rr("garantia") = 0
                                    End If
                                End If
                                If Fondeo = "03" And con <> "PAGO" Then
                                    If AplicaFega = False Then
                                        rr("fega") = 0
                                    Else
                                        If FegaFLAT = 0 Then
                                            dias = DateDiff("d", CadenaFecha(r("FechaFinal")), FechaVen)
                                            rr("fega") = Math.Round(r("importe") * (0.0174 / 360) * dias, 2)
                                        End If
                                    End If
                                End If




                                Saldofin = rr("importe") + rr("fega") + rr("garantia") + rr("intereses") + Saldoini
                                rr("SALDOFinal") = Saldofin
                                rr("SALDOinicial") = Saldoini
                                UltimoCorte = rr("fechaFinal")
                            Else
                                FechaCorte = CadenaFecha(Fecha.ToString("yyyyMM01"))
                                FechaCorte = FechaCorte.AddMonths(1)
                                FechaCorte = FechaCorte.AddDays(-1)
                                dias = DateDiff(DateInterval.Day, Fecha, FechaCorte)
                                FechaCorte = Fecha
                                Fecha = fechaINI
                                'genera intereses hasta el corte
                                'Console.WriteLine("GeneraInteresesII1")
                                GeneraInteresesII(FechaCorte.ToString("yyyyMMdd"), Consec, r, Saldofin, 0, Fecha.ToString("yyyyMMdd"))
                                'Console.WriteLine("GeneraInteresesII2")
                                'LINEA(NORMAL)
                                r.Consecutivo = Consec * 102
                                Me.ProductionDataSet1.DetalleFINAGIL.ImportRow(r)
                                rr = Me.ProductionDataSet1.DetalleFINAGIL.Rows(Consec)
                                Consec += 1
                                rr("Consecutivo") = Consec
                                rr("Factura") = r.Factura
                                rr("FolioFiscal") = r.FolioFiscal
                                rr("Facturado") = r.Facturado
                                rr("dias") = 0
                                rr("fechaInicial") = r("FechaFinal")
                                Saldoini = Saldofin

                                If con <> "PAGO" And con <> "BONIFICACIÓN GARANTÍA LÍQUIDA" And Mid(con, 1, 2) <> "NC" Then
                                    rr("garantia") = 0
                                    rr("fega") = 0
                                End If
                                FijaTasa(r("anexo"), r("ciclo"), fechaINI.AddMonths(-1))
                                If FechaVen < fechaINI And Tipar = "C" Then
                                    Tasa = Tasa * 2
                                ElseIf FechaVen < fechaINI And (Tipar = "H" Or Tipar = "A") Then
                                    Tasa = Tasa * 3
                                End If
                                rr("tasabp") = Tasa

                                If Tipar = "C" And Fondeo = "03" And con <> "PAGO" And Mid(con, 1, 2) <> "NC" Then
                                    If FechaAutorizacion >= "20160101" Then
                                        rr("fega") = r("importe") * 0.0174
                                        If PorcFega > 0 Then
                                            rr("fega") = r("importe") * PorcFega
                                        End If
                                    Else
                                        rr("fega") = r("importe") * 0.0116
                                    End If
                                    If r("Anexo") = "030140002" And r("ciclo") = "02" And Consec = 1 Then
                                        rr("fega") = 43913.27
                                    End If
                                    If r("Anexo") = "030140002" And r("ciclo") = "03" And Consec = 1 Then
                                        rr("fega") = 33694.46
                                    End If
                                    rr("garantia") = 0
                                End If
                                If r("Anexo") = "030140003" And r("ciclo") = "01" And Consec = 1 Then rr("fega") = 58000
                                If r("Anexo") = "030140003" And r("ciclo") = "01" And Consec = 19 Then rr("fega") = 14210
                                If r("Anexo") = "043980001" And r("ciclo") = "01" And Consec = 1 Then rr("fega") = 454.14
                                If r("Anexo") = "043980001" And r("ciclo") = "01" And Consec = 3 Then rr("fega") = 25360.91
                                If Tipar = "H" And Fondeo = "03" And con <> "PAGO" And con <> "BONIFICACIÓN GARANTÍA LÍQUIDA" And Mid(con, 1, 2) <> "NC" Then
                                    If FechaAutorizacion >= "20160101" Then
                                        rr("fega") = r("importe") * 0.0174
                                        If PorcFega > 0 Then
                                            rr("fega") = r("importe") * PorcFega
                                        End If
                                    Else
                                        rr("fega") = r("importe") * 0.0116
                                    End If
                                    rr("garantia") = r("importe") * 0.1
                                    'If (txtanexo.Text = "030140002" Or txtanexo.Text = "x85140012") And txtCiclo.Text = "10" Then
                                    '    rr("garantia") = 0
                                    'End If
                                    If UCase(AplicaGarantiaLIQ) = "NO" Then
                                        rr("garantia") = 0
                                    End If
                                End If
                                If Fondeo = "03" And con <> "PAGO" Then
                                    If AplicaFega = False Then
                                        rr("fega") = 0
                                    Else
                                        If FegaFLAT = 0 Then
                                            dias = DateDiff("d", CadenaFecha(r("FechaFinal")), FechaVen)
                                            rr("fega") = Math.Round(r("importe") * (0.0174 / 360) * dias, 2)
                                        End If
                                    End If
                                End If


                                Saldofin = rr("importe") + rr("fega") + rr("garantia") + rr("intereses") + Saldoini
                                rr("SALDOFinal") = Saldofin
                                rr("saldoInicial") = Saldoini
                                'Me.ProductionDataSet1.DetalleFINAGIL.ImportRow(rr)
                                UltimoCorte = rr("fechaFinal")

                            End If
                        Else
                            r.Consecutivo = Consec * 100
                            Me.ProductionDataSet1.DetalleFINAGIL.ImportRow(r)
                            rr = Me.ProductionDataSet1.DetalleFINAGIL.Rows(Consec)
                            Saldoini = Saldofin
                            Consec += 1
                            rr("dias") = 0
                            rr("fechaFinal") = fechaINI.ToString("yyyyMMdd")
                            rr("fechaInicial") = fechaINI.ToString("yyyyMMdd")
                            rr("Consecutivo") = Consec
                            rr("Factura") = r.Factura
                            rr("FolioFiscal") = r.FolioFiscal
                            rr("Facturado") = r.Facturado
                            If con <> "PAGO" And con <> "BONIFICACIÓN GARANTÍA LÍQUIDA" And Mid(con, 1, 2) <> "NC" Then
                                rr("garantia") = 0
                                rr("fega") = 0
                            End If

                            FijaTasa(r("anexo"), r("ciclo"), fechaINI.AddMonths(-1))
                            If FechaVen < fechaINI And Tipar = "C" Then
                                Tasa = Tasa * 2
                            ElseIf FechaVen < fechaINI And (Tipar = "H" Or Tipar = "A") Then
                                Tasa = Tasa * 3
                            End If
                            rr("tasabp") = Tasa

                            If Tipar = "C" And Fondeo = "03" And con <> "PAGO" And Mid(con, 1, 2) <> "NC" Then
                                If FechaAutorizacion >= "20160101" Then
                                    rr("fega") = r("importe") * 0.0174
                                    If PorcFega > 0 Then
                                        rr("fega") = r("importe") * PorcFega
                                    End If
                                Else
                                    rr("fega") = r("importe") * 0.0116
                                End If
                                If r("Anexo") = "030140002" And r("ciclo") = "02" And Consec = 1 Then
                                    rr("fega") = 43913.27
                                End If
                                If r("Anexo") = "030140002" And r("ciclo") = "03" And Consec = 1 Then
                                    rr("fega") = 33694.46
                                End If
                                rr("garantia") = 0
                            End If
                            If r("Anexo") = "030140003" And r("ciclo") = "01" And Consec = 1 Then rr("fega") = 58000
                            If r("Anexo") = "030140003" And r("ciclo") = "01" And Consec = 19 Then rr("fega") = 14210
                            If r("Anexo") = "043980001" And r("ciclo") = "01" And Consec = 1 Then rr("fega") = 454.14
                            If r("Anexo") = "043980001" And r("ciclo") = "01" And Consec = 3 Then rr("fega") = 25360.91
                            If Tipar = "H" And Fondeo = "03" And con <> "PAGO" And con <> "BONIFICACIÓN GARANTÍA LÍQUIDA" And Mid(con, 1, 2) <> "NC" Then
                                If FechaAutorizacion >= "20160101" Then
                                    rr("fega") = r("importe") * 0.0174
                                    If PorcFega > 0 Then
                                        rr("fega") = r("importe") * PorcFega
                                    End If
                                Else
                                    rr("fega") = r("importe") * 0.0116
                                End If
                                rr("garantia") = r("importe") * 0.1
                                'If (txtanexo.Text = "030140002" Or txtanexo.Text = "x85140012") And txtCiclo.Text = "10" Then
                                '    rr("garantia") = 0
                                'End If
                                If UCase(AplicaGarantiaLIQ) = "NO" Then
                                    rr("garantia") = 0
                                End If
                            End If
                            If Fondeo = "03" And con <> "PAGO" Then
                                If AplicaFega = False Then
                                    rr("fega") = 0
                                Else
                                    If FegaFLAT = 0 Then
                                        dias = DateDiff("d", CadenaFecha(r("FechaFinal")), FechaVen)
                                        rr("fega") = Math.Round(r("importe") * (0.0174 / 360) * dias, 2)
                                    End If
                                End If
                            End If
                            Saldofin = rr("importe") + rr("fega") + rr("garantia") + r("intereses") + Saldoini
                            rr("SALDOFinal") = Saldofin
                            rr("SALDOinicial") = Saldoini
                            UltimoCorte = rr("fechaFinal")
                        End If
                    Else
                        Me.ProductionDataSet1.DetalleFINAGIL.ImportRow(r)
                        rr = Me.ProductionDataSet1.DetalleFINAGIL.Rows(Consec)
                        Saldoini = Saldofin
                        Consec += 1
                        rr("dias") = 0
                        rr("fechaFinal") = fechaINI.ToString("yyyyMMdd")
                        rr("fechaInicial") = fechaINI.ToString("yyyyMMdd")
                        rr("Consecutivo") = Consec
                        rr("Factura") = r.Factura
                        rr("FolioFiscal") = r.FolioFiscal
                        rr("Facturado") = r.Facturado

                        If con <> "PAGO" And con <> "BONIFICACIÓN GARANTÍA LÍQUIDA" And Mid(con, 1, 2) <> "NC" Then
                            rr("garantia") = 0
                            rr("fega") = 0
                        End If

                        FijaTasa(r("anexo"), r("ciclo"), fechaINI.AddMonths(-1))
                        If FechaVen < fechaINI And Tipar = "C" Then
                            Tasa = Tasa * 2
                        ElseIf FechaVen < fechaINI And (Tipar = "H" Or Tipar = "A") Then
                            Tasa = Tasa * 3
                        End If
                        rr("tasabp") = Tasa

                        If Tipar = "C" And Fondeo = "03" And con <> "PAGO" And Mid(con, 1, 2) <> "NC" Then
                            If FechaAutorizacion >= "20160101" Then
                                rr("fega") = r("importe") * 0.0174
                                If PorcFega > 0 Then
                                    rr("fega") = r("importe") * PorcFega
                                End If
                            Else
                                rr("fega") = r("importe") * 0.0116
                            End If
                            If r("Anexo") = "030140002" And r("ciclo") = "02" And Consec = 1 Then
                                rr("fega") = 43913.27
                            End If
                            If r("Anexo") = "030140002" And r("ciclo") = "03" And Consec = 1 Then
                                rr("fega") = 33694.46
                            End If
                            rr("garantia") = 0
                        End If
                        If r("Anexo") = "030140003" And r("ciclo") = "01" And Consec = 1 Then rr("fega") = 58000
                        If r("Anexo") = "030140003" And r("ciclo") = "01" And Consec = 19 Then rr("fega") = 14210
                        If r("Anexo") = "043980001" And r("ciclo") = "01" And Consec = 1 Then rr("fega") = 454.14
                        If r("Anexo") = "043980001" And r("ciclo") = "01" And Consec = 3 Then rr("fega") = 25360.91
                        If Tipar = "H" And Fondeo = "03" And con <> "PAGO" And con <> "BONIFICACIÓN GARANTÍA LÍQUIDA" And Mid(con, 1, 2) <> "NC" Then
                            If FechaAutorizacion >= "20160101" Then
                                rr("fega") = r("importe") * 0.0174
                                If PorcFega > 0 Then
                                    rr("fega") = r("importe") * PorcFega
                                End If
                            Else
                                rr("fega") = r("importe") * 0.0116
                            End If
                            rr("garantia") = r("importe") * 0.1
                            'If (txtanexo.Text = "030140002" Or txtanexo.Text = "x85140012") And txtCiclo.Text = "10" Then
                            '    rr("garantia") = 0
                            'End If
                            If UCase(AplicaGarantiaLIQ) = "NO" Then
                                rr("garantia") = 0
                            End If
                        End If
                        If Fondeo = "03" And con <> "PAGO" Then
                            If AplicaFega = False Then
                                rr("fega") = 0
                            Else
                                If FegaFLAT = 0 Then
                                    dias = DateDiff("d", CadenaFecha(r("FechaFinal")), FechaVen)
                                    rr("fega") = Math.Round(r("importe") * (0.0174 / 360) * dias, 2)
                                End If
                            End If
                        End If
                        Saldofin = rr("importe") + rr("fega") + rr("garantia") + rr("intereses") + Saldoini
                        rr("SALDOFinal") = Saldofin
                        rr("SALDOinicial") = Saldoini
                        UltimoCorte = rr("fechaFinal")
                    End If
                End If
                Fecha = fechaINI
                FechaANT = Fecha.AddMonths(-1)
                uno = False
            Next
            'Console.WriteLine("recalc22")
            If CheckProyectado.Checked = False Then
                If HastaVENC = True Then
                    FecTope = CadenaFecha(AHORA.AddDays(DiasMenos).ToString("yyyyMMdd"))
                    GeneraIntereses(UltimoCorte, Consec, r, Saldofin, 0, AHORA.AddDays(DiasMenos).ToString("yyyyMMdd"))
                Else
                    FecTope = CadenaFecha(AHORA.AddDays(DiasMenos).ToString("yyyyMM01"))
                    GeneraIntereses(UltimoCorte, Consec, r, Saldofin, 0, AHORA.AddDays(DiasMenos).ToString("yyyyMM01"))
                End If

            Else
                If HastaVENC = True Then
                    FecTope = CadenaFecha(AHORA.AddDays(DiasMenos).ToString("yyyyMMdd"))
                    GeneraIntereses(UltimoCorte, Consec, r, Saldofin, 0, AHORA.AddDays(DiasMenos).ToString("yyyyMMdd"))
                Else
                    FecTope = CadenaFecha(AHORA.AddMonths(1).ToString("yyyyMM01"))
                    GeneraIntereses(UltimoCorte, Consec, r, Saldofin, 0, AHORA.AddMonths(1).ToString("yyyyMM01"))
                End If

            End If
        End If
        'Console.WriteLine("recalc23")
        Me.DetalleFINAGILTableAdapter.FillByAnexo(Me.ProductionDataSet.DetalleFINAGIL, txtanexo.Text, txtCiclo.Text)
        LBtotal2.Text = Format(Saldofin, "N")
        LbTotal3.Text = Format(LBtotal.Text - LBtotal2.Text, "N")

        If Val(LbTotal3.Text) = 0 Then
            LbTotal3.ForeColor = Color.Black
        ElseIf Val(LbTotal3.Text) < 0 Then
            LbTotal3.ForeColor = Color.Red
        Else
            LbTotal3.ForeColor = Color.Blue
        End If

    End Sub

    Function CadenaFecha(ByVal f As String) As Date
        Dim ff As New System.DateTime(CInt(Mid(f, 1, 4)), Mid(f, 5, 2), Mid(f, 7, 2))
        Return ff
    End Function

    Sub FijaTasa(ByVal a As String, ByVal c As String, ByVal F As Date)
        Dim TaAvios As New Estado_de_Cuenta.ProductionDataSetTableAdapters.AviosTableAdapter
        Dim TaTasaMora As New Estado_de_Cuenta.ProductionDataSetTableAdapters.AnexosTasaMoraFecORdTableAdapter
        Dim y As New Estado_de_Cuenta.ProductionDataSet.AviosDataTable
        Dim Diferencial As Double
        TaAvios.UpdateAplicaGAR()
        TaAvios.FillAnexo(y, a, c)
        UltimoPAGO_CON = TaAvios.UltimoPago(a, c)
        UltimoCORTE_CON = TaAvios.UltimaFecha(a, c)
        UltimoPAGO_GEN = TaTasaMora.SacaFecha(a, c)
        If TaTasaMora.TieneTercioDeTasa(a, c, True) > 0 Then
            TercioDeTasa = True
        Else
            TercioDeTasa = False
        End If
        Dim v As String = y.Rows(0).Item(0)
        If v = "7" Then
            Tasa = y.Rows(0).Item(5)
        Else
            Dim TIIE As New Estado_de_Cuenta.ProductionDataSetTableAdapters.TIIEpromedioTableAdapter
            Diferencial = y.Rows(0).Item(1)
            Tasa = TIIE.SacaTIIE(F.ToString("yyyyMM")) + Diferencial
        End If
        'SOLICITADO POR ELISANDER DEBIDO A PRORROGA, SE SUMA UN PUNTO PORCENTUAL A PARTIR DE ENERO++++++++++++
        If (a = "070320012" Or a = "070860007" Or a = "070790010" Or a = "070780012" Or a = "070600011" Or a = "071330006") And F >= CDate("01/01/2015") Then
            Tasa += 1
        End If
        TasaAUX = Tasa
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Fondeo = y.Rows(0).Item("Fondeo")
        Tipar = y.Rows(0).Item("Tipar")
        AplicaGarantiaLIQ = y.Rows(0).Item("AplicaGarantiaLIQ")
        FechaAutorizacion = y.Rows(0).Item("FechaAutorizacion")
        FechaVen = CadenaFecha(y.Rows(0).Item("FechaTerminacion"))
        SinMoratorios = y.Rows(0).Item("SinMoratorios")
        InteresMensual = y.Rows(0).Item("InteresMensual")
        PorcFega = y.Rows(0).Item("PorcFega")
        AplicaFega = y.Rows(0).Item("AplicaFega")
        FegaFLAT = y.Rows(0).Item("FegaFlat")
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Try
        Dim tx As New ProductionDataSetTableAdapters.AviosTableAdapter
        FECHA_APLICACION = tx.FechaAplicacion()
        AHORA = Now
        arg = Environment.GetCommandLineArgs()
        ButtonReCalc.Enabled = False
        ButtonCargar.Enabled = True
        ButtonSave.Enabled = False
        If arg.Length > 1 Then
            If arg.Length >= 6 Then
                Usuario = arg(5)
                If arg.Length >= 7 Then DiasMenos = arg(6)
                If Usuario = "lhernandez" Then
                    AHORA = FECHA_APLICACION
                End If
            End If
            If arg(3) = "FIN" Then
                ''Console.WriteLine("1")
                txtanexo.Text = arg(1)
                txtCiclo.Text = arg(2)
                CheckProyectado.Checked = arg(4)
                If CheckProyectado.Checked = True Then
                    If arg.Length >= 8 Then
                        AHORA = CadenaFecha(arg(7))
                    End If
                End If
                'Console.WriteLine("carga")
                ButtonCargar_Click(Nothing, Nothing)
                'Console.WriteLine("recalc")
                ButtonRecalc_Click(Nothing, Nothing)
                'Console.WriteLine("save")
                ButtonSave_Click(Nothing, Nothing)
                'Console.WriteLine("end")
                End
            ElseIf UCase(arg(3)) = "ECT" Then
                ProcesaAvio()
                End
            End If
        Else
            ReDim arg(5)
            arg(1) = ""
            arg(2) = ""
            arg(4) = 0
            End
        End If
        txtanexo.Text = arg(1)
        txtCiclo.Text = arg(2)

        CheckProyectado.Checked = arg(4)
        If CheckProyectado.Checked = True Then
            If arg.Length >= 8 Then
                AHORA = CadenaFecha(arg(7))
            End If
        End If

        'Catch ex As Exception
        '    Console.WriteLine(ex.Message)
        '    End
        'End Try
        'DiasMenos = -190
        'AHORA = Now
        'MessageBox.Show(AHORA.AddDays(DiasMenos).ToString("dd/MM/yyyy"))
        Label3.Text = AHORA.ToShortDateString
    End Sub

    Sub ProcesaAvio()
        'procesa Avios con Saldo
        Dim contador As Integer = 0
        Dim ww As New Estado_de_Cuenta.ProductionDataSetTableAdapters.SaldosAvioTableAdapter
        Dim TT As New Estado_de_Cuenta.ProductionDataSet.SaldosAvioDataTable
        ww.TerminaContratos(AHORA.ToString("yyyyMMdd"))
        ww.Fill(TT)
        For Each r As Estado_de_Cuenta.ProductionDataSet.SaldosAvioRow In TT.Rows
            txtanexo.Text = r.Anexo
            txtCiclo.Text = r.Ciclo
            CheckProyectado.Checked = False
            ButtonCargar_Click(Nothing, Nothing)
            ButtonRecalc_Click(Nothing, Nothing)
            If r.Flcan = "T" Then
                ww.ActivaContrato("A", r.Anexo, r.Ciclo)
            End If
            contador += 1
        Next
    End Sub

    Private Sub ButtonCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCancelar.Click
        ButtonReCalc.Enabled = False
        ButtonCargar.Enabled = True
        ButtonSave.Enabled = False
        txtanexo.Enabled = True
        txtCiclo.Enabled = True
        Me.ProductionDataSet.DetalleFINAGIL.Clear()
        Me.ProductionDataSet1.DetalleFINAGIL.Clear()

        LBtotal.Text = 0.0
        LBtotal2.Text = 0.0
        LbTotal3.Text = 0.0

    End Sub

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        Dim Respuesta As String = ""
        If arg(3) = "FIN" Or UCase(arg(3)) = "ECT" Then
            Respuesta = "SI"
        Else
            Respuesta = InputBox("Dame contraseña", "Contraseña de Cambio")
        End If

        If (Respuesta = "SI" Or arg(3) = "FIN") And Val(LBtotal.Text) > 0 And Me.ProductionDataSet1.DetalleFINAGIL.Rows.Count > 0 Then
            Dim r As Estado_de_Cuenta.ProductionDataSet.DetalleFINAGILRow
            Me.DetalleFINAGILTableAdapter.DeleteAnexo(txtanexo.Text, txtCiclo.Text)
            Me.ProductionDataSet.DetalleFINAGIL.Clear()
            For Each r In Me.ProductionDataSet1.DetalleFINAGIL.Rows
                Me.DetalleFINAGILTableAdapter.Insert(r.Anexo, r.Ciclo, r.Cliente, r.Consecutivo, r.FechaInicial, r.FechaFinal, r.Dias, r.TasaBP, r.SaldoInicial, r.Concepto, r.Importe, r.FEGA, r.Garantia, r.Intereses, r.SaldoFinal, "01/02/2099", False, r.Minds, r.Facturado, r.Factura, r.FolioFiscal)
                UltimoCorte = r.FechaFinal
            Next
            If UltimoCorte <> Nothing Or UltimoCorte <> "" Then
                Me.DetalleFINAGILTableAdapter.Ultimocorte(UltimoCorte, txtCiclo.Text, txtanexo.Text)
            End If
            UltimoCorte = ""

            ButtonReCalc.Enabled = False
            ButtonCargar.Enabled = True
            ButtonSave.Enabled = False
            txtanexo.Enabled = True
            txtCiclo.Enabled = True
            LBtotal.Text = 0.0
            LBtotal2.Text = 0.0
            LbTotal3.Text = 0.0
            SumaFin = DetalleFINAGILTableAdapter.SumaCargosPag(txtanexo.Text, txtCiclo.Text)
            Me.ProductionDataSet1.DetalleFINAGIL.Clear()
            Me.DetalleFINAGILTableAdapter.FillByAnexo(Me.ProductionDataSet.DetalleFINAGIL, txtanexo.Text, txtCiclo.Text)
            Me.ClientesTableAdapter.Fill(Me.ProductionDataSet.Clientes, LabelCliente.Text)
            LBtotal.Text = Format(Me.DetalleFINAGILTableAdapter.Total(txtanexo.Text, txtCiclo.Text), "N")
            Me.DetalleFINAGILTableAdapter.UpdateUltimoSaldo(LBtotal.Text, txtanexo.Text, txtCiclo.Text)
            If SumaIni <> SumaFin Then
                Dim MSG As String = "Contrato: "
                MSG += txtanexo.Text & "-" & txtCiclo.Text & " <br/>"
                MSG += "Monto  Inicial: " & SumaIni.ToString("n2") & " <br/>"
                MSG += "Monto  Final: " & SumaFin.ToString("n2") & " <br/>"
                MSG += "Usuario: " & Usuario & " <br/>"
                EnviaError("ecacerest@lamoderna.com.mx", MSG, "Anexo Procesado: " & txtanexo.Text & "-" & txtCiclo.Text)
            End If

        End If
    End Sub

    Private Sub LBtotal2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LBtotal2.Click
        Me.DetalleFINAGILTableAdapter.Update(Me.ProductionDataSet.DetalleFINAGIL)
        Me.ProductionDataSet.DetalleFINAGIL.AcceptChanges()
    End Sub

    Sub GeneraIntereses(ByRef f As String, ByRef Consec As Integer, ByRef r As Estado_de_Cuenta.ProductionDataSet.DetalleFINAGILRow, ByRef SaldoFin As Double, ByVal Correcion As Integer, ByVal H As String)
        Dim fec As Date = CadenaFecha(f)
        Dim Intereses As Double
        Dim SaldoIni As Double
        Dim dias As Integer
        Dim hoy As Date = CadenaFecha(H)
        Dim aux As Date = CadenaFecha(fec.ToString("yyyyMM01"))
        Dim FechaAnt As Date = Now
        Dim Mora As Boolean = False
        Dim FinMes As Integer = 0
        Dim Vencido As Integer = 0
        ' Dim Bandera As Boolean = False
        hoy = hoy.AddDays(-1)

        Dim rri As Estado_de_Cuenta.ProductionDataSet.DetalleFINAGILRow
        If fec <= hoy And SaldoFin > 0 Then
            FechaAnt = aux.AddMonths(-1)
            aux = aux.AddMonths(1)
            aux = aux.AddDays(-1)

            'FijaTasa(r("anexo"), r("ciclo"), FechaAnt, r("tasabp"))
            'If r.Anexo = "085930008" And aux = CDate("31/03/2015") Then
            '    aux = CadenaFecha("20150325")
            '    Bandera = True
            'End If

            If fec < FechaVen And FechaVen < aux And fec.Month = FechaVen.Month And aux.Month = FechaVen.Month Then
                aux = FechaVen
                FechaAnt = aux.AddMonths(-1)
                Vencido = 1
            End If
            If fec > FechaVen And aux = fec And hoy > fec Then
                aux = hoy
                FechaAnt = aux.AddMonths(-1)
            End If

            FijaTasa(r("anexo"), r("ciclo"), FechaAnt)

            If FechaVen < aux And Tipar = "C" And SinMoratorios = "N" Then
                Tasa = Tasa * 2
                Mora = True
            ElseIf FechaVen < aux And (Tipar = "H" Or Tipar = "A") And SinMoratorios = "N" Then
                Tasa = Tasa * 3
                Mora = True
            End If

            'If UltimoPAGO = f And Correcion = 1 And CadenaFecha(UltimoPAGO).AddDays(1).Day = 1 And CadenaFecha(UltimoPAGO).AddDays(1) = CadenaFecha(H).AddDays(1) Then
            '    Correcion = 0
            '    PasaUltimoPago = False
            'End If


            dias = DateDiff(DateInterval.Day, fec, aux) + Correcion
            If dias > 0 Then
                If UltimoPAGO = f And Correcion = 1 And CadenaFecha(UltimoPAGO).AddDays(1).Day = 1 Then
                    UltimoCorte = aux.ToString("yyyyMMdd")
                    aux = aux.AddDays(1)
                    f = aux.ToString("yyyyMMdd")
                    GeneraIntereses(f, Consec, r, SaldoFin, 1, H)
                Else
                    r.Consecutivo = Consec * 100
                    Me.ProductionDataSet1.DetalleFINAGIL.ImportRow(r)
                    rri = Me.ProductionDataSet1.DetalleFINAGIL.Rows(Consec)
                    'calcula interes dentro de un mes
                    SaldoIni = SaldoFin
                    'LINEA INETRES

                    Consec += 1
                    '

                    rri("Consecutivo") = Consec
                    fec = fec.AddDays(Correcion * -1)
                    rri("fechainicial") = fec.ToString("yyyyMMdd")
                    fec = fec.AddDays(Correcion)
                    rri("fechafinal") = aux.ToString("yyyyMMdd")
                    'If Array.IndexOf(SinMoraHastaPago, txtanexo.Text) >= 0 Then 'ELISANDER sin moratorios despues del ultimo pago
                    If UltimoPAGO_GEN.Trim <> "" And UltimoPAGO_GEN < aux.ToString("yyyyMMdd") Then
                        'If Array.IndexOf(SinMoraHastaPagoTasamenor, txtanexo.Text) >= 0 Then 'ELISANDER sin moratorios despues del ultimo pago
                        If TercioDeTasa = True Then
                            Tasa = TasaAUX / 3
                        Else
                            Tasa = TasaAUX
                        End If
                    End If
                    ' End If
                    rri("dias") = dias
                    rri("tasabp") = Tasa
                    rri("saldoinicial") = SaldoIni
                    rri("concepto") = "INTERESES"
                    rri("Facturado") = 1
                    rri("Factura") = ""
                    rri("FolioFiscal") = ""
                    rri("importe") = 0
                    rri("fega") = 0
                    rri("garantia") = 0
                    Intereses = rri("saldoinicial") * (Tasa / 100 / 360) * dias
                    rri("intereses") = Intereses
                    SaldoFin = SaldoIni + Intereses
                    rri("saldofinal") = SaldoFin
                    UltimoCorte = rri("fechafinal")
                    aux = aux.AddDays(1)
                    f = aux.ToString("yyyyMMdd")
                    If Vencido = 1 Then
                        UltimoCorte = FechaVen.ToString("yyyyMMdd")
                    End If
                    '+++++Seguro de Vida+++++++++++++++++++++++++++++++++++++++++++++++++++
                    FecVid = CadenaFecha(UltimoCorte)
                    If FecVid.AddDays(1).Day = 1 Then
                        FinMes = 1
                    ElseIf FecVid.AddDays(1).Day = 2 Then
                        FinMes = 0
                    End If
                    If FecVid = FecTope Then
                        FinMes = 0
                    End If

                    If FecVid.AddDays(FinMes).Day = 1 And TasaSegVid > 0 And UltimoCorte > "20140901" And CalcSEGVID = True Then 'aplica seguro de vida
                        r.Consecutivo = Consec * 100
                        Me.ProductionDataSet1.DetalleFINAGIL.ImportRow(r)
                        rri = Me.ProductionDataSet1.DetalleFINAGIL.Rows(Consec)
                        SaldoIni = SaldoFin
                        If FecVid.Day = 1 Then
                            FecVid = FecVid.AddDays(-1)
                        End If
                        DatFact = SacaDatosFactura(FecVid.ToString("yyyyMMdd"))
                        'LINEA Seguro de Vida
                        Consec += 1
                        rri("Consecutivo") = Consec
                        rri("fechainicial") = FecVid.ToString("yyyyMMdd")
                        rri("fechafinal") = FecVid.ToString("yyyyMMdd")
                        rri("dias") = 0
                        rri("tasabp") = 0
                        rri("Intereses") = 0
                        rri("saldoinicial") = SaldoIni
                        rri("concepto") = "SEGURO DE VIDA"
                        rri("Facturado") = DatFact(0)
                        rri("Factura") = DatFact(1)
                        rri("FolioFiscal") = DatFact(2)
                        SegVid = CalculaPrima(rri.Cliente, UltimoCorte, SaldoIni)
                        rri("Importe") = SegVid
                        rri("fega") = 0
                        rri("garantia") = 0
                        If Tipar = "C" And Fondeo = "03" And Mora = False Then
                            If FechaAutorizacion >= "20160101" Then
                                rri("fega") = SegVid * 0.0174
                                If PorcFega > 0 Then
                                    rri("fega") = SegVid * PorcFega
                                End If
                            Else
                                rri("fega") = SegVid * 0.0116
                            End If

                            rri("garantia") = 0
                        End If
                        If Tipar = "H" And Fondeo = "03" And Mora = False Then
                            If FechaAutorizacion >= "20160101" Then
                                rri("fega") = SegVid * 0.0174
                                If PorcFega > 0 Then
                                    rri("fega") = SegVid * PorcFega
                                End If
                            Else
                                rri("fega") = SegVid * 0.0116
                            End If
                            rri("garantia") = SegVid * 0.1
                            If UCase(AplicaGarantiaLIQ) = "NO" Then
                                rri("garantia") = 0
                            End If
                        End If
                        '***** SE QUITA FEGA Y GL DEL SEGURO DE VIDA A PARTIR DEL MES DE SEPTIEMBRE, ELISANDER PINEDA #ect 26092015.N
                        If FecVid > FechaQuitaFega Then
                            rri("fega") = 0
                            rri("garantia") = 0
                        End If
                        '**************************************************************************
                        SaldoFin = SaldoIni + SegVid + rri("fega") + rri("garantia")
                        rri("saldofinal") = SaldoFin
                    End If
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    If Vencido = 1 Then
                        'If Bandera = False Then
                        GeneraIntereses(f, Consec, r, SaldoFin, 0, H)
                        'End If
                    Else
                        'If Bandera = False Then
                        GeneraIntereses(f, Consec, r, SaldoFin, 1, H)
                        ' End If
                    End If

                End If

            Else
                If SaldoFin > 0 Then
                    GeneraIntereses(f, Consec, r, SaldoFin, 1, H)
                End If
            End If

        End If
    End Sub

    Sub GeneraInteresesII(ByRef f As String, ByRef Consec As Integer, ByRef r As Estado_de_Cuenta.ProductionDataSet.DetalleFINAGILRow, ByRef SaldoFin As Double, ByVal Correcion As Integer, ByVal H As String)
        Dim fec As Date = CadenaFecha(f)
        Dim FinMes As Integer = 0
        Dim Intereses As Double
        Dim SaldoIni As Double
        Dim dias As Integer
        Dim hoy As Date = CadenaFecha(H)
        Dim aux As Date = CadenaFecha(fec.ToString("yyyyMMdd"))
        Dim FechaAnt As Date = Now
        Dim rri As Estado_de_Cuenta.ProductionDataSet.DetalleFINAGILRow
        Dim Mora As Boolean = False
        If fec <= hoy And SaldoFin > 0 Then
            'Console.WriteLine("GeneraInteresesII1")
            FechaAnt = aux.AddMonths(-1)
            If aux.Month <> hoy.Month Then
                aux = CadenaFecha(fec.ToString("yyyyMM01"))
                aux = aux.AddMonths(1)
                aux = aux.AddDays(-1)
            Else
                If aux.Year = hoy.Year Then
                    aux = hoy
                Else
                    aux = CadenaFecha(fec.ToString("yyyyMM01"))
                    aux = aux.AddMonths(1)
                    aux = aux.AddDays(-1)
                End If

            End If
            'Console.WriteLine("fijatasa1")
            FijaTasa(r("anexo"), r("ciclo"), FechaAnt)
            'Console.WriteLine("fijatasa2")

            If fec < FechaVen And FechaVen < aux And fec.Month = FechaVen.Month And aux.Month = FechaVen.Month Then
                aux = FechaVen
            End If

            If FechaVen < aux And Tipar = "C" And SinMoratorios = "N" Then
                Tasa = Tasa * 2
                Mora = True
            ElseIf FechaVen < fec And (Tipar = "H" Or Tipar = "A") And SinMoratorios = "N" Then
                Tasa = Tasa * 3
                Mora = True
            End If



            dias = DateDiff(DateInterval.Day, fec, aux) + Correcion
            If dias > 0 Then
                'Console.WriteLine("interes")
                r.Consecutivo = Consec * 1030
                Me.ProductionDataSet1.DetalleFINAGIL.ImportRow(r)
                rri = Me.ProductionDataSet1.DetalleFINAGIL.Rows(Consec)
                'calcula interes dentro de un mes
                SaldoIni = SaldoFin
                'LINEA INETRES
                Consec += 1
                rri("Consecutivo") = Consec
                fec = fec.AddDays(Correcion * -1)
                rri("fechainicial") = fec.ToString("yyyyMMdd")
                fec = fec.AddDays(Correcion)
                rri("fechafinal") = aux.ToString("yyyyMMdd")
                'If Array.IndexOf(SinMoraHastaPago, txtanexo.Text) >= 0 Then 'ELISANDER sin moratorios despues del ultimo pago
                If UltimoPAGO_GEN.Trim <> "" And UltimoPAGO_GEN < aux.ToString("yyyyMMdd") Then
                    'If Array.IndexOf(SinMoraHastaPagoTasamenor, txtanexo.Text) >= 0 Then 'ELISANDER sin moratorios despues del ultimo pago
                    If TercioDeTasa = True Then
                        Tasa = TasaAUX / 3
                    Else
                        Tasa = TasaAUX
                    End If
                End If
                'End If
                rri("dias") = dias
                rri("tasabp") = Tasa
                rri("saldoinicial") = SaldoIni
                rri("concepto") = "INTERESES"
                rri("Facturado") = 1
                rri("Factura") = ""
                rri("FolioFiscal") = ""
                rri("importe") = 0
                rri("fega") = 0
                rri("garantia") = 0
                Intereses = rri("saldoinicial") * (Tasa / 100 / 360) * dias
                rri("intereses") = Intereses
                SaldoFin = SaldoIni + Intereses
                rri("saldofinal") = SaldoFin
                UltimoCorte = rri("fechafinal")
                '+++++Seguro de Vida+++++++++++++++++++++++++++++++++++++++++++++++++++
                FecVid = CadenaFecha(UltimoCorte)
                If FecVid.AddDays(1).Day = 1 Then
                    FinMes = 1
                ElseIf FecVid.AddDays(1).Day = 2 Then
                    FinMes = 1
                End If

                'If FecVid.AddDays(FinMes).Day = 1 And TasaSegVid > 0 Then 'aplica seguro de vida
                If FecVid.AddDays(FinMes).Day = 1 And TasaSegVid > 0 And UltimoCorte > "20140901" And CalcSEGVID = True Then
                    'Console.WriteLine("Vida")
                    r.Consecutivo = Consec * 100
                    Me.ProductionDataSet1.DetalleFINAGIL.ImportRow(r)
                    rri = Me.ProductionDataSet1.DetalleFINAGIL.Rows(Consec)
                    SaldoIni = SaldoFin
                    'LINEA Seguro de Vida
                    Consec += 1
                    'Console.WriteLine("VidaFactura1" & UltimoCorte)
                    DatFact = SacaDatosFactura(UltimoCorte)
                    'Console.WriteLine("VidaFactura2" & UltimoCorte)
                    rri("Consecutivo") = Consec
                    rri("fechainicial") = UltimoCorte
                    rri("fechafinal") = UltimoCorte
                    rri("dias") = 0
                    rri("tasabp") = 0
                    rri("Intereses") = 0
                    rri("saldoinicial") = SaldoIni
                    rri("concepto") = "SEGURO DE VIDA"
                    rri("Facturado") = DatFact(0)
                    rri("Factura") = DatFact(1)
                    rri("FolioFiscal") = DatFact(2)
                    'Console.WriteLine("prima1")
                    SegVid = CalculaPrima(rri.Cliente, UltimoCorte, SaldoIni)
                    'Console.WriteLine("prima2")
                    rri("Importe") = SegVid
                    rri("fega") = 0
                    rri("garantia") = 0
                    'Console.WriteLine("C")
                    If Tipar = "C" And Fondeo = "03" And Mora = False Then
                        If FechaAutorizacion >= "20160101" Then
                            rri("fega") = SegVid * 0.0174
                            If PorcFega > 0 Then
                                rri("fega") = SegVid * PorcFega
                            End If
                        Else
                            rri("fega") = SegVid * 0.0116
                        End If
                        rri("garantia") = 0
                    End If
                    'Console.WriteLine("h")
                    If Tipar = "H" And Fondeo = "03" And Mora = False Then
                        If FechaAutorizacion >= "20160101" Then
                            rri("fega") = SegVid * 0.0174
                            If PorcFega > 0 Then
                                rri("fega") = SegVid * PorcFega
                            End If
                        Else
                            rri("fega") = SegVid * 0.0116
                        End If
                        rri("garantia") = SegVid * 0.1
                        If UCase(AplicaGarantiaLIQ) = "NO" Then
                            rri("garantia") = 0
                        End If
                    End If
                    'Console.WriteLine("h2")
                    '***** SE QUITA FEGA Y GL DEL SEGURO DE VIDA A PARTIR DEL MES DE SEPTIEMBRE, ELISANDER PINEDA #ect 26092015.N
                    If FecVid > FechaQuitaFega Then
                        rri("fega") = 0
                        rri("garantia") = 0
                    End If
                    '**************************************************************************
                    'Console.WriteLine("h3")
                    SaldoFin = SaldoIni + SegVid + rri("fega") + rri("garantia")
                    rri("saldofinal") = SaldoFin
                    'Console.WriteLine("hc fin")
                End If
                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            End If
            aux = aux.AddDays(1)
            f = aux.ToString("yyyyMMdd")
            'Console.WriteLine("GeneraInteresesIIinterno1")
            GeneraInteresesII(f, Consec, r, SaldoFin, 1, H)
            'Console.WriteLine("GeneraInteresesIIinterno2")
        Else
            'Console.WriteLine("GeneraInteresesII2")
            fec = CadenaFecha(UltimoCorte)
            If fec < hoy And SaldoFin > 0 Then
                aux = hoy
                FechaAnt = CadenaFecha(UltimoCorte)
                dias = DateDiff(DateInterval.Day, fec, aux)
                If dias = 1 Then
                    r.Consecutivo = Consec * 1040
                    Me.ProductionDataSet1.DetalleFINAGIL.ImportRow(r)
                    rri = Me.ProductionDataSet1.DetalleFINAGIL.Rows(Consec)
                    'calcula interes dentro de un mes
                    SaldoIni = SaldoFin
                    'LINEA INETRES

                    Consec += 1
                    rri("Consecutivo") = Consec
                    rri("fechainicial") = fec.ToString("yyyyMMdd")
                    rri("fechafinal") = aux.ToString("yyyyMMdd")
                    rri("dias") = dias
                    FijaTasa(r("anexo"), r("ciclo"), FechaAnt)
                    If FechaVen < fec And Tipar = "C" Then
                        Tasa = Tasa * 2
                    ElseIf FechaVen < fec And (Tipar = "H" Or Tipar = "A") Then
                        Tasa = Tasa * 3
                    End If

                    'If Array.IndexOf(SinMoraHastaPago, txtanexo.Text) >= 0 Then 'ELISANDER sin moratorios despues del ultimo pago
                    If UltimoPAGO_GEN.Trim <> "" And UltimoPAGO_GEN < aux.ToString("yyyyMMdd") Then
                        'If Array.IndexOf(SinMoraHastaPagoTasamenor, txtanexo.Text) >= 0 Then 'ELISANDER sin moratorios despues del ultimo pago
                        If TercioDeTasa = True Then
                            Tasa = TasaAUX / 3
                        Else
                            Tasa = TasaAUX
                        End If
                    End If
                    'End If

                    rri("tasabp") = Tasa
                    rri("saldoinicial") = SaldoIni
                    rri("concepto") = "INTERESES"
                    rri("Facturado") = 1
                    rri("Factura") = ""
                    rri("FolioFiscal") = ""
                    rri("importe") = 0
                    rri("fega") = 0
                    rri("garantia") = 0
                    Intereses = rri("saldoinicial") * (Tasa / 100 / 360) * dias
                    rri("intereses") = Intereses
                    SaldoFin = SaldoIni + Intereses
                    rri("saldofinal") = SaldoFin
                    UltimoCorte = rri("fechafinal")
                    '+++++Seguro de Vida+++++++++++++++++++++++++++++++++++++++++++++++++++
                    FecVid = CadenaFecha(UltimoCorte)
                    If FecVid.AddDays(1).Day = 1 Then
                        FinMes = 1
                    ElseIf FecVid.AddDays(1).Day = 2 Then
                        FinMes = 1
                    End If

                    'If FecVid.AddDays(FinMes).Day = 1 And TasaSegVid > 0 Then 'aplica seguro de vida
                    If FecVid.AddDays(FinMes).Day = 1 And TasaSegVid > 0 And UltimoCorte > "20140901" And CalcSEGVID = True Then
                        r.Consecutivo = Consec * 100
                        Me.ProductionDataSet1.DetalleFINAGIL.ImportRow(r)
                        rri = Me.ProductionDataSet1.DetalleFINAGIL.Rows(Consec)
                        SaldoIni = SaldoFin
                        'LINEA Seguro de Vida
                        DatFact = SacaDatosFactura(UltimoCorte)
                        Consec += 1
                        rri("Consecutivo") = Consec
                        rri("fechainicial") = UltimoCorte
                        rri("fechafinal") = UltimoCorte
                        rri("dias") = 0
                        rri("tasabp") = 0
                        rri("Intereses") = 0
                        rri("saldoinicial") = SaldoIni
                        rri("concepto") = "SEGURO DE VIDA"
                        rri("Facturado") = DatFact(0)
                        rri("Factura") = DatFact(1)
                        rri("FolioFiscal") = DatFact(2)
                        SegVid = CalculaPrima(rri.Cliente, UltimoCorte, SaldoIni)
                        rri("Importe") = SegVid
                        rri("fega") = 0
                        rri("garantia") = 0
                        If Tipar = "C" And Fondeo = "03" Then
                            If FechaAutorizacion >= "20160101" Then
                                rri("fega") = SegVid * 0.0174
                                If PorcFega > 0 Then
                                    rri("fega") = SegVid * PorcFega
                                End If
                            Else
                                rri("fega") = SegVid * 0.0116
                            End If
                            rri("garantia") = 0
                        End If
                        If Tipar = "H" And Fondeo = "03" Then
                            If FechaAutorizacion >= "20160101" Then
                                rri("fega") = SegVid * 0.0174
                                If PorcFega > 0 Then
                                    rri("fega") = SegVid * PorcFega
                                End If
                            Else
                                rri("fega") = SegVid * 0.0116
                            End If
                            rri("garantia") = SegVid * 0.1
                            If UCase(AplicaGarantiaLIQ) = "NO" Then
                                rri("garantia") = 0
                            End If
                        End If
                        '***** SE QUITA FEGA Y GL DEL SEGURO DE VIDA A PARTIR DEL MES DE SEPTIEMBRE, ELISANDER PINEDA #ect 26092015.N
                        If FecVid > FechaQuitaFega Then
                            rri("fega") = 0
                            rri("garantia") = 0
                        End If
                        '**************************************************************************
                        SaldoFin = SaldoIni + SegVid + rri("fega") + rri("garantia")
                        rri("saldofinal") = SaldoFin
                    End If
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                End If
            End If

        End If
    End Sub

    Function SeguroVida(ByVal a As String, ByVal c As String) As Decimal
        Dim Ta As New Estado_de_Cuenta.ProductionDataSet.DetalleFINAGILDataTable
        Me.DetalleFINAGILTableAdapter.FillByAnexoFecha(Ta, txtanexo.Text, txtCiclo.Text)
        Dim x As Integer = 0
        If Ta.Rows.Count > 0 Then
            For Each r As ProductionDataSet.DetalleFINAGILRow In Ta.Rows
                If Trim(r("Concepto")) = "SEGURO DE VIDA" Then
                    ArreSegVid(x, 0) = r.Facturado
                    ArreSegVid(x, 1) = r.Factura
                    ArreSegVid(x, 2) = r.FolioFiscal
                    ArreSegVid(x, 3) = r.FechaFinal
                    x += 1
                End If
            Next
        End If

        'TASA seguro de vida
        Dim XX As New Estado_de_Cuenta.ProductionDataSetTableAdapters.AviosTableAdapter
        Dim y As New Estado_de_Cuenta.ProductionDataSet.AviosDataTable
        XX.FillAnexo(y, a, c)
        Dim Tasa As Decimal = 0
        If Not IsDBNull(y.Rows(0).Item("SeguroVida")) And y.Rows.Count > 0 Then
            Tasa = y.Rows(0).Item("SeguroVida")
        End If
        If LabelTipo.Text = "M" Then 'Elisander pidio para Fisicas y fisicas con act. empresarial 
            Tasa = 0
        End If
        Return (Tasa)
    End Function

    Private Function SacaDatosFactura(ByVal f As String) As String()
        Dim REspuesta(3) As String
        REspuesta(0) = "False"
        REspuesta(1) = ""
        REspuesta(2) = ""
        REspuesta(3) = ""
        For x = 0 To 50
            If f = ArreSegVid(x, 3) Then
                REspuesta(0) = ArreSegVid(x, 0)
                REspuesta(1) = ArreSegVid(x, 1)
                REspuesta(2) = ArreSegVid(x, 2)
                REspuesta(3) = ArreSegVid(x, 3)
                Exit For
            End If
            If ArreSegVid(x, 0) = "" Then Exit For
        Next
        SacaDatosFactura = REspuesta
    End Function

    Private Function CalculaPrima(ByVal Cli As String, ByVal Fec As String, ByVal SaldoAnexo As Double) As Double
        Dim Factor As Double = 0
        Dim Saldos As New ProductionDataSetTableAdapters.SaldoClienteTableAdapter
        Dim T As New ProductionDataSet.SaldoClienteDataTable
        Dim Saldo As Double = 0
        Dim SegMes As Double = 0
        Dim Prima As Double = 0
        Saldos.Fill(T, Fec, Cli, txtanexo.Text, txtCiclo.Text)
        For Each r As ProductionDataSet.SaldoClienteRow In T.Rows
            If r.Anexo <> txtanexo.Text And r.Ciclo <> txtCiclo.Text Then
                SegMes += Saldos.SeguroVidaDelMes(Fec, txtanexo.Text, txtCiclo.Text)
                Saldo += r.Saldo
                Saldo -= SegMes
            End If
        Next
        If SaldoAnexo + Saldo > 5000000 Then
            Factor = SaldoAnexo / SaldoAnexo + Saldo
            Prima = ((5000000 / 1000) * TasaSegVid) * Factor
        Else
            Prima = (SaldoAnexo / 1000) * TasaSegVid
        End If
        If Prima < 0 Then Prima = 0
        Return Prima
    End Function

    Private Sub EnviaError(ByVal Para As String, ByVal Mensaje As String, ByVal Asunto As String)
        Dim Mensage As New MailMessage("Avio@cmoderna.com", Trim(Para), Trim(Asunto), Mensaje)
        Dim Cliente As New SmtpClient("smtp01.cmoderna.com", 26)
        Mensage.IsBodyHtml = True
        Mensage.Priority = MailPriority.High
        Cliente.Send(Mensage)
    End Sub

End Class
