Public Class Form1



    Private Sub btn_aceptar_Click(sender As Object, e As EventArgs) Handles btn_aceptar.Click

        If index < 10 Then
            If txt_nombre.Text <> "" Then
                nombre(index) = txt_nombre.Text
            Else
                MsgBox("Debe ingresar el nombre del cliente..")
            End If

            If txt_valor.Text <> "" Then
                valor(index) = txt_valor.Text
            Else
                MsgBox("Debe ingresar el valor del paquete..")
            End If

            If cb_tipopaquete.SelectedIndex <> -1 Then
                tipo_paquete(index) = cb_tipopaquete.Text
            Else
                MsgBox("Debe seleccionar el tipo de paquete..")
            End If

            If rb_camion.Checked Or rb_moto.Checked Then
                If rb_moto.Checked Then
                    forma_envio(index) = "Moto"
                ElseIf rb_camion.Checked Then
                    forma_envio(index) = "Camion"
                End If
            Else
                MsgBox("Debe seleccionar la forma de envio..")
            End If

            peso(index) = InputBox("", "Ingrese peso del paquete")

            porcentaje_peso(index) = System.Math.Round(Resultados.PorcentajePeso, 2)
            pago_parcial(index) = System.Math.Round(Resultados.PagoParcial, 2)
            pago_total(index) = System.Math.Round(Resultados.PagoTotal, 2)

            ListBox1.Items.Add(nombre(index))
            ListBox2.Items.Add(valor(index))
            ListBox3.Items.Add(peso(index))
            ListBox4.Items.Add(tipo_paquete(index))
            ListBox5.Items.Add(forma_envio(index))
            ListBox6.Items.Add(porcentaje_peso(index))
            ListBox7.Items.Add(pago_parcial(index))
            ListBox8.Items.Add(pago_total(index))

            index += 1

        Else
            MsgBox("No se pueden ingresar mas paquetes..")
        End If


    End Sub

    Private Sub MOSTRARRESULTADOSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MOSTRARRESULTADOSToolStripMenuItem.Click
        Dim total_plociones As Double
        Dim total_pdoc As Double
        Dim cant_ropa As Double
        Dim cant_plastico As Double

        For i = 0 To 9
            If tipo_paquete(i) = "Lociones" Then
                total_plociones += valor(i)
            ElseIf tipo_paquete(i) = "Documentos" Then
                total_pdoc += valor(i)
            ElseIf tipo_paquete(i) = "Articulos de plastico" Then
                cant_plastico += 1
            ElseIf tipo_paquete(i) = "Ropa" Then
                cant_ropa += 1
            End If

        Next i

        lb_q_lociones.Text = "Q" + total_plociones.ToString
        lb_q_doc.Text = "Q" + total_pdoc.ToString
        lb_cant_ropa.Text = cant_ropa
        lb_cant_plastico.Text = cant_plastico
    End Sub

    Private Sub LIMPIARENTRADASToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LIMPIARENTRADASToolStripMenuItem.Click
        txt_nombre.Clear()
        txt_nombre.Focus()
        txt_valor.Clear()
        cb_tipopaquete.SelectedIndex = -1
        rb_camion.Checked = False
        rb_moto.Checked = False

    End Sub

    Private Sub MOSTRARVECTORESToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MOSTRARVECTORESToolStripMenuItem.Click

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        ListBox5.Items.Clear()
        ListBox6.Items.Clear()
        ListBox7.Items.Clear()
        ListBox8.Items.Clear()


        For i = 0 To 9
            If nombre(i) = Nothing Then
                Exit For
            Else

                ListBox1.Items.Add(nombre(i))
                ListBox2.Items.Add(valor(i))
                ListBox3.Items.Add(peso(i))
                ListBox4.Items.Add(tipo_paquete(i))
                ListBox5.Items.Add(forma_envio(i))
                ListBox6.Items.Add(porcentaje_peso(i))
                ListBox7.Items.Add(pago_parcial(i))
                ListBox8.Items.Add(pago_total(i))
            End If

        Next i
    End Sub

    Private Sub LIMPIARVECTORESToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LIMPIARVECTORESToolStripMenuItem.Click

        For i = 0 To 9
            nombre(i) = Nothing
            valor(i) = Nothing
            peso(i) = Nothing
            tipo_paquete(i) = Nothing
            forma_envio(i) = Nothing
            porcentaje_peso(i) = Nothing
            pago_parcial(i) = Nothing
            pago_total(i) = Nothing
        Next i

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        ListBox5.Items.Clear()
        ListBox6.Items.Clear()
        ListBox7.Items.Clear()
        ListBox8.Items.Clear()
    End Sub

    Private Sub SALIRToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SALIRToolStripMenuItem.Click
        If MsgBox("Deseas salir", vbYesNo, "") = vbYes Then
            End
        End If
    End Sub


End Class
