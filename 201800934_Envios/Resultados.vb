Module Resultados
    Public index As Byte = 0

    Public nombre(9) As String
    Public valor(9) As Double
    Public peso(9) As Double
    Public tipo_paquete(9) As String
    Public forma_envio(9) As String

    Public pago_parcial(9) As Double
    Public porcentaje_peso(9) As Double
    Public pago_total(9) As Double

    Const envio_camion As Double = 50
    Const envio_moto As Double = 15
    Const doc As Double = 0.015
    Const ropa As Double = 0.06
    Const perecedero As Double = 0.055
    Const plastico As Double = 0.045
    Const lociones As Double = 0.02

    Function PorcentajePeso() As Double
        Dim porcentaje As Double


        Dim cant As Double = peso(index) / 5

        Select Case Form1.cb_tipopaquete.SelectedIndex
            Case 0
                porcentaje = Form1.txt_valor.Text * doc * cant
            Case 1
                porcentaje = Form1.txt_valor.Text * ropa * cant
            Case 2
                porcentaje = Form1.txt_valor.Text * perecedero * cant
            Case 3
                porcentaje = Form1.txt_valor.Text * plastico * cant
            Case 4
                porcentaje = Form1.txt_valor.Text * lociones * cant
        End Select

        Return porcentaje
    End Function

    Function PagoParcial() As Double
        Dim pago As Double

        If Form1.rb_camion.Checked Then
            pago = Form1.txt_valor.Text + envio_camion
        ElseIf Form1.rb_moto.Checked Then
            pago = Form1.txt_valor.Text + envio_moto
        End If

        Return pago
    End Function

    Function PagoTotal() As Double
        Dim total As Double

        total = PorcentajePeso() + PagoParcial()
        Return total
    End Function
End Module
