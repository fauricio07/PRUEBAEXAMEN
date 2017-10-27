Class MainWindow

    Dim valor, resultado As Integer

#Region "Metodos varios"

    Sub cargaCombo1()

        cbxSeleccion.Items.Add("")
        cbxSeleccion.Items.Add("Centímetro")
        cbxSeleccion.Items.Add("Pulgada")
        cbxSeleccion.Items.Add("Pie")
        cbxSeleccion.Items.Add("Kilómetro")
        cbxSeleccion.Items.Add("Metro")

    End Sub

    Sub cargaCombo2()

        cbxSeleccion2.Items.Add("")
        cbxSeleccion2.Items.Add("Centímetro")
        cbxSeleccion2.Items.Add("Pulgada")
        cbxSeleccion2.Items.Add("Pie")
        cbxSeleccion2.Items.Add("Kilómetro")
        cbxSeleccion2.Items.Add("Metro")

    End Sub

    Sub centimetros()

        If (txtIngreso.Text <> String.Empty And IsNumeric(txtIngreso.Text)) Then

            valor = 0
            resultado = 0
            valor = Val(txtIngreso.Text)

            If (cbxSeleccion.Text <> String.Empty And cbxSeleccion2.Text <> String.Empty) Then

                'Inicio de las opciones del comboBox
                If (cbxSeleccion2.Text.Equals("Centímetro")) Then

                    resultado = valor * 1
                ElseIf (cbxSeleccion2.Text.Equals("Pulgada")) Then

                    resultado = valor * 0.3937
                ElseIf (cbxSeleccion2.Text.Equals("Pie")) Then

                    resultado = valor * 0.03281
                ElseIf (cbxSeleccion2.Text.Equals("Kilómetro")) Then

                    resultado = valor * 0.00001
                ElseIf (cbxSeleccion2.Text.Equals("Metro")) Then

                    resultado = valor * 0.01
                Else

                    MsgBox("No has seleccionado una de las unidades de medición existentes")
                End If

                'Fin de opciones del comboBox
            Else

                MsgBox("Debes seleccionar la unidad de medición actual y a la que desea convertir")
            End If

        Else

            MsgBox("Debes ingresar un valor a convertir primero")
        End If

    End Sub

    Sub pulgadas()

        If (txtIngreso.Text <> String.Empty And IsNumeric(txtIngreso.Text)) Then

            valor = 0
            resultado = 0
            valor = Val(txtIngreso.Text)

            If (cbxSeleccion.Text <> String.Empty And cbxSeleccion2.Text <> String.Empty) Then

                'Inicio de las opciones del comboBox
                If (cbxSeleccion2.Text.Equals("Centímetro")) Then

                    resultado = valor * 2.54
                ElseIf (cbxSeleccion2.Text.Equals("Pulgada")) Then

                    resultado = valor * 1
                ElseIf (cbxSeleccion2.Text.Equals("Pie")) Then

                    resultado = valor * 0.0833
                ElseIf (cbxSeleccion2.Text.Equals("Kilómetro")) Then

                    resultado = valor * 0.0000254
                ElseIf (cbxSeleccion2.Text.Equals("Metro")) Then

                    resultado = valor * 0.0254
                Else

                    MsgBox("No has seleccionado una de las unidades de medición existentes")
                End If

                'Fin de opciones del comboBox
            Else

                MsgBox("Debes seleccionar la unidad de medición actual y a la que desea convertir")
            End If

        Else

            MsgBox("Debes ingresar un valor a convertir primero")
        End If

    End Sub

    Sub pies()

        If (txtIngreso.Text <> String.Empty And IsNumeric(txtIngreso.Text)) Then

            valor = 0
            resultado = 0
            valor = Val(txtIngreso.Text)

            If (cbxSeleccion.Text <> String.Empty And cbxSeleccion2.Text <> String.Empty) Then

                'Inicio de las opciones del comboBox
                If (cbxSeleccion2.Text.Equals("Centímetro")) Then

                    resultado = valor * 30.48
                ElseIf (cbxSeleccion2.Text.Equals("Pulgada")) Then

                    resultado = valor * 12
                ElseIf (cbxSeleccion2.Text.Equals("Pie")) Then

                    resultado = valor * 1
                ElseIf (cbxSeleccion2.Text.Equals("Kilómetro")) Then

                    resultado = valor * 0.0003048
                ElseIf (cbxSeleccion2.Text.Equals("Metro")) Then

                    resultado = valor * 0.3048
                Else

                    MsgBox("No has seleccionado una de las unidades de medición existentes")
                End If

                'Fin de opciones del comboBox
            Else

                MsgBox("Debes seleccionar la unidad de medición actual y a la que desea convertir")
            End If

        Else

            MsgBox("Debes ingresar un valor a convertir primero")
        End If

    End Sub

    Sub kilometros()

        If (txtIngreso.Text <> String.Empty And IsNumeric(txtIngreso.Text)) Then

            valor = 0
            resultado = 0
            valor = Val(txtIngreso.Text)

            If (cbxSeleccion.Text <> String.Empty And cbxSeleccion2.Text <> String.Empty) Then

                'Inicio de las opciones del comboBox
                If (cbxSeleccion2.Text.Equals("Centímetro")) Then

                    resultado = valor * 100000
                ElseIf (cbxSeleccion2.Text.Equals("Pulgada")) Then

                    resultado = valor * 3.937
                ElseIf (cbxSeleccion2.Text.Equals("Pie")) Then

                    resultado = valor * 3281
                ElseIf (cbxSeleccion2.Text.Equals("Kilómetro")) Then

                    resultado = valor * 1
                ElseIf (cbxSeleccion2.Text.Equals("Metro")) Then

                    resultado = valor * 1000
                Else

                    MsgBox("No has seleccionado una de las unidades de medición existentes")
                End If

                'Fin de opciones del comboBox
            Else

                MsgBox("Debes seleccionar la unidad de medición actual y a la que desea convertir")
            End If

        Else

            MsgBox("Debes ingresar un valor a convertir primero")
        End If

    End Sub

    Sub metros()

        If (txtIngreso.Text <> String.Empty And IsNumeric(txtIngreso.Text)) Then

            valor = 0
            resultado = 0
            valor = Val(txtIngreso.Text)

            If (cbxSeleccion.Text <> String.Empty And cbxSeleccion2.Text <> String.Empty) Then

                'Inicio de las opciones del comboBox
                If (cbxSeleccion2.Text.Equals("Centímetro")) Then

                    resultado = valor * 100
                ElseIf (cbxSeleccion2.Text.Equals("Pulgada")) Then

                    resultado = valor * 39.37
                ElseIf (cbxSeleccion2.Text.Equals("Pie")) Then

                    resultado = valor * 3.281
                ElseIf (cbxSeleccion2.Text.Equals("Kilómetro")) Then

                    resultado = valor * 0.001
                ElseIf (cbxSeleccion2.Text.Equals("Metro")) Then

                    resultado = valor * 1
                Else

                    MsgBox("No has seleccionado una de las unidades de medición existentes")
                End If

                'Fin de opciones del comboBox
            Else

                MsgBox("Debes seleccionar la unidad de medición actual y a la que desea convertir")
            End If

        Else

            MsgBox("Debes ingresar un valor a convertir primero")
        End If

    End Sub

#End Region

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        cargaCombo1()
        cargaCombo2()

    End Sub

    Private Sub btnCalcular_Click(sender As Object, e As RoutedEventArgs) Handles btnCalcular.Click

        If (txtIngreso.Text <> String.Empty And IsNumeric(txtIngreso.Text)) Then

            If (cbxSeleccion2.Text <> String.Empty) Then

                If (cbxSeleccion.Text.Equals("Centímetro")) Then

                    centimetros()
                    txtResultado.Text = resultado.ToString
                    btnCalcular.IsEnabled = False
                    txtIngreso.IsEnabled = False
                    cbxSeleccion.IsEnabled = False
                    cbxSeleccion2.IsEnabled = False

                ElseIf (cbxSeleccion.Text.Equals("Pulgada")) Then

                    pulgadas()
                    txtResultado.Text = resultado.ToString
                    btnCalcular.IsEnabled = False
                    txtIngreso.IsEnabled = False
                    cbxSeleccion.IsEnabled = False
                    cbxSeleccion2.IsEnabled = False

                ElseIf (cbxSeleccion.Text.Equals("Pie")) Then

                    pies()
                    txtResultado.Text = resultado.ToString
                    btnCalcular.IsEnabled = False
                    txtIngreso.IsEnabled = False
                    cbxSeleccion.IsEnabled = False
                    cbxSeleccion2.IsEnabled = False

                ElseIf (cbxSeleccion.Text.Equals("Kilómetro")) Then

                    kilometros()
                    txtResultado.Text = resultado.ToString
                    btnCalcular.IsEnabled = False
                    txtIngreso.IsEnabled = False
                    cbxSeleccion.IsEnabled = False
                    cbxSeleccion2.IsEnabled = False

                ElseIf (cbxSeleccion.Text.Equals("Metro")) Then

                    metros()
                    txtResultado.Text = resultado.ToString
                    btnCalcular.IsEnabled = False
                    txtIngreso.IsEnabled = False
                    cbxSeleccion.IsEnabled = False
                    cbxSeleccion2.IsEnabled = False

                Else

                    MsgBox("Debes seleccionar la unidad de medición actual")
                End If

            Else

                MsgBox("Debes seleccionar la unidad de medición a la que desea convertir el dato")
            End If

        Else

            MsgBox("Verifique que haya ingresado un valor a convertir y que éste sea de tipo numérico")
        End If

    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As RoutedEventArgs) Handles btnLimpiar.Click

        If (txtIngreso.Text <> String.Empty Or txtResultado.Text <> String.Empty Or cbxSeleccion.Text <> String.Empty Or cbxSeleccion2.Text <> String.Empty) Then

            txtIngreso.Text = String.Empty
            txtResultado.Text = String.Empty
            cbxSeleccion.SelectedIndex = 0
            cbxSeleccion2.SelectedIndex = 0
            btnCalcular.IsEnabled = True
            txtIngreso.IsEnabled = True
            cbxSeleccion.IsEnabled = True
            cbxSeleccion2.IsEnabled = True

        Else

            MsgBox("El formulario ahora mismo se encuentra límpio")
        End If

    End Sub

End Class
