Public Class Form1
    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        If MsgBox(" ¿ Desea salir del programa ? ", vbQuestion + vbYesNo, "salir") = vbYes Then
            End
        End If
    End Sub

    Private Sub EliminarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EliminarToolStripMenuItem.Click
        Dim I As Byte

        tiposala(vector) = Nothing
        preciosala(vector) = Nothing
        nit(vector) = Nothing
        nompelicula(vector) = Nothing
        cantidadboletos(vector) = Nothing


        For I = vector To 6

            tiposala(I) = tiposala(I + 1)
            preciosala(I) = preciosala(I + 1)
            nit(I) = nit(I + 1)
            nompelicula(I) = nompelicula(I + 1)
            cantidadboletos(I) = cantidadboletos(vector)

        Next I
        MsgBox("Registro Eliminado exitosamente")

        tiposala(vector) = Nothing
        preciosala(vector) = Nothing
        nit(vector) = Nothing
        nompelicula(vector) = Nothing
        cantidadboletos(vector) = Nothing


        vector = I
        limpiar_entradas()
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub ActualizarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ActualizarToolStripMenuItem.Click

        preciosala(vector) = TextBox2.Text
        nit(vector) = Val(TextBox2.Text)
        cantidadboletos(vector) = Val(TextBox3.Text)
        tiposala(vector) = ComboBox1.Text

        MsgBox("Registro actualizado correctamente ")

        vector = 6
    End Sub



    Private Sub LimpiarVectoresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LimpiarVectoresToolStripMenuItem.Click
        limpiar_vectores()
    End Sub

    Private Sub MostrarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MostrarToolStripMenuItem.Click
        DataGridView1.Rows.Clear()
        For I = 0 To 6
            If (nit(I) <> Nothing) Then
                DataGridView1.Rows.Add(nompelicula(I), (nit(I)), (cantidadboletos(I)), (tiposala(I)), (ptotal(I)))
            End If
        Next
    End Sub

    Private Sub CalcularToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalcularToolStripMenuItem.Click
        If vector <= 6 Then

            If Val(TextBox1.Text <> "") Then
                nompelicula(vector) = Val(TextBox1.Text)
            Else
                MsgBox("Por favor ingresar nombre de la pelicula")
                TextBox1.Focus() : Exit Sub
            End If
            If Val(TextBox2.Text <> "") Then
                nit(vector) = Val(TextBox2.Text)
            Else
                MsgBox("Por favor ingresar NIT")
                TextBox2.Focus() : Exit Sub
            End If


            If Val(TextBox3.Text <> "") Then
                cantidadboletos(vector) = Val(TextBox3.Text)
            Else
                MsgBox("Por favor ingresar nombre de la pelicula")
                TextBox1.Focus() : Exit Sub
            End If



            Select Case ComboBox1.SelectedIndex
                Case 0
                    ptotal(vector) = Val(TextBox3.Text) * cnormal
                Case 1
                    ptotal(vector) = Val(TextBox3.Text) * c3d
                Case 2
                    ptotal(vector) = Val(TextBox3.Text) * c4dx
                Case Else
                    MsgBox(" No se seleciono ningun tipo de sala, por favor seleccione uno ")
                    TextBox1.Focus()
            End Select


            nompelicula(vector) = TextBox1.Text
            nit(vector) = TextBox2.Text
            cantidadboletos(vector) = TextBox3.Text
            tiposala(vector) = ComboBox1.Text


            vector = vector + 1
            limpiar_entradas()

            If vector = 7 Then
                MsgBox("Se llego al limite del almacenamiento")

            End If

        End If
    End Sub

    Private Sub ConsultarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsultarToolStripMenuItem.Click
        Dim existe As Boolean = False

        Dim I As Byte
        I = 0
        While (I <= 6) And Not (existe)

            If (nit(I) = Val(TextBox4.Text)) Then
                existe = True
            Else
                I = I + 1
            End If
        End While

        If (existe) Then
            MsgBox("Registro Encontrado exitosamente")

            TextBox1.Text = nompelicula(I)
            TextBox2.Text = nit(I)
            TextBox3.Text = cantidadboletos(I)
            ComboBox1.Text = tiposala(I)

            DataGridView1.Rows.Clear()
            DataGridView1.Rows.Add((nompelicula(I)), nit(I), (cantidadboletos(I)), (tiposala(I)), (ptotal(I)))

            vector = I
        Else
            MsgBox("Carnet no encontrado")
            TextBox4.Focus()
        End If

    End Sub
End Class
