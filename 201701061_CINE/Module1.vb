Module Module1

    Public tiposala(7)
    Public preciosala(7)
    Public nit(7)
    Public nompelicula(7)
    Public cantidadboletos(7)
    Public ptotal(7)
    Public proedio(7)

    Public Const cnormal = 44
    Public Const c3d = 62
    Public Const c4dx = 104

    Public vector As Byte = 0
    Public I As Byte
    Sub limpiar_entradas()
        Form1.TextBox1.Clear()
        Form1.TextBox2.Clear()
        Form1.TextBox3.Clear()
        Form1.TextBox4.Clear()
        Form1.ComboBox1.Text = ""
        Form1.TextBox1.Focus()
    End Sub

    Sub limpiar_vectores()

        Form1.DataGridView1.Rows.Clear()

        vector = 0

        For I = 0 To 6

            tiposala(vector) = Nothing
            preciosala(vector) = Nothing
            nit(vector) = Nothing
            nompelicula(vector) = Nothing
            cantidadboletos(vector) = Nothing

        Next I

    End Sub
End Module
