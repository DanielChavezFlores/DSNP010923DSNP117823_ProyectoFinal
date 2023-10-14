Public Class Principal

    Dim coincidencia_contador As Integer
    Dim Grupo As Object
    Dim Jugar As Boolean
    Dim intentos As Integer
    Dim Fallos As Integer


    Sub Ganador()

        Dim J As Integer

        For Each Grupo In Controls

            If TypeOf Grupo Is TextBox Then

                If Grupo.tag = 1 Then

                    If Grupo.forecolor = Color.Black Then

                        J = J + 1

                    End If

                End If

            End If

        Next

        If J = 10 Then
            MsgBox("Felicidades has ganado")
            Lista_palabras.Items.RemoveAt(Lista_palabras.SelectedIndex)
            Volver_a_jugar()
        End If


    End Sub


    Private Sub Principal_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        intentos = 0
        coincidencia_contador = 0
        Fallos = 0
        Jugar = False

    End Sub

    Private Sub Btn_Jugar_Click(sender As Object, e As EventArgs) Handles Btn_Jugar.Click

        Volver_a_jugar()

    End Sub

    Private Sub Principal_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Seleccionarletras(e.KeyCode.ToString)

    End Sub


    Sub Seleccionarletras(cadena As String)
        If Jugar Then

            Dim coincidencia As Boolean

            For Each Grupo In Controls
                If TypeOf Grupo Is TextBox Then
                    If Grupo.text = cadena Then
                        Grupo.forecolor = Color.Black
                        coincidencia = True
                    End If
                End If

                If TypeOf Grupo Is Button Then

                    If Grupo.text = cadena Then

                        Grupo.enabled = False

                    End If
                End If

            Next
            If coincidencia Then
                coincidencia_contador = coincidencia_contador + 1
            Else
                Fallos = Fallos + 1
            End If

            intentos = intentos + 1
            Caja_de_intentos.Text = intentos
            Caja_aciertos.Text = coincidencia_contador
            Caja_fallos.Text = Fallos

            If Fallos < 6 Then
                PictureBox1.Image = ImageList1.Images(Fallos)
            ElseIf Fallos >= 6 Then
                PictureBox1.Image = ImageList1.Images(Fallos)
                MsgBox("Perdiste")
                Volver_a_jugar()
            End If

            Ganador()

        End If

    End Sub

    Private Sub Teclado_virtual(sender As Object, e As EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Button4.Click, Button5.Click, Button6.Click, Button7.Click, Button8.Click, Button9.Click,
        Button10.Click, Button11.Click, Button12.Click, Button13.Click, Button14.Click, Button15.Click, Button16.Click, Button17.Click, Button18.Click, Button19.Click, Button20.Click, Button21.Click, Button22.Click,
        Button23.Click, Button24.Click, Button25.Click, Button26.Click

        Seleccionarletras(sender.text)

    End Sub


    Sub Volver_a_jugar()

        Dim Aleatorio As New Random
        Dim N, i As Integer

        N = 0
        N = Aleatorio.Next(0, Lista_palabras.Items.Count)
        Lista_palabras.SelectedIndex = N

        i = 11

        For Each Grupo In Controls

            If TypeOf Grupo Is TextBox Then

                i = i - 1
                Grupo.text = Mid(Lista_palabras.SelectedItem, i, 1)
                Grupo.forecolor = Color.White
            End If

            If TypeOf Grupo Is Button Then

                Grupo.enabled = True

            End If

        Next

        Caja_aciertos.Text = 0
        Caja_de_intentos.Text = 0
        Caja_fallos.Text = 0
        coincidencia_contador = 0
        intentos = 0
        Fallos = 0
        Jugar = True
        PictureBox1.Image = ImageList1.Images(0)

    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        MsgBox("La aplicacion se cerrara")
        End
    End Sub
End Class
