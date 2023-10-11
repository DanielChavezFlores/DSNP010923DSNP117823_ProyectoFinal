Public Class Principal


    Dim Grupo As Object

    Private Sub Principal_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Btn_Jugar_Click(sender As Object, e As EventArgs) Handles Btn_Jugar.Click

        Dim Aleatorio As New Random
        Dim N, i As Integer



        N = Aleatorio.Next(0, Lista_palabras.Items.Count)
        MsgBox(N)
        Lista_palabras.SelectedIndex = N

        i = 11

        For Each Grupo In Controls

            If TypeOf Grupo Is TextBox Then

                i = i - 1
                Grupo.text = Mid(Lista_palabras.SelectedItem, i, 1)

            End If


        Next
        MsgBox(Lista_palabras.SelectedItem, i, 9)



    End Sub

    Private Sub Principal_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Seleccionarletras(e.KeyCode.ToString)

    End Sub


    Sub Seleccionarletras(cadena As String)

        For Each Grupo In Controls
            If TypeOf Grupo Is TextBox Then
                If Grupo.text = cadena Then
                    Grupo.forecolor = Color.Black

                End If
            End If

            If TypeOf Grupo Is Button Then

                If Grupo.text = cadena Then

                    Grupo.enabled = False

                End If
            End If

        Next


    End Sub

    Private Sub Teclado_virtual(sender As Object, e As EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Button4.Click, Button5.Click, Button6.Click, Button7.Click, Button8.Click, Button9.Click,
        Button10.Click, Button11.Click, Button12.Click, Button13.Click, Button14.Click, Button15.Click, Button16.Click, Button17.Click, Button18.Click, Button19.Click, Button20.Click, Button21.Click, Button22.Click,
        Button23.Click, Button24.Click, Button25.Click, Button26.Click

        Seleccionarletras(sender.text)

    End Sub
End Class
