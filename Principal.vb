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

        For Each Grupo In Controls
            If TypeOf Grupo Is TextBox Then
                If Grupo.text = e.KeyCode.ToString Then
                    Grupo.forecolor = Color.Black

                End If
            End If

            If TypeOf Grupo Is Button Then

                If Grupo.text = e.KeyCode.ToString Then

                    Grupo.enabled = False

                End If
            End If

        Next

    End Sub
End Class
