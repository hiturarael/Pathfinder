Public Class Dice

    Private Function rollDice()
        Dim stat As Integer
        Dim d6(4) As Integer
        Dim lower, upper As Integer
        lower = 1
        upper = 6

        For counter As Integer = 0 To 4
            d6(counter) = CInt(Int((upper - lower + 1) * Rnd() + lower))

            If d6(counter) = 1 Then
                Do While d6(counter) <= 1
                    d6(counter) = CInt(Int((upper - lower + 1) * Rnd() + lower))
                Loop
            End If

        Next

        Array.Sort(d6)

        d6(0) = CInt(Int((upper - lower + 1) * Rnd() + lower))

        Array.Sort(d6)

        stat = d6(1) + d6(2) + d6(3)
        ' MessageBox.Show(d6(0) & " " & d6(1) & " " & d6(2) & " " & d6(3))
        Return stat
    End Function

End Class
