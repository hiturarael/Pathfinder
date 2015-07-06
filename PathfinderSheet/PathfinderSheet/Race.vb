Module Race
    Sub setRaces()
        Dim races As ArrayList = XML.setRaceList

        AddRace.ddlRace.Items.Clear()
        For Each strRace In races
            AddRace.ddlRace.Items.Add(strRace)
        Next
    End Sub
End Module
