Module Race
    Sub setRaces()
        Dim races As ArrayList = XML.setRaceList

        AddRace.ddlRace.DataSource = races
    End Sub
End Module
