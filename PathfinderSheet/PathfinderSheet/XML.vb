﻿'Option Strict On
Imports Excel = Microsoft.Office.Interop.Excel

Module XML
    Public APP As New Excel.Application
    Public workbook As Excel.Workbook = APP.Workbooks.Open(CurDir() & "\data\Pathfinder Matrix.xlsx")

    Function setRaceList()
        Dim worksheet As New Excel.Worksheet
        Dim raceID As Integer = 0
        Dim subraceID As Integer = 0
        Dim raceName As String = ""
        Dim x As Integer = 0
        Dim y As Integer = 0
        Dim races As New ArrayList

        worksheet = CType(workbook.Worksheets("Race"), Excel.Worksheet)
        For x = 2 To worksheet.UsedRange.Rows.Count
            If worksheet.Cells(x, 2).value = 0 Then
                Dim race As String = worksheet.Cells(x, 3).value
                races.Add(race)
            End If
        Next

        Return races
    End Function

    Function setToNextLevel(ByVal level As String, ByVal expSpeed As String)
        Dim worksheet As New Excel.Worksheet
        Dim exp As String = ""
        Dim column As Integer = CInt(getColumnOfHeader(expSpeed, "Levels"))
        Dim row As Integer = 2
        worksheet = CType(workbook.Worksheets("Levels"), Excel.Worksheet)

        Do Until exp <> "" Or CStr(row) = ""

            If row = CInt(level) + 2 And level <> "20" Then
                exp = "0"
            ElseIf level = "20" Then
                exp = "0"
            End If

            row += 1
        Loop

        Return exp
    End Function

    Function setCurrentExp(ByVal level As String, ByVal expspeed As String)
        Dim worksheet As New Excel.Worksheet
        Dim exp As String = ""
        Dim column As Integer = CInt(getColumnOfHeader(expspeed, "Levels"))
        Dim row As Integer = 2
        worksheet = CType(workbook.Worksheets("Levels"), Excel.Worksheet)

        Do Until exp <> "" Or CStr(row) = ""

            If row = CInt(level) + 1 And level <> "20" Then
                exp = worksheet.Cells(row, column).value
            ElseIf level = "20" Then
                exp = worksheet.Cells(21, column).value
            End If

            row += 1
        Loop

        Return exp
    End Function

    Function getColumnOfHeader(ByVal header As String, ByVal worksheetName As String)
        Dim worksheet As New Excel.Worksheet
        Dim column As String = "0"
        Dim x As Integer = 1
        worksheet = CType(workbook.Worksheets(worksheetName), Excel.Worksheet)

        Do Until column <> "0"
            Dim colName As String = worksheet.Cells(1, x).value

            If colName = header Then
                column = x
            ElseIf colName = "" Then
                Exit Do
            End If

            x += 1
        Loop

        Return column
    End Function
End Module
