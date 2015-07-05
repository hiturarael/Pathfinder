Imports System.Xml

Public Module ReadXML

    Sub readXML(ByVal book As String)
        Dim reader As XmlTextReader = New XmlTextReader(book)

        Do While (reader.Read())

        Loop
    End Sub

End Module
