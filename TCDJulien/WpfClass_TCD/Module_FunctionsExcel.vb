Public Module Module_FunctionsExcel
    Public Function GetNoCol(ByRef sheet As Microsoft.Office.Interop.Excel.Worksheet, key As String, Optional NoLigne As Integer = 1) As Integer
        For i As Integer = 0 To 1000
            Dim champs As String = CType(sheet.Range(xlCol(i) + NoLigne.ToString).Value2, String)
            If champs IsNot Nothing AndAlso key.ToUpperInvariant = champs.ToUpperInvariant Then
                Return i
            End If
        Next
        Return -1
    End Function

    Public Function GetlasteCol(ByRef sheet As Microsoft.Office.Interop.Excel.Worksheet, Optional NoLigne As Integer = 1) As Integer
        For i As Integer = 0 To 1000
            Dim champs As Object = sheet.Range(xlCol(i) + NoLigne.ToString).Value2
            If champs Is Nothing Then
                Return i
            End If
        Next
        Return -1
    End Function
    Public Function GetLastLigne(Sheet As Microsoft.Office.Interop.Excel.Worksheet, Optional Start As Integer = 1) As Integer
        For i As Integer = Start To 1000000

            Dim Line2 = Sheet.Range("A" + i.ToString + ":G" + i.ToString).Value2
            Dim Line2Array = CType(Line2, Array)
            Dim state As Boolean = False
            For Each ItemArray In Line2Array
                If ItemArray IsNot Nothing Then
                    state = True
                    Exit For
                End If
            Next

            If state = False Then
                Return i
            End If
        Next
        Return 1000000
    End Function


    Public Function xlCol(ByVal col As Integer) As String
        'col -= 1 'Uncomment this line if you are using 1-based column indices
        Dim s As String = ""
        If col < 0 Or col > 16383 Then
            Throw New ArgumentException(String.Format("{0} is an invalid column", col), "col")
        End If
        If col >= 26 ^ 2 Then
            s = Chr(64 + (col \ 26 \ 26) Mod 26)
        End If
        If col >= 26 Then
            s &= Chr(64 + (col \ 26) Mod 26)
        End If
        s &= Chr(65 + (col Mod 26))
        Return s
    End Function
End Module

