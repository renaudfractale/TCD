Public Class Class_Analyse
    Property Liste_AllValue As New List(Of String)
    Property Liste_OcValue As New List(Of String)
    Property Liste_NumericValue As New List(Of String)

End Class
Public Class Class_AnalyseV2s
    Inherits Dictionary(Of String, Class_AnalyseV2)
    Property ListeTilte As New List(Of String)
    Public Sub New(ListeTilte As List(Of String))
        Me.ListeTilte = ListeTilte
    End Sub

    Public Overloads Sub Add(AnalyseV2 As Class_AnalyseV2)
        If Me.ContainsKey(AnalyseV2.Key(ListeTilte)) Then
            Me.Item(AnalyseV2.Key(ListeTilte)).Add(AnalyseV2)
        Else
            Me.Add(AnalyseV2.Key(ListeTilte), AnalyseV2)
        End If
    End Sub

End Class
Public Class Class_AnalyseV2
    Public Function Key(liste As List(Of String)) As String

        Dim L = New List(Of String)
        For Each Tilte In liste
            L.Add(DicoKey.Item(Tilte))
        Next
        Return Join(L.ToArray, " | ")
    End Function


    Property DicoKey As New Dictionary(Of String, String)
    Property DicoValueNum As New Dictionary(Of Integer, Integer)
    Property DicoValueOc As New Dictionary(Of String, Integer)

    Public Sub Add(AnalyseV2 As Class_AnalyseV2)
        Dim ListeKey = AnalyseV2.DicoValueOc.Keys.ToList
        ListeKey.AddRange(Me.DicoValueOc.Keys.ToList)
        ListeKey = ListeKey.Distinct.ToList

        fore
    End Sub


End Class
