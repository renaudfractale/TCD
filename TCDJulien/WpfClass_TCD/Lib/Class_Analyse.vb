Public Class Class_Analyse
    Property Liste_AllValue As New List(Of String)
    Property Liste_OcValue As New List(Of String)
    Property Liste_NumericValue As New List(Of String)
    Property Liste_OcAndNu As New List(Of Class_Analyse_Sub)


    Public Function Verification() As String
        Dim Liste As New List(Of String)
        Dim Dico As New Dictionary(Of String, Integer)

        For Each txt In Liste_AllValue
            If Not Dico.ContainsKey(txt) Then
                Dico.Add(txt, 0)
            End If
            Dico.Item(txt) += 1
        Next


        For Each txt In Liste_OcValue
            If Not Dico.ContainsKey(txt) Then
                Dico.Add(txt, 0)
            End If
            Dico.Item(txt) += 1
        Next



        For Each txt In Liste_NumericValue
            If Not Dico.ContainsKey(txt) Then
                Dico.Add(txt, 0)
            End If
            Dico.Item(txt) += 1
        Next

        For Each txt2 In Liste_OcAndNu
            If Not Dico.ContainsKey(txt2.NumericValue) Then
                Dico.Add(txt2.NumericValue, 0)
            End If
            If Not Dico.ContainsKey(txt2.OcValue) Then
                Dico.Add(txt2.OcValue, 0)
            End If
            Dico.Item(txt2.NumericValue) += 1
            Dico.Item(txt2.OcValue) += 1
        Next

        For Each KV In Dico
            If KV.Value >= 2 Then
                Liste.Add("Le champs '" + KV.Key + "'  est presents " + KV.Value.ToString + " fois")
            End If
        Next

        If Liste.Count = 0 Then
            Return ""
        Else
            Return Join(Liste.ToArray, ", ")
        End If

    End Function
End Class

Public Class Class_Analyse_Sub
    Property OcValue As String
    Property NumericValue As String

    Public Sub New()

    End Sub

    Public Sub New(OcValue As String, NumericValue As String)
        Me.OcValue = OcValue
        Me.NumericValue = NumericValue
    End Sub
End Class
Public Class Class_AnalyseV2s
    Inherits Dictionary(Of String, Class_AnalyseV2)
    Property ListeTilte As New List(Of String)
    Public Sub New(ListeTilte As List(Of String))
        Me.ListeTilte = ListeTilte
    End Sub
    Public Sub New()

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
    Public Sub New()

    End Sub
    Public Sub New(Dico As Dictionary(Of String, String), Analyse As Class_Analyse)

        For Each KV In Dico
            Dim Tilte As String = KV.Key
            Dim Value As String = KV.Value

            If Value Is Nothing Then Value = ""


            For Each K In Analyse.Liste_AllValue
                If K = Tilte Then
                    DicoKey.Add(Tilte, Value)
                    Exit For
                End If
            Next

            For Each K In Analyse.Liste_OcValue
                If K = Tilte Then
                    DicoValueOc.Add(Tilte, New Dictionary(Of String, Integer))
                    DicoValueOc.Item(K).Add(Value, 1)
                    Exit For
                End If
            Next

            For Each K In Analyse.Liste_NumericValue
                If K = Tilte Then
                    Try
                        DicoValueNum.Add(Tilte, New Dictionary(Of Integer, Integer))
                        DicoValueNum.Item(K).Add(CInt(Value), 1)
                    Catch ex As Exception

                    End Try

                    Exit For
                End If
            Next
        Next
    End Sub


    Public Function Key(liste As List(Of String)) As String

        Dim L = New List(Of String)
        For Each Tilte In liste
            L.Add(DicoKey.Item(Tilte))
        Next
        Return Join(L.ToArray, " | ")
    End Function


    Property DicoKey As New Dictionary(Of String, String)
    Property DicoValueNum As New Dictionary(Of String, Dictionary(Of Integer, Integer))
    Property DicoValueOc As New Dictionary(Of String, Dictionary(Of String, Integer))
    Property DicoValueOcNum As New Dictionary(Of Class_Analyse_Sub, Class_AnalyseV2_Sub)
    Public Sub Add(AnalyseV2 As Class_AnalyseV2)

        For Each KTilte In DicoValueOc.Keys
            Dim ListeKey = AnalyseV2.DicoValueOc.Item(KTilte).Keys.ToList
            ListeKey.AddRange(Me.DicoValueOc.Item(KTilte).Keys.ToList)
            ListeKey = ListeKey.Distinct.ToList

            For Each K In ListeKey

                Dim NB = 0
                If AnalyseV2.DicoValueOc.Item(KTilte).ContainsKey(K) Then
                    NB = AnalyseV2.DicoValueOc.Item(KTilte).Item(K)
                End If
                If DicoValueOc.Item(KTilte).ContainsKey(K) Then
                    DicoValueOc.Item(KTilte).Item(K) += NB
                Else
                    DicoValueOc.Item(KTilte).Add(K, NB)
                End If
            Next
        Next
        For Each KTilte In DicoValueNum.Keys
            Dim ListeKeyInt = AnalyseV2.DicoValueNum.Item(KTilte).Keys.ToList
            ListeKeyInt.AddRange(Me.DicoValueNum.Item(KTilte).Keys.ToList)
            ListeKeyInt = ListeKeyInt.Distinct.ToList

            For Each K In ListeKeyInt

                Dim NB = 0
                If AnalyseV2.DicoValueNum.Item(KTilte).ContainsKey(K) Then
                    NB = AnalyseV2.DicoValueNum.Item(KTilte).Item(K)
                End If
                If DicoValueNum.Item(KTilte).ContainsKey(K) Then
                    DicoValueNum.Item(KTilte).Item(K) += NB
                Else
                    DicoValueNum.Item(KTilte).Add(K, NB)
                End If
            Next
        Next
    End Sub



End Class

Public Class Class_AnalyseV2_Sub
    Property OcValue As New Dictionary(Of String, Integer)
    Property NumericValue As New Dictionary(Of Integer, Integer)
End Class

Public Class Class_AnalyseSave
    Property Alllines As Class_AnalyseV2s
    Property Tiltes As Class_Analyse
End Class