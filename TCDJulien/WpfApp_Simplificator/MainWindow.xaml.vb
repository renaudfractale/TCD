Imports WpfClass_TCD
Imports System.IO
Class MainWindow
    Property ListeChamps As New Dictionary(Of String, Dictionary(Of String, Integer))
    Private Sub Button_Save_Click(sender As Object, e As RoutedEventArgs)

    End Sub

    Private Sub Button_AddChoix_Click(sender As Object, e As RoutedEventArgs)
        If ListeChamps.Count = 0 Then Exit Sub

        Dim UusCtrl As New UserControl_Simplificator(ListeChamps)

        StackPanel_Choix.Children.Add(UusCtrl)

    End Sub

    Private Sub TextBox_FileName_TextChanged(sender As Object, e As TextChangedEventArgs)
        Dim FileName = TextBox_FileName.Text
        If Not File.Exists(FileName) Then Exit Sub
        Dim ext = Path.GetExtension(FileName)
        If ext <> ".xlsx" Then Exit Sub

        ListeChamps = GetListeChamps(FileName)

    End Sub

    Private Sub Button_Go_Click(sender As Object, e As RoutedEventArgs)
        Dim FileName = TextBox_FileName.Text
        If Not File.Exists(FileName) Then Exit Sub
        Dim ext = Path.GetExtension(FileName)
        If ext <> ".xlsx" Then Exit Sub
        Simplifie(FileName, StackPanel_Choix)



    End Sub
End Class
