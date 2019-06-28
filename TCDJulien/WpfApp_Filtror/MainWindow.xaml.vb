Imports WpfClass_TCD
Imports System.IO
Class MainWindow

    Property ListeChamps As New Dictionary(Of String, Dictionary(Of String, Integer))
    Private Sub Button_AddFilter_Click(sender As Object, e As RoutedEventArgs)
        If ListeChamps.Count = 0 Then Exit Sub

        Dim UusCtrl As New UserControl_Filter_KEY(ListeChamps)

        StackPanel_Filter.Children.Add(UusCtrl)

    End Sub

    Private Sub TextBox_FileName_TextChanged(sender As Object, e As TextChangedEventArgs)
        Dim FileName = TextBox_FileName.Text
        If Not File.Exists(FileName) Then Exit Sub
        Dim ext = Path.GetExtension(FileName)
        If ext <> ".xlsx" Then Exit Sub

        ListeChamps = GetListeChamps(FileName)

    End Sub

    Private Sub Button_Go_Click(sender As Object, e As RoutedEventArgs)
        Dim FileName = Me.TextBox_FileName.Text
        If Not File.Exists(FileName) Then Exit Sub
        Dim ext = Path.GetExtension(FileName)
        If ext <> ".xlsx" Then Exit Sub
        Export(FileName, StackPanel_Filter)
    End Sub

    Private Sub Button_AddFilter_LIGHT_Click(sender As Object, e As RoutedEventArgs)
        If ListeChamps.Count = 0 Then Exit Sub

        Dim UusCtrl As New UserControl_Filter_light(ListeChamps)

        StackPanel_Filter.Children.Add(UusCtrl)

    End Sub
End Class
