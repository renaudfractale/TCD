Imports System.Threading
Class Page_01_SelectFile
    Private Sub Page_Loaded(sender As Object, e As RoutedEventArgs)

    End Sub

    Private Sub Button_OpenFile_Click(sender As Object, e As RoutedEventArgs)
        Dim openFileDlg = New Microsoft.Win32.OpenFileDialog()
        'Set filter for file extension And default file extension  
        openFileDlg.DefaultExt = ".xlsx"
        openFileDlg.Filter = "Excel (.xlsx)|*.xlsx"
        'Set initial directory    
        openFileDlg.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments


        If openFileDlg.ShowDialog() = True Then
            Application.PathFile = openFileDlg.FileName

            Application.MonThreadAnalyse.Start()

        End If
    End Sub
End Class
