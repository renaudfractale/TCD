Imports System.Threading
Class Page_01_SelectFile
    Private Main As MainWindow
    Public Sub New(Main As MainWindow)

        ' Cet appel est requis par le concepteur.
        InitializeComponent()

        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
        Me.Main = Main
    End Sub

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
            Main.PathFile = openFileDlg.FileName
            Main.Add()
        Else
            Main.PathFile = ""
        End If

    End Sub


End Class
