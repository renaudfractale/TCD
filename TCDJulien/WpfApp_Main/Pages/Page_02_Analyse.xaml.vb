Class Page_02_Analyse
    Private Main As MainWindow
    Public Sub New(Main As MainWindow)

        ' Cet appel est requis par le concepteur.
        InitializeComponent()

        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
        Me.Main = Main
    End Sub
    Public Sub UpdateShow()
        If Main.PathFile.Length <> 0 Then
            Me.Border_Info.Visibility = Visibility.Visible
            Me.TextBlock_PathDir.Text = System.IO.Path.GetDirectoryName(Main.PathFile)
            Me.TextBlock_NameFile.Text = System.IO.Path.GetFileName(Main.PathFile)
        Else
            Me.Border_Info.Visibility = Visibility.Collapsed
            Me.TextBlock_PathDir.Text = ""
            Me.TextBlock_NameFile.Text = ""
        End If
    End Sub
End Class
