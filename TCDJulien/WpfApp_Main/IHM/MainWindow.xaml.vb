Class MainWindow
    Public Sub New()

        ' Cet appel est requis par le concepteur.
        InitializeComponent()

        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
        Me.UpdatePage()
    End Sub

    Private Sub Button_Next_Click(sender As Object, e As RoutedEventArgs)
        Application.NoPage += 1
        Me.UpdatePage()
    End Sub

    Private Sub Button_Previous_Click(sender As Object, e As RoutedEventArgs)
        Application.NoPage -= 1
        Me.UpdatePage()
    End Sub

    Public Sub UpdatePage()
        Select Case Application.NoPage
            Case 0
                Me.ParentFrame.Navigate(Application.Page_00)
                Me.Button_Previous.Visibility = Visibility.Hidden
                Me.Button_Next.Visibility = Visibility.Visible
            Case 1
                Me.ParentFrame.Navigate(Application.Page_01)
                Me.Button_Next.Visibility = Visibility.Hidden
                Me.Button_Previous.Visibility = Visibility.Visible
                Application.PathFile = ""
            Case 2
                Me.ParentFrame.Navigate(Application.Page_02)
                Me.Button_Next.Visibility = Visibility.Hidden
                Me.Button_Previous.Visibility = Visibility.Visible
                MsgBox(Application.PathFile)
        End Select

        Me.ParentFrame.NavigationUIVisibility = NavigationUIVisibility.Hidden
    End Sub
End Class
