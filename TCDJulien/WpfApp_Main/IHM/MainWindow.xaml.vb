Class MainWindow

    Public Page_01 As Page_01_SelectFile
    Public Page_02 As Page_02_Analyse

    Public NoPage As Integer = 0
    Public NoPage_Old As Integer = -1

    Public PathFile As String

    Public Sub New()

        ' Cet appel est requis par le concepteur.
        InitializeComponent()

        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
        Page_01 = New Page_01_SelectFile(Me)
        Page_02 = New Page_02_Analyse(Me)

        Me.UpdatePage()

    End Sub
 
    Private Sub Button_Next_Click(sender As Object, e As RoutedEventArgs)
        Add()
    End Sub

    Private Sub Button_Previous_Click(sender As Object, e As RoutedEventArgs)
        Remove()
    End Sub

    Public Sub Add()
        NoPage_Old = NoPage
        NoPage += 1
        UpdatePage()
    End Sub


    Public Sub Remove()
        NoPage_Old = NoPage
        NoPage -= 1
        UpdatePage()
    End Sub

    Public Sub UpdatePage()
        Select Case NoPage
            Case 0
                Me.ParentFrame.Navigate(New Page_00_Welcom())
                Me.Button_Previous.Visibility = Visibility.Hidden
                Me.Button_Next.Visibility = Visibility.Visible
            Case 1
                Me.ParentFrame.Navigate(Page_01)
                Me.Button_Next.Visibility = Visibility.Hidden
                Me.Button_Previous.Visibility = Visibility.Visible
                PathFile = ""
            Case 2
                Page_02.UpdateShow()
                Me.ParentFrame.Navigate(Page_02)
                Me.Button_Next.Visibility = Visibility.Hidden
                Me.Button_Previous.Visibility = Visibility.Visible
                MsgBox(PathFile)
        End Select

        Me.ParentFrame.NavigationUIVisibility = NavigationUIVisibility.Hidden
    End Sub
End Class
