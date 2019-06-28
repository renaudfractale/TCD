Public Class UserControl_Simplificator
    Property ListeChamps As New Dictionary(Of String, Dictionary(Of String, Integer))
    Public Sub New()

        ' Cet appel est requis par le concepteur.
        InitializeComponent()

        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().

    End Sub

    Public Sub New(ListeChamps As Dictionary(Of String, Dictionary(Of String, Integer)))

        ' Cet appel est requis par le concepteur.
        InitializeComponent()

        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
        Me.ListeChamps = ListeChamps
        Dim ListeValue = ListeChamps.Keys.ToList
        ListeValue.Sort()

        For Each Champs In ListeValue
            Me.ComboBox_Champs.Items.Add(Champs)
        Next
        Me.ComboBox_Choix.Items.Add("")
        Me.ComboBox_Choix.Items.Add("All Value")
        Me.ComboBox_Choix.Items.Add("Ocurence Value")
        Me.ComboBox_Choix.Items.Add("Numeric Value")


    End Sub
    Private Sub Button_Close_Click(sender As Object, e As RoutedEventArgs)
        Me.Visibility = Visibility.Collapsed
    End Sub

    Private Sub TextBox_Tilte_TextChanged(sender As Object, e As TextChangedEventArgs)

    End Sub

    Private Sub ComboBox_Champs_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)

    End Sub

    Public Function ExportTiClass() As Class_Simplificator
        Dim Simplificator As New Class_Simplificator
        Simplificator.Tilte = TextBox_Tilte.Text
        Dim Champs = CType(ComboBox_Champs.SelectedItem, String)
        If Champs Is Nothing Then
            Simplificator.ChampsName = ""
        Else
            Simplificator.ChampsName = Champs
        End If
        Dim Choix = CType(ComboBox_Choix.SelectedIndex, String)
        If Choix Is Nothing Then
            Simplificator.ChoixSelected = Class_Simplificator.Choix.None
        Else
            Select Case Choix
                Case ""
                    Simplificator.ChoixSelected = Class_Simplificator.Choix.None
                Case "All Value"
                    Simplificator.ChoixSelected = Class_Simplificator.Choix.All_Values
                Case "Ocurence Value"
                    Simplificator.ChoixSelected = Class_Simplificator.Choix.Ocurence_Value
                Case "Numeric Value"
                    Simplificator.ChoixSelected = Class_Simplificator.Choix.NumericValue

            End Select

        End If


        Return Simplificator
    End Function
End Class
