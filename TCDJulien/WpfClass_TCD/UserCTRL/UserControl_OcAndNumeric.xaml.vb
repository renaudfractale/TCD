Public Class UserControl_OcAndNumeric
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
            Me.ComboBox_Champs_Oc.Items.Add(Champs)
            Me.ComboBox_Champs_Nu.Items.Add(Champs)
        Next
    End Sub
End Class
