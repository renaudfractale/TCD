Public Class UserControl_Filter
    Property ListeChamps As New Dictionary(Of String, Dictionary(Of String, Integer))
    Public Sub New()

        ' Cet appel est requis par le concepteur.
        InitializeComponent()

        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
        TextBox_Name_TextChanged(Nothing, Nothing)

    End Sub

    Public Sub New(ListeChamps As Dictionary(Of String, Dictionary(Of String, Integer)))

        ' Cet appel est requis par le concepteur.
        InitializeComponent()

        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
        Me.ListeChamps = ListeChamps
        Dim ListeValue = ListeChamps.Keys.ToList
        ListeValue.Sort()

        For Each Champs In ListeValue
            Me.ComboBox_KEYS_Select.Items.Add(Champs)
        Next

        Me.ComboBox_Action.Items.Add("Remove")
        Me.ComboBox_Action.Items.Add("Add")
        TextBox_Name_TextChanged(Nothing, Nothing)
    End Sub
    Private Sub Button_Close_Click(sender As Object, e As RoutedEventArgs)
        Me.Visibility = Visibility.Collapsed
    End Sub

    Private Sub TextBox_Name_TextChanged(sender As Object, e As TextChangedEventArgs)
        If TextBox_Name.Text.Length > 0 Then
            ComboBox_KEYS_Select.IsEnabled = True
        Else
            ComboBox_KEYS_Select.IsEnabled = False
        End If
    End Sub

    Private Sub ComboBox_KEYS_Select_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        ComboBox_KEYS_Values.IsEnabled = True
        ComboBox_KEYS_Values.Items.Clear()
        If ComboBox_KEYS_Select.SelectedItem Is Nothing Then
            ComboBox_KEYS_Values.IsEnabled = False
        Else
            ComboBox_KEYS_Values.IsEnabled = True
            Dim key = ComboBox_KEYS_Select.SelectedItem.ToString
            Dim ListeValue = ListeChamps.Item(key).Keys.ToList
            ListeValue.Sort()

            ComboBox_KEYS_Values.Items.Add("*")
            For Each Value In ListeValue
                ComboBox_KEYS_Values.Items.Add(Value)
            Next

        End If
    End Sub

    Private Sub ComboBox_KEYS_Values_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        ComboBox_Filter_Select.IsEnabled = True
        ComboBox_Filter_Select.Items.Clear()
        If ComboBox_KEYS_Select.SelectedItem Is Nothing Then
            ComboBox_Filter_Select.IsEnabled = False
        Else

            Dim ListeValue = ListeChamps.Keys.ToList
            ListeValue.Sort()


            For Each Value In ListeValue
                If Value = ComboBox_KEYS_Select.SelectedItem.ToString Then Continue For
                ComboBox_Filter_Select.Items.Add(Value)
            Next

        End If
    End Sub

    Private Sub ComboBox_Filter_Select_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        ComboBox_Filter_Comp.IsEnabled = True
        ComboBox_Filter_Comp.Items.Clear()
        If ComboBox_Filter_Select.SelectedItem Is Nothing Then
            ComboBox_Filter_Comp.IsEnabled = False
        Else
            Dim Values = {"=", "<>", ">", ">=", "<", "<="}
            For Each Value In Values
                ComboBox_Filter_Comp.Items.Add(Value)
            Next
        End If
    End Sub

    Private Sub ComboBox_Filter_Comp_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        ComboBox_Filter_Value.IsEnabled = True
        ComboBox_Filter_Value.Items.Clear()
        If ComboBox_Filter_Comp.SelectedItem Is Nothing Then
            ComboBox_Filter_Value.IsEnabled = False
        Else
            Dim key = ComboBox_Filter_Select.SelectedItem.ToString
            Dim ListeValue = ListeChamps.Item(key).Keys.ToList
            ListeValue.Sort()


            For Each Value In ListeValue
                ComboBox_Filter_Value.Items.Add(Value)
            Next

        End If
    End Sub

    Private Sub ComboBox_Filter_Value_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        ComboBox_Action.IsEnabled = True
        ComboBox_Action.Items.Clear()
        If ComboBox_Filter_Value.SelectedItem Is Nothing Then
            ComboBox_Action.IsEnabled = False
        Else
            Dim Values = {"Add", "Remove"}
            For Each Value In Values
                ComboBox_Action.Items.Add(Value)
            Next
        End If
    End Sub
End Class
