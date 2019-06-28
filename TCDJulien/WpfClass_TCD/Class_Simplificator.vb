Public Class Class_Simplificator
    Property Tilte As String
    Property ChampsName As String

    Property ChoixSelected As Choix

    Enum Choix
        None
        All_Values
        Ocurence_Value
        NumericValue
    End Enum
End Class



Public Class Class_Simplificators
    Inherits Dictionary(Of Integer, Class_Simplificator)
    Property NameProject As String
End Class