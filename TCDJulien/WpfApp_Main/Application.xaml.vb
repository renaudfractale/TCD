Imports System.Threading
Imports System.Windows.Threading

Class Application
    Public Sub New()

    End Sub


    ' Les événements de niveau application, par exemple Startup, Exit et DispatcherUnhandledException
    ' peuvent être gérés dans ce fichier.
    Public Shared ParentWindowRef As New MainWindow()



    Public Shared MonThreadAnalyse As New Thread(AddressOf Module_RunTime.Analyse)






End Class
