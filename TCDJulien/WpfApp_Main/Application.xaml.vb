Imports System.Threading

Class Application

    ' Les événements de niveau application, par exemple Startup, Exit et DispatcherUnhandledException
    ' peuvent être gérés dans ce fichier.
    Public Shared ParentWindowRef As New MainWindow()

    Public Shared Page_00 As New Page_00_Welcom()
    Public Shared Page_01 As New Page_01_SelectFile()
    Public Shared Page_02 As New Page_02_Analyse()

    Public Shared NoPage As Integer = 0

    Public Shared PathFile As String

    Public Shared MonThreadAnalyse As New Thread(AddressOf Module_RunTime.Analyse)
End Class
