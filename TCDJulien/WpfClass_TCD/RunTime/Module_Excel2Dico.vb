Imports Microsoft.Win32
Imports System.IO
Public Module Module_Excel2Dico
    Property MyExcel As New Microsoft.Office.Interop.Excel.Application
    Public Function GetListeChamps(FileName As String) As Dictionary(Of String, Dictionary(Of String, Integer))
        Dim ListeChamps As New Dictionary(Of String, Dictionary(Of String, Integer))
        Dim MyFile = MyExcel.Workbooks.Open(FileName, False, True)
        Dim MySheet = CType(MyFile.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        Dim LastLine = Module_FunctionsExcel.GetLastLigne(MySheet) - 1
        Dim LastCol = Module_FunctionsExcel.GetlasteCol(MySheet) - 1
        Dim Pose = "A1:" + Module_FunctionsExcel.xlCol(LastCol) + LastLine.ToString

        Dim Array = CType(MySheet.Range(Pose).Value2, Array)
        ListeChamps = New Dictionary(Of String, Dictionary(Of String, Integer))

        For NoCol = 1 To LastCol + 1
            Dim value = CType(Array.GetValue(1, NoCol), String)
            ListeChamps.Add(value, New Dictionary(Of String, Integer))

            For i As Integer = 2 To LastLine
                Dim value2 = CType(Array.GetValue(i, NoCol), String)
                If value2 Is Nothing Then value2 = ""
                If Not ListeChamps.Item(value).ContainsKey(value2) Then
                    ListeChamps.Item(value).Add(value2, 0)
                End If
                ListeChamps.Item(value).Item(value2) += 1
            Next

        Next

        MyFile.Close(False)
        Return ListeChamps
    End Function

    Public Sub Export(FileName As String, StackPanel_Filter As StackPanel)

        Dim MyFile = MyExcel.Workbooks.Open(FileName, False, True)
        Dim MySheet = CType(MyFile.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        Dim LastLine = Module_FunctionsExcel.GetLastLigne(MySheet) - 1
        Dim LastCol = Module_FunctionsExcel.GetlasteCol(MySheet) - 1
        Dim Pose = "A1:" + Module_FunctionsExcel.xlCol(LastCol) + LastLine.ToString

        Dim Array = CType(MySheet.Range(Pose).Value2, Array)

        Dim Dico As New Dictionary(Of Integer, Dictionary(Of String, String))

        For NoLigne = 2 To LastLine
            Dico.Add(NoLigne, New Dictionary(Of String, String))
            For NoCol As Integer = 1 To LastCol + 1
                Dim value = CType(Array.GetValue(NoLigne, NoCol), String)
                Dim Tilte = CType(Array.GetValue(1, NoCol), String)

                If value Is Nothing Then value = ""
                Dico.Item(NoLigne).Add(Tilte, value)
            Next
        Next

        MyFile.Close(False)



        For Each CtrtObject As Object In StackPanel_Filter.Children

            If CtrtObject.GetType() = GetType(UserControl_Filter_KEY) Then
                Dim Ctrt = DirectCast(CtrtObject, UserControl_Filter_KEY)
                If Ctrt.Visibility = Visibility.Collapsed Then Continue For
                Dim Dico2 As New Dictionary(Of String, Dictionary(Of Integer, Dictionary(Of String, String)))
                Dim Key = Ctrt.ComboBox_KEYS_Select.SelectedItem.ToString
                Dim Value = Ctrt.ComboBox_KEYS_Values.SelectedItem.ToString
                If Value = "*" Then Value = Nothing
                For Each KV In Dico
                    Dim V = KV.Value.Item(Key)
                    If Value Is Nothing Then
                        If Not Dico2.ContainsKey(V) Then
                            Dico2.Add(V, New Dictionary(Of Integer, Dictionary(Of String, String)))
                        End If
                        Dico2.Item(V).Add(KV.Key, KV.Value)
                    ElseIf Value = V Then
                        If Not Dico2.ContainsKey(V) Then
                            Dico2.Add(V, New Dictionary(Of Integer, Dictionary(Of String, String)))
                        End If
                        Dico2.Item(V).Add(KV.Key, KV.Value)
                    End If
                Next

                Dim ListeOK As New List(Of String)

                Dim KeyF = Ctrt.ComboBox_Filter_Select.SelectedItem.ToString
                Dim ValueF = Ctrt.ComboBox_Filter_Value.SelectedItem.ToString
                Dim CompF = Ctrt.ComboBox_Filter_Comp.SelectedItem.ToString
                For Each KV In Dico2
                    Dim state As Boolean = False
                    Dim StateMAster As Boolean = False
                    If CompF <> "<>" Then
                        For Each NoLigne In KV.Value.Keys.ToArray
                            Dim V = KV.Value.Item(NoLigne).Item(KeyF)
                            Select Case CompF
                                Case "="
                                    state = V = ValueF
                                Case ">"
                                    If V = "" Or ValueF = "" Then Exit Select
                                    state = CInt(V) > CInt(ValueF)
                                Case ">="
                                    If V = "" Or ValueF = "" Then Exit Select
                                    state = CInt(V) >= CInt(ValueF)
                                Case "<"
                                    If V = "" Or ValueF = "" Then Exit Select
                                    state = CInt(V) < CInt(ValueF)
                                Case "<="
                                    If V = "" Or ValueF = "" Then Exit Select
                                    state = CInt(V) <= CInt(ValueF)
                            End Select
                            If state Then Exit For
                        Next
                    Else
                        For Each NoLigne In KV.Value.Keys.ToArray
                            Dim V = KV.Value.Item(NoLigne).Item(KeyF)

                            state = V = ValueF

                            If state Then Exit For
                        Next
                        state = Not state
                    End If

                    If state Then ListeOK.Add(KV.Key)

                Next
                Dim Action = Ctrt.ComboBox_Action.SelectedItem.ToString
                Dico = New Dictionary(Of Integer, Dictionary(Of String, String))

                For Each KV In Dico2
                    If Action = "Add" Then
                        For Each K In ListeOK
                            If KV.Key = K Then
                                For Each D In KV.Value
                                    Dico.Add(D.Key, D.Value)
                                Next
                            End If
                        Next
                    ElseIf Action = "Remove" Then
                        Dim state = True
                        For Each K In ListeOK
                            If KV.Key = K Then
                                state = False
                                Exit For
                            End If
                        Next
                        If state Then
                            For Each D In KV.Value
                                Dico.Add(D.Key, D.Value)
                            Next
                        End If
                    Else
                        For Each D In KV.Value
                            Dico.Add(D.Key, D.Value)
                        Next
                    End If
                Next
            ElseIf CtrtObject.GetType() = GetType(UserControl_Filter_light) Then
                Dim Ctrt = DirectCast(CtrtObject, UserControl_Filter_light)
                If Ctrt.Visibility = Visibility.Collapsed Then Continue For
                Dim Key = Ctrt.ComboBox_KEYS_Select.SelectedItem.ToString
                Dim Value = Ctrt.ComboBox_KEYS_Values.SelectedItem.ToString

                Dim KeyF = Ctrt.ComboBox_Filter_Select.SelectedItem.ToString
                Dim ValueF = Ctrt.ComboBox_Filter_Value.SelectedItem.ToString
                Dim CompF = Ctrt.ComboBox_Filter_Comp.SelectedItem.ToString
                Dim Dico_New As New Dictionary(Of Integer, Dictionary(Of String, String))
                For Each KV In Dico
                    If KV.Value.Item(Key) = Value Then
                        Dim Dic = KV.Value
                        Dim V = Dic.Item(KeyF)
                        Dim state As Boolean = False
                        Select Case CompF
                            Case "="
                                state = V = ValueF
                            Case "<>"
                                state = V <> ValueF
                            Case ">"
                                If V = "" Or ValueF = "" Then Exit Select
                                state = CInt(V) > CInt(ValueF)
                            Case ">="
                                If V = "" Or ValueF = "" Then Exit Select
                                state = CInt(V) >= CInt(ValueF)
                            Case "<"
                                If V = "" Or ValueF = "" Then Exit Select
                                state = CInt(V) < CInt(ValueF)
                            Case "<="
                                If V = "" Or ValueF = "" Then Exit Select
                                state = CInt(V) <= CInt(ValueF)
                        End Select

                        Dim Action = Ctrt.ComboBox_Action.SelectedItem.ToString
                        If Action = "Add" Then
                            If state = True Then Dico_New.Add(KV.Key, KV.Value)
                        ElseIf Action = "Remove" Then
                            If state = False Then Dico_New.Add(KV.Key, KV.Value)
                        Else
                            Dico_New.Add(KV.Key, KV.Value)
                        End If
                    Else
                        Dico_New.Add(KV.Key, KV.Value)
                    End If
                Next
                Dico = Dico_New
            End If
        Next


        Dim saveFileDialog = New SaveFileDialog()
        If saveFileDialog.ShowDialog() = True Then
            Dim MyFileS = MyExcel.Workbooks.Add
            MyFileS.SaveAs(saveFileDialog.FileName)
            Dim MySheetS = CType(MyFileS.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            Dim Tiltes = Dico.First.Value.Keys.ToArray

            Dim LastColS = Tiltes.Length - 1
            Dim PoseTilte = "A1:" + Module_FunctionsExcel.xlCol(LastColS) + "1"
            MySheetS.Range(PoseTilte).Value2 = Tiltes
            Dim Line As Integer = 2
            For Each KV In Dico
                Dim listevalues As New List(Of String)
                For Each Tilte In Tiltes
                    listevalues.Add(KV.Value.Item(Tilte))
                Next
                Dim PoseV = "A" + Line.ToString + ":" + Module_FunctionsExcel.xlCol(LastColS) + Line.ToString
                MySheetS.Range(PoseV).Value2 = listevalues.ToArray
                Line += 1
                '   MyFileS.Save()
            Next

            MyFileS.Save()
            MyFileS.Close(True)
            'saveFileDialog.FileName
        End If



    End Sub


    Public Sub Simplifie(FileName As String, StackPanel_Choix As StackPanel)

        Dim MyFile = MyExcel.Workbooks.Open(FileName, False, True)
        Dim MySheet = CType(MyFile.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        Dim LastLine = Module_FunctionsExcel.GetLastLigne(MySheet) - 1
        Dim LastCol = Module_FunctionsExcel.GetlasteCol(MySheet) - 1
        Dim Pose = "A1:" + Module_FunctionsExcel.xlCol(LastCol) + LastLine.ToString

        Dim Array = CType(MySheet.Range(Pose).Value2, Array)

        Dim Dico As New Dictionary(Of Integer, Dictionary(Of String, String))

        For NoLigne = 2 To LastLine
            Dico.Add(NoLigne, New Dictionary(Of String, String))
            For NoCol As Integer = 1 To LastCol + 1
                Dim value = CType(Array.GetValue(NoLigne, NoCol), String)
                Dim Tilte = CType(Array.GetValue(1, NoCol), String)

                If value Is Nothing Then value = ""
                Dico.Item(NoLigne).Add(Tilte, value)
            Next
        Next

        MyFile.Close(False)
        Dim Analyse As New Class_Analyse
        For Each CtrlObject As Object In StackPanel_Choix.Children

            If CtrlObject.GetType() = GetType(UserControl_Simplificator) Then
                Dim Ctrl = DirectCast(CtrlObject, UserControl_Simplificator)

                If Ctrl.Visibility = Visibility.Collapsed Then Continue For

                Dim Champs = DirectCast(Ctrl.ComboBox_Champs.SelectedItem, String)
                If Champs Is Nothing Then Continue For
                Dim Choix = DirectCast(Ctrl.ComboBox_Choix.SelectedItem, String)
                If Choix Is Nothing OrElse Choix = "" Then Continue For

                If Choix = "All Value" Then
                    Analyse.Liste_AllValue.Add(Champs)
                ElseIf Choix = "Ocurence Value" Then
                    Analyse.Liste_OcValue.Add(Champs)
                Else 'Numeric value
                    Analyse.Liste_NumericValue.Add(Champs)
                End If
            ElseIf CtrlObject.GetType() = GetType(UserControl_OcAndNumeric) Then
                Dim Ctrl = DirectCast(CtrlObject, UserControl_OcAndNumeric)

                If Ctrl.Visibility = Visibility.Collapsed Then Continue For
                Dim ChampsOc = DirectCast(Ctrl.ComboBox_Champs_Oc.SelectedItem, String)
                If ChampsOc Is Nothing Then Continue For
                Dim ChampsNu = DirectCast(Ctrl.ComboBox_Champs_Nu.SelectedItem, String)
                If ChampsNu Is Nothing Then Continue For
                Analyse.Liste_OcAndNu.Add(New Class_Analyse_Sub(ChampsOc, ChampsNu))
            End If
        Next
        If Analyse.Liste_AllValue.Count = 0 Then Exit Sub
        If Analyse.Liste_OcValue.Count + Analyse.Liste_NumericValue.Count = 0 Then Exit Sub

        If Analyse.Verification.Length > 0 Then
            MsgBox(Analyse.Verification)
            Exit Sub
        End If

        Dim AnalyseV2s As New Class_AnalyseV2s(Analyse.Liste_AllValue)


        For Each KV In Dico
            Dim AnalyseV2 As New Class_AnalyseV2(KV.Value, Analyse)
            AnalyseV2s.Add(AnalyseV2)
        Next


        Dim saveFileDialog = New SaveFileDialog()
        saveFileDialog.Filter = "Analyse|*.jsonA"
        If saveFileDialog.ShowDialog() = True Then
            If File.Exists(saveFileDialog.FileName) Then File.Delete(saveFileDialog.FileName)

            Dim Class_AnalyseSave As New Class_AnalyseSave
            Class_AnalyseSave.Alllines = AnalyseV2s
            Class_AnalyseSave.Tiltes = Analyse

            Class_SerialisationAndUnSerialisation.Serialisation(Of Class_AnalyseSave)(Class_AnalyseSave, saveFileDialog.FileName)
        End If

    End Sub

End Module
