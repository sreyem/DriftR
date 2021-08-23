
Imports core


Module main

    Sub main()


        mylog(LogTxtArray:=getStartInfo)

        showform = New frmPropGrid(
            class2Show:=test,
             classType:=test.GetType)


        Try

            If Not IsNothing(showform) Then

                With showform

                    .Width = 1200
                    .Height = 1300
                    .Text = test.GetType.ToString.Split(".").Last.ToUpper
                    .ShowDialog()

                End With


            End If

        Catch ex As Exception

        End Try


        End

    End Sub

    Public showform As frmPropGrid
    Public WithEvents test As New driftPercent

End Module
