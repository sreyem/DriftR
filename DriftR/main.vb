
Imports core


Module main

    Sub main()

restart:

        fillTest()

        log.LogList.Clear()
        Console.Clear()
        mylog(LogTxtArray:=getStartInfo)

        showform = New frmPropGrid(
            class2Show:=test,
             classType:=test.GetType,
             restart:=True)


        Try

            If Not IsNothing(showform) Then

                With showform

                    .Width = 1200
                    .Height = 1300
                    .Text = test.GetType.ToString.Split(".").Last.ToUpper
                    If .ShowDialog() = Windows.Forms.DialogResult.Retry Then

                        showform.Close()
                        GoTo restart

                    End If

                End With


            End If

        Catch ex As Exception

        End Try


        End

    End Sub

    Public showform As frmPropGrid
    Public WithEvents test As New driftPercent

    Public Sub fillTest()

        With test

            .FOCUSswDriftCrop = eFOCUSswDriftCrop.PFE
            .noOfApplns = eNoOfApplns.one
            .rate = 1
            .FOCUSswWaterBody = eFOCUSswWaterBody.ditch

        End With

    End Sub


End Module
