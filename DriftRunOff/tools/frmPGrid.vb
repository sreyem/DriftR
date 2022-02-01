

Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Runtime.Serialization
Imports System.Text
Imports System.Web.Script.Serialization
Imports System.Windows.Forms
Imports System.Xml.Serialization

Partial Class frmPropGrid

    Inherits Form

    Public Sub New()
        InitializeComponent()
    End Sub


    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Friend WithEvents tsmiRestart As ToolStripMenuItem
    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.MenuStrip = New System.Windows.Forms.MenuStrip()
        Me.tsmiFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLoad = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiSave = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiRestart = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiGrid = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiCollapseAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiCollapseStd = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiExpandAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiRefresh = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiEnlargeLeft = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.lblStatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.progressBar = New System.Windows.Forms.ToolStripProgressBar()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.propGridMain = New System.Windows.Forms.PropertyGrid()
        Me.propGridDetails = New System.Windows.Forms.PropertyGrid()
        Me.MenuStrip.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip
        '
        Me.MenuStrip.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiFile, Me.tsmiGrid})
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Size = New System.Drawing.Size(800, 27)
        Me.MenuStrip.TabIndex = 0
        Me.MenuStrip.Text = "MenuStrip"
        '
        'tsmiFile
        '
        Me.tsmiFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiLoad, Me.tsmiSave, Me.ToolStripSeparator, Me.tsmiExit, Me.tsmiRestart})
        Me.tsmiFile.Name = "tsmiFile"
        Me.tsmiFile.Size = New System.Drawing.Size(41, 23)
        Me.tsmiFile.Text = "&File"
        '
        'tsmiLoad
        '
        Me.tsmiLoad.Name = "tsmiLoad"
        Me.tsmiLoad.Size = New System.Drawing.Size(180, 24)
        Me.tsmiLoad.Text = "&Load"
        '
        'tsmiSave
        '
        Me.tsmiSave.Name = "tsmiSave"
        Me.tsmiSave.Size = New System.Drawing.Size(180, 24)
        Me.tsmiSave.Text = "&Save"
        '
        'ToolStripSeparator
        '
        Me.ToolStripSeparator.Name = "ToolStripSeparator"
        Me.ToolStripSeparator.Size = New System.Drawing.Size(177, 6)
        '
        'tsmiExit
        '
        Me.tsmiExit.Name = "tsmiExit"
        Me.tsmiExit.Size = New System.Drawing.Size(180, 24)
        Me.tsmiExit.Text = "OK and E&xit"
        '
        'tsmiRestart
        '
        Me.tsmiRestart.Name = "tsmiRestart"
        Me.tsmiRestart.Size = New System.Drawing.Size(180, 24)
        Me.tsmiRestart.Text = "Restart"
        Me.tsmiRestart.Visible = False
        '
        'tsmiGrid
        '
        Me.tsmiGrid.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiCollapseAll, Me.tsmiCollapseStd, Me.tsmiExpandAll, Me.tsmiRefresh, Me.tsmiEnlargeLeft})
        Me.tsmiGrid.Name = "tsmiGrid"
        Me.tsmiGrid.Size = New System.Drawing.Size(47, 23)
        Me.tsmiGrid.Text = "&Grid"
        '
        'tsmiCollapseAll
        '
        Me.tsmiCollapseAll.Name = "tsmiCollapseAll"
        Me.tsmiCollapseAll.Size = New System.Drawing.Size(146, 24)
        Me.tsmiCollapseAll.Text = "Collapse"
        '
        'tsmiCollapseStd
        '
        Me.tsmiCollapseStd.Name = "tsmiCollapseStd"
        Me.tsmiCollapseStd.Size = New System.Drawing.Size(146, 24)
        Me.tsmiCollapseStd.Text = "    std."
        '
        'tsmiExpandAll
        '
        Me.tsmiExpandAll.Name = "tsmiExpandAll"
        Me.tsmiExpandAll.Size = New System.Drawing.Size(146, 24)
        Me.tsmiExpandAll.Text = "Expand"
        '
        'tsmiRefresh
        '
        Me.tsmiRefresh.Name = "tsmiRefresh"
        Me.tsmiRefresh.Size = New System.Drawing.Size(146, 24)
        Me.tsmiRefresh.Text = "Refresh"
        '
        'tsmiEnlargeLeft
        '
        Me.tsmiEnlargeLeft.Name = "tsmiEnlargeLeft"
        Me.tsmiEnlargeLeft.Size = New System.Drawing.Size(146, 24)
        Me.tsmiEnlargeLeft.Text = "Enlarge left"
        '
        'StatusStrip
        '
        Me.StatusStrip.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblStatus, Me.progressBar})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 426)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(800, 24)
        Me.StatusStrip.TabIndex = 1
        Me.StatusStrip.Text = "StatusStrip1"
        '
        'lblStatus
        '
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(18, 19)
        Me.lblStatus.Text = "..."
        '
        'progressBar
        '
        Me.progressBar.Name = "progressBar"
        Me.progressBar.Size = New System.Drawing.Size(100, 18)
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 27)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.propGridMain)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.propGridDetails)
        Me.SplitContainer1.Size = New System.Drawing.Size(800, 399)
        Me.SplitContainer1.SplitterDistance = 414
        Me.SplitContainer1.TabIndex = 2
        '
        'propGridMain
        '
        Me.propGridMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.propGridMain.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.propGridMain.Location = New System.Drawing.Point(0, 0)
        Me.propGridMain.Name = "propGridMain"
        Me.propGridMain.PropertySort = System.Windows.Forms.PropertySort.Categorized
        Me.propGridMain.Size = New System.Drawing.Size(414, 399)
        Me.propGridMain.TabIndex = 0
        Me.propGridMain.ToolbarVisible = False
        '
        'propGridDetails
        '
        Me.propGridDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.propGridDetails.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.propGridDetails.Location = New System.Drawing.Point(0, 0)
        Me.propGridDetails.Name = "propGridDetails"
        Me.propGridDetails.PropertySort = System.Windows.Forms.PropertySort.Categorized
        Me.propGridDetails.Size = New System.Drawing.Size(382, 399)
        Me.propGridDetails.TabIndex = 0
        Me.propGridDetails.ToolbarVisible = False
        '
        'frmPropGrid
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.StatusStrip)
        Me.Controls.Add(Me.MenuStrip)
        Me.MainMenuStrip = Me.MenuStrip
        Me.Name = "frmPropGrid"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " ... "
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#Region "    controls"

    Friend WithEvents MenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents propGridMain As PropertyGrid
    Friend WithEvents propGridDetails As System.Windows.Forms.PropertyGrid
    Friend WithEvents lblStatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tsmiFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLoad As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiSave As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiGrid As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiCollapseAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiCollapseStd As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiExpandAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiRefresh As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiEnlargeLeft As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents progressBar As System.Windows.Forms.ToolStripProgressBar

#End Region

End Class


Public Class frmPropGrid

#Region "    Constructor and init"

    ''' <summary>
    ''' starts form
    ''' </summary>
    ''' <param name="class2Show">
    ''' class shown in main
    ''' </param>
    ''' <param name="classType">
    ''' class type for de serializing
    ''' </param>
    Public Sub New(
                ByRef class2Show As Object,
       Optional ByVal classType As Type = Nothing,
       Optional ByVal restart As Boolean = False)

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.

        If restart Then Me.tsmiRestart.Visible = True

        Me.class2Show = class2Show

        If IsNothing(classType) Then
            classType = class2Show.GetType
        End If

        Me.Text = class2Show.GetType.ToString.Split(".").Last.ToUpper

        init(classType:=classType)

    End Sub

    ''' <summary>
    ''' init form
    ''' </summary>
    ''' <param name="class2Show">
    ''' class shown in main
    ''' </param>
    ''' <param name="classType">
    ''' class type for de serializing
    ''' </param>
    Public Sub init(
            Optional ByRef class2Show As Object = Nothing,
            Optional ByVal classType As Type = Nothing)


        If Not IsNothing(class2Show) Then
            Me.class2Show = class2Show
        End If

        Me.propGridMain.SelectedObject = Me.class2Show

        If Not IsNothing(classType) Then
            Me.classType = classType
        Else
            Me.classType = class2Show.GetType
        End If

        collapseStd()
        setSplitterPosPercent(PositionPercent:=50)
        initPgrid()

    End Sub

    Public Sub initPgrid(Optional myFont As String = "Courier New", Optional mySize As Integer = 10)

        With Me.propGridDetails

            .Font = New Font(familyName:=myFont, emSize:=mySize)

        End With

        With Me.propGridMain

            .Font = New Font(familyName:=myFont, emSize:=mySize)

        End With

    End Sub

    Public Sub setProgressBar(actual As Integer, Optional max As Integer = Integer.MaxValue, Optional msg As String = "?")

        If actual = -1 Then
            progressBar.Visible = False
            Me.lblStatus.Text = ""
            Exit Sub
        Else
            progressBar.Visible = True
        End If

        If max = Integer.MaxValue AndAlso actual > Me.progressBar.Maximum Then
            max = actual + 1
        End If

        If Me.progressBar.Maximum <> max Then
            Me.progressBar.Maximum = max
        End If

        Me.progressBar.Value = actual

        If msg <> "?" Then
            Me.lblStatus.Text = msg
        End If


    End Sub

    Public Sub setStatus(msg As String)
        Me.lblStatus.Text = msg
    End Sub

    Public Sub setSplitterPosPercent(PositionPercent As Integer)

        With Me.SplitContainer1

            .SplitterDistance =
                .ClientSize.Width * (PositionPercent / 100)

        End With

    End Sub

    Public Enum eMainDetailPGrid
        main
        detail
    End Enum

    <TypeConverter(GetType(enumConverter(Of eViewEdit)))>
    Public Enum eViewEdit

        <Description("View only")>
        view
        <Description("Edit or load")>
        edit
        <Description(enumConverter(Of Type).not_defined)>
        not_def = -1

    End Enum

#End Region

#Region "    Splitter"

    Public Shared Sub moveSplitter(ByVal propertyGrid As PropertyGrid, ByVal x As Integer)
        Dim propertyGridView As Object = GetType(PropertyGrid).InvokeMember("gridView", BindingFlags.GetField Or BindingFlags.NonPublic Or BindingFlags.Instance, Nothing, propertyGrid, Nothing)
        propertyGridView.[GetType]().InvokeMember("MoveSplitterTo", BindingFlags.InvokeMethod Or BindingFlags.NonPublic Or BindingFlags.Instance, Nothing, propertyGridView, New Object() {x})
    End Sub

    Public Shared Sub moveVerticalSplitter(grid As PropertyGrid, LabelRatio As Double)
        Try
            Dim info = grid.[GetType]().GetProperty("Controls")
            Dim collection = DirectCast(info.GetValue(grid, Nothing), Control.ControlCollection)

            For Each control As Object In collection
                Dim type = control.[GetType]()

                If "PropertyGridView" = type.Name Then
                    control.LabelRatio = LabelRatio

                    grid.HelpVisible = True
                    Exit For
                End If
            Next

        Catch ex As Exception
            Trace.WriteLine(ex)
        End Try
    End Sub

    Public Sub setPGridSplitterLabelRatio(
                            LabelRatio As Double,
                            MainOrDetailPGrid As eMainDetailPGrid)

        If MainOrDetailPGrid = eMainDetailPGrid.main Then

            moveVerticalSplitter(
                            grid:=Me.propGridMain,
                            LabelRatio:=LabelRatio)
        Else

            moveVerticalSplitter(
                            grid:=Me.propGridDetails,
                            LabelRatio:=LabelRatio)
        End If


    End Sub

#End Region

    Public Property class2Show As Object

    Public Property classType As Type

#Region "Form-Stuff"

    Private restart As Boolean = False

    Private Sub gridItemChangedMain(
            sender As Object,
            e As SelectedGridItemChangedEventArgs) Handles propGridMain.SelectedGridItemChanged

        If Not IsNothing(e.OldSelection) AndAlso
           Not IsNothing(e.NewSelection.Value) AndAlso
           e.NewSelection.Value.GetType.ToString <> (New Date).GetType.ToString Then

            Try

                With propGridDetails

                    .Font = New Font(
                                    familyName:="Courier New",
                                    emSize:=10)

                    .PropertySort = PropertySort.Categorized
                    .SelectedObject = e.NewSelection.Value
                    '.Refresh()

                End With

            Catch ex As Exception
                Console.Write(ex.Message)
            End Try

        End If

    End Sub

    Private Sub propertyValue_Changed(s As Object, e As PropertyValueChangedEventArgs) Handles _
                                propGridMain.PropertyValueChanged,
                                propGridDetails.PropertyValueChanged

        Dim oldValue As String

        '1st change
        If IsNothing(e.OldValue) Then
            oldValue = ""
        Else
            oldValue = e.OldValue.ToString
        End If

        Dim NewValue As String
        Dim ChangedItemLabel As String = e.ChangedItem.Label.ToString
        Dim Status As String

        If Not IsNothing(e.ChangedItem.Value) Then

            NewValue = e.ChangedItem.Value.ToString

            oldValue = Replace(
                            Expression:=oldValue,
                                  Find:=Environment.NewLine,
                           Replacement:=" ")

            If oldValue <> NewValue Then

                Status =
                    "Value for '" &
                        ChangedItemLabel & "' changed: " &
                        IIf(
                            oldValue = "",
                            "",
                            " from " & oldValue).ToString &
                            " to " & NewValue

                Status = Replace(
                                Expression:=Status,
                                      Find:=vbCrLf,
                               Replacement:="")

                Me.lblStatus.Text = Status

            End If

        End If

        ' refresh()

    End Sub

    Private Sub tsmiExit_Click(sender As Object, e As EventArgs) Handles tsmiExit.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub


    Private Sub tsmiRestart_Click(sender As Object, e As EventArgs) Handles tsmiRestart.Click
        Me.DialogResult = DialogResult.Retry
        Me.Close()
    End Sub


    Private Sub frmPGrid_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        Try

            If Not IsNothing(class2Show) AndAlso
               Not class2Show.inputComplete.ToString.ToString.Contains("OK") Then

                If MsgBox(
                    Prompt:=class2Show.inputComplete & " is missing, exit anyway?",
                    Buttons:=MsgBoxStyle.YesNo,
                    Title:="Input incomplete") = MsgBoxResult.No Then

                    e.Cancel = True

                End If

            End If

        Catch ex As Exception
            Console.WriteLine("incomplete not defined for " & class2Show.ToString)
        End Try

    End Sub


#Region "load and save"

    Private Sub save_click(sender As Object, e As EventArgs) Handles tsmiSave.Click

        Try
            Me.lblStatus.Text = serialize(targetClass:=class2Show)
        Catch ex As Exception
            Me.lblStatus.Text = ex.Message
        End Try

    End Sub

    Public Function serialize(targetClass As Object) As String

        Dim mySaveFileDialog As New SaveFileDialog

        With mySaveFileDialog

#Region "            SaveFileDialog settings"

            .Filter =
                "XML files (*.xml)|*.xml|" &
                "Binary files (*.soap)|*.soap|" &
                "JSON files (*.json)|*.json|" &
                "All files (*.*)|*.*"

            .FilterIndex = 0

            .AddExtension = True
            .AutoUpgradeEnabled = True

            .CheckFileExists = False
            .CheckPathExists = True

            .CreatePrompt = False
            .OverwritePrompt = True

#End Region

            'get filename
            If .ShowDialog <> DialogResult.OK Then
                Return " Save canceled by user"
            End If

            Select Case Path.GetExtension(.FileName).ToLower

                Case ".xml"

                    Dim myWriter As IO.StreamWriter = Nothing

                    Try

                        Dim mySerializer As XmlSerializer =
                            New XmlSerializer(targetClass.GetType)

                        myWriter = New IO.StreamWriter(.FileName)

                        mySerializer.Serialize(
                            textWriter:=myWriter,
                            o:=targetClass)

                    Catch ex As Exception

                        Return "Error serializing (saving) class " &
                            ex.Message

                    Finally
                        myWriter.Close()
                        myWriter = Nothing
                    End Try

                Case ".soap"

                    Dim formatter As IFormatter = Nothing
                    Dim fileStream As FileStream = Nothing

                    Try

                        fileStream =
                            New FileStream(
                        path:= .FileName,
                        mode:=FileMode.Create,
                      access:=FileAccess.Write)

                        formatter =
                            New Formatters.Binary.BinaryFormatter

                        formatter.Serialize(
                                serializationStream:=fileStream,
                                              graph:=targetClass)

                    Catch ex As Exception
                        Return "Error serializing to SOAP " & ex.Message
                    Finally

                        If fileStream IsNot Nothing Then
                            fileStream.Close()
                        End If

                    End Try

                Case ".json"

                    Try

                        Dim serializer As New JavaScriptSerializer

                        Dim serializedResult =
                            serializer.Serialize(targetClass)

                        serializedResult =
                            New jsonFormatter(json:=serializedResult).Format

                        File.WriteAllText(
                            path:= .FileName,
                            contents:=serializedResult)

                        Return True

                    Catch ex As Exception
                        Return "Error serializing to JSON " & ex.Message
                    End Try

            End Select

            Return "OK : Serializing class to file " & .FileName

        End With

        Return False

    End Function


    Private Sub load_click(sender As Object, e As EventArgs) Handles tsmiLoad.Click

        Try
            Me.lblStatus.Text = deSerialize(targetClass:=class2Show)
        Catch ex As Exception
            Me.lblStatus.Text = ex.Message
        End Try

    End Sub

    Public Function deSerialize(ByRef targetClass As Object) As String

        Dim myOpenFileDialog As New OpenFileDialog
        Dim xmlFile As String() = {}
        Dim myFileStream As IO.FileStream = Nothing
        Dim jsonString As String = String.Empty
        Dim jss = New JavaScriptSerializer()

        With myOpenFileDialog

#Region "            OpenFileDialog settings"

            .Filter =
                "XML files (*.xml)|*.xml|" &
                "Binary files (*.soap)|*.soap|" &
                "JSON files (*.json)|*.json|" &
                "All files (*.*)|*.*"

            .FilterIndex = 0

            .AddExtension = True
            .AutoUpgradeEnabled = True

            .CheckFileExists = False
            .CheckPathExists = True

#End Region

            'get filename
            If .ShowDialog <> DialogResult.OK Then
                Return "Canceled by user"
            End If

            Select Case Path.GetExtension(.FileName).ToLower

                Case ".xml"

                    'read xml file, text and stream
                    Try

                        xmlFile = File.ReadAllLines(.FileName)

                        'check if type fits file
                        If Not xmlFile(1).StartsWith(
                            "<" & classType.ToString.Split(".").Last) Then

                            Return "File and class don't match : class= " &
                                classType.ToString.Split(".").Last & " ; file = " &
                            Replace(Expression:=xmlFile(1).Split.First, Find:="<", Replacement:="") & " ; " & Path.GetFileName(.FileName)

                        End If


                        ' To read the file, create a FileStream.
                        myFileStream =
                            New IO.FileStream(
                                    path:= .FileName,
                                    mode:=IO.FileMode.Open)

                    Catch ex As Exception

                        If Not IsNothing(myFileStream) Then
                            myFileStream.Close()
                        End If

                        Return "IO error : " & ex.Message

                    End Try

                    'de-serialize xml string to class2Show
                    Try

                        Dim mySerializer As XmlSerializer =
                            New XmlSerializer(targetClass.GetType)

                        targetClass = Nothing
                        targetClass =
                            mySerializer.Deserialize(myFileStream)

                        classType = targetClass.GetType

                    Catch ex As Exception
                        Return "Error de-serializing xml string to " &
                               classType.ToString
                    Finally
                        myFileStream.Close()
                    End Try

                    'check type and show in property grid
                    Try

                        With Me.propGridMain

                            .SelectedObject = Nothing
                            .SelectedObject = targetClass
                            .Refresh()

                        End With

                        Return "OK : Load file " & .FileName

                    Catch ex As Exception

                        Return ex.Message

                    End Try


                Case ".soap"

                    Dim formatter As IFormatter = Nothing
                    Dim fileStream As FileStream = Nothing

                    'getting file stream
                    Try
                        fileStream = New FileStream(
                                             path:= .FileName,
                                             mode:=FileMode.Open,
                                           access:=FileAccess.Read)
                    Catch ex As Exception

                        If fileStream IsNot Nothing Then
                            fileStream.Close()
                        End If

                        Return "IO Error De-Serializing : " &
                            ex.Message

                    End Try

                    'de-serialization SOAP binary and 
                    'show in property grid
                    Try

                        formatter =
                            New Formatters.Binary.BinaryFormatter

                        targetClass =
                            formatter.Deserialize(
                            serializationStream:=fileStream)


                        With Me.propGridMain

                            .SelectedObject = Nothing
                            .SelectedObject = targetClass
                            .Refresh()

                        End With

                        Return "OK : SOAP De-Serializing from " &
                            .FileName

                    Catch ex As Exception
                        Return "ERROR SOAP De-Serializing : " &
                            ex.Message

                    Finally

                        If fileStream IsNot Nothing Then
                            fileStream.Close()
                        End If
                    End Try

                Case ".json"

                    'get json string
                    Try
                        jsonString =
                            File.ReadAllText(
                                path:= .FileName)
                    Catch ex As Exception
                        Return "IO ERROR reading JSON file : " & ex.Message
                    End Try

                    'de-serialization JSON string and 
                    'show in property grid
                    Try

                        targetClass = jss.Deserialize(
                                       input:=jsonString,
                                       targetType:=targetClass.GetType)


                        With Me.propGridMain

                            .SelectedObject = Nothing
                            .SelectedObject = targetClass
                            .Refresh()

                        End With


                    Catch ex As Exception
                        Return "ERROR de-serialization JSON " & ex.Message
                    End Try

                Case Else

                    Return "Unknown file type "

            End Select

        End With

        Return "?"

    End Function

#Region "    json formating"

    Public Class jsonFormatter

        Private ReadOnly _walker As StringWalker
        Private ReadOnly _writer As IndentWriter = New IndentWriter()
        Private ReadOnly _currentLine As StringBuilder = New StringBuilder()
        Private _quoted As Boolean

        Public Sub New(ByVal json As String)
            _walker = New StringWalker(json)
            ResetLine()
        End Sub

        Public Sub ResetLine()
            _currentLine.Length = 0
        End Sub

        Public Function Format() As String
            While MoveNextChar()

                If Not Me._quoted AndAlso Me.isOpenBracket() Then
                    Me.writeCurrentLine()
                    Me.addCharToLine()
                    Me.writeCurrentLine()
                    _writer.Indent()
                ElseIf Not Me._quoted AndAlso Me.isCloseBracket() Then
                    Me.writeCurrentLine()
                    _writer.UnIndent()
                    Me.addCharToLine()
                ElseIf Not Me._quoted AndAlso Me.isColon() Then
                    Me.addCharToLine()
                    Me.writeCurrentLine()
                Else
                    addCharToLine()
                End If
            End While

            Me.writeCurrentLine()
            Return _writer.ToString()
        End Function

        Private Function MoveNextChar() As Boolean
            Dim success As Boolean = _walker.MoveNext()

            If Me.isApostrophe() Then
                Me._quoted = Not _quoted
            End If

            Return success
        End Function

        Public Function isApostrophe() As Boolean
            Return Me._walker.currentChar = """"c AndAlso Me._walker.isEscaped = False
        End Function

        Public Function isOpenBracket() As Boolean
            Return Me._walker.currentChar = "{"c OrElse Me._walker.currentChar = "["c
        End Function

        Public Function isCloseBracket() As Boolean
            Return Me._walker.currentChar = "}"c OrElse Me._walker.currentChar = "]"c
        End Function

        Public Function isColon() As Boolean
            Return Me._walker.currentChar = ","c
        End Function

        Private Sub addCharToLine()
            Me._currentLine.Append(_walker.currentChar)
        End Sub

        Private Sub writeCurrentLine()
            Dim line As String = Me._currentLine.ToString().Trim()

            If line.Length > 0 Then
                _writer.WriteLine(line)
            End If

            Me.ResetLine()
        End Sub
    End Class

    Public Class StringWalker

        Private ReadOnly _s As String
        Public Property index As Integer
        Public Property isEscaped As Boolean
        Public Property currentChar As Char

        Public Sub New(ByVal s As String)
            _s = s
            Me.index = -1
        End Sub

        Public Function MoveNext() As Boolean

            If Me.index = _s.Length - 1 Then Return False

            If Not isEscaped Then
                isEscaped = currentChar = "\"c
            Else
                isEscaped = False
            End If

            Me.index += 1
            currentChar = _s(index)
            Return True

        End Function
    End Class

    Public Class IndentWriter

        Private ReadOnly _result As StringBuilder = New StringBuilder()
        Private _indentLevel As Integer

        Public Sub Indent()
            _indentLevel += 1
        End Sub

        Public Sub UnIndent()
            If _indentLevel > 0 Then _indentLevel -= 1
        End Sub

        Public Sub WriteLine(ByVal line As String)
            _result.AppendLine(CreateIndent() & line)
        End Sub

        Private Function CreateIndent() As String
            Dim indent As StringBuilder = New StringBuilder()

            For i As Integer = 0 To _indentLevel - 1
                indent.Append("    ")
            Next

            Return indent.ToString()
        End Function

        Public Overrides Function ToString() As String
            Return _result.ToString()
        End Function
    End Class


#End Region

#End Region

#Region "collapse and expand"

    Public Shared refreshOngoing As Boolean = False

    Private Sub tsmiCollapseStd_Click(sender As Object, e As EventArgs) Handles tsmiCollapseStd.Click
        collapseStd()
    End Sub

    Public Sub collapseStd()

        Dim root As GridItem

        If Not IsNothing(propGridMain.SelectedGridItem) AndAlso
           Not IsNothing(propGridMain.SelectedGridItem.Parent) AndAlso
           Not IsNothing(propGridMain.SelectedGridItem.Parent.Parent) Then

            root = propGridMain.SelectedGridItem.Parent.Parent

            Try

                For Each CollapesCategorie As String In class2Show.collapseStd

                    For counter As Integer = 0 To root.GridItems.Count - 1

                        If root.GridItems(counter).Label =
                                   CollapesCategorie Then

                            If Not refreshOngoing Then root.GridItems(counter).Expanded = False

                        End If

                    Next

                Next
            Catch ex As Exception
                lblStatus.Text = "No Public Property 'collapseStd' defined"
            End Try

        End If

        If Not IsNothing(propGridDetails.SelectedGridItem) AndAlso
           Not IsNothing(propGridDetails.SelectedGridItem.Parent) AndAlso
           Not IsNothing(propGridDetails.SelectedGridItem.Parent.Parent) Then

            root = propGridDetails.SelectedGridItem.Parent.Parent

            Try

                For Each CollapesCategorie As String In propGridDetails.SelectedObject.collapseStd

                    For counter As Integer = 0 To root.GridItems.Count - 1

                        If root.GridItems(counter).Label =
                                CollapesCategorie Then

                            If Not refreshOngoing Then root.GridItems(counter).Expanded = False

                        End If

                    Next

                Next
            Catch ex As Exception
                lblStatus.Text = "No Public Property 'collapseStd' defined"
            End Try

        End If

    End Sub

    Private Sub tsmiCollapseAll_Click(sender As Object, e As EventArgs) Handles tsmiCollapseAll.Click
        Me.propGridMain.CollapseAllGridItems()
        Me.propGridDetails.CollapseAllGridItems()
    End Sub

    Private Sub tsmiExpandAll_Click(sender As Object, e As EventArgs) Handles tsmiExpandAll.Click
        Me.propGridMain.ExpandAllGridItems()
        Me.propGridDetails.ExpandAllGridItems()
    End Sub

    Private Sub tsmiRefresh_Click(sender As Object, e As EventArgs) Handles tsmiRefresh.Click
        refreshPGrid()
    End Sub

    Public Sub refreshPGrid()

        With Me.propGridMain

            refreshOngoing = True

            .Refresh()
            .RefreshTabs(tabScope:=PropertyTabScope.Component)

            refreshOngoing = False

        End With

    End Sub


    Private Sub frmPGrid_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        lblStatus.Text = "Width : " & Me.Width & "  Height: " & Me.Height
    End Sub

    Private Sub tsmiEnlargeLeft_Click(sender As Object, e As EventArgs) Handles tsmiEnlargeLeft.Click

        Me.Size =
            New Size(
            width:=1300,
            height:=1300)

        Me.propGridMain.CollapseAllGridItems()
        Me.propGridDetails.CollapseAllGridItems()

        setSplitterPosPercent(20)

        setPGridSplitterLabelRatio(
            LabelRatio:=1,
            MainOrDetailPGrid:=eMainDetailPGrid.detail)

        setPGridSplitterLabelRatio(
            LabelRatio:=15,
            MainOrDetailPGrid:=eMainDetailPGrid.detail)

    End Sub

#End Region

#End Region

End Class


