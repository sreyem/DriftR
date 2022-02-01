
Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.IO
Imports System.Web.Script.Serialization
Imports System.Xml.Serialization


Module Main

    Sub main()


restart:


        Console.Clear()

        'mylog(LogTxtArray:=getStartInfo)

        'mylog(
        '    LogTxt:="ztsParser v " &
        '    getExecVersionInfo(
        '    filename:=Path.Combine(Environment.CurrentDirectory, "ztsParser.exe")).ToString)

        'DriftR = New DriftR
        'DriftR.DrainHyd553 = File.ReadAllLines(path:="DrainHyd5553.csv")
        'DriftR.SWASHBBCHout = File.ReadAllLines(path:="SWASH_BBCH.out")

        BBCH = New BBCH
        BBCH.SWASH_BBCHout = File.ReadAllLines(path:="SWASH_BBCH.out")



        With DriftR

            '.FOCUSswDriftCrop = eFOCUSswDriftCrop.CW
            '.FOCUSswScenario = eFOCUSswScenario.D5
            '.FOCUSswWaterBody = eFOCUSswWaterBody.ST
            '.ApplnMethodStep03 = eApplnMethodStep03.GS
            '.NoOfApplns = eNoOfApplns._02
            '.Rate = 0.25
            '.ApplnDate = New Date(year:=1901, month:=4, day:=22)

            '.RAC = 0.3

        End With


        'showform = New frmPropGrid(
        '    class2Show:=DriftR,
        '     classType:=DriftR.GetType,
        '       restart:=True)

        showform = New frmPropGrid(
            class2Show:=BBCH,
             classType:=BBCH.GetType,
               restart:=True)

        Try

            If Not IsNothing(showform) Then

                With showform

                    .Width = 700
                    .Height = 800

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

    Public Property DriftR As New DriftR

    Public Property BBCH As New BBCH

End Module

<Serializable>
<DefaultProperty("FOCUSswDriftCrop")>
<TypeConverter(GetType(PropGridConverter))>
Public Class DriftR

    Public Sub New()

    End Sub

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const catInput As String = " 01 Main"

#Region "    Main"

#Region "    FOCUS Crop & Scenario"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _FOCUSswDriftCrop As eFOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined

    ''' <summary>
    ''' FOCUSswDriftCrop
    ''' Target crop out of the
    ''' available FOCUSsw crops
    ''' </summary>
    <Category(catInput)>
    <DisplayName(
    "SWASH Crop")>
    <Description(
    "FOCUS SWASH crop for drift" & vbCrLf &
    "triggers Ganzelmeier crop group")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(CInt(eFOCUSswDriftCrop.not_defined))>
    Public Property FOCUSswDriftCrop As eFOCUSswDriftCrop
        Get
            Return _FOCUSswDriftCrop
        End Get
        Set

            'set std appln. method for air blast or ground spray
            Select Case Value

                Case _
                    eFOCUSswDriftCrop.HP,
                    eFOCUSswDriftCrop.CI,
                    eFOCUSswDriftCrop.OL,
                    eFOCUSswDriftCrop.PFE,
                    eFOCUSswDriftCrop.PFL,
                    eFOCUSswDriftCrop.VI,
                    eFOCUSswDriftCrop.VIL

                    _applnMethodStep03 = eApplnMethodStep03.AB

                    enumConverter(Of eApplnMethodStep03).dontShow =
                    {enumConverter(Of eApplnMethodStep03).getEnumDescription(eApplnMethodStep03.SI)}

                Case Else

                    _applnMethodStep03 = eApplnMethodStep03.GS

                    enumConverter(Of eApplnMethodStep03).dontShow =
                    {enumConverter(Of eApplnMethodStep03).getEnumDescription(eApplnMethodStep03.AB)}

            End Select

            _FOCUSswDriftCrop = Value

            UpdateGanzelmeier(
                Ganzelmeier:=Me.Ganzelmeier,
                FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                ApplnMethodStep03:=_applnMethodStep03)

            'only scenarios defined for this crop
            enumConverter(Of eFOCUSswScenario).onlyShow =
                GetScenariosFromCrop(Crop:=_FOCUSswDriftCrop)

            'reset if crop changed
            _FOCUSswScenario = eFOCUSswScenario.not_defined
            _FOCUSswWaterBody = eFOCUSswWaterBody.not_defined
            _waterDepth = 0

            With Me.Distances
                .Crop2Bank = 0
                .Bank2Water = 0
                .Closest2EdgeOfField = 0
                .Farthest2EdgeOfWB = 0
            End With

        End Set
    End Property

    <Category(catInput)>
    <RefreshProperties(RefreshProperties.All)>
    Public ReadOnly Property BBCHCrop As String
        Get
            Return GetBBCHSearchStringsFromCrop(
                FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                Season:=eSeason.first)

        End Get
    End Property


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _FOCUSswScenario As eFOCUSswScenario = eFOCUSswScenario.not_defined

    ''' <summary>
    ''' FOCUSswScenario
    ''' FOCUSsw Scenario
    ''' D1 - D6 And R1 - R4
    ''' </summary>
    <Category(catInput)>
    <DisplayName(
    "Scenario")>
    <Description(
    "All FOCUSsw scenarios " & vbCrLf &
    "D1 - D6 And R1 - R4")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(CInt(eFOCUSswScenario.not_defined))>
    Public Property FOCUSswScenario As eFOCUSswScenario
        Get
            Return _FOCUSswScenario
        End Get
        Set

            If _FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined Then
                _FOCUSswScenario = eFOCUSswScenario.not_defined
            Else

                If enumConverter(Of eFOCUSswScenario).onlyShow.Contains(
                   enumConverter(Of eFOCUSswScenario).getEnumDescription(EnumConstant:=Value)) Then

                    _FOCUSswScenario = Value

                    enumConverter(Of eFOCUSswWaterBody).onlyShow =
                            GetWaterBodysFromScenario(_FOCUSswScenario)

                    If _FOCUSswScenario.ToString.ToUpper.StartsWith("R") Then
                        Me.FOCUSswWaterBody = eFOCUSswWaterBody.ST
                    Else

                        _FOCUSswWaterBody = eFOCUSswScenario.not_defined

                        With Me.Distances

                            .Crop2Bank = 0
                            .Bank2Water = 0
                            .Closest2EdgeOfField = 0
                            .Farthest2EdgeOfWB = 0

                        End With

                    End If

                End If

            End If

        End Set
    End Property

#End Region

#Region "    Water body and depth in m"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _FOCUSswWaterBody As eFOCUSswWaterBody = eFOCUSswWaterBody.not_defined

    ''' <summary>
    ''' FOCUSswWaterBody
    ''' FOCUS water body
    ''' Ditch, pond or stream
    ''' </summary>
    <Category(catInput)>
    <DisplayName(
    "Water  Body")>
    <Description(
    "Ditch, stream : 100m x 1m x 0.3m ;  30,000L" & vbCrLf &
    "Pond          :  30m rad. x 1.0m ; ~94,000L")>
    <RefreshProperties(RefreshProperties.All)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(CInt(eFOCUSswWaterBody.not_defined))>
    Public Property FOCUSswWaterBody As eFOCUSswWaterBody
        Get
            Return _FOCUSswWaterBody
        End Get
        Set

            If _FOCUSswScenario = eFOCUSswScenario.not_defined Then

                _FOCUSswWaterBody = eFOCUSswScenario.not_defined

            Else

                If enumConverter(Of eFOCUSswWaterBody).onlyShow.Contains(
                   enumConverter(Of eFOCUSswWaterBody).getEnumDescription(
                   EnumConstant:=Value)) OrElse
                   Value = eFOCUSswWaterBody.not_defined Then

                    _FOCUSswWaterBody = Value

                    Me.Distances.GetDistance(
                        Distances:=Me.Distances,
                        FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                        FOCUSswWaterBody:=_FOCUSswWaterBody,
                        ApplnMethodStep03:=_applnMethodStep03,
                        Buffer:=_buffer)

                    If _FOCUSswWaterBody = eFOCUSswWaterBody.PO Then
                        _waterDepth = 1
                    ElseIf _FOCUSswScenario.ToString.ToUpper.StartsWith("R") Then
                        _waterDepth = CDbl(
                           RunOffWaterDepthsEstimates(_FOCUSswScenario.ToString.Substring(1, 1) - 1))
                    Else
                        _waterDepth = 0.3
                    End If

                    CalcDriftPercentDistances(
                        Distances:=Me.Distances,
                        Nozzle:=Me.Nozzle)

                End If

            End If

        End Set
    End Property


    Public RunOffWaterDepthsEstimates As Double() =
        {
        0.41, 'R1
        0.305,'R2
        0.29, 'R3
        0.41  'R4 
        }

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _waterDepth As Double = 0

    ''' <summary>
    ''' WaterDepth
    ''' Depth of water body in m
    ''' </summary>
    ''' <returns></returns>
    <Category(catInput)>
    <DisplayName(
    "       Depth")>
    <Description(
    "in m" & vbCrLf &
    "")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '0.000'|" &
    "unit='m'")>
    Public Property WaterDepth As Double
        Get
            Return _waterDepth
        End Get
        Set
            _waterDepth = Value
        End Set
    End Property

#End Region

#Region "    Application number, method and rate"


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _noOfApplns As eNoOfApplns = eNoOfApplns.not_defined

    ''' <summary>
    ''' NoOfApplns
    ''' Number of applications
    ''' 1 - 8+
    ''' </summary>
    <Category(catInput)>
    <DisplayName(
    "Appln. Number")>
    <Description(
    "Max. number Of applications" & vbCrLf &
    "1 - 8 (Or more)")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(CInt(eNoOfApplns.not_defined))>
    Public Property NoOfApplns As eNoOfApplns
        Get
            Return _noOfApplns
        End Get
        Set

            _noOfApplns = Value

            If _noOfApplns = eNoOfApplns._01 Then
                Interval = Integer.MaxValue
            End If

            Me.Regression.Update(
                    Regression:=Me.Regression,
                    Ganzelmeier:=Ganzelmeier,
                    NoOfApplns:=_noOfApplns)

            CalcDriftPercentDistances(
                        Distances:=Me.Distances,
                        Nozzle:=Me.Nozzle)

        End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _applnMethodStep03 As eApplnMethodStep03 = eApplnMethodStep03.not_defined

    ''' <summary>
    ''' ApplnMethodStep03
    ''' Application Methods
    ''' </summary>
    ''' <returns></returns>
    <Category(catInput)>
    <DisplayName(
    "       Method")>
    <Description(
    "Appln. method, basically appln. to :" & vbCrLf &
    "Crop canopy, soil or incorporation")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(CInt(eApplnMethodStep03.not_defined))>
    Public Property ApplnMethodStep03 As eApplnMethodStep03
        Get
            Return _applnMethodStep03
        End Get
        Set

            If _FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined Then
                Value = eApplnMethodStep03.not_defined
            Else

                Select Case _FOCUSswDriftCrop

                    Case _
                        eFOCUSswDriftCrop.HP,
                        eFOCUSswDriftCrop.CI,
                        eFOCUSswDriftCrop.OL,
                        eFOCUSswDriftCrop.PFE,
                        eFOCUSswDriftCrop.PFL,
                        eFOCUSswDriftCrop.VI,
                        eFOCUSswDriftCrop.VIL

                        If Value = eApplnMethodStep03.SI Then
                            Value = eApplnMethodStep03.not_defined
                        End If

                    Case Else

                        If Value = eApplnMethodStep03.AB Then
                            Value = eApplnMethodStep03.not_defined
                        End If

                End Select

            End If

            _applnMethodStep03 = Value

            UpdateGanzelmeier(
                Ganzelmeier:=Me.Ganzelmeier,
                FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                ApplnMethodStep03:=_applnMethodStep03)

            Me.Distances.GetDistance(
                        Distances:=Me.Distances,
                FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                FOCUSswWaterBody:=_FOCUSswWaterBody,
                ApplnMethodStep03:=_applnMethodStep03,
                Buffer:=_buffer)

        End Set
    End Property

#Region "    Rate"

    ''' <summary>
    ''' GUI : Time between applns. in days
    ''' </summary>
    ''' <returns></returns>
    <Category(catInput)>
    <DisplayName(
    "       Rate")>
    <Description(
    "Application rate" & vbCrLf &
    "in kg as/ha")>
    <TypeConverter(GetType(dropDownList))>
    <RefreshProperties(RefreshProperties.All)>
    <DefaultValue(" - ")>
    <Browsable(True)>
    <XmlIgnore> <ScriptIgnore>
    Public Property RateGUI As String
        Get

            dropDownList.dropDownEntries =
                    {
                    " - ",
                    "0.005",
                    "0.010",
                    "0.050",
                    "0.075",
                    "0.100",
                    "0.150",
                    "0.200",
                    "0.250",
                    "0.300",
                    "0.400",
                    "0.500",
                    "0.750",
                    "1.000"
                    }

            If Rate > 0 Then
                Return Rate.ToString() & " kg as/ha"
            Else
                Return " - "
            End If

        End Get
        Set

            If IsNothing(Value) Then Value = " - "

            If Value <> " - " Then

                Value =
                    Replace(
                    Expression:=Value.ToUpper,
                    Find:="KG",
                    Replacement:="",
                    Compare:=CompareMethod.Text)

                Value =
                    Replace(
                    Expression:=Value.ToUpper,
                    Find:="AS",
                    Replacement:="",
                    Compare:=CompareMethod.Text)

                Value =
                    Replace(
                    Expression:=Value.ToUpper,
                    Find:="AI",
                    Replacement:="",
                    Compare:=CompareMethod.Text)

                Value =
                    Replace(
                    Expression:=Value.ToUpper,
                    Find:="/",
                    Replacement:="",
                    Compare:=CompareMethod.Text)

                Value =
                    Replace(
                    Expression:=Value.ToUpper,
                    Find:="HA",
                    Replacement:="",
                    Compare:=CompareMethod.Text)

                Try
                    _rate = Double.Parse(Trim(Value))
                Catch ex As Exception
                    _rate = 0
                End Try

            Else

                _rate = 0

            End If

        End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _rate As Double = 0

    ''' <summary>
    ''' Rate
    ''' Appln Rate
    ''' </summary>
    ''' <returns></returns>
    <Category(catInput)>
    <DisplayName(
    "       Rate")>
    <Description(
    "Application rate" & vbCrLf &
    "in kg as/ha")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(False)>
    <[ReadOnly](False)>
    <DefaultValue(0)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= 'G4'|" &
    "unit=' kg as/ha'")>
    Public Property Rate As Double
        Get
            Return _rate
        End Get
        Set
            _rate = Value
        End Set
    End Property


#End Region


    Public DrainHyd553 As String() = {}

    Private _ApplnDate As New Date

    <Category(catInput)>
    <DisplayName(
    "       Date")>
    <Description(
    "Appln. date" & vbCrLf &
    "")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <TypeConverter(GetType(DateConv))>
    <AttributeProvider("format= 'dd-MMM'")>
    Public Property ApplnDate As Date
        Get
            Return _ApplnDate
        End Get
        Set
            _ApplnDate = Value

            If _FOCUSswScenario.ToString.ToUpper.StartsWith("D") Then

                Dim Header As String() = DrainHyd553.First.Split(",")
                Dim year As Integer = 1901
                Dim SearchHeader As String
                Dim TargetColumn As Integer = -1
                Dim TargetDate As Date
                Dim depth As Double = -1
                Dim temp As String


                SearchHeader = _FOCUSswScenario.ToString & "_" & _FOCUSswWaterBody.ToString & "_" & _FOCUSswDriftCrop.ToString.Substring(0, 2)
                For ColumnCounter As Integer = 0 To Header.Count - 1
                    If Trim(Header(ColumnCounter)) = SearchHeader Then
                        TargetColumn = ColumnCounter
                        Exit For
                    End If
                Next

                year = DrainHyd553(1).Split(",")(TargetColumn)
                TargetDate =
                    New Date(
                        year:=year,
                        month:=Value.Month,
                        day:=Value.Day)
                temp = DrainHyd553(1 + TargetDate.DayOfYear)
                Header = DrainHyd553(1 + TargetDate.DayOfYear).Split(",")

                WaterDepth = DrainHyd553(1 + TargetDate.DayOfYear).Split(",")(TargetColumn)
                _ApplnDate = TargetDate

            End If

        End Set
    End Property


#End Region

#End Region

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const catResults As String = " 02 Results"

#Region "    Results"

#Region "    Step 02"

    Public _Step02 As Double = 0

    <Category(catResults)>
    Public ReadOnly Property Step02 As String
        Get

            Dim Distances As New Distances

            Select Case _FOCUSswDriftCrop

                Case eFOCUSswDriftCrop.HP

                    If _noOfApplns <> eNoOfApplns.not_defined Then
                        _Step02 =
                            {
                            19.326,
                            17.723,
                            15.928,
                            15.378,
                            15.114,
                            14.902,
                            14.628,
                            13.52,
                            13.52
                            }(_noOfApplns)
                    End If

                Case eFOCUSswDriftCrop.CI,
                     eFOCUSswDriftCrop.OL,
                     eFOCUSswDriftCrop.VI,
                     eFOCUSswDriftCrop.VIL,
                     eFOCUSswDriftCrop.PF,
                     eFOCUSswDriftCrop.PFE,
                     eFOCUSswDriftCrop.PFL

                    Distances.GetDistance(
                        Distances:=Distances,
                        FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                        FOCUSswWaterBody:=eFOCUSswWaterBody.DI,
                        ApplnMethodStep03:=_applnMethodStep03,
                        Buffer:=3)

                Case Else

                    Distances.GetDistance(
                        Distances:=Distances,
                        FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                        FOCUSswWaterBody:=eFOCUSswWaterBody.DI,
                        ApplnMethodStep03:=_applnMethodStep03,
                        Buffer:=1)

            End Select


            If _FOCUSswDriftCrop <> eFOCUSswDriftCrop.HP Then

                CalcDriftPercentDistances(Distances:=Distances, Nozzle:=eNozzles._0)

                If _noOfApplns = eNoOfApplns._01 Then
                    _Step02 = Distances.NearestDriftPercentSingle
                Else
                    _Step02 = Distances.NearestDriftPercentMulti
                End If

            End If


            If _Step02 = 0 Then Return " - "

            _Step02 = Math.Round(_Step02, digits:=3)

            If _rate <> 0 Then

                _Step02 = _Step02 * _rate / 0.3

                If _noOfApplns > eNoOfApplns._01 AndAlso MAF > 1 Then
                    _Step02 *= MAF
                End If

                Return "PECsw : " & _Step02.ToString("0.0000") & " μg/L"

            Else
                Return "Drift : " & _Step02.ToString("0.0000") & " %"
            End If

        End Get
    End Property


#End Region

#Region "    Step 03"

    Public _Step03 As Double

    <Category(catResults)>
    <Editor(GetType(buttonEmulator), GetType(UITypeEditor))>
    Public Property Step03 As String
        Get

            _Step03 =
                CalcDriftPercent(
                noOfApplns:=_noOfApplns,
                FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                FOCUSswWaterBody:=_FOCUSswWaterBody,
                ApplnMethodStep03:=_applnMethodStep03)

            If _rate = 0 Then

                If _Step03 <> 0 AndAlso Not Double.IsNaN(_Step03) Then
                    Return "Drift : " & _Step03.ToString("0.0000") & " %"
                Else
                    _Step03 = 0
                    Return " - "
                End If


            Else

                If _waterDepth > 0 AndAlso Not Double.IsNaN(_Step03) Then
                    _Step03 = _rate * _Step03 / _waterDepth

                    If _noOfApplns > eNoOfApplns._01 AndAlso
                       _FOCUSswWaterBody = eFOCUSswWaterBody.PO Then
                        _Step03 *= MAF
                    End If

                    Return "PECsw : " & _Step03.ToString("0.0000") & " μg/L"
                Else
                    _Step03 = 0
                    Return "depth?"
                End If

            End If

        End Get
        Set(value As String)
            If value = buttonEmulator.Clicked Then

                Dim filepath As String = String.Empty
                Dim CSVfiles As String() = {}
                Dim CSVfile As String() = {}
                Dim rate As Double = -1
                Dim appln As Double = -1
                Dim temp() As String = {}
                Dim maxRow() As String = {}
                Dim max As Double = -1
                Dim maxdate As New Date
                Dim dummy As String()


                filepath = InputBox(Prompt:="CSV filepath")
                filepath = Replace(Expression:=filepath, Find:=Chr(34), Replacement:="", Compare:=CompareMethod.Text)

                Try
                    CSVfiles =
                    Directory.GetFiles(
                    path:=Path.GetDirectoryName(filepath),
                    searchPattern:="*" & Path.GetFileName(filepath).Split("-")(1) & "*.csv")

                    If CSVfiles.Count > 0 Then
                        Console.WriteLine(CSVfiles.Count & " files found")
                    Else
                        Console.WriteLine("No files found : *" & Path.GetFileName(filepath).Split("-")(1) & "*.csv")
                        Exit Property
                    End If
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                    Exit Property
                End Try

                For Each member As String In CSVfiles

                    Me.FOCUSswScenario =
                    [Enum].Parse(
                    enumType:=GetType(eFOCUSswScenario),
                    value:=member.Split("-").Last.Substring(0, 2))

                    Select Case member.Split("-").Last.Substring(2, 1)
                        Case "D"
                            Me.FOCUSswWaterBody = eFOCUSswWaterBody.DI
                        Case "S"
                            Me.FOCUSswWaterBody = eFOCUSswWaterBody.ST
                        Case Else
                            Me.FOCUSswWaterBody = eFOCUSswWaterBody.PO
                    End Select

                    Try
                        CSVfile = File.ReadAllLines(member)
                        'Console.WriteLine(Path.GetFileName(member))
                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                        Exit Property
                    End Try


                    Me.Rate =
                    Double.Parse(Filter(
                    Source:=CSVfile,
                    Match:="* Rate(s)",
                    Include:=True,
                    Compare:=CompareMethod.Text).First.Split(",").Last) / 1000

                    temp = Filter(
                    Source:=CSVfile,
                    Match:="* Appln(s)",
                    Include:=True,
                    Compare:=CompareMethod.Text).First.Split(",")

                    Me.NoOfApplns = temp.Count - 2

                    max = -1
                    maxdate = New Date

                    For counter As Integer = 1 To temp.Count - 1

                        maxRow = Filter(
                    Source:=CSVfile,
                    Match:=temp(counter) & " 09:00",
                    Include:=True,
                    Compare:=CompareMethod.Text).First.Split(",")

                        If Double.Parse(maxRow(1)) > max Then
                            max = Double.Parse(maxRow(1))
                            dummy = maxRow.First.Split
                            dummy = dummy.First.Split("-")

                            maxdate =
                                New Date(
                                year:=CInt(dummy(0)),
                                month:=CInt(dummy(1)),
                                day:=CInt(dummy(2)))
                        End If

                    Next

                    dummy = Path.GetFileNameWithoutExtension(member).Split("-")

                    Me.ApplnDate = maxdate
                    Console.WriteLine(
                        maxdate.ToShortDateString & " " & dummy(0).PadLeft(2) & "-" & dummy(2) & " : " &
                            max.ToString("0.0000") & "μg/L: " & Me.Step03.PadLeft(15) & " : " &
                           (100 - max / _Step03 * 100).ToString("0.00") & "%")

                Next

            End If

        End Set
    End Property

#End Region

#Region "    Step 04"

    Public _Step04 As Double

    <Category(catResults)>
    Public ReadOnly Property Step04 As String
        Get

            If _nozzle = 0 AndAlso _buffer = 0 Then
                Return " - "
            End If

            Dim DriftPercent As Double

            DriftPercent =
                CalcDriftPercent(
                noOfApplns:=_noOfApplns,
                FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                FOCUSswWaterBody:=_FOCUSswWaterBody,
                ApplnMethodStep03:=_applnMethodStep03,
                Nozzle:=_nozzle,
                Buffer:=_buffer)

            If _rate = 0 Then

                If DriftPercent <> 0 AndAlso Not Double.IsNaN(DriftPercent) Then

                    _Step04 = DriftPercent

                    Return "Drift : " & _Step04.ToString("0.0000") & " %"
                Else
                    _Step04 = 0
                    Return " - "
                End If


            Else

                If _waterDepth > 0 AndAlso Not Double.IsNaN(DriftPercent) Then

                    _Step04 = _rate * DriftPercent / _waterDepth

                    If _noOfApplns > eNoOfApplns._01 AndAlso
                       _FOCUSswWaterBody = eFOCUSswWaterBody.PO Then
                        _Step04 *= MAF
                    End If

                    Return "PECsw : " & _Step04.ToString("0.0000") & " μg/L"
                Else
                    _Step04 = 0
                    Return "depth?"
                End If

            End If


        End Get
    End Property


#Region "    Buffer"

    ''' <summary>
    ''' GUI : Time between applns. in days
    ''' </summary>
    ''' <returns></returns>
    <Category(catResults)>
    <DisplayName(
    "Buffer")>
    <Description(
    " in m")>
    <TypeConverter(GetType(dropDownList))>
    <RefreshProperties(RefreshProperties.All)>
    <DefaultValue("Step03, no add. buffer")>
    <Browsable(True)>
    <XmlIgnore> <ScriptIgnore>
    Public Property BufferGUI As String
        Get

            dropDownList.dropDownEntries =
                    {
                    "Step03" &
                    IIf(Me.Distances.Farthest2EdgeOfWB <> 0 AndAlso _FOCUSswWaterBody <> eFOCUSswWaterBody.not_defined,
                        " (" & Me.Distances.Farthest2EdgeOfWB & "m)", ""),
                    " 1m ( < FOCUS std. buffer? )",
                    " 2m ( < FOCUS std. buffer? )",
                    " 3m ( < FOCUS std. buffer? )",
                    " 4m ( < FOCUS std. buffer? )",
                    " 5m",
                    "10m",
                    "15m",
                    "20m",
                    "30m",
                    "40m",
                    "50m"
                    }

            If _buffer > 0 AndAlso _buffer <= 100 Then
                Return _buffer.ToString() & "m"
            Else
                Return dropDownList.dropDownEntries.First
            End If

        End Get
        Set

            Try

                Value =
                    Value.Split(
                    separator:={" "c},
                    options:=StringSplitOptions.RemoveEmptyEntries).First

                Buffer =
                    Double.Parse(
                    Trim(
                        Replace(
                        Expression:=Value,
                        Find:="m",
                        Replacement:="",
                        Compare:=CompareMethod.Text)
                        )
                                )

            Catch ex As Exception
                Buffer = 0
            End Try



        End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _buffer As Double = 0

    ''' <summary>
    ''' Buffer
    ''' Buffer in m
    ''' </summary>
    ''' <returns></returns>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Buffer")>
    <DisplayName("Buffer")>
    <Category(catResults)>
    <DefaultValue(0)>
    <Browsable(False)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "unit=' m'")>
    Public Property Buffer As Double
        Get
            Return _buffer
        End Get
        Set

            _buffer = Value

            Me.Distances.GetDistance(
                       Distances:=Me.Distances,
                       FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                       FOCUSswWaterBody:=_FOCUSswWaterBody,
                       ApplnMethodStep03:=_applnMethodStep03,
                       Buffer:=_buffer)

            CalcDriftPercentDistances(
                        Distances:=Me.Distances,
                        Nozzle:=Me.Nozzle)
        End Set
    End Property

#End Region

#Region "    Nozzle"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _nozzle As eNozzles = eNozzles.not_defined

    ''' <summary>
    ''' Drift reducing nozzle in %
    ''' </summary>
    ''' <returns></returns>
    <Category(catResults)>
    <DisplayName(
        "Nozzle")>
    <Description(
        "Drift reducing nozzle in %" & vbCrLf &
        "")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(CInt(eBufferWidth.FOCUSStep03))>
    Public Property Nozzle As eNozzles
        Get
            Return _nozzle
        End Get
        Set
            _nozzle = Value

            CalcDriftPercentDistances(
                        Distances:=Me.Distances,
                        Nozzle:=Me.Nozzle)

        End Set
    End Property

#End Region

    ''' <summary>
    ''' Total drift reduction compared to Step03
    ''' "Buffer + Nozzle in percent, single appln.
    ''' </summary>
    ''' <returns></returns>
    <Category(catResults)>
    <DisplayName(
    "Total Reduction")>
    <Description(
    "Total drift reduction compared to Step03" & vbCrLf &
    "Buffer + Nozzle in percent, single appln")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '0'" &
    "|unit='%'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public ReadOnly Property TotalReduction As Double
        Get

            Dim Step03Drift As Double
            Dim Step04Drift As Double

            If _nozzle = 0 AndAlso _buffer = 0 Then
                Return 0
            Else
                Step03Drift =
                CalcDriftPercent(
                noOfApplns:=_noOfApplns,
                FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                FOCUSswWaterBody:=_FOCUSswWaterBody,
                ApplnMethodStep03:=_applnMethodStep03)

                Step04Drift =
                CalcDriftPercent(
                noOfApplns:=_noOfApplns,
                FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                FOCUSswWaterBody:=_FOCUSswWaterBody,
                ApplnMethodStep03:=_applnMethodStep03,
                Nozzle:=_nozzle,
                Buffer:=_buffer)


                If Step03Drift > 0 AndAlso Step03Drift >= Step04Drift Then
                    Return Math.Ceiling(a:=100 - (Step04Drift * 100 / Step03Drift))
                Else
                    Return 0
                End If

            End If


        End Get
    End Property

#End Region

#Region "    RAC"

    <Category(catResults)>
    <DisplayName(
    "RAC")>
    <Description(
    "Regulatory acceptable concentration")>
    <TypeConverter(GetType(DblConv))>
    <RefreshProperties(RefreshProperties.All)>
    <DefaultValue(0)>
    <Browsable(True)>
    Public Property RAC As Double = 0


    <Category(catResults)>
    <RefreshProperties(RefreshProperties.All)>
    <DisplayName(
    "Buffer : Nozzle")>
    <Description(
    "Min Nozzle in % to beat RAC" & vbCrLf &
    "for Step03, 5m ,10, 15m and 20m Buffer")>
    Public ReadOnly Property GUINozzlePerBuffer4RAC As String()
        Get


            Dim temp As Double
            Dim out As New List(Of Double)
            Dim gui As New List(Of String)

            If _rate = 0 OrElse RAC = 0 OrElse _waterDepth < 0 Then
                Return {}
            Else

                For buffer As Integer = 0 To 20 Step 5

                    temp =
                        CalcDriftPercent(
                            noOfApplns:=_noOfApplns,
                            FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                            FOCUSswWaterBody:=_FOCUSswWaterBody,
                            ApplnMethodStep03:=_applnMethodStep03,
                            Buffer:=buffer)

                    If Not Double.IsNaN(temp) Then
                        temp = _rate * temp / _waterDepth
                        temp = 100 - (RAC / temp * 100)

                        If temp > 0 Then

                            temp = Math.Ceiling(temp)
                            out.Add(Math.Ceiling(temp))

                            If buffer = 0 Then
                                gui.Add("Step03 : min " & temp & "%")
                            Else
                                gui.Add(buffer.ToString("00") & "m    : min " & temp & "%")
                            End If

                        Else

                            If buffer = 0 Then
                                gui.Add("Step03 : - ")
                            Else
                                gui.Add(buffer.ToString("00") & "m    : - ")
                            End If

                            out.Add(Double.NaN)
                        End If

                    Else
                        out.Add(Double.NaN)
                    End If
                Next
            End If

            NozzlePerBuffer4RAC = out.ToArray
            Return gui.ToArray

        End Get
    End Property

    <XmlIgnore> <ScriptIgnore>
    <Category(catResults)>
    <RefreshProperties(RefreshProperties.All)>
    <Browsable(False)>
    Public Property NozzlePerBuffer4RAC As Double()

#End Region

#End Region


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const catBBCH As String = " 02 BBCH"

#Region "   BBCH Test"


    Public SWASH_BBCHout As String() = {}


    Private _BBCH1st As eBBCH = eBBCH._0

    <Category(catBBCH)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property BBCH1st As eBBCH
        Get
            Return _BBCH1st
        End Get
        Set


            If _FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined OrElse
               _FOCUSswScenario = eFOCUSswScenario.not_defined Then
                Exit Property
            End If


            _BBCH1st = Value

            Dim test As String() = {}

            test =
                Filter(
                    Source:=SWASH_BBCHout,
                    Match:=GetBBCHSearchStringsFromCrop(
                        FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                        Season:=eSeason.first).First)

            test =
                Filter(Source:=test, Match:=_FOCUSswScenario.ToString)

            test =
                Filter(Source:=test, Match:=vbTab & _BBCH1st & vbTab)

            test = test.First.Split(vbTab)(6).Split(".")

            Date1st = New Date(year:=test(2), month:=test(1), day:=test(0))

            BBCHEnd = _BBCH1st

        End Set
    End Property

    <Category(catBBCH)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property Date1st As New Date


    Private _BBCHEnd As eBBCH = eBBCH._0

    <Category(catBBCH)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property BBCHEnd As eBBCH
        Get
            Return _BBCHEnd
        End Get
        Set
            If Value >= _BBCH1st Then
                _BBCHEnd = Value
            End If

            Dim test As String() = {}

            test =
                Filter(
                    Source:=SWASH_BBCHout,
                    Match:=GetBBCHSearchStringsFromCrop(
                        FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                        Season:=eSeason.first).First)

            test =
                Filter(Source:=test, Match:=_FOCUSswScenario.ToString)

            test =
                Filter(Source:=test, Match:=vbTab & _BBCHEnd & vbTab)

            test = test.First.Split(vbTab)(6).Split(".")

            Datelast = New Date(year:=test(2), month:=test(1), day:=test(0))


        End Set
    End Property

    <Category(catBBCH)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property Datelast As New Date

    <Category(catBBCH)>
    <RefreshProperties(RefreshProperties.All)>
    Public ReadOnly Property Window As Integer
        Get
            If _noOfApplns = 1 Then
                Return 30
            Else
                Return 30 + _noOfApplns * _interval
            End If
        End Get
    End Property




    <Category(catBBCH)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property PHI As Integer = 0


#End Region


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const catSpecial As String = " 03 Special"

#Region "    Special"

    <Category(catSpecial)>
    <DisplayName(
    "Aqua Met Factor")>
    <Description(
    "Factor for calculate aquatic metabolites")>
    <TypeConverter(GetType(DblConv))>
    <RefreshProperties(RefreshProperties.All)>
    <DefaultValue(0)>
    <Browsable(True)>
    Public Property AquaMetFactor As Double = 0

    <Category(catSpecial)>
    <DisplayName(
    "DT50sw")>
    <Description(
    "DT50sw for calculation of multi applns. ponds")>
    <TypeConverter(GetType(DblConv))>
    <RefreshProperties(RefreshProperties.All)>
    <DefaultValue(100)>
    <Browsable(True)>
    Public Property DT50sw As Double = 100


#Region "    Interval : Time between applns. in days"

    ''' <summary>
    ''' GUI : Time between applns. in days
    ''' </summary>
    ''' <returns></returns>
    <Category(catSpecial)>
    <DisplayName(
    "Appln. Interval")>
    <Description(
    "Time between applns. in days" & vbCrLf &
    "for calculation of multi applns. ponds")>
    <TypeConverter(GetType(dropDownList))>
    <RefreshProperties(RefreshProperties.All)>
    <DefaultValue(" - ")>
    <Browsable(True)>
    <XmlIgnore> <ScriptIgnore>
    Public Property IntervalGUI As String
        Get

            dropDownList.dropDownEntries =
                    {
                    " - ",
                    "5",
                    "7",
                    "10",
                    "12",
                    "14",
                    "21",
                    "28",
                    "30",
                    "42",
                    "50"
                    }

            If Interval > 0 AndAlso Interval <= 365 Then
                Return Interval.ToString()
            Else
                Return " - "
            End If

        End Get
        Set

            If _noOfApplns = eNoOfApplns.not_defined OrElse
                    _noOfApplns = eNoOfApplns._01 Then
                Interval = 0
                Exit Property
            End If

            If Value <> " - " Then

                Try
                    _interval = Integer.Parse(Trim(Value))
                Catch ex As Exception
                    _interval = 0
                End Try

            Else

                Interval = 0

            End If

        End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _interval As Integer = 7

    ''' <summary>
    ''' Time between applns. in days
    ''' </summary>
    ''' <returns></returns>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Time between applns. in days")>
    <DisplayName("Interval")>
    <Category(catSpecial)>
    <DefaultValue(7)>
    <Browsable(False)>
    <TypeConverter(GetType(IntConv))>
    <AttributeProvider(
    "unit=' days'")>
    Public Property Interval As Integer
        Get
            Return _interval
        End Get
        Set
            _interval = Value
        End Set
    End Property


    <Category(catSpecial)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '0.0000'")>
    <Description(
    "(1 - e^(-1 * noOfApplns * _interval * k)) /" & vbCrLf &
    "(1 - e^(-1              * _interval * k))")>
    Public ReadOnly Property MAF As Double
        Get
            If _noOfApplns = eNoOfApplns._01 OrElse DT50sw = 0 OrElse _interval = 0 Then
                Return 1
            Else
                Try

                    Dim k As Double
                    Dim out As Double
                    Dim noOfApplns As Double

                    noOfApplns = _noOfApplns + 1
                    k = Math.Log(2) / DT50sw


                    out = (1 - Math.Exp(-1 * noOfApplns * _interval * k)) / (1 - Math.Exp(-1 * _interval * k))
                    Return (1 - Math.Exp(-1 * noOfApplns * _interval * k)) / (1 - Math.Exp(-1 * _interval * k))
                Catch ex As Exception
                    Return 1
                End Try
            End If


        End Get
    End Property

    <Category(catSpecial)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '0.0000'")>
    Public ReadOnly Property Loading As Double
        Get

            If Step03 <> " - " AndAlso _rate > 0 Then
                Return _rate * CalcDriftPercent(
                                    noOfApplns:=_noOfApplns,
                                    FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                                    FOCUSswWaterBody:=_FOCUSswWaterBody,
                                    ApplnMethodStep03:=_applnMethodStep03)
            Else
                Return 0
            End If


        End Get
    End Property

#End Region


#End Region


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const catOut As String = " 04 Output "

#Region "    Output"

#Region "    Buffer & Nozzles"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _outBuffer As Double() =
                    {
                    0,
                    5,
                    10,
                    20
                    }

    ''' <summary>
    ''' OutBuffer
    ''' OutBuffer in m
    ''' </summary>
    ''' <returns></returns>
    <RefreshProperties(RefreshProperties.All)>
    <DisplayName(
    "Buffers")>
    <Description(
    "Buffer distances for output in m" & vbCrLf &
    "std.: Step03, 5m, 10m, 20m")>
    <Category(catOut)>
    <DefaultValue(0)>
    <Browsable(True)>
    Public Property OutBuffer As Double()
        Get
            Return _outBuffer
        End Get
        Set
            _outBuffer = Value
        End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _outNozzle As eNozzles() =
                        {
                        eNozzles._0,
                        eNozzles._50,
                        eNozzles._75,
                        eNozzles._90
                        }

    ''' <summary>
    ''' OutNozzle
    ''' OutNozzle in m
    ''' </summary>
    ''' <returns></returns>
    <RefreshProperties(RefreshProperties.All)>
    <DisplayName(
    "Nozzles")>
    <Description(
    "Nozzles for output in %" & vbCrLf &
    "std.: Step03 (0%), 50%, 75%, 90%")>
    <Category(catOut)>
    <DefaultValue(0)>
    <Browsable(True)>
    Public Property OutNozzle As eNozzles()
        Get
            Return _outNozzle
        End Get
        Set
            _outNozzle = Value
        End Set
    End Property


#End Region

#Region "    DGR & PMT"

    <Category(catOut)>
    <DefaultValue(0)>
    <TypeConverter(GetType(IntConv))>
    Public Property DGR As Integer = 0

    <Category(catOut)>
    <DefaultValue(0)>
    <TypeConverter(GetType(IntConv))>
    Public Property PMT As Integer = 0

#End Region


    Public Class out

        Public Sub New()

        End Sub

        Public Enum eEntry

            Entry
            MaxSum
            ACT01
            ACT02
            ACT03
            ACT04
            ACT07
            ACT14
            ACT21
            ACT28
            ACT42
            ACT50
            ACT00
            TWA01
            TWA02
            TWA03
            TWA04
            TWA07
            TWA14
            TWA21
            TWA28
            TWA42
            TWA50
            TWA00

        End Enum

        Public Property Entry As eEntry = eEntry.Entry


        Public Enum eParMet
            Par
            Met
        End Enum

        Public Property ParMet As eParMet = eParMet.Par

        Public Property DGR As Integer = 0

        Public Property PMT As Integer = 0

        Public Property WaterBody As eFOCUSswWaterBody = eFOCUSswWaterBody.not_defined

        Public Property Buffer As Integer = 0

        Public Property Nozzle As Integer = 0

        Public Property NoOfApplns As Integer = 0

        Public Property Interval As Integer = 0

        Public Property Rate As Double = 0

        Public Enum eMainRoute
            R
            S
            D
        End Enum

        Public Property MainRoute As eMainRoute = eMainRoute.D

        Public Property MAX As String = ""

        Public Property PECsw As Double = 0

        Public Property PECswRACsw As Double = 0

        Public Property MaxDateYear As Integer = 0

        Public Property MaxDateMonth As Integer = 0

        Public Property MaxDateDay As Integer = 0


        Public Property PECsed As Double = 0

        Public Property PECsedRACsed As Double = 0

        Public Property FUllPath As String = ""


    End Class


#End Region


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const catUse As String = " 98 Use "

#Region "    Use"

    '#Region "    BBCH"

    '    <Category(catUse)>
    '    <DisplayName(
    '    "BBCH Start")>
    '    <Description(
    '    "" & vbCrLf &
    '    "")>
    '    <RefreshProperties(RefreshProperties.All)>
    '    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    '    <Browsable(True)>
    '    <[ReadOnly](False)>
    '    <TypeConverter(GetType(IntConv))>
    '    <DefaultValue(0)>
    '    Public Property BBCHStart As Integer = 0

    '    <Category(catUse)>
    '    <DisplayName(
    '    "     End")>
    '    <Description(
    '    "" & vbCrLf &
    '    "")>
    '    <RefreshProperties(RefreshProperties.All)>
    '    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    '    <Browsable(True)>
    '    <[ReadOnly](False)>
    '    <TypeConverter(GetType(IntConv))>
    '    <DefaultValue(0)>
    '    Public Property BBCHend As Integer = 0

    '#End Region

#End Region

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const catCalculations As String = " 05 Calculations "

#Region "    Calculations"

#Region "    functions"

    ''' <summary>
    ''' calculation of FOCUS drift values
    ''' </summary>   
    ''' <remarks></remarks>
    '<DebuggerStepThrough>
    Public Sub CalcDriftPercentDistances(ByRef Distances As Distances, Nozzle As eNozzles)

        'check inputs
        If NoOfApplns = eNoOfApplns.not_defined OrElse
           FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined OrElse
           FOCUSswWaterBody = eFOCUSswWaterBody.not_defined Then

            Exit Sub

        End If


        With Me.Regression.SingleApplnRegression

            If Distances.Farthest2EdgeOfWB < .HingePoint Then
                Distances.FarthestDriftPercentSingle = .A * (Distances.Farthest2EdgeOfWB ^ .B)
            Else
                Distances.FarthestDriftPercentSingle = .C * (Distances.Farthest2EdgeOfWB ^ .D)
            End If

            If Distances.Closest2EdgeOfField < .HingePoint Then
                Distances.NearestDriftPercentSingle = .A * (Distances.Closest2EdgeOfField ^ .B)
            Else
                Distances.NearestDriftPercentSingle = .C * (Distances.Closest2EdgeOfField ^ .D)
            End If

            Try

                If Distances.Farthest2EdgeOfWB < .HingePoint Then
                    Distances.AverageDriftPercentSingle = (.A / (.B + 1) * (Distances.Farthest2EdgeOfWB ^ (.B + 1) - Distances.Closest2EdgeOfField ^ (.B + 1))) /
                                           (Distances.Farthest2EdgeOfWB - Distances.Closest2EdgeOfField)
                ElseIf Distances.Closest2EdgeOfField > .HingePoint Then
                    Distances.AverageDriftPercentSingle = .C / (.D + 1) * (Distances.Farthest2EdgeOfWB ^ (.D + 1) - Distances.Closest2EdgeOfField ^ (.D + 1)) /
                                          (Distances.Farthest2EdgeOfWB - Distances.Closest2EdgeOfField)
                Else
                    Distances.AverageDriftPercentSingle = (.A / (.B + 1) * (.HingePoint ^ (.B + 1) - Distances.Closest2EdgeOfField ^ (.B + 1)) + .C / (.D + 1) * (Distances.Farthest2EdgeOfWB ^ (.D + 1) - .HingePoint ^ (.D + 1))) * 1 /
                                           (Distances.Farthest2EdgeOfWB - Distances.Closest2EdgeOfField)
                End If

            Catch ex As Exception

                Throw New _
                        ArithmeticException(
                        message:="Error during drift calc.",
                        innerException:=ex)

            End Try

            ' if stream then apply upstream catchment factor 1.2
            Distances.NearestDriftPercentSingle *=
                IIf(
                    Expression:=_FOCUSswWaterBody = eFOCUSswWaterBody.ST,
                    TruePart:=1.2,
                    FalsePart:=1)


            'drift reducing nozzles?
            If Nozzle > eNozzles._0 Then
                Distances.NearestDriftPercentSingle *= (100 - _nozzle) / 100
            End If

            Distances.FarthestDriftPercentSingle *=
               IIf(
                   Expression:=_FOCUSswWaterBody = eFOCUSswWaterBody.ST,
                   TruePart:=1.2,
                   FalsePart:=1)


            'drift reducing nozzles?
            If Nozzle > eNozzles._0 Then
                Distances.FarthestDriftPercentSingle *= (100 - _nozzle) / 100
            End If

            Distances.AverageDriftPercentSingle *=
               IIf(
                   Expression:=_FOCUSswWaterBody = eFOCUSswWaterBody.ST,
                   TruePart:=1.2,
                   FalsePart:=1)


            'drift reducing nozzles?
            If Nozzle > eNozzles._0 Then
                Distances.AverageDriftPercentSingle *= (100 - _nozzle) / 100
            End If

        End With


        With Regression.MultiApplnRegression

            If Distances.Farthest2EdgeOfWB < .HingePoint Then
                Distances.FarthestDriftPercentMulti = .A * (Distances.Farthest2EdgeOfWB ^ .B)
            Else
                Distances.FarthestDriftPercentMulti = .C * (Distances.Farthest2EdgeOfWB ^ .D)
            End If

            If Distances.Closest2EdgeOfField < .HingePoint Then
                Distances.NearestDriftPercentMulti = .A * (Distances.Closest2EdgeOfField ^ .B)
            Else
                Distances.NearestDriftPercentMulti = .C * (Distances.Closest2EdgeOfField ^ .D)
            End If

            Try

                If Distances.Farthest2EdgeOfWB < .HingePoint Then
                    Distances.AverageDriftPercentMulti = (.A / (.B + 1) * (Distances.Farthest2EdgeOfWB ^ (.B + 1) - Distances.Closest2EdgeOfField ^ (.B + 1))) /
                                           (Distances.Farthest2EdgeOfWB - Distances.Closest2EdgeOfField)
                ElseIf Distances.Closest2EdgeOfField > .HingePoint Then
                    Distances.AverageDriftPercentMulti = .C / (.D + 1) * (Distances.Farthest2EdgeOfWB ^ (.D + 1) - Distances.Closest2EdgeOfField ^ (.D + 1)) /
                                          (Distances.Farthest2EdgeOfWB - Distances.Closest2EdgeOfField)
                Else
                    Distances.AverageDriftPercentMulti = (.A / (.B + 1) * (.HingePoint ^ (.B + 1) - Distances.Closest2EdgeOfField ^ (.B + 1)) + .C / (.D + 1) * (Distances.Farthest2EdgeOfWB ^ (.D + 1) - .HingePoint ^ (.D + 1))) * 1 /
                                           (Distances.Farthest2EdgeOfWB - Distances.Closest2EdgeOfField)
                End If

            Catch ex As Exception

                Throw New _
                        ArithmeticException(
                        message:="Error during drift calc.",
                        innerException:=ex)

            End Try

        End With

        ' if stream then apply upstream catchment factor 1.2
        Distances.NearestDriftPercentMulti *=
                    IIf(
                        Expression:=_FOCUSswWaterBody = eFOCUSswWaterBody.ST,
                        TruePart:=1.2,
                        FalsePart:=1)

        'drift reducing nozzles?
        If Nozzle > eNozzles._0 Then
            Distances.NearestDriftPercentMulti *= (100 - _nozzle) / 100
        End If

        Distances.NearestDriftPercentMulti *=
                   IIf(
                       Expression:=_FOCUSswWaterBody = eFOCUSswWaterBody.ST,
                       TruePart:=1.2,
                       FalsePart:=1)


        'drift reducing nozzles?
        If Nozzle > eNozzles._0 Then
            Distances.FarthestDriftPercentMulti *= (100 - _nozzle) / 100
        End If

        Distances.AverageDriftPercentMulti *=
                   IIf(
                       Expression:=_FOCUSswWaterBody = eFOCUSswWaterBody.ST,
                       TruePart:=1.2,
                       FalsePart:=1)


        'drift reducing nozzles?
        If Nozzle > eNozzles._0 Then
            Distances.AverageDriftPercentMulti *= (100 - _nozzle) / 100
        End If



    End Sub


    ''' <summary>
    ''' calculation of FOCUS drift values
    ''' </summary>
    ''' <param name="noOfApplns">
    ''' # of applications, 1 - 8, as enum
    ''' </param>
    ''' <param name="FOCUSswDriftCrop">
    ''' FOCUS crop as enum
    ''' </param>
    ''' <param name="FOCUSswWaterBody">
    ''' Ditch, pond or stream, as enum
    ''' </param>
    ''' <param name="nozzle">
    ''' Drift red. nozzle as percent 0 - 100
    ''' </param>   
    ''' <param name="Buffer">
    ''' Buffer width in m, -1 = FOCUS std. width
    ''' </param>
    ''' <remarks></remarks>
    <DebuggerStepThrough>
    Public Function CalcDriftPercent(
                            noOfApplns As eNoOfApplns,
                            FOCUSswDriftCrop As eFOCUSswDriftCrop,
                            FOCUSswWaterBody As eFOCUSswWaterBody,
                            ApplnMethodStep03 As eApplnMethodStep03,
                   Optional Nozzle As Integer = 0,
                   Optional Buffer As Double = 0) As Double

        'check inputs
        If noOfApplns = eNoOfApplns.not_defined OrElse
           FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined OrElse
           FOCUSswWaterBody = eFOCUSswWaterBody.not_defined Then

            Return 0

        End If

        If Nozzle = CInt(eNozzles.not_defined) Then
            Nozzle = eNozzles._0
        End If

        'drift for output
        Dim driftPercent As Double = Double.NaN

        'Application
        Dim Ganzelmeier As eGanzelmeier

        Dim Distance As New Distances

        Dim Regression As New Regression

        Dim A As Double : Dim B As Double
        Dim C As Double : Dim D As Double
        Dim HingePoint As Double

        Distance.GetDistance(
            Distances:=Distance,
            FOCUSswDriftCrop:=FOCUSswDriftCrop,
            FOCUSswWaterBody:=FOCUSswWaterBody,
            ApplnMethodStep03:=ApplnMethodStep03,
            Buffer:=Buffer)

        UpdateGanzelmeier(
            Ganzelmeier:=Ganzelmeier,
            FOCUSswDriftCrop:=FOCUSswDriftCrop,
            ApplnMethodStep03:=ApplnMethodStep03)

        Regression.Update(
            Regression:=Regression,
            Ganzelmeier:=Ganzelmeier,
            NoOfApplns:=noOfApplns)

        'get regression parameter

        If noOfApplns = eNoOfApplns._01 Then

            With Regression.SingleApplnRegression

                A = .A

                B = .B

                C = .C

                D = .D

                HingePoint = .HingePoint

            End With

        Else

            With Regression.MultiApplnRegression

                A = .A

                B = .B

                C = .C

                D = .D

                HingePoint = .HingePoint

            End With

        End If


        With Distance

            Try

                If .Farthest2EdgeOfWB < HingePoint Then
                    driftPercent = (A / (B + 1) * (.Farthest2EdgeOfWB ^ (B + 1) - .Closest2EdgeOfField ^ (B + 1))) /
                        (.Farthest2EdgeOfWB - .Closest2EdgeOfField)
                ElseIf .Closest2EdgeOfField > HingePoint Then
                    driftPercent = C / (D + 1) * (.Farthest2EdgeOfWB ^ (D + 1) - .Closest2EdgeOfField ^ (D + 1)) /
                        (.Farthest2EdgeOfWB - .Closest2EdgeOfField)
                Else
                    driftPercent = (A / (B + 1) * (HingePoint ^ (B + 1) - .Closest2EdgeOfField ^ (B + 1)) + C / (D + 1) * (.Farthest2EdgeOfWB ^ (D + 1) - HingePoint ^ (D + 1))) * 1 /
                        (.Farthest2EdgeOfWB - .Closest2EdgeOfField)
                End If

            Catch ex As Exception

                Throw New _
                    ArithmeticException(
                    message:="Error during drift calc.",
                    innerException:=ex)

            End Try

        End With

        driftPercent *=
                IIf(
                    Expression:=_FOCUSswWaterBody = eFOCUSswWaterBody.ST,
                    TruePart:=1.2,
                    FalsePart:=1)


        'drift reducing nozzles?
        If Nozzle > eNozzles._0 Then
            driftPercent *= (100 - Nozzle) / 100
        End If

        Return driftPercent

    End Function

    ''' <summary>
    ''' get Ganzelmeier Crop Group
    ''' </summary>
    ''' <param name="Ganzelmeier">
    ''' enum to change
    ''' </param>
    ''' <param name="FOCUSswDriftCrop">
    ''' 
    ''' </param>
    ''' <param name="ApplnMethodStep03">
    ''' 
    ''' </param>
    Public Sub UpdateGanzelmeier(
                    ByRef Ganzelmeier As eGanzelmeier,
                    FOCUSswDriftCrop As eFOCUSswDriftCrop,
                    ApplnMethodStep03 As eApplnMethodStep03)


        If FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined OrElse
           ApplnMethodStep03 = eApplnMethodStep03.not_defined Then
            Ganzelmeier = eGanzelmeier.not_defined
            Exit Sub
        End If

        Select Case ApplnMethodStep03

            Case eApplnMethodStep03.SI, eApplnMethodStep03.GR
                Ganzelmeier = eGanzelmeier.noDrift

            Case eApplnMethodStep03.GS
                Ganzelmeier = eGanzelmeier.ArableCrops

            Case eApplnMethodStep03.AA
                Ganzelmeier = eGanzelmeier.AerialAppln

            Case eApplnMethodStep03.AB

                Select Case FOCUSswDriftCrop

                    Case _
                        eFOCUSswDriftCrop.HP,
                        eFOCUSswDriftCrop.CI,
                        eFOCUSswDriftCrop.OL

                        Ganzelmeier = eGanzelmeier.FruitCrops_Late


                    Case eFOCUSswDriftCrop.PFE
                        Ganzelmeier = eGanzelmeier.FruitCrops_Early

                    Case eFOCUSswDriftCrop.PFL
                        Ganzelmeier = eGanzelmeier.FruitCrops_Late

                    Case eFOCUSswDriftCrop.VI
                        Ganzelmeier = eGanzelmeier.Vines_Early

                    Case eFOCUSswDriftCrop.VIL
                        Ganzelmeier = eGanzelmeier.Vines_Late

                End Select

            Case eApplnMethodStep03.HL
                Ganzelmeier = eGanzelmeier.ArableCrops

            Case eApplnMethodStep03.HH
                Ganzelmeier = eGanzelmeier.FruitCrops_Late

            Case eApplnMethodStep03.not_defined
                Ganzelmeier = eGanzelmeier.not_defined

        End Select

        If Ganzelmeier <> eGanzelmeier.not_defined Then

            Me.Regression.Update(
                    Regression:=Me.Regression,
                    Ganzelmeier:=Ganzelmeier,
                    NoOfApplns:=_noOfApplns)

        End If

    End Sub


#End Region


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _ganzelmeier As eGanzelmeier = eGanzelmeier.not_defined

    ''' <summary>
    ''' Ganzelmeier
    ''' Ganzelmeier crop group
    ''' based on selected FOCUS crop
    ''' </summary>
    <Category(catCalculations)>
    <DisplayName(
    "Ganzelmeier Crop")>
    <Description(
    "Ganzelmeier crop group" & vbCrLf &
    "base for further calculations")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](True)>
    <XmlIgnore> <ScriptIgnore>
    Public Property Ganzelmeier As eGanzelmeier
        Get
            Return _ganzelmeier
        End Get
        Set(value As eGanzelmeier)
            _ganzelmeier = value
        End Set
    End Property


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _Regression As New Regression

    ''' <summary>
    ''' Regression parameters for
    ''' drift curve calculation
    ''' </summary>
    ''' <returns></returns>
    <XmlIgnore> <ScriptIgnore>
    <Category(catCalculations)>
    <DisplayName(
    "Regression")>
    <Description(
    "Regression parameters for" & vbCrLf &
    "drift curve calculation")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](True)>
    Public Property Regression As Regression
        Get
            Return _Regression
        End Get
        Set
            _Regression = Value
        End Set
    End Property


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _Distances As New Distances

    ''' <summary>
    ''' Distances near/fare field
    ''' </summary>
    ''' <returns></returns>
    <XmlIgnore> <ScriptIgnore>
    <Category(catCalculations)>
    <DisplayName(
    "Distances")>
    <Description(
    "" & vbCrLf &
    "")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](True)>
    Public Property Distances As Distances
        Get
            Return _Distances
        End Get
        Set
            _Distances = Value
        End Set
    End Property

#End Region

End Class

#Region "    Calculations"

<TypeConverter(GetType(PropGridConverter))>
Public Class Regression

    Public Sub New()

    End Sub

    Public Sub New(Ganzelmeier As eGanzelmeier, NoOfApplns As eNoOfApplns)
        Update(
            Regression:=Me,
            Ganzelmeier:=Ganzelmeier,
            NoOfApplns:=NoOfApplns)
    End Sub

    <DisplayName("Single appln.")>
    Public Property SingleApplnRegression As New RegressionBase

    <DisplayName("Multi applns.")>
    Public Property MultiApplnRegression As New RegressionBase

    Public Sub Update(
                ByRef Regression As Regression,
                      Ganzelmeier As eGanzelmeier,
                      NoOfApplns As eNoOfApplns)

        If Ganzelmeier = eGanzelmeier.not_defined OrElse
            NoOfApplns = eNoOfApplns.not_defined Then

            With Regression.SingleApplnRegression

                .A = 0
                .B = 0
                .C = 0
                .D = 0
                .HingePoint = 0

            End With

            With Regression.MultiApplnRegression

                .A = 0
                .B = 0
                .C = 0
                .D = 0
                .HingePoint = 0

            End With

        Else

            With Regression.SingleApplnRegression

                .A = RegressionA(
                    Ganzelmeier,
                     eNoOfApplns._01)

                .B = RegressionB(
                    Ganzelmeier,
                    eNoOfApplns._01)

                .C = RegressionC(
                    Ganzelmeier,
                    eNoOfApplns._01)

                .D = RegressionD(
                    Ganzelmeier,
                    eNoOfApplns._01)

                .HingePoint = HingePointDB(
                    Ganzelmeier,
                    eNoOfApplns._01)

            End With

            If NoOfApplns > eNoOfApplns._01 Then

                With Regression.MultiApplnRegression

                    .A = RegressionA(
                        Ganzelmeier,
                         NoOfApplns)

                    .B = RegressionB(
                        Ganzelmeier,
                        NoOfApplns)

                    .C = RegressionC(
                        Ganzelmeier,
                        NoOfApplns)

                    .D = RegressionD(
                        Ganzelmeier,
                        NoOfApplns)

                    .HingePoint = HingePointDB(
                        Ganzelmeier,
                        NoOfApplns)

                End With

            Else

                With Regression.MultiApplnRegression

                    .A = 0
                    .B = 0
                    .C = 0
                    .D = 0
                    .HingePoint = 0

                End With

            End If

        End If

    End Sub

End Class


<TypeConverter(GetType(PropGridConverter))>
Public Class RegressionBase

    Public Sub New()

    End Sub

    <Category()>
    <Description(
        "" & vbCrLf &
        "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '0.0000'")>
    Public Property A As Double = 0

    <Category()>
    <Description(
        "" & vbCrLf &
        "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '0.0000'")>
    Public Property B As Double = 0

    <Category()>
    <Description(
        "" & vbCrLf &
        "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '0.0000'")>
    Public Property C As Double = 0

    <Category()>
    <Description(
        "" & vbCrLf &
        "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '0.0000'")>
    Public Property D As Double = 0

    <Category()>
    <DisplayName(
        "Hinge Point")>
    <Description(
        "Distance limit for each regression in m" & vbCrLf &
        "to switch from A/B to C/D")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    <TypeConverter(GetType(IntConv))>
    <AttributeProvider(
    "format= '0.00'")>
    Public Property HingePoint As Double = 0

End Class

<TypeConverter(GetType(PropGridConverter))>
Public Class Distances

    Public Sub New()

    End Sub

    Public Enum eFOCUSStdDistancesMember

        crop
        waterbody
        edgeField2topBank
        topBank2edgeWaterbody

    End Enum

    '<DebuggerStepThrough()>
    Public Sub GetDistance(
                    ByRef Distances As Distances,
                          FOCUSswDriftCrop As eFOCUSswDriftCrop,
                          FOCUSswWaterBody As eFOCUSswWaterBody,
                          ApplnMethodStep03 As eApplnMethodStep03,
                          Buffer As Double)

        Dim SearchString As String
        Dim FOCUSstdDistancesEntry As String = ""

        If FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined OrElse
           FOCUSswWaterBody = eFOCUSswWaterBody.not_defined OrElse
           ApplnMethodStep03 = eApplnMethodStep03.not_defined Then

            With Distances

                .Crop2Bank = 0
                .Bank2Water = 0
                .Closest2EdgeOfField = 0
                .Farthest2EdgeOfWB = 0

            End With

            Exit Sub

        End If

        Try

            Select Case ApplnMethodStep03

                Case eApplnMethodStep03.HH,
                     eApplnMethodStep03.HL,
                     eApplnMethodStep03.AA

                    SearchString =
                        ApplnMethodStep03.ToString & "|" &
                        FOCUSswWaterBody.ToString

                Case Else

                    SearchString =
                        FOCUSswDriftCrop.ToString & "|" &
                        FOCUSswWaterBody.ToString

            End Select

            FOCUSstdDistancesEntry =
                        Filter(
                            Source:=FOCUSStdDistances,
                            Match:=SearchString,
                            Include:=True,
                            Compare:=CompareMethod.Text).First

            With Distances

                If Buffer <> 0 Then

                    .Crop2Bank = Buffer
                    .Bank2Water = 0

                    .Closest2EdgeOfField = Buffer

                Else

                    .Crop2Bank =
                        FOCUSstdDistancesEntry.Split({"|"c})(eFOCUSStdDistancesMember.edgeField2topBank)

                    .Bank2Water =
                        FOCUSstdDistancesEntry.Split({"|"c})(eFOCUSStdDistancesMember.topBank2edgeWaterbody)

                    .Closest2EdgeOfField = _crop2Bank + _bank2Water

                End If

                If FOCUSswWaterBody = eFOCUSswWaterBody.PO Then
                    .Farthest2EdgeOfWB = .Closest2EdgeOfField + 30
                Else
                    .Farthest2EdgeOfWB = .Closest2EdgeOfField + 1
                End If

            End With

        Catch ex As Exception

        End Try

    End Sub


    Public FOCUSStdDistances As String() =
        {
            "CS|DI|0.5|0.5",
            "CS|ST|0.5|1",
            "CS|PO|0.5|3",
            "CW|DI|0.5|0.5",
            "CW|ST|0.5|1",
            "CW|PO|0.5|3",
            "GA|DI|0.5|0.5",
            "GA|ST|0.5|1",
            "GA|PO|0.5|3",
            "OS|DI|0.5|0.5",
            "OS|ST|0.5|1",
            "OS|PO|0.5|3",
            "OW|DI|0.5|0.5",
            "OW|ST|0.5|1",
            "OW|PO|0.5|3",
            "VB|DI|0.5|0.5",
            "VB|ST|0.5|1",
            "VB|PO|0.5|3",
            "VF|DI|0.5|0.5",
            "VF|ST|0.5|1",
            "VF|PO|0.5|3",
            "VL|DI|0.5|0.5",
            "VL|ST|0.5|1",
            "VL|PO|0.5|3",
            "VR|DI|0.5|0.5",
            "VR|ST|0.5|1",
            "VR|PO|0.5|3",
            "PS|DI|0.8|0.5",
            "PS|ST|0.8|1",
            "PS|PO|0.8|3",
            "SY|DI|0.8|0.5",
            "SY|ST|0.8|1",
            "SY|PO|0.8|3",
            "SB|DI|0.8|0.5",
            "SB|ST|0.8|1",
            "SB|PO|0.8|3",
            "SU|DI|0.8|0.5",
            "SU|ST|0.8|1",
            "SU|PO|0.8|3",
            "CO|DI|0.8|0.5",
            "CO|ST|0.8|1",
            "CO|PO|0.8|3",
            "FB|DI|0.8|0.5",
            "FB|ST|0.8|1",
            "FB|PO|0.8|3",
            "LG|DI|0.8|0.5",
            "LG|ST|0.8|1",
            "LG|PO|0.8|3",
            "MZ|DI|0.8|0.5",
            "MZ|ST|0.8|1",
            "MZ|PO|0.8|3",
            "TB|DI|1|0.5",
            "TB|ST|1|1",
            "TB|PO|1|3",
            "CI|DI|3|0.5",
            "CI|ST|3|1",
            "CI|PO|3|3",
            "HP|DI|3|0.5",
            "HP|ST|3|1",
            "HP|PO|3|3",
            "OL|DI|3|0.5",
            "OL|ST|3|1",
            "OL|PO|3|3",
            "PFE|DI|3|0.5",
            "PFE|ST|3|1",
            "PFE|PO|3|3",
            "PFL|DI|3|0.5",
            "PFL|ST|3|1",
            "PFL|PO|3|3",
            "VI|DI|3|0.5",
            "VI|ST|3|1",
            "VI|PO|3|3",
            "VIL|DI|3|0.5",
            "VIL|ST|3|1",
            "VIL|PO|3|3",
            "HH|DI|3|0.5",
            "HH|ST|3|1",
            "HH|PO|3|3",
            "HL|DI|0.5|0.5",
            "HL|ST|0.5|1",
            "HL|PO|0.5|3",
            "AA|DI|5|0.5",
            "AA|ST|5|1",
            "AA|PO|5|3"
        }

#Region "    Distances from crop, drift percentages"

    Private Const driftPercentFormat As String = "0.0000"
    Private Const driftPercentUnit As String = " %"

    Private Const bufferFormat As String = "0.0"
    Private Const bufferUnit As String = " m"

#Region "   -   "

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step12Single As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step12SinglePEC As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step03SingleLoading As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step03MultiLoading As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step04SingleLoading As Double = Double.NaN


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step04MultiLoading As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step03SinglePEC As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step04SinglePEC As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step04MultiPEC As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step03MultiPEC As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step12MultiPEC As Double = Double.NaN



    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step12Multi As Double = Double.NaN



    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step04SingleSingle As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step03Single As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _totalDriftSingle As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step04Multi As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step03Multi As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _totalDriftMulti As Double = Double.NaN



#End Region

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _crop2Bank As Double = 0

    ''' <summary>
    ''' Distance  Crop to Bank in m 
    ''' </summary>
    <Category()>
    <DisplayName(
    "Crop <-> Bank")>
    <Description(
    "Distance Crop <-> Bank" & vbCrLf &
    "in m")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '" & bufferFormat &
    "'|unit='" & bufferUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    Public Property Crop2Bank As Double
        Get
            Return _crop2Bank
        End Get
        Set
            _crop2Bank = Value
        End Set
    End Property


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _bank2Water As Double = 0

    ''' <summary>
    ''' Distance  Bank to Water in m
    ''' </summary>
    <Category()>
    <DisplayName(
    "Bank <-> Water")>
    <Description(
    "Distance Bank <-> Water" & vbCrLf &
    "in m")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '" & bufferFormat &
    "'|unit='" & bufferUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    Public Property Bank2Water As Double
        Get
            Return _bank2Water
        End Get
        Set
            _bank2Water = Value
        End Set
    End Property


    ''' <summary>
    ''' Closest to the edge of the field in m
    ''' </summary>
    <Category()>
    <RefreshProperties(RefreshProperties.All)>
    <DisplayName(
    "edge nearest  field")>
    <Description(
    "Closest to the edge of the field in m" & vbCrLf &
    "FOCUS std. Buffer length: Crop2Bank + Bank2Water")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '" & bufferFormat &
    "'|unit='" & bufferUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    Public Property Closest2EdgeOfField As Double


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _nearestDriftPercentSingle As Double = 0

    ''' <summary>
    ''' Drift at edge nearest field in %, single appln.
    ''' </summary>
    ''' <returns></returns>
    <Category()>
    <DisplayName(
    "Drift  , single")>
    <Description(
    "Drift at edge nearest field" & vbCrLf &
    "in %, single appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '" & driftPercentFormat &
    "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    Public Property NearestDriftPercentSingle As Double
        Get
            Return _nearestDriftPercentSingle
        End Get
        Set
            _nearestDriftPercentSingle = Value
        End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _nearestDriftPercentMulti As Double = 0

    ''' <summary>
    ''' Drift at edge nearest field in %, multi applns.
    ''' </summary>
    ''' <returns></returns>
    <Category()>
    <DisplayName(
    "         multi")>
    <Description(
    "Drift at edge nearest field" & vbCrLf &
    "in %, multi applns.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '" & driftPercentFormat &
    "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    Public Property NearestDriftPercentMulti As Double
        Get
            Return _nearestDriftPercentMulti
        End Get
        Set
            _nearestDriftPercentMulti = Value
        End Set
    End Property



    ''' <summary>
    ''' Farthest to the edge of the field in m
    ''' BufferWidth + WaterBodyWidth
    ''' </summary>
    <Category()>
    <RefreshProperties(RefreshProperties.All)>
    <DisplayName(
    "farthest from field")>
    <Description(
    "Farthest to the edge of the field in m" & vbCrLf &
    "Closest2Edge + water body width")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '" & bufferFormat &
    "'|unit='" & bufferUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    Public Property Farthest2EdgeOfWB As Double


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _farthestDriftPercentSingle As Double = 0

    ''' <summary>
    ''' Drift farthest to the 
    ''' edge of the field in %, single appln.
    ''' </summary>
    ''' <returns></returns>
    <Category()>
    <DisplayName(
    "Drift  , single")>
    <Description(
    "Drift farthest to the edge of the field" & vbCrLf &
    "in %, single appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '" & driftPercentFormat &
    "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    Public Property FarthestDriftPercentSingle As Double
        Get
            Return _farthestDriftPercentSingle
        End Get
        Set
            _farthestDriftPercentSingle = Value
        End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _farthestDriftPercentMulti As Double = 0

    ''' <summary>
    ''' Drift farthest to the 
    ''' edge of the field in %, multi applns.
    ''' </summary>
    ''' <returns></returns>
    <Category()>
    <DisplayName(
    "         multi")>
    <Description(
    "Drift farthest to the edge of the field" & vbCrLf &
    "in %, multi applns.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '" & driftPercentFormat &
    "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    Public Property FarthestDriftPercentMulti As Double
        Get
            Return _farthestDriftPercentMulti
        End Get
        Set
            _farthestDriftPercentMulti = Value
        End Set
    End Property


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _averageDriftPercentSingle As Double = 0

    ''' <summary>
    ''' Drift at average distance %, single appln.
    ''' </summary>
    ''' <returns></returns>
    <Category()>
    <DisplayName(
    "Average, single")>
    <Description(
    "Drift at average distance" & vbCrLf &
    "in %, single appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '" & driftPercentFormat &
    "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    Public Property AverageDriftPercentSingle As Double
        Get
            Return _averageDriftPercentSingle
        End Get
        Set
            _averageDriftPercentSingle = Value
        End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _averageDriftPercentMulti As Double = 0

    ''' <summary>
    ''' Drift at average distance %, multi applns.
    ''' </summary>
    ''' <returns></returns>
    <Category()>
    <DisplayName(
    "         multi")>
    <Description(
    "Drift at average distance" & vbCrLf &
    "in %, multi applns.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format= '" & driftPercentFormat &
    "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    Public Property AverageDriftPercentMulti As Double
        Get
            Return _averageDriftPercentMulti
        End Get
        Set
            _averageDriftPercentMulti = Value
        End Set
    End Property


#End Region


End Class

#End Region

