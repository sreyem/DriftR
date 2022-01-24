Imports System.ComponentModel
Imports System.IO
Imports System.Web.Script.Serialization
Imports System.Xml.Serialization
Imports core


Module Main

    Sub main()


restart:


        Console.Clear()

        mylog(LogTxtArray:=getStartInfo)

        'mylog(
        '    LogTxt:="ztsParser v " &
        '    getExecVersionInfo(
        '    filename:=Path.Combine(Environment.CurrentDirectory, "ztsParser.exe")).ToString)

        DriftR = New DriftR


        showform = New frmPropGrid(
            class2Show:=DriftR,
             classType:=DriftR.GetType,
               restart:=True)


        Try

            If Not IsNothing(showform) Then

                With showform

                    .Width = 1200
                    .Height = 1300

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



End Module

<TypeConverter(GetType(PropGridConverter))>
Public Class DriftR

    Public Sub New()

    End Sub


    Public Const catFOCUS As String = " 01 FOCUS "

#Region "    FOCUS"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _FOCUSswDriftCrop As eFOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined

    ''' <summary>
    ''' Target crop out of the
    ''' available FOCUSsw crops
    ''' </summary>
    <Category(catFOCUS)>
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

            'only scenarios defined for this crop
            enumConverter(Of eFOCUSswScenario).onlyShow =
                GetScenariosFromCrop(Crop:=_FOCUSswDriftCrop)

            'reset if crop changed
            _FOCUSswScenario = eFOCUSswScenario.not_defined
            _FOCUSswWaterBody = eFOCUSswWaterBody.not_defined
            _wbDepths.Data = {}

        End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _FOCUSswScenario As eFOCUSswScenario = eFOCUSswScenario.not_defined

    ''' <summary>
    ''' FOCUSsw Scenario
    ''' D1 - D6 And R1 - R4
    ''' </summary>
    <Category(catFOCUS)>
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
                    End If

                End If

            End If

        End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _FOCUSswWaterBody As eFOCUSswWaterBody = eFOCUSswWaterBody.not_defined

    ''' <summary>
    ''' FOCUS water body
    ''' Ditch, pond or stream
    ''' </summary>
    <Category(catFOCUS)>
    <DisplayName(
    "Water Body")>
    <Description(
    "Ditch, stream : 100m x 1m x 0.3m ;  30,000L" & vbCrLf &
    "Pond          : 30m radius x 1m  ; ~94,000L")>
    <RefreshProperties(RefreshProperties.All)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(CInt(eFOCUSswWaterBody.not_defined))>
    Public Property FOCUSswWaterBody As eFOCUSswWaterBody
        Get
            Return _FOCUSswWaterBody
        End Get
        Set

            If _FOCUSswScenario = eFOCUSswDriftCrop.not_defined Then
                _FOCUSswWaterBody = eFOCUSswScenario.not_defined
            Else

                If enumConverter(Of eFOCUSswWaterBody).onlyShow.Contains(
                   enumConverter(Of eFOCUSswWaterBody).getEnumDescription(EnumConstant:=Value)) Then

                    _FOCUSswWaterBody = Value

                    Me.Distances.GetFOCUSStdDistance(
                        FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                        FOCUSswWaterBody:=_FOCUSswWaterBody,
                        ApplnMethodStep03:=_applnMethodStep03)

                    If _FOCUSswWaterBody = eFOCUSswWaterBody.PO Then
                        Me.WBDepths.Data = {1}
                        Me.waterDepth = 1
                    ElseIf _FOCUSswScenario.ToString.ToUpper.StartsWith("R") Then
                        Me.WBDepths.Data =
                                {RunOffWaterDepthsEstimates(_FOCUSswScenario.ToString.Substring(1, 1) - 1)}
                    End If

                    Me.DepthValueMode = _DepthValueMode

                End If

            End If

        End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _waterDepth As Double = 0.3

    ''' <summary>
    ''' Depth in m
    ''' </summary>
    ''' <returns></returns>
    <Category(catFOCUS)>
    <DisplayName(
    "      Depth")>
    <Description(
    "in m" & vbCrLf &
    "")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    Public Property waterDepth As Double
        Get
            Return _waterDepth
        End Get
        Set
            _waterDepth = Value
        End Set
    End Property


    Public RunOffWaterDepthsEstimates As Double() = {0.41, 0.305, 0.29, 0.41}

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _wbDepths As New StatsPEC


#If DEBUG Then
    Private Const statsVisible As Boolean = True
#Else
    Private Const statsVisible As Boolean = false
#End If

    <Category(catCalculations)>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(statsVisible)>
    <[ReadOnly](False)>
    Public Property WBDepths As StatsPEC
        Get
            Return _wbDepths
        End Get
        Set
            _wbDepths = Value
        End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _DepthValueMode As eDepthValueMode = eDepthValueMode.last


    <RefreshProperties(RefreshProperties.All)>
    <Category(catCalculations)>
    Public Property DepthValueMode As eDepthValueMode
        Get
            Return _DepthValueMode
        End Get
        Set

            Dim year As Integer = -1

            With _wbDepths

                If _FOCUSswWaterBody = eFOCUSswWaterBody.PO Then
                    .Data = {1}
                Else
                    If .Data.Count > 0 Then

                        Select Case Value

                            Case eDepthValueMode.last
                                _waterDepth = .Data.Last

                            Case eDepthValueMode.max
                                _waterDepth = .Max

                            Case eDepthValueMode.min
                                _waterDepth = .Min

                            Case Else

                                If Value.ToString.ToUpper.StartsWith("Y") Then
                                    year =
                                    CInt(
                                    Replace(
                                        Expression:=Value.ToString.ToUpper,
                                        Find:="Y",
                                        Replacement:="",
                                        Compare:=CompareMethod.Text))

                                    If _wbDepths.Data.Count >= year Then
                                        _waterDepth = _wbDepths.Data(year - 1)
                                    Else
                                        'ToDo
                                    End If

                                End If

                        End Select

                    End If
                End If

            End With

        End Set
    End Property

#Region "    Buffer"

    ''' <summary>
    ''' GUI : Time between applns. in days
    ''' </summary>
    ''' <returns></returns>
    <Category(catFOCUS)>
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
                    "Step03, no add. buffer" &
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
                Return _buffer.ToString()
            Else
                Return dropDownList.dropDownEntries.First
            End If

        End Get
        Set

            Select Case Value

                Case dropDownList.dropDownEntries(0)
                    _buffer = 0

                Case dropDownList.dropDownEntries(1)
                    _buffer = 1

                Case dropDownList.dropDownEntries(2)
                    _buffer = 2

                Case dropDownList.dropDownEntries(3)
                    _buffer = 3

                Case dropDownList.dropDownEntries(4)
                    _buffer = 4

                Case Else

                    Try

                        _buffer =
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
                        _buffer = 0
                    End Try

            End Select

        End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _buffer As Integer = 0

    ''' <summary>
    ''' Buffer in m
    ''' </summary>
    ''' <returns></returns>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Buffer")>
    <DisplayName("Buffer")>
    <Category(catUse)>
    <DefaultValue(0)>
    <Browsable(False)>
    <TypeConverter(GetType(IntConv))>
    <AttributeProvider(
    "unit=' m'")>
    Public Property Buffer As Integer
        Get
            Return _buffer
        End Get
        Set
            _buffer = Value
        End Set
    End Property

#End Region

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _nozzle As eNozzles

    ''' <summary>
    ''' Drift reducing nozzle in %
    ''' </summary>
    ''' <returns></returns>
    <Category(catFOCUS)>
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
        End Set
    End Property

#End Region

    Public Const catUse As String = " 02 Use "

#Region "    Use"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _applnMethodStep03 As eApplnMethodStep03 = eApplnMethodStep03.not_defined

    ''' <summary>
    ''' Application Methods
    ''' </summary>
    ''' <returns></returns>
    <Category(catUse)>
    <DisplayName(
    "Method")>
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

            Me.Distances.GetFOCUSStdDistance(
                FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                FOCUSswWaterBody:=_FOCUSswWaterBody,
                ApplnMethodStep03:=_applnMethodStep03)

        End Set
    End Property


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _noOfApplns As eNoOfApplns = eNoOfApplns.not_defined

    ''' <summary>
    ''' Number of applications
    ''' 1 - 8 or more
    ''' </summary>
    <Category(catUse)>
    <DisplayName(
    "Number")>
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
                    Ganzelmeier:=Ganzelmeier,
                    NoOfApplns:=_noOfApplns)

        End Set
    End Property


#Region "    Interval : Time between applns. in days"

    ''' <summary>
    ''' GUI : Time between applns. in days
    ''' </summary>
    ''' <returns></returns>
    <Category(catUse)>
    <DisplayName(
        "Interval")>
    <Description(
        "Time between applns. in days")>
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
    Private _interval As Integer = 0

    ''' <summary>
    ''' Time between applns. in days
    ''' </summary>
    ''' <returns></returns>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Time between applns. in days")>
    <DisplayName("Interval")>
    <Category(catUse)>
    <DefaultValue(0)>
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

#End Region

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _rate As Double = 0

    ''' <summary>
    ''' Appln Rate
    ''' </summary>
    ''' <returns></returns>
    <Category(catUse)>
    <DisplayName(
    "Rate")>
    <Description(
    "Application rate" & vbCrLf &
    "in kg as/ha")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
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

    Public Const catCalculations As String = " 03 Calculations "

#Region "    Calculations"

#Region "    functions"

    ''' <summary>
    ''' Distances
    ''' nearest, farthest, average
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eDriftDistance)))>
    Public Enum eDriftDistance

        <Description("Drift at edge nearest field")>
        closest

        <Description("Drift farthest to the edge of the field")>
        farthest

        <Description("Areic mean drift")>
        average

    End Enum

    ''' <summary>
    ''' calculation of FOCUS drift values
    ''' </summary>   
    ''' <remarks></remarks>
    <DebuggerStepThrough>
    Public Sub CalcDriftPercent()

        'check inputs
        If NoOfApplns = eNoOfApplns.not_defined OrElse
           FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined OrElse
           FOCUSswWaterBody = eFOCUSswWaterBody.not_defined Then

            Exit Sub

        End If

        If Nozzle = CInt(eNozzles.not_defined) Then
            Nozzle = eNozzles._0
        End If

        With Regression.SingleApplnRegression

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
                    Distances.NearestDriftPercentSingle = (.A / (.B + 1) * (Distances.Farthest2EdgeOfWB ^ (.B + 1) - Distances.Closest2EdgeOfField ^ (.B + 1))) /
                                           (Distances.Farthest2EdgeOfWB - Distances.Closest2EdgeOfField)
                ElseIf Distances.Closest2EdgeOfField > .HingePoint Then
                    Distances.NearestDriftPercentSingle = .C / (.D + 1) * (Distances.Farthest2EdgeOfWB ^ (.D + 1) - Distances.Closest2EdgeOfField ^ (.D + 1)) /
                                          (Distances.Farthest2EdgeOfWB - Distances.Closest2EdgeOfField)
                Else
                    Distances.NearestDriftPercentSingle = (.A / (.B + 1) * (.HingePoint ^ (.B + 1) - Distances.Closest2EdgeOfField ^ (.B + 1)) + .C / (.D + 1) * (Distances.Farthest2EdgeOfWB ^ (.D + 1) - .HingePoint ^ (.D + 1))) * 1 /
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
            Distances.NearestDriftPercentSingle *= 1 - _nozzle / 100

        End With


        With Regression.MultiApplnRegression

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
                    Distances.NearestDriftPercentSingle = (.A / (.B + 1) * (Distances.Farthest2EdgeOfWB ^ (.B + 1) - Distances.Closest2EdgeOfField ^ (.B + 1))) /
                                           (Distances.Farthest2EdgeOfWB - Distances.Closest2EdgeOfField)
                ElseIf Distances.Closest2EdgeOfField > .HingePoint Then
                    Distances.NearestDriftPercentSingle = .C / (.D + 1) * (Distances.Farthest2EdgeOfWB ^ (.D + 1) - Distances.Closest2EdgeOfField ^ (.D + 1)) /
                                          (Distances.Farthest2EdgeOfWB - Distances.Closest2EdgeOfField)
                Else
                    Distances.NearestDriftPercentSingle = (.A / (.B + 1) * (.HingePoint ^ (.B + 1) - Distances.Closest2EdgeOfField ^ (.B + 1)) + .C / (.D + 1) * (Distances.Farthest2EdgeOfWB ^ (.D + 1) - .HingePoint ^ (.D + 1))) * 1 /
                                           (Distances.Farthest2EdgeOfWB - Distances.Closest2EdgeOfField)
                End If

            Catch ex As Exception

                Throw New _
                        ArithmeticException(
                        message:="Error during drift calc.",
                        innerException:=ex)

            End Try

            ' if stream then apply upstream catchment factor 1.2
            Distances.NearestDriftPercentMulti *=
                IIf(
                    Expression:=_FOCUSswWaterBody = eFOCUSswWaterBody.ST,
                    TruePart:=1.2,
                    FalsePart:=1)

            'drift reducing nozzles?     
            Distances.NearestDriftPercentMulti *= 1 - _nozzle / 100

        End With

    End Sub


#End Region


    ''' <summary>
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
    Public ReadOnly Property Ganzelmeier As eGanzelmeier
        Get

            Dim returnValue As eGanzelmeier = eGanzelmeier.not_defined

            If _FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined Then
                returnValue = eGanzelmeier.not_defined
            End If

            Select Case _applnMethodStep03

                Case eApplnMethodStep03.SI, eApplnMethodStep03.GR
                    returnValue = eGanzelmeier.noDrift

                Case eApplnMethodStep03.GS
                    returnValue = eGanzelmeier.ArableCrops

                Case eApplnMethodStep03.AA
                    returnValue = eGanzelmeier.AerialAppln

                Case eApplnMethodStep03.AB

                    Select Case _FOCUSswDriftCrop

                        Case _
                            eFOCUSswDriftCrop.HP,
                            eFOCUSswDriftCrop.CI,
                            eFOCUSswDriftCrop.OL

                            returnValue = eGanzelmeier.FruitCrops_Late


                        Case eFOCUSswDriftCrop.PFE
                            returnValue = eGanzelmeier.FruitCrops_Early

                        Case eFOCUSswDriftCrop.PFL
                            returnValue = eGanzelmeier.FruitCrops_Late

                        Case eFOCUSswDriftCrop.VI
                            returnValue = eGanzelmeier.Vines_Early

                        Case eFOCUSswDriftCrop.VIL
                            returnValue = eGanzelmeier.Vines_Late

                    End Select

                Case eApplnMethodStep03.HL
                    returnValue = eGanzelmeier.ArableCrops

                Case eApplnMethodStep03.HH
                    returnValue = eGanzelmeier.FruitCrops_Late

                Case eApplnMethodStep03.not_defined
                    returnValue = eGanzelmeier.not_defined

            End Select

            If returnValue <> eGanzelmeier.not_defined AndAlso _noOfApplns <> eNoOfApplns.not_defined Then
                Me.Regression.Update(
                        Ganzelmeier:=returnValue,
                        NoOfApplns:=_noOfApplns)
            End If


            Return returnValue

        End Get
    End Property


    <XmlIgnore> <ScriptIgnore>
    <Category(catCalculations)>
    <DisplayName(
    "Regression Parameters")>
    <Description(
    "" & vbCrLf &
    "")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](True)>
    Public Property Regression As New Regression


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
    Public Property Distances As New Distances

#End Region

End Class

<TypeConverter(GetType(PropGridConverter))>
Public Class Regression


    Public Sub New()

    End Sub

    Public Sub New(Ganzelmeier As eGanzelmeier, NoOfApplns As eNoOfApplns)
        Update(Ganzelmeier:=Ganzelmeier, NoOfApplns:=NoOfApplns)
    End Sub

    <DisplayName("Single appln.")>
    Public Property SingleApplnRegression As New RegressionBase

    <DisplayName("Multi applns.")>
    Public Property MultiApplnRegression As New RegressionBase

    Public Sub Update(Ganzelmeier As eGanzelmeier, NoOfApplns As eNoOfApplns)

        If Ganzelmeier = eGanzelmeier.not_defined OrElse
           NoOfApplns = eNoOfApplns.not_defined Then

            With Me.SingleApplnRegression

                .A = 0
                .B = 0
                .C = 0
                .D = 0
                .HingePoint = 0

            End With

            With Me.MultiApplnRegression

                .A = 0
                .B = 0
                .C = 0
                .D = 0
                .HingePoint = 0

            End With

        Else

            With Me.SingleApplnRegression

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

                With Me.MultiApplnRegression

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

                With Me.MultiApplnRegression

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
    Public Property A As Double = 0

    <Category()>
    <Description(
        "" & vbCrLf &
        "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    Public Property B As Double = 0

    <Category()>
    <Description(
        "" & vbCrLf &
        "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
    Public Property C As Double = 0

    <Category()>
    <Description(
        "" & vbCrLf &
        "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(0)>
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

    <DebuggerStepThrough()>
    Public Sub GetFOCUSStdDistance(
                                FOCUSswDriftCrop As eFOCUSswDriftCrop,
                                FOCUSswWaterBody As eFOCUSswWaterBody,
                                ApplnMethodStep03 As eApplnMethodStep03)

        Dim SearchString As String
        Dim FOCUSstdDistancesEntry As String = ""

        If FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined OrElse
           FOCUSswWaterBody = eFOCUSswWaterBody.not_defined Then

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

            With Me

                .Crop2Bank =
                    FOCUSstdDistancesEntry.Split({"|"c})(eFOCUSStdDistancesMember.edgeField2topBank)

                .Bank2Water =
                    FOCUSstdDistancesEntry.Split({"|"c})(eFOCUSStdDistancesMember.topBank2edgeWaterbody)

                .Closest2EdgeOfField = _crop2Bank + _bank2Water

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

    Private Const driftPercentFormat As String = "0.000"
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
    "Drift ,single")>
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
    "       multi")>
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
    <DefaultValue(Double.NaN)>
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
    "Drift ,single")>
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
    Private _farthestDriftPercentMulti As Double = Double.NaN

    ''' <summary>
    ''' Drift farthest to the 
    ''' edge of the field in %, multi applns.
    ''' </summary>
    ''' <returns></returns>
    <Category()>
    <DisplayName(
    "       multi")>
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

#End Region


End Class


Public Class Results


    Public Sub New()

    End Sub

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _FOCUSswScenario As eFOCUSswScenario = eFOCUSswScenario.not_defined

    ''' <summary>
    ''' FOCUSsw Scenario
    ''' D1 - D6 And R1 - R4
    ''' </summary>
    <Category()>
    <DisplayName(
    "Scenario")>
    <Description(
    "FOCUSsw scenario" & vbCrLf &
    "D1 - D6, R1 - R4")>
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
            _FOCUSswScenario = Value
        End Set
    End Property


End Class












<TypeConverter(GetType(PropGridConverter))>
Public Class Inputs

    Public Sub New()

    End Sub


    Public Const catInputs As String = " 01 Inputs "

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _FOCUSswDriftCrop As eFOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined

    ''' <summary>
    ''' Target crop out of the
    ''' available FOCUS crops
    ''' </summary>
    <Category(catInputs)>
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

        End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _applnMethodStep03 As eApplnMethodStep03 = eApplnMethodStep03.not_defined

    ''' <summary>
    ''' Methods
    ''' </summary>
    ''' <returns></returns>
    <Category(catInputs)>
    <DisplayName("Method")>
    <RefreshProperties(RefreshProperties.All)>
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

        End Set
    End Property


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _noOfApplns As eNoOfApplns = eNoOfApplns.not_defined

    ''' <summary>
    ''' Number of applications
    ''' 1 - 8 or more
    ''' </summary>
    <Category(catInputs)>
    <DisplayName(
        "Number")>
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
                interval = Integer.MaxValue
            End If



        End Set
    End Property



#Region "    Interval : Time between applns. in days"

    ''' <summary>
    ''' GUI : Time between applns. in days
    ''' </summary>
    ''' <returns></returns>
    <Category(catInputs)>
    <DisplayName(
        "Interval")>
    <Description(
        "Time between applns. in days")>
    <TypeConverter(GetType(dropDownList))>
    <RefreshProperties(RefreshProperties.All)>
    <DefaultValue(" - ")>
    <Browsable(False)>
    <XmlIgnore> <ScriptIgnore>
    Public Property intervalGUI As String
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
                    "42",
                    "50"
                    }

            If interval <> Integer.MaxValue Then
                Return interval.ToString()
            Else
                Return " - "
            End If

        End Get
        Set

            If Value <> " - " Then

                Try
                    interval = Integer.Parse(Trim(Value))
                Catch ex As Exception
                    interval = Integer.MaxValue
                End Try

            Else

                interval = Integer.MaxValue

            End If

        End Set
    End Property


    ''' <summary>
    ''' Time between applns. in days
    ''' </summary>
    ''' <returns></returns>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Time between applns. in days")>
    <DisplayName("Interval")>
    <Category(catInputs)>
    <DefaultValue(Integer.MaxValue)>
    <Browsable(False)>
    <TypeConverter(GetType(IntConv))>
    Public Property interval As Integer


#End Region

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _rate As Double = Double.NaN

    ''' <summary>
    ''' Appln Rate
    ''' </summary>
    ''' <returns></returns>
    <Category(catInputs)>
    <DisplayName(
        "Rate")>
    <Description(
        "Application rate" & vbCrLf &
        "in kg as/ha")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
        "format= 'G4'|unit=' kg as/ha'")>
    Public Property rate As Double
        Get
            Return _rate
        End Get
        Set

            _rate = Value

        End Set
    End Property

    Private _scenario As eFOCUSswScenario = eFOCUSswScenario.not_defined

    ''' <summary>
    ''' PRZMsw Scenario
    ''' R1 - R4, not def.
    ''' </summary>
    <Category(catInputs)>
    <DisplayName(
    "Scenario")>
    <Description(
    "PRZM Scenario, R1 - R4" & vbCrLf &
    "")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(CInt(eFOCUSswScenario.not_defined))>
    Public Property scenario As eFOCUSswScenario
        Get
            Return _scenario
        End Get
        Set
            _scenario = Value

            'If _scenario <> eFOCUSswScenario.not_defined Then
            '    waterDepth = runOffWaterDepths(_scenario)
            'Else
            '    waterDepth = 0.3
            'End If

        End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _FOCUSswWaterBody As eFOCUSswWaterBody = eFOCUSswWaterBody.not_defined

    ''' <summary>
    ''' FOCUS water body
    ''' Ditch, pond or stream
    ''' </summary>
    <Category(catInputs)>
    <DisplayName(
    "Water Body")>
    <Description(
    "Ditch, pond or stream" & vbCrLf &
    "")>
    <RefreshProperties(RefreshProperties.All)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(CInt(eFOCUSswWaterBody.not_defined))>
    Public Property FOCUSswWaterBody As eFOCUSswWaterBody
        Get
            Return _FOCUSswWaterBody
        End Get
        Set

            _FOCUSswWaterBody = Value

        End Set
    End Property

    ''' <summary>
    ''' Ganzelmeier crop group
    ''' based on selected FOCUS crop
    ''' </summary>
    <Category(catInputs)>
    <DisplayName(
            "Ganzelmeier Crop")>
    <Description(
            "Ganzelmeier crop group" & vbCrLf &
            "base for further calculations")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](True)>
    <DefaultValue(CInt(eGanzelmeier.not_defined))>
    Public ReadOnly Property Ganzelmeier As eGanzelmeier
        Get

            If _FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined OrElse
               _applnMethodStep03 = eApplnMethodStep03.not_defined Then
                Return eGanzelmeier.not_defined
            End If

            If _applnMethodStep03 = eApplnMethodStep03.AA Then
                Return eGanzelmeier.AerialAppln
            End If

            If _applnMethodStep03 = eApplnMethodStep03.SI OrElse
               _applnMethodStep03 = eApplnMethodStep03.GR Then
                Return eGanzelmeier.noDrift
            End If

            If _applnMethodStep03 = eApplnMethodStep03.GS AndAlso
               (_FOCUSswDriftCrop = eFOCUSswDriftCrop.PFE OrElse
                _FOCUSswDriftCrop = eFOCUSswDriftCrop.PFL OrElse
                _FOCUSswDriftCrop = eFOCUSswDriftCrop.VI OrElse
                _FOCUSswDriftCrop = eFOCUSswDriftCrop.VIL) Then

                Return eGanzelmeier.ArableCrops

            End If

            Select Case _FOCUSswDriftCrop

                Case _
                    eFOCUSswDriftCrop.HP,
                    eFOCUSswDriftCrop.CI,
                    eFOCUSswDriftCrop.OL

                    If _applnMethodStep03 = eApplnMethodStep03.GS Then
                        Return eGanzelmeier.ArableCrops
                    Else
                        Return eGanzelmeier.FruitCrops_Late
                    End If

                Case eFOCUSswDriftCrop.PFE
                    Return eGanzelmeier.FruitCrops_Early

                Case eFOCUSswDriftCrop.PFL
                    Return eGanzelmeier.FruitCrops_Late

                Case eFOCUSswDriftCrop.VI
                    Return eGanzelmeier.Vines_Early

                Case eFOCUSswDriftCrop.VIL
                    Return eGanzelmeier.Vines_Late

            End Select

            Return eGanzelmeier.ArableCrops

        End Get
    End Property

    ''' <summary>
    ''' Ganzelmeier crop groups
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eGanzelmeier)))>
    Public Enum eGanzelmeier

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        ''' <summary>
        ''' Arable crops 1.9274 %
        ''' </summary>
        <Description(
                "Arable crops " & vbCrLf &
                "1.9274 %")>
        ArableCrops = 0

        ''' <summary>
        ''' Fruit crops, early
        ''' BBCH 01 - 71, 96 - 99, 23.599 %
        ''' </summary>
        <Description(
                "Fruit crops, early " & vbCrLf &
                "23.599 %")>
        FruitCrops_Early

        ''' <summary>
        ''' Fruit crops, late
        ''' BBCH 72 - 95, 11.134 %
        ''' </summary>
        <Description(
                "Fruit crops, late " & vbCrLf &
                "11.134 %")>
        FruitCrops_Late

        ''' <summary>
        ''' Hops 14.554 %
        ''' </summary>
        <Description(
                "Hops " & vbCrLf &
                "14.554 %")>
        Hops

        ''' <summary>
        ''' Vines, early 1.7184 %
        ''' </summary>
        <Description(
                "Vines (early, std.) " & vbCrLf &
                "1.7184 %")>
        Vines_Early

        ''' <summary>
        ''' Vines, late 5.173 %
        ''' Just for compatibility reasons
        ''' </summary>
        <Description(
                "Vines, late " & vbCrLf &
                "5.173 %")>
        Vines_Late

        ''' <summary>
        ''' Aerial appln 25.476 %
        ''' </summary>
        <Description(
                "Aerial appln " & vbCrLf &
                "25.476 %")>
        AerialAppln

        ''' <summary>
        ''' no drift 0 % ;-)
        ''' </summary>
        <Description(
                "no drift " & vbCrLf &
                "0 % ;-)")>
        noDrift

    End Enum


    Public Const catRegressionParameters As String = "03  Regression Parameter"

#Region "    Regression Parameter"

    <Category(catRegressionParameters)>
    <Description(
        "" & vbCrLf &
        "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property A As Double = Double.NaN

    <Category(catRegressionParameters)>
    <Description(
        "" & vbCrLf &
        "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property B As Double = Double.NaN

    <Category(catRegressionParameters)>
    <Description(
        "" & vbCrLf &
        "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property C As Double = Double.NaN

    <Category(catRegressionParameters)>
    <Description(
        "" & vbCrLf &
        "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property D As Double = Double.NaN

#End Region




End Class


