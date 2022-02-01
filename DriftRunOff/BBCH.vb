
Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.IO
Imports System.Web.Script.Serialization
Imports System.Xml.Serialization

Imports DriftRunOff.fo

<Serializable>
<DefaultProperty("FOCUSswDriftCrop")>
<TypeConverter(GetType(PropGridConverter))>
Public Class BBCH

    Public Sub New()

    End Sub

    <DisplayName(
    "SWASH_BBCH.out")>
    <Description(
    "AppDate output for all" & vbCrLf &
    "Crop X Scenario X BBCH combinations")>
    <RefreshProperties(RefreshProperties.All)>
    <Browsable(False)>
    <XmlIgnore> <ScriptIgnore>
    Public Shared Property SWASH_BBCHout As String() = {}


    Public Shared Function GetDateFromBBCH(BBCH As Integer,
                                FOCUSswDriftCrop As eFOCUSswDriftCrop,
                                FOCUSswScenario As eFOCUSswScenario,
                       Optional Season As eSeason = eSeason.first) As Date

        Dim Target As String()

        Target =
            Filter(
                Source:=SWASH_BBCHout,
                Match:=GetBBCHSearchStringsFromCrop(
                                FOCUSswDriftCrop:=FOCUSswDriftCrop,
                                Season:=Season))

        Target =
            Filter(Source:=Target, Match:=FOCUSswScenario.ToString)

        Target =
            Filter(Source:=Target, Match:=vbTab & BBCH & vbTab)

        Target = Target.First.Split(vbTab)(6).Split(".")

        Return New Date(year:=Target(2), month:=Target(1), day:=Target(0))



    End Function

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const catInput As String = " 01 Main"

#Region "    Main"

#End Region

#Region "    FOCUS Crop, Season & Scenario"

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

            _FOCUSswDriftCrop = Value

            If enumConverter(Of Type).getEnumDescription(
                    EnumConstant:=_FOCUSswDriftCrop).Contains("2nd") Then
                _Season = eSeason.first
            End If

            'only scenarios defined for this crop
            enumConverter(Of eFOCUSswScenario).onlyShow =
                GetScenariosFromCrop(Crop:=_FOCUSswDriftCrop)

            _FOCUSswScenario = eFOCUSswScenario.not_defined

        End Set
    End Property

    Private _Season As eSeason = eSeason.not_defined

    ''' <summary>
    ''' Crop season, 1st or 2nd
    ''' </summary>
    ''' <returns></returns>
    <Category(catInput)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property Season As eSeason
        Get
            Return _Season
        End Get
        Set

            If Value = eSeason.second AndAlso
                    enumConverter(Of Type).getEnumDescription(
                    EnumConstant:=_FOCUSswDriftCrop).Contains("2nd") Then

                _Season = Value
            Else

                _Season = eSeason.first
            End If

        End Set
    End Property

    ''' <summary>
    ''' Corp for BBCH.out
    ''' </summary>
    ''' <returns></returns>
    <Category(catInput)>
    <RefreshProperties(RefreshProperties.All)>
    Public ReadOnly Property BBCHCrop As String
        Get
            Return GetBBCHSearchStringsFromCrop(
                FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                Season:=_Season)
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

                End If

            End If

        End Set
    End Property

#End Region


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const catPAT As String = " 02 PAT window"

#Region "    PAT Window"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _noOfApplns As eNoOfApplns = eNoOfApplns.not_defined

    ''' <summary>
    ''' NoOfApplns
    ''' Number of applications
    ''' 1 - 8+
    ''' </summary>
    <Category(catPAT)>
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
                _interval = Integer.MaxValue
            End If

            _Window = 0

            CheckWindow(
                NoOfApplns:=_noOfApplns,
                Interval:=Interval,
                Window:=_Window)

        End Set
    End Property



#Region "    Interval : Time between applns. in days"

    ''' <summary>
    ''' GUI : Time between applns. in days
    ''' </summary>
    ''' <returns></returns>
    <Category(catPAT)>
    <DisplayName(
    "Appln. Interval")>
    <Description(
    "Time between applns. in days" & vbCrLf &
    "")>
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
                    Interval = Integer.Parse(Trim(Value))
                Catch ex As Exception
                    Interval = 0
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
    <Description(
    "Time between applns. in days")>
    <DisplayName(
    "Interval")>
    <Category(catPAT)>
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
            _Window = 0

            CheckWindow(
                NoOfApplns:=_noOfApplns,
                Interval:=Interval,
                Window:=_Window)

        End Set
    End Property

#End Region

    Private _PHI As Integer = 0

    <Category(catPAT)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(IntConv))>
    Public Property PHI As Integer
        Get
            Return _PHI
        End Get
        Set(value As Integer)
            _PHI = value
        End Set
    End Property

    Private _Window As Integer = 0

    ''' <summary>
    ''' PAT window
    ''' mmin : 30 + (# of applns -1 ) * appln interval
    ''' </summary>
    ''' <returns></returns>
    <Category(catPAT)>
    <RefreshProperties(RefreshProperties.All)>
    <DefaultValue(30)>
    <TypeConverter(GetType(IntConv))>
    Public Property Window As Integer
        Get
            Return _Window
        End Get
        Set(value As Integer)

            CheckWindow(
                NoOfApplns:=_noOfApplns,
                Interval:=Interval,
                Window:=value)

            _Window = value

        End Set
    End Property

    Public Shared Function CheckWindow(
                    NoOfApplns As eNoOfApplns,
                    Interval As Integer,
     Optional ByRef Window As Integer = -1) As Integer

        If NoOfApplns = eNoOfApplns._01 Then
            Window = 30
        ElseIf NoOfApplns = eNoOfApplns.not_defined OrElse
               Interval = Integer.MaxValue Then
            Window = 0
        Else
            If Window < 30 + NoOfApplns * Interval Then
                Window = 30 + NoOfApplns * Interval
            End If
        End If

        Return Window

    End Function

#End Region


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const catBBCH As String = " 03 BBCH min max"

    Private _BBCHmin As eBBCH = eBBCH.not_defined

    <Category(catBBCH)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property BBCHmin As eBBCH
        Get
            Return _BBCHmin
        End Get
        Set

            If _FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined OrElse
               _FOCUSswScenario = eFOCUSswScenario.not_defined Then
                _BBCHmin = eBBCH.not_defined
                _BBCHmax = eBBCH.not_defined
                Exit Property
            End If


            _BBCHmin = Value



            Date1st = GetDateFromBBCH(
                            BBCH:=_BBCHmin,
                            FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                            FOCUSswScenario:=_FOCUSswScenario,
                            Season:=
                            IIf(
                                Expression:=_Season = eSeason.not_defined,
                                TruePart:=eSeason.first,
                                FalsePart:=_Season))

            If BBCHmax < _BBCHmin Then
                BBCHmax = _BBCHmin
            End If

        End Set
    End Property

    <Category(catBBCH)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DateConv))>
    <AttributeProvider("format= 'dd-MMM-yy'")>
    Public Property Date1st As New Date


    Private _BBCHmax As eBBCH = eBBCH.not_defined

    <Category(catBBCH)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property BBCHmax As eBBCH
        Get
            Return _BBCHmax
        End Get
        Set

            If _FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined OrElse
               _FOCUSswScenario = eFOCUSswScenario.not_defined Then
                Exit Property
            End If

            If Value < _BBCHmin Then
                Value = _BBCHmin
            End If

            _BBCHmax = Value

            Datelast = GetDateFromBBCH(
                            BBCH:=_BBCHmax,
                            FOCUSswDriftCrop:=_FOCUSswDriftCrop,
                            FOCUSswScenario:=_FOCUSswScenario,
                            Season:=
                            IIf(
                                Expression:=_Season = eSeason.not_defined,
                                TruePart:=eSeason.first,
                                FalsePart:=_Season))

            'Dim tempDate As Date
            'tempDate.AddDays(-PHI)

            'If tempDate.AddDays(-Window) >= Date1st Then
            '    Datelast = tempDate
            'Else
            '    Datelast = New Date
            'End If

        End Set
    End Property

    <Category(catBBCH)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DateConv))>
    <AttributeProvider("format= 'dd-MMM-yy'")>
    Public Property Datelast As New Date

    <Category(catBBCH)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DateConv))>
    <AttributeProvider("format= 'dd-MMM-yy'")>
    Public ReadOnly Property DateFinal As Date
        Get

            If Range < Window + PHI Then
                Return New Date
            Else
                Return Datelast.AddDays(-PHI)
            End If

        End Get
    End Property


    <Category(catBBCH)>
    <TypeConverter(GetType(IntConv))>
    Public ReadOnly Property Range As Integer
        Get

            If Datelast > Date1st Then

                Return (Datelast - Date1st).Days
            Else
                Return 0

            End If

        End Get
    End Property


    Public Property BBCHearliest As eBBCH = eBBCH.not_defined

    Public Property DateEarliest As New Date

    Public Property BBCHlatest As eBBCH = eBBCH.not_defined



    Public Property DateLatest As New Date




End Class


Public Class ApplnWindow

    Public Sub New()

    End Sub

    Public Property BBCH_Start As eBBCH = eBBCH._0

    Public Property Start_Date As Date

End Class