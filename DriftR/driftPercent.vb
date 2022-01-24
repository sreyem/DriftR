


Imports System.ComponentModel
Imports System.Web.Script.Serialization
Imports System.Xml.Serialization
Imports System.Text

Imports core
Imports core.FOCUSdriftDB
Imports System.Drawing.Design

<TypeConverter(GetType(propGridConverter))>
Public Class appln

    Public Sub New()

    End Sub


#Region "   - "

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _rate As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _applnMethodStep03 As eApplnMethodStep03 = eApplnMethodStep03.not_defined

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _BBCH As Integer = 0

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _interval As Integer = Integer.MaxValue

#End Region


    ''' <summary>
    ''' Application rate in kg as/ha
    ''' </summary>
    <Category()>
    <DisplayName(
        "Rate")>
    <Description(
        "Application rate in kg as/ha" & vbCrLf &
        "")>
    <RefreshProperties(RefreshProperties.All)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= 'G4'|" &
        "unit=' kg as/ha'")>
    Public Property rate As Double
        Get
            Return _rate
        End Get
        Set
            _rate = Value
        End Set
    End Property

    ''' <summary>
    ''' Time between applns. in days
    ''' </summary>
    ''' <returns></returns>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Time between applns. in days")>
    <DisplayName("Interval")>
    <Category()>
    <DefaultValue(Integer.MaxValue)>
    <Browsable(False)>
    <TypeConverter(GetType(intConv))>
    Public Property interval As Integer
        Get
            Return _interval
        End Get
        Set
            _interval = Value
        End Set
    End Property

    <TypeConverter(GetType(dropDownList))>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Time between applns. in days")>
    <DisplayName("Interval")>
    <Category()>
    <DefaultValue(" - ")>
    <XmlIgnore> <ScriptIgnore>
    Public Property intervalGUI As String
        Get

            dropDownList.dropDownEntries =
                    {
                    " - ",
                    "5",
                    "7",
                    "10",
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

    <Category()>
    <DisplayName("Method")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(CInt(eApplnMethodStep03.not_defined))>
    Public Property applnMethodStep03 As eApplnMethodStep03
        Get
            Return _applnMethodStep03
        End Get
        Set
            _applnMethodStep03 = Value
        End Set
    End Property

    Public Property cam As eCAM = eCAM.not_defined

    Public Property depth As Double = 0

    ''' <summary>
    ''' BBCH
    ''' </summary>
    ''' <returns></returns>
    <Category()>
    <DisplayName(
        "BBCH")>
    <Description(
        "" & vbCrLf &
        "")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    Public Property BBCH As Integer
        Get
            Return _BBCH
        End Get
        Set
            _BBCH = Value
        End Set
    End Property

    Public Property interception As Double

    Public Property applnDate As New Date

End Class

<TypeConverter(GetType(propGridConverter))>
Public Class driftPercent

#Region "    Constructor"

    Public Sub New()

    End Sub

    Public Sub New(aquaMetFactor As Double)

    End Sub

    <DefaultValue(Double.NaN)>
    <TypeConverter(GetType(dblConv))>
    <Browsable(True)>
    <Category(catInputs)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property aquaMetFactor As Double = Double.NaN



    <Browsable(False)>
    <XmlIgnore> <ScriptIgnore>
    Public Property collapseStd As String() =
    {catDistances, catRegressionParameters, catSingle, catMulti}

#End Region

    ''' <summary>
    ''' Input complete ?
    ''' </summary>
    ''' <returns></returns>
    <Category(catInputs)>
    <XmlIgnore> <ScriptIgnore>
    <Browsable(False)>
    Public ReadOnly Property inputComplete As String
        Get

            If FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined Then
                Return "SWASH Crop ?"
            End If

            If FOCUSswWaterBody = eFOCUSswWaterBody.not_defined Then
                Return "Water-body ?"
            End If

            If noOfApplns = eNoOfApplns.not_defined Then
                Return "# of applns ?"
            End If

            If bufferWidth = eBufferWidth.not_defined Then
                Return "Buffer with ?"
            End If

            Return "OK!"

        End Get
    End Property

    Public Const catAppln As String = " 01 Application "

#Region "    Appln Info"

    ''' <summary>
    ''' Name
    ''' </summary>
    ''' <returns></returns>
    <Category(catAppln)>
    <Browsable(True)>
    <DisplayName(
    "Name")>
    Public ReadOnly Property name As String

        Get

            Dim out As New StringBuilder
            Dim space As String = "                      " & vbCrLf
            Dim DateString = ""
            Dim aqMetString As String = ""
            Dim multiString As String
            Dim singlString As String
            Dim offset As Integer = 7

            If Double.IsNaN(step03Single) Then
                Return " - "
            End If

            If Me.eventDate <> New Date Then
                DateString = dateConv.convDate2String(value:=Me.eventDate, format:="dd-MMM-yy")
            Else
                DateString = ""
            End If


            'aqua met
            If Double.IsNaN(aquaMetFactor) Or aquaMetFactor <= 0 Then
                aqMetString = ""
            Else

                If Me.noOfApplns > eNoOfApplns.one Then
                    multiString = step03MultiPECAquaMet.ToString("0.0000").PadLeft(offset) & " µg/L "
                Else
                    multiString = ""
                End If

                singlString = step03SinglePECAquaMet.ToString("0.0000").PadLeft(offset) & " µg/L "
                aqMetString = "Fs3 AqMet " & multiString & singlString & space
                'step03SinglePECAquaMet.ToString("0.0000").PadLeft(7) & " µg/L" & multiString & DateString & space
            End If

            singlString = step03SinglePEC.ToString("0.0000").PadLeft(offset) & " µg/L "

            If Me.noOfApplns > eNoOfApplns.one Then

                multiString = step03MultiPEC.ToString("0.0000").PadLeft(offset) & " µg/L "

                Return _
                    "Fs3 Par   " & multiString & singlString & DateString & space &
                    aqMetString &
                    "Fs4 05m   " & step04_05mMulti.ToString("0.0000").PadLeft(offset) & " µg/L " &
                                   step04_05m.ToString("0.0000").PadLeft(offset) & " µg/L" & " " & space &
                    "Fs4 10m   " & step04_10mMulti.ToString("0.0000").PadLeft(offset) & " µg/L " &
                                   step04_10m.ToString("0.0000").PadLeft(offset) & " µg/L" & " " & space &
                    "Fs4 15m   " & step04_15mMulti.ToString("0.0000").PadLeft(offset) & " µg/L " &
                                   step04_15m.ToString("0.0000").PadLeft(offset) & " µg/L" & " " & space &
                    "Fs4 20m   " & step04_20mMulti.ToString("0.0000").PadLeft(offset) & " µg/L " &
                                   step04_20m.ToString("0.0000").PadLeft(offset) & " µg/L"

            Else

                multiString = ""

                Return _
                    "Fs3 Par   " & multiString & singlString & DateString & space &
                    aqMetString &
                    "Fs4 05m   " & step04_05m.ToString("0.0000").PadLeft(offset) & " µg/L" & " " & space &
                    "Fs4 10m   " & step04_10m.ToString("0.0000").PadLeft(offset) & " µg/L" & " " & space &
                    "Fs4 15m   " & step04_15m.ToString("0.0000").PadLeft(offset) & " µg/L" & " " & space &
                    "Fs4 20m   " & step04_20m.ToString("0.0000").PadLeft(offset) & " µg/L" & " "

            End If








            out.Append(
                conv2String(
                value:=step03SinglePEC,
                format:="0.000",
                unit:=" µg/L"))

            If noOfApplns > eNoOfApplns.one Then

                out.Append(" " & vbCrLf)

                out.Append(
                           conv2String(
                           value:=step04Multi,
                           format:=driftPercentFormat,
                           unit:=driftPercentUnit))

            End If

            Return out.ToString

        End Get

    End Property

    ''' <summary>
    ''' Show item in own window
    ''' </summary>
    ''' <returns></returns>
    <Category(catAppln)>
    <Editor(GetType(buttonEmulator), GetType(UITypeEditor))>
    <DisplayName(
    "Zoom ...")>
    <Description(
    "Show item In own window")>
    <XmlIgnore> <ScriptIgnore>
    <Browsable(True)>
    <DefaultValue("")>
    Public Property show As frmPropGrid.eViewEdit
        Get
            Return frmPropGrid.eViewEdit.not_def
        End Get
        Set

            If Value = frmPropGrid.eViewEdit.not_def Then Exit Property

            Dim frm As New frmPropGrid(class2Show:=Me, classType:=Me.GetType)

            With frm

                .Width = 900
                .Height = 550

            End With

            If Value = frmPropGrid.eViewEdit.edit Then

                frm.ShowDialog()
                Me.CopyPropertiesByName(src:=frm.class2Show)

            Else
                frm.Show()
            End If

        End Set
    End Property

    ''' <summary>
    ''' Event date
    ''' </summary>
    ''' <returns></returns>
    <Category(catAppln)>
    <DisplayName(
        "Event Date")>
    <Description(
        "" & vbCrLf &
        "")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <XmlIgnore> <ScriptIgnore>
    Public Property eventDate As Date
        Get
            Return _eventDate
        End Get
        Set
            _eventDate = Value
        End Set
    End Property

    ''' <summary>
    ''' Target crop out of the
    ''' available FOCUS crops
    ''' </summary>
    <Category(catAppln)>
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
                Case eFOCUSswDriftCrop.PFE, eFOCUSswDriftCrop.PFL,
                          eFOCUSswDriftCrop.OL,
                          eFOCUSswDriftCrop.CI,
                          eFOCUSswDriftCrop.VI,
                          eFOCUSswDriftCrop.HP,
                          eFOCUSswDriftCrop.VIL
                    _applnMethodStep03 = eApplnMethodStep03.airBlast

                Case Else
                    _applnMethodStep03 = eApplnMethodStep03.groundSpray
            End Select

            _FOCUSswDriftCrop = Value
            RaiseEvent update()

        End Set
    End Property

    ''' <summary>
    ''' If crop = pome fruits then
    ''' early or late appln.?
    ''' </summary>
    ''' <returns></returns>
    <Category(catAppln)>
    <DisplayName(
        "PF Early or Late?")>
    <Description(
        "If crop = pome fruits then" & vbCrLf &
        "early or late appln.?")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(CInt(eEarlyLate.not_defined))>
    Public Property earlyLate As eEarlyLate
        Get
            Return _earlyLate
        End Get
        Set

            If _FOCUSswDriftCrop <> eFOCUSswDriftCrop.PFE OrElse _FOCUSswDriftCrop <> eFOCUSswDriftCrop.PFL Then
                _earlyLate = eEarlyLate.not_defined
            Else
                _earlyLate = Value
            End If

            RaiseEvent update()

        End Set
    End Property

    ''' <summary>
    ''' Number of applications
    ''' 1 - 8 or more
    ''' </summary>
    <Category(catAppln)>
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
    Public Property noOfApplns As eNoOfApplns
        Get
            Return _noOfApplns
        End Get
        Set

            _noOfApplns = Value

            If _noOfApplns = eNoOfApplns.one Then
                interval = Integer.MaxValue
            End If

            RaiseEvent update()

        End Set
    End Property

#Region "    Interval : Time between applns. in days"

    ''' <summary>
    ''' GUI : Time between applns. in days
    ''' </summary>
    ''' <returns></returns>
    <Category(catAppln)>
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
    <Category(catAppln)>
    <DefaultValue(Integer.MaxValue)>
    <Browsable(False)>
    <TypeConverter(GetType(intConv))>
    Public Property interval As Integer


#End Region

    ''' <summary>
    ''' Appln Rate
    ''' </summary>
    ''' <returns></returns>
    <Category(catAppln)>
    <DisplayName(
        "Rate")>
    <Description(
        "Application rate" & vbCrLf &
        "in kg as/ha")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= 'G4'|unit=' kg as/ha'")>
    Public Property rate As Double
        Get
            Return _rate
        End Get
        Set

            _rate = Value
            RaiseEvent update()

        End Set
    End Property

    ''' <summary>
    ''' Appln method
    ''' </summary>
    ''' <returns></returns>
    <Category(catAppln)>
    <DisplayName("Method")>
    Public Property applnMethodStep03 As eApplnMethodStep03
        Get
            Return _applnMethodStep03
        End Get
        Set

            Select Case Value

                Case eApplnMethodStep03.soilIncorp

                    If _FOCUSswDriftCrop = eFOCUSswDriftCrop.PFE OrElse
                            _FOCUSswDriftCrop = eFOCUSswDriftCrop.PFL OrElse
                        _FOCUSswDriftCrop = eFOCUSswDriftCrop.OL OrElse
                        _FOCUSswDriftCrop = eFOCUSswDriftCrop.CI OrElse
                        _FOCUSswDriftCrop = eFOCUSswDriftCrop.HP OrElse
                        _FOCUSswDriftCrop = eFOCUSswDriftCrop.VI OrElse
                        _FOCUSswDriftCrop = eFOCUSswDriftCrop.VIL Then

                        Value = _applnMethodStep03

                    End If

                    _Ganzelmeier = eGanzelmeier.noDrift

                Case eApplnMethodStep03.airBlast

                    If _FOCUSswDriftCrop = eFOCUSswDriftCrop.PFE AndAlso
                            _FOCUSswDriftCrop = eFOCUSswDriftCrop.PFL AndAlso
                        _FOCUSswDriftCrop <> eFOCUSswDriftCrop.OL AndAlso
                        _FOCUSswDriftCrop <> eFOCUSswDriftCrop.CI AndAlso
                        _FOCUSswDriftCrop <> eFOCUSswDriftCrop.HP AndAlso
                        _FOCUSswDriftCrop <> eFOCUSswDriftCrop.VI AndAlso
                        _FOCUSswDriftCrop <> eFOCUSswDriftCrop.VIL Then

                        Value = _applnMethodStep03

                    End If

                Case eApplnMethodStep03.granular


                Case eApplnMethodStep03.aerial

                    _Ganzelmeier = eGanzelmeier.AerialAppln

                Case eApplnMethodStep03.handHigh

                    _Ganzelmeier = eGanzelmeier.Vines_Late

                Case eApplnMethodStep03.handLow

                    _Ganzelmeier = eGanzelmeier.ArableCrops

            End Select

            _applnMethodStep03 = Value
            RaiseEvent update()

        End Set
    End Property

    Private Sub setbyApplnMethod()

        Select Case _applnMethodStep03

            Case eApplnMethodStep03.soilIncorp,
                 eApplnMethodStep03.granular

                _Ganzelmeier = eGanzelmeier.noDrift

            Case eApplnMethodStep03.aerial

                _Ganzelmeier = eGanzelmeier.AerialAppln

            Case eApplnMethodStep03.handHigh

                _Ganzelmeier = eGanzelmeier.Vines_Late

            Case eApplnMethodStep03.handLow

                _Ganzelmeier = eGanzelmeier.ArableCrops

        End Select

    End Sub

    Private Sub checkInputs()

        Select Case _applnMethodStep03

            ' special appln methods
            Case eApplnMethodStep03.soilIncorp,
                 eApplnMethodStep03.granular

                _Ganzelmeier = eGanzelmeier.noDrift

            Case eApplnMethodStep03.aerial

                _Ganzelmeier = eGanzelmeier.AerialAppln

            Case eApplnMethodStep03.handHigh

                _Ganzelmeier = eGanzelmeier.Vines_Late

            Case eApplnMethodStep03.handLow

                _Ganzelmeier = eGanzelmeier.ArableCrops

            Case Else

                'std appln method : ground spray
                Select Case _FOCUSswDriftCrop

                    Case eFOCUSswDriftCrop.not_defined

                        _earlyLate = eEarlyLate.not_defined
                        _Ganzelmeier = eGanzelmeier.not_defined
                        _applnMethodStep03 = eApplnMethodStep03.not_defined

                    Case eFOCUSswDriftCrop.PFE, eFOCUSswDriftCrop.PFL,
                         eFOCUSswDriftCrop.OL,
                         eFOCUSswDriftCrop.CI,
                         eFOCUSswDriftCrop.VI,
                         eFOCUSswDriftCrop.VIL,
                         eFOCUSswDriftCrop.HP

                        Select Case _applnMethodStep03

                            Case eApplnMethodStep03.not_defined

                                _applnMethodStep03 = eApplnMethodStep03.airBlast

                                Select Case _FOCUSswDriftCrop

                                    Case eFOCUSswDriftCrop.PFE, eFOCUSswDriftCrop.PFL

                                        Select Case _earlyLate

                                            Case eEarlyLate.not_defined

                                                _earlyLate = eEarlyLate.early
                                                _Ganzelmeier = eGanzelmeier.FruitCrops_Early

                                            Case eEarlyLate.early

                                                _Ganzelmeier = eGanzelmeier.FruitCrops_Early

                                            Case eEarlyLate.late

                                                _Ganzelmeier = eGanzelmeier.FruitCrops_Late

                                        End Select

                                    Case eFOCUSswDriftCrop.VI

                                        _Ganzelmeier = eGanzelmeier.Vines_Late
                                        _earlyLate = eEarlyLate.not_defined

                                    Case Else

                                        _Ganzelmeier = convertFOCUSCrop2Ganzelmeier(FOCUSswDriftCrop)
                                        _earlyLate = eEarlyLate.not_defined

                                End Select

                            Case eApplnMethodStep03.groundSpray

                                _Ganzelmeier = eGanzelmeier.ArableCrops

                            Case eApplnMethodStep03.airBlast

                                Select Case _FOCUSswDriftCrop

                                    Case eFOCUSswDriftCrop.PFE, eFOCUSswDriftCrop.PFL

                                        Select Case _earlyLate

                                            Case eEarlyLate.not_defined

                                                _earlyLate = eEarlyLate.early
                                                _Ganzelmeier = eGanzelmeier.FruitCrops_Early

                                            Case eEarlyLate.early

                                                _Ganzelmeier = eGanzelmeier.FruitCrops_Early

                                            Case eEarlyLate.late

                                                _Ganzelmeier = eGanzelmeier.FruitCrops_Late

                                        End Select

                                    Case eFOCUSswDriftCrop.VI

                                        _Ganzelmeier = eGanzelmeier.Vines_Late
                                        _earlyLate = eEarlyLate.not_defined

                                    Case Else

                                        _Ganzelmeier = convertFOCUSCrop2Ganzelmeier(FOCUSswDriftCrop)
                                        _earlyLate = eEarlyLate.not_defined

                                End Select
                                _Ganzelmeier = convertFOCUSCrop2Ganzelmeier(FOCUSswDriftCrop)

                            Case eApplnMethodStep03.handHigh

                                _applnMethodStep03 = eApplnMethodStep03.groundSpray
                                _Ganzelmeier = eGanzelmeier.Vines_Late

                            Case eApplnMethodStep03.handLow
                                _applnMethodStep03 = eApplnMethodStep03.groundSpray
                                _Ganzelmeier = eGanzelmeier.ArableCrops

                        End Select

                        'Case eFOCUSswDriftCrop.HL

                        'Case eFOCUSswDriftCrop.HH


                    Case Else

                        If _applnMethodStep03 = eApplnMethodStep03.airBlast Then
                            _applnMethodStep03 = eApplnMethodStep03.groundSpray
                        End If

                        _earlyLate = eEarlyLate.not_defined

                        _Ganzelmeier = convertFOCUSCrop2Ganzelmeier(FOCUSswDriftCrop)

                End Select


        End Select

    End Sub

    Private Function updateRegression() As Boolean

        If _applnMethodStep03 = eApplnMethodStep03.aerial Then

            A = regressionA(eGanzelmeier.AerialAppln,
                            eNoOfApplns.one)

            B = regressionB(eGanzelmeier.AerialAppln,
                            eNoOfApplns.one)

            C = regressionC(eGanzelmeier.AerialAppln,
                            eNoOfApplns.one)

            D = regressionD(eGanzelmeier.AerialAppln,
                            eNoOfApplns.one)

            Return True

        ElseIf Me.FOCUSswDriftCrop <> eFOCUSswDriftCrop.not_defined AndAlso
               Me.Ganzelmeier <> eGanzelmeier.not_defined AndAlso
               Me.noOfApplns <> eNoOfApplns.not_defined AndAlso
               Me.Ganzelmeier <> eGanzelmeier.noDrift AndAlso
               Me.applnMethodStep03 <> eApplnMethodStep03.granular AndAlso
               Me.applnMethodStep03 <> eApplnMethodStep03.soilIncorp Then

            A = regressionA(Ganzelmeier,
                            noOfApplns)

            B = regressionB(Ganzelmeier,
                            noOfApplns)

            C = regressionC(Ganzelmeier,
                            noOfApplns)

            D = regressionD(Ganzelmeier,
                            noOfApplns)

            Return True

        Else

            A = Double.NaN
            B = Double.NaN
            C = Double.NaN
            D = Double.NaN

            Return False

        End If

    End Function

    Public Event update()


    Public Enum eChangePart

        FOCUScrop
        earlyLate
        number
        rate
        applnMethod
        rest

    End Enum

    Private Sub driftPercent_update() Handles Me.update


        checkInputs()

        If Not updateRegression() Then

            With Me

                .distanceCrop2Bank = Double.NaN
                .distanceBank2Water = Double.NaN
                .closest2EdgeOfField = Double.NaN


                .nearestDriftPercentSingle = Double.NaN
                .nearestDriftPercentMulti = Double.NaN

                .farthestDriftPercentSingle = Double.NaN
                .farthestDriftPercentMulti = Double.NaN

                .step12Single = Double.NaN
                .step12SinglePEC = Double.NaN
                .step03SingleLoading = Double.NaN
                .step03SinglePEC = Double.NaN
                .step04Single = Double.NaN
                .step04SingleLoading = Double.NaN
                .step04SinglePEC = Double.NaN
                .step04_05m = Double.NaN
                .step04_10m = Double.NaN
                .step04_15m = Double.NaN
                .step04_20m = Double.NaN

                .step12Multi = Double.NaN
                .step12MultiPEC = Double.NaN
                .step03MultiLoading = Double.NaN
                .step03MultiPEC = Double.NaN
                .step04Multi = Double.NaN
                .step04MultiLoading = Double.NaN
                .step04MultiPEC = Double.NaN
                .step04_05mMulti = Double.NaN
                .step04_10mMulti = Double.NaN
                .step04_15mMulti = Double.NaN
                .step04_20mMulti = Double.NaN

            End With

            Exit Sub

        End If

        recalc(driftPercent:=Me)


        Dim step04Drift As Double

#Region "    single"

        Me.step12Single =
            calcDriftPercentStep12(
            noOfApplns:=eNoOfApplns.one,
            FOCUSswDriftCrop:=FOCUSswDriftCrop)

        Me.step12SinglePEC = rate * Me.step12Single / 0.3

        Me.step03SingleLoading = Me.rate * Me.step03Single

        Me.step03SinglePEC = Me.step03SingleLoading / Me.waterDepth

        Me.step04SingleLoading = Me.rate * Me.step04Single

        Me.step04SinglePEC = Me.step04SingleLoading / Me.waterDepth

#Region "    5 - 20m"

        step04Drift = calcDriftPercent(
                              noOfApplns:=eNoOfApplns.one,
                        FOCUSswDriftCrop:=Me.FOCUSswDriftCrop,
                        FOCUSswWaterBody:=Me.FOCUSswWaterBody,
                           driftDistance:=eDriftDistance.average,
                             bufferWidth:=5,
                                  Nozzle:=0,
                                  herbicideUse:=IIf(applnMethodStep03 = eApplnMethodStep03.groundSpray, True, False),
                                  earlyLate:=earlyLate)

        Me.step04_05m = Me.rate * step04Drift / Me.waterDepth


        step04Drift = calcDriftPercent(
                              noOfApplns:=eNoOfApplns.one,
                        FOCUSswDriftCrop:=Me.FOCUSswDriftCrop,
                        FOCUSswWaterBody:=Me.FOCUSswWaterBody,
                           driftDistance:=eDriftDistance.average,
                             bufferWidth:=10,
                                  Nozzle:=0,
                                  herbicideUse:=IIf(applnMethodStep03 = eApplnMethodStep03.groundSpray, True, False),
                                  earlyLate:=earlyLate)

        Me.step04_10m = Me.rate * step04Drift / Me.waterDepth


        step04Drift = calcDriftPercent(
                              noOfApplns:=eNoOfApplns.one,
                        FOCUSswDriftCrop:=Me.FOCUSswDriftCrop,
                        FOCUSswWaterBody:=Me.FOCUSswWaterBody,
                           driftDistance:=eDriftDistance.average,
                             bufferWidth:=15,
                                  Nozzle:=0,
                                  herbicideUse:=IIf(applnMethodStep03 = eApplnMethodStep03.groundSpray, True, False),
                                  earlyLate:=earlyLate)

        Me.step04_15m = Me.rate * step04Drift / Me.waterDepth

        step04Drift = calcDriftPercent(
                              noOfApplns:=eNoOfApplns.one,
                        FOCUSswDriftCrop:=Me.FOCUSswDriftCrop,
                        FOCUSswWaterBody:=Me.FOCUSswWaterBody,
                           driftDistance:=eDriftDistance.average,
                             bufferWidth:=20,
                                  Nozzle:=0,
                                  herbicideUse:=IIf(applnMethodStep03 = eApplnMethodStep03.groundSpray, True, False),
                                  earlyLate:=earlyLate)

        Me.step04_20m = Me.rate * step04Drift / Me.waterDepth


#End Region

        'Me.step2aSingle = calcDriftPercentStep12(
        '        noOfApplns:=eNoOfApplns.one,
        '        FOCUSswDriftCrop:=FOCUSswDriftCrop,
        '        bufferWidth:=bufferWidth)

        maxNozzleSingle =
                   calcMaxNozzle(
                   maxReduction:=maxTotalReduction,
                   noOfApplns:=eNoOfApplns.one,
                   FOCUSswDriftCrop:=FOCUSswDriftCrop,
                   FOCUSswWaterBody:=FOCUSswWaterBody,
                   bufferWidth:=_bufferWidth)

        maxBufferSingle = calcMaxBuffer(
                    maxReduction:=maxTotalReduction,
                    noOfApplns:=eNoOfApplns.one,
                    FOCUSswDriftCrop:=FOCUSswDriftCrop,
                    FOCUSswWaterBody:=FOCUSswWaterBody,
                    nozzle:=nozzle)

#End Region

#Region "    multi"

        If Me.applnMethodStep03 = eApplnMethodStep03.aerial AndAlso
           Me.noOfApplns > eNoOfApplns.one Then

            Me.step12Multi = Me.step12Single
            Me.step12MultiPEC = Me.step12SinglePEC
            Me.step03SingleLoading = Me.step03MultiLoading
            Me.step03SinglePEC = Me.step03MultiPEC
            Me.step04SingleLoading = Me.step04MultiLoading
            Me.step04SinglePEC = Me.step04MultiPEC
            Me.step04_05mMulti = Me.step04_05m
            Me.step04_10mMulti = Me.step04_10m
            Me.step04_15mMulti = Me.step04_15m
            Me.step04_20mMulti = Me.step04_20m

        ElseIf noOfApplns > eNoOfApplns.one Then


            Me.step12Multi =
                    calcDriftPercentStep12(
                    noOfApplns:=noOfApplns,
                    FOCUSswDriftCrop:=FOCUSswDriftCrop)

            Me.step12MultiPEC = rate * Me.step12Multi / 0.3

            Me.step03MultiLoading = Me.rate * Me.step03Multi

            Me.step03MultiPEC = Me.step03MultiLoading / Me.waterDepth

            Me.step04MultiLoading = Me.rate * Me.step04Multi

            Me.step04MultiPEC = Me.step04MultiLoading / Me.waterDepth

#Region "    5 - 20m"


            step04Drift = calcDriftPercent(
                              noOfApplns:=noOfApplns,
                        FOCUSswDriftCrop:=Me.FOCUSswDriftCrop,
                        FOCUSswWaterBody:=Me.FOCUSswWaterBody,
                           driftDistance:=eDriftDistance.average,
                             bufferWidth:=5,
                                  nozzle:=0,
                                  herbicideUse:=IIf(applnMethodStep03 = eApplnMethodStep03.groundSpray, True, False),
                                  earlyLate:=earlyLate)

            Me.step04_05mMulti = Me.rate * step04Drift / Me.waterDepth


            step04Drift = calcDriftPercent(
                              noOfApplns:=noOfApplns,
                        FOCUSswDriftCrop:=Me.FOCUSswDriftCrop,
                        FOCUSswWaterBody:=Me.FOCUSswWaterBody,
                           driftDistance:=eDriftDistance.average,
                             bufferWidth:=10,
                                  nozzle:=0,
                                  herbicideUse:=IIf(applnMethodStep03 = eApplnMethodStep03.groundSpray, True, False),
                                  earlyLate:=earlyLate)

            Me.step04_10mMulti = Me.rate * step04Drift / Me.waterDepth


            step04Drift = calcDriftPercent(
                              noOfApplns:=noOfApplns,
                        FOCUSswDriftCrop:=Me.FOCUSswDriftCrop,
                        FOCUSswWaterBody:=Me.FOCUSswWaterBody,
                           driftDistance:=eDriftDistance.average,
                             bufferWidth:=15,
                                  nozzle:=0,
                                  herbicideUse:=IIf(applnMethodStep03 = eApplnMethodStep03.groundSpray, True, False),
                                  earlyLate:=earlyLate)

            Me.step04_15mMulti = Me.rate * step04Drift / Me.waterDepth

            step04Drift = calcDriftPercent(
                              noOfApplns:=noOfApplns,
                        FOCUSswDriftCrop:=Me.FOCUSswDriftCrop,
                        FOCUSswWaterBody:=Me.FOCUSswWaterBody,
                           driftDistance:=eDriftDistance.average,
                             bufferWidth:=20,
                                  nozzle:=0,
                                  herbicideUse:=IIf(applnMethodStep03 = eApplnMethodStep03.groundSpray, True, False),
                                  earlyLate:=earlyLate)

            Me.step04_20mMulti = Me.rate * step04Drift / Me.waterDepth


#End Region

            Me.step2aMulti =
                               calcDriftPercentStep12(
                                noOfApplns:=noOfApplns,
                                FOCUSswDriftCrop:=FOCUSswDriftCrop,
                                bufferWidth:=bufferWidth)

            maxNozzleMulti =
                          calcMaxNozzle(
                          maxReduction:=maxTotalReduction,
                          noOfApplns:=noOfApplns,
                          FOCUSswDriftCrop:=FOCUSswDriftCrop,
                          FOCUSswWaterBody:=FOCUSswWaterBody,
                          bufferWidth:=_bufferWidth)


            maxBufferMulti = calcMaxBuffer(
                   maxReduction:=maxTotalReduction,
                   noOfApplns:=noOfApplns,
                   FOCUSswDriftCrop:=FOCUSswDriftCrop,
                   FOCUSswWaterBody:=FOCUSswWaterBody,
                   nozzle:=nozzle)

        Else

            nearestDriftPercentMulti = Double.NaN
            farthestDriftPercentMulti = Double.NaN

            step12Multi = Double.NaN
            step2aMulti = Double.NaN
            step03Multi = Double.NaN
            step04Multi = Double.NaN
            step04_05mMulti = Double.NaN
            step04_10mMulti = Double.NaN
            step04_15mMulti = Double.NaN
            step04_20mMulti = Double.NaN

            maxNozzleMulti = eNozzles.not_defined
            maxBufferMulti = eNozzles.not_defined

        End If

#End Region

    End Sub

#End Region

    Public Const catInputs As String = "01  FOCUS"

#Region "    FOCUS"

#Region "    -  "

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _eventDate As New Date

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _FOCUSswDriftCrop As eFOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _Ganzelmeier As eGanzelmeier = eGanzelmeier.not_defined

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _FOCUSswWaterBody As eFOCUSswWaterBody = eFOCUSswWaterBody.not_defined

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _noOfApplns As eNoOfApplns = eNoOfApplns.not_defined

    '<DebuggerBrowsable(DebuggerBrowsableState.Never)>
    'Private _BBCH As New BBCH

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _bufferWidth As eBufferWidth = eBufferWidth.FOCUSStep03

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _nozzle As eNozzles

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _maxTotalReduction As Integer = 95

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _waterDepth As Double = 0.3

#End Region


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
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(CInt(eGanzelmeier.not_defined))>
    Public Property Ganzelmeier As eGanzelmeier
        Get
            Return _Ganzelmeier
        End Get
        Set

            _Ganzelmeier = Value

            updateRegression()
            RaiseEvent update()

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
    Public Property scenario As ePRZMScenario
        Get
            Return _scenario
        End Get
        Set
            _scenario = Value

            If _scenario <> eFOCUSswScenario.not_defined Then
                waterDepth = runOffWaterDepths(_scenario)
            Else
                waterDepth = 0.3
            End If

        End Set
    End Property

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
            RaiseEvent update()

        End Set
    End Property

    ''' <summary>
    ''' Dimensions
    ''' lenght x width x depth; volume
    ''' </summary>
    ''' <returns></returns>
    <Category(catInputs)>
    <DisplayName(
    "Dimensions")>
    <Description(
    "lenght x width x depth; volume" & vbCrLf &
    "")>
    <RefreshProperties(RefreshProperties.All)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(enumConverter(Of Type).not_defined)>
    Public ReadOnly Property dimensions As String
        Get

            If _FOCUSswWaterBody = eFOCUSswWaterBody.not_defined Then
                Return enumConverter(Of Type).not_defined
            Else
                Return wbDimensions(_FOCUSswWaterBody)
            End If

        End Get
    End Property


    ''' <summary>
    ''' Depth in m
    ''' </summary>
    ''' <returns></returns>
    <Category(catInputs)>
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
            RaiseEvent update()

        End Set
    End Property

    ''' <summary>
    ''' Buffer width in m
    ''' Step03 = std. FOCUS buffer width
    ''' </summary>
    ''' <returns></returns>
    <Category(catInputs)>
    <DisplayName(
        "Buffer")>
    <Description(
        "Buffer width in m" & vbCrLf &
        "Step03 = std. FOCUS buffer width")>
    <RefreshProperties(RefreshProperties.All)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(CInt(eBufferWidth.FOCUSStep03))>
    Public Property bufferWidth As eBufferWidth
        Get
            Return _bufferWidth
        End Get
        Set(value As eBufferWidth)

            _bufferWidth = value
            RaiseEvent update()

        End Set
    End Property

    ''' <summary>
    ''' Drift reducing nozzle in %
    ''' </summary>
    ''' <returns></returns>
    <Category(catInputs)>
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
    Public Property nozzle As eNozzles
        Get
            Return _nozzle
        End Get
        Set

            _nozzle = Value
            RaiseEvent update()

        End Set
    End Property



    ''' <summary>
    ''' Total drift reduction compared to Step03
    ''' "Buffer + Nozzle in percent, single appln.
    ''' </summary>
    ''' <returns></returns>
    <Category(catInputs)>
    <DisplayName(
    "Total Reduction")>
    <Description(
    "Total drift reduction compared to Step03" & vbCrLf &
    "Buffer + Nozzle in percent, single appln")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
    "format= '0'" &
    "|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public ReadOnly Property totalDriftSingle As Double
        Get

            If Double.IsNaN(step04Single) OrElse
                   Double.IsNaN(step03Single) OrElse
                   step03Single = 0 Then

                Return Double.NaN

            Else

                Return 100 - (step04Single * 100 / step03Single)

            End If

        End Get
    End Property

    ''' <summary>
    ''' Maximum allowed total reduction in %
    ''' </summary>
    ''' <returns></returns>
    <Category(catInputs)>
    <DisplayName(
    "Max Allowed")>
    <Description(
    "Maximum allowed total reduction in %" & vbCrLf &
    "")>
    <RefreshProperties(RefreshProperties.All)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <TypeConverter(GetType(intConv))>
    <AttributeProvider(
    "format= '0" &
    "'|unit='" & driftPercentUnit & "'")>
    <DefaultValue(95)>
    Public Property maxTotalReduction As Integer
        Get
            Return _maxTotalReduction
        End Get
        Set

            _maxTotalReduction = Value
            RaiseEvent update()

        End Set
    End Property

    ''' <summary>
    ''' Max nozzle at given buffer 
    ''' to stay within the max reduction, single appln
    ''' </summary>
    ''' <returns></returns>
    <Category(catInputs)>
    <DisplayName(
    "Max Nozzle")>
    <Description(
    "Max nozzle at given buffer " & vbCrLf &
    "to stay within the max reduction")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(intConv))>
    <AttributeProvider(
    "format= '0" &
    "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    Public Property maxNozzleSingle As eNozzles
        Get
            Return _maxNozzleSingle
        End Get
        Set
            _maxNozzleSingle = Value
        End Set
    End Property

    ''' <summary>
    ''' Max buffer at given nozzle 
    ''' to stay within the max reduction, single appln
    ''' </summary>
    ''' <returns></returns>
    <Category(catInputs)>
    <DisplayName(
    "Max Buffer")>
    <Description(
    "Max buffer at given nozzle " & vbCrLf &
    "to stay within the max reduction")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <XmlIgnore> <ScriptIgnore>
    Public Property maxBufferSingle As eBufferWidth
        Get
            Return _maxBufferSingle
        End Get
        Set
            _maxBufferSingle = Value
        End Set
    End Property

    ''' <summary>
    ''' PEC single appln at Step04 level incl. Buffer + Nozzle
    ''' </summary>
    ''' <returns></returns>
    <Category(catInputs)>
    <DisplayName(
    "Step 04 Single")>
    <Description(
    "PEC at Step04 level incl. Buffer + Nozzle" & vbCrLf &
    "in µg/L, single appln")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " µg/L" & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step04SinglePEC As Double
        Get
            Return _step04SinglePEC
        End Get
        Set
            _step04SinglePEC = Value
        End Set
    End Property

    ''' <summary>
    ''' PEC multi applns at Step04 level incl. Buffer + Nozzle
    ''' </summary>
    ''' <returns></returns>
    <Category(catInputs)>
    <DisplayName(
    "Step 04 Multi")>
    <Description(
    "PEC at Step04 level incl. Buffer + Nozzle" & vbCrLf &
    "in µg/L, multi applns")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
    "format= '" & "G4" &
    "'|unit='" & " µg/L" & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step04MultiPEC As Double
        Get
            Return _step04MultiPEC
        End Get
        Set
            _step04MultiPEC = Value
        End Set
    End Property


#End Region

    Public Const catDistances As String = "02  Distances"

#Region "    Distances from crop, drift percentages"

    Private Const offset As String = "    "
    Private Const driftPercentFormat As String = "0.000"
    Private Const driftPercentUnit As String = " %"

    Private Const bufferFormat As String = "0.0"
    Private Const bufferUnit As String = " m"

#Region "   -   "

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _distanceCrop2Bank As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _distanceBank2Water As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _hingePoint As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _closest2EdgeOfField As Double = Double.NaN

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
    Private _nearestDriftPercentSingle As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _nearestDriftPercentMulti As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _farthest2EdgeOfWB As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _step12Multi As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _farthestDriftPercentSingle As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _farthestDriftPercentMulti As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _maxNozzleSingle As eNozzles = eNozzles.not_defined

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _maxBufferSingle As eBufferWidth = eBufferWidth.not_defined

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

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _rate As Double = Double.NaN

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _applnMethodStep03 As eApplnMethodStep03 = eApplnMethodStep03.not_defined

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _intervalGUI As String

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _earlyLate As eEarlyLate = eEarlyLate.not_defined



#End Region


    ''' <summary>
    ''' Distance  Crop to Bank in m 
    ''' </summary>
    <Category(catDistances)>
    <DisplayName(
        "Crop <-> Bank")>
    <Description(
        "Distance Crop <-> Bank" & vbCrLf &
        "in m")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & bufferFormat &
        "'|unit='" & bufferUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property distanceCrop2Bank As Double
        Get
            Return _distanceCrop2Bank
        End Get
        Set
            _distanceCrop2Bank = Value
        End Set
    End Property

    ''' <summary>
    ''' Distance  Bank to Water in m
    ''' </summary>
    <Category(catDistances)>
    <DisplayName(
        "Bank <-> Water")>
    <Description(
        "Distance Bank <-> Water" & vbCrLf &
        "in m")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & bufferFormat &
        "'|unit='" & bufferUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property distanceBank2Water As Double
        Get
            Return _distanceBank2Water
        End Get
        Set
            _distanceBank2Water = Value
        End Set
    End Property

    ''' <summary>
    ''' Hinge Point
    ''' Distance limit for each regression in m
    ''' to switch from A + B to C + D
    ''' </summary>
    <Category(catDistances)>
    <DisplayName(
        "Hinge Point")>
    <Description(
        "Distance limit for each regression in m" & vbCrLf &
        "to switch from A/B to C/D")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & bufferFormat &
        "'|unit='" & bufferUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property hingePoint As Double
        Get
            Return _hingePoint
        End Get
        Set
            _hingePoint = Value
        End Set
    End Property

    ''' <summary>
    ''' Closest to the edge of the field in m
    ''' </summary>
    <Category(catDistances)>
    <RefreshProperties(RefreshProperties.All)>
    <DisplayName(
        "edge nearest  field")>
    <Description(
        "Closest to the edge of the field in m" & vbCrLf &
        "FOCUS std. Buffer length: Crop2Bank + Bank2Water")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & bufferFormat &
        "'|unit='" & bufferUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property closest2EdgeOfField As Double
        Get
            Return _closest2EdgeOfField
        End Get
        Set
            _closest2EdgeOfField = Value
        End Set
    End Property

    ''' <summary>
    ''' Drift at edge nearest field in %, single appln.
    ''' </summary>
    ''' <returns></returns>
    <Category(catDistances)>
    <DisplayName(
        "Drift ,single")>
    <Description(
        "Drift at edge nearest field" & vbCrLf &
        "in %, single appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & driftPercentFormat &
        "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property nearestDriftPercentSingle As Double
        Get
            Return _nearestDriftPercentSingle
        End Get
        Set
            _nearestDriftPercentSingle = Value
        End Set
    End Property

    ''' <summary>
    ''' Drift at edge nearest field in %, multi applns.
    ''' </summary>
    ''' <returns></returns>
    <Category(catDistances)>
    <DisplayName(
        "       multi")>
    <Description(
        "Drift at edge nearest field" & vbCrLf &
        "in %, multi applns.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & driftPercentFormat &
        "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property nearestDriftPercentMulti As Double
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
    <Category(catDistances)>
    <RefreshProperties(RefreshProperties.All)>
    <DisplayName(
        "farthest from field")>
    <Description(
        "Farthest to the edge of the field in m" & vbCrLf &
        "Closest2Edge + water body width")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & bufferFormat &
        "'|unit='" & bufferUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property farthest2EdgeOfWB As Double
        Get
            Return _farthest2EdgeOfWB
        End Get
        Set
            _farthest2EdgeOfWB = Value
        End Set
    End Property

    ''' <summary>
    ''' Drift farthest to the 
    ''' edge of the field in %, single appln.
    ''' </summary>
    ''' <returns></returns>
    <Category(catDistances)>
    <DisplayName(
         "Drift ,single")>
    <Description(
        "Drift farthest to the edge of the field" & vbCrLf &
        "in %, single appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & driftPercentFormat &
        "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property farthestDriftPercentSingle As Double
        Get
            Return _farthestDriftPercentSingle
        End Get
        Set
            _farthestDriftPercentSingle = Value
        End Set
    End Property

    ''' <summary>
    ''' Drift farthest to the 
    ''' edge of the field in %, multi applns.
    ''' </summary>
    ''' <returns></returns>
    <Category(catDistances)>
    <DisplayName(
         "       multi")>
    <Description(
        "Drift farthest to the edge of the field" & vbCrLf &
        "in %, multi applns.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & driftPercentFormat &
        "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property farthestDriftPercentMulti As Double
        Get
            Return _farthestDriftPercentMulti
        End Get
        Set
            _farthestDriftPercentMulti = Value
        End Set
    End Property

#End Region

    Public Const catRegressionParameters As String = "03  Regression Parameter"

#Region "    Regression Parameter"

    <Category(catRegressionParameters)>
    <Description(
        "" & vbCrLf &
        "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property A As Double = Double.NaN

    <Category(catRegressionParameters)>
    <Description(
        "" & vbCrLf &
        "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    <AttributeProvider(
        "format= '" & driftPercentFormat & "'")>
    Public Property B As Double = Double.NaN

    <Category(catRegressionParameters)>
    <Description(
        "" & vbCrLf &
        "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property C As Double = Double.NaN

    <Category(catRegressionParameters)>
    <Description(
        "" & vbCrLf &
        "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property D As Double = Double.NaN

#End Region

    Public Const catSingle As String = "04  Single Drift % "

#Region "    Drift % Single Appln."

    ''' <summary>
    ''' Drift at Step 12 level
    ''' in %, single appln.
    ''' </summary>
    ''' <returns></returns>
    <Category(catSingle)>
    <DisplayName(
        "Step 12")>
    <Description(
        "Drift at Step 12 level" & vbCrLf &
        "in %, single appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & driftPercentFormat &
        "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step12Single As Double
        Get
            Return _step12Single
        End Get
        Set
            _step12Single = Value
        End Set
    End Property

    ''' <summary>
    ''' Step03 drift for comparison
    ''' </summary>
    ''' <returns></returns>
    <Category(catSingle)>
    <DisplayName(
        "Step 03")>
    <Description(
        "Step03 drift value for comparison" & vbCrLf &
        "in %, single appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & driftPercentFormat &
        "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step03Single As Double
        Get
            Return _step03Single
        End Get
        Set
            _step03Single = Value
        End Set
    End Property

    ''' <summary>
    ''' Areic mean drift in %
    ''' </summary>
    ''' <returns></returns>
    <Category(catSingle)>
    <DisplayName(
        "Step 04")>
    <Description(
        "Areic mean drift" & vbCrLf &
        "in %, single appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & driftPercentFormat &
        "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step04Single As Double
        Get
            Return _step04SingleSingle
        End Get
        Set
            _step04SingleSingle = Value
        End Set
    End Property

#End Region

    Public Const catSinglePECs As String = "05         PECs "

#Region "    PECs Single Appln"

    ''' <summary>
    ''' Step03 drift for comparison
    ''' </summary>
    ''' <returns></returns>
    <Category(catSinglePECs)>
    <DisplayName(
        "PEC Step 03")>
    <Description(
        "PEC  at Step03 level" & vbCrLf &
        "in µg/L, single appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " µg/L" & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step03SinglePEC As Double
        Get
            Return _step03SinglePEC
        End Get
        Set
            _step03SinglePEC = Value
        End Set
    End Property


#Region "    std. distances 5, 10 ,15 and 20m"

    ''' <summary>
    ''' Step 04 single, 05m
    ''' </summary>
    ''' <returns></returns>
    <Category(catSinglePECs)>
    <DisplayName(
    "Step 04, 05m")>
    <Description(
    "" & vbCrLf &
    "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " µg/L" & "'")>
    <DefaultValue(Double.NaN)>
    Public Property step04_05m As Double = Double.NaN

    ''' <summary>
    ''' Step 04 single, 05m
    ''' </summary>
    ''' <returns></returns>
    <Category(catSinglePECs)>
    <DisplayName(
    "         10m")>
    <Description(
    "" & vbCrLf &
    "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " µg/L" & "'")>
    <DefaultValue(Double.NaN)>
    Public Property step04_10m As Double = Double.NaN

    ''' <summary>
    ''' Step 04 single, 05m
    ''' </summary>
    ''' <returns></returns>
    <Category(catSinglePECs)>
    <DisplayName(
    "         15m")>
    <Description(
    "" & vbCrLf &
    "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " µg/L" & "'")>
    <DefaultValue(Double.NaN)>
    Public Property step04_15m As Double = Double.NaN

    ''' <summary>
    ''' Step 04 single, 05m
    ''' </summary>
    ''' <returns></returns>
    <Category(catSinglePECs)>
    <DisplayName(
    "         20m")>
    <Description(
    "" & vbCrLf &
    "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " µg/L" & "'")>
    <DefaultValue(Double.NaN)>
    Public Property step04_20m As Double = Double.NaN

#End Region


    ''' <summary>
    ''' PEC at Step 12 level
    ''' in µg/L, single appln.
    ''' </summary>
    ''' <returns></returns>
    <Category(catSinglePECs)>
    <DisplayName(
        "PEC Step 12")>
    <Description(
        "PEC at Step 12 level" & vbCrLf &
        "in µg/L, single appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " µg/L" & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step12SinglePEC As Double
        Get
            Return _step12SinglePEC
        End Get
        Set
            _step12SinglePEC = Value
        End Set
    End Property


    ''' <summary>
    ''' Step03 aquatic metabolite
    ''' </summary>
    ''' <returns></returns>
    <Category(catSinglePECs)>
    <DisplayName(
        "    Step 03 Aquatic. Met")>
    <Description(
        "Aquatic. Met PEC at Step03 level" & vbCrLf &
        "in µg/L, single appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " µg/L" & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    <RefreshProperties(RefreshProperties.All)>
    Public ReadOnly Property step03SinglePECAquaMet As Double
        Get

            If Double.IsNaN(step03Single) OrElse Double.IsNaN(aquaMetFactor) Then
                Return Double.NaN
            Else
                Return step03SinglePEC * aquaMetFactor
            End If

        End Get

    End Property

    ''' <summary>
    ''' Step03 drift for comparison
    ''' </summary>
    ''' <returns></returns>
    <Category(catSinglePECs)>
    <DisplayName(
    "Loading Step 03")>
    <Description(
        "Mass loading per drift event" & vbCrLf &
        "in mg/m², single appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " mg/m²" & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step03SingleLoading As Double
        Get
            Return _step03SingleLoading
        End Get
        Set
            _step03SingleLoading = Value
        End Set
    End Property

    ''' <summary>
    ''' Step03 drift for comparison
    ''' </summary>
    ''' <returns></returns>
    <Category(catSinglePECs)>
    <DisplayName(
    "        Step 04")>
    <Description(
        "Mass loading per drift event" & vbCrLf &
        "in mg/m², single appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " mg/m²" & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step04SingleLoading As Double
        Get
            Return _step04SingleLoading
        End Get
        Set
            _step04SingleLoading = Value
        End Set
    End Property

#End Region

    Public Const catMulti As String = "06  Multi  Drift %"

#Region "    Drift % Multi Applns"

    ''' <summary>
    ''' Drift at Step 1/2 level
    ''' in %, multi applns.
    ''' </summary>
    ''' <returns></returns>
    <Category(catMulti)>
    <DisplayName(
        "Step 12")>
    <Description(
        "Drift at Step 12 level" & vbCrLf &
        "in %, multi applns.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & driftPercentFormat &
        "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step12Multi As Double
        Get
            Return _step12Multi
        End Get
        Set
            _step12Multi = Value
        End Set
    End Property


    <Category(catMulti)>
    <DisplayName(
        "Step 2a")>
    <Description(
        "Drift at Step 2 level" & vbCrLf &
        "in % incl. buffer, multi applns")>
    <Browsable(False)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & driftPercentFormat &
        "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step2aMulti As Double


    ''' <summary>
    ''' Step03 drift for comparison
    ''' </summary>
    ''' <returns></returns>
    <Category(catMulti)>
    <DisplayName(
    "Step 03")>
    <Description(
    "Step03 drift value for comparison" & vbCrLf &
    "in %, multi applns.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
    "format= '" & driftPercentFormat &
    "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step03Multi As Double
        Get
            Return _step03Multi
        End Get
        Set
            _step03Multi = Value
        End Set
    End Property

    ''' <summary>
    ''' Areic mean drift in %
    ''' </summary>
    ''' <returns></returns>
    <Category(catMulti)>
    <DisplayName(
    "Step 04")>
    <Description(
    "Areic mean drift" & vbCrLf &
    "in %, multi applns.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
    "format= '" & driftPercentFormat &
    "'|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step04Multi As Double
        Get
            Return _step04Multi
        End Get
        Set
            _step04Multi = Value
        End Set
    End Property

    ''' <summary>
    ''' Total drift reduction compared to Step03
    ''' Buffer + Nozzle in percent, multi applns.
    ''' </summary>
    ''' <returns></returns>
    <Category(catMulti)>
    <DisplayName(
    "Total Reduction")>
    <Description(
    "Total drift reduction compared to Step03 " & vbCrLf &
    "Buffer + Nozzle in percent, multi applns.")>
    <Browsable(False)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
    "format= '0'" &
    "|unit='" & driftPercentUnit & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public ReadOnly Property totalDriftMulti As Double
        Get

            Dim temp As Double

            If Double.IsNaN(step04Multi) OrElse
                   Double.IsNaN(step03Multi) OrElse
                   step03Multi = 0 Then

                Return Double.NaN

            Else

                temp = (step04Multi * 100 / step03Multi)

                Return 100 - (step04Multi * 100 / step03Multi)

            End If

        End Get

    End Property

    <Category(catMulti)>
    <DisplayName(
            "Max Nozzle")>
    <Description(
            "Max nozzle at given buffer " & vbCrLf &
            "to stay within the max reduction, multi applns.")>
    <Browsable(False)>
    <[ReadOnly](True)>
    <XmlIgnore> <ScriptIgnore>
    Public Property maxNozzleMulti As eNozzles = eNozzles.not_defined

    <Category(catMulti)>
    <DisplayName(
            "Max Buffer")>
    <Description(
            "Max buffer at given nozzle " & vbCrLf &
            "to stay within the max reduction, multi applns.")>
    <Browsable(False)>
    <[ReadOnly](True)>
    <XmlIgnore> <ScriptIgnore>
    Public Property maxBufferMulti As eBufferWidth = eBufferWidth.not_defined


#End Region

    Public Const catMultiPECs As String = "07         PECs"

#Region "    PECs Multi Applns"

    ''' <summary>
    ''' Step03 drift for comparison
    ''' </summary>
    ''' <returns></returns>
    <Category(catMultiPECs)>
    <DisplayName(
        "PEC Step 03")>
    <Description(
        "PEC at Step03 level" & vbCrLf &
        "in µg/L, multi appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " µg/L" & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step03MultiPEC As Double
        Get
            Return _step03MultiPEC
        End Get
        Set
            _step03MultiPEC = Value
        End Set
    End Property


#Region "    std. distances 5, 10 ,15 and 20m"

    ''' <summary>
    ''' Step 04 multi, 05m
    ''' </summary>
    ''' <returns></returns>
    <Category(catMultiPECs)>
    <DisplayName(
    "Step 04, 05m")>
    <Description(
    "" & vbCrLf &
    "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " µg/L" & "'")>
    <DefaultValue(Double.NaN)>
    Public Property step04_05mMulti As Double = Double.NaN

    ''' <summary>
    ''' Step 04 multi, 05m
    ''' </summary>
    ''' <returns></returns>
    <Category(catMultiPECs)>
    <DisplayName(
    "         10m")>
    <Description(
    "" & vbCrLf &
    "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " µg/L" & "'")>
    <DefaultValue(Double.NaN)>
    Public Property step04_10mMulti As Double = Double.NaN

    ''' <summary>
    ''' Step 04 multi, 05m
    ''' </summary>
    ''' <returns></returns>
    <Category(catMultiPECs)>
    <DisplayName(
    "         15m")>
    <Description(
    "" & vbCrLf &
    "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " µg/L" & "'")>
    <DefaultValue(Double.NaN)>
    Public Property step04_15mMulti As Double = Double.NaN

    ''' <summary>
    ''' Step 04 multi, 05m
    ''' </summary>
    ''' <returns></returns>
    <Category(catMultiPECs)>
    <DisplayName(
    "         20m")>
    <Description(
    "" & vbCrLf &
    "")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " µg/L" & "'")>
    <DefaultValue(Double.NaN)>
    Public Property step04_20mMulti As Double = Double.NaN

#End Region


    ''' <summary>
    ''' PEC at Step 12 level
    ''' in µg/L, multi appln.
    ''' </summary>
    ''' <returns></returns>
    <Category(catMultiPECs)>
    <DisplayName(
        "PEC Step 12")>
    <Description(
        "PEC at Step 12 level" & vbCrLf &
        "in mg/L, Multi appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " µg/L" & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step12MultiPEC As Double
        Get
            Return _step12MultiPEC
        End Get
        Set
            _step12MultiPEC = Value
        End Set
    End Property


    ''' <summary>
    ''' Step03 drift for comparison
    ''' </summary>
    ''' <returns></returns>
    <Category(catMultiPECs)>
    <DisplayName(
        "    Step 03 Aquatic. Met")>
    <Description(
        "Aquatic. Met PEC at Step03 level" & vbCrLf &
        "in µg/L, multi appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " µg/L" & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public ReadOnly Property step03MultiPECAquaMet As Double
        Get

            If Double.IsNaN(step03Multi) OrElse Double.IsNaN(aquaMetFactor) Then
                Return Double.NaN
            Else
                Return step03MultiPEC * aquaMetFactor
            End If


            Return _step03MultiPEC
        End Get
    End Property


    ''' <summary>
    ''' Step03 drift for comparison
    ''' </summary>
    ''' <returns></returns>
    <Category(catMultiPECs)>
    <DisplayName(
    "Loading Step 03")>
    <Description(
        "Mass loading per drift event" & vbCrLf &
        "in mg/m², multi appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " mg/m²" & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step03MultiLoading As Double
        Get
            Return _step03MultiLoading
        End Get
        Set
            _step03MultiLoading = Value
        End Set
    End Property

    ''' <summary>
    ''' Step04 drift for comparison
    ''' </summary>
    ''' <returns></returns>
    <Category(catMultiPECs)>
    <DisplayName(
    "        Step 04")>
    <Description(
        "Mass loading per drift event" & vbCrLf &
        "in mg/m², multi appln.")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <TypeConverter(GetType(dblConv))>
    <AttributeProvider(
        "format= '" & "G4" &
        "'|unit='" & " mg/m²" & "'")>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(Double.NaN)>
    Public Property step04MultiLoading As Double
        Get
            Return _step04MultiLoading
        End Get
        Set
            _step04MultiLoading = Value
        End Set
    End Property

#End Region

End Class

<TypeConverter(GetType(propGridConverter))>
Public Class BBCHalt

    Public Sub New()

    End Sub

    <Browsable(False)>
    Public ReadOnly Property name As String
        Get

            If BBCHmin = -1 OrElse BBCHmax = 101 Then
                Return " - "
            End If

            Return guiBBCHmin & " - " & guiBBCHmax
        End Get
    End Property

    Private BBCHsmin As String() =
        {
        "0",
        "10",
        "20",
        "30",
        "40",
        "50",
        "60",
        "70",
        "80",
        "90"
        }

    Private BBCHsmax As String() =
        {
        "9",
        "19",
        "29",
        "39",
        "49",
        "59",
        "69",
        "79",
        "89",
        "99"
        }


    <Browsable(False)>
    Public Property BBCHmin As Integer = -1

    <TypeConverter(GetType(dropDownList))>
    <DisplayName(
        "BBCH min")>
    <Description(
        "" & vbCrLf &
        "")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <XmlIgnore> <ScriptIgnore>
    Public Property guiBBCHmin As String
        Get

            dropDownList.dropDownEntries = BBCHsmin
            Return BBCHmin.ToString

        End Get
        Set(value As String)

            Dim temp As Integer

            Try
                temp = Integer.Parse(s:=value)
            Catch ex As Exception
                temp = -99
            End Try

            If temp >= 0 AndAlso temp <= 99 AndAlso temp <= BBCHmax Then
                BBCHmin = temp
            End If

        End Set
    End Property

    <Browsable(False)>
    Public Property BBCHmax As Integer = 101

    <TypeConverter(GetType(dropDownList))>
    <DisplayName(
        "BBCH max")>
    <Description(
        "" & vbCrLf &
        "")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <XmlIgnore> <ScriptIgnore>
    Public Property guiBBCHmax As String
        Get

            dropDownList.dropDownEntries = BBCHsmax
            Return BBCHmax.ToString

        End Get
        Set(value As String)

            Dim temp As Integer

            Try
                temp = Integer.Parse(s:=value)
            Catch ex As Exception
                temp = -99
            End Try

            If temp >= 0 AndAlso temp <= 99 AndAlso temp >= BBCHmin Then
                BBCHmax = temp
            End If

        End Set
    End Property


End Class

Public Module FOCUSdriftDB

#Region "    enums"

    ''' <summary>
    ''' Drift red. nozzles
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eNozzles)))>
    Public Enum eNozzles

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        <Description("0")>
        _0 = 0

        <Description("10")>
        _10 = 10

        <Description("20")>
        _20 = 20

        <Description("25")>
        _25 = 25

        <Description("30")>
        _30 = 30

        <Description("40")>
        _40 = 40

        <Description("50")>
        _50 = 50

        <Description("60")>
        _60 = 60

        <Description("70")>
        _70 = 70

        <Description("75")>
        _75 = 75

        <Description("80")>
        _80 = 80

        <Description("90")>
        _90 = 90

        <Description("95")>
        _95 = 95

    End Enum



    ''' <summary>
    ''' FOCUSsw DRIFT crops as enumeration
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eFOCUSswDriftCrop)))>
    Public Enum eFOCUSswDriftCrop

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        ''' <summary>
        ''' CS Cereals, spring
        ''' </summary>
        <Description(
                "CS " & vbCrLf &
                "Cereals, spring")>
        CS = 0

        ''' <summary>
        ''' CW Cereals, winter
        ''' </summary>
        <Description(
                "CW " & vbCrLf &
                "Cereals, winter")>
        CW

        '-------------------------------------------------------------

        ''' <summary>
        ''' CI Citrus
        ''' </summary>
        <Description(
                "CI " & vbCrLf &
                "Citrus")>
        CI

        ''' <summary>
        ''' CO Cotton
        ''' </summary>
        <Description(
                "CO " & vbCrLf &
                "Cotton")>
        CO

        '-------------------------------------------------------------

        ''' <summary>
        ''' FB Field beans 1st/2nd
        ''' </summary>
        <Description(
                "FB " & vbCrLf &
                "Field beans" & vbCrLf &
                "1st/2nd")>
        FB

        ''' <summary>
        ''' GA Grass/alfalfa
        ''' </summary>
        <Description(
                "GA " & vbCrLf &
                "Grass/alfalfa")>
        GA

        '-------------------------------------------------------------

        ''' <summary>
        ''' HP Hops
        ''' </summary>
        <Description(
                "HP " & vbCrLf &
                "Hops")>
        HP

        '-------------------------------------------------------------

        ''' <summary>
        ''' LG Legumes
        ''' </summary>
        <Description(
                "LG " & vbCrLf &
                "Legumes")>
        LG

        ''' <summary>
        ''' MZ Maize
        ''' </summary>
        <Description(
                "MZ " & vbCrLf &
                "Maize")>
        MZ

        '-------------------------------------------------------------

        ''' <summary>
        ''' OS Oil seed rape, sprin
        ''' </summary>
        <Description(
                "OS " & vbCrLf &
                "Oilseed rape, spring")>
        OS

        ''' <summary>
        ''' OW Oil seed rape, winter
        ''' </summary>
        <Description(
                "OW " & vbCrLf &
                "Oilseed rape, winter")>
        OW

        '-------------------------------------------------------------

        ''' <summary>
        ''' OL Olives
        ''' </summary>
        <Description(
                "OL " & vbCrLf &
                "Olives")>
        OL


        ''' <summary>
        ''' PFE Pome/stone fruits
        ''' early BBCH 01 - 71, 96 - 99
        ''' </summary>
        <Description(
                "PFE " & vbCrLf &
                "Pome fruit" & vbCrLf &
                "early BBCH 01 - 71, 96 - 99")>
        PFE

        ''' <summary>
        ''' PFL Pome/stone fruits
        ''' late  BBCH 72 - 95
        ''' </summary>
        <Description(
                "PFL " & vbCrLf &
                "Pome fruit" & vbCrLf &
                "late  BBCH 72 - 95")>
        PFL

        '-------------------------------------------------------------

        ''' <summary>
        ''' PS Potatoes, 1st/2nd
        ''' </summary>
        <Description(
                "PS " & vbCrLf &
                "Potatoes" & vbCrLf &
                "1st/2nd")>
        PS

        ''' <summary>
        ''' SY Soybean
        ''' </summary>
        <Description(
                "SY " & vbCrLf &
                "Soybeans")>
        SY

        ''' <summary>
        ''' SB Sugar beets
        ''' </summary>
        <Description(
                "SB " & vbCrLf &
                "Sugar beets")>
        SB_Sugar_beets

        ''' <summary>
        ''' SU Sunflowers
        ''' </summary>
        <Description(
                "SU " & vbCrLf &
                "Sunflowers")>
        SU

        '-------------------------------------------------------------

        ''' <summary>
        ''' TB Tobacco
        ''' </summary>
        <Description(
                "TB " & vbCrLf &
                "Tobacco")>
        TB

        '-------------------------------------------------------------

        ''' <summary>
        ''' VB Vegetables, bulb, 1st/2nd
        ''' </summary>
        <Description(
                "VB " & vbCrLf &
                "Vegetables, bulb, 1st/2nd")>
        VB

        ''' <summary>
        ''' 
        ''' </summary>
        <Description(
                "VF " & vbCrLf &
                "Vegetables, fruiting")>
        VF

        ''' <summary>
        ''' VL Vegetables, leafy, 1st/2nd
        ''' </summary>
        <Description(
                "VL " & vbCrLf &
                "Vegetables, leafy, 1st/2nd")>
        VL

        ''' <summary>
        ''' VR Vegetables, root, 1st/2nd
        ''' </summary>
        <Description(
                "VR " & vbCrLf &
                "Vegetables, root, 1st/2nd")>
        VR

        '-------------------------------------------------------------

        ''' <summary>
        ''' VI Vines
        ''' </summary>
        <Description(
                "VI " & vbCrLf &
                "Vines")>
        VI

        '-------------------------------------------------------------


        ''' <summary>
        ''' VIL Vines, late, for compatibility reasons
        ''' </summary>
        <Description(
                "VIL " & vbCrLf &
                "Vines, late")>
        VIL


        '''' <summary>
        '''' Aerial appln.
        '''' </summary>
        '<Description(
        '    "AA " & vbCrLf &
        '    "Aerial appln." & vbCrLf &
        '    "")>
        'AA = 24

        '''' <summary>
        '''' HL Appln, hand (crop < 50 cm)
        '''' </summary>
        '<Description(
        '    "HL " & vbCrLf &
        '    "Appln, hand (crop < 50 cm)")>
        'HL

        '''' <summary>
        '''' HH Appln, hand (crop > 50 cm)
        '''' </summary>
        '<Description(
        '    "HH " & vbCrLf &
        '    "Appln, hand (crop > 50 cm)")>
        'HH

        '''' <summary>
        '''' no drift (incorporatio or seed treatment)
        '''' </summary>
        '<Description(
        '    "ND " & vbCrLf &
        '    "No drift ")>
        'ND_NoDrift

    End Enum

    Public cropXScenarios As String() =
                {
                    "D1|D3|D4|D5|R4",
                    "D1|D2|D3|D4|D5|D6|R1|R3|R4",
                    "D6|R4",
                    "D6",
                    "D2|D3|D4|D6|R1|R2|R3|R4",
                    "D1|D2|D3|D4|D5|R2|R3",
                    "R1",
                    "D3|D4|D5|D6|R1|R2|R3|R4",
                    "D3|D4|D5|D6|R1|R2|R3|R4",
                    "D1|D3|D4|D5|R1",
                    "D2|D3|D4|D5|R1|R3",
                    "D6|R4",
                    "D3|D4|D5|R1|R2|R3|R4",
                    "D3|D4|D6|R1|R2|R3",
                    "R3|R4",
                    "D3|D4|R1|R3",
                    "D5|R1|R3|R4",
                    "R3",
                    "D3|D4|D6|R1|R2|R3|R4",
                    "D6|R2|R3|R4",
                    "D3|D4|D6|R1|R2|R3|R4",
                    "D3|D6|R1|R2|R3|R4",
                    "D6|R1|R2|R3|R4"
                }



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
                "BBCH 01 - 71, 96 - 99 " & vbCrLf &
                "23.599 %")>
        FruitCrops_Early

        ''' <summary>
        ''' Fruit crops, late
        ''' BBCH 72 - 95, 11.134 %
        ''' </summary>
        <Description(
                "Fruit crops, late " & vbCrLf &
                "BBCH 72 - 95 " & vbCrLf &
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
                "Vines, early " & vbCrLf &
                "1.7184 %")>
        Vines_Early

        ''' <summary>
        ''' Vines, late 5.173 %
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
        <Description("no drift " & vbCrLf &
                "0 % ;-)")>
        noDrift

    End Enum

    ''' <summary>
    ''' all FOCUSsw scenarios D1-6 and R1-4
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eFOCUSswScenario)))>
    Public Enum eFOCUSswScenario

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        <Description(
                "D1 " & vbCrLf &
                "Lanna " & vbCrLf &
                "ditch|stream")>
        D1 = 0

        <Description(
                "D2 " & vbCrLf &
                "Brimstone " & vbCrLf &
                "ditch|stream")>
        D2

        <Description(
                "D3 " & vbCrLf &
                "Vredepeel " & vbCrLf &
                "ditch")>
        D3

        <Description(
                "D4 " & vbCrLf &
                "Skousbo " & vbCrLf &
                "pond|stream")>
        D4

        <Description(
                "D5 " & vbCrLf &
                "La Jailliere " & vbCrLf &
                "pond|stream")>
        D5

        <Description(
                "D6 " & vbCrLf &
                "Thiva " & vbCrLf &
                "ditch")>
        D6

        <Description(
                "R1 " & vbCrLf &
                "Weiherbach " & vbCrLf &
                "pond|stream")>
        R1

        <Description(
                "R2 " & vbCrLf &
                "Porto " & vbCrLf &
                "stream")>
        R2

        <Description(
                "R3 " & vbCrLf &
                "Bologna " & vbCrLf &
                "stream")>
        R3

        <Description(
                "R4 " & vbCrLf &
                "Roujan " & vbCrLf &
                "stream")>
        R4

    End Enum

    ''' <summary>
    ''' FOCUSsw water body
    ''' ditch, stream or pond
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eFOCUSswWaterBody)))>
    Public Enum eFOCUSswWaterBody

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        <Description(
                "DI " & vbCrLf &
                "ditch ")>
        ditch = 0

        <Description(
                "ST " & vbCrLf &
                "stream ")>
        stream

        <Description(
                "PO " & vbCrLf &
                "pond ")>
        pond

    End Enum

    Public wbDimensions As String() =
        {"100m " & myConst.multiply & " 1m " & myConst.multiply & " 0.3m; 30,000L",
         "100m " & myConst.multiply & " 1m " & myConst.multiply & " 0.3m; 30,000L",
         "30m radius " & myConst.multiply & " 1.0m; ~94,000L"}

    ''' <summary>
    ''' Number of applns.
    ''' 1 - 8 (or more)
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eNoOfApplns)))>
    Public Enum eNoOfApplns

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        <Description("1")>
        one = 0

        <Description("2")>
        two

        <Description("3")>
        three

        <Description("4")>
        four

        <Description("5")>
        fife

        <Description("6")>
        six

        <Description("7")>
        seven

        <Description("8+")>
        eight

    End Enum

    ''' <summary>
    ''' appln Interval
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eApplnInterval)))>
    Public Enum eApplnInterval

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        <Description(" 1")> _01 = 1
        <Description(" 2")> _02 = 2
        <Description(" 3")> _03 = 3
        <Description(" 4")> _04 = 4

        <Description(" 5")>
        _05 = 5
        <Description(" 7")>
        _07 = 7
        <Description("10")>
        _10 = 10
        <Description("12")>
        _12 = 12
        <Description("14")>
        _14 = 14
        <Description("21")>
        _21 = 21
        <Description("28")>
        _28 = 28


        '<Description(" 6")> _06 = 6

        '<Description(" 8")> _08 = 8
        '<Description(" 9")> _09 = 9

        '<Description("11")> _11 = 11

        '<Description("13")> _13 = 13

        <Description("15")> _15 = 15
        '<Description("16")> _16 = 16
        '<Description("17")> _17 = 17
        '<Description("18")> _18 = 18
        '<Description("19")> _19 = 19
        '<Description("20")> _20 = 20

        '<Description("22")> _22 = 22
        '<Description("23")> _23 = 23
        '<Description("24")> _24 = 24
        '<Description("25")> _25 = 25
        '<Description("26")> _26 = 26
        '<Description("27")> _27 = 27

        '<Description("29")> _29 = 29
        <Description("30")> _30 = 30
        '<Description("31")> _31 = 31
        '<Description("32")> _32 = 32
        '<Description("33")> _33 = 33
        '<Description("34")> _34 = 34
        '<Description("35")> _35 = 35
        '<Description("36")> _36 = 36
        '<Description("37")> _37 = 37
        '<Description("38")> _38 = 38
        '<Description("39")> _39 = 39
        '<Description("40")> _40 = 40
        '<Description("41")> _41 = 41
        <Description("42")> _42 = 42
        '<Description("43")> _43 = 43
        '<Description("44")> _44 = 44
        '<Description("45")> _45 = 45
        '<Description("46")> _46 = 46
        '<Description("47")> _47 = 47
        '<Description("48")> _48 = 48
        '<Description("49")> _49 = 49
        <Description("50")> _50 = 50
        <Description("60")> _60 = 60

    End Enum

    ''' <summary>
    ''' Buffer width as enum for simple input, -1 = Step03 std ;-)
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eBufferWidth)))>
    Public Enum eBufferWidth

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        <Description("Step03, no add. buffer")>
        FOCUSStep03 = 0

        <Description(" 1 ( < FOCUS std. buffer? )")>
        _01 = 1

        <Description(" 2 ( < FOCUS std. buffer? )")>
        _02 = 2

        <Description(" 3 ( < FOCUS std. buffer? )")>
        _03 = 3

        <Description(" 4 ( < FOCUS std. buffer? )")>
        _04 = 4

        <Description(" 5")>
        _05 = 5

        <Description("10")>
        _10 = 10

        <Description("15")>
        _15 = 15

        <Description("20")>
        _20 = 20

        <Description("30")>
        _30 = 30

        <Description("40")>
        _40 = 40

        <Description("50")>
        _50 = 50

    End Enum

    ''' <summary>
    ''' PAT version
    ''' one FOCUS year or 20 year assessment
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eSWASHVersion)))>
    Public Enum eSWASHVersion
        <Description("5.5.3, one target year")>
        _553
        <Description("6.6.4, 20 years")>
        _664
    End Enum

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
    ''' appln method 
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eApplnMethodStep03)))>
    Public Enum eApplnMethodStep03

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        ''' <summary>
        ''' aerial appln.
        ''' </summary>
        <Description("aerial appln.")>
        aerial

        ''' <summary>
        ''' air blast, for hops, pome/stone fruits, vines, olives and citrus
        ''' </summary>
        <Description("air blast")>
        airBlast

        ''' <summary>
        ''' granular appln., no drift and interception
        ''' </summary>
        <Description("granular appln.")>
        granular

        ''' <summary>
        ''' ground spray, std. incl. drift and interception
        ''' </summary>
        <Description("ground spray")>
        groundSpray

        ''' <summary>
        ''' soil incorp. needs depth (std. 4cm)
        ''' </summary>
        <Description("soil incorp.")>
        soilIncorp

        ''' <summary>
        ''' HL Appln, hand (crop less than 50 cm)
        ''' </summary>
        <Description(
               "Hand appln. (crop < 50 cm)")>
        handLow

        ''' <summary>
        ''' HH Appln, hand (crop higher than 50 cm)
        ''' </summary>
        <Description(
                "Hand appln.  (crop > 50 cm)")>
        handHigh


    End Enum

    <TypeConverter(GetType(enumConverter(Of eCAM)))>
    Public Enum eCAM

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        SoilLinear = 1
        FoliarLinear = 2
        IncorpUniform = 4
        IncorpLinIncrease = 5
        IncorpLinDecrease = 6
        IncorpOneDepth = 8
        FoliarLinearInclInterception = 12

    End Enum

    <TypeConverter(GetType(enumConverter(Of eEarlyLate)))>
    Public Enum eEarlyLate

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        ''' <summary>
        ''' early/harvest appln BBCH 1-71 / 96-99, 23.6%
        ''' </summary>
        <Description("early: BBCH 1-71 / 96-99, 23.6%")>
        early

        ''' <summary>
        ''' late appln BBCH 72-95, 11.1%
        ''' </summary>
        <Description("late: BBCH 72-95, 11.1%")>
        late

    End Enum

#End Region

#Region "    functions"

    Public Sub recalc(ByRef driftPercent As driftPercent)

        Dim herbicideUse As Boolean = False
        Dim dummyFOCUSswDriftCrop As eFOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined

        With driftPercent

            dummyFOCUSswDriftCrop = .FOCUSswDriftCrop

            If _
                .FOCUSswDriftCrop <> eFOCUSswDriftCrop.not_defined AndAlso
                .FOCUSswWaterBody <> eFOCUSswWaterBody.not_defined AndAlso
                .noOfApplns <> eNoOfApplns.not_defined AndAlso
                .applnMethodStep03 <> eApplnMethodStep03.soilIncorp AndAlso
                .applnMethodStep03 <> eApplnMethodStep03.granular Then


                If .applnMethodStep03 = eApplnMethodStep03.handHigh Then

                    '    dummyFOCUSswDriftCrop = eFOCUSswDriftCrop.HH

                    'ElseIf .applnMethodStep03 = eApplnMethodStep03.handLow Then

                    '    dummyFOCUSswDriftCrop = eFOCUSswDriftCrop.HL

                    'ElseIf .applnMethodStep03 = eApplnMethodStep03.aerial Then

                    '    dummyFOCUSswDriftCrop = eFOCUSswDriftCrop.AA

                Else

                    dummyFOCUSswDriftCrop = .FOCUSswDriftCrop

                End If

                If .bufferWidth = eBufferWidth.FOCUSStep03 Then

                    .distanceCrop2Bank =
                        getFOCUSStdDistance(
                                FOCUSswDriftCrop:=dummyFOCUSswDriftCrop,
                                FOCUSswWaterBody:= .FOCUSswWaterBody,
                                FOCUSstdDistancesMember:=eFOCUSStdDistancesMember.edgeField2topBank)

                    .distanceBank2Water =
                            getFOCUSStdDistance(
                                    FOCUSswDriftCrop:=dummyFOCUSswDriftCrop,
                                    FOCUSswWaterBody:= .FOCUSswWaterBody,
                                    FOCUSstdDistancesMember:=eFOCUSStdDistancesMember.topBank2edgeWaterbody)

                    .closest2EdgeOfField =
                                    .distanceCrop2Bank + .distanceBank2Water

                    .farthest2EdgeOfWB =
                                    .distanceCrop2Bank +
                                    .distanceBank2Water +
                                    WaterBodyDB(.FOCUSswWaterBody, eWBMember.Witdth)

                Else

                    .distanceCrop2Bank = Double.NaN
                    .distanceBank2Water = Double.NaN

                    .closest2EdgeOfField = .bufferWidth
                    .farthest2EdgeOfWB =
                                .bufferWidth + WaterBodyDB(.FOCUSswWaterBody, eWBMember.Witdth)

                End If

                .hingePoint =
                        hingePointDB(
                                convertFOCUSCrop2Ganzelmeier(dummyFOCUSswDriftCrop),
                                .noOfApplns)

                .nearestDriftPercentSingle = calcDriftPercent(
                              noOfApplns:=eNoOfApplns.one,
                        FOCUSswDriftCrop:=dummyFOCUSswDriftCrop,
                        FOCUSswWaterBody:= .FOCUSswWaterBody,
                           driftDistance:=eDriftDistance.closest,
                             bufferWidth:= .bufferWidth,
                                  nozzle:= .nozzle,
                                  herbicideUse:=herbicideUse,
                                  earlyLate:= .earlyLate)

                .farthestDriftPercentSingle = calcDriftPercent(
                              noOfApplns:=eNoOfApplns.one,
                        FOCUSswDriftCrop:=dummyFOCUSswDriftCrop,
                        FOCUSswWaterBody:= .FOCUSswWaterBody,
                           driftDistance:=eDriftDistance.farthest,
                             bufferWidth:= .bufferWidth,
                                  nozzle:= .nozzle,
                                  herbicideUse:=herbicideUse,
                                  earlyLate:= .earlyLate)

                .step03Single = calcDriftPercent(
                              noOfApplns:=eNoOfApplns.one,
                        FOCUSswDriftCrop:=dummyFOCUSswDriftCrop,
                        FOCUSswWaterBody:= .FOCUSswWaterBody,
                           driftDistance:=eDriftDistance.average,
                             bufferWidth:=eBufferWidth.FOCUSStep03,
                                  nozzle:=eNozzles._0,
                                  herbicideUse:=herbicideUse,
                                  earlyLate:= .earlyLate)

                .step04Single = calcDriftPercent(
                            noOfApplns:=eNoOfApplns.one,
                    FOCUSswDriftCrop:=dummyFOCUSswDriftCrop,
                    FOCUSswWaterBody:= .FOCUSswWaterBody,
                        driftDistance:=eDriftDistance.average,
                        bufferWidth:= .bufferWidth,
                                nozzle:= .nozzle,
                                herbicideUse:=herbicideUse,
                                earlyLate:= .earlyLate)


                If .noOfApplns > eNoOfApplns.one Then

                    If .applnMethodStep03 = eApplnMethodStep03.aerial Then

                        .nearestDriftPercentMulti = .nearestDriftPercentSingle
                        .farthestDriftPercentMulti = .farthestDriftPercentSingle
                        .step03Multi = .step03Single
                        .step04Multi = .step04Single

                    Else

                        .nearestDriftPercentMulti = calcDriftPercent(
                                                      noOfApplns:= .noOfApplns,
                                                FOCUSswDriftCrop:=dummyFOCUSswDriftCrop,
                                                FOCUSswWaterBody:= .FOCUSswWaterBody,
                                                   driftDistance:=eDriftDistance.closest,
                                                     bufferWidth:= .bufferWidth,
                                                          nozzle:= .nozzle,
                                      herbicideUse:=herbicideUse,
                                      earlyLate:= .earlyLate)


                        .farthestDriftPercentMulti = calcDriftPercent(
                                                      noOfApplns:= .noOfApplns,
                                                FOCUSswDriftCrop:=dummyFOCUSswDriftCrop,
                                                FOCUSswWaterBody:= .FOCUSswWaterBody,
                                                   driftDistance:=eDriftDistance.farthest,
                                                     bufferWidth:= .bufferWidth,
                                                          nozzle:= .nozzle,
                                      herbicideUse:=herbicideUse,
                                      earlyLate:= .earlyLate)


                        .step03Multi = calcDriftPercent(
                                  noOfApplns:= .noOfApplns,
                            FOCUSswDriftCrop:=dummyFOCUSswDriftCrop,
                            FOCUSswWaterBody:= .FOCUSswWaterBody,
                               driftDistance:=eDriftDistance.average,
                                 bufferWidth:=eBufferWidth.FOCUSStep03,
                                      nozzle:=eNozzles._0,
                                      herbicideUse:=herbicideUse,
                                      earlyLate:= .earlyLate)

                        .step04Multi = calcDriftPercent(
                                            noOfApplns:= .noOfApplns,
                                    FOCUSswDriftCrop:=dummyFOCUSswDriftCrop,
                                    FOCUSswWaterBody:= .FOCUSswWaterBody,
                                        driftDistance:=eDriftDistance.average,
                                        bufferWidth:= .bufferWidth,
                                                nozzle:= .nozzle,
                            herbicideUse:=herbicideUse,
                            earlyLate:= .earlyLate)



                    End If

                Else

                    .nearestDriftPercentMulti = Double.NaN
                    .farthestDriftPercentMulti = Double.NaN
                    .step04Multi = Double.NaN

                End If

            Else

                .distanceCrop2Bank = Double.NaN
                .distanceBank2Water = Double.NaN
                .closest2EdgeOfField = Double.NaN


                .nearestDriftPercentSingle = Double.NaN
                .nearestDriftPercentMulti = Double.NaN

                .farthestDriftPercentSingle = Double.NaN
                .farthestDriftPercentMulti = Double.NaN

                .step12Single = Double.NaN
                .step12SinglePEC = Double.NaN
                .step03SingleLoading = Double.NaN
                .step03SinglePEC = Double.NaN
                .step04Single = Double.NaN
                .step04SingleLoading = Double.NaN
                .step04SinglePEC = Double.NaN
                .step04_05m = Double.NaN
                .step04_10m = Double.NaN
                .step04_15m = Double.NaN
                .step04_20m = Double.NaN

                .step12Multi = Double.NaN
                .step12MultiPEC = Double.NaN
                .step03MultiLoading = Double.NaN
                .step03MultiPEC = Double.NaN
                .step04Multi = Double.NaN
                .step04MultiLoading = Double.NaN
                .step04MultiPEC = Double.NaN
                .step04_05mMulti = Double.NaN
                .step04_10mMulti = Double.NaN
                .step04_15mMulti = Double.NaN
                .step04_20mMulti = Double.NaN

            End If

        End With

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
    ''' <param name="bufferWidth">
    ''' Buffer width in m, -1 = FOCUS std. width
    ''' </param>
    ''' <param name="driftDistance">
    ''' Nearest, farthest or average distance 
    ''' std. = average
    ''' </param>
    ''' <param name="herbicideUse">
    ''' use Ganzelmeier 'Arable Crops' 
    ''' and ignore 'FOCUSswDriftCrop'
    ''' </param>
    ''' <param name="earlyLate">
    ''' if crop  = 'PF Pome Fruit' select early or late use
    ''' early/harvest appln BBCH 1-71 / 96-99
    ''' late appln BBCH 72-95
    ''' </param>
    ''' <returns>
    ''' FOCUS drift in percent as double
    ''' </returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough>
    Public Function calcDriftPercent(
                            noOfApplns As eNoOfApplns,
                            FOCUSswDriftCrop As eFOCUSswDriftCrop,
                            FOCUSswWaterBody As eFOCUSswWaterBody,
                   Optional nozzle As Integer = eNozzles._0,
                   Optional bufferWidth As Integer = eBufferWidth.FOCUSStep03,
                   Optional driftDistance As eDriftDistance = eDriftDistance.average,
                   Optional herbicideUse As Boolean = False,
                   Optional earlyLate As eEarlyLate = eEarlyLate.not_defined) As Double

        'check inputs
        If noOfApplns = eNoOfApplns.not_defined OrElse
           FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined OrElse
           FOCUSswWaterBody = eFOCUSswWaterBody.not_defined Then

            Return Double.NaN

        End If

        If nozzle = CInt(eNozzles.not_defined) Then
            nozzle = eNozzles._0
        End If

        'drift for output
        Dim driftPercent As Double = Double.NaN

        'Application
        Dim Ganzelmeier As eGanzelmeier

        'regression parameters
        Dim A As Double
        Dim B As Double

        Dim C As Double
        Dim D As Double

        'set buffer width to FOCUS std.
        Dim hingePoint As Double
        Dim edgeField2topBank As Double
        Dim topBank2edgeWaterbody As Double
        Dim closest2EdgeOfField As Double
        Dim farthest2EdgeOfField As Double

        'drift for Step03 std. buffer distance?
        Dim step03stdBuffer As Boolean

        'check if buffer = Step03 std.
        step03stdBuffer = checkStep03StdBuffer(bufferWidth:=bufferWidth)

        'get Ganzelmeier crop group
        Select Case FOCUSswDriftCrop

            Case eFOCUSswDriftCrop.PFE, eFOCUSswDriftCrop.PFL

                If herbicideUse Then
                    Ganzelmeier = eGanzelmeier.ArableCrops

                ElseIf earlyLate = eEarlyLate.early Then
                    Ganzelmeier = eGanzelmeier.FruitCrops_Early
                Else
                    Ganzelmeier = eGanzelmeier.FruitCrops_Late
                End If

            Case eFOCUSswDriftCrop.VI,
                 eFOCUSswDriftCrop.VIL,
                 eFOCUSswDriftCrop.HP,
                 eFOCUSswDriftCrop.OL,
                 eFOCUSswDriftCrop.CI

                If herbicideUse Then
                    Ganzelmeier = eGanzelmeier.ArableCrops
                Else
                    Ganzelmeier = convertFOCUSCrop2Ganzelmeier(FOCUSswDriftCrop)
                End If

            Case Else

                Ganzelmeier = convertFOCUSCrop2Ganzelmeier(FOCUSswDriftCrop)

        End Select

        'distances
        'edge of the field to top of the bank
        If step03stdBuffer Then

            edgeField2topBank =
                    getFOCUSStdDistance(
                                FOCUSswDriftCrop:=FOCUSswDriftCrop,
                                FOCUSswWaterBody:=FOCUSswWaterBody,
                                FOCUSstdDistancesMember:=eFOCUSStdDistancesMember.edgeField2topBank)
        Else
            edgeField2topBank = bufferWidth
        End If

        'top of the bank to edge of water body
        'FOCUS Step03 std. 'buffer' distance
        If step03stdBuffer Then

            topBank2edgeWaterbody =
                    getFOCUSStdDistance(
                                FOCUSswDriftCrop:=FOCUSswDriftCrop,
                                FOCUSswWaterBody:=FOCUSswWaterBody,
                                FOCUSstdDistancesMember:=eFOCUSStdDistancesMember.topBank2edgeWaterbody)

        Else
            topBank2edgeWaterbody = bufferWidth
        End If

        'distance limit for each regression in m, hinge point
        hingePoint = hingePointDB(Ganzelmeier, noOfApplns)


        'distance from crop at edge nearest to field
        If step03stdBuffer Then
            closest2EdgeOfField = edgeField2topBank + topBank2edgeWaterbody
        Else
            closest2EdgeOfField = bufferWidth
        End If

        'distance from crop at edge farthest from field
        If step03stdBuffer Then

            farthest2EdgeOfField =
                    edgeField2topBank + topBank2edgeWaterbody + WaterBodyDB(FOCUSswWaterBody, eWBMember.Witdth)

        Else

            farthest2EdgeOfField =
                    bufferWidth + WaterBodyDB(FOCUSswWaterBody, eWBMember.Witdth)

        End If

        'get regression parameter
        A = regressionA(Ganzelmeier, noOfApplns)

        B = regressionB(Ganzelmeier, noOfApplns)

        C = regressionC(Ganzelmeier, noOfApplns)

        D = regressionD(Ganzelmeier, noOfApplns)

        'drift distances farthest, closest or average
        Select Case driftDistance

            Case eDriftDistance.farthest

                Try

                    If farthest2EdgeOfField < hingePoint Then
                        driftPercent = A * (farthest2EdgeOfField ^ B)
                    Else
                        driftPercent = C * (farthest2EdgeOfField ^ D)
                    End If

                Catch ex As Exception

                    Throw New _
                            ArithmeticException(
                            message:="Error during drift calc.",
                            innerException:=ex)

                End Try

            Case eDriftDistance.closest

                Try

                    If closest2EdgeOfField < hingePoint Then
                        driftPercent = A * (closest2EdgeOfField ^ B)
                    Else
                        driftPercent = C * (closest2EdgeOfField ^ D)
                    End If

                Catch ex As Exception

                    Throw New _
                            ArithmeticException(
                            message:="Error during drift calc.",
                            innerException:=ex)

                End Try

            Case eDriftDistance.average

                Try

                    If farthest2EdgeOfField < hingePoint Then
                        driftPercent = (A / (B + 1) * (farthest2EdgeOfField ^ (B + 1) - closest2EdgeOfField ^ (B + 1))) /
                                               (farthest2EdgeOfField - closest2EdgeOfField)
                    ElseIf closest2EdgeOfField > hingePoint Then
                        driftPercent = C / (D + 1) * (farthest2EdgeOfField ^ (D + 1) - closest2EdgeOfField ^ (D + 1)) /
                                              (farthest2EdgeOfField - closest2EdgeOfField)
                    Else
                        driftPercent = (A / (B + 1) * (hingePoint ^ (B + 1) - closest2EdgeOfField ^ (B + 1)) + C / (D + 1) * (farthest2EdgeOfField ^ (D + 1) - hingePoint ^ (D + 1))) * 1 /
                                               (farthest2EdgeOfField - closest2EdgeOfField)
                    End If

                Catch ex As Exception

                    Throw New _
                            ArithmeticException(
                            message:="Error during drift calc.",
                            innerException:=ex)

                End Try

        End Select

        ' if stream then apply upstream catchment factor 1.2
        driftPercent *= WaterBodyDB(FOCUSswWaterBody, eWBMember.Factor)

        'drift reducing nozzles?     
        driftPercent *= 1 - nozzle / 100

        Return driftPercent

    End Function


    Public Function calcDriftPercentStep12(
                        noOfApplns As eNoOfApplns,
                        FOCUSswDriftCrop As eFOCUSswDriftCrop,
               Optional bufferWidth As Integer = eBufferWidth.not_defined) As Double

        If noOfApplns = eNoOfApplns.not_defined OrElse
                    FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined Then
            Return Double.NaN

        End If

        Dim buffer As Integer

        Select Case FOCUSswDriftCrop

            Case _
                    eFOCUSswDriftCrop.CI,
                    eFOCUSswDriftCrop.OL,
                    eFOCUSswDriftCrop.VI,
                    eFOCUSswDriftCrop.PFE, eFOCUSswDriftCrop.PFL,
                    eFOCUSswDriftCrop.HP

                buffer = eBufferWidth._03

                'Case _
                '    eFOCUSswDriftCrop.HH

                '    buffer = eBufferWidth._03

                '    FOCUSswDriftCrop = eFOCUSswDriftCrop.OL

                'Case _
                '     eFOCUSswDriftCrop.HL

                '    buffer = eBufferWidth._01

                '    FOCUSswDriftCrop = eFOCUSswDriftCrop.CS

            Case Else
                buffer = eBufferWidth._01

        End Select

        If bufferWidth <> eBufferWidth.not_defined AndAlso
           bufferWidth <> eBufferWidth.FOCUSStep03 Then

            buffer = bufferWidth

        End If

        Dim driftPercent As Double

        driftPercent = calcDriftPercent(
                noOfApplns:=noOfApplns,
                FOCUSswDriftCrop:=FOCUSswDriftCrop,
                FOCUSswWaterBody:=eFOCUSswWaterBody.ditch,
                bufferWidth:=buffer,
                driftDistance:=eDriftDistance.closest)

        Return driftPercent

    End Function

    Public Function calcMaxNozzle(
                            maxReduction As Integer,
                            noOfApplns As eNoOfApplns,
                            FOCUSswDriftCrop As eFOCUSswDriftCrop,
                            FOCUSswWaterBody As eFOCUSswWaterBody,
                            bufferWidth As Integer) As eNozzles

        Dim base As Double = Double.NaN
        Dim nozzleDrift As Double = Double.NaN
        Dim temp As Double = Double.NaN


        Dim nozzles As eNozzles() =
                {eNozzles._0,
                eNozzles._10,
                eNozzles._20,
                eNozzles._25,
                eNozzles._30,
                eNozzles._40,
                eNozzles._50,
                eNozzles._60,
                eNozzles._70,
                eNozzles._75,
                eNozzles._80,
                eNozzles._90,
                eNozzles._95
                }

        base =
                calcDriftPercent(
                noOfApplns:=noOfApplns,
                FOCUSswDriftCrop:=FOCUSswDriftCrop,
                FOCUSswWaterBody:=FOCUSswWaterBody,
                Nozzle:=eNozzles._0,
                bufferWidth:=eBufferWidth.FOCUSStep03)

        For counter As Integer = 1 To nozzles.Count - 1

            nozzleDrift =
                    calcDriftPercent(
                    noOfApplns:=noOfApplns,
                    FOCUSswDriftCrop:=FOCUSswDriftCrop,
                    FOCUSswWaterBody:=FOCUSswWaterBody,
                    Nozzle:=nozzles(counter),
                    bufferWidth:=bufferWidth)


            temp = Math.Round(100 - (nozzleDrift * 100 / base), digits:=0)

            If temp >= maxReduction Then
                Return nozzles(counter - 1)
            End If

        Next

        Return eNozzles._0

    End Function

    Public Function calcMaxBuffer(
                            maxReduction As Integer,
                            noOfApplns As eNoOfApplns,
                            FOCUSswDriftCrop As eFOCUSswDriftCrop,
                            FOCUSswWaterBody As eFOCUSswWaterBody,
                            nozzle As eNozzles) As eBufferWidth

        Dim base As Double = Double.NaN
        Dim bufferDrift As Double = Double.NaN
        Dim temp As Double = Double.NaN


        Dim buffers As eBufferWidth() =
                {eBufferWidth._01,
                eBufferWidth._02,
                eBufferWidth._03,
                eBufferWidth._04,
                eBufferWidth._05,
                eBufferWidth._10,
                eBufferWidth._15,
                eBufferWidth._20,
                eBufferWidth._30,
                eBufferWidth._40,
                eBufferWidth._50
                }

        base =
                calcDriftPercent(
                noOfApplns:=noOfApplns,
                FOCUSswDriftCrop:=FOCUSswDriftCrop,
                FOCUSswWaterBody:=FOCUSswWaterBody,
                Nozzle:=eNozzles._0,
                bufferWidth:=eBufferWidth.FOCUSStep03)

        For counter As Integer = 1 To buffers.Count - 1

            bufferDrift =
                    calcDriftPercent(
                    noOfApplns:=noOfApplns,
                    FOCUSswDriftCrop:=FOCUSswDriftCrop,
                    FOCUSswWaterBody:=FOCUSswWaterBody,
                    Nozzle:=nozzle,
                    bufferWidth:=buffers(counter))


            temp = Math.Round(100 - (bufferDrift * 100 / base), digits:=0)

            If temp >= maxReduction Then
                Return buffers(counter - 1)
            End If

        Next

        Return eBufferWidth.FOCUSStep03

    End Function


    ''' <summary>
    ''' check if buffer is std Step03 or something else
    ''' </summary>
    ''' <param name="bufferWidth"></param>
    ''' <returns></returns>
    <DebuggerStepThrough>
    Public Function checkStep03StdBuffer(bufferWidth As eBufferWidth) As Boolean

        Try

            If [Enum].Parse(
                enumType:=GetType(eBufferWidth),
                value:=bufferWidth) = eBufferWidth.FOCUSStep03 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Return False
        End Try

    End Function



#Region "    Std. FOCUS Step03 buffer distance in m"

    <DebuggerStepThrough()>
    Public Function getFOCUSStdDistance(
                                FOCUSswDriftCrop As eFOCUSswDriftCrop,
                                FOCUSswWaterBody As eFOCUSswWaterBody,
                                FOCUSstdDistancesMember As eFOCUSStdDistancesMember) As Double

        Dim FOCUSstdDistancesEntry As String = ""
        Dim CropString As String = String.Empty

        If FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined OrElse
               FOCUSswWaterBody = eFOCUSswWaterBody.not_defined Then
            Return Double.NaN
        End If

        Try

            CropString =
                enumConverter(Of eFOCUSswDriftCrop).getEnumDescription(
                        EnumConstant:=FOCUSswDriftCrop)

            FOCUSstdDistancesEntry =
                    Array.Find(
                    array:=
                        Filter(
                            Source:=FOCUSStdDistances,
                            Match:=FOCUSswWaterBody.ToString,
                            Include:=True,
                            Compare:=CompareMethod.Text),
                    match:=Function(x) x.StartsWith(CropString.Substring(0, 2)))


            FOCUSstdDistancesEntry =
                    FOCUSstdDistancesEntry.Split({"|"c})(FOCUSstdDistancesMember)

            Return CDbl(FOCUSstdDistancesEntry)

        Catch ex As Exception
            Return Double.NaN
        End Try

    End Function


    Public Enum eFOCUSStdDistancesMember

        cropShort
        cropString
        waterbody
        edgeField2topBank
        topBank2edgeWaterbody

    End Enum


    Public FOCUSStdDistances As String() =
        {
            "CS|cereals, spring|ditch|0.5|0.5",
            "CS|cereals, spring|stream|0.5|1",
            "CS|cereals, spring|pond|0.5|3",
            "CW|cereals, winter|ditch|0.5|0.5",
            "CW|cereals, winter|stream|0.5|1",
            "CW|cereals, winter|pond|0.5|3",
            "GA|grass/alfalfa|ditch|0.5|0.5",
            "GA|grass/alfalfa|stream|0.5|1",
            "GA|grass/alfalfa|pond|0.5|3",
            "OS|oil seed rape, spring|ditch|0.5|0.5",
            "OS|oil seed rape, spring|stream|0.5|1",
            "OS|oil seed rape, spring|pond|0.5|3",
            "OW|oil seed rape, winter|ditch|0.5|0.5",
            "OW|oil seed rape, winter|stream|0.5|1",
            "OW|oil seed rape, winter|pond|0.5|3",
            "VB|vegetables, bulb|ditch|0.5|0.5",
            "VB|vegetables, bulb|stream|0.5|1",
            "VB|vegetables, bulb|pond|0.5|3",
            "VF|vegetables, fruiting|ditch|0.5|0.5",
            "VF|vegetables, fruiting|stream|0.5|1",
            "VF|vegetables, fruiting|pond|0.5|3",
            "VL|vegetables, leafy|ditch|0.5|0.5",
            "VL|vegetables, leafy|stream|0.5|1",
            "VL|vegetables, leafy|pond|0.5|3",
            "VR|vegetables, root|ditch|0.5|0.5",
            "VR|vegetables, root|stream|0.5|1",
            "VR|vegetables, root|pond|0.5|3",
            "PS|potatoes|ditch|0.8|0.5",
            "PS|potatoes|stream|0.8|1",
            "PS|potatoes|pond|0.8|3",
            "SY|soybeans|ditch|0.8|0.5",
            "SY|soybeans|stream|0.8|1",
            "SY|soybeans|pond|0.8|3",
            "SB|sugar beets|ditch|0.8|0.5",
            "SB|sugar beets|stream|0.8|1",
            "SB|sugar beets|pond|0.8|3",
            "SU|sunflowers|ditch|0.8|0.5",
            "SU|sunflowers|stream|0.8|1",
            "SU|sunflowers|pond|0.8|3",
            "CO|cotton|ditch|0.8|0.5",
            "CO|cotton|stream|0.8|1",
            "CO|cotton|pond|0.8|3",
            "FB|field beans|ditch|0.8|0.5",
            "FB|field beans|stream|0.8|1",
            "FB|field beans|pond|0.8|3",
            "LG|legumes|ditch|0.8|0.5",
            "LG|legumes|stream|0.8|1",
            "LG|legumes|pond|0.8|3",
            "MZ|maize|ditch|0.8|0.5",
            "MZ|maize|stream|0.8|1",
            "MZ|maize|pond|0.8|3",
            "TB|tobacco|ditch|1|0.5",
            "TB|tobacco|stream|1|1",
            "TB|tobacco|pond|1|3",
            "CI|citrus|ditch|3|0.5",
            "CI|citrus|stream|3|1",
            "CI|citrus|pond|3|3",
            "HP|hops|ditch|3|0.5",
            "HP|hops|stream|3|1",
            "HP|hops|pond|3|3",
            "OL|olives|ditch|3|0.5",
            "OL|olives|stream|3|1",
            "OL|olives|pond|3|3",
            "PF|pome/stone fruits|ditch|3|0.5",
            "PF|pome/stone fruits|stream|3|1",
            "PF|pome/stone fruits|pond|3|3",
            "VI|vines|ditch|3|0.5",
            "VI|vines|stream|3|1",
            "VI|vines|pond|3|3",
            "VT|vines, late|ditch|3|0.5",
            "VT|vines, late|stream|3|1",
            "VT|vines, late|pond|3|3",
            "HH|appln, hand (crop > 50 cm)|ditch|3|0.5",
            "HH|appln, hand (crop > 50 cm)|stream|3|1",
            "HH|appln, hand (crop > 50 cm)|pond|3|3",
            "HL|appln, hand (crop < 50 cm)|ditch|0.5|0.5",
            "HL|appln, hand (crop < 50 cm)|stream|0.5|1",
            "HL|appln, hand (crop < 50 cm)|pond|0.5|3",
            "AA|appln, aerial|ditch|5|0.5",
            "AA|appln, aerial|stream|5|1",
            "AA|appln, aerial|pond|5|3"
        }

#End Region

#Region "    Water body data"

    Public Enum eWBMember

        Witdth
        Length
        Depth
        DistanceBankWater
        Factor

    End Enum

    Public WaterBodyDB As Double(,) =
        {
            {1, 100, 0.3, 0.5, 1},
            {1, 100, 0.3, 1, 1.2},
            {30, 30, 1, 3, 1}
        }


    Public runOffWaterDepths As Double() = {0.41, 0.305, 0.29, 0.41}


#End Region

#Region "    Regression parameters and hinge point"

    Public Property regressionA As Double(,) =
        {
            {2.7593, 2.4376, 2.0244, 1.8619, 1.7942, 1.6314, 1.5784, 1.5119},
            {66.702, 62.272, 58.796, 58.947, 58.111, 58.829, 59.912, 59.395},
            {60.396, 42.002, 40.12, 36.273, 34.591, 31.64, 31.561, 29.136},
            {58.247, 66.243, 60.397, 58.559, 59.548, 60.136, 59.774, 53.2},
            {15.793, 15.461, 16.887, 16.484, 15.648, 15.119, 14.675, 14.948},
            {44.769, 40.262, 39.314, 37.401, 37.767, 36.908, 35.498, 35.094},
            {50.47, 50.47, 50.47, 50.47, 50.47, 50.47, 50.47, 50.47}
        }

    Public Property regressionB As Double(,) =
        {
            {-0.9778, -1.01, -0.9956, -0.9861, -0.9943, -0.9861, -0.9811, -0.9832},
            {-0.752, -0.8116, -0.8171, -0.8331, -0.8391, -0.8644, -0.8838, -0.8941},
            {-1.2249, -1.1306, -1.1769, -1.1616, -1.1533, -1.1239, -1.1318, -1.1048},
            {-1.0042, -1.2001, -1.2132, -1.2171, -1.2481, -1.2699, -1.2813, -1.2469},
            {-1.608, -1.6599, -1.7223, -1.7172, -1.7072, -1.6999, -1.6936, -1.7177},
            {-1.5643, -1.5771, -1.5842, -1.5746, -1.5829, -1.5905, -1.5844, -1.5819},
            {-0.3819, -0.3819, -0.3819, -0.3819, -0.3819, -0.3819, -0.3819, -0.3819}
        }

    'only used for fruit/early, fruit/late and hops
    Public Property regressionC As Double(,) =
        {
            {2.7593, 2.4376, 2.0244, 1.8619, 1.7942, 1.6314, 1.5784, 1.5119},
            {3867.9, 7961.7, 9598.8, 8609.8, 7684.6, 7065.6, 7292.9, 7750.9},
            {210.7, 298.76, 247.78, 201.98, 197.08, 228.69, 281.84, 256.33},
            {8654.9, 5555.3, 4060.9, 3670.4, 2860.6, 2954, 3191.6, 3010.1},
            {15.793, 15.461, 16.887, 16.484, 15.648, 15.119, 14.675, 14.948},
            {44.769, 40.262, 39.314, 37.401, 37.767, 36.908, 35.498, 35.094},
            {281.06, 281.06, 281.06, 281.06, 281.06, 281.06, 281.06, 281.06}
        }

    'only used for fruit/early, fruit/late and hops
    Public Property regressionD As Double(,) =
        {
            {-0.9778, -1.01, -0.9956, -0.9861, -0.9943, -0.9861, -0.9811, -0.9832},
            {-2.4183, -2.6854, -2.7706, -2.7592, -2.7366, -2.7323, -2.7463, -2.7752},
            {-1.7599, -1.9464, -1.9299, -1.8769, -1.8799, -1.9519, -2.0087, -1.9902},
            {-2.8354, -2.8231, -2.7625, -2.7619, -2.7036, -2.7269, -2.7665, -2.7549},
            {-1.608, -1.6599, -1.7223, -1.7172, -1.7072, -1.6999, -1.6936, -1.7177},
            {-1.5643, -1.5771, -1.5842, -1.5746, -1.5829, -1.5905, -1.5844, -1.5819},
            {-0.9989, -0.9989, -0.9989, -0.9989, -0.9989, -0.9989, -0.9989, -0.9989}
        }

    'distance limit for each regression (m), also called hinge point.
    Public Property hingePointDB As Double(,) =
        {
            {1, 1, 1, 1, 1, 1, 1, 1},
            {11.4, 13.3, 13.6, 13.3, 13.1, 13, 13.2, 13.3},
            {10.3, 11.1, 11.2, 11, 11, 10.9, 12.1, 11.7},
            {15.3, 15.3, 15.1, 14.6, 14.3, 14.5, 14.6, 14.6},
            {1, 1, 1, 1, 1, 1, 1, 1},
            {1, 1, 1, 1, 1, 1, 1, 1},
            {16.2, 16.2, 16.2, 16.2, 16.2, 16.2, 16.2, 16.2}
        }

#End Region


#End Region

    ''' <summary>
    ''' convert FOCUSsw DRIFT crop to Ganzelmeier crop group
    ''' </summary>
    Public convertFOCUSCrop2Ganzelmeier As eGanzelmeier() =
        {
           eGanzelmeier.ArableCrops,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.FruitCrops_Late,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.Hops,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.FruitCrops_Late,
           eGanzelmeier.FruitCrops_Early,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.Vines_Early,
           eGanzelmeier.Vines_Late,
           eGanzelmeier.AerialAppln,
           eGanzelmeier.ArableCrops,
           eGanzelmeier.Vines_Late,
           eGanzelmeier.noDrift,
           eGanzelmeier.not_defined
        }


End Module