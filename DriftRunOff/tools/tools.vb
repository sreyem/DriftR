

Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Design
Imports System.Globalization
Imports System.Reflection
Imports System.Text
Imports System.Web.Script.Serialization
Imports System.Windows.Forms
Imports System.Windows.Forms.Design
Imports System.Xml.Serialization

''' <summary>
''' TypeConverter(GetType(propGridConverter))
''' Make a class brows-able in the property grid
''' Name and collapseStd property should be defined!
''' </summary>
Public Class PropGridConverter

    'usage : <TypeConverter(GetType(propGridConverter))>

    Inherits ExpandableObjectConverter

    Public Const noName As String = " ... "

#Region "    Engine"

    <DebuggerStepThrough>
    Public Overloads Overrides Function CanConvertTo(
                                                    ByVal context As ITypeDescriptorContext,
                                                    ByVal destinationType As Type) As Boolean
        Try

            If (destinationType Is GetType(PropGridConverter)) Then
                Return True
            End If

        Catch ex As Exception

        End Try

        Return MyBase.CanConvertTo(
                                context,
                                destinationType)

    End Function

    <DebuggerStepThrough>
    Public Overloads Overrides Function ConvertTo(
                                     ByVal context As ITypeDescriptorContext,
                                     ByVal culture As Globalization.CultureInfo,
                                     ByVal value As Object,
                                     ByVal destinationType As System.Type) As Object

        If (destinationType Is GetType(System.String)) Then

            Try

                Return CallByName(
                                        value,
                                        "Name",
                                        CallType.Get)

            Catch ex As Exception
                Return " ... "
            End Try

        End If

        Return MyBase.ConvertTo(
                                    context,
                                    culture,
                                    value,
                                    destinationType)

    End Function

#End Region

End Class


''' <summary>
''' adds a drop down field
''' </summary>
Public Class dropDownList

    Inherits StringConverter

    ' usage :  add 
    ' <TypeConverter(GetType(dropDownList))>
    ' to property
    Public Overloads Shared Property dropDownEntries As String() = {"add", "own", "entries"}

#Region "Overloads Overrides"

    Public Overloads Overrides Function GetStandardValuesSupported(ByVal context As ITypeDescriptorContext) As Boolean
        Return True
    End Function

    Public Overloads Overrides Function GetStandardValues(ByVal context As ITypeDescriptorContext) As StandardValuesCollection
        Return New StandardValuesCollection(dropDownEntries)
    End Function

    Public Overloads Overrides Function GetStandardValuesExclusive(ByVal context As ITypeDescriptorContext) As Boolean
        Return False
    End Function

#End Region

End Class



''' <summary>
''' enum, dbl, date, int
''' </summary>
#Region "    converter"

Public Module pGridSettings

#Region "    settings"
    Public Enum eCountry
        de_DE
        en_US
        fr_FR
    End Enum

    Public Enum eJulian
        add
        only
        none
    End Enum

    Public Const stdCountry As eCountry = eCountry.en_US
    Public Const stdEmptyString As String = " - "

    Public Const stdDblformat As String = "G4"
    Public Const stdMinSign As String = "<"
    Public Const stdMinValue As Double = Double.NaN
    Public Const stdDigits As Integer = -1
    Public Const stdUnit As String = ""
    Public Const stdnegDef As Boolean = False

    Public Const stdDateFormat As String = "dd. MMM"
    Public Const stdJulian As eJulian = eJulian.add

    Public Const stdDblEmpty As Double = 0
    Public Const zeroIsEmpty As Boolean = False

#End Region

#Region "    functions"

    Public Function Conv2String(
                             value As Double,
                    Optional format As String = stdDblformat,
                    Optional country As eCountry = stdCountry,
                    Optional minValue As Double = stdMinValue,
                    Optional minSign As String = stdMinSign,
                    Optional digits As Integer = stdDigits,
                    Optional unit As String = stdUnit,
                    Optional stdDblEmpty As Double = stdDblEmpty,
                    Optional zeroIsEmpty As Boolean = True) As String

        Dim countryString As String

        countryString =
                Replace(
                Expression:=country.ToString,
                Find:="_", Replacement:="-",
                Compare:=CompareMethod.Text)

        'if minvalue is set and value < minvalue
        If Not Double.IsNaN(minValue) AndAlso
                    value < minValue Then

            Return minSign & minValue.ToString & unit

        End If

        If digits > -1 Then

            value =
                Math.Round(
                    value,
                    digits:=digits)

        End If

        If value = stdDblEmpty Then

            If value = 0 Then

                If zeroIsEmpty Then
                    Return stdEmptyString
                Else
                    Return value.ToString(
                          format:=format,
                        provider:=CultureInfo.CreateSpecificCulture(countryString)) & unit
                End If

            Else
                Return stdEmptyString
            End If

        End If


        Return value.ToString(
                          format:=format,
                        provider:=CultureInfo.CreateSpecificCulture(countryString)) & unit

    End Function

    Public Function CleanInputString(
                   value As Object, unit As String, minSign As String) As String

        Try

            value = CType(value, String)

            value =
            Replace(
            Expression:=value,
            Find:=unit,
            Replacement:="",
            Compare:=CompareMethod.Text)

            value =
            Replace(
            Expression:=value,
            Find:=stdEmptyString,
            Replacement:="",
            Compare:=CompareMethod.Text)

            value =
            Replace(
            Expression:=value,
            Find:=minSign,
            Replacement:="",
            Compare:=CompareMethod.Text)

            value = Trim(value)

            Return value

        Catch ex As Exception
            Return stdEmptyString
        End Try

    End Function

    Public Function GetNumberParameter(context As ITypeDescriptorContext) As String()

        Dim attributes As AttributeCollection
        Dim provider As AttributeProviderAttribute

        Dim formatArray As String() = {}

        Try

            With context

                attributes = .PropertyDescriptor.Attributes

                For Each attribute In attributes

                    If attribute.ToString =
                        "System.ComponentModel.AttributeProviderAttribute" Then

                        provider = attribute
                        formatArray = provider.TypeName.Split("|")

                    End If

                Next

            End With
        Catch ex As Exception
            Console.WriteLine(
                "Error getting attributes from attribute provider")
        End Try

        For counter As Integer = 0 To formatArray.Count - 1
            formatArray(counter) = Trim(formatArray(counter))
        Next

        Return formatArray

    End Function

#End Region

End Module


''' <summary>
''' Show enum descriptions
''' TypeConverter(GetType(EnumConverter(of enumType))
''' </summary>
''' <typeparam name="T">
''' enum type
''' </typeparam>
Public Class enumConverter(Of T)

    'usage :  <TypeConverter(GetType(EnumConverter(of <Type>))>

    Inherits EnumConverter

    Public Const not_defined As String = " - "

    '' <summary>
    '' Initializing instance
    '' </summary>       
    '' <remarks></remarks>
    Public Sub New()
        MyBase.New(GetType(T))
    End Sub

    ''' <summary>
    ''' don't show enum members with this description
    ''' </summary>
    ''' <returns></returns>
    Public Shared Property dontShow As String() = {}

    ''' <summary>
    ''' only show enum members with this description
    ''' </summary>
    ''' <returns></returns>
    Public Shared Property onlyShow As String() = {}

    ''' <summary>
    ''' check chosen value for being valid
    ''' </summary>
    ''' <param name="value">
    ''' chosen value as enum
    ''' </param>
    ''' <returns></returns>
    Public Shared Function checkEntry(value As Object) As Boolean

        Dim entry As String

        Try
            entry = enumConverter(Of T).getEnumDescription(value)
        Catch ex As Exception
            Return False
        End Try


        If onlyShow.Contains(value) AndAlso
   Not dontShow.Contains(value) Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Shared no2Show As Integer = -1

#Region "    Engine"

    Public Overrides Function CanConvertTo(context As ITypeDescriptorContext, destType As Type) As Boolean
        Return destType = GetType(String)
    End Function

    Public Overrides Function ConvertTo(
                                context As ITypeDescriptorContext,
                                culture As CultureInfo,
                                value As Object,
                                destType As Type) As Object

        Dim fi As FieldInfo = GetType(T).GetField([Enum].GetName(GetType(T), value))
        Dim dna As DescriptionAttribute = DirectCast(
        Attribute.GetCustomAttribute(fi, GetType(DescriptionAttribute)), DescriptionAttribute)

        If dna IsNot Nothing Then

            If no2Show > -1 AndAlso dna.Description.Split(vbLf).Count > no2Show Then
                Return dna.Description.Split(vbLf)(no2Show).Split(vbCr).First
            Else
                Return dna.Description
            End If

        Else
            Return value.ToString()
        End If

    End Function

    Public Overrides Function CanConvertFrom(
                        context As ITypeDescriptorContext,
                        srcType As Type) As Boolean
        Return srcType = GetType(String)
    End Function

    Public Overrides Function ConvertFrom(
                                context As ITypeDescriptorContext,
                                culture As CultureInfo,
                                value As Object) As Object

        Dim found As Boolean = False

        For Each member As String In dontShow

            If CStr(value).Contains(member) Then
                value = not_defined
            End If

        Next


        If onlyShow.Count <> 0 Then

            For Each member As String In onlyShow
                If CStr(value).Contains(member) Then
                    found = True
                    Exit For
                End If
            Next

            If Not found Then
                value = not_defined
            End If

        End If


        For Each fi As FieldInfo In GetType(T).GetFields()

            Dim dna As DescriptionAttribute =
            DirectCast(
                    Attribute.GetCustomAttribute(
                            element:=fi,
                            attributeType:=GetType(DescriptionAttribute)),
                    DescriptionAttribute)

            If (dna IsNot Nothing) AndAlso
                (dna.Description).Contains(DirectCast(value, String)) Then
                '(DirectCast(value, String).Contains(dna.Description)) Then
                Return [Enum].Parse(GetType(T), fi.Name)
            End If

        Next

        Return [Enum].Parse(GetType(T), DirectCast(value, String))

    End Function

#End Region

    ''' <summary>
    ''' returns the enum description
    ''' </summary>
    ''' <param name="EnumConstant"></param>
    ''' <returns></returns>
    <DebuggerStepThrough>
    Public Shared Function getEnumDescription(ByVal EnumConstant As [Enum]) As String

        Dim fi As FieldInfo =
        EnumConstant.GetType().GetField(EnumConstant.ToString())

        Dim attr() As DescriptionAttribute =
                  DirectCast(fi.GetCustomAttributes(GetType(DescriptionAttribute),
                  False), DescriptionAttribute())

        If attr.Length > 0 Then
            Return attr(0).Description
        Else
            Return EnumConstant.ToString()
        End If

    End Function

End Class


''' <summary>
''' double to string converter
''' AttributeProvider(
''' format , minValue , unit , country , minSign , digits , negDef  
''' </summary>
Public Class DblConv

    Inherits DoubleConverter

    'usage : 
    '<TypeConverter(GetType(dblConv))>
    '<AttributeProvider("format= 'G5'|unit=' kg/ha'")>

#Region "    set parameters"

    Public Sub setNumberParameters(formatArray As String())

        Dim parameterValue As String()

        setStd()

        For Each member As String In formatArray

            parameterValue = member.Split("=")
            parameterValue(0) = Trim(parameterValue(0))
            parameterValue(1) = parameterValue(1).Split("'")(1)

            Select Case parameterValue.First.ToLower

                Case "format"
                    dblformat = parameterValue.Last

                Case "country"

                    Try
                        country =
                        [Enum].Parse(
                        enumType:=GetType(eCountry),
                        value:=parameterValue.Last)
                    Catch ex As Exception
                        country = stdCountry
                    End Try

                Case "minsign"
                    minSign = parameterValue.Last

                Case "minvalue"
                    minValue = parameterValue.Last

                Case "digits"
                    digits = parameterValue.Last

                Case "negdef"
                    negDef = parameterValue.Last

                Case "unit"
                    unit = parameterValue.Last

            End Select

        Next

    End Sub

    Public Sub setStd()

        country = stdCountry
        dblformat = stdDblformat
        emptyString = stdEmptyString
        minSign = stdMinSign
        minValue = stdMinValue
        digits = stdDigits
        unit = ""
        negDef = stdnegDef

    End Sub

#End Region

#Region "    engine"

    ''' <summary>
    ''' out
    ''' </summary>
    ''' <param name="context"></param>
    ''' <param name="culture"></param>
    ''' <param name="value"></param>
    ''' <param name="destinationType"></param>
    ''' <returns></returns>
    <DebuggerStepThrough>
    Public Overrides Function ConvertTo(
                        context As ITypeDescriptorContext,
                        culture As CultureInfo,
                        value As Object,
                        destinationType As Type) As Object

        setNumberParameters(
        formatArray:=GetNumberParameter(context:=context))

        Try

            Return _
          Conv2String(
                value:=value,
                format:=dblformat,
                country:=country,
                minSign:=minSign,
                minValue:=minValue,
                unit:=unit,
                digits:=digits)

        Catch ex As Exception
            Return value
        End Try

    End Function

    ''' <summary>
    ''' convert to double if possible
    ''' </summary>
    ''' <param name="context"></param>
    ''' <param name="culture"></param>
    ''' <param name="value"></param>
    ''' <returns></returns>
    <DebuggerStepThrough>
    Public Overrides Function ConvertFrom(
                        context As ITypeDescriptorContext,
                        culture As CultureInfo,
                        value As Object) As Object

        Dim valueString As String

        Try

            valueString =
            CleanInputString(
                value:=value,
                unit:=Me.unit,
                minSign:=Me.minSign)

            If Double.IsNaN(Double.Parse(valueString)) Then
                Return Double.NaN
            Else
                Return Double.Parse(valueString)
            End If

        Catch ex As Exception
            Return Double.NaN
        End Try

    End Function

#End Region

#Region "    definitions"

    Public country As eCountry = stdCountry
    Public dblformat As String = stdDblformat
    Public emptyString As String = stdEmptyString
    Public minSign As String = stdMinSign
    Public minValue As Double = stdMinValue
    Public digits As Integer = stdDigits
    Public unit As String = ""
    Public negDef As Boolean = stdnegDef

#End Region

End Class

''' <summary>
''' integer to string converter
''' </summary>
Public Class IntConv

    Inherits Int16Converter

    'usage : 
    '<TypeConverter(GetType(intConv))>
    '<AttributeProvider("format= 'G5'|unit=' kg/ha'")>

#Region "    set parameters"

    Public Sub SetStd()

        country = stdCountry
        dblformat = stdDblformat
        emptyString = stdEmptyString
        minSign = stdMinSign
        minValue = stdMinValue
        unit = ""
        negDef = stdnegDef

    End Sub


    Public Sub SetNumberParameters(formatArray As String())

        Dim parameterValue As String()

        SetStd()

        For Each member As String In formatArray

            parameterValue = member.Split("=")
            parameterValue(0) = Trim(parameterValue(0))
            parameterValue(1) = parameterValue(1).Split("'")(1)

            Select Case parameterValue.First.ToLower

                Case "format"
                    dblformat = parameterValue.Last

                Case "country"

                    Try
                        country =
                        [Enum].Parse(
                        enumType:=GetType(eCountry),
                        value:=parameterValue.Last)
                    Catch ex As Exception
                        country = stdCountry
                    End Try

                Case "minsign"
                    minSign = parameterValue.Last

                Case "minvalue"
                    minValue = parameterValue.Last

                    'Case "digits"
                    '    digits = parameterValue.Last

                Case "negdef"
                    negDef = parameterValue.Last

                Case "unit"
                    unit = parameterValue.Last

            End Select

        Next

    End Sub

#End Region

#Region "    engine"

    <DebuggerStepThrough>
    Public Overrides Function ConvertTo(
                    context As ITypeDescriptorContext,
                    culture As CultureInfo,
                    value As Object,
                    destinationType As Type) As Object


        SetNumberParameters(formatArray:=GetNumberParameter(context:=context))

        Try

            Return _
      Conv2String(
            value:=value,
            format:=dblformat,
            country:=country,
            minSign:=minSign,
            minValue:=minValue,
            unit:=unit)

        Catch ex As Exception
            Return value
        End Try

    End Function

    <DebuggerStepThrough>
    Public Overrides Function ConvertFrom(
                    context As ITypeDescriptorContext,
                    culture As CultureInfo,
                    value As Object) As Object

        Dim valueString As String

        Try

            valueString =
            CleanInputString(
                value:=value,
                unit:=Me.unit,
                minSign:=Me.minSign)

            If Double.IsNaN(Double.Parse(valueString)) Then
                Return Integer.MaxValue
            Else
                Return Integer.Parse(valueString)
            End If

        Catch ex As Exception
            Return Integer.MaxValue
        End Try

    End Function

#End Region

#Region "    definitions"

    Public country As eCountry = stdCountry
    Public dblformat As String = stdDblformat
    Public emptyString As String = stdEmptyString
    Public minSign As String = stdMinSign
    Public minValue As Double = stdMinValue
    Public digits As Integer = stdDigits
    Public unit As String = ""
    Public negDef As Boolean = stdnegDef

#End Region

End Class

''' <summary>
''' date to string converter
''' AttributeProvider(
''' format, julian = add/only/none, country
''' </summary>
Public Class DateConv

    Inherits DateTimeConverter

    '<TypeConverter(GetType(dateConv)>

#Region "    set parameters"

    Public Sub setStd()

        country = stdCountry
        dateFormat = stdDateFormat
        emptyString = stdEmptyString

    End Sub


    Public Sub SetDateParameters(formatArray As String())

        Dim parameterValue As String()

        setStd()

        For Each member As String In formatArray

            parameterValue = member.Split("=")
            parameterValue(0) = Trim(parameterValue(0))
            parameterValue(1) = parameterValue(1).Split("'")(1)

            Select Case parameterValue.First

                Case "format"
                    dateFormat = parameterValue.Last

                Case "country"

                    Try
                        country =
                        [Enum].Parse(
                        enumType:=GetType(eCountry),
                        value:=parameterValue.Last)
                    Catch ex As Exception
                        country = stdCountry
                    End Try

                Case "julian"

                    Try

                        julian =
                        [Enum].Parse(
                        enumType:=GetType(eJulian),
                        value:=parameterValue.Last)

                    Catch ex As Exception
                        julian = stdJulian
                    End Try

            End Select

        Next

    End Sub

#End Region

#Region "    engine"

    ''' <summary>
    ''' out : convert date to string
    ''' </summary>
    ''' <param name="context"></param>
    ''' <param name="culture"></param>
    ''' <param name="value"></param>
    ''' <param name="destType"></param>
    ''' <returns></returns>
    <DebuggerStepThrough>
    Public Overrides Function ConvertTo(
                        context As ITypeDescriptorContext,
                        culture As CultureInfo,
                        value As Object,
                        destType As Type) As Object

        SetDateParameters(
        formatArray:=GetNumberParameter(
                        context:=context))

        Try

            If CType(value, Date) = Date.MinValue OrElse
           CType(value, Date) = New Date Then

                Return emptyString

            Else

                Return _
            ConvDate2String(
                    value:=value,
                    format:=dateFormat,
                    julian:=julian,
                    emptyString:=emptyString,
                    country:=country)

            End If
        Catch ex As Exception
            Return emptyString
        End Try



    End Function

    ''' <summary>
    ''' in  : convert string to date
    ''' </summary>
    ''' <param name="context"></param>
    ''' <param name="culture"></param>
    ''' <param name="value"></param>
    ''' <returns></returns>
    <DebuggerStepThrough>
    Public Overrides Function ConvertFrom(
                        context As ITypeDescriptorContext,
                        culture As CultureInfo,
                        value As Object) As Object

        Try

            If CType(value, Date) = New Date Then
                Return New Date
            Else
                Return Date.Parse(CType(value, String))
            End If

        Catch ex As Exception
            Return New Date
        End Try

    End Function

#End Region

#Region "    functions"

    ''' <summary>
    ''' converts a date to a string
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="format"></param>
    ''' <param name="julian"></param>
    ''' <param name="emptyString"></param>
    ''' <param name="country"></param>
    ''' <returns></returns>
    Public Shared Function ConvDate2String(
                                   value As Date,
                          Optional format As String = stdDateFormat,
                          Optional julian As eJulian = stdJulian,
                          Optional emptyString As String = stdEmptyString,
                          Optional country As eCountry = stdCountry) As String

        Dim countryString As String
        Dim out As String = ""

        countryString =
                Replace(
                Expression:=country.ToString,
                Find:="_", Replacement:="-",
                Compare:=CompareMethod.Text)
        Try

            If value = New Date OrElse IsNothing(value) Then
                Return emptyString
            End If

            out = value.ToString(
                            format:=format,
                            provider:=CultureInfo.CreateSpecificCulture(countryString))

            Select Case julian

                Case eJulian.none

                Case eJulian.add
                    out &= " (" & value.DayOfYear.ToString.PadLeft(3) & ")"

                Case eJulian.only
                    out = value.DayOfYear.ToString()

            End Select

            Return out

        Catch ex As Exception
            Return emptyString
        End Try


    End Function

#End Region

#Region "    definitions"

    Public country As eCountry = stdCountry
    Public julian As eCountry = stdJulian
    Public dateFormat As String = stdDateFormat
    Public emptyString As String = stdEmptyString

#End Region


End Class

#End Region


''' <summary>
''' on click enabler
''' [Editor(GetType(buttonEmulator), GetType(UITypeEditor))]
''' </summary>
Public Class buttonEmulator

    Inherits UITypeEditor

    'use : <Editor(GetType(buttonEmulator), GetType(UITypeEditor))>
    Public Shared Clicked As String = "Clicked"

#Region "Overloads Overrides"

    Public Overrides Function GetEditStyle(context As ITypeDescriptorContext) As UITypeEditorEditStyle
        Return UITypeEditorEditStyle.Modal
    End Function

    Public Overrides Function EditValue(context As ITypeDescriptorContext,
                                        provider As IServiceProvider,
                                        value As Object) As Object

        value = Clicked

        Return MyBase.EditValue(context,
                                    provider,
                                    value)

    End Function

#End Region

End Class



''' <summary>
''' Multi line editor for string properties
''' change FontName and FontSize
''' used by adding
''' [Editor(GetType(multiLineDoubleArrayEditor), GetType(UITypeEditor))]
''' </summary>
''' <remarks></remarks>
Public Class multiLineDoubleArrayEditor

    Inherits UITypeEditor

    ' usage :  add 
    '<Editor(GetType(multiLineDoubleArrayEditor), GetType(UITypeEditor))>
    ' to property

    Public Shared FontName As String = "Courier New"
    Public Shared FontSize As Integer = 10
    Public Shared Width As Integer = 250
    Public Shared Height As Integer = 150
    Public Shared AcceptsReturn As Boolean = True


    Private editorService As IWindowsFormsEditorService

    Public Overrides Function GetEditStyle(
                        context As ITypeDescriptorContext) As UITypeEditorEditStyle
        Return UITypeEditorEditStyle.Modal
    End Function

    Public Overrides Function EditValue(
                        context As ITypeDescriptorContext,
                        provider As IServiceProvider,
                            value As Object) As Object

        editorService = DirectCast(
        provider.GetService(GetType(IWindowsFormsEditorService)),
        IWindowsFormsEditorService)

        Dim txtBox As New TextBox()
        Dim temp As New List(Of String)
        Dim out As New List(Of Double)

        With txtBox

            .Multiline = True

            .Font = New Font(
            familyName:=FontName,
                emSize:=FontSize)

            .ScrollBars = ScrollBars.Both
            .BorderStyle = BorderStyle.None

            .Width = Width
            .Height = Height

            .AcceptsReturn = AcceptsReturn

            For Each member As Double In value

                Try
                    temp.Add(member.ToString)
                Catch ex As Exception
                    temp.Add("parsing error")
                End Try
            Next

            txtBox.Text = Join(SourceArray:=temp.ToArray, Delimiter:=vbCrLf)

        End With

        editorService.DropDownControl(txtBox)

        temp.Clear()
        temp.AddRange(txtBox.Text.Split(CChar(vbCrLf)))

        For Each member As String In temp

            Try
                out.Add(Double.Parse(Trim(member)))
            Catch ex As Exception
                out.Add(Double.NaN)
            End Try

        Next

        Return out.ToArray

    End Function

End Class


Public Class StatsPEC

    Inherits Statistic

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const CATMetaData As String = "00 Meta data"

#Region "    MetaData"

    ''' <summary>
    ''' name to display and for .tostring
    ''' </summary>
    ''' <returns></returns>
    <DisplayName(
    "Name")>
    <Browsable(True)>
    <[ReadOnly](False)>
    <Category(CATMetaData)>
    <RefreshProperties(RefreshProperties.All)>
    <XmlIgnore> <ScriptIgnore>
    Public ReadOnly Overrides Property name As String
        Get

            Dim temp As New StringBuilder

            If Double.IsNaN(Me.Percentile) Then
                Return " - "
            Else

                If Not Double.IsNaN(Me.PEC) Then
                    temp.Append(
                            Conv2String(
                            value:=Me.PEC,
                            unit:=" µg/L",
                            format:="G4"))
                End If

                Return temp.ToString
            End If

        End Get
    End Property



    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _percentileFOCUS As Double = 80

    <DisplayName(
    "FOCUS Percentile")>
    <Description(
    "Percentile for the FOCUS result, 5 - 95" & vbCrLf &
    "-1 for last")>
    <Browsable(True)>
    <[ReadOnly](False)>
    <Category(CATMetaData)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    Public Property PercentileFOCUS As Double
        Get
            Return _percentileFOCUS
        End Get
        Set

            If Value = -1 Then
                _percentileFOCUS = Value
            Else
                _percentileFOCUS = Math.Round(Value / 5, 0) * 5

                Lower = CInt(_percentileFOCUS / 5)
                Upper = Lower + 1

            End If

        End Set
    End Property

    <DisplayName(
    "Upper")>
    <Description(
    "Upper member for FOCUS Percentile")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATMetaData)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    Public Property Lower As Integer = 16

    <DisplayName(
    "Lower")>
    <Description(
    "Lower member for FOCUS Percentile")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATMetaData)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    Public Property Upper As Integer = 17

    <DisplayName("PEC")>
    <Description("Mean of upper and lower member, in µg/L")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATMetaData)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider("format= 'G3'|unit=' µg/L'")>
    <XmlIgnore> <ScriptIgnore>
    Public ReadOnly Property PEC As Double
        Get

            If Not IsNothing(Me.SortedData) Then

                If _percentileFOCUS = -1 Then

                    Return Me.SortedData.Last

                ElseIf SortedData.Count >= Lower AndAlso
                       SortedData.Count >= Upper Then

                    Return (Me.SortedData(Lower - 1) + Me.SortedData(Upper - 1)) / 2

                Else
                    Return Double.NaN
                End If

            Else
                Return Double.NaN
            End If

        End Get
    End Property

#End Region

End Class


''' <summary>
''' Descriptive statistic
''' </summary>
<TypeConverter(GetType(PropGridConverter))>
<DefaultProperty("dataImporter")>
<Serializable>
Public Class Statistic

#Region "    Constructors"

    ''' <summary>
    ''' Default constructor
    ''' </summary>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' Complete constructor
    ''' </summary>
    ''' <param name="Data">
    ''' Array of double to analyze
    ''' </param>
    ''' <param name="Percentile">
    ''' Percentile, from 1 to 100
    ''' std. = 80
    ''' </param>
    ''' <param name="Base">
    ''' std deviation and variance
    ''' from sample or totality
    ''' std. sample
    ''' </param>
    ''' <remarks></remarks>
    Public Sub New(
                Data As Double(),
                Optional Percentile As Integer = 80,
                Optional Base As eSampleTotality = eSampleTotality.Sample,
                Optional gaussNorm As Boolean = True)


        Me.Data = Data
        Me.Percent = Percentile
        Me.Base = Base
        Me.Norm = gaussNorm

        Analyze()

    End Sub

#End Region

#Region "    Calculations"

    ''' <summary>
    ''' Run the analysis to obtain descriptive information of the data
    ''' </summary>
    Public Function Analyze() As Statistic

        Dim ZeroOrNeg As Boolean = False
        Dim SkewHelper As Double = 0
        Dim KurtHelper As Double = 0

        If Data.Count = 0 Then Return New Statistic

        Me.SortedData = New Double(Data.Length - 1) {}
        Me.Gauss = New Double(Data.Length - 1) {}
        Me.Order = New Integer(Data.Length - 1) {}

#Region "        sort data, get order"

        For counter As Integer = 0 To Order.Count - 1
            Me.Order(counter) = counter + 1
            Me.SortedData(counter) = Me.Data(counter)
        Next

        Dim temp As Double
        Dim tempOrder As Short

        For outerCounter As Integer = 0 To Order.Count - 1

            For innerCounter As Integer = 0 To Order.Count - 1

                If SortedData(innerCounter) > SortedData(outerCounter) Then

                    temp = SortedData(innerCounter)

                    SortedData(innerCounter) = SortedData(outerCounter)
                    SortedData(outerCounter) = temp

                    tempOrder = Order(innerCounter)

                    Order(innerCounter) = Order(outerCounter)
                    Order(outerCounter) = tempOrder

                End If

            Next

        Next

        'Data.CopyTo(sortedData, 0)
        'Array.Sort(sortedData)

#End Region

        Me.N = Data.Count
        Me.Min = SortedData.First
        Me.Max = SortedData.Last
        Me.Range = Data.Max - Data.Min

        'sum, .. of products, ... of reciprocals
        For i As Integer = 0 To Data.Length - 1

            If i = 0 Then

                'init
                Me.Sum = Data(i)
                Me.SumOfProducts = Data(i)
                Me.SumOfReciprocalProducts = 1.0 / Data(i)
                'Me.Arithmetic = data(i)

            Else

                Me.Sum += Data(i)
                Me.SumOfProducts *= Data(i)
                Me.SumOfReciprocalProducts += 1.0 / Data(i)

            End If

            'data  <= 0  no geometric and harmonic mean
            If Data(i) <= 0 Then ZeroOrNeg = True

        Next

        Me.Arithmetic = Me.Sum / N

        'data  <= 0  no geometric and harmonic mean
        If ZeroOrNeg Then

            Me.Geometric = Double.NaN
            Me.Harmonic = Double.NaN

        Else

            'https://support.office.com/en-us/article/geomean-function-045ff75c-0bb8-4748-831d-16c395804586?redirectSourcePath=%252fde-de%252farticle%252fGEOMITTEL-Funktion-d97b6117-3280-475d-ac88-7c7858adf58e
            Me.Geometric = Math.Pow(Me.SumOfProducts, 1.0 / N)

            'https://support.office.com/en-us/article/harmean-function-5efd9184-fab5-42f9-b1d3-57883a1d3bc6?redirectSourcePath=%252fde-de%252farticle%252fHARMITTEL-Funktion-b2d540ce-445f-4d31-a09d-021ca38fb416
            Me.Harmonic = 1.0 / (Me.SumOfReciprocalProducts / N)

        End If

        'sum of errors and error square, helper for kurtosis
        For i As Integer = 0 To Data.Length - 1

            If i = 0 Then

                Me.SumOfError = Math.Abs(Data(i) - Me.Arithmetic)
                Me.SumOfErrorSquare = (Math.Abs(Data(i) - Me.Arithmetic)) ^ 2

                KurtHelper = (Math.Abs(Data(i) - Me.Arithmetic)) ^ 4

            Else

                Me.SumOfError += Math.Abs(Data(i) - Me.Arithmetic)
                Me.SumOfErrorSquare += (Math.Abs(Data(i) - Me.Arithmetic)) ^ 2

                KurtHelper += (Math.Abs(Data(i) - Me.Arithmetic)) ^ 4

            End If

        Next

        'https://support.office.com/de-de/article/VAR-S-Funktion-913633DE-136B-449D-813E-65A00B2B990B
        Me.Variance = Me.SumOfErrorSquare / (N - IIf(Me.Base = eSampleTotality.Sample, 1, 0))

        'https://support.office.com/en-us/article/STDEV-S-function-7D69CF97-0C1F-4ACF-BE27-F3E83904CC23
        Me.StdDeviation = Math.Sqrt(Me.Variance)

        'http://jumbo.uni-muenster.de/index.php?id=185
#Region "        Gauss norm. distribution"

        For i As Integer = 0 To Gauss.Length - 1

            Try
                Gauss(i) =
                  Double.Parse(getGauss(
                    value:=Data(i),
                    stdDeviation:=Me.StdDeviation,
                    arithmeticMean:=Me.Arithmetic))

            Catch ex As Exception

            End Try

        Next

        If Norm Then

            Dim maxgauss As Double = Gauss.Max

            For i As Integer = 0 To Gauss.Length - 1
                Gauss(i) = Double.Parse((Gauss(i) * 100 / maxgauss))
            Next

        End If

#End Region

        'https://welt-der-bwl.de/Variationskoeffizient
        Me.RelativeVariance = Me.StdDeviation / Me.Arithmetic * 100


        ' the cum part of SKEW formula
        For i As Integer = 0 To Data.Length - 1
            SkewHelper += Math.Pow((Data(i) - Me.Arithmetic) / (Math.Sqrt(Me.SumOfErrorSquare / (N - 1))), 3)
        Next

        'https://support.office.com/en-us/article/skew-function-bdf49d86-b1ef-4804-a046-28eaea69c9fa
        Me.Skewness = N / (N - 1) / (N - 2) * SkewHelper

        'https://support.office.com/de-de/article/kurt-funktion-bc3a265c-5da4-4dcb-b7fd-c237789095ab
        Me.Kurtosis = ((N + 1) * N * (N - 1)) /
                            ((N - 2) * (N - 3)) * (KurtHelper /
                                               Me.SumOfErrorSquare ^ 2) - 3 * Math.Pow(N - 1, 2) /
                                      ((N - 2) * (N - 3))

        Me.FirstQuartile = calcPercentile(SortedData, 25)
        Me.ThirdQuartile = calcPercentile(SortedData, 75)
        Me.IQR = Me.ThirdQuartile - Me.FirstQuartile
        Me.Median = calcPercentile(SortedData, 50)
        Me.Percentile = calcPercentile(SortedData, Me.Percent)

        Return Me

    End Function

#Region "     Percentile"

    ''' <summary>
    ''' Percentile
    ''' </summary>
    ''' <param name="percent">Percentile, between 0 to 100</param>
    ''' <returns>Percentile</returns>
    Public Function calcPercentile(percent As Double) As Double
        Return Me.calcPercentile(SortedData, percent)
    End Function


    ''' <summary>
    ''' Calculate percentile of a sorted data set
    ''' </summary>
    ''' <param name="sortedData"></param>
    ''' <param name="p">Percentile as percent, 1 - 100</param>
    ''' <returns></returns>
    Public Function calcPercentile(sortedData As Double(), p As Double) As Double

        If IsNothing(sortedData) Then Return 0

        Try

            ' algorithm derived from Aczel pg 15 bottom
            If p >= 100.0 Then
                Return sortedData(sortedData.Length - 1)
            End If

            Dim position As Double = CDbl(sortedData.Length + 1) * p / 100.0
            Dim leftNumber As Double = 0.0, rightNumber As Double = 0.0

            Dim n As Double = p / 100.0 * (sortedData.Length - 1) + 1.0

            If position >= 1 Then
                leftNumber = sortedData(CInt(System.Math.Floor(n)) - 1)
                rightNumber = sortedData(CInt(System.Math.Floor(n)))
            Else
                leftNumber = sortedData(0)
                ' first data
                ' first data
                rightNumber = sortedData(1)
            End If

            If leftNumber = rightNumber Then
                Return leftNumber
            Else
                Dim part As Double = n - System.Math.Floor(n)
                Return leftNumber + part * (rightNumber - leftNumber)
            End If

        Catch ex As Exception
            Return Double.NaN
        End Try

    End Function

    'Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
    '    target = value
    '    Return value
    'End Function

#End Region

    'http://jumbo.uni-muenster.de/index.php?id=185
    Public Function getGauss(
                            value As Double,
                            stdDeviation As Double,
                            arithmeticMean As Double) As Double

        Dim temp01 As Double
        Dim temp02 As Double

        Try

            temp01 = 1 /
                            (stdDeviation * Math.Sqrt(2 * Math.PI))

            temp02 = Math.Exp(-0.5 * Math.Pow(
                                              x:=(value - arithmeticMean) / stdDeviation,
                                              y:=2))

        Catch ex As Exception
            Return Double.NaN
        End Try


        Return temp01 * temp02

    End Function

#End Region

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const CATBasic As String = "01 Basic"
#Region "    Basic"


    ''' <summary>
    ''' name to display and for .tostring
    ''' </summary>
    ''' <returns></returns>
    <DisplayName(
            "Name")>
    <Browsable(True)>
    <[ReadOnly](False)>
    <Category(CATBasic)>
    <RefreshProperties(RefreshProperties.All)>
    <XmlIgnore> <ScriptIgnore>
    Public ReadOnly Overridable Property Name As String
        Get
            Return Me.Percent & "th Percentile" & Me.Percentile
        End Get
    End Property

#Region "    Data"

    ''' <summary>
    ''' add a single value to the dataset
    ''' </summary>
    ''' <param name="value">data point as double</param>
    Public Sub AddValue(value As Double)

        Dim temp As New List(Of Double)

        If Not IsNothing(Me.Data) AndAlso
                   Not IsNothing(value) Then

            temp.AddRange(Me.Data)
            temp.Add(value)

            Me.Data = temp.ToArray

        End If
    End Sub

    ''' <summary>
    ''' add multiple values to the dataset
    ''' </summary>
    ''' <param name="values">data points as array of double</param>
    Public Sub AddValue(values As Double())

        Dim temp As New List(Of Double)

        If Not IsNothing(values) Then

            If Not IsNothing(Me.Data) Then
                temp.AddRange(Me.Data)
            End If

            temp.AddRange(values)

            Me.Data = temp.ToArray

        End If

    End Sub

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _dataImporter As String() = {}

    ''' <summary>
    ''' orig.data is used to calculate percentiles
    ''' </summary>
    ''' <returns></returns>
    <DisplayName(
    "Data  importer")>
    <Browsable(True)>
    <[ReadOnly](False)>
    <Category(CATBasic)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property DataImporter As String()
        Get
            Return _dataImporter
        End Get
        Set

            Dim tempDbl As New List(Of Double)
            Dim tempString As New List(Of String)

            For Each member As String In Value

                Try
                    tempDbl.Add(Double.Parse(Trim(member)))
                    tempString.Add(tempDbl.Last.ToString)
                Catch ex As Exception

                End Try

            Next

            _data = tempDbl.ToArray
            Analyze()

            _dataImporter = Value

        End Set
    End Property


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _data As Double()

    ''' <summary>
    ''' orig.data is used to calculate percentiles
    ''' </summary>
    ''' <returns></returns>
    <DisplayName(
            "Data, as double")>
    <Browsable(False)>
    <[ReadOnly](True)>
    <Category(CATBasic)>
    <RefreshProperties(RefreshProperties.All)>
    <Editor(GetType(multiLineDoubleArrayEditor), GetType(UITypeEditor))>
    Public Property Data As Double()
        Get
            Return _data
        End Get
        Set(value As Double())

            Dim temp As New List(Of String)

            For Each member As Double In value
                temp.Add(member.ToString)
            Next

            _dataImporter = temp.ToArray
            _data = value

            Analyze()

        End Set
    End Property


    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _sortedData As Double()

    ''' <summary>
    ''' sortedData is used to calculate percentiles
    ''' </summary>
    ''' <returns></returns>
    <XmlIgnore> <ScriptIgnore>
    <DisplayName(
            " '' , sorted")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATBasic)>
    <RefreshProperties(RefreshProperties.All)>
    <Editor(GetType(multiLineDoubleArrayEditor), GetType(UITypeEditor))>
    Public Property SortedData As Double()
        Get
            Return _sortedData
        End Get
        Set
            _sortedData = Value
        End Set
    End Property

    ''' <summary>
    ''' sortedData is used to calculate percentiles
    ''' </summary>
    ''' <returns></returns>
    <DisplayName(
            " '' , order")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATBasic)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property Order As Integer()

#End Region

    ''' <summary>
    ''' Count
    ''' </summary>
    <Description(
            "Number of elements")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATBasic)>
    <XmlIgnore> <ScriptIgnore>
    Public Property N As Integer

    ''' <summary>
    ''' Sum
    ''' </summary>
    <Description(
            "Sum of elements")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATBasic)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property Sum As Double = Double.NaN


    ''' <summary>
    ''' The range of the values
    ''' </summary>
    <Description(
            "Range of the values")>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATBasic)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property Range As Double = Double.NaN

#End Region

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const CATMeans As String = "02 Mean"
#Region "    Mean"

    ''' <summary>
    ''' Arithmetic mean
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATMeans)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    <Description(
            "(a1 + a2  +...+ an)/n")>
    Public Property Arithmetic As Double = Double.NaN

    ''' <summary>
    ''' Geometric mean
    ''' </summary>
    ''' 
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATMeans)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    <Description(
            "(a1 * a2 *...* an)^1/n" & vbCrLf &
            "not def. if any value is 0 or neg.")>
    Public Property Geometric As Double = Double.NaN

    ''' <summary>
    ''' Harmonic mean
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATMeans)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    <Description(
            "n(1/a1 + 1/a2 +...+ 1/an)" & vbCrLf &
            "not def. if any value is 0 or neg.")>
    Public Property Harmonic As Double = Double.NaN

    ''' <summary>
    ''' Median, or second quartile, or at 50 percentile
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATMeans)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    <Description("= second quartile or 50th percentile")>
    Public Property Median As Double = Double.NaN

#End Region

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const CATWhisker As String = "03 Whisker"
#Region "    Whisker"

    ''' <summary>
    ''' Minimum value
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATWhisker)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property Min As Double = Double.NaN

    ''' <summary>
    ''' First quartile, at 25 percentile
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATWhisker)>
    <DisplayName(
            "1st Quartile")>
    <Description(
            "First quartile = 25th percentile")>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property FirstQuartile As Double = Double.NaN

    ''' <summary>
    ''' Interquartile range
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATWhisker)>
    <DisplayName(
            "IQR")>
    <Description(
            "Inter Quartile Range")>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property IQR As Double = Double.NaN

    ''' <summary>
    ''' Third quartile, at 75 percentile
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATWhisker)>
    <DisplayName(
            "3rd Quartile")>
    <Description(
            "Third Quartile")>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property ThirdQuartile As Double = Double.NaN

    ''' <summary>
    ''' Maximum value
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATWhisker)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property Max As Double = Double.NaN

    <DebuggerBrowsable(False)>
    Private m_Percent As Double = 80

    ''' <summary>
    ''' Percentile
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](False)>
    <Category(CATWhisker)>
    <DisplayName(
            "Percentile")>
    <Description(
            "Percentile in percent 0 - 100%")>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    <DefaultValue(80)>
    Public Property Percent As Double
        Get
            Return m_Percent
        End Get
        Set(value As Double)

            m_Percent = value
            Percentile = calcPercentile(percent:=Me.Percent)

        End Set
    End Property

    ''' <summary>
    ''' Percentile
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATWhisker)>
    <DisplayName(
                "Value")>
    <Description("Percentile value" & vbCrLf &
                "EXCEL QUANTIL.INKL")>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property Percentile As Double = Double.NaN

#End Region

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const CATStatistic As String = "04 Statistic"
#Region "    Statistic"

    Private Const StatisticDetailsBrowsable As Boolean = False

    Public Enum eSampleTotality
        Sample
        Totality
    End Enum

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private m_Base As eSampleTotality = eSampleTotality.Sample

    ''' <summary>
    ''' Sample or Totality
    ''' for std. deviation and variance
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](False)>
    <Category(CATStatistic)>
    <DisplayName("Sample or Totality")>
    <RefreshProperties(RefreshProperties.All)>
    <DefaultValue(CInt(eSampleTotality.Sample))>
    <XmlIgnore> <ScriptIgnore>
    Public Property Base As eSampleTotality
        Get
            Return m_Base
        End Get
        Set(value As eSampleTotality)
            m_Base = value
            Analyze()
        End Set
    End Property

    ''' <summary>
    ''' Sample variance
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATStatistic)>
    <DisplayName(
            "Variance")>
    <Description(
            "Sum of error^2 / (n - 1{Sample}, 0{Total})" & vbCrLf &
            "EXCEL : VARIANZA sample; VARIANZENA totality" & vbCrLf &
            "1/n * (a1^2 + a2^2 + ... + an^2) - mean^2")>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property Variance As Double = Double.NaN

    ''' <summary>
    ''' Sample variance
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATStatistic)>
    <DisplayName(
            "Relative Variance in %")>
    <Description(
            "std deviation / mean" & vbCrLf &
            "")>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property RelativeVariance As Double = Double.NaN

    ''' <summary>
    ''' Sample standard deviation
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATStatistic)>
    <DisplayName(
            "Std. Deviation")>
    <Description(
            "sqrt(Variance)" & vbCrLf &
            "EXCEL : STABWA sample; STABWNA totality")>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property StdDeviation As Double = Double.NaN

    ''' <summary>
    ''' orig.data is used to calculate percentiles
    ''' </summary>
    ''' <returns></returns>
    <DisplayName(
            "Gaussian distribution")>
    <Browsable(StatisticDetailsBrowsable)>
    <[ReadOnly](False)>
    <Category(CATStatistic)>
    <RefreshProperties(RefreshProperties.All)>
    <XmlIgnore> <ScriptIgnore>
    <Editor(GetType(multiLineDoubleArrayEditor), GetType(UITypeEditor))>
    Public Property Gauss As Double()

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private _norm As Boolean = False

    ''' <summary>
    ''' orig.data is used to calculate percentiles
    ''' </summary>
    ''' <returns></returns>
    <DisplayName(
            " '     ' norm. to 100?")>
    <Browsable(StatisticDetailsBrowsable)>
    <[ReadOnly](False)>
    <Category(CATStatistic)>
    <RefreshProperties(RefreshProperties.All)>
    <DefaultValue(False)>
    <XmlIgnore> <ScriptIgnore>
    Public Property Norm As Boolean
        Get
            Return _norm
        End Get
        Set(value As Boolean)
            _norm = value
            Analyze()
        End Set
    End Property

    ''' <summary>
    ''' Skewness of the data distribution
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATStatistic)>
    <DisplayName(
            "Skewness")>
    <Description(
            "Measure of the asymmetry" & vbCrLf &
            "negative/positive: left/right tail; EXCEL Schief")>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property Skewness As Double = Double.NaN

    ''' <summary>
    ''' Kurtosis of the data distribution
    ''' </summary>
    <Browsable(True)>
    <[ReadOnly](True)>
    <Category(CATStatistic)>
    <DisplayName(
            "Kurtosis")>
    <Description(
            "measure of the 'taillessness' of the probability distribution" & vbCrLf &
            "positive/negative: sharp/wide distribution; EXCEL Kurt")>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property Kurtosis As Double = Double.NaN

#End Region

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const CATStatHelper As String = "05 Helper"
#Region "    05 Helper"

    Private Const HelperBrowsable As Boolean = False

    ''' <summary>
    ''' Sum of Error
    ''' </summary>
    <Browsable(HelperBrowsable)>
    <[ReadOnly](True)>
    <Category(CATStatHelper)>
    <DisplayName(
            "Sum of error")>
    <Description(
            "(|a1| - aver.) + (|a2| - aver.) + ... + (|a3| - aver.)")>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property SumOfError As Double = Double.NaN

    ''' <summary>
    ''' The sum of the squares of errors
    ''' </summary>
    <Browsable(HelperBrowsable)>
    <[ReadOnly](True)>
    <Category(CATStatHelper)>
    <DisplayName(
            "Sum of error^2")>
    <Description(
            "(|a1| - aver.)^2 + (|a2| - aver.)^2 + ... + (|a3| - aver.)^2")>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property SumOfErrorSquare As Double = Double.NaN


    ''' <summary>
    ''' Sum of Products
    ''' </summary>
    <Browsable(HelperBrowsable)>
    <[ReadOnly](True)>
    <Category(CATStatHelper)>
    <DisplayName(
            "Sum of products")>
    <Description(
            "a1 * a2 * ... * an")>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property SumOfProducts As Double = Double.NaN

    ''' <summary>
    ''' Sum of reciprocal Products
    ''' </summary>
    <Browsable(HelperBrowsable)>
    <[ReadOnly](True)>
    <Category(CATStatHelper)>
    <DisplayName(
            "Sum of reciprocal products")>
    <Description(
            "1/a1 + 1/a2 + ... + 1/an")>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(DblConv))>
    <XmlIgnore> <ScriptIgnore>
    Public Property SumOfReciprocalProducts As Double = Double.NaN

#End Region

End Class
