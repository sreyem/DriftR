
Imports System.ComponentModel

Public Module Enums

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
                "Field beans " & vbCrLf &
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
        ''' OS Oil seed rape, spring
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
        ''' PF Pome/stone fruits
        ''' add. parameter BBCH needed
        ''' </summary>
        <Description(
                "PF " & vbCrLf &
                "Pome fruit " & vbCrLf &
                "add. parameter BBCH needed")>
        PF

        ''' <summary>
        ''' PFE Pome/stone fruits
        ''' early BBCH 01 - 71, 96 - 99
        ''' </summary>
        <Description(
                "PFE " & vbCrLf &
                "Pome fruit " & vbCrLf &
                "early BBCH 01-71, 96-99")>
        PFE

        ''' <summary>
        ''' PFL Pome/stone fruits
        ''' late  BBCH 72 - 95
        ''' </summary>
        <Description(
                "PFL " & vbCrLf &
                "Pome fruit " & vbCrLf &
                "late  BBCH 72-95")>
        PFL

        '-------------------------------------------------------------

        ''' <summary>
        ''' PS Potatoes, 1st/2nd
        ''' </summary>
        <Description(
                "PS " & vbCrLf &
                "Potatoes " & vbCrLf &
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
        SB

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
                "Vegetables, bulb " & vbCrLf &
                "1st/2nd")>
        VB

        ''' <summary>
        ''' VF Vegetables, fruiting
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
                "Vegetables, leafy " & vbCrLf &
                "1st/2nd")>
        VL

        ''' <summary>
        ''' VR Vegetables, root, 1st/2nd
        ''' </summary>
        <Description(
                "VR " & vbCrLf &
                "Vegetables, root " & vbCrLf &
                "1st/2nd")>
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
        '''' no drift (incorporation or seed treatment)
        '''' </summary>
        '<Description(
        '    "ND " & vbCrLf &
        '    "No drift ")>
        'ND_NoDrift

    End Enum

    ''' <summary>
    ''' Target crop season
    ''' 1st, 2nd or both
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eSeason)))>
    Public Enum eSeason

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        <Description("1st")>
        first

        <Description("2nd")>
        second

        '<Description("1&2")>
        'both

    End Enum

    Public Function GetBBCHSearchStringsFromCrop(
                                   FOCUSswDriftCrop As eFOCUSswDriftCrop,
                          Optional Season As eSeason = eSeason.first) As String

        Dim temp As String()

        If FOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined Then Return ""

        Select Case FOCUSswDriftCrop

            Case eFOCUSswDriftCrop.HP,
                 eFOCUSswDriftCrop.CI,
                 eFOCUSswDriftCrop.OL,
                 eFOCUSswDriftCrop.GA

                Return ""

            Case eFOCUSswDriftCrop.CW

                Return "Cereals, Winter"

            Case eFOCUSswDriftCrop.CS

                Return "Cereals, Spring"

            Case eFOCUSswDriftCrop.PF,
                 eFOCUSswDriftCrop.PFE,
                 eFOCUSswDriftCrop.PFL

                Return "Pome fruit"

            Case eFOCUSswDriftCrop.VI,
                 eFOCUSswDriftCrop.VIL

                Return "Vines"

            Case eFOCUSswDriftCrop.SY

                Return "Soybean"

            Case eFOCUSswDriftCrop.SU
                Return "Sunflower"

        End Select

        temp =
            enumConverter(Of eFOCUSswDriftCrop).getEnumDescription(FOCUSswDriftCrop).Split(vbCrLf)

        For counter As Integer = 0 To temp.Count - 1

            temp(counter) =
                Replace(temp(counter), vbLf, "")

        Next

        If temp.Last.Contains("1st/2nd") Then

            'If Season = eSeason.both Then
            '    Return {description(1) & "1st", description(1) & "2nd"}
            'Else
            If Season = eSeason.first Then
                Return temp(1) & "1st"
            Else
                Return temp(1) & "2nd"
            End If

        Else
            Return temp(1)
        End If

    End Function


    ''' <summary>
    ''' simple enum 0 - 99
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eBBCH)))>
    Public Enum eBBCH

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        <Description("00")>
        _0 = 0
        <Description("01")>
        _1
        <Description("02")>
        _2
        <Description("03")>
        _3
        <Description("04")>
        _4
        <Description("05")>
        _5
        <Description("06")>
        _6
        <Description("07")>
        _7
        <Description("08")>
        _8
        <Description("09")>
        _9
        <Description("10")>
        _10
        <Description("11")>
        _11
        <Description("12")>
        _12
        <Description("13")>
        _13
        <Description("14")>
        _14
        <Description("15")>
        _15
        <Description("16")>
        _16
        <Description("17")>
        _17
        <Description("18")>
        _18
        <Description("19")>
        _19
        <Description("20")>
        _20
        <Description("21")>
        _21
        <Description("22")>
        _22
        <Description("23")>
        _23
        <Description("24")>
        _24
        <Description("25")>
        _25
        <Description("26")>
        _26
        <Description("27")>
        _27
        <Description("28")>
        _28
        <Description("29")>
        _29
        <Description("30")>
        _30
        <Description("31")>
        _31
        <Description("32")>
        _32
        <Description("33")>
        _33
        <Description("34")>
        _34
        <Description("35")>
        _35
        <Description("36")>
        _36
        <Description("37")>
        _37
        <Description("38")>
        _38
        <Description("39")>
        _39
        <Description("40")>
        _40
        <Description("41")>
        _41
        <Description("42")>
        _42
        <Description("43")>
        _43
        <Description("44")>
        _44
        <Description("45")>
        _45
        <Description("46")>
        _46
        <Description("47")>
        _47
        <Description("48")>
        _48
        <Description("49")>
        _49
        <Description("50")>
        _50
        <Description("51")>
        _51
        <Description("52")>
        _52
        <Description("53")>
        _53
        <Description("54")>
        _54
        <Description("55")>
        _55
        <Description("56")>
        _56
        <Description("57")>
        _57
        <Description("58")>
        _58
        <Description("59")>
        _59
        <Description("60")>
        _60
        <Description("61")>
        _61
        <Description("62")>
        _62
        <Description("63")>
        _63
        <Description("64")>
        _64
        <Description("65")>
        _65
        <Description("66")>
        _66
        <Description("67")>
        _67
        <Description("68")>
        _68
        <Description("69")>
        _69
        <Description("70")>
        _70
        <Description("71")>
        _71
        <Description("72")>
        _72
        <Description("73")>
        _73
        <Description("74")>
        _74
        <Description("75")>
        _75
        <Description("76")>
        _76
        <Description("77")>
        _77
        <Description("78")>
        _78
        <Description("79")>
        _79
        <Description("80")>
        _80
        <Description("81")>
        _81
        <Description("82")>
        _82
        <Description("83")>
        _83
        <Description("84")>
        _84
        <Description("85")>
        _85
        <Description("86")>
        _86
        <Description("87")>
        _87
        <Description("88")>
        _88
        <Description("89")>
        _89
        <Description("90")>
        _90
        <Description("91")>
        _91
        <Description("92")>
        _92
        <Description("93")>
        _93
        <Description("94")>
        _94
        <Description("95")>
        _95
        <Description("96")>
        _96
        <Description("97")>
        _97
        <Description("98")>
        _98
        <Description("99")>
        _99


    End Enum


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
    ''' all FOCUSsw scenarios D1 - D6 and R1 - R4
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eFOCUSswScenario)))>
    Public Enum eFOCUSswScenario

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        ''' <summary>
        ''' D1 Lanna Ditch|Stream
        ''' </summary>
        <Description(
                "D1 " & vbCrLf &
                "Lanna")>
        D1 = 0

        ''' <summary>
        ''' D2 Brimstone Ditch|Stream
        ''' </summary>
        <Description(
                "D2 " & vbCrLf &
                "Brimstone")>
        D2

        ''' <summary>
        ''' D3 Vredepeel Ditch
        ''' </summary>
        <Description(
                "D3 " & vbCrLf &
                "Vredepeel")>
        D3

        ''' <summary>
        ''' D4 Skousbo Pond|Stream
        ''' </summary>
        <Description(
                "D4 " & vbCrLf &
                "Skousbo")>
        D4

        ''' <summary>
        ''' D5 La Jailliere Pond|Stream
        ''' </summary>
        <Description(
                "D5 " & vbCrLf &
                "La Jailliere")>
        D5

        ''' <summary>
        ''' D6 Thiva Ditch
        ''' </summary>
        <Description(
                "D6 " & vbCrLf &
                "Thiva")>
        D6

        ''' <summary>
        ''' R1 Weiherbach Pond|Stream
        ''' </summary>
        <Description(
                "R1 " & vbCrLf &
                "Weiherbach")>
        R1

        ''' <summary>
        ''' R2 Porto Stream
        ''' </summary>
        <Description(
                "R2 " & vbCrLf &
                "Porto")>
        R2

        ''' <summary>
        ''' R3 Bologna Stream
        ''' </summary>
        <Description(
                "R3 " & vbCrLf &
                "Bologna")>
        R3

        ''' <summary>
        ''' R4 Roujan Stream
        ''' </summary>
        <Description(
                "R4 " & vbCrLf &
                "Roujan")>
        R4

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
        <Description(
                "AA " & vbCrLf &
                "Aerial appln.")>
        AA

        ''' <summary>
        ''' air blast, for hops, pome/stone fruits, vines, olives and citrus
        ''' </summary>
        <Description(
                "AB " & vbCrLf &
                "Air Blast")>
        AB

        ''' <summary>
        ''' granular appln., no drift and interception
        ''' </summary>
        <Description(
                "GR " & vbCrLf &
                "Granular appln.")>
        GR

        ''' <summary>
        ''' ground spray, std. incl. drift and interception
        ''' </summary>
        <Description(
                "GS " & vbCrLf &
                "Ground spray")>
        GS

        ''' <summary>
        ''' soil incorp. needs depth (std. 4cm)
        ''' </summary>
        <Description(
                "SI " & vbCrLf &
                "Soil incorp.")>
        SI

        ''' <summary>
        ''' HL Appln, hand (crop less than 50 cm)
        ''' </summary>
        <Description(
                "HL " & vbCrLf &
                "Hand appln. (crop < 50 cm)")>
        HL

        ''' <summary>
        ''' HH Appln, hand (crop higher than 50 cm)
        ''' </summary>
        <Description(
                "HH " & vbCrLf &
                "Hand appln. (crop > 50 cm)")>
        HH


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
        DI = 0

        <Description(
                "ST " & vbCrLf &
                "stream ")>
        ST

        <Description(
                "PO " & vbCrLf &
                "pond ")>
        PO

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
    ''' Drift red. nozzles
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eNozzles)))>
    Public Enum eNozzles

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        <Description(" 0%")>
        _0 = 0

        <Description("10%")>
        _10 = 10

        <Description("20%")>
        _20 = 20

        <Description("25%")>
        _25 = 25

        <Description("30%")>
        _30 = 30

        <Description("40%")>
        _40 = 40

        <Description("50%")>
        _50 = 50

        <Description("60%")>
        _60 = 60

        <Description("70%")>
        _70 = 70

        <Description("75%")>
        _75 = 75

        <Description("80%")>
        _80 = 80

        <Description("90%")>
        _90 = 90

        <Description("95%")>
        _95 = 95

    End Enum

    ''' <summary>
    ''' Pic water body depth value out of sim. years
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eDepthValueMode)))>
    Public Enum eDepthValueMode

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        <Description("Last")>
        last = 1
        <Description("Min")>
        min = 2
        <Description("Max")>
        max = 3

        <Description(" 5th perc")>
        _5 = 5
        <Description("10th perc")>
        _10 = 10
        <Description("15th perc")>
        _15 = 15
        <Description("20th perc")>
        _20 = 20
        <Description("25th perc")>
        _25 = 25
        <Description("30th perc")>
        _30 = 30
        <Description("35th perc")>
        _35 = 35
        <Description("40th perc")>
        _40 = 40
        <Description("45th perc")>
        _45 = 45
        <Description("50th perc")>
        _50 = 50
        <Description("55th perc")>
        _55 = 55
        <Description("60th perc")>
        _60 = 60
        <Description("65th perc")>
        _65 = 65
        <Description("70th perc")>
        _70 = 70
        <Description("75th perc")>
        _75 = 75
        <Description("80th perc")>
        _80 = 80
        <Description("85th perc")>
        _85 = 85
        <Description("90th perc")>
        _90 = 90
        <Description("95th perc")>
        _95 = 95

        <Description("Year  1")>
        _Y1 = 101
        <Description("Year  2")>
        _Y2 = 102
        <Description("Year  3")>
        _Y3 = 103
        <Description("Year  4")>
        _Y4 = 104
        <Description("Year  5")>
        _Y5 = 105
        <Description("Year  6")>
        _Y6 = 106
        <Description("Year  7")>
        _Y7 = 107
        <Description("Year  8")>
        _Y8 = 108
        <Description("Year  9")>
        _Y9 = 109
        <Description("Year 10")>
        _Y10 = 110
        <Description("Year 11")>
        _Y11 = 111
        <Description("Year 12")>
        _Y12 = 112
        <Description("Year 13")>
        _Y13 = 113
        <Description("Year 14")>
        _Y14 = 114
        <Description("Year 15")>
        _Y15 = 115
        <Description("Year 16")>
        _Y16 = 116
        <Description("Year 17")>
        _Y17 = 117
        <Description("Year 18")>
        _Y18 = 118
        <Description("Year 19")>
        _Y19 = 119
        <Description("Year 20")>
        _Y20 = 120

    End Enum


    ''' <summary>
    ''' Number of applns.
    ''' 1 - 8 (or more)
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of eNoOfApplns)))>
    Public Enum eNoOfApplns

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

        <Description("1  90th Percentile")>
        _01 = 0

        <Description("2  82th Percentile")>
        _02

        <Description("3  77th Percentile")>
        _03

        <Description("4  74th Percentile")>
        _04

        <Description("5  72th Percentile")>
        _05

        <Description("6  70th Percentile")>
        _06

        <Description("7  69th Percentile")>
        _07

        <Description("8+ 67th Percentile")>
        _08

    End Enum


    ''' <summary>
    ''' RunOff Scenarios R1 - R4
    ''' </summary>
    <TypeConverter(GetType(enumConverter(Of ePRZMScenario)))>
    Public Enum ePRZMScenario

        <Description(enumConverter(Of Type).not_defined)>
        not_defined = -1

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


End Module

Public Module DBs

#Region "    Scenario X Water Body"

    Public scenarioXWaterBody As String() =
               {
                   "D1|DI|ST",
                   "D2|DI|ST",
                   "D3|DI",
                   "D4|PO|ST",
                   "D5|PO|ST",
                   "D6|DI",
                   "R1|PO|ST",
                   "R2|ST",
                   "R3|ST",
                   "R4|ST"
               }

    Public Function GetWaterBodysFromScenario(Scenario As eFOCUSswScenario) As String()

        Dim WBs As String()
        Dim out As New List(Of String)
        Dim enumWaterBody As eFOCUSswWaterBody

        WBs =
            Filter(
                Source:=scenarioXWaterBody,
                Match:=Scenario.ToString,
                Include:=True,
                Compare:=CompareMethod.Text)


        If WBs.Count = 0 Then Return {}

        WBs = WBs.First.Split("|")

        If WBs.Count < 1 Then Return {}

        For counter As Integer = 1 To WBs.Count - 1

            enumWaterBody = [Enum].Parse(
                        enumType:=GetType(eFOCUSswWaterBody),
                        value:=WBs(counter))
            out.Add(
            enumConverter(Of eFOCUSswWaterBody).getEnumDescription(
                EnumConstant:=enumWaterBody))

        Next

        Return out.ToArray

    End Function

    Public Function GetScenariosFromCrop(Crop As eFOCUSswDriftCrop) As String()

        Dim Scenarios As String()
        Dim out As New List(Of String)
        Dim enumScenario As eFOCUSswScenario

        Scenarios =
            Filter(
                Source:=cropXScenarios,
                Match:=Crop.ToString.Substring(0, 2),
                Include:=True,
                Compare:=CompareMethod.Text)

        If Scenarios.Count = 0 Then Return {}

        Scenarios = Scenarios.First.Split("|")

        If Scenarios.Count < 1 Then Return {}

        For counter As Integer = 1 To Scenarios.Count - 1

            enumScenario = [Enum].Parse(
                        enumType:=GetType(eFOCUSswScenario),
                        value:=Scenarios(counter))
            out.Add(
            enumConverter(Of eFOCUSswWaterBody).getEnumDescription(
                EnumConstant:=enumScenario))

        Next

        Return out.ToArray

    End Function


#End Region


    Public cropXScenarios As String() =
                {
                    "CS|D1|D3|D4|D5|R4",
                    "CW|D1|D2|D3|D4|D5|D6|R1|R3|R4",
                    "CI|D6|R4",
                    "CO|D6",
                    "FB|D2|D3|D4|D6|R1|R2|R3|R4",
                    "GA|D1|D2|D3|D4|D5|R2|R3",
                    "HP|R1",
                    "LG|D3|D4|D5|D6|R1|R2|R3|R4",
                    "MZ|D3|D4|D5|D6|R1|R2|R3|R4",
                    "OS|D1|D3|D4|D5|R1",
                    "OW|D2|D3|D4|D5|R1|R3",
                    "OL|D6|R4",
                    "PF|D3|D4|D5|R1|R2|R3|R4",
                    "PS|D3|D4|D6|R1|R2|R3",
                    "SY|R3|R4",
                    "SB|D3|D4|R1|R3",
                    "SU|D5|R1|R3|R4",
                    "TB|R3",
                    "VB|D3|D4|D6|R1|R2|R3|R4",
                    "VF|D6|R2|R3|R4",
                    "VL|D3|D4|D6|R1|R2|R3|R4",
                    "VR|D3|D6|R1|R2|R3|R4",
                    "VI|D6|R1|R2|R3|R4"
                }


    Public Function GetMACRO_pFileFromCropScenario(
                                    Crop As eFOCUSswDriftCrop,
                                    Scenario As eFOCUSswScenario) As String


        Try
            Return Filter(
            Source:=cropScenarioXMACRO_p,
            Match:=Scenario.ToString & "," & Crop.ToString,
            Include:=True,
            Compare:=CompareMethod.Text).First.Split(",").Last

        Catch ex As Exception
            Return ""
        End Try

    End Function


    Public cropScenarioXMACRO_p As String() =
            {
            "D3,MZ,vr_ima_p.BIN",
            "D3,OS,vrede_p.BIN",
            "D3,OW,vrede_p.BIN",
            "D3,PF,vr_ipf_p.BIN",
            "D3,PS,vr_ipo_p.BIN",
            "D3,SB,vr_isb_p.BIN",
            "D3,VB,vr_ibv_p.BIN",
            "D3,VL,vr_ilv_p.BIN",
            "D3,VR,vr_irv_p.BIN",
            "D4,CS,skous_p.BIN",
            "D4,CW,skous_p.BIN",
            "D4,FB,skous_p.BIN",
            "D4,GA,skous_p.BIN",
            "D4,LG,sk_ile_p.BIN",
            "D4,MZ,skous_p.BIN",
            "D4,OS,skous_p.BIN",
            "D4,OW,skous_p.BIN",
            "D4,PF,skous_p.BIN",
            "D4,PS,sk_ipo_p.BIN",
            "D4,SB,sk_isb_p.BIN",
            "D4,VB,sk_ibv_p.BIN",
            "D4,VL,sk_ilv_p.BIN",
            "D5,CS,jaill_p.BIN",
            "D5,CW,jaill_p.BIN",
            "D5,GA,jaill_p.BIN",
            "D5,LG,jaill_p.BIN",
            "D5,MZ,jaill_p.BIN",
            "D5,OS,jaill_p.BIN",
            "D5,OW,jaill_p.BIN",
            "D5,PF,jaill_p.BIN",
            "D5,SU,jaill_p.BIN",
            "D6,CI,th_ici_p.BIN",
            "D6,CO,th_ico_p.BIN",
            "D6,CW,thebe_p.BIN",
            "D6,FB,th_ifb_p.BIN",
            "D6,LG,th_ile_p.BIN",
            "D6,MZ,th_ima_p.BIN",
            "D6,OL,thebe_p.BIN",
            "D6,PS,th_ipo_p.BIN",
            "D6,VB,th_ibv_p.BIN",
            "D6,VF,th_ifv_p.BIN",
            "D6,VI,thebe_p.BIN",
            "D6,VL,th_ilv_p.BIN",
            "D6,VR,th_irv_p.BIN"
            }


#Region "    Regression parameters and hinge point"

    Public Property RegressionA As Double(,) =
        {
            {2.7593, 2.4376, 2.0244, 1.8619, 1.7942, 1.6314, 1.5784, 1.5119},
            {66.702, 62.272, 58.796, 58.947, 58.111, 58.829, 59.912, 59.395},
            {60.396, 42.002, 40.12, 36.273, 34.591, 31.64, 31.561, 29.136},
            {58.247, 66.243, 60.397, 58.559, 59.548, 60.136, 59.774, 53.2},
            {15.793, 15.461, 16.887, 16.484, 15.648, 15.119, 14.675, 14.948},
            {44.769, 40.262, 39.314, 37.401, 37.767, 36.908, 35.498, 35.094},
            {50.47, 50.47, 50.47, 50.47, 50.47, 50.47, 50.47, 50.47}
        }

    Public Property RegressionB As Double(,) =
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
    Public Property RegressionC As Double(,) =
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
    Public Property RegressionD As Double(,) =
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
    Public Property HingePointDB As Double(,) =
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



End Module
