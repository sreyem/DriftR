Imports System
Imports System.Collections.Generic

Namespace GAPExporterLib
    Public Class GAPData
        Public Property ActiveIngredients As List(Of String) = New List(Of String)(9)
        Public Property Biologics As String
        Public Property Description As String
        Public Property ESPID As Integer
        Public Property ExportedBy As String
        Public Property ExportedOn As DateTime?
        Public Property GAPMetadataID As Decimal
        Public Property GAPSummaryList As List(Of GAPSummary) = New List(Of GAPSummary)()
        Public Property ProductConcentrationUnit As String
        Public Property SeedTreatment As String
        Public Property TemplateVersion As String
        Public Property HasGapGroups As Boolean = False
    End Class

    Public Class GAPSummary
        Public Property Active As Boolean
        Public Property ApplicationMaxNumber As String
        Public Property ApplicationMethod As String
        Public Property ApplicationMinInterval As String
        Public Property ApplicationRateActive As String
        Public Property ApplicationRateProduct As String
        Public Property ApplicationTiming As String
        Public Property BBCHEarliest As Decimal?
        Public Property BBCHLatest As Decimal?
        Public Property BBCHRange As String
        Public Property ConcentrationRule As String
        Public Property CountryZoneID As Integer
        Public Property Crop As String
        Public Property CropDataID As Integer
        Public Property FGI As String
        Public Property GAPMetadataID As Integer
        Public Property GAPSummaryID As Integer
        Public Property GHType As String
        Public Property MaxRateMaxAIConcs As List(Of Double?) = New List(Of Double?)(9)
        Public Property MaxRateMaxProdConc As Double?
        Public Property MaxRateMinAIConcs As List(Of Double?) = New List(Of Double?)(9)
        Public Property MaxRateMinProdConc As Double?
        Public Property MaxSingleAIRates As List(Of Double?) = New List(Of Double?)(9)
        Public Property MaxSingleProdRate As Double?
        Public Property MaxSlurryVolume As Double
        Public Property MaxSowingDensity As Double
        Public Property MaxTGW As Double
        Public Property MaxTotalAIRates As List(Of Double?) = New List(Of Double?)(9)
        Public Property MaxTotalNoOfApps As Decimal?
        Public Property MaxTotalProdRate As Double?
        Public Property MemberState As String
        Public Property MinSingleProdRate As String
        Public Property MinSlurryVolume As Double
        Public Property MinSowingDensity As Double
        Public Property MinTGW As Double
        Public Property MinorCrop As String
        Public Property NumberOfBlock As Integer
        Public Property OverallMaxAIConcs As List(Of Double?) = New List(Of Double?)(9)
        Public Property OverallMaxProdConc As Double?
        Public Property OverallMinAIConcs As List(Of Double?) = New List(Of Double?)(9)
        Public Property OverallMinInterval As Integer
        Public Property OverallMinProdConc As Double?
        Public Property OverallMinSingleProdRate As Double?
        Public Property PHI As String
        Public Property Pests As String
        Public Property ProductSeedLoading As Double
        Public Property RegulatoryZone As String
        Public Property Remarks As String
        Public Property SeedLoadingAIs As List(Of Double) = New List(Of Double)(9)
        Public Property TreatmentProductLoading As Double
        Public Property UnitSize As String
        Public Property UseNumber As String
        Public Property Water As String
        Public Property Zone As String
        Public Property GAPGroupIdentifier As String
        Public Property GAPGroupTitle As String
    End Class
End Namespace

