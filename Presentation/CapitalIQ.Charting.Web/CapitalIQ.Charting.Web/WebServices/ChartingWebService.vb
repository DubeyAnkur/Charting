'''' -----------------------------------------------------------------------------
'''' <summary>
'''' 
'''' </summary>
'''' <remarks>
'''' </remarks>
'''' <history>
'''' 	[wmassie]	12/12/2005	Created
'''' </history>
'''' -----------------------------------------------------------------------------
Imports CapitalIQ.Constants
Imports CapitalIQ.Web.Services
Imports CapitalIQ.Charting.Series
Imports System.Web.Services
Imports System.Xml.Serialization
Imports System.Collections.Generic
Imports System.Linq

Imports ccs = CapitalIQ.Charting.Services
Imports ccc = CapitalIQ.Charting.Common

Namespace CapitalIQ.Charting.Web.WebServices
    <System.Web.Services.WebService(Namespace:="https://www.capitaliq.com/CIQDotNet/Charting/ChartingWebService", Description:="WebService for getting chart data")> _
        Public Class ChartingWebService
        Inherits CapitalIQ.Web.Services.BaseWebService

        ' This is a list of the valid cachemoduleid's to pass to GetChartableList
        Private CHARTABLE_LISTS() As CapitalIQ.Caching.CacheModuleId = {CapitalIQ.Caching.CacheModuleId.UserTargetLists, _
                 CapitalIQ.Caching.CacheModuleId.UserWatchLists, _
                CapitalIQ.Caching.CacheModuleId.AllUserCoverageListsWithConstituents, _
                CapitalIQ.Caching.CacheModuleId.UserCompSets, _
                CapitalIQ.Caching.CacheModuleId.CommonIndices_refdata, _
                CapitalIQ.Caching.CacheModuleId.InterestRate_refData, _
                CapitalIQ.Caching.CacheModuleId.ChartCurrencies}

        ''' <summary>Creates an instance of IChartingBusinessMgr through RemoteUtil</summary>
        Private Function GetChartingBusinessMgr() As IChartingBusinessMgr
            Return DirectCast(RemoteUtil.GetRemote(GetType(IChartingBusinessMgr)), IChartingBusinessMgr)
        End Function

        <WebMethod(Description:="Get a dataset populated with the requested list", EnableSession:=True), _
            XmlInclude(GetType(CapitalIQ.Charting.Series.SeriesComponent)), _
            XmlInclude(GetType(CapitalIQ.Charting.Series.SeriesComponent.EstimateInfo)), _
            XmlInclude(GetType(CapitalIQ.Charting.Common.EstimatesRegion))> _
        Public Function GetChartableList(ByVal inCacheModuleId As Integer) As SeriesComponent()

            ' Make sure they only ask for one of the modules we want to let them see
            Dim cm As CapitalIQ.Caching.CacheModuleId
            If Array.IndexOf(CHARTABLE_LISTS, cm) >= 0 Then
                Throw New ArgumentException("Invalid CacheModuelId")
            End If
            cm = CType(inCacheModuleId, CapitalIQ.Caching.CacheModuleId)

            Dim ret As SeriesComponent() = CapitalIQ.Charting.Chartables.ChartableFactory.GetSearchComponentList(CapitalIQ.Caching.Cache.Get(Of DataSet)(New CapitalIQ.Caching.UserKey(cm)).Tables(0), "primaryTradingItemId")
            AddPlatformDefaultEstimateInfo(ret)
            Return ret
        End Function

        'Private Shared Sub CheckCapability()
        '    'Dim role As CapitalIQ.SessionState.UserRole = CapitalIQ.SessionState.UserRole.Current
        '    'If Not role.HasAny(CapabilityId.WebServiceCharting, CapabilityId.ExcelPlugIn) Then
        '    '    Throw New CapabilityException("User must have either WebServiceCharting capability or CapabilityId.ExcelPlugIn capability")
        '    'End If
        'End Sub

        ''' <remarks>I don't think this is used any more</remarks>
        <WebMethod(Description:="Gets the chartable financial dataitems.", EnableSession:=True)> _
        Public Function GetFinancialDataItems() As DataSet

            ' Now we can load the data
            Dim mgr As IChartingBusinessMgr = GetChartingBusinessMgr()
            Dim ds As DataSet = mgr.GetFinancialDataItems()
            Return ds

        End Function

        <WebMethod(Description:="Get an array of Strings representing the names of the members of this list", EnableSession:=True)> _
        Public Function GetListConstituentNames(ByVal inListTypeId As Integer, ByVal inListId As Integer) As List(Of String)
            Dim lt As ListTypeId = CType(inListTypeId, ListTypeId)

            Dim cm As Caching.CacheModuleId
            Dim ds As DataSet

            Select Case lt
                Case CapitalIQ.Constants.ListTypeId.CompSet
                    cm = Caching.CacheModuleId.UserCompSets
                Case CapitalIQ.Constants.ListTypeId.CoverageList
                    cm = Caching.CacheModuleId.AllUserCoverageListsWithConstituents
                Case CapitalIQ.Constants.ListTypeId.TargetList
                    cm = Caching.CacheModuleId.UserTargetLists
                Case CapitalIQ.Constants.ListTypeId.WatchList
                    cm = Caching.CacheModuleId.GetAllWatchListsDataSet
                Case Else
                    Throw New InvalidOperationException("Unhandled listType: " & lt)
            End Select

            ds = CapitalIQ.Charting.Web.Pages.Builder.Utility.GetListConstituentsDS(cm, inListId)

            ' Load the rows from the constituent table
            Dim ret As New List(Of String)
            For Each dr As DataRow In ds.Tables(1).Rows()
                If Not String.IsNullOrEmpty(CiqCore.Utility.NullConvert.ToStr(dr.Item("CompanyNamePlus"))) Then
                    ret.Add(CiqCore.Utility.NullConvert.ToStr(dr.Item("CompanyNamePlus")))
                ElseIf Not String.IsNullOrEmpty(CiqCore.Utility.NullConvert.ToStr(dr.Item("TradingItemNamePlus"))) Then
                    ret.Add(CiqCore.Utility.NullConvert.ToStr(dr.Item("TradingItemNamePlus")))
                End If

                ' Put a cap at 2001(we'll truncate to 2000 and tell them that's the limit.)
                If ret.Count = 2001 Then Exit For
            Next

            Return ret
        End Function


        <WebMethod(Description:="Get an array of SeriesComponents representing the members of this list", EnableSession:=True), _
            XmlInclude(GetType(CapitalIQ.Charting.Series.SeriesComponent)), _
            XmlInclude(GetType(CapitalIQ.Charting.Series.SeriesComponent.EstimateInfo)), _
            XmlInclude(GetType(CapitalIQ.Charting.Common.EstimatesRegion))> _
        Public Function GetListConstituents(ByVal inListTypeId As Integer, ByVal inListId As Integer) As SeriesComponent()

            ' Make sure they only ask for one of the modules we want to let them see
            Dim lt As ListTypeId = CType(inListTypeId, ListTypeId)

            Dim ret As SeriesComponent() = CapitalIQ.Charting.Web.Pages.Builder.Utility.GetListConstituents(lt, inListId)
            AddPlatformDefaultEstimateInfo(ret)
            Return ret
        End Function

        ''' <remarks>Curse you whoever decided it'd be a good idea to send Excel Plug-in a DataSet</remarks>
        <WebMethod(Description:="Get all the lists for current user", EnableSession:=True)> _
        Public Function GetListsForUser() As DataSet

            Dim mgr As IChartingBusinessMgr = GetChartingBusinessMgr()
            Dim lists As List(Of ListInfo) = mgr.GetAllListsForUser(True)

            Dim t As New DataTable()
            t.Columns.Add("listTypeId", GetType(Integer))
            t.Columns.Add("listId", GetType(Integer))
            t.Columns.Add("listName", GetType(String))
            t.Columns.Add("validForMarketCap", GetType(Boolean))

            For Each l As ListInfo In lists
                t.Rows.Add(l.ListTypeId, l.ListId, l.ListName, l.ValidForMarketCap.Value)
            Next

            Dim ds As New DataSet()
            ds.Tables.Add(t)

            Return ds

        End Function

        <WebMethod(Description:="Get Chart Data.", EnableSession:=True), _
         XmlInclude(GetType(ccs.Series)), _
         XmlInclude(GetType(ccs.FPSeries)), _
         XmlInclude(GetType(ccc.FPDataItem)), _
         XmlInclude(GetType(ccs.DTSChart)), _
         XmlInclude(GetType(ccs.FPChart)), _
         XmlInclude(GetType(ccs.STAChart)), _
         XmlInclude(GetType(ccc.CIQChartType))> _
        Public Function GetChartData(ByVal inChart As ccs.Chart) As ccs.Chart

            Dim comment As String = String.Format("Get data for {0} chart", inChart.Type.Name)
            CapitalIQ.ActivityManager.LogAsync(New CapitalIQ.Activity(ActivityTypeId.WebServiceGetChartData, comment))
            Dim cl As ChartLoader = ChartLoader.GetLoader(inChart)
            cl.LoadChart()

            Dim ret As ccs.Chart = cl.GetReturnChart()

            FixDates(ret.Data)

            Return ret

        End Function

        <WebMethod(Description:="Gets the requested key devs for the passed in company in the specified time frame", EnableSession:=True)> _
        Public Function GetCorporateTimeline(ByVal inCompanyId As Integer, _
                                            ByVal inFromDate As Date, _
                                            ByVal inToDate As Date, _
                                            ByVal inIncludeSubs As Boolean, _
                                            ByVal inCategoryArray As ArrayList) As DataSet

            ' Now we can load the data
            Dim mgr As IChartingBusinessMgr = GetChartingBusinessMgr()

            Dim cti As Chartables.ChartableTradingItem = Chartables.ChartableFactory.GetChartableTradingItemFromCompanyId(inCompanyId)

            ' The second parm keeps it from getting quote summary data
            ' modified [gsirois] 9/11/06
            Dim ds As DataSet = mgr.GetCorporateTimeLine(cti, 0, inIncludeSubs, inFromDate, inToDate, GetKeyDevFlags(inCategoryArray), GetKeyDevCategoryXML(inCategoryArray), SessionState.UserSession.Current.UserRole.CapabilityBitMask)

            FixDates(ds)

            Dim a As New Activity(CapitalIQ.Constants.ActivityTypeId.ExcelCharting_RetrievedKeyDevs, String.Empty)
            ActivityManager.LogAsync(a)

            Return ds

        End Function

        Private Shared Function GetKeyDevCategory(c As Object) As Types.KeyDevCategory
            Dim k As Types.KeyDevCategory
            If TypeOf c Is Types.KeyDevCategory Then
                k = DirectCast(c, Types.KeyDevCategory)
            ElseIf TypeOf c Is Integer Then
                k = DirectCast(c, Types.KeyDevCategory)
            Else
                Dim node As System.Xml.XmlNode = DirectCast(c, System.Xml.XmlNode())(1)
                k = DirectCast([Enum].Parse(GetType(Types.KeyDevCategory), node.Value), Types.KeyDevCategory)
            End If
            Return k
        End Function

        Private Shared Function GetKeyDevFlags(ByVal inCategoryArray As ArrayList) As Integer
            Dim flags As Integer = 0
            For Each c As Object In inCategoryArray
                Dim k As Types.KeyDevCategory = GetKeyDevCategory(c)

                If k = Types.KeyDevCategory.SECFilings Then
                    flags = flags Or 4
                Else
                    flags = flags Or 2
                End If
            Next

            Return flags
        End Function

        Private Shared Function GetKeyDevCategoryXML(ByVal inCategoryArray As ArrayList) As String
            Dim sb As New System.Text.StringBuilder
            With sb
                .Append("<itemlist>")
                For Each type As Object In inCategoryArray
                    Dim k As Integer = DirectCast(GetKeyDevCategory(type), Integer)
                    .Append("<item itemId=""").Append(k).Append("""/>")
                Next
                .Append("</itemlist>")
            End With

            Return sb.ToString()

        End Function

#Region "Search"
        <WebMethod(Description:="Get a SeriesComponent array of tradingItems that match the search text", EnableSession:=True), _
            XmlInclude(GetType(CapitalIQ.Charting.Series.SeriesComponent)), _
            XmlInclude(GetType(CapitalIQ.Charting.Series.SeriesComponent.EstimateInfo)), _
            XmlInclude(GetType(CapitalIQ.Charting.Common.EstimatesRegion))> _
        Public Function SearchTradingItems(ByVal searchText As String, ByVal inPrimaryOnly As Boolean) As SeriesComponent()

            Dim ret As SeriesComponent() = SearchMethods.SearchTradingItems(searchText, inPrimaryOnly)
            AddPlatformDefaultEstimateInfo(ret)
            Return ret
        End Function

        ''' <summary>
        ''' The SeriesComponent EstimateInfo array list doesn't know about the 
        ''' PlatformDefault setting in excel.  So loop through all the components
        ''' and copy the info that corresponds to the PlatformDefault setting into
        ''' a new estimate info set to VendorId -1.
        ''' </summary>
        ''' <param name="items"></param>
        ''' <remarks></remarks>
        Private Shared Sub AddPlatformDefaultEstimateInfo(ByVal items As SeriesComponent())
            ' PlatformDefault
            Dim userDefault As DataVendorId
            userDefault = DirectCast(CapitalIQ.SessionState.UserSession.Current.UserRole.UserPreference.EstimatesDataVendorId(), DataVendorId)
            Dim ei As SeriesComponent.EstimateInfo
            For Each sc As SeriesComponent In items
                ei = sc.GetEstimateInfo(userDefault)
                ' in the plug-in PlatformDefault = -1
                sc.AddEstimateInfo(-1, ei.HasEstimates, ei.RegionId)
            Next
        End Sub

        <WebMethod(Description:="Get a SeriesComponent array of tradingItems that match the search text using the ExcelSearchCompanyNames_tbl", EnableSession:=True), _
            XmlInclude(GetType(CapitalIQ.Charting.Series.SeriesComponent)), _
            XmlInclude(GetType(CapitalIQ.Charting.Series.SeriesComponent.EstimateInfo)), _
            XmlInclude(GetType(CapitalIQ.Charting.Common.EstimatesRegion))> _
        Public Function QuickCompanySearch(ByVal searchText As String) As SeriesComponent()


            Dim ret As SeriesComponent() = SearchMethods.QuickCompanySearch(SearchMethods.GetExcelTerm(searchText))
            AddPlatformDefaultEstimateInfo(ret)
            Return ret
        End Function

        <WebMethod(Description:="Get a SeriesComponent array of companies that match the search text", EnableSession:=True), _
            XmlInclude(GetType(CapitalIQ.Charting.Series.SeriesComponent)), _
            XmlInclude(GetType(CapitalIQ.Charting.Series.SeriesComponent.EstimateInfo)), _
            XmlInclude(GetType(CapitalIQ.Charting.Common.EstimatesRegion))> _
        Public Function SearchCompanies(ByVal searchText As String) As SeriesComponent()


            Dim ret As SeriesComponent() = SearchMethods.SearchCompanies(searchText)
            AddPlatformDefaultEstimateInfo(ret)
            Return ret
        End Function

        <WebMethod(Description:="Get a SeriesComponent array of indices that match the search text", EnableSession:=True), _
            XmlInclude(GetType(CapitalIQ.Charting.Series.SeriesComponent)), _
            XmlInclude(GetType(CapitalIQ.Charting.Series.SeriesComponent.EstimateInfo)), _
            XmlInclude(GetType(CapitalIQ.Charting.Common.EstimatesRegion))> _
        Public Function SearchMarketIndices(ByVal searchText As String) As SeriesComponent()


            Dim ret As SeriesComponent() = SearchMethods.SearchMarketIndices(searchText)
            AddPlatformDefaultEstimateInfo(ret)
            Return ret
        End Function

#End Region

        ''' <summary>
        ''' Sets the DateTimeMode to Unspecified for any date columns in the dataset
        ''' This fixes an issue where dates are assumed to be localtime in webservices
        ''' so 1/1/2007 becomes 12/31/2006 9:00 pm on the west coast.
        ''' </summary>
        Private Sub FixDates(ByVal ds As DataSet)
            If ds Is Nothing Then
                Return
            End If
            For Each t As DataTable In ds.Tables
                For Each c As DataColumn In t.Columns
                    If c.DataType Is GetType(System.DateTime) Then
                        c.DateTimeMode = DataSetDateTime.Unspecified
                    End If
                Next
            Next
        End Sub


    End Class
End Namespace
