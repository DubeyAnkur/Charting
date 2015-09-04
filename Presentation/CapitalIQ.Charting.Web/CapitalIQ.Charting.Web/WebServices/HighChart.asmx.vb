Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports ccc = CapitalIQ.Charting.Common
Imports CapitalIQ.Constants.ActivityTypeId

Imports CapitalIQ.Charting
Imports CapitalIQ.Constants
Imports System.Web.Script.Serialization
Imports System.Xml
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.IO
Imports System.Collections.Generic
Imports Newtonsoft.Json.Serialization

Namespace CapitalIQ.Charting.Web.WebServices
    <WebService(Namespace:="http://www.capitaliq.com/ciqdotnet/Charting/HighChart.asmx")> _
    <WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    <System.Web.Script.Services.ScriptService()> _
    Public Class HighChart
        Inherits System.Web.Services.WebService

#Region "AJAX Methods"

        Private Function GetChartingBusinessMgr() As IChartingBusinessMgr
            Return DirectCast(RemoteUtil.GetRemote(GetType(IChartingBusinessMgr)), IChartingBusinessMgr)
        End Function

        <WebMethod(Description:="Save PCXML to DB", EnableSession:=True)> _
        Public Function SavePCXML(pcXML As String) As String
            Dim mgr As IChartingBusinessMgr = GetChartingBusinessMgr()

            Dim g As String
            g = Guid.NewGuid().ToString()
            mgr.SavePCXML(g, pcXML)
            Return g
        End Function

        <WebMethod(Description:="Get PCXML for givengUID", EnableSession:=True)> _
        Public Function GetPCXMLsJSON(GUID As String) As String
            Dim mgr As IChartingBusinessMgr = GetChartingBusinessMgr()
            Dim pcXML As String = mgr.GetPCXML(GUID)

            Dim Options As String = ""
            Options = ConvertXMLtoJSON(pcXML)

            Return Options
        End Function
        'strXML = EscapeXMLValue(strXML)

        Private Function ConvertXMLtoJSON(strXML As String) As String
            Dim chartModel As New ChartModel()

            strXML = EscapeXMLValue(strXML)
            Dim document As New XmlDocument()
            document.LoadXml(strXML)

            For Each node As XmlNode In document.GetElementsByTagName("Chart")
                For Each node1 As XmlNode In node.ChildNodes
                    If node1.Name = "Graph" Then
                        GetChartModel(node1, chartModel)
                    ElseIf node1.Name = "GraphData" Then
                        Dim chartSeries As New List(Of ChartSeries)()
                        Dim dataPoints As New List(Of DataPoint)()
                        For Each node2 As XmlNode In node1
                            If node2.Name = "Categories" AndAlso node2.ChildNodes.Count > 0 Then
                                chartModel.Categories = GetCategories(node2)
                            ElseIf node2.Name = "Series" Then
                                If Not String.Equals(chartModel.Type, "pie", StringComparison.InvariantCultureIgnoreCase) Then
                                    chartSeries.Add(GetChartSeries(node2, chartModel))
                                Else
                                    dataPoints.AddRange(GetPieChartSeries(node2))
                                End If
                            End If
                        Next

                        If String.Equals(chartModel.Type, "pie", StringComparison.InvariantCultureIgnoreCase) Then
                            chartSeries.Add(New ChartSeries() With {.DataPoints = dataPoints, .Type = "pie", .PieInnerSize = "45%", .PieShowInLegend = True})
                        End If

                        chartModel.ChartSeries = chartSeries
                    End If
                Next
            Next

            Dim result As String = JsonConvert.SerializeObject(chartModel, Newtonsoft.Json.Formatting.Indented, New JsonSerializerSettings() With {.ContractResolver = New CamelCasePropertyNamesContractResolver(), .NullValueHandling = NullValueHandling.Ignore})

            Return result
        End Function

        Private Function GetCategories(node2 As XmlNode) As List(Of String)
            Dim chartCategories As New List(Of String)()
            For Each node3 As XmlNode In node2.ChildNodes
                If node3.Name = "Category" Then
                    If node3.Attributes IsNot Nothing Then
                        chartCategories.Add(node3.Attributes("Name").Value)
                    End If
                End If
            Next

            Return chartCategories
        End Function

        Private Function GetPieChartSeries(node2 As XmlNode) As List(Of DataPoint)
            Dim dataPoints As New List(Of DataPoint)()

            Dim dataPoint As New DataPoint()

            If node2.Attributes IsNot Nothing Then
                dataPoint.Name = node2.Attributes("Name").Value
            End If

            For Each node3 As XmlNode In node2.ChildNodes
                If node3.Name = "Data" AndAlso node3.Attributes IsNot Nothing Then
                    For Each attr As XmlAttribute In node3.Attributes
                        If attr.Name = "Date" Then
                            dataPoint.Xpoint = CLng(Convert.ToDateTime(attr.Value).ToUniversalTime().Subtract(New DateTime(1970, 1, 1, 0, 0, 0, _
                                            DateTimeKind.Utc)).TotalMilliseconds)
                        ElseIf attr.Name = "Value" Then
                            dataPoint.Ypoint = If(attr.Value = String.Empty, CType(Nothing, System.Nullable(Of Decimal)), Convert.ToDecimal(attr.Value))
                        ElseIf attr.Name = "High" Then
                            dataPoint.High = Convert.ToDecimal(attr.Value)
                        ElseIf attr.Name = "Low" Then
                            dataPoint.Low = Convert.ToDecimal(attr.Value)
                        ElseIf attr.Name = "Open" Then
                            dataPoint.Open = Convert.ToDecimal(attr.Value)
                        ElseIf attr.Name = "Close" Then
                            dataPoint.Close = Convert.ToDecimal(attr.Value)
                        ElseIf attr.Name = "Popup" Then
                            dataPoint.ToolTip = attr.Value
                        ElseIf attr.Name = "Name" Then
                            dataPoint.Name = attr.Value
                        ElseIf attr.Name = "Bubble" Then
                            dataPoint.Bubble = attr.Value
                        ElseIf attr.Name = "DrilldownUrl" Then
                            dataPoint.DrilldownUrl = attr.Value
                        End If
                    Next

                    dataPoints.Add(dataPoint)
                End If
            Next

            Return dataPoints
        End Function

        Private Function GetChartSeries(node2 As XmlNode, chartModel As ChartModel) As ChartSeries
            Dim chartCategories As List(Of String) = chartModel.Categories
            Dim series As New ChartSeries()
            Dim dataPoints As New List(Of DataPoint)()

            series.Type = chartModel.Type.ToLower()

            If node2.Attributes IsNot Nothing Then
                series.Name = node2.Attributes("Name").Value
            End If

            For Each node3 As XmlNode In node2.ChildNodes
                If node3.Name = "Data" AndAlso node3.Attributes IsNot Nothing Then
                    Dim dataPoint As New DataPoint()
                    For Each attr As XmlAttribute In node3.Attributes
                        If attr.Name = "Date" AndAlso chartCategories Is Nothing Then
                            dataPoint.Xpoint = CLng(Convert.ToDateTime(attr.Value).ToUniversalTime().Subtract(New DateTime(1970, 1, 1, 0, 0, 0, _
                                            DateTimeKind.Utc)).TotalMilliseconds)
                        ElseIf attr.Name = "Value" Then
                            dataPoint.Ypoint = If(attr.Value = String.Empty, CType(Nothing, System.Nullable(Of Decimal)), Convert.ToDecimal(attr.Value))
                        ElseIf attr.Name = "High" Then
                            dataPoint.High = Convert.ToDecimal(attr.Value)
                        ElseIf attr.Name = "Low" Then
                            dataPoint.Low = Convert.ToDecimal(attr.Value)
                        ElseIf attr.Name = "Open" Then
                            dataPoint.Open = Convert.ToDecimal(attr.Value)
                        ElseIf attr.Name = "Close" Then
                            dataPoint.Close = Convert.ToDecimal(attr.Value)
                        ElseIf attr.Name = "Popup" Then
                            dataPoint.ToolTip = attr.Value
                        ElseIf attr.Name = "Name" Then
                            dataPoint.Name = attr.Value
                        ElseIf attr.Name = "Bubble" Then
                            dataPoint.Bubble = attr.Value
                        ElseIf attr.Name = "DrilldownUrl" Then
                            dataPoint.DrilldownUrl = attr.Value
                        End If
                    Next

                    dataPoints.Add(dataPoint)
                End If
            Next

            series.DataPoints = dataPoints
            Return series
        End Function

        Private Sub GetChartModel(node1 As XmlNode, chartModel As ChartModel)
            If node1.Attributes IsNot Nothing Then
                For Each attr As XmlAttribute In node1.Attributes
                    If attr.Name = "Type" Then
                        If String.Equals(attr.Value, "bar", StringComparison.OrdinalIgnoreCase) Then
                            chartModel.Type = "column"
                        ElseIf String.Equals(attr.Value, "time", StringComparison.OrdinalIgnoreCase) Then
                            chartModel.Type = "line"
                        Else
                            chartModel.Type = attr.Value.ToLower()
                        End If
                    ElseIf attr.Name = "SubType" Then
                        chartModel.SubType = attr.Value
                    ElseIf attr.Name = "Anchor" Then
                        chartModel.Anchor = attr.Value
                    ElseIf attr.Name = "Name" Then
                        chartModel.Name = attr.Value
                    ElseIf attr.Name = "Left" Then
                        chartModel.Left = Convert.ToInt32(attr.Value)
                    ElseIf attr.Name = "Width" Then
                        chartModel.Width = Convert.ToInt32(attr.Value)
                    ElseIf attr.Name = "Top" Then
                        chartModel.Top = Convert.ToInt32(attr.Value)
                    ElseIf attr.Name = "Height" Then
                        chartModel.Height = Convert.ToInt32(attr.Value)
                    End If
                Next
            End If

            ' chartModel.ChartPlotOptions = GetChartPlotOptions(node1);
        End Sub



        Private Function UnescapeXMLValue(xmlString As String) As String
            If xmlString Is DBNull.Value Then
                Throw New ArgumentNullException("xmlString")
            End If

            Return xmlString.Replace("&apos;", "'").Replace("&quot;", "\""").Replace("&gt;", ">").Replace("&lt;", "<").Replace("&amp;", "&")
        End Function


        Private Function EscapeXMLValue(xmlString As String) As String
            If xmlString Is DBNull.Value Then
                Throw New ArgumentNullException("xmlString")
            End If

            Return xmlString.Replace("&", "&amp;") ' .Replace("'", "&apos;").Replace("\""", "&quot;").Replace(">", "&gt;").Replace("<", "&lt;")
        End Function
#End Region
    End Class


#Region "Models"
    <JsonObject(MemberSerialization.OptIn)> _
    Public Class ChartModel
        <JsonProperty(PropertyName:="type")> _
        Public Property Type() As String
            Get
                Return m_Type
            End Get
            Set(value As String)
                m_Type = value
            End Set
        End Property
        Private m_Type As String

        <JsonProperty(PropertyName:="subType")> _
        Public Property SubType() As String
            Get
                Return m_SubType
            End Get
            Set(value As String)
                m_SubType = value
            End Set
        End Property
        Private m_SubType As String

        <JsonProperty(PropertyName:="anchor")> _
        Public Property Anchor() As String
            Get
                Return m_Anchor
            End Get
            Set(value As String)
                m_Anchor = value
            End Set
        End Property
        Private m_Anchor As String

        <JsonProperty(PropertyName:="name")> _
        Public Property Name() As String
            Get
                Return m_Name
            End Get
            Set(value As String)
                m_Name = value
            End Set
        End Property
        Private m_Name As String

        <JsonProperty(PropertyName:="left")> _
        Public Property Left() As Integer
            Get
                Return m_Left
            End Get
            Set(value As Integer)
                m_Left = value
            End Set
        End Property
        Private m_Left As Integer

        <JsonProperty(PropertyName:="width")> _
        Public Property Width() As Integer
            Get
                Return m_Width
            End Get
            Set(value As Integer)
                m_Width = value
            End Set
        End Property
        Private m_Width As Integer

        <JsonProperty(PropertyName:="top")> _
        Public Property Top() As Integer
            Get
                Return m_Top
            End Get
            Set(value As Integer)
                m_Top = value
            End Set
        End Property
        Private m_Top As Integer

        <JsonProperty(PropertyName:="height")> _
        Public Property Height() As Integer
            Get
                Return m_Height
            End Get
            Set(value As Integer)
                m_Height = value
            End Set
        End Property
        Private m_Height As Integer

        <JsonProperty(PropertyName:="categories")> _
        Public Property Categories() As List(Of String)
            Get
                Return m_Categories
            End Get
            Set(value As List(Of String))
                m_Categories = value
            End Set
        End Property
        Private m_Categories As List(Of String)

        '[JsonProperty(PropertyName = "chartPlotOptions")]
        'public ChartPlotOptions ChartPlotOptions { get; set; }

        <JsonProperty(PropertyName:="series")> _
        Public Property ChartSeries() As List(Of ChartSeries)
            Get
                Return m_ChartSeries
            End Get
            Set(value As List(Of ChartSeries))
                m_ChartSeries = value
            End Set
        End Property
        Private m_ChartSeries As List(Of ChartSeries)
    End Class

    Public Class ChartSeries
        <JsonProperty(PropertyName:="type")> _
        Public Property Type() As String
            Get
                Return m_Type
            End Get
            Set(value As String)
                m_Type = value
            End Set
        End Property
        Private m_Type As String

        <JsonProperty(PropertyName:="innerSize")> _
        Public Property PieInnerSize() As String
            Get
                Return m_PieInnerSize
            End Get
            Set(value As String)
                m_PieInnerSize = value
            End Set
        End Property
        Private m_PieInnerSize As String

        <JsonProperty(PropertyName:="showInLegend")> _
        Public Property PieShowInLegend() As System.Nullable(Of Boolean)
            Get
                Return m_PieShowInLegend
            End Get
            Set(value As System.Nullable(Of Boolean))
                m_PieShowInLegend = value
            End Set
        End Property
        Private m_PieShowInLegend As System.Nullable(Of Boolean)

        <JsonProperty(PropertyName:="turboThreshold")> _
        Public ReadOnly Property TurboThreshold() As System.Nullable(Of Integer)
            Get
                Return 0
            End Get
        End Property

        <JsonProperty(PropertyName:="name")> _
        Public Property Name() As String
            Get
                Return m_Name
            End Get
            Set(value As String)
                m_Name = value
            End Set
        End Property
        Private m_Name As String

        <JsonProperty(PropertyName:="data")> _
        Public Property DataPoints() As List(Of DataPoint)
            Get
                Return m_DataPoints
            End Get
            Set(value As List(Of DataPoint))
                m_DataPoints = value
            End Set
        End Property
        Private m_DataPoints As List(Of DataPoint)

    End Class

    Public Class DataPoint
        <JsonProperty(PropertyName:="x")> _
        Public Property Xpoint() As System.Nullable(Of Long)
            Get
                Return m_Xpoint
            End Get
            Set(value As System.Nullable(Of Long))
                m_Xpoint = value
            End Set
        End Property
        Private m_Xpoint As System.Nullable(Of Long)

        <JsonProperty(PropertyName:="y")> _
        Public Property Ypoint() As System.Nullable(Of Decimal)
            Get
                Return m_Ypoint
            End Get
            Set(value As System.Nullable(Of Decimal))
                m_Ypoint = value
            End Set
        End Property
        Private m_Ypoint As System.Nullable(Of Decimal)

        <JsonProperty(PropertyName:="name")> _
        Public Property Name() As String
            Get
                Return m_Name
            End Get
            Set(value As String)
                m_Name = value
            End Set
        End Property
        Private m_Name As String

        <JsonProperty(PropertyName:="bubble")> _
        Public Property Bubble() As String
            Get
                Return m_Bubble
            End Get
            Set(value As String)
                m_Bubble = value
            End Set
        End Property
        Private m_Bubble As String

        <JsonProperty(PropertyName:="drilldownUrl")> _
        Public Property DrilldownUrl() As String
            Get
                Return m_DrilldownUrl
            End Get
            Set(value As String)
                m_DrilldownUrl = value
            End Set
        End Property
        Private m_DrilldownUrl As String

        <JsonProperty(PropertyName:="tooltip")> _
        Public Property ToolTip() As String
            Get
                Return m_ToolTip
            End Get
            Set(value As String)
                m_ToolTip = value
            End Set
        End Property
        Private m_ToolTip As String

        <JsonProperty(PropertyName:="high")> _
        Public Property High() As System.Nullable(Of Decimal)
            Get
                Return m_High
            End Get
            Set(value As System.Nullable(Of Decimal))
                m_High = value
            End Set
        End Property
        Private m_High As System.Nullable(Of Decimal)

        <JsonProperty(PropertyName:="low")> _
        Public Property Low() As System.Nullable(Of Decimal)
            Get
                Return m_Low
            End Get
            Set(value As System.Nullable(Of Decimal))
                m_Low = value
            End Set
        End Property
        Private m_Low As System.Nullable(Of Decimal)

        <JsonProperty(PropertyName:="open")> _
        Public Property Open() As System.Nullable(Of Decimal)
            Get
                Return m_Open
            End Get
            Set(value As System.Nullable(Of Decimal))
                m_Open = value
            End Set
        End Property
        Private m_Open As System.Nullable(Of Decimal)

        <JsonProperty(PropertyName:="close")> _
        Public Property Close() As System.Nullable(Of Decimal)
            Get
                Return m_Close
            End Get
            Set(value As System.Nullable(Of Decimal))
                m_Close = value
            End Set
        End Property
        Private m_Close As System.Nullable(Of Decimal)
    End Class

#End Region
End Namespace

