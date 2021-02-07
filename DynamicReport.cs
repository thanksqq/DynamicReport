/*
 * Copyright 2014
 * Author: Quyan
 * Date: 2014-05-02
 * Version: 0.3 
 * 更新历史
 * 2014-5-19 v0.3 添加表格内分组
 * 2014-5-28 v0.4 增加参数类，开始做图表显示功能
 */
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using Microsoft.Reporting.WebForms;

namespace Cloud.Utility
{
    public interface IDynamicReport
    {
        void SetReport(ReportViewer reportViewer);
        void AddData(DataTable dataTable, string groupName = null);
        void AddData<T>(IEnumerable<T> data, string groupName = null);
        void ShowPager(bool showPager);
        void AddData(ReportDataParameter reportDataParameter);
        void ShowReport();
        void LoadReport(string reportPath);
        void SetColoumStyle(IEnumerable<ReportColoumStyle> coloumStyle);
        void AddText(string text);
    }

    public class ReportColoumStyle
    {
        public string ColoumName { get; set; }
        public float ColoumWidth { get; set; }
        public TextAlign TextAlign { get; set; }
        public ReportColoumHierarchy ReportColoumHierarchy { get; set; }

        public ReportColoumStyle()
        {
            ColoumWidth = DynamicReport.ColoumWidth;
        }
    }

    public class ReportDataParameter
    {
        public ReportDataParameter()
        {
            TableOrChart = true;
        }
        public dynamic Data { get; set; }
        public bool TableOrChart { get; set; }
        public ChartType ChartType { get; set; }
        public ChartSubtype ChartSubtype { get; set; }
        public float? ChartHeight { get; set; }
        public bool Enabled3D { get; set; }
        public string GroupedColoumName { get; set; }
        public IEnumerable<ReportColoumStyle> ReportColoumStyles { get; set; }
        public string Title { get; set; }
    }

    public enum TextAlign
    {
        Left,
        Center,
        Right
    }

    public enum ReportType
    {
        Tables,
        Chart,
        Finally
    }

    public enum ReportColoumHierarchy
    {
        CategoryHierarchy,
        SeriesHierarchy,
        DataHierarchy,
        DataLabelHierarchy
    }

    public enum ChartType
    {
        /// <summary>
        /// 柱状图
        /// </summary>
        Column,
        /// <summary>
        /// 折线图
        /// </summary>
        Line,
        /// <summary>
        /// 饼图
        /// </summary>
        Shape,
        /// <summary>
        /// 条状图
        /// </summary>
        Bar,
    }
    public enum ChartSubtype
    {
        /// <summary>
        /// 平面（默认的）
        /// </summary>
        Plain,
        /// <summary>
        /// 平滑的
        /// </summary>
        Smooth,
        /// <summary>
        /// 甜甜圈
        /// </summary>
        ExplodedDoughnut,
    }


    internal class ReportItemPattern
    {
        public string DataSetName { get; set; }
        public string DataSetString { get; set; }
        public string TablixString { get; set; }
        public dynamic Data { get; set; }
        public ReportType ReportType { get; set; }
        public IEnumerable<ReportColoumStyle> ReportColoumStyles { get; set; }
    }

    /// <summary>
    /// v0.4版
    /// </summary>
    public class DynamicReport : IDynamicReport
    {
        #region 空白文档

        #region 空白文档代码
        /// <summary>
        /// 空白文档的xml文件
        /// </summary>
        private string _docTemplate =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?><Report xmlns:rd=\"http://schemas.microsoft.com/SQLServer/reporting/reportdesigner\" xmlns=\"http://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition\">" +
            "<DataSources>" +
            "   <DataSource Name=\"DummyDataSource\">" +
            "       <ConnectionProperties>" +
            "           <DataProvider>SQL</DataProvider>" +
            "           <ConnectString />" +
            "       </ConnectionProperties>" +
            "       <rd:DataSourceID>3eecdab9-6b4b-4836-ad62-95e4aee65ea8</rd:DataSourceID>" +
            "   </DataSource>" +
            "</DataSources>" +
            "<DataSets>@DataSets</DataSets>" +
            "<Body>" +
            "<ReportItems>@Title@Tablix" +
            "</ReportItems>" +
            "<Style />" +
            "<Height>8cm</Height>" +
            "</Body>" +
            "<Width>17cm</Width>" +
            "<Page>" +
            "<PageHeight>29.7cm</PageHeight>" +
            "<PageWidth>21cm</PageWidth>" +
            "@InteractivePage" +
            "<LeftMargin>1.8cm</LeftMargin>" +
            "<RightMargin>1.8cm</RightMargin>" +
            "<TopMargin>1.8cm</TopMargin>" +
            "<BottomMargin>1.8cm</BottomMargin>" +
            "<ColumnSpacing>0.13cm</ColumnSpacing>" +
            "<Style />" +
            "</Page>" +
            "<rd:ReportID>809f16cf-ea78-4469-bf43-965c4afe69d0</rd:ReportID>" +
            "<rd:ReportUnitType>Cm</rd:ReportUnitType>" +
            "</Report>";
        #endregion

        #region 文字标签
        private const string TitlePattern = " <Textbox Name=\"Textbo@TextboxName\"> " + @"<CanGrow>true</CanGrow>
        <KeepTogether>true</KeepTogether>
        <Paragraphs>
          <Paragraph>
            <TextRuns>
              <TextRun>
                <Value>@Title</Value>
                <Style>@FontStyle</Style>
              </TextRun>
            </TextRuns>
            <Style>@Style</Style>
          </Paragraph>
        </Paragraphs>
        <rd:DefaultName>Textbo@TextboxName</rd:DefaultName>
        <Top>@TopPositioncm</Top>
        <Left>1cm</Left>
        <Height>0.83813cm</Height>
        <Width>14.35207cm</Width>
        <ZIndex>1</ZIndex>
        <Style>
          <Border>
            <Style>None</Style>
          </Border>
          <PaddingLeft>2pt</PaddingLeft>
          <PaddingRight>2pt</PaddingRight>
          <PaddingTop>2pt</PaddingTop>
          <PaddingBottom>2pt</PaddingBottom>
        </Style>
      </Textbox>";
        #endregion

        #region DataSet的代码段
        private static string DataSetPattern
        {
            get
            {
                return "    <DataSet Name=\"@DataSetNameData\">" +
                       "       <Fields>@Fields</Fields>" +
                       "       <Query>" +
                       "           <DataSourceName>DummyDataSource</DataSourceName>" +
                       "           <CommandText />" +
                       "       </Query>" +
                       "    </DataSet>";
            }
        }
        #endregion

        #region 表格的代码段
        private string TablixPattern
        {
            get
            {
                return " <Tablix Name=\"Tablix@DataSetName\">" +
                       "   <TablixBody>" +
                       "       <TablixColumns>@TablixColumns</TablixColumns>" +
                       "       <TablixRows>" +
                       "           <TablixRow>" +
                       "               <Height>0.23622in</Height>" +
                       "               <TablixCells>@TablixHeader</TablixCells>" +
                       "           </TablixRow>" +
                       "           <TablixRow>" +
                       "               <Height>0.23622in</Height>" +
                       "               <TablixCells>@TablixCells</TablixCells>" +
                       "           </TablixRow>" +
                       "       </TablixRows>" +
                       "   </TablixBody>" +
                       "   <TablixColumnHierarchy>" +
                       "       <TablixMembers>@TablixMember</TablixMembers>" +
                       "   </TablixColumnHierarchy>" +
                       "   <TablixRowHierarchy>" +
                       "       <TablixMembers>" +
                       "           <TablixMember>@GroupHeader" +
                       "               <KeepWithGroup>After</KeepWithGroup>" +
                       "           </TablixMember>" +
                       "           <TablixMember>@GroupMember" +
                       "               <TablixMembers>" +
                       "                   <TablixMember>" +
                       "                        <Group Name=\"详细信息@DataSetName\" />" +
                       "                    </TablixMember>" +
                       "                </TablixMembers>" +
                       "           </TablixMember>" +
                       "       </TablixMembers>" +
                       "   </TablixRowHierarchy>" +
                       "   <DataSetName>@DataSetNameData</DataSetName>" +
                       "   <Top>@TopPositioncm</Top>" +
                       "   <Left>@LeftPostioncm</Left>" +
                       "   <Height>1.2cm</Height>" +
                       "   <Width>14.35207cm</Width>" +
                       "   <Style>" +
                       "       <Border>" +
                       "           <Style>None</Style>" +
                       "       </Border>" +
                       "   </Style>" +
                       "</Tablix>";
            }
        }
        #endregion

        #region 图形的代码段
        private const string ChartPattern =
            "      <Chart Name=\"Chart@DataSetName\" > "
            + "        <ChartCategoryHierarchy>"
            + "          <ChartMembers>"
            + "            <ChartMember>"
            + "              <Group Name=\"Chart@DataSetName_CategoryGroup1\">"
            + "                <GroupExpressions>"
            + "                  <GroupExpression>=Fields!@CategoryHierarchy.Value</GroupExpression>"
            + "                </GroupExpressions>"
            + "              </Group>"
            + "              <Label>=Fields!@CategoryHierarchy.Value</Label>"
            + "            </ChartMember>"
            + "          </ChartMembers>"
            + "        </ChartCategoryHierarchy>"
            + "        <ChartSeriesHierarchy>"
            + "          <ChartMembers>"
            + "            <ChartMember>@SeriesHierarchy"
            + "            </ChartMember>"
            + "          </ChartMembers>"
            + "        </ChartSeriesHierarchy>"
            + "        <ChartData>"
            + "          <ChartSeriesCollection>"
            + "            <ChartSeries Name=\"@DataHierarchy\">"
            + "              <ChartDataPoints>"
            + "                <ChartDataPoint>"
            + "                  <ChartDataPointValues>"
            + "                    <Y>=Fields!@DataHierarchy.Value</Y>"
            + "                  </ChartDataPointValues>"
            + "                  <ChartDataLabel>"
            + "                    <Style />@DataLabelHierarchy"
            + "                  </ChartDataLabel>"
            + "                  <Style />"
            + "                  <ChartMarker>"
            + "                    <Style />"
            + "                  </ChartMarker>"
            + "                  <DataElementOutput>Output</DataElementOutput>"
            + "                </ChartDataPoint>"
            + "              </ChartDataPoints>@ChartType"
            + "              <Style />"
            + "              <ChartEmptyPoints>"
            + "                <Style />"
            + "                <ChartMarker>"
            + "                  <Style />"
            + "                </ChartMarker>"
            + "                <ChartDataLabel>"
            + "                  <Style />"
            + "                </ChartDataLabel>"
            + "              </ChartEmptyPoints>"
            + "              <ValueAxisName>Primary</ValueAxisName>"
            + "              <CategoryAxisName>Primary</CategoryAxisName>"
            + "              <ChartSmartLabel>"
            + "                <CalloutLineColor>Black</CalloutLineColor>"
            + "                <MinMovingDistance>0pt</MinMovingDistance>"
            + "              </ChartSmartLabel>"
            + "            </ChartSeries>"
            + "          </ChartSeriesCollection>"
            + "        </ChartData>"
            + "        <ChartAreas>"
            + "          <ChartArea Name=\"Default\">"
            + "            <ChartCategoryAxes>"
            + "              <ChartAxis Name=\"Primary\">"
            + "                <Style>"
            + "                  <FontSize>8pt</FontSize>"
            + "                </Style>"
            + "                <ChartAxisTitle>"
            + "                  <Caption />"
            + "                  <Style>"
            + "                    <FontSize>8pt</FontSize>"
            + "                  </Style>"
            + "                </ChartAxisTitle>"
            + "                <Interval>1</Interval>"
            + "                <ChartMajorGridLines>"
            + "                  <Enabled>False</Enabled>"
            + "                  <Style>"
            + "                    <Border>"
            + "                      <Color>Gainsboro</Color>"
            + "                    </Border>"
            + "                  </Style>"
            + "                </ChartMajorGridLines>"
            + "                <ChartMinorGridLines>"
            + "                  <Style>"
            + "                    <Border>"
            + "                      <Color>Gainsboro</Color>"
            + "                      <Style>Dotted</Style>"
            + "                    </Border>"
            + "                  </Style>"
            + "                </ChartMinorGridLines>"
            + "                <ChartMinorTickMarks>"
            + "                  <Length>0.5</Length>"
            + "                </ChartMinorTickMarks>"
            + "                <CrossAt>NaN</CrossAt>"
            + "                <Minimum>NaN</Minimum>"
            + "                <Maximum>NaN</Maximum>"
            + "              </ChartAxis>"
            + "              <ChartAxis Name=\"Secondary\">"
            + "                <Style>"
            + "                  <FontSize>8pt</FontSize>"
            + "                </Style>"
            + "                <ChartAxisTitle>"
            + "                  <Caption>轴标题</Caption>"
            + "                  <Style>"
            + "                    <FontSize>8pt</FontSize>"
            + "                  </Style>"
            + "                </ChartAxisTitle>"
            + "                <ChartMajorGridLines>"
            + "                  <Enabled>False</Enabled>"
            + "                  <Style>"
            + "                    <Border>"
            + "                      <Color>Gainsboro</Color>"
            + "                    </Border>"
            + "                  </Style>"
            + "                </ChartMajorGridLines>"
            + "                <ChartMinorGridLines>"
            + "                  <Style>"
            + "                    <Border>"
            + "                      <Color>Gainsboro</Color>"
            + "                      <Style>Dotted</Style>"
            + "                    </Border>"
            + "                  </Style>"
            + "                </ChartMinorGridLines>"
            + "                <ChartMinorTickMarks>"
            + "                  <Length>0.5</Length>"
            + "                </ChartMinorTickMarks>"
            + "                <CrossAt>NaN</CrossAt>"
            + "                <Location>Opposite</Location>"
            + "                <Minimum>NaN</Minimum>"
            + "                <Maximum>NaN</Maximum>"
            + "              </ChartAxis>"
            + "            </ChartCategoryAxes>"
            + "            <ChartValueAxes>"
            + "              <ChartAxis Name=\"Primary\">"
            + "                <Style>"
            + "                  <FontSize>8pt</FontSize>"
            + "                </Style>"
            + "                <ChartAxisTitle>"
            + "                  <Caption />"
            + "                  <Style>"
            + "                    <FontSize>8pt</FontSize>"
            + "                  </Style>"
            + "                </ChartAxisTitle>"
            + "                <ChartMajorGridLines>"
            + "                  <Style>"
            + "                    <Border>"
            + "                      <Color>Gainsboro</Color>"
            + "                    </Border>"
            + "                  </Style>"
            + "                </ChartMajorGridLines>"
            + "                <ChartMinorGridLines>"
            + "                  <Style>"
            + "                    <Border>"
            + "                      <Color>Gainsboro</Color>"
            + "                      <Style>Dotted</Style>"
            + "                    </Border>"
            + "                  </Style>"
            + "                </ChartMinorGridLines>"
            + "                <ChartMinorTickMarks>"
            + "                  <Length>0.5</Length>"
            + "                </ChartMinorTickMarks>"
            + "                <CrossAt>NaN</CrossAt>"
            + "                <Minimum>NaN</Minimum>"
            + "                <Maximum>NaN</Maximum>"
            + "              </ChartAxis>"
            + "              <ChartAxis Name=\"Secondary\">"
            + "                <Style>"
            + "                  <FontSize>8pt</FontSize>"
            + "                </Style>"
            + "                <ChartAxisTitle>"
            + "                  <Caption>轴标题</Caption>"
            + "                  <Style>"
            + "                    <FontSize>8pt</FontSize>"
            + "                  </Style>"
            + "                </ChartAxisTitle>"
            + "                <ChartMajorGridLines>"
            + "                  <Style>"
            + "                    <Border>"
            + "                      <Color>Gainsboro</Color>"
            + "                    </Border>"
            + "                  </Style>"
            + "                </ChartMajorGridLines>"
            + "                <ChartMinorGridLines>"
            + "                  <Style>"
            + "                    <Border>"
            + "                      <Color>Gainsboro</Color>"
            + "                      <Style>Dotted</Style>"
            + "                    </Border>"
            + "                  </Style>"
            + "                </ChartMinorGridLines>"
            + "                <ChartMinorTickMarks>"
            + "                  <Length>0.5</Length>"
            + "                </ChartMinorTickMarks>"
            + "                <CrossAt>NaN</CrossAt>"
            + "                <Location>Opposite</Location>"
            + "                <Minimum>NaN</Minimum>"
            + "                <Maximum>NaN</Maximum>"
            + "              </ChartAxis>"
            + "            </ChartValueAxes>"
            + "            <ChartThreeDProperties>"
            + "              <Enabled>@Enable3D</Enabled>"
            + "            </ChartThreeDProperties>"
            + "            <Style>"
            + "              <BackgroundGradientType>None</BackgroundGradientType>"
            + "            </Style>"
            + "          </ChartArea>"
            + "        </ChartAreas>"
            + "        <ChartLegends>"
            + "          <ChartLegend Name=\"ChartLegend@DataSetName\">"
            + "            <Style>"
            + "              <BackgroundGradientType>None</BackgroundGradientType>"
            + "              <FontSize>8pt</FontSize>"
            + "            </Style>"
            + "            <ChartLegendTitle>"
            + "              <Caption />"
            + "              <Style>"
            + "                <FontSize>8pt</FontSize>"
            + "                <FontWeight>Bold</FontWeight>"
            + "                <TextAlign>Center</TextAlign>"
            + "              </Style>"
            + "            </ChartLegendTitle>"
            + "            <HeaderSeparatorColor>Black</HeaderSeparatorColor>"
            + "            <ColumnSeparatorColor>Black</ColumnSeparatorColor>"
            + "          </ChartLegend>"
            + "        </ChartLegends>"
            + "        <ChartTitles>"
            + "          <ChartTitle Name=\"Default\">"
            + "            <Caption>@Title</Caption>"
            + "            <Style>"
            + "              <BackgroundGradientType>None</BackgroundGradientType>"
            + "              <FontWeight>Bold</FontWeight>"
            + "              <TextAlign>General</TextAlign>"
            + "              <VerticalAlign>Top</VerticalAlign>"
            + "            </Style>"
            + "          </ChartTitle>"
            + "        </ChartTitles>"
            + "        <Palette>BrightPastel</Palette>"
            + "        <ChartBorderSkin>"
            + "          <Style>"
            + "            <BackgroundColor>Gray</BackgroundColor>"
            + "            <BackgroundGradientType>None</BackgroundGradientType>"
            + "            <Color>White</Color>"
            + "          </Style>"
            + "        </ChartBorderSkin>"
            + "        <ChartNoDataMessage Name=\"NoDataMessage\">"
            + "          <Caption>没有可用数据</Caption>"
            + "          <Style>"
            + "            <BackgroundGradientType>None</BackgroundGradientType>"
            + "            <TextAlign>General</TextAlign>"
            + "            <VerticalAlign>Top</VerticalAlign>"
            + "          </Style>"
            + "        </ChartNoDataMessage>"
            + "        <DataSetName>@DataSetNameData</DataSetName>"
            + "        <Top>@TopPositioncm</Top>"
            + "        <Height>@ChartHeightcm</Height>"
            + "        <Width>18cm</Width>"
            + "        <Style>"
            + "          <Border>"
            + "            <Color>LightGrey</Color>"
            + "            <Style>Solid</Style>"
            + "          </Border>"
            + "          <BackgroundColor>White</BackgroundColor>"
            + "          <BackgroundGradientType>None</BackgroundGradientType>"
            + "        </Style>"
            + "      </Chart>";
        #endregion

        #endregion

        private ReportViewer _report;
        private IEnumerable<ReportColoumStyle> _coloumStyle;
        private readonly List<ReportItemPattern> _reportItemPatterns = new List<ReportItemPattern>();
        private readonly List<string> _reportTitlePatterns = new List<string>();
        private readonly List<ReportItemPattern> _reportHeadPatterns = new List<ReportItemPattern>();
        internal const float ColoumWidth = 1.6F; //行宽
        private bool _showPager = true;
        public ReportType ReportType { get; set; }

        /// <summary>
        /// 从现有报表中加载报表并进行修改
        /// </summary>
        /// <param name="url"></param>
        public void LoadReport(string url)
        {
            try
            {
                _docTemplate = File.ReadAllText(url);
            }
            catch (Exception ex)
            {

            }
        }

        public void SetReport(ReportViewer reportViewer)
        {
            this._report = reportViewer;
        }

        public void SetColoumStyle(IEnumerable<ReportColoumStyle> coloumStyle)
        {
            this._coloumStyle = coloumStyle;
        }

        public void AddText(string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                var pos = CaculatePlacePostion();
                var titlePattern = TitlePattern
                    .Replace("@Title", text)
                    .Replace("@TopPosition", pos.ToString())
                    .Replace("@TextboxName", _reportTitlePatterns.Count.ToString())
                    .Replace("@FontStyle", "<FontFamily>微软雅黑</FontFamily><FontSize>12pt</FontSize>")
                    .Replace("@Style", "<TextAlign>Center</TextAlign>");
                _reportTitlePatterns.Add(titlePattern);
            }
        }

        public void AddText(string text, int chapterGrade)
        {
            if (!string.IsNullOrEmpty(text))
            {
                var pos = CaculatePlacePostion();
                var titlePattern = TitlePattern
                    .Replace("@Title", text)
                    .Replace("@TopPosition", pos.ToString())
                    .Replace("@TextboxName", _reportTitlePatterns.Count.ToString());
                switch (chapterGrade)
                {
                    case 1:
                        titlePattern = titlePattern.Replace("@FontStyle",
                                                            "<FontFamily>宋体</FontFamily><FontSize>18pt</FontSize><Color>#000000</Color>")
                                                   .Replace("@Style", "<TextAlign>Center</TextAlign>");
                        break;
                    case 2:
                        titlePattern = titlePattern.Replace("@FontStyle",
                                                            "<FontFamily>黑体</FontFamily><FontSize>14pt</FontSize><Color>#000000</Color>")
                                                   .Replace("@Style", "<TextAlign>Left</TextAlign>");
                        break;
                    case 3:
                        titlePattern = titlePattern.Replace("@FontStyle",
                                                            "<FontFamily>宋体</FontFamily><FontSize>12pt</FontSize><FontWeight>Bold</FontWeight>")
                                                   .Replace("@Style", "<TextAlign>Left</TextAlign>");
                        break;
                    default:
                    case 10:
                        titlePattern = titlePattern.Replace("@FontStyle",
                                                            "<FontFamily>宋体</FontFamily><FontSize>12pt</FontSize>")
                                                   .Replace("@Style", "<LineHeight>22pt</LineHeight>");
                        break;
                }
                _reportTitlePatterns.Add(titlePattern);
            }
        }

        public void AddData(DataTable dataTable, string groupName = null)
        {
            if (dataTable != null)
            {
                var coloumNames = new List<string>();
                foreach (DataColumn dataColumn in dataTable.Columns)
                {
                    var protertyName = dataColumn.ColumnName;
                    coloumNames.Add(protertyName);
                }
                AddReportItemPattern(coloumNames.ToArray(), dataTable, groupName);
            }
        }

        public void AddData<T>(IEnumerable<T> data, string groupName = null)
        {
            if (data.Count() != 0)
            {
                var properites = typeof(T).GetProperties(); //得到实体类属性的集合
                AddReportItemPattern(properites.Select(p => p.Name).ToArray(), data, groupName);
            }
        }

        public void ShowPager(bool showPage)
        {
            _showPager = showPage;
        }

        /// <summary>
        /// 14-05-28新增，将参数做成一个集合，添加
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="reportDataParameter"></param>
        public void AddData(ReportDataParameter reportDataParameter)
        {
            //this._coloumStyle = reportDataParameter.ReportColoumStyles;
            if (reportDataParameter.TableOrChart) //说明是表格
            {
                if (reportDataParameter.Data != null)
                {
                    if (reportDataParameter.Data is DataTable)
                    {
                        AddData((DataTable)reportDataParameter.Data, reportDataParameter.GroupedColoumName);
                    }
                    else if (reportDataParameter.Data is IEnumerable)
                    {
                        AddData(Enumerable.AsEnumerable(reportDataParameter.Data), reportDataParameter.GroupedColoumName);
                    }
                }
            }
            else //说明是图形
            {
                //先取表头
                //var coloumNames = new List<string>();
                DataTable dataTable = null;

                if (reportDataParameter.Data != null)
                {
                    if (reportDataParameter.Data is DataTable)
                    {
                        dataTable = (DataTable)reportDataParameter.Data;
                    }
                    else if (reportDataParameter.Data is IEnumerable)
                    {
                        dataTable = ((IEnumerable)reportDataParameter.Data).Cast<object>().CopyToDataTable();
                    }
                    //coloumNames.AddRange(
                    //        dataTable.Columns.Cast<DataColumn>()
                    //                  .Select(dataColumn => dataColumn.ColumnName));
                }

                //这个是DataTable的定义
                var fields = new StringBuilder();
                //定义datasetName
                var dataSetName = string.Format("Data{0}", _reportItemPatterns.Count + _reportHeadPatterns.Count + 1);
                //定义整个chart的字符
                var chartString = ChartPattern;
                //定义DataHierarchy的名称
                var dataHierarchy = string.Empty;

                var dataFieldName = reportDataParameter.ReportColoumStyles.Select(r => r.ColoumName).Distinct();
                foreach (var name in dataFieldName)
                {
                    fields.AppendFormat(
                    "<Field Name=\"{0}\"><DataField>{0}</DataField><rd:TypeName>System.String</rd:TypeName></Field>",
                    name);
                }

                foreach (var coloumStyle in reportDataParameter.ReportColoumStyles)
                {

                    //取出列的样式
                    var reportColoumStyle = coloumStyle;
                    switch (reportColoumStyle.ReportColoumHierarchy)
                    {
                        case ReportColoumHierarchy.CategoryHierarchy:
                            //把主要轴的字符替换掉
                            chartString = chartString.Replace("@CategoryHierarchy", coloumStyle.ColoumName);
                            break;
                        case ReportColoumHierarchy.SeriesHierarchy:
                            chartString = chartString.Replace("@SeriesHierarchy",
                                                              string.Format(
                                                                  "<Group Name=\"Chart@DataSetName_SeriesGroup1\">"
                                                                  + "<GroupExpressions>"
                                                                  +
                                                                  "<GroupExpression>=Fields!{0}.Value</GroupExpression>"
                                                                  + "</GroupExpressions>"
                                                                  + "</Group>"
                                                                  + "<Label>=Fields!{0}.Value</Label>",
                                                                  coloumStyle.ColoumName));
                            break;
                        case ReportColoumHierarchy.DataHierarchy:
                            chartString = chartString.Replace("@DataHierarchy", coloumStyle.ColoumName);
                            dataHierarchy = coloumStyle.ColoumName;
                            break;
                        case ReportColoumHierarchy.DataLabelHierarchy:
                            //如果是横向的条形图，则不显示标签上的字
                            if (reportDataParameter.ChartType == ChartType.Bar)
                            {
                                chartString = chartString
                                    .Replace("@DataLabelHierarchy",
                                        string.Format(
                                            "<Label>=Fields!{0}.Value</Label><Visible>true</Visible>",
                                            coloumStyle.ColoumName));
                            }
                            else
                            {
                                chartString = chartString
                                    .Replace("@DataLabelHierarchy",
                                        string.Format(
                                            "<Label>=Fields!{1}.Value + \": \" + Fields!{0}.Value</Label><Visible>true</Visible>",
                                            coloumStyle.ColoumName,
                                            reportDataParameter.ReportColoumStyles.Single(
                                                s => s.ReportColoumHierarchy == ReportColoumHierarchy.CategoryHierarchy)
                                                .ColoumName));
                            }
                            break;
                    }
                }

                //ChartType
                if (!string.IsNullOrEmpty(reportDataParameter.ChartType.ToString()))
                {
                    chartString = chartString.Replace("@ChartType",
                        string.Format("<Type>{0}</Type><Subtype>{1}</Subtype>", reportDataParameter.ChartType, reportDataParameter.ChartSubtype));
                }

                //如果没有group，则把@SeriesHierarchy替换成name
                if (chartString.IndexOf("@SeriesHierarchy", StringComparison.Ordinal) > -1)
                {
                    chartString = chartString.Replace("@SeriesHierarchy", string.Format("<Label>{0}</Label>", dataHierarchy));
                }

                //如果没有数据标签，则把@DataLabelHierarchy替换成name
                if (chartString.IndexOf("@DataLabelHierarchy", StringComparison.Ordinal) > -1)
                {
                    chartString = chartString.Replace("@DataLabelHierarchy", "");
                }

                //设置图形表的高度（某些序列多的图形表比较有用）
                if (reportDataParameter.ChartHeight != null && reportDataParameter.ChartHeight > 0)
                {
                    chartString = chartString.Replace("@ChartHeight",
                        reportDataParameter.ChartHeight.ToString());
                }
                else
                {
                    chartString = chartString.Replace("@ChartHeight", "13");//默认13cm
                }

                //3D形状


                var reportItemPattern = new ReportItemPattern
                {
                    Data = dataTable,
                    DataSetName = dataSetName,
                    DataSetString = DataSetPattern
                        .Replace("@DataSetName", dataSetName)
                        .Replace("@Fields", fields.ToString()),
                    TablixString = chartString
                        .Replace("@DataSetName", dataSetName)
                        .Replace("@TopPosition", CaculatePlacePostion().ToString())
                        .Replace("@Title", reportDataParameter.Title)
                        .Replace("@Enable3D", reportDataParameter.Enabled3D.ToString().ToLower()),
                    ReportType = ReportType.Chart,
                };
                _reportItemPatterns.Add(reportItemPattern);
            }
        }

        /// <summary>
        /// 计算开始摆放的位置
        /// </summary>
        /// <returns></returns>
        protected float CaculatePlacePostion()
        {
            //每个标题的高度
            float titleCount = _reportTitlePatterns.Count * 1f;
            //每个数据表的高度
            float itemCount = _reportItemPatterns.Count(c => c.ReportType == ReportType.Tables) * 2f;
            //每个Chart的高度
            float chartCount = _reportItemPatterns.Count(c => c.ReportType == ReportType.Chart) * 13f;
            // 每个空表头的高度
            float emptyItemCount = _reportHeadPatterns.Count * 0.5f;
            switch (ReportType)
            {
                case ReportType.Chart:
                case ReportType.Tables:
                    return titleCount + itemCount + emptyItemCount + 0.5f;
                case ReportType.Finally:
                    return titleCount + itemCount + chartCount + emptyItemCount + 25.7f;
            }
            return 0f;
        }

        /// <summary>
        /// 增加一个报表
        /// </summary>
        /// <param name="coloumNames"></param>
        /// <param name="data"></param>
        /// <param name="groupName"></param>
        /// <param name="isTable"></param>
        protected void AddReportItemPattern(string[] coloumNames, dynamic data, string groupName = null, bool isTable = true)
        {
            var fields = new StringBuilder();
            var coloums = new StringBuilder();
            var tablixHearders = new StringBuilder();
            var tablixCells = new StringBuilder();
            var tablixMembers = new StringBuilder();
            var currentNamePrefix = _reportItemPatterns.Count + _reportHeadPatterns.Count + 1;
            var tableWidth = 0F;
            var dataRows = GetDataRowsCount(data); //数据行数
            var groupHeader = string.Empty;
            var groupMember = string.Empty;

            foreach (var coloumName in coloumNames)
            {
                var coloumValue = coloumName;
                //根据coloumStyle设定和计算列宽度
                var coloumWidth = ColoumWidth;
                var textAlign = TextAlign.Right;
                var reportColoumStyle = _coloumStyle.FirstOrDefault(r => r.ColoumName == coloumName);
                if (reportColoumStyle != null)
                {
                    textAlign = reportColoumStyle.TextAlign;
                    coloumWidth = reportColoumStyle.ColoumWidth;
                }
                //例外，如果字段开始为__x,则该字段为空并且设定宽度
                if (coloumName.StartsWith("cm"))
                {
                    if (float.TryParse(coloumName.Replace("cm", ""), out coloumWidth))
                    {
                        coloumValue = "　";
                        coloumWidth = coloumWidth / 10;
                    }
                }

                tableWidth += coloumWidth;

                var bottomBorder = string.Empty; //每个单元格底部border
                if (dataRows == 0)
                {
                    bottomBorder = "<BottomBorder><Style>None</Style></BottomBorder>";
                }
                //例外,如果coloumName包含Coloum之类的字段,则将value设成空
                if (coloumName.IndexOf("Column", System.StringComparison.Ordinal) > -1)
                {
                    coloumValue = "　";
                }
                //例外，替换coloumName中的+号
                coloumValue = coloumValue.Replace("_plus_", "+");

                //例外，如果coloumName包含序号，则替换
                for (int i = 1; i <= 10; i++)
                {
                    coloumValue = coloumValue.Replace(string.Format("_{0}_", i), "");
                }
                //例外，替换coloumName中的单引号和双引号
                coloumValue = coloumValue.Replace("9_", "(").Replace("_9", ")");


                fields.AppendFormat(
                    "<Field Name=\"{0}\"><DataField>{0}</DataField><rd:TypeName>System.String</rd:TypeName></Field>",
                    coloumName);

                if (coloumName != groupName)
                {

                    coloums.AppendFormat("<TablixColumn><Width>{0}cm</Width></TablixColumn>", coloumWidth);
                    tablixHearders.AppendFormat("<TablixCell><CellContents>" +
                                                "<Textbox Name=\"Textbox{0}{1}\"><CanGrow>true</CanGrow><KeepTogether>true</KeepTogether><Paragraphs><Paragraph>" +
                                                "<TextRuns><TextRun><Value>{2}</Value><Style /></TextRun></TextRuns><Style><TextAlign>Center</TextAlign></Style></Paragraph></Paragraphs>" +
                                                "<rd:DefaultName>Textbox{0}{1}</rd:DefaultName><Style><Border><Color>LightGrey</Color><Style>Solid</Style></Border>{3}" +
                                                "<PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight><PaddingTop>2pt</PaddingTop><PaddingBottom>2pt</PaddingBottom></Style></Textbox></CellContents></TablixCell>",
                                                coloumName, currentNamePrefix, coloumValue, bottomBorder);
                    tablixCells.AppendFormat(
                        "<TablixCell><CellContents><Textbox Name=\"{0}{1}1\"><CanGrow>true</CanGrow><KeepTogether>true</KeepTogether>" +
                        "<Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!{0}.Value</Value><Style /></TextRun></TextRuns><Style><TextAlign>{2}</TextAlign></Style></Paragraph></Paragraphs>" +
                        "<rd:DefaultName>{0}{1}1</rd:DefaultName><Style><Border><Color>LightGrey</Color><Style>Solid</Style></Border>" +
                        "<PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight><PaddingTop>2pt</PaddingTop><PaddingBottom>2pt</PaddingBottom></Style></Textbox></CellContents></TablixCell>",
                        coloumName, currentNamePrefix, textAlign);

                    tablixMembers.AppendFormat("<TablixMember />");
                }
                else //该列作为分组列展示
                {
                    groupHeader = string.Format(
                    "<TablixHeader><Size>{3}cm</Size><CellContents><Textbox Name=\"Textbox{0}{1}\"><CanGrow>true</CanGrow><KeepTogether>true</KeepTogether>" +
                    "<Paragraphs><Paragraph><TextRuns><TextRun><Value>{2}</Value><Style /></TextRun></TextRuns><Style><TextAlign>Center</TextAlign></Style></Paragraph></Paragraphs>" +
                    "<rd:DefaultName>Textbox{0}{1}</rd:DefaultName><Style><Border><Color>LightGrey</Color><Style>Solid</Style></Border><PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight><PaddingTop>2pt</PaddingTop><PaddingBottom>2pt</PaddingBottom></Style></Textbox></CellContents></TablixHeader>" +
                    "<TablixMembers><TablixMember /></TablixMembers>", coloumName, currentNamePrefix, coloumValue, coloumWidth);
                    groupMember = string.Format(
                        "<Group Name=\"Group{1}\"><GroupExpressions><GroupExpression>=Fields!{0}.Value</GroupExpression></GroupExpressions></Group>" +
                        //"<SortExpressions><SortExpression><Value>=Fields!{0}.Value</Value></SortExpression></SortExpressions>" +
                        "<TablixHeader><Size>{3}cm</Size><CellContents><Textbox Name=\"Group{0}{1}\"><CanGrow>true</CanGrow><KeepTogether>true</KeepTogether>" +
                        "<Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!{0}.Value</Value><Style /></TextRun></TextRuns><Style><TextAlign>{2}</TextAlign></Style></Paragraph>" +
                        "</Paragraphs><rd:DefaultName>Group1</rd:DefaultName><Style><Border><Color>LightGrey</Color><Style>Solid</Style></Border>" +
                        "<PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight><PaddingTop>2pt</PaddingTop><PaddingBottom>2pt</PaddingBottom>" +
                        "</Style></Textbox></CellContents></TablixHeader>", coloumName, currentNamePrefix, textAlign, coloumWidth);
                }
            }

            //计算表格应该离左边多少距离
            var leftPosition = 0F;
            if (tableWidth < 17)
            {
                leftPosition = (17F - tableWidth) / 2;
            }

            var dataSetName = string.Format("Data{0}", _reportItemPatterns.Count + _reportHeadPatterns.Count + 1);
            var reportItemPattern = new ReportItemPattern();
            reportItemPattern.Data = DynamicReportExtension.RemoveZeroData(data);
            reportItemPattern.DataSetName = dataSetName;
            reportItemPattern.DataSetString =
                DataSetPattern
                    .Replace("@DataSetName", dataSetName)
                    .Replace("@Fields", fields.ToString());
            reportItemPattern.TablixString =
                TablixPattern
                    .Replace("@DataSetName", dataSetName)
                    .Replace("@TablixColumns", coloums.ToString())
                    .Replace("@TablixHeader", tablixHearders.ToString())
                    .Replace("@TablixCells", tablixCells.ToString())
                    .Replace("@TablixMember", tablixMembers.ToString())
                    .Replace("@TopPosition", CaculatePlacePostion().ToString())
                    .Replace("@LeftPostion", leftPosition.ToString())
                    .Replace("@GroupHeader", groupHeader)
                    .Replace("@GroupMember", groupMember);


            //读取行数,如果是空行就加到新的
            if (dataRows == 0)
            {
                _reportHeadPatterns.Add(reportItemPattern);
            }
            else
            {
                reportItemPattern.ReportType = ReportType.Tables;
                _reportItemPatterns.Add(reportItemPattern);
            }
        }

        /// <summary>
        /// 得到某种类型数据的数量
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private int GetDataRowsCount(dynamic data)
        {
            if (data is DataTable)
            {
                return ((DataTable)data).Rows.Count;
            }
            else if (data is IEnumerable)
            {
                return Enumerable.Count(data);
            }
            else return 0;
        }

        /// <summary>
        /// 最终显示报表
        /// </summary>
        public void ShowReport()
        {
            //定义输出的doc
            var outPutDoc = string.Empty;

            //将每一个patter转换
            if (_reportItemPatterns.Count > 0 || _reportTitlePatterns.Count > 0)
            {
                var dataSetsString = new StringBuilder();
                var tablixString = new StringBuilder();

                foreach (var reportItemPattern in _reportItemPatterns)
                {
                    dataSetsString.Append(reportItemPattern.DataSetString);
                    tablixString.Append(reportItemPattern.TablixString);
                }
                foreach (var reportItemPattern in _reportHeadPatterns)
                {
                    dataSetsString.Append(reportItemPattern.DataSetString);
                    tablixString.Append(reportItemPattern.TablixString);
                }

                var reportTitleString = new StringBuilder();
                foreach (var reportTitlePattern in _reportTitlePatterns)
                {
                    reportTitleString.Append(reportTitlePattern);
                }

                //把文档中的文字替换掉
                switch (ReportType)
                {
                    case ReportType.Tables:
                        outPutDoc = _docTemplate.Replace("@DataSets", dataSetsString.ToString())
                                                   .Replace("@Tablix", tablixString.ToString())
                                                   .Replace("@Title", reportTitleString.ToString());
                        break;
                    case ReportType.Chart:
                        outPutDoc = _docTemplate.Replace("@DataSets", dataSetsString.ToString())
                                                   .Replace("@Tablix", tablixString.ToString())
                                                   .Replace("@Title", reportTitleString.ToString());
                        break;
                    case ReportType.Finally:
                        //替换datasetstring
                        outPutDoc = _docTemplate;
                        var pos = outPutDoc.IndexOf("<Body>", StringComparison.Ordinal);
                        outPutDoc = _docTemplate.Insert(pos,
                                                           string.Format(
                                                               "<DataSources><DataSource Name=\"DummyDataSource\"><ConnectionProperties><DataProvider>SQL</DataProvider><ConnectString /></ConnectionProperties><rd:DataSourceID>3eecdab9-6b4b-4836-ad62-95e4aee65ea8</rd:DataSourceID></DataSource></DataSources><DataSets>{0}</DataSets>",
                                                               dataSetsString));
                        //替换Tablix
                        pos = outPutDoc.IndexOf("<ReportItems>", StringComparison.Ordinal);
                        outPutDoc = outPutDoc.Insert(pos + 13, tablixString.ToString());
                        //替换title
                        outPutDoc = outPutDoc.Insert(pos + 13, reportTitleString.ToString());
                        break;
                }

                outPutDoc = outPutDoc.Replace("@InteractivePage", _showPager ? string.Empty : "<InteractiveHeight>9999cm</InteractiveHeight><InteractiveWidth>21cm</InteractiveWidth>");

                var doc = new XmlDocument();
                doc.LoadXml(outPutDoc);
                MemoryStream stream = GetRdlcStream(doc);

                //using (FileStream fs = new FileStream(@"c:\test.rdlc", FileMode.Create))
                //{
                //    stream.WriteTo(fs);
                //}

                //加载报表定义
                _report.LocalReport.LoadReportDefinition(stream);
                _report.LocalReport.DataSources.Clear();
                foreach (var reportItemPattern in _reportItemPatterns)
                {
                    _report.LocalReport.DataSources
                           .Add(new ReportDataSource(reportItemPattern.DataSetName + "Data",
                                                     reportItemPattern.Data));
                }
                foreach (var reportItemPattern in _reportHeadPatterns)
                {
                    _report.LocalReport.DataSources
                           .Add(new ReportDataSource(reportItemPattern.DataSetName + "Data",
                                                     reportItemPattern.Data));
                }

                _report.LocalReport.Refresh();
            }
        }

        /// <summary>
        /// 序列化到内存流
        /// </summary>
        /// <returns></returns>
        protected MemoryStream GetRdlcStream(XmlDocument xmlDoc)
        {
            MemoryStream ms = new MemoryStream();
            XmlSerializer serializer = new XmlSerializer(typeof(XmlDocument));
            serializer.Serialize(ms, xmlDoc);

            ms.Position = 0;
            return ms;
        }
    }

    internal static class DynamicReportExtension
    {
        public static dynamic RemoveZeroData(this object data)
        {
            if (data is DataTable)
            {
                return ((DataTable)data).ChangeEachColumnTypeToString();
            }
            else if (data is IEnumerable)
            {
                var _data = ((IEnumerable)data).Cast<object>();
                return _data.CopyToDataTable().RemoveZeroData();
            }
            return data;
        }

        public static DataTable ChangeEachColumnTypeToString(this DataTable dt)
        {
            DataTable tempdt = new DataTable();
            foreach (DataColumn dc in dt.Columns)
            {
                DataColumn tempdc = new DataColumn();

                tempdc.ColumnName = dc.ColumnName;
                tempdc.DataType = typeof(String);
                tempdt.Columns.Add(tempdc);
            }
            int coloumCount = dt.Columns.Count;
            foreach (DataRow dr in dt.Rows)
            {
                var newrow = tempdt.NewRow();

                for (int i = 0; i < coloumCount; i++)
                {
                    var value = dr[i].ToString();
                    switch (value)
                    {
                        case "0":
                        case "0.00%":
                            newrow[i] = "-";
                            break;
                        default:
                            newrow[i] = value;
                            break;
                    }

                }
                tempdt.Rows.Add(newrow);
            }
            return tempdt;
        }
    }

    internal static class DataSetLinqOperators
    {
        public static DataTable CopyToDataTable<T>(this IEnumerable<T> source)
        {
            return new ObjectShredder<T>().Shred(source, null, null);
        }

        public static DataTable CopyToDataTable<T>(this IEnumerable<T> source,
                                                   DataTable table, LoadOption? options)
        {
            return new ObjectShredder<T>().Shred(source, table, options);
        }

    }

    internal class ObjectShredder<T>
    {
        private FieldInfo[] _fi;
        private PropertyInfo[] _pi;
        private Dictionary<string, int> _ordinalMap;
        private Type _type;

        public ObjectShredder()
        {
            _type = typeof(T);
            _fi = _type.GetFields();
            _pi = _type.GetProperties();
            _ordinalMap = new Dictionary<string, int>();
        }

        public DataTable Shred(IEnumerable<T> source, DataTable table, LoadOption? options)
        {
            if (typeof(T).IsPrimitive)
            {
                return ShredPrimitive(source, table, options);
            }


            if (table == null)
            {
                table = new DataTable(typeof(T).Name);
            }

            // now see if need to extend datatable base on the type T + build ordinal map
            table = ExtendTable(table, typeof(T));

            table.BeginLoadData();
            using (IEnumerator<T> e = source.GetEnumerator())
            {
                while (e.MoveNext())
                {
                    if (options != null)
                    {
                        table.LoadDataRow(ShredObject(table, e.Current), (LoadOption)options);
                    }
                    else
                    {
                        table.LoadDataRow(ShredObject(table, e.Current), true);
                    }
                }
            }
            table.EndLoadData();
            return table;
        }

        public DataTable ShredPrimitive(IEnumerable<T> source, DataTable table, LoadOption? options)
        {
            if (table == null)
            {
                table = new DataTable(typeof(T).Name);
            }

            if (!table.Columns.Contains("Value"))
            {
                table.Columns.Add("Value", typeof(T));
            }

            table.BeginLoadData();
            using (IEnumerator<T> e = source.GetEnumerator())
            {
                Object[] values = new object[table.Columns.Count];
                while (e.MoveNext())
                {
                    values[table.Columns["Value"].Ordinal] = e.Current;

                    if (options != null)
                    {
                        table.LoadDataRow(values, (LoadOption)options);
                    }
                    else
                    {
                        table.LoadDataRow(values, true);
                    }
                }
            }
            table.EndLoadData();
            return table;
        }

        public DataTable ExtendTable(DataTable table, Type type)
        {
            // value is type derived  T, may need to extend table.
            foreach (FieldInfo f in type.GetFields())
            {
                if (!_ordinalMap.ContainsKey(f.Name))
                {
                    DataColumn dc = table.Columns.Contains(f.Name) ?
                    table.Columns[f.Name]
                    :
                    table.Columns.Add(f.Name, f.FieldType);
                    _ordinalMap.Add(f.Name, dc.Ordinal);
                }
            }
            foreach (PropertyInfo p in type.GetProperties())
            {
                if (!_ordinalMap.ContainsKey(p.Name))
                {
                    DataColumn dc = table.Columns.Contains(p.Name) ?
                    table.Columns[p.Name]
                    :
                    table.Columns.Add(p.Name);
                    _ordinalMap.Add(p.Name, dc.Ordinal);
                }
            }
            return table;
        }

        public object[] ShredObject(DataTable table, T instance)
        {

            FieldInfo[] fi = _fi;
            PropertyInfo[] pi = _pi;

            if (instance.GetType() != typeof(T))
            {
                ExtendTable(table, instance.GetType());
                fi = instance.GetType().GetFields();
                pi = instance.GetType().GetProperties();
            }

            Object[] values = new object[table.Columns.Count];
            foreach (FieldInfo f in fi)
            {
                values[_ordinalMap[f.Name]] = f.GetValue(instance);
            }

            foreach (PropertyInfo p in pi)
            {
                values[_ordinalMap[p.Name]] = p.GetValue(instance, null);
            }
            return values;
        }
    }
}
