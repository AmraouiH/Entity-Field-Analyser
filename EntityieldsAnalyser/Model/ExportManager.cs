using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk.Metadata;

namespace EntityieldsAnalyser
{
    public class FileManaged
    {

        public static void ExportFile(Dictionary<AttributeTypeCode, List<entityParam>> entityParam, EntityInfo entityInfo)
        {
            SaveFileDialog sfd = null;
            DataColumns[] fromatedList = FormatDataForExport(entityParam);
            String[] columnsHeaderName = new string[] {
                "Display Name","Schema Name","Type","Target", "Managed/Unmanaged","IsAuditable","IsSearchable","Required Level","Introduced Version","CreatedOn","Percentage Of Use"
            };
            int headerIndex = 9;
            int lineIndex = headerIndex + 1;

            if (fromatedList.Length > 0)
            {
                sfd = new SaveFileDialog();
                sfd.Filter = "Excel (.xlsx)|  *.xlsx;*.xls;";
                sfd.FileName = entityInfo.entityName + "_EntityKPIsExport_" + DateTime.Now.ToShortDateString().Replace('/','-')+".xlsx";
                bool fileError = false;

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(sfd.FileName))
                    {
                        try
                        {
                            File.Delete(sfd.FileName);
                        }
                        catch (IOException ex)
                        {
                            fileError = true;
                            MessageBox.Show("It wasn't possible to write the data to the disk." + ex.Message);
                        }
                    }
                    if (!fileError)
                    {
                        try
                        {
                            Microsoft.Office.Interop.Excel._Application XcelApp = new Microsoft.Office.Interop.Excel.Application();
                            Microsoft.Office.Interop.Excel._Workbook workbook = XcelApp.Workbooks.Add(Type.Missing);
                            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                            #region DataGrid
                            worksheet = workbook.ActiveSheet;
                            worksheet.Name = entityInfo.entityName+"_MetaData";

                            if (entityInfo != null)
                            {
                                for (int i = 1; i <= 6; i++)
                                {
                                    worksheet.Cells[i, 1].Font.Bold = true;
                                    worksheet.Cells[i, 1].Interior.Color = Color.Wheat;
                                    worksheet.Cells[i, 1].Font.Size = 12;
                                }

                                worksheet.Cells[1, 1]                = "Entity Display Name";
                                worksheet.Cells[1, 2]                = entityInfo.entityName;
                                worksheet.Cells[2, 1]                = "Entity Technical Name";
                                worksheet.Cells[2, 2]                = entityInfo.entityTechnicalName;
                                worksheet.Cells[3, 1]                = "CreatedOn";
                                worksheet.Cells[3, 2]                = entityInfo.entityDateOfCreation != DateTime.MinValue ? entityInfo.entityDateOfCreation.ToShortDateString() : String.Empty;
                                worksheet.Cells[4, 1]                = "Number Of Fields";
                                worksheet.Cells[4, 2]                = entityInfo.entityFieldsCount;
                                worksheet.Cells[5, 1]                = "Number Of Records";
                                worksheet.Cells[5, 2]                = entityInfo.entityRecordsCount;
                                worksheet.Cells[6, 1]                = "Entity Fields Volume Usage";
                                worksheet.Cells[6, 2]                = ((entityInfo.entityTotalUseOfColumns * 100) / entityInfo.entityDefaultColumnSize).ToString("0.##\\%");
                            }

                            for (int i = 1; i < columnsHeaderName.Length + 1; i++)
                            {
                                worksheet.Cells[headerIndex, i]                = columnsHeaderName[i - 1];
                                worksheet.Cells[headerIndex, i].Font.NAME      = "Calibri";
                                worksheet.Cells[headerIndex, i].Font.Bold      = true;
                                worksheet.Cells[headerIndex, i].Interior.Color = Color.Wheat;
                                worksheet.Cells[headerIndex, i].Font.Size      = 12;
                            }

                            for (int i = 0; i < fromatedList.Length; i++)
                            {
                                worksheet.Cells[i + lineIndex, 1]  = fromatedList[i].displayName;
                                worksheet.Cells[i + lineIndex, 2] = fromatedList[i].fieldName;
                                worksheet.Cells[i + lineIndex, 3]  = fromatedList[i].fieldType;
                                worksheet.Cells[i + lineIndex, 4]  = fromatedList[i].target;
                                worksheet.Cells[i + lineIndex, 5]  = fromatedList[i].isManaged;
                                worksheet.Cells[i + lineIndex, 6]  = fromatedList[i].isAuditable;
                                worksheet.Cells[i + lineIndex, 7]  = fromatedList[i].isSearchable; ;
                                worksheet.Cells[i + lineIndex, 8]  = fromatedList[i].requiredLevel; ;
                                worksheet.Cells[i + lineIndex, 9]  = fromatedList[i].introducedVersion;
                                worksheet.Cells[i + lineIndex, 10]  = fromatedList[i].dateOfCreation != DateTime.MinValue ? fromatedList[i].dateOfCreation.ToShortDateString() : String.Empty;
                                worksheet.Cells[i + lineIndex, 11] = fromatedList[i].percentageOfUse.Replace(",",".");
                                if (fromatedList[i].target == String.Empty) {
                                    worksheet.Cells[i + lineIndex, 4].Interior.Color = Color.Gainsboro;
                                }
                            }



                            worksheet.Columns.AutoFit();
                            #endregion
                            var xlSheets = workbook.Sheets as Microsoft.Office.Interop.Excel.Sheets;
                            var xlNewSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlSheets.Add(Type.Missing, xlSheets[1], Type.Missing, Type.Missing);
                            xlNewSheet.Name = "Charts";
                            #region chartManagedUnmanaged
                            //add data 
                            xlNewSheet.Cells[2, 2] = "Managed";
                            xlNewSheet.Cells[2, 2].Font.Color = Color.White;
                            xlNewSheet.Cells[2, 3] = entityInfo.managedFieldsCount;
                            xlNewSheet.Cells[2, 3].Font.Color = Color.White;


                            xlNewSheet.Cells[3, 2] = "Unmanaged";
                            xlNewSheet.Cells[3, 2].Font.Color = Color.White;
                            xlNewSheet.Cells[3, 3] = entityInfo.unmanagedFieldsCount;
                            xlNewSheet.Cells[3, 3].Font.Color = Color.White;


                            Microsoft.Office.Interop.Excel.Range chartRange;
                            Microsoft.Office.Interop.Excel.ChartObjects xlCharts = (Microsoft.Office.Interop.Excel.ChartObjects)xlNewSheet.ChartObjects(Type.Missing);
                            Microsoft.Office.Interop.Excel.ChartObject myChart   = (Microsoft.Office.Interop.Excel.ChartObject)xlCharts.Add(10, 10, 300, 250);
                            Microsoft.Office.Interop.Excel.Chart chartPage       = myChart.Chart;

                            chartPage.HasTitle = true;
                            chartPage.ChartTitle.Text = @"Managed\Unmanaged Fields";
                            chartRange = xlNewSheet.get_Range("B2", "C3");
                            chartPage.SetSourceData(chartRange, System.Reflection.Missing.Value);
                            chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlDoughnut;
                            #endregion
                            #region EntityFieldsCreated
                            //add data 
                            xlNewSheet.Cells[2, 10] = "Available Fields To Create";
                            xlNewSheet.Cells[2, 10].Font.Color = Color.White;
                            xlNewSheet.Cells[2, 11] = entityInfo.entityDefaultColumnSize - entityInfo.entityTotalUseOfColumns;
                            xlNewSheet.Cells[2, 11].Font.Color = Color.White;


                            xlNewSheet.Cells[3, 10] = "Created Fields";
                            xlNewSheet.Cells[3, 10].Font.Color = Color.White;
                            xlNewSheet.Cells[3, 11] = entityInfo.entityTotalUseOfColumns;
                            xlNewSheet.Cells[3, 11].Font.Color = Color.White;


                            Microsoft.Office.Interop.Excel.Range chartRangeTotaluse;
                            Microsoft.Office.Interop.Excel.ChartObjects xlChartsTotalUse     = (Microsoft.Office.Interop.Excel.ChartObjects)xlNewSheet.ChartObjects(Type.Missing);
                            Microsoft.Office.Interop.Excel.ChartObject totalUseChartChart    = (Microsoft.Office.Interop.Excel.ChartObject)xlChartsTotalUse.Add(510, 10, 300, 250);
                            Microsoft.Office.Interop.Excel.Chart chartPageTotalUse           = totalUseChartChart.Chart;

                            chartPageTotalUse.HasTitle = true;
                            chartPageTotalUse.ChartTitle.Text = @"Entity Fields Created";
                            chartRangeTotaluse = xlNewSheet.get_Range("J2", "K3");
                            chartPageTotalUse.SetSourceData(chartRangeTotaluse, System.Reflection.Missing.Value);
                            chartPageTotalUse.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlDoughnut;
                            #endregion
                            #region CustomStandar
                            //add data 
                            xlNewSheet.Cells[2, 20] = "Standard Fields";
                            xlNewSheet.Cells[2, 20].Font.Color = Color.White;
                            xlNewSheet.Cells[2, 21] = entityInfo.entityStandardFieldsCount;
                            xlNewSheet.Cells[2, 21].Font.Color = Color.White;


                            xlNewSheet.Cells[3, 20] = "Custom Fields";
                            xlNewSheet.Cells[3, 20].Font.Color = Color.White;
                            xlNewSheet.Cells[3, 21] = entityInfo.entityCustomFieldsCount;
                            xlNewSheet.Cells[3, 21].Font.Color = Color.White;

                            Microsoft.Office.Interop.Excel.Range chartRangeCustomStandard;
                            Microsoft.Office.Interop.Excel.ChartObjects xlChartsCustomStandard = (Microsoft.Office.Interop.Excel.ChartObjects)xlNewSheet.ChartObjects(Type.Missing);
                            Microsoft.Office.Interop.Excel.ChartObject customStandardChart     = (Microsoft.Office.Interop.Excel.ChartObject)xlChartsCustomStandard.Add(1010, 10, 300, 250);
                            Microsoft.Office.Interop.Excel.Chart chartPageCustomStandard       = customStandardChart.Chart;

                            chartPageCustomStandard.HasTitle        = true;
                            chartPageCustomStandard.ChartTitle.Text = @"Custom\Standard Fields";
                            chartRangeCustomStandard                = xlNewSheet.get_Range("T2", "U3");
                            chartPageCustomStandard.SetSourceData(chartRangeCustomStandard, System.Reflection.Missing.Value);
                            chartPageCustomStandard.ChartType       = Microsoft.Office.Interop.Excel.XlChartType.xlDoughnut;
                            #endregion
                            #region FieldsType
                            //add data 
                            int indicator = 0;
                            foreach (var item in entityParam)
                            {
                                xlNewSheet.Cells[indicator + 21, 10] = item.Key.ToString();
                                xlNewSheet.Cells[indicator + 21, 10].Font.Color = Color.White;
                                xlNewSheet.Cells[indicator + 21, 11] = item.Value.Count;
                                xlNewSheet.Cells[indicator + 21, 11].Font.Color = Color.White;
                                indicator++;
                            }

                            Microsoft.Office.Interop.Excel.Range chartRangeFieldType;
                            Microsoft.Office.Interop.Excel.ChartObjects xlChartsFieldTypes = (Microsoft.Office.Interop.Excel.ChartObjects)xlNewSheet.ChartObjects(Type.Missing);
                            Microsoft.Office.Interop.Excel.ChartObject fieldTypes          = (Microsoft.Office.Interop.Excel.ChartObject)xlCharts.Add(485, 270, 350, 300);
                            Microsoft.Office.Interop.Excel.Chart chartPageFieldTypes       = fieldTypes.Chart;

                            chartPageFieldTypes.HasTitle        = true;
                            chartPageFieldTypes.ChartTitle.Text = @"Entity Fields Types";
                            chartRangeFieldType                 = xlNewSheet.get_Range("J21", ("K"+(21+(indicator-1))).ToString());
                            chartPageFieldTypes.SetSourceData(chartRangeFieldType, System.Reflection.Missing.Value);
                            chartPageFieldTypes.ChartType       = Microsoft.Office.Interop.Excel.XlChartType.xlDoughnut;
                            #endregion
                            Microsoft.Office.Interop.Excel.Worksheet sheet = workbook.Worksheets[1];
                            sheet.Activate();

                            workbook.SaveAs(sfd.FileName);
                            XcelApp.Quit();

                            ReleaseObject(worksheet);
                            ReleaseObject(xlNewSheet);
                            ReleaseObject(workbook);
                            ReleaseObject(XcelApp);

                            if (MessageBox.Show("Would you like to open it?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            {
                                Process.Start(sfd.FileName);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error :" + ex.Message);
                        }
                    }
                }
            }
        }

        public static DataColumns[] FormatDataForExport(Dictionary<AttributeTypeCode, List<entityParam>> entityParams) {
            List<DataColumns> _dataToExport = new List<DataColumns>();
            foreach (var entityPara in entityParams) {
                foreach(var element in entityPara.Value)
                {
                    _dataToExport.Add(new DataColumns {
                        displayName         = element.displayName,
                        fieldName           = element.fieldName,
                        fieldType           = entityPara.Key.ToString(),
                        dateOfCreation      = element.dateOfCreation,
                        introducedVersion   = element.introducedVersion,
                        isAuditable         = element.isAuditable,
                        isManaged           = element.isManaged,
                        isOnForm            = element.isOnForm,
                        isSearchable        = element.isSearchable,
                        requiredLevel       = element.requiredLevel,
                        target              = element.target,
                        percentageOfUse     = element.percentageOfUse
                    });
                }
            }
            return _dataToExport.ToArray();
        }

        public static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.Message, "Error");
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
