using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;

using log4net;
using Newtonsoft.Json;
using OfficeOpenXml;
using Newtonsoft.Json.Linq;
using DosageManagement.DataAccess;

namespace DosageManagement.Logic
{
    public class DosageExcel
    {
        private const string MTH = "MTH";
        private const string MAT = "MAT";
        private const string MQT = "MQT";
        private const string DAILY_DOSAGE = "DAILYDOSAGE";
        private const string CELL_POISTION_VALUE_START = "W";
        private const string CELL_POISTION_VALUE_END = "BS";
        private const string CELL_POISTION_PTD_START = "BT";
        private const string CELL_POISTION_PTD_END = "DP";


        private ILog m_logger = DosageLog.GetLogger(typeof(DosageExcel));
        public string ExportExcels(string dataType, string selectedMarket, DateTime calculatedForm)
        {
            this.CleanTempDirectory();
            try
            {
                var fileNames = new List<string>();
                using (var ctx = new DosageDataContext())
                {
                    var marketList = from market in ctx.MarketSettings
                                     group market by market.Market1Name into m
                                     select m.Key;
                    foreach (var market in marketList)
                    {
                        if (selectedMarket == "all" || selectedMarket == market)
                        {
                            var fileName = NewFileName(market);
                            using (ExcelReport excelReport = new ExcelReport(fileName))
                            {
                                excelReport.AddSheet(DBUtility.QueryDosageResult(market, dataType, calculatedForm), market);
                                excelReport.Close();
                                fileNames.Add(fileName);
                            }
                        }
                    }
                }
                var zipName = NewFileName(null, true);
                ZipHelper.Zip(fileNames.ToArray(), zipName);
                return zipName;
            }
            catch (System.Data.SqlClient.SqlException sqlExp)
            {
                m_logger.Error($"error occurred while comunicating with database, message: {sqlExp.Message}");
                return null;
            }
            catch (System.IO.IOException ioExp)
            {
                m_logger.Error($"excel file build fail, message: {ioExp.Message}");
                return null;
            }
        }

        public void ImportExcel(Stream postStream)
        {
            LoadExcelToDB(postStream);

            /*do procedure to handle data*/
            DBUtility.ExecuteProcecure("usp_initBizData");
        }

        private void LoadExcelToDB(Stream postStream)
        {

            ExcelPackage document = null;
            try
            {
                document = new ExcelPackage(postStream);
                var sheets = document.Workbook.Worksheets;
                var config = GetColumnMapping();
                var tempDataHasCleaned = false; //determine whether the db is clean

                foreach (var sheet in sheets)
                {
                    if (sheet.Name.ToUpper() != MTH && sheet.Name.ToUpper() != MAT && sheet.Name.ToUpper() != DAILY_DOSAGE) { continue; }
                    if (sheet.Name.ToUpper() == DAILY_DOSAGE)
                    {
                        ImportDailyDosageToDB((JObject)config.Property("DailyDosageColumns").Value, sheet);
                        continue;
                    }

                    if (sheet.Name.ToUpper() != MTH && sheet.Name.ToUpper() != MAT) { continue; }
                    //clean temp db
                    if (!tempDataHasCleaned)
                    {
                        ClearTempData();
                        tempDataHasCleaned = true;
                    }
                    var dateCols = new List<string>();
                    int valueStartIndex, valueEndIndex, ptdStartIndex, ptdEndIndex, volumeStart, volumeEnd, datePartLength;
                    var dateColumnRange = (JObject)config.Property(sheet.Name.ToUpper()).Value;

                    datePartLength = Convert.ToInt32(dateColumnRange.Property("DateLengh").Value);
                    valueStartIndex = GetColumnIndex(dateColumnRange.Property("VALUEStart").Value.ToString());
                    valueEndIndex = GetColumnIndex(dateColumnRange.Property("VALUEEnd").Value.ToString());
                    ptdStartIndex = GetColumnIndex(dateColumnRange.Property("PTDStart").Value.ToString());
                    ptdEndIndex = GetColumnIndex(dateColumnRange.Property("PTDEnd").Value.ToString());
                    volumeStart = GetColumnIndex(dateColumnRange.Property("VOLUMEStart").Value.ToString());
                    volumeEnd = GetColumnIndex(dateColumnRange.Property("VOLUMEEnd").Value.ToString());

                    //VALUE
                    for (var i = valueStartIndex; i <= valueEndIndex; i++)
                    {
                        dateCols.Add("VAL_" + sheet.Cells[1, i].Value.ToString().Right(datePartLength).Replace("/", "_"));
                    }

                    //PTD
                    for (var i = ptdStartIndex; i <= ptdEndIndex; i++)
                    {
                        dateCols.Add("PTD_" + sheet.Cells[1, i].Value.ToString().Right(datePartLength).Replace("/", "_"));
                    }

                    //VOLUME
                    for (var i = volumeStart; i <= volumeEnd; i++)
                    {
                        dateCols.Add("VOL_" + sheet.Cells[1, i].Value.ToString().Right(datePartLength).Replace("/", "_"));
                    }

                    //create data table schema in memory
                    var dt = CreateMemoryDataTable((JObject)config.Property("InputColumns").Value, dateCols);
                    //fill data from excel sheet 


                    for (var i = 2; i <= sheet.Dimension.End.Row; i++)
                    {
                        var row = dt.NewRow();
                        for (var j = 0; j < sheet.Dimension.End.Column; j++)
                        {
                            try
                            {
                                row[j] = sheet.Cells[i, j + 1].Value?.ToString();
                            }
                            catch (InvalidCastException castExp)
                            {
                                m_logger.Error($"Invalid cast for sheet {sheet.Name} at Cells[{i}, {j + 1}], target type:{dt.Columns[j].DataType.ToString()}, cell value:{sheet.Cells[i, j + 1].Value.ToString()}, message:{castExp.Message}");
                                throw new ApplicationException($"there is invalid cell value, please hava a check at [{i}, {j + 1}] in sheet {sheet.Name} ");
                            }
                        }
                        dt.Rows.Add(row);
                    }
                    dt.AcceptChanges();

                    if (dt.Rows.Count > 0)
                    {
                        //create table into database
                        var dbTableName = CreateDBTable(sheet.Name, (JObject)config.Property("InputColumns").Value, dateCols);
                        //copy data to database
                        DBUtility.BulkInsert(dbTableName, dt);
                    }
                }
            }
            catch (DocumentFormat.OpenXml.Packaging.OpenXmlPackageException openXmlExp)
            {
                m_logger.Error($"please check whether the file is valid excel support by open xml, message: {openXmlExp.Message}");
            }
            finally
            {
                if (document != null)
                {
                    document.Dispose();
                }
            }
        }

        private void ImportDailyDosageToDB(JObject dailyDosgaeMap, ExcelWorksheet dataSheet)
        {
            try
            {

                var dtDosage = new DataTable();
                foreach (var prop in dailyDosgaeMap.Properties())
                {
                    DataColumn col = null;
                    if (prop.Value is JObject)
                    {
                        var dtType = Type.GetType(((JObject)prop.Value).Property("CLRType").Value.ToString());
                        col = new DataColumn(prop.Name, dtType);
                    }
                    else
                    {
                        col = new DataColumn(prop.Name, typeof(string));
                    }
                    dtDosage.Columns.Add(col);
                }
                dtDosage.AcceptChanges();
                var colProperties = dailyDosgaeMap.Properties().ToList();

                for (var i = 2; i <= dataSheet.Dimension.End.Row; i++)
                {
                    DataRow row = dtDosage.NewRow();
                    for (var j = 0; j < colProperties.Count; j++)
                    {
                        if (dataSheet.Cells[i, j + 1].Value != null)
                        {
                            if (row[j].GetType().ToString() == "System.Decimal")
                            {
                                if (dataSheet.Cells[i, j + 1].Value != null)
                                {
                                    row[j] = Convert.ToDecimal(dataSheet.Cells[i, j + 1].Value);
                                }
                            }
                            else
                            {
                                row[j] = dataSheet.Cells[i, j + 1].Value.ToString();
                            }
                        }
                    }
                    dtDosage.Rows.Add(row);
                }
                dtDosage.AcceptChanges();

                DBUtility.CleanData("DailyDosage");
                DBUtility.BulkInsert("DailyDosage", dtDosage);

            }
            catch (Exception exp)
            {
                m_logger.Error($"error message:{exp.Message}, statck info:{exp.StackTrace}");
                throw new ApplicationException("error occurred, please contact support!");
            }
        }

        /// <summary>
        /// create db table, and return the table name
        /// </summary>
        /// <param name="periodName"></param>
        /// <param name="cols"></param>
        /// <param name="dateCols"></param>
        /// <returns></returns>
        private string CreateDBTable(string periodName, JObject cols, List<string> dateCols)
        {
            string tableName = "BIZ_" + periodName + DateTime.Now.ToString("MMdd");
            StringBuilder sql = new StringBuilder();
            sql.Append($"create table {tableName}({Environment.NewLine}");
            foreach (JProperty property in cols.Properties())
            {
                var col = ((JObject)property.Value);
                sql.Append($" {Environment.NewLine}{property.Name} {col.Property("Type").Value.ToString()}");
                if (col.Property("PrimaryKey") != null)
                {
                    sql.Append(" identity(1,1)");
                }
                sql.Append(",");
            }

            foreach (var date in dateCols)
            {
                sql.Append($"{Environment.NewLine} {date} decimal(38,10),");
            }
            sql.Remove(sql.Length - 1, 1);
            sql.Append(")");
            DBUtility.ExecuteNoneQuery(sql.ToString());
            return tableName;
        }

        /// <summary>
        /// this method will drop temp table and shrink db to fress disk
        /// </summary>
        private void ClearTempData()
        {
            DBUtility.ExecuteProcecure("usp_cleanBizTable");
        }

        private DataTable CreateMemoryDataTable(JObject cols, List<string> dateCols)
        {
            DataTable dt = new DataTable();
            foreach (JProperty property in cols.Properties())
            {
                var col = ((JObject)property.Value);
                if (col.Property("PrimaryKey") != null)
                {
                    continue;
                }
                dt.Columns.Add(new DataColumn(property.Name, typeof(string)));
            }

            foreach (var col in dateCols)
            {
                dt.Columns.Add(new DataColumn(col, typeof(Decimal)));
            }
            dt.AcceptChanges();
            return dt;
        }


        private JObject GetColumnMapping()
        {
            string configFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
            if (!configFilePath.EndsWith("\\")) configFilePath += "\\";
            configFilePath += "bin\\ExcelColumnMapping.json";

            StreamReader sr = new StreamReader(configFilePath);
            JsonTextReader jreader = new JsonTextReader(sr);
            var configObj = (JObject)JObject.ReadFrom(jreader);
            jreader.Close();
            return configObj;


            /*
             * var columnMapping = (JObject)jobj["InputColumns"];
                ((JObject)property).Property("Type");
                ((JObject)property).Property("OrginName");
                property.Name;
             */
        }


        /// <summary>
        /// the column index will start with 1
        /// </summary>
        /// <returns></returns>
        private int GetColumnIndex(string cellPostionCol)
        {
            cellPostionCol = cellPostionCol.Trim().ToUpper();
            int colIndex = 0;
            for (var i = 0; i < cellPostionCol.Length; i++)
            {
                var idx = cellPostionCol.Length - i - 1;
                var assciDec = Convert.ToInt32(cellPostionCol[idx]);

                var colDec = (assciDec - 64) % 26;
                colIndex += Convert.ToInt32(Math.Pow(26, i)) * colDec;
            }
            return colIndex;
        }

        private string NewFileName(string marketName, bool isZip = false)
        {
            var path = System.Web.HttpContext.Current.Server.MapPath("~/temp_excel");
            if (!path.EndsWith("\\")) path += "\\";

            if (isZip)
            {
                return $"{path}myResult_{DateTime.Now.ToString("MMdd")}.zip";
            }
            else
            {
                return $"{path}{marketName}_{DateTime.Now.ToString("HHmm")}.xlsx";
            }
        }

        private void CleanTempDirectory()
        {
            var path = System.Web.HttpContext.Current.Server.MapPath("~/temp_excel");
            DirectoryInfo directory = new DirectoryInfo(path);
            foreach (var fileInfo in directory.GetFiles())
            {
                File.Delete(fileInfo.FullName);
            }
        }
    }
}
