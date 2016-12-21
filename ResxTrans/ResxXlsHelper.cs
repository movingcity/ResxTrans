using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Resources;
using System.Threading.Tasks;
using System.Windows.Forms;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using Application = System.Windows.Application;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace ResxTrans
{
    public static class ResxXlsHelper
    {
        private const string COLUMNCODE = "Name";
        private static string COLUMNDESC = "Chinese";
        private const string SHEETNAME = "StringResources$";

        private static object m_objOpt = System.Reflection.Missing.Value;

        public static async void ExportToXls(string fileName, string xlsFileName)
        {
            IList<ResxEntity> resxList = await ReadResx(fileName);
            ResxListToXls(resxList, xlsFileName);
            ShowXls(xlsFileName);
        }

        public static async Task<IList<ResxEntity>> ReadResx(string fileName)
        {
            IList<ResxEntity> resxList = new List<ResxEntity>();

            ResXResourceReader reader = new ResXResourceReader(fileName);
            FileInfo fi = new FileInfo(fileName);
            reader.BasePath = fi.DirectoryName;

            try
            {
                #region read

                foreach (DictionaryEntry de in reader)
                {
                    if (de.Value is string)
                    {
                        string key = (string)de.Key;

                        string value = de.Value.ToString().Replace("\r", "\\r").Replace("\n", "\\n");

                        ResxEntity r = new ResxEntity() { Name = key, English = value };

                        // Add value to English
                        resxList.Add(r);

                        // Add Chinese column
                    }
                }

                return resxList;

                #endregion
            }
            catch (Exception ex)
            {
                await ViewHelper.ShowMessageAsync("Information",
                     "A problem occured reading " + fileName + "\n" + ex.Message);
                return null;
            }
            finally
            {
                reader.Close();
            }
        }

        public static void ResxListToXls(IList<ResxEntity> resxList, string xlsFileName)
        {

            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

            Excel.Sheets sheets = wb.Worksheets;
            Excel.Worksheet sheet = (Excel.Worksheet)sheets.Item[1];
            sheet.Name = "StringResources";

            sheet.Cells[1, 1] = "Name";
            sheet.Cells[1, 2] = "English";
            sheet.Cells[1, 3] = "Chinese";

            int row = 2;

            //var controller = await ViewHelper.ShowProgressAsync("Please wait...", "Exporting to excel...");
            //controller.Maximum = resxList.Count;

            foreach (ResxEntity r in resxList)
            {
                sheet.Cells[row, 1] = r.Name;
                sheet.Cells[row, 2] = r.English;
                sheet.Cells[row, 3] = r.Chinese;
                row++;
                //controller.SetProgress(row);
            }
            //controller.SetProgress(resxList.Count);

            sheet.Cells.Range["A1", "Z1"].EntireColumn.AutoFit();

            // Save the Workbook and quit Excel.
            wb.SaveAs(xlsFileName, m_objOpt, m_objOpt,
                m_objOpt, m_objOpt, m_objOpt, Excel.XlSaveAsAccessMode.xlNoChange,
                m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);

            wb.Close(false, m_objOpt, m_objOpt);
            app.Quit();
            //await controller.CloseAsync();

        }

        public static void ShowXls(string xslFilePath)
        {
            if (!System.IO.File.Exists(xslFilePath))
                return;

            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(xslFilePath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
        true, false, 0, true, false, false);

            app.Visible = true;
        }

        public static void ExportToResx(string fileName, string resxFileName)
        {
            string resxName = resxFileName;

            try
            {
                // to StringResources.resx
                ResXResourceWriter rw = new ResXResourceWriter(resxName);
                using (rw)
                {
                    foreach (var pair in GetSECCodes(fileName, ResxType.English))
                    {

                        rw.AddResource(pair.Key, pair.Value);
                    }
                }

                Console.WriteLine("RESX generated successfully");

                // to StringResources.zh.resx
                ResXResourceWriter rwcn = new ResXResourceWriter(resxName.Remove(resxName.LastIndexOf('.')) + "zh.resx");
                using (rwcn)
                {
                    foreach (var pair in GetSECCodes(fileName, ResxType.Chinese))
                    {

                        rwcn.AddResource(pair.Key, pair.Value);
                    }
                }

                Console.WriteLine("RESX cn generated successfully");
            }
            catch (Exception e)
            {
                Console.WriteLine("Error while reading excel: " + e);
                Console.WriteLine("RESX not generated");
            }
        }

        // read the passed excel file and enumerate the key/value pairs for SEC
        private static IEnumerable<KeyValuePair<string, string>> GetSECCodes(string file, ResxType type)
        {
            if (type == ResxType.English)
            {
                COLUMNDESC = "English";
            }
            else
            {
                COLUMNDESC = "Chinese";
            }

            DataTable table = new DataTable("ResxTable");

            // connect to XLSX format MS2007++ file format
            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file
                + @";Extended Properties=""Excel 12.0;HDR=Yes""";

            OleDbConnection dbConn = new OleDbConnection(connString);
            using (dbConn)
            {
                dbConn.Open();
                string fetch = @"SELECT [" + COLUMNCODE + @"], [" + COLUMNDESC + @"] FROM [" + SHEETNAME + @"]";
                OleDbCommand cmdSelect = new OleDbCommand(fetch, dbConn);
                OleDbDataAdapter dbAdapter = new OleDbDataAdapter();
                dbAdapter.SelectCommand = cmdSelect;
                dbAdapter.Fill(table);
            }

            for (int i = 0; i < table.Rows.Count; i++)
            {
                yield return new KeyValuePair<string, string>(table.Rows[i][COLUMNCODE].ToString(),
                    table.Rows[i][COLUMNDESC].ToString());
            }
        }
    }
    enum ResxType
    {
        English,
        Chinese
    }
}
