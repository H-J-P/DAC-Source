#region Lizenz
//----------------------------------------------------------------------
//    Copyright (c) 2014 Heinz-Joerg Puhlmann
//
// Hiermit wird unentgeltlich, jeder Person, die eine Kopie der Software
// und der zugehörigen Dokumentationen (die "Software") erhält, die
// Erlaubnis erteilt, uneingeschränkt zu benutzen, inklusive und ohne
// Ausnahme, dem Recht, sie zu verwenden, kopieren, ändern, fusionieren,
// verlegen, verbreiten, unter-lizenzieren und/oder zu verkaufen, und 
// Personen, die diese Software erhalten, diese Rechte zu geben, unter
// den folgenden Bedingungen:
//
// Der obige Urheberrechtsvermerk und dieser Erlaubnisvermerk sind in 
// alle Kopien oder Teilkopien der Software beizulegen.
//
// DIE SOFTWARE WIRD OHNE JEDE AUSDRÜCKLICHE ODER IMPLIZIERTE GARANTIE 
// BEREITGESTELLT, EINSCHLIESSLICH DER GARANTIE ZUR BENUTZUNG FÜR DEN
// VORGESEHENEN ODER EINEM BESTIMMTEN ZWECK SOWIE JEGLICHER 
// RECHTSVERLETZUNG, JEDOCH NICHT DARAUF BESCHRÄNKT. IN KEINEM FALL SIND
// DIE AUTOREN ODER COPYRIGHTINHABER FÜR JEGLICHEN SCHADEN ODER SONSTIGE
// ANSPRUCH HAFTBAR ZU MACHEN, OB INFOLGE DER ERFÜLLUNG VON EINEM 
// VERTRAG, EINEM DELIKT ODER ANDERS IM ZUSAMMENHANG MIT DER BENUTZUNG
// ODER SONSTIGE VERWENDUNG DER SOFTWARE ENTSTANDEN
//----------------------------------------------------------------------
#endregion

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Drawing.Printing;

namespace DAC
{
    class ImportExport
    {
        #region Member

        public enum XlWBATemplate
        {
            xlWBATChart,
            xlWBATExcel4IntlMacroSheet,
            xlWBATExcel4MacroSheet,
            xlWBATWorksheet
        }

        static string newline = System.Environment.NewLine;
        public static List<String> log = new List<string>();
        private static int prozessID;

        #endregion

        #region static Memberfunctions

        [DllImport("User32.dll", SetLastError = true)]
        public static extern int SetForegroundWindow(IntPtr hwnd);

        public static void CheckProzessAndSendKeys(string programName, string sendKeys)
        {
            try
            {
                Process[] processes = Process.GetProcessesByName(programName);

                if (processes.Length > 0)
                    prozessID = processes[0].Id;
                else
                {
                    Process process = new Process();
                    process.StartInfo.FileName = programName;
                    process.Start();
                    prozessID = process.Id;
                }

                System.IntPtr MainHandle = Process.GetProcessById(prozessID).MainWindowHandle;
                SetForegroundWindow(MainHandle);
                SendKeys.SendWait(sendKeys);
            }
            catch
            {
                LogMessage("Program " + programName + " not found.", true);
            }
        }

        public static void CreateCSVFile(DataTable dt, string strFilePath)
        {
            if (dt == null)
                return;

            // Create the CSV file to which grid data will be exported.
            StreamWriter sw = new StreamWriter(strFilePath, false);
            int iColCount = dt.Columns.Count;

            // First we will write the headers.
            for (int i = 0; i < iColCount; i++)
            {
                sw.Write(dt.Columns[i]);

                if (i < iColCount - 1)
                    sw.Write(";");
            }
            sw.Write(sw.NewLine);

            // Now write all the rows.
            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < iColCount; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                        sw.Write(dr[i].ToString());

                    if (i < iColCount - 1)
                        sw.Write(";");
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        public static void ComboboxToXml(String startElement, String fileName, ComboBox cbo, String elementString)
        {
            XmlTextWriter xtw = null;
            String fileLoction = System.Windows.Forms.Application.StartupPath + fileName;

            try
            {
                xtw = new XmlTextWriter(fileLoction, Encoding.UTF8);
                xtw.Formatting = Formatting.Indented;

                xtw.WriteStartDocument();
                xtw.WriteStartElement(startElement);

                for (int i = 0; i < cbo.Items.Count; i++)
                {
                    xtw.WriteElementString(elementString, (string)cbo.Items[i]);
                }
                xtw.WriteEndElement();
                xtw.WriteEndDocument();
            }
            catch (Exception e)
            {
                LogMessage("Achtung Fehler, " + e, true);
            }
            finally
            {
                if (xtw != null)
                    xtw.Close();
            }
        }

        public static void LogMessage(String message, bool withTimeStemp)
        {
            try
            {
                if (withTimeStemp)
                    log.Add(DateTime.Now.ToString("HH:mm:ss.fff") + "  " + message);
                else
                    log.Add(message);
            }
            catch { }
        }

        public static void DatasetToXml(String fileName, DataSet ds)
        {
            String fileLoction = "";

            if (fileName.IndexOf("\\") > 0)
                fileLoction = fileName;
            else
                fileLoction = System.Windows.Forms.Application.StartupPath + "\\" + fileName;

            try
            {
                if (ds != null)
                    ds.WriteXml(fileLoction);
            }
            catch (Exception e)
            {
                LogMessage("DatasetToXml: " + fileName + " ... " + e, true);
            }
        }

        /// <summary>
        /// Solution for .net 4.0 ******
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="tableName"></param>
        /// <param name="fileName"></param> without extension
        /// <param name="selectWorkSheet"></param>
        //public static void ExportDataTableToExcel(ref System.Data.DataTable[] dt, ref string[] tableName, string fileName, int selectWorkSheet)
        //{
        //    Excel.Application excel = new Excel.Application();
        //    excel.SheetsInNewWorkbook = tableName.Length;
        //    excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        //    Excel._Worksheet worksheet = excel.ActiveSheet(0);

        //    for (int i = 0; i < tableName.Length; i++)  // Table
        //    {
        //        worksheet = excel.ActiveSheet[i];
        //        worksheet.Name = tableName[i];

        //        int iCol = 0;
        //        foreach (DataColumn c in dt[i].Columns) // Header / Kopfzeile
        //        {
        //            iCol++;
        //            worksheet.Cells[1, iCol] = c.ColumnName;
        //        }

        //        int iRow = 0;
        //        foreach (DataRow r in dt[i].Rows) // Row / Zeile
        //        {
        //            iRow++;
        //            iCol = 0;

        //            foreach (DataColumn c in dt[i].Columns) // Column / Spalte
        //            {
        //                iCol++;
        //                worksheet.Cells[iRow + 1, iCol] = r[c.ColumnName].ToString();
        //            }
        //        }
        //        worksheet.Range["A"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatTable10);
        //    }
        //    worksheet = excel.ActiveSheet[selectWorkSheet];
        //    worksheet.SaveAs(string.Format(@"{0}\" + fileName + ".xlsx", Environment.CurrentDirectory));
        //    excel.Visible = true;
        //    //excel.Quit();
        //}

         //<summary>
         //Solution for .net 3.5 ********
         //</summary>
         //<param name="dt"></param>
         //<param name="tableName"></param>
         //<param name="selectWorkSheet"></param>

        //public static void ExportDataTableToExcel(ref System.Data.DataTable[] dt, ref string[] tableName, string fileName, int selectWorkSheet)
        //{
        //    if (dt == null || tableName == null)
        //        return;

        //    // Create an Excel object and add workbook...
        //    ApplicationClass excel = new ApplicationClass();
        //    Workbook workbook = excel.Application.Workbooks.Add(true);
        //    Sheets sheets = null;
        //    Worksheet worksheet = null;

        //    try
        //    {
        //        // for every table
        //        for (int i = 0; i < tableName.Length; i++)
        //        {
        //            sheets = workbook.Sheets;
        //            worksheet = (Worksheet)sheets.Add(sheets[i + 1], Type.Missing, Type.Missing, Type.Missing);

        //            Range SourceRange = (Range)worksheet.get_Range("A1", (char)((int)'A' + dt[i].Columns.Count - 1) + (dt[i].Rows.Count + 1).ToString());

        //            // Format as string
        //            SourceRange.NumberFormat = "@";

        //            // Add column headings...
        //            int iCol = 0;
        //            foreach (DataColumn c in dt[i].Columns)
        //            {
        //                iCol++;
        //                excel.Cells[1, iCol] = c.ColumnName;
        //            }

        //            // for each row and column of data...
        //            int iRow = 0;

        //            foreach (DataRow r in dt[i].Rows)
        //            {
        //                iRow++;
        //                iCol = 0;

        //                foreach (DataColumn c in dt[i].Columns)
        //                {
        //                    iCol++;
        //                    excel.Cells[iRow + 1, iCol] = r[c.ColumnName].ToString();
        //                }
        //            }
        //            worksheet.Name = tableName[i];
        //            FormatAsTable(SourceRange, tableName[i], "TableStyleMedium15");
        //        }
        //        // select the first worksheet
        //        worksheet = (Worksheet)workbook.Sheets[tableName[selectWorkSheet]];
        //        worksheet.Select(Type.Missing);

        //        // make Excel visible and activate the worksheet...
        //        excel.Visible = true;
        //    }
        //    catch
        //    {
        //        excel.Quit();
        //    }
        //}

        //public static void FormatAsTable(Range SourceRange, string TableName, string TableStyleName)
        //{
        //    SourceRange.Worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange,
        //    SourceRange, System.Type.Missing, XlYesNoGuess.xlYes, System.Type.Missing).Name = TableName;

        //    SourceRange.Select();
        //    SourceRange.Worksheet.ListObjects[TableName].TableStyle = TableStyleName;
        //    SourceRange.EntireColumn.AutoFit();
        //}

        public static string GetDefaultPrinter()
        {
            PrinterSettings settings = new PrinterSettings();
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                settings.PrinterName = printer;
                if (settings.IsDefaultPrinter)
                    return printer;
            }
            return string.Empty;
        }

        public static void TableToXml(String fileName, DataTable tb)
        {
            if (tb == null)
                return;

            String fileLoction = System.Windows.Forms.Application.StartupPath + "\\" + fileName;

            try
            {
                tb.WriteXml(fileLoction, true);
                
            }
            catch (Exception e)
            {
                LogMessage("TableToXml Error: " + e, true);
            }
        }

        public static int WriteListBoxToFile(ListBox lb, String path)
        {
            if (lb == null)
                return -1;

            StringBuilder buffer = new StringBuilder();
            StreamWriter sw = null;

            for (int i = 0; i < lb.Items.Count; i++)
            {
                buffer.Append(lb.Items[i].ToString());
                buffer.Append(newline);
            }

            if (buffer.Length > 0)
            {
                try
                {
                    if (!File.Exists(path))
                    {
                        try
                        {
                            sw = File.CreateText(path);
                        }
                        catch { }
                    }
                    else
                    {
                        sw = File.AppendText(path);
                    }
                    sw.Write(buffer);
                    sw.Close();
                    lb.Items.Clear();
                }
                catch
                {
                   return -1;
                }
            }
            return 0;
        }

        public static void XmlToDataSet(String fileName, DataSet ds)
        {
            if (ds == null)
                return;

            String fileLoction = "";

            if (fileName.IndexOf("\\") > 0)
                fileLoction = fileName;
            else
                fileLoction = System.Windows.Forms.Application.StartupPath + "\\" + fileName;

            try
            {
                ds.ReadXml(fileLoction);
            }
            catch (Exception e)
            {
                LogMessage("XmlToDataSet (read file to dataset): " + fileName + " ... " + e.ToString()/*.Substring(0,100)*/, true);
            }
        }

        public static String XmlReadElement(String fileName, String element, bool useApplicationPath)
        {
            XmlTextReader xtr = null;
            String fileLoction = "";

            if (useApplicationPath)
                fileLoction = System.Windows.Forms.Application.StartupPath + fileName;
            else
                fileLoction = fileName;

            String readName = "";
            String readValue = "";

            try
            {
                // Initialize the XmlTextReader variable with the name of the file
                xtr = new XmlTextReader(fileLoction);
                xtr.WhitespaceHandling = WhitespaceHandling.None;

                // Moves the reader to the root element.
                xtr.MoveToContent();

                // Scan the XML file
                while (xtr.Read())
                {
                    readValue = "";
                    // every time you find an element, find out what type it is
                    switch (xtr.NodeType)
                    {
                        case XmlNodeType.Element:
                            readName = xtr.Name;
                            break;
                        case XmlNodeType.Text:
                            readValue = xtr.Value;
                            break;
                        case XmlNodeType.CDATA:
                            break;
                        case XmlNodeType.ProcessingInstruction:
                            break;
                        case XmlNodeType.Comment:
                            break;
                        case XmlNodeType.XmlDeclaration:
                            break;
                        case XmlNodeType.Document:
                            break;
                        case XmlNodeType.DocumentType:
                            break;
                        case XmlNodeType.EntityReference:
                            break;
                        case XmlNodeType.EndElement:
                            readName = "";
                            break;
                    }
                    if (readName == element && readValue != "")
                        return readValue;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lese/Schreibfehler in: " + fileName + " Element: " + element + " / " + ex);
            }
            finally
            {
                // Close the XmlTextReader
                if (xtr != null)
                    xtr.Close();
            }
            return readValue;
        }

        public static void XmlToTable(String fileName, DataTable tb)
        {
            if (tb == null)
                return;

            String fileLoction = System.Windows.Forms.Application.StartupPath + "\\" + fileName;

            try
            {
                tb.ReadXml(fileLoction);
            }
            catch (Exception e)
            {
                LogMessage("Achtung Fehler in XmlToTable, " + e, true);
            }
        }

        public static void XmlToCombobox(String fileName, ComboBox cbo, String elementString)
        {
            XmlTextReader xtr = null;
            String fileLoction = System.Windows.Forms.Application.StartupPath + fileName;

            try
            {
                // Initialize the XmlTextReader variable with the name of the file
                xtr = new XmlTextReader(fileLoction);
                xtr.WhitespaceHandling = WhitespaceHandling.None;

                // Scan the XML file
                while (xtr.Read())
                {
                    // every time you find an element, find out what type it is
                    switch (xtr.NodeType)
                    {
                        case XmlNodeType.Text:
                            // If you find text, put it in the combobox list
                            cbo.Items.Add(xtr.Value);
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                LogMessage("Achtung Fehler, " + e, true);
            }
            finally
            {
                // Close the XmlTextReader
                if (xtr != null)
                    xtr.Close();
            }
        }

        public static void ReadTextFileToStringArray(String filename, ref String[] lines)
        {
            lines = System.IO.File.ReadAllLines(filename);
        }

        #endregion
    }
}
