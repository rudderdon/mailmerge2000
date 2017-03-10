using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace MailMerge2000.Data
{

  public class ExcelApp
  {

    private Application _excelApp;
    private int _processId = 0;

    public Workbook ExcelWorkbook { get; set; }
    public Dictionary<string, Array> WorkSheetData { get; set; }
    public Dictionary<string, DataTable> WorkSheetTables { get; set; }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="filePath">Excel Workbook File Path</param>
    /// <remarks>Reads data into both array and datatable form for your enjoyment</remarks>
    public ExcelApp(string filePath)
    {
      // Widen Scope
      ExcelStart(filePath);
      WorkSheetData = new Dictionary<string, Array>();
      WorkSheetTables = new Dictionary<string, DataTable>();
      if (ExcelWorkbook != null)
      {
        foreach (Worksheet x in ExcelWorkbook.Sheets)
        {
          GetValues(x);
        }
      }
    }

    #region Private Members - Reading Data

    /// <summary>
    /// Return an Empty Array with Bounds set to maximum row and column counts
    /// </summary>
    /// <param name="ws">Worksheet to Check</param>
    /// <returns>Range for all Data - no header</returns>
    /// <remarks></remarks>
    private Range GetFullDataRange(Worksheet ws)
    {
      try
      {
        // Count Columns
        int m_iCol = 0;
        bool m_noData = false;
        while (!m_noData)
        {
          try
          {
            // Test for Value - Make sure we are testing the header row
            m_iCol++;
            if (!string.IsNullOrEmpty(ws.Cells[1, m_iCol].Value.ToString())) { }

          }
          catch
          {
            m_noData = true;
            m_iCol--;
          }
        }

        // Count Rows
        int m_iRow = 0;
        m_noData = false;
        while (!m_noData)
        {
          try
          {
            // Test for Value
            m_iRow++;
            if (!string.IsNullOrEmpty(ws.Cells[m_iRow, 1].Value.ToString())) { }
          }
          catch
          {
            m_iRow--;
            m_noData = true;
          }
        }

        // Return Range
        Range m_range = ws.Range[ws.Cells[1, 1], ws.Cells[m_iRow, m_iCol]];
        return m_range;

      }
      catch (Exception ex)
      { }
      return null;
    }
    
    /// <summary>
    /// Range to Table
    /// </summary>
    /// <param name="name"></param>
    /// <param name="r"></param>
    /// <returns></returns>
    private DataTable GetTable(string name, Range r)
    {
      DataTable m_dt = new DataTable(name);
      try
      {
        object[,] arr = r.Value;
        // Rows
        for (int i = 1; i < arr.GetUpperBound(0) + 1; i++)
        {
          DataRow m_row = m_dt.NewRow();
          // Columns
          for (int ii = 1; ii < arr.GetUpperBound(1) + 1; ii++)
          {
            string m_value = "";
            object m_x = arr.GetValue(i, ii);
            if (m_x != null)
            {
              m_value = m_x.ToString();
            }
            if (i == 1)
            {
              // Headers
              m_dt.Columns.Add(m_value);
            }
            else
            {
              // Rows
              m_row[ii - 1] = m_value;
            }
          }
          if (i > 1) m_dt.Rows.Add(m_row);
        }
      } 
      catch (Exception)
      {
      }
      return m_dt;
    }

    /// <summary>
    /// Get the Excel Data
    /// </summary>
    /// <param name="ws">Worksheet</param>
    /// <remarks></remarks>
    private void GetValues(Worksheet ws)
    {
      try
      {
        // Worksheet Data
        Range m_rangeAll = GetFullDataRange(ws);
        WorkSheetData.Add(ws.Name, m_rangeAll.Value);
        WorkSheetTables.Add(ws.Name, GetTable(ws.Name, m_rangeAll));
      }
      catch (Exception ex)
      { }
    }

    #endregion

    #region Internal Members - Excel Startup and Shutdown

    /// <summary>
    /// Start Excel, open spreadsheet file
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns></returns>
    internal bool ExcelStart(string filePath)
    {
      // Does the File Exist?
      if (!File.Exists(filePath))
      {
        // Too lazy to add an error
        return false;
      }
      else
      {
        try
        {
          // Store the Excel processes before opening.
          Process[] processesBefore = Process.GetProcessesByName("excel");
          _excelApp = new Application();
          ExcelWorkbook = _excelApp.Workbooks.Open(filePath);
          ExcelWorkbook.Application.Windows[Path.GetFileName(filePath)].Visible = true;
          Process[] processesAfter = Process.GetProcessesByName("excel");

          foreach (Process process in processesAfter)
          {
            if (!processesBefore.Select(p => p.Id).Contains(process.Id))
            {
              _processId = process.Id;
            }
          }

          return true;
        }
        catch (Exception ex)
        {
          // Too lazy to add an error
          return false;
        }
      }
      return false;
    }

    /// <summary>
    /// Terminate the Excel Application
    /// </summary>
    /// <remarks></remarks>
    internal void ExcelShutDown()
    {
      try
      {

        // Quit if we have an Application Reference
        if (_excelApp != null)
        {
          ExcelWorkbook.Close(true);
          _excelApp.Workbooks.Close();
          _excelApp.Quit();
          if (_processId != 0)
          {
            Process process = Process.GetProcessById(_processId);
            process.Kill();
            _processId = 0;
          }
        }

        if (_processId != 0)
        {
          Process process = Process.GetProcessById(_processId);
          process.Kill();
        }

        // Set References to Nothing
        ExcelWorkbook = null;
        _excelApp = null;
        GC.Collect();

      }
      catch { }

    }

    #endregion

  }
}