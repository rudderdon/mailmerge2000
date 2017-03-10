using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using MailMerge2000.Data;
using Microsoft.Win32;

namespace MailMerge2000.UI
{
  public partial class MainWindow : Window
  {

    private const string _outlookTemplatesPath = @"C:\$bim\3 - Documentation\1 - Outlook Email Templates";

    private ExcelApp _excel;
    private readonly OutlookApp _outlook;

    /// <summary>
    /// Constructor
    /// </summary>
    public MainWindow()
    {
      InitializeComponent();
      _outlook = new OutlookApp();
      GetOutlookTemplates();
      FormIsReady();
    }

    #region Private Members

    /// <summary>
    /// Form State
    /// </summary>
    private void FormIsReady()
    {
      bool m_state = _excel != null;
      ButtonOk.IsEnabled = m_state;
      ButtonSave.IsEnabled = m_state;
    }

    /// <summary>
    /// Get the Outlook Templates
    /// </summary>
    private void GetOutlookTemplates()
    {
      ComboBoxTemplates.Items.Clear();
      if (!Directory.Exists(_outlookTemplatesPath)) return;
      DirectoryInfo m_diNet = new DirectoryInfo(_outlookTemplatesPath);
      List<FileHelper> m_files = m_diNet.EnumerateFiles("*.oft", 
        SearchOption.TopDirectoryOnly).Select(x => new FileHelper(x)).ToList();

      if (m_files.Count > 0)
      {
        ComboBoxTemplates.ItemsSource = m_files;
        ComboBoxTemplates.DisplayMemberPath = "Name";
        ComboBoxTemplates.SelectedIndex = 0;
      }
    }

    /// <summary>
    /// Send some emails!
    /// </summary>
    /// <param name="sendToRecipients"></param>
    private void SendEmails(bool sendToRecipients)
    {
      try
      {
        FileHelper m_template = (FileHelper)ComboBoxTemplates.SelectedItem;
        KeyValuePair<string, DataTable> m_obj = (KeyValuePair<string, DataTable>)ComboBoxWorksheets.SelectedItem;
        DataTable m_dt = m_obj.Value;
        if (m_dt != null)
        {

          // Email per Row
          foreach (DataRow x in m_dt.Rows)
          {
            Dictionary<string, string> m_values = new Dictionary<string, string>();
            List<string> m_attachments = new List<string>();
            string m_emailSubject = x["EmailSubject"].ToString();
            string m_emailTo = x["EmailTo"].ToString();
            string m_emailCc = x["EmailCc"].ToString();
            string m_emailBcc = x["EmailBcc"].ToString();
            string m_emailAtt = x["EmailAttachments"].ToString();
            if (!string.IsNullOrEmpty(m_emailAtt))
            {
              m_attachments = m_emailAtt.Split(',').ToList();
            }

            for (int i = 0; i < m_dt.Columns.Count; i++)
            {
              m_values.Add(m_dt.Columns[i].ToString(), x[i].ToString());
            }

            // Send it
            _outlook.EmailFromTemplate(
              m_template.FullName,
              m_emailSubject,
              m_emailTo,
              m_emailCc,
              m_emailBcc, 
              m_values, 
              m_attachments, 
              sendToRecipients);

          }
        }
      } 
      catch (Exception)
      {
      }
    }

    #endregion

    #region Private Members - Form Controls & Events

    /// <summary>
    /// Send Mail
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void ButtonOk_OnClick(object sender, RoutedEventArgs e)
    {
      SendEmails(true);
    }

    /// <summary>
    /// Create and Save Mail
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void ButtonSave_OnClick(object sender, RoutedEventArgs e)
    {
      SendEmails(false);
    }

    /// <summary>
    /// Close and Cancel
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void ButtonCancel_OnClick(object sender, RoutedEventArgs e)
    {
      Close();
    }

    /// <summary>
    /// Open Excel
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void ButtonOopenExcel_OnClick(object sender, RoutedEventArgs e)
    {
      // Open Excel and Read Data
      OpenFileDialog m_dlg = new OpenFileDialog
      {
        Filter = "Excel Docs (*.xlsx)|*.xlsx",
        Title = "Select an Excel Mail Merge Source",
        Multiselect = false
      };
      if (m_dlg.ShowDialog() != true) return;
      if (!string.IsNullOrEmpty(m_dlg.FileName))
      {
        _excel = new ExcelApp(m_dlg.FileName);
      }

      // Worksheets
      ComboBoxWorksheets.ItemsSource = _excel.WorkSheetTables;
      ComboBoxWorksheets.DisplayMemberPath = "Key";
      if (ComboBoxWorksheets.Items.Count > 0)
      {
        ComboBoxWorksheets.SelectedIndex = 0;
      }
      _excel.ExcelShutDown();
      FormIsReady();
    }

    /// <summary>
    /// Outlook Template Select
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void ComboBoxTemplates_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
      // Check for %%%% compatibility?
    } 

    /// <summary>
    /// Excel Worksheet Select
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void ComboBoxWorksheet_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
      // Load Datagrid
      try
      {
        KeyValuePair<string, DataTable> m_obj = (KeyValuePair<string, DataTable>) ComboBoxWorksheets.SelectedItem;

        DataTable m_dt = m_obj.Value;
        GridMain.DataContext = m_dt.DefaultView;
        GridMain.AutoGenerateColumns = true;
      } 
      catch (Exception ex)
      {
      }
    }

    #endregion

  }
}