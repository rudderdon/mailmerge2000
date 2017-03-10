using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace MailMerge2000.Data
{

  /// <summary>
  /// General documentation can be found at:
  /// https://msdn.microsoft.com/en-us/library/office/bb612741.aspx
  /// </summary>
  public class OutlookApp
  {

    private int _processId = 0;

    internal Application App;

    /// <summary>
    /// Constructor
    /// </summary>
    public OutlookApp()
    {
      OutlookStart();
    }

    /// <summary>
    /// Email Processing
    /// </summary>
    /// <param name="templatePath"></param>
    /// <param name="subject"></param>
    /// <param name="toMail"></param>
    /// <param name="ccMail"></param>
    /// <param name="bccMail"></param>
    /// <param name="values"></param>
    /// <param name="attachFilePaths"></param>
    /// <param name="sendMail"></param>
    /// <returns></returns>
    internal bool EmailFromTemplate(
      string templatePath, 
      string subject, 
      string toMail, 
      string ccMail, 
      string bccMail, 
      Dictionary<string, string> values, 
      List<string> attachFilePaths,
      bool sendMail)
    {
      try
      {
        Folder m_draftFolder = App.Session.GetDefaultFolder(
          OlDefaultFolders.olFolderDrafts) as Folder;
        MailItem m_mail = App.CreateItemFromTemplate(templatePath, m_draftFolder) as MailItem;
        if (m_mail != null)
        {
          
          m_mail.Subject = subject;
          m_mail.To = toMail;
          m_mail.CC = ccMail;
          m_mail.BCC = bccMail;

          foreach (var x in attachFilePaths)
          {
            m_mail.Attachments.Add(x);
          }

          // Replace %%%% contents with matches in the body
          StringBuilder m_sb = new StringBuilder(m_mail.HTMLBody);
          foreach (var x in values)
          {
            m_sb.Replace(string.Format("%%{0}%%", x.Key.ToUpper()), x.Value);
          }
          m_mail.HTMLBody = m_sb.ToString();

          if (sendMail)
            m_mail.Send();
          else
            m_mail.Save();

        }
        return true;
      }
      catch (Exception ex)
      {
        // way too lazy to do anything here
      }
      return false;
    }

    /// <summary>
    /// Start
    /// </summary>
    /// <returns></returns>
    internal bool OutlookStart()
    {
      try
      {
        // Store the Excel processes before opening.
        Process[] m_processesBefore = Process.GetProcessesByName("outlook");
        App = new Application();
        Process[] m_processesAfter = Process.GetProcessesByName("outlook");
        foreach (Process process in m_processesAfter)
        {
          if (!m_processesBefore.Select(p => p.Id).Contains(process.Id))
          {
            _processId = process.Id;
          }
        }
        return true;
      }
      catch (Exception ex)
      {
        // Still way too lazy to do anything here
      }
      return false;
    }

  }
}