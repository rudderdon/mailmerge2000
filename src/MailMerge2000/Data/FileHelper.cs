using System.IO;

namespace MailMerge2000.Data
{

  /// <summary>
  /// Just a convenient way to deal with files in UI binding
  /// </summary>
  public class FileHelper
  {

    public FileInfo File { get; set; }

    public string Name
    {
      get { return File.Name; }
    }
    public string FullName
    {
      get { return File.FullName; }
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="f"></param>
    public FileHelper(FileInfo f)
    {
      File = f;
    }

  }
}