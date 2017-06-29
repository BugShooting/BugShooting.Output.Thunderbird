namespace BS.Output.Thunderbird
{

  public class Output: IOutput 
  {

    string name;
    string fileName;
    string fileFormat;

    public Output(string name,
                  string fileName,
                  string fileFormat)
    {
      this.name = name;
      this.fileName = fileName;
      this.fileFormat = fileFormat;
    }

    public string Name
    {
      get { return name; }
    }

    public string Information
    {
      get { return string.Empty; }
    }

    public string FileName
    {
      get { return fileName; }
    }

    public string FileFormat
    {
      get { return fileFormat; }
    }

  }
}
