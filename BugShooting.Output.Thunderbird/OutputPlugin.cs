using BS.Plugin.V3.Common;
using BS.Plugin.V3.Output;
using BS.Plugin.V3.Utilities;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BugShooting.Output.Thunderbird
{
  public class OutputPlugin: OutputPlugin<Output>
  {

    protected override string Name
    {
      get { return "Thunderbird"; }
    }

    protected override Image Image64
    {
      get  { return Properties.Resources.logo_64; }
    }

    protected override Image Image16
    {
      get { return Properties.Resources.logo_16 ; }
    }

    protected override bool Editable
    {
      get { return true; }
    }

    protected override string Description
    {
      get { return "Attach screenshots to your Thunderbird emails."; }
    }
    
    protected override Output CreateOutput(IWin32Window Owner)
    {

      Output output = new Output(Name,
                                 "Screenshot",
                                 FileHelper.GetFileFormats().First().ID,
                                 false);

      return EditOutput(Owner, output);

    }

    protected override Output EditOutput(IWin32Window Owner, Output Output)
    {

      Edit edit = new Edit(Output);

      var ownerHelper = new System.Windows.Interop.WindowInteropHelper(edit);
      ownerHelper.Owner = Owner.Handle;

      if (edit.ShowDialog() == true)
      {

        return new Output(edit.OutputName,
                          edit.FileName,
                          edit.FileFormatID,
                          edit.EditFileName);
      }
      else
      {
        return null;
      }

    }

    protected override OutputValues SerializeOutput(Output Output)
    {

      OutputValues outputValues = new OutputValues();

      outputValues.Add("Name", Output.Name);
      outputValues.Add("FileName", Output.FileName);
      outputValues.Add("FileFormatID", Output.FileFormatID.ToString());
      outputValues.Add("EditFileName", Output.EditFileName.ToString());

      return outputValues;

    }

    protected override Output DeserializeOutput(OutputValues OutputValues)
    {
      return new Output(OutputValues["Name", this.Name],
                        OutputValues["FileName", "Screenshot"],
                        new Guid(OutputValues["FileFormatID", ""]),
                        Convert.ToBoolean(OutputValues["EditFileName", false.ToString()]));
    }

    protected async override Task<SendResult> Send(IWin32Window Owner, Output Output, ImageData ImageData)
    {
      try
      {

        string applicationPath = string.Empty;

        using (RegistryKey localMachineKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32))
        {
          using (RegistryKey key = localMachineKey.OpenSubKey("Software\\Mozilla\\Mozilla Thunderbird", false))
          {
            if (key != null)
            {
              string currentVersion = Convert.ToString(key.GetValue("CurrentVersion", string.Empty));

              using (RegistryKey pathKey = localMachineKey.OpenSubKey("Software\\Mozilla\\Mozilla Thunderbird\\" + currentVersion + "\\Main", false))
              {
                if (pathKey != null)
                  applicationPath = Convert.ToString(pathKey.GetValue("PathToExe", string.Empty));
              }
            }
          }
        }

        if (!File.Exists(applicationPath))
        {
          return new SendResult(Result.Failed, "Thunderbird is not installed.");
        }


        string fileName = AttributeHelper.ReplaceAttributes(Output.FileName, ImageData);

        if (Output.EditFileName)
        {

          Send send = new Send(fileName);

          var ownerHelper = new System.Windows.Interop.WindowInteropHelper(send);
          ownerHelper.Owner = Owner.Handle;
       
          if (send.ShowDialog() != true)
          {
            return new SendResult(Result.Canceled);
          }

          fileName = send.FileName;

        }

        IFileFormat fileFormat = FileHelper.GetFileFormat(Output.FileFormatID);

        string filePath = Path.Combine(Path.GetTempPath(), fileName + "." + fileFormat.FileExtension);

        Byte[] fileBytes = FileHelper.GetFileBytes(Output.FileFormatID, ImageData);

        using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite))
        {
          file.Write(fileBytes, 0, fileBytes.Length);
          file.Close();
        }
        
        Process.Start(applicationPath, "-compose \"attachment='" + filePath + "'\"");

        return new SendResult(Result.Success);
                
      }
      catch (Exception ex)
      {
        return new SendResult(Result.Failed, ex.Message);
      }
      
    }
      
  }

}