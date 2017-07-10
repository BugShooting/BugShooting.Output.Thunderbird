using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using Microsoft.Win32;

namespace BS.Output.Thunderbird
{
  public class OutputAddIn: V3.OutputAddIn<Output>
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
                                 String.Empty,
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
                          edit.FileFormat,
                          edit.EditFileName);
      }
      else
      {
        return null;
      }

    }

    protected override OutputValueCollection SerializeOutput(Output Output)
    {

      OutputValueCollection outputValues = new OutputValueCollection();

      outputValues.Add(new OutputValue("Name", Output.Name));
      outputValues.Add(new OutputValue("FileName", Output.FileName));
      outputValues.Add(new OutputValue("FileFormat", Output.FileFormat));
      outputValues.Add(new OutputValue("EditFileName", Output.EditFileName.ToString()));

      return outputValues;

    }

    protected override Output DeserializeOutput(OutputValueCollection OutputValues)
    {
      return new Output(OutputValues["Name", this.Name].Value,
                        OutputValues["FileName", "Screenshot"].Value,
                        OutputValues["FileFormat", ""].Value,
                        Convert.ToBoolean(OutputValues["EditFileName", false.ToString()].Value));
    }

    protected async override Task<V3.SendResult> Send(Output Output, V3.ImageData ImageData)
    {
      try
      {

        string fileName = V3.FileHelper.GetFileName(Output.FileName, Output.FileFormat, ImageData);

        if (Output.EditFileName)
        {

          Send send = new Send(fileName);

          if (send.ShowDialog() != true)
          {
            return new V3.SendResult(V3.Result.Canceled);
          }

          fileName = send.FileName;

        }

        string filePath = Path.Combine(Path.GetTempPath(), fileName + "." + V3.FileHelper.GetFileExtention(Output.FileFormat));

        Byte[] fileBytes = V3.FileHelper.GetFileBytes(Output.FileFormat, ImageData);

        using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite))
        {
          file.Write(fileBytes, 0, fileBytes.Length);
          file.Close();
        }

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
          return new V3.SendResult(V3.Result.Failed, "Thunderbird is not installed.");
        }
        
        Process.Start(applicationPath, "-compose \"attachment='" + filePath + "'\"");

        return new V3.SendResult(V3.Result.Success);
                
      }
      catch (Exception ex)
      {
        return new V3.SendResult(V3.Result.Failed, ex.Message);
      }
      
    }
      
  }

}