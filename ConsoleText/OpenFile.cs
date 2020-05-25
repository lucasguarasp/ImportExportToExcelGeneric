using System;
using System.Diagnostics;
using System.IO;
using System.Security;
using System.Windows.Forms;

public class OpenFile
{
    OpenFileDialog openFileDialog = new OpenFileDialog();
    public OpenFile()
    {
        openFileDialog = new OpenFileDialog()
        {
            FileName = "Select a CSV file",
            Filter = "csv files(*.csv)|*.csv|All files(*.*)|*.*",
            Title = "Open csv file"
        };

    }

    public string GetStreamSelectedPath()
    {
        //openFileDialog.FilterIndex = 2;
        //openFileDialog.RestoreDirectory = true;

        if (openFileDialog.ShowDialog() == DialogResult.OK)
            return openFileDialog.FileName;
            //return openFileDialog.OpenFile();

        return null;
    }



}