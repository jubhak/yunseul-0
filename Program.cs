using System;
using System.IO;
using System.Windows.Forms;

namespace DataParser;

static class Program
{
    [STAThread]
    static void Main()
    {
        try
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }
        catch (Exception ex)
        {
            string logPath = Path.Combine(AppContext.BaseDirectory, "crash.txt");
            File.WriteAllText(logPath, DateTime.Now + "\n" + ex.ToString());
            MessageBox.Show("오류 발생:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
