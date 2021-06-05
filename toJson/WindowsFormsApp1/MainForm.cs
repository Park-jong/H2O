﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btn_load_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var filePath = ofd.FileName;
                    var size = new FileInfo(filePath).Length;
                    var extension = Path.GetExtension(filePath);

                    MessageBoxButtons button = MessageBoxButtons.OK;

                    using (Stream str = ofd.OpenFile())
                    {
                        if (extension != ".hwp")
                        {
                            MessageBox.Show("Hwp 파일만 불러올 수 있습니다.", "Fail", button);
                        }
                        else if (size > 20971520)
                        {
                            MessageBox.Show("20mb 이하의 파일만 불러올 수 있습니다.", "Fail", button);
                        }
                        else
                        {
                            MessageBox.Show("불러오기 완료.", "Success", button);
                            ExecuteCommandSync(filePath);

                            string path = @"C:\Documents\MasterDocumentFolder\";
                            DirectoryInfo parentDir = Directory.GetParent(filePath);

                            currentPath = parentDir.FullName;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                }
            }
        }
        public void ExecuteCommandSync(Object filepath)
        {
            String path = System.IO.Directory.GetCurrentDirectory();

            try

            {
                // create the ProcessStartInfo using "cmd" as the program to be run,
                // and "/c " as the parameters.
                // Incidentally, /c tells cmd that we want it to execute the command that follows,
                // and then exit.
                System.Diagnostics.ProcessStartInfo procStartInfo =
                    new System.Diagnostics.ProcessStartInfo("java", @"-jar temp44-0.0.1-jar-with-dependencies.jar " + filepath + " " + Environment.NewLine);

                // The following commands are needed to redirect the standard output.
                // This means that it will be redirected to the Process.StandardOutput StreamReader.
                procStartInfo.WorkingDirectory = path;
                procStartInfo.RedirectStandardOutput = true;
                procStartInfo.UseShellExecute = false;
                // Do not create the black window.
                procStartInfo.CreateNoWindow = false;

                // Now we create a process, assign its ProcessStartInfo and start it
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo = procStartInfo;
                proc.Start();
                /*             proc.StandardInput.WriteLine(filepath);
                             proc.StandardInput.Close();
                */             // Get the output into a string
                string result = proc.StandardOutput.ReadToEnd();
                // Display the command output.
                Console.WriteLine(result);

            }
            catch (Exception objException)
            {
                // Log the exception
            }
        }

        string currentPath;
        private void bnt_convert_Click(object sender, EventArgs e)
        {
            FileManager fm = new FileManager(currentPath);
        }
    }
}
