using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public class HwpToOdt
    {
        public HwpToOdt()
        {

        }


        public void Convert(string filePath, string currentPath)
        {
            ExecuteCommandSync(filePath);

            FileManager fm = new FileManager(currentPath);

            string[] files = new string[] { "content.xml", "manifest.xml", "settings.xml", "styles.xml" };
            foreach (string filename in files)
            {
                string rewrite = System.IO.File.ReadAllText(Application.StartupPath + @"\New File\" + filename);
                int index = rewrite.IndexOf("\n");
                string subString1 = rewrite.Substring(0, index + 1);
                string subString2 = rewrite.Substring(index + 1);
                rewrite = subString1 + Regex.Replace(subString2, @">\r\n( )*<", "><");
                //  (, ) 가 json에 (,\n+공백)로 줄바꿈 입력됨 이유는 모름 나중에 수정가능성
                rewrite = Regex.Replace(rewrite, @",\n             ", ", ");
                System.IO.File.WriteAllText(Application.StartupPath + @"\New File\" + filename, rewrite);
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
                    new System.Diagnostics.ProcessStartInfo("java", @"-jar temp44-0.0.1-jar-with-dependencies.jar " + "\"" + filepath + "\" " + Environment.NewLine);

                // The following commands are needed to redirect the standard output.
                // This means that it will be redirected to the Process.StandardOutput StreamReader.
                procStartInfo.WorkingDirectory = path;
                procStartInfo.RedirectStandardOutput = true;
                procStartInfo.UseShellExecute = false;
                // Do not create the black window.
                procStartInfo.CreateNoWindow = true;

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
    }
}
