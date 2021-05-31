using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//using System.Windows.Forms;
using System.IO;
using System.IO.Compression;

namespace H2O__
{
    class FileManager
    {
        public FileManager()
        {
            Initial_Folder();
        }

        private string appPath;
        private string fileName;
        private string loadName;

        private void Initial_Folder()
        {
            appPath = @"C:\Users\park\source\repos\H2O_odf\H2O_odf";


            //fileName = @"\TestODT";
            //loadName = @"\data";

            XmlManager xm = new XmlManager();

            xm.CreateODT();
            xm.AddContentP("asd");
            xm.AddContentP("qwe");
            xm.SetFontColor("P1", "#FF0000");
            xm.SetFont("P2", "궁서");
            xm.SaveODT(xm.root);


            try
            {
                ZipFile.CreateFromDirectory(appPath + @"\New File", appPath + @"\New.odt");
            }
            catch (IOException)
            {
                File.Delete(appPath + @"\New.odt");
                ZipFile.CreateFromDirectory(appPath + @"\New File", appPath + @"\New.odt");
            }

            //xm.Create_Document(appPath + loadName);

            //xm.Update_Text("Test Test one two one two");

            //xm.Save_Document(appPath + fileName);
        }

        private void Initial_XML()
        {

        }
    }
}
