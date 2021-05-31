using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using System.IO;

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
            appPath = Application.StartupPath;


            //fileName = @"\TestODT";
            //loadName = @"\data";

            XmlManager xm = new XmlManager();

            xm.CreateODT();
            xm.Update_Bold();
            xm.SaveODT(xm.root);

            //xm.Create_Document(appPath + loadName);

            //xm.Update_Text("Test Test one two one two");

            //xm.Save_Document(appPath + fileName);
        }

        private void Initial_XML()
        {

        }
    }
}
