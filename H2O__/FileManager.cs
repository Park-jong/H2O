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

        private void Initial_Folder()
        {
            XmlManager xm = new XmlManager();

            xm.CreateODT();
            xm.Update_Text("Test");
            //xm.Update_Bold();
            xm.SaveODT(xm.root);
            xm.SaveZIP(xm.root);
        }
    }
}
