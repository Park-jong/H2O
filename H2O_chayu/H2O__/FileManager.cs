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
            
            
            fileName = @"\TestODT";
            loadName = @"\data";

            Create_Folder();

            XmlManager xm = new XmlManager();
            xm.Create_Document(appPath + loadName);

            xm.Update_Text("Test Test one two one two");

            xm.Save_Document(appPath + fileName);
        }

        private void Create_Folder()
        {
            //Root
            DirectoryInfo di = new DirectoryInfo(appPath + fileName);
            if (!di.Exists) { di.Create(); }

            //Configurations2
            di = new DirectoryInfo(appPath + fileName + @"\Configurations2");
            if (!di.Exists) { di.Create(); }

            //accelerator
            di = new DirectoryInfo(appPath + fileName + @"\Configurations2" + @"\accelerator");
            if (!di.Exists) { di.Create(); }
            //floater
            di = new DirectoryInfo(appPath + fileName + @"\Configurations2" + @"\floater");
            if (!di.Exists) { di.Create(); }

            //images
            di = new DirectoryInfo(appPath + fileName + @"\Configurations2" + @"\images");
            if (!di.Exists) { di.Create(); }
            //Bitmaps
            di = new DirectoryInfo(appPath + fileName + @"\Configurations2" + @"\images" + @"\Bitmaps");

            if (!di.Exists) { di.Create(); }
            //menubar
            di = new DirectoryInfo(appPath + fileName + @"\Configurations2" + @"\menubar");
            if (!di.Exists) { di.Create(); }
            //popupmenu
            di = new DirectoryInfo(appPath + fileName + @"\Configurations2" + @"\popupmenu");
            if (!di.Exists) { di.Create(); }
            //progressbar
            di = new DirectoryInfo(appPath + fileName + @"\Configurations2" + @"\progressbar");
            if (!di.Exists) { di.Create(); }
            //statusbar
            di = new DirectoryInfo(appPath + fileName + @"\Configurations2" + @"\statusbar");
            if (!di.Exists) { di.Create(); }
            //toolbar
            di = new DirectoryInfo(appPath + fileName + @"\Configurations2" + @"\toolbar");
            if (!di.Exists) { di.Create(); }
            //toolpanel
            di = new DirectoryInfo(appPath + fileName + @"\Configurations2" + @"\toolpanel");
            if (!di.Exists) { di.Create(); }


            //META_INF
            di = new DirectoryInfo(appPath + fileName + @"\META-INF");
            if (!di.Exists) { di.Create(); }
            //Thumbnails
            di = new DirectoryInfo(appPath + fileName + @"\Thumbnails");
            if (!di.Exists) { di.Create(); }
        }

        private void Initial_XML()
        {

        }
    }
}
