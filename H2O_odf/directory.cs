using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace H2O_odf
{
    class Configuration
    {
        String path;
        public Configuration(String path)
        {
            this.path = path + @"\Configurations2";
        }
        public void createDirectory()
        {
            DirectoryInfo conf = new DirectoryInfo(path);
            if (conf.Exists)
            {
                conf.Delete(true);
            }
            conf.Create();
            Directory.CreateDirectory(path + @"\accelerator");
            Directory.CreateDirectory(path + @"\floater");
            Directory.CreateDirectory(path + @"\images");
            Directory.CreateDirectory(path + @"\images\Bitmaps");
            Directory.CreateDirectory(path + @"\menubar");
            Directory.CreateDirectory(path + @"\popupmenu");
            Directory.CreateDirectory(path + @"\progressbar");
            Directory.CreateDirectory(path + @"\statusbar");
            Directory.CreateDirectory(path + @"\toolbar");
            Directory.CreateDirectory(path + @"\toolpanel");

        }
    }

    class Metainf
    {
        string path;
        string str = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n" +
                "<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.3\" xmlns:loext=\"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0\">\r\n" +
                " <manifest:file-entry manifest:full-path=\"/\" manifest:version=\"1.3\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/>\r\n" +
                " <manifest:file-entry manifest:full-path=\"Configurations2/\" manifest:media-type=\"application/vnd.sun.xml.ui.configuration\"/>\r\n" +
                " <manifest:file-entry manifest:full-path=\"manifest.rdf\" manifest:media-type=\"application/rdf+xml\"/>\r\n" +
                " <manifest:file-entry manifest:full-path=\"styles.xml\" manifest:media-type=\"text/xml\"/>\r\n" +
                " <manifest:file-entry manifest:full-path=\"meta.xml\" manifest:media-type=\"text/xml\"/>\r\n" +
                " <manifest:file-entry manifest:full-path=\"settings.xml\" manifest:media-type=\"text/xml\"/>\r\n" +
                " <manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/>\r\n" +
                " <manifest:file-entry manifest:full-path=\"Thumbnails/thumbnail.png\" manifest:media-type=\"image/png\"/>\r\n" +
                "</manifest:manifest>";
        public Metainf(String path)
        {
            this.path = path + @"\META-INF";
        }
        public void createDirectory()
        {
            DirectoryInfo meta = new DirectoryInfo(path);
            if (meta.Exists)
            {
                meta.Delete(true);
            }
            meta.Create();
            StreamWriter writer = new StreamWriter(path + @"\manifest.xml");
            writer.Write(str);
            writer.Close();
        }

    }
    class Thumbnails
    {
        String path;
        public Thumbnails(String path)
        {
            this.path = path + @"\Thumbnails";
        }
        public void createDirectory()
        {
            DirectoryInfo thumbnail = new DirectoryInfo(path);
            if (thumbnail.Exists)
            {
                thumbnail.Delete(true);
            }
            thumbnail.Create();
        }
    }
}
