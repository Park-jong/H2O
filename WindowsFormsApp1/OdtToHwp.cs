﻿using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace WindowsFormsApp1
{
    class OdtToHwp
    {
        public void Convert(string filePath, string currentPath)
        {
            Directory.CreateDirectory(currentPath + @"\OdtToHwp");
            try
            {
                ZipFile.ExtractToDirectory(filePath, currentPath + @"\OdtToHwp");
            }
            catch (IOException e)
            {
                Directory.Delete(currentPath + @"\OdtToHwp", true);
                ZipFile.ExtractToDirectory(filePath, currentPath + @"\OdtToHwp");
            }
            string xmlContent = currentPath + @"\OdtToHwp\content.xml";
            string xmlStyles = currentPath + @"\OdtToHwp\styles.xml";
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlContent);

            string json;
            json = JsonConvert.SerializeXmlNode(doc);

            File.WriteAllText(currentPath+@"\content.json", json);

            doc.Load(xmlStyles);
            json = JsonConvert.SerializeXmlNode(doc);
            File.WriteAllText(currentPath + @"\styles.json", json);
        }
    }
}
