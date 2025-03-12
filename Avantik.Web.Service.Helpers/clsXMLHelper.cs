using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Runtime.Serialization.Json;

namespace Avantik.Web.Service.Helpers
{
    public static class XMLHelper
    {
        public static string Serialize(object o, bool withXmlHeader)
        {
            if (o != null)
            {
                XmlSerializer s = new XmlSerializer(o.GetType());
                XmlSerializer serializer = new XmlSerializer(o.GetType());
                XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                ns.Add("", "");

                XmlWriterSettings writerSettings = new XmlWriterSettings();
                if (withXmlHeader == true)
                { writerSettings.OmitXmlDeclaration = false; }
                else
                { writerSettings.OmitXmlDeclaration = true; }

                StringWriter writer = new StringWriter();
                using (XmlWriter xmlWriter = XmlWriter.Create(writer, writerSettings))
                {
                    serializer.Serialize(xmlWriter, o, ns);
                }

                return writer.ToString();
            }

            return string.Empty;

        }
        public static object Deserialize(string xml, Type t)
        {
            try
            {
                if (!string.IsNullOrEmpty(xml))
                {
                    xml = xml.Replace("encoding=\"utf-16\"", "encoding=\"utf-8\"");
                    using (StringReader reader = new StringReader(xml))
                    {
                        XmlSerializer serializer = new XmlSerializer(t);
                        return serializer.Deserialize(reader);
                    }
                }
                else
                    return null;
            }
            catch
            {
                return null;
            }

        }

        public static string JsonSerializer(Type classReq, Object objectReq)
        {
            try
            {
                string returnStr = string.Empty;
                DataContractJsonSerializer js = new DataContractJsonSerializer(classReq);
                MemoryStream ms = new MemoryStream();
                js.WriteObject(ms, objectReq);
                ms.Position = 0;
                StreamReader sr = new StreamReader(ms);
                returnStr = sr.ReadToEnd().ToString();
                sr.Dispose();
                sr.Close();
                ms.Dispose();
                ms.Close();
                return returnStr;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
