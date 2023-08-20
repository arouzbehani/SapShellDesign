using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace SapShellDesign
{
    public class GenericSerializer
    {
        public static string Serialize(object input)
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(input.GetType());

                MemoryStream stream = new MemoryStream();

                serializer.Serialize(stream, input);

                byte[] bytes = new byte[(int)stream.Length];
                stream.Position = 0;
                stream.Read(bytes, 0, (int)stream.Length);
                stream.Close();

                return Encoding.UTF8.GetString(bytes); //  Encoding.ASCII.GetString(bytes);
            }
            catch (Exception exc)
            {

                return "SERIALIZATION FAILED!";
            }
        }

        public static bool DeSerialize<T>(string input, ref T output)
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(T));

                MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(input));

                output = (T)serializer.Deserialize(stream);

                return true;
            }
            catch
            {
                return false;
            }
        }

    }

}
