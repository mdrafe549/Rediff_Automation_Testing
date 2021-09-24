
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace rediff_minni_project
{
    class TextFileWriteRead
    {
        
            public void DirectoryOperation()
            {
            string path = @"C:\Users\mdraf\Desktop\rediff minni project\data1.text";

                if (File.Exists(path))
                {
                    File.Delete(path);
                    Console.WriteLine("File deleted");
                }

                using (StreamWriter sw = File.CreateText(path))
                {
                    Console.WriteLine("New Text File Created!!!!!");
                    sw.WriteLine("https://money.rediff.com/index.html");
                }

                using (StreamReader rdr = File.OpenText(path))
                {
                    Console.WriteLine("Reading from file");
                    string message = string.Empty;
                    Console.WriteLine(rdr.ReadToEnd());
                }

            }

            

       
    }
}
