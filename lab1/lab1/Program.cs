using System;
using System.IO;

namespace lab1
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {
                startProgram(args[0], args[1], args[2]);
                // startProgram("test_input", "test_output", "xlsx");
            }
            catch (Exception ex)
            {
                Reader.log(ex.Message);
                Console.WriteLine(ex.Message);
            }
        }

        static void startProgram(string inputFile = "", string outputFile = "", string type = "")
        {
            if (!type.Equals("json") &&  !type.Equals("xlsx"))
            {
                Reader.log("You can enter only json or xlsx format");
                Console.WriteLine("You can enter only json or xlsx format");
                return;
            }   

            if(!File.Exists(inputFile + ".csv"))
            {
                Reader.log("Output file not found");
                Console.WriteLine("File " + inputFile + ".csv not found");
                return;
            }

            Reader reader = new Reader();
   
            try
            {
                reader.read(inputFile + ".csv");

                switch(type)
                {
                    case "json":
                        {
                            reader.writeJSON(outputFile + ".json");
                            break;
                        }
                    case "xlsx":
                        {
                            reader.writeXLSX(outputFile + ".xlsx");
                            break;
                        }
                }
            }
            catch(Exception ex)
            {
                Reader.log(ex.Message);
                Console.WriteLine(ex.Message);
            }
        }
    }
}
