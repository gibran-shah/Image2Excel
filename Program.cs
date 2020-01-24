using System;

namespace Image2Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            string imageFilename = "";
            string excelFilename = "";

            switch (args.Length) {
                case 0:
                    Console.WriteLine("Please supply an image file name and an optional Excel file name.");
                    Console.WriteLine("Example: dotnet run [path/imagefile.ext [path/excelfile.xls]]");
                    return;
                case 1:
                    imageFilename = args[0];
                    break;
                case 2:
                default:
                    imageFilename = args[0];
                    excelFilename = args[1];
                    break;
            }

            Engine.go(imageFilename, excelFilename);
        }
    }
}

// https://docs.microsoft.com/en-us/dotnet/core/tutorials/with-visual-studio-code

// Add project to github
