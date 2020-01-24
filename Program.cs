using System;

namespace Image2Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            string imageFile = "";
            string excelFilename = "";

            switch (args.Length) {
                case 0:
                    Console.WriteLine("Please supply an image file name and an optional Excel file name.");
                    Console.WriteLine("Example: dotnet run [imagefile.ext [excelfile.xls]]");
                    return;
                case 1:
                    imageFile = args[0];
                    break;
                case 2:
                default:
                    imageFile = args[0];
                    excelFilename = args[1];
                    break;
            }
        }
    }
}

// https://docs.microsoft.com/en-us/dotnet/core/tutorials/with-visual-studio-code

// Add project to github
