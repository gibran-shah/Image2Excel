using System;
using System.Drawing;

namespace Image2Excel {
    class Engine {
        public static void go(string imageFilename, string excelFilename = null) {
            Image img = Image.FromFile(imageFilename);

            if (img != null) {
                Console.WriteLine("img good");
            } else {
                Console.WriteLine("img bad");
            }
        } 
    }
}