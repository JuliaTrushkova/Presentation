using Presentation;

Console.WriteLine("Hello, World!");



string path = @"C:\Users\trushkova\Desktop\test";

string[] files = Directory.GetFiles(path, "*.png");


MyPresentation.CreatePresentation(files, 2, 3);




