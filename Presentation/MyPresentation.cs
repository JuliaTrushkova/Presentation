using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Drawing;



namespace Presentation
{
    internal class MyPresentation
    {
        public static void CreatePresentation()
        {
            PowerPoint.Application PPApp = new PowerPoint.Application();
            string filePath = @"C:\Users\trushkova\Desktop\MiMGO_short_profile_2022_00000002.pptx";
            //PPApp.Visible = Microsoft.Office.Core.MsoTriState.msoCTrue;
            PowerPoint.Presentation presentation = PPApp.Presentations.Add(Office.MsoTriState.msoTrue);
           // PowerPoint.Presentation presentationTemplate = PPApp.Presentations.Open(filePath);
            string filePathSave = @"C:\Users\trushkova\Desktop\test\test.pptx";
            presentation.SaveAs(filePathSave,
                PowerPoint.PpSaveAsFileType.ppSaveAsDefault,
                Office.MsoTriState.msoTrue);

            string[] PictureFiles = { @"C:\Users\trushkova\Desktop\test\26_CUBE_Ampl_-400-2000Average.png",
            @"C:\Users\trushkova\Desktop\test\26_CUBE_DF_-400-2000Average.png",
            @"C:\Users\trushkova\Desktop\test\26_CUBE_DFInst_-400-2000Average.png" };
            //string[] PictureFiles = { @"C:\Users\abc\Desktop\test\2.jpg",
            //@"C:\Users\abc\Desktop\test\3.jpg",
            //@"C:\Users\abc\Desktop\test\4.jpg" };
            //string[] PictureFiles = { @"C:\Users\abc\Desktop\test\2.jpg",
            //@"C:\Users\abc\Desktop\test\3.jpg"            };
            
            PowerPoint.Slides slides;
            PowerPoint._Slide slide;
            PowerPoint.TextRange textRange;

            PowerPoint.CustomLayout customLayout = presentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutText];
            //PowerPoint.CustomLayout customLayout = presentationTemplate.SlideMaster.CustomLayouts[3];
            slides = presentation.Slides;
            slide = slides.AddSlide(1, customLayout);

            float height = customLayout.Height;
            float width = customLayout.Width;

            float widthBlock = customLayout.Width / PictureFiles.Length;
            float widthLeftIndent = 10;
            
            float widthBlockWithoutIndent = widthBlock - widthLeftIndent;

            Image image = Image.FromFile(PictureFiles[0]);
            int heightPicture = image.Height;
            int widthPicture = image.Width;

            //float indent = widthLeftIndent;

            PowerPoint.Shape shape = slide.Shapes[2];

            textRange = slide.Shapes[1].TextFrame.TextRange;
            textRange.Text = "Title of Page ";
            textRange.Font.Name = "Arial";
            textRange.Font.Size = 32;

            float widthInitShape = widthBlockWithoutIndent;
            float heightPictureWork = shape.Height;

            float widthPictureWork = widthPicture * heightPictureWork / heightPicture;  

            if (widthPictureWork > widthBlockWithoutIndent)
            {
                widthPictureWork = widthBlockWithoutIndent;
                heightPictureWork = heightPicture * widthBlockWithoutIndent / widthPicture;
            }  

            float topShape = shape.Top;
            float indent = widthLeftIndent;
            float indentBlock = shape.Left;
            slide.Shapes.AddPicture(PictureFiles[0],
                    Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                    indent, topShape, widthPictureWork, heightPictureWork);
            int i = 0;

            for (int i = 0; i < PictureFiles.Length; i++)
            {

                //textRange = slide.Shapes[2].TextFrame.TextRange;
                //textRange.Text = "Content goes here\nYou can add text\nItem 3";
               // PowerPoint.Shape shapePic = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, indent, topShape, widthBlockWithoutIndent, heightPictureWork);

                //slide.Shapes.AddPicture(PictureFiles[i],
                //    Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                //    indent, topShape, widthInitShape, heightInitShape);
                
                slide.Shapes.AddPicture(PictureFiles[i],
                    Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                    indent, topShape, widthPictureWork, heightPictureWork);

                indent += widthBlock;

            }
            slide.Shapes[2].Delete();
            presentation.Save();

           KillProcessesPowerPoint();
        }

        private static void KillProcessesPowerPoint()
        {
            var processes = System.Diagnostics.Process.GetProcessesByName("POWERPNT");
            foreach (var process in processes)
            {
                process.Kill();
            }
        }

    }
}
