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
            
            

            for (int i = 0; i < PictureFiles.Length; i++)
            {
                PowerPoint.Shape shape = slide.Shapes[2];
                

                textRange = slide.Shapes[1].TextFrame.TextRange;
                textRange.Text = "Title of Page ";
                textRange.Font.Name = "Arial";
                textRange.Font.Size = 32;

                textRange = slide.Shapes[2].TextFrame.TextRange;
                textRange.Text = "Content goes here\nYou can add text\nItem 3";

                float otstup = shape.Width / PictureFiles.Length;
                otstup *= 1f;
                slide.Shapes.AddPicture(PictureFiles[i],
                    Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                    otstup, shape.Top, shape.Width, shape.Height);

            }

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
