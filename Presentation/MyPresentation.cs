using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace Presentation
{
    internal class MyPresentation
    {        
        public static void CreatePresantation()
        {
            PowerPoint.Application PPApp = new PowerPoint.Application();
            string filePath = @"C:\Users\trushkova\Desktop\MiMGO_short_profile_2022_00000002.pptx";
            //PPApp.Visible = Microsoft.Office.Core.MsoTriState.msoCTrue;
            PowerPoint.Presentation presentation = PPApp.Presentations.Add(Office.MsoTriState.msoTrue);
            //PowerPoint.Presentation presentationTemplate = PPApp.Presentations.Open(filePath);
            string filePathSave = @"C:\Users\trushkova\Desktop\newTestPic.pptx";
            presentation.SaveAs(filePathSave, 
                PowerPoint.PpSaveAsFileType.ppSaveAsDefault, 
                Office.MsoTriState.msoTrue);

            string[] PictureFiles = { @"C:\Users\trushkova\Desktop\test\26_CUBE_Ampl_-400-2000Average.png",
            @"C:\Users\trushkova\Desktop\test\26_CUBE_DF_-400-2000Average.png",
            @"C:\Users\trushkova\Desktop\test\26_CUBE_DFInst_-400-2000Average.png" };

            for (int i = 0; i < PictureFiles.Length; i++)
            {
                PowerPoint.Slides slides;
                PowerPoint._Slide slide;
                PowerPoint.TextRange textRange;

                PowerPoint.CustomLayout customLayout = presentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutText];

                slides = presentation.Slides;
                slide = slides.AddSlide(i+1, customLayout);

                textRange = slide.Shapes[1].TextFrame.TextRange;
                textRange.Text = "Title of Page " + i;
                textRange.Font.Name = "Arial";
                textRange.Font.Size = 32;

                PowerPoint.Shape shape = slide.Shapes[2];
                slide.Shapes.AddPicture(PictureFiles[i], 
                    Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 
                    shape.Left, shape.Top, shape.Width, shape.Height);
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
