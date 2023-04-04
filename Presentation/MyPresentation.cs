using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Drawing;



namespace Presentation
{
    internal class MyPresentation
    {
        public static void CreatePresentation()
        {
            //Создаем объект приложения PowerPoint, в который потом будем добавлять презентации 
            PowerPoint.Application PPApp = new PowerPoint.Application();

            //Создаем презентацию добавляя в коллекцию нашего приложения
            //Чтобы был доступ к Office.MsoTriState нужно напрямую добавить библиотеку office из папки C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c. Иначе будет ошибка, что нужна библиотека версии 15.0.0.0
            //Просто установка нугета Microsoft.Office.Interop.PowerPoint и COM'а Microsoft.Office16.Objects не помогает
            //MsoTriState - не понятно зачем, пробовала разные значения, презентация одинаковая получается

            PowerPoint.Presentation presentation = PPApp.Presentations.Add(Office.MsoTriState.msoFalse);

            //Можно добавить существующую: PowerPoint.Presentation presentationTemplate = PPApp.Presentations.Open(filePath);

            //сохраняем презентацию
            //string filePathSave = @"C:\Users\trushkova\Desktop\test\test.pptx";
            //presentation.SaveAs(filePathSave,
            //    PowerPoint.PpSaveAsFileType.ppSaveAsDefault,
            //    Office.MsoTriState.msoTrue);

            string filePathSave = @"C:\Users\abc\Desktop\test\test.pptx";
            presentation.SaveAs(filePathSave,
                PowerPoint.PpSaveAsFileType.ppSaveAsDefault,
                Office.MsoTriState.msoFalse);

            //string[] PictureFiles = { @"C:\Users\trushkova\Desktop\test\26_CUBE_Ampl_-400-2000Average.png",
            //@"C:\Users\trushkova\Desktop\test\26_CUBE_DF_-400-2000Average.png",
            //@"C:\Users\trushkova\Desktop\test\26_CUBE_DFInst_-400-2000Average.png" };
            string[] PictureFiles = { @"C:\Users\abc\Desktop\test\picture2.jpg",
            @"C:\Users\abc\Desktop\test\picture3.jpg",
            @"C:\Users\abc\Desktop\test\picture4.jpg" };
            //string[] PictureFiles = { @"C:\Users\abc\Desktop\test\2.jpg",
            //@"C:\Users\abc\Desktop\test\3.jpg"            };
          

            PowerPoint.CustomLayout customLayout = presentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutText];
            
            PowerPoint.Slides slides = presentation.Slides;
            PowerPoint._Slide slide = slides.AddSlide(1, customLayout);

            AddText(slide.Shapes[1], "Title of Page ", "Arial", 32);              

            (float widthOfBlock, float heightOfBlock) = CalculateHeightWidthOfBlock(customLayout, 1, PictureFiles.Length);

            float widthLeftIndent = 20;            
            float widthBlockWithoutIndent = widthOfBlock - widthLeftIndent;
            
            float topShape = 100;
            float indent = widthLeftIndent;

            //костыль. Первая вставка картинки косячная
            (float widthPictureWork, float heightPictureWork) = CalculateHeightWidthOfPicture(PictureFiles[0], widthBlockWithoutIndent, heightOfBlock);
            slide.Shapes.AddPicture(PictureFiles[0],
                    Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                    indent, topShape, widthPictureWork, heightPictureWork);
            //конец костыля
            

            for (int i = 0; i < PictureFiles.Length; i++)
            {
                (widthPictureWork, heightPictureWork) = CalculateHeightWidthOfPicture(PictureFiles[i], widthBlockWithoutIndent, heightOfBlock);
                
                slide.Shapes.AddPicture(PictureFiles[i],
                    Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                    indent, topShape, widthPictureWork, heightPictureWork);

                string fileName = Path.GetFileNameWithoutExtension(PictureFiles[i]);

                PowerPoint.Shape shapeLabel = slide.Shapes.AddLabel(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    indent, topShape - 25, widthPictureWork, 25);
                AddText(shapeLabel, fileName, "Arial", 16);
                
                indent += widthOfBlock;

            }
            //костыль. Первая вставка картинки косячная
            slide.Shapes[2].Delete();
            //конец костыля

            presentation.Save();

           KillProcessesPowerPoint();
        }

        private static (float width, float height) CalculateHeightWidthOfBlock(PowerPoint.CustomLayout customLayout, int numberOfRows, int numberOfColumns, float scale = 1)
        {   
            float widthBlock = customLayout.Width / numberOfColumns;
            float heightBlock = (customLayout.Height - 120) * scale / numberOfRows;
            return (widthBlock, heightBlock);
        }        

        private static (float width, float height) CalculateHeightWidthOfPicture(string fileOfPicture, float widthOfBlock, float heightOfBlock)
        {    
            Image image = Image.FromFile(fileOfPicture);
            int heightInitialPicture = image.Height;
            int widthInitialPicture = image.Width;
            
            float heightPictureWork = heightOfBlock;
            float widthPictureWork = widthInitialPicture * heightPictureWork / heightInitialPicture;

            if (widthPictureWork > widthOfBlock)
            {
                widthPictureWork = widthOfBlock;
                heightPictureWork = heightInitialPicture * widthOfBlock / widthInitialPicture;
            }

            return (widthPictureWork, heightPictureWork);
        }

        private static PowerPoint.TextRange AddText(PowerPoint.Shape shape, string text, string fontName, float fontSize, PowerPoint.PpParagraphAlignment alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter)
        {
            PowerPoint.TextRange textRangeLabel = shape.TextFrame.TextRange;

            textRangeLabel.Text = text;
            textRangeLabel.Font.Name = fontName;
            textRangeLabel.Font.Size = fontSize;
            textRangeLabel.ParagraphFormat.Alignment = alignment;

            return textRangeLabel;
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
