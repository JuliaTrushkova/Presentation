﻿using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Drawing;


namespace Presentation
{    
    internal class MyPresentation
    {
        public static void CreatePresentation(string[] PictureFiles, int countOfRows, int countOfColumns, bool hasTitle)
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
            string filePathSave = Path.GetDirectoryName(PictureFiles[0]) + @"\test.pptx";
            presentation.SaveAs(filePathSave,
                PowerPoint.PpSaveAsFileType.ppSaveAsDefault,
                Office.MsoTriState.msoTrue);

            //создаем шаблон. PowerPoint.PpSlideLayout.ppLayoutText - встроенный формат шаблона как в PowerPoint
            PowerPoint.CustomLayout customLayout = presentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutText];
            
            //Добавляем слайды в презентацию
            PowerPoint.Slides slides = presentation.Slides;
            int slideID = 1;
            PowerPoint._Slide slide = slides.AddSlide(slideID, customLayout);

            //Добавляем титул на слайд. slide.Shapes[1] - здесь 1 - это порядковый номер фигуры. Каждый объект на слайде - это фигура (Shape)
            AddText(slide.Shapes[1], "Title of Page ", "Arial", 32);              

            //Расчет размера номинального блока для каждого рисунка в зависимости от заданного количества столбцов и рядов для картинок
            //просто разбивает область для картинок на слайде по столбцам и рядам
            (float widthOfBlock, float heightOfBlock) = CalculateHeightWidthOfBlock(customLayout, countOfRows, countOfColumns);

            //Расчет максимально возможного размера для отдельной картинки (размера номинального блока за вычетом отступа слева)
            float widthLeftIndent = 20;            
            float widthBlockWithoutIndent = widthOfBlock - widthLeftIndent;

            //Начальный отступ по вертикали (положение верхней границы картинки) и горизонтали (положение левой стороны картинки)
            float initialTopShape = 100;
            float topShape = initialTopShape;
            float indent = widthLeftIndent;

            //Костыль. Первая вставка картинки косячная
            List<PowerPoint.Shape> shapesToDelete = new List<PowerPoint.Shape>();           
            shapesToDelete.Add(AddKostyl(slide, PictureFiles[0]));
            //Конец костыля

            //Высота подписи к картинке
            float heightOfLabel = 25;

            for (int i = 0; i < PictureFiles.Length; i++)
            {               
                //Расчет размера картинки на слайде
                (float widthPictureWork, float heightPictureWork) = CalculateHeightWidthOfPicture(PictureFiles[i], widthBlockWithoutIndent, heightOfBlock);

                //Добавление картинки на слайд
                PowerPoint.Shape shapeLabelPic = slide.Shapes.AddPicture2(PictureFiles[i],
                    Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                    indent, topShape, widthPictureWork, heightPictureWork, Office.MsoPictureCompress.msoPictureCompressFalse);

                //Добавление подписи к картинке
                if (hasTitle)
                {
                    string fileName = Path.GetFileNameWithoutExtension(PictureFiles[i]);

                    PowerPoint.Shape shapeLabel = slide.Shapes.AddLabel(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                                     indent, topShape - heightOfLabel, widthPictureWork, heightOfLabel);
                    
                    AddText(shapeLabel, fileName, "Arial", 16);
                }
                
                //Смещение левой стороны на одну картинку вправо
                indent += widthOfBlock;

                //При превышении номера картинки над количеством столбцов переходит на следующую строку
                if (((i + 1) >= countOfColumns) && ((i + 1) % countOfColumns == 0))
                {
                    indent = widthLeftIndent;
                    topShape += (heightPictureWork + heightOfLabel);
                }

                //При заданном количестве картинок на слайде создает новый слайд
                if (topShape > customLayout.Height)
                {
                    ++slideID;
                    slide = slides.AddSlide(slideID, customLayout);

                    shapesToDelete.Add(AddKostyl(slide, PictureFiles[0]));

                    indent = widthLeftIndent;
                    topShape = initialTopShape;
                }
            }

            //!!!Костыль. Первая вставка картинки косячная, удаляем все первые вставленные картинки со всех слайдов.
            foreach (PowerPoint.Shape shapeToDelete in shapesToDelete)
                shapeToDelete.Delete();
            //Конец костыля

            presentation.Save();
           
            KillProcessesPowerPoint();
        }

        //Расчет размеров блока путем деления слайда на количество строк и столбцов
        private static (float width, float height) CalculateHeightWidthOfBlock(PowerPoint.CustomLayout customLayout, int numberOfRows, int numberOfColumns, float scale = 1)
        {   
            float widthBlock = customLayout.Width / numberOfColumns;
            float heightBlock = (customLayout.Height - 130) * scale / numberOfRows;
            return (widthBlock, heightBlock);
        }        

        //Расчет размеров картинки на слайд
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

        //Добавление текста на слайд
        private static PowerPoint.TextRange AddText(PowerPoint.Shape shape, string text, string fontName, float fontSize, PowerPoint.PpParagraphAlignment alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter)
        {
            shape.TextFrame.WordWrap = Office.MsoTriState.msoTrue;

            PowerPoint.TextRange textRangeLabel = shape.TextFrame.TextRange;

            textRangeLabel.Text = text;
            textRangeLabel.Font.Name = fontName;
            textRangeLabel.Font.Size = fontSize;
            textRangeLabel.ParagraphFormat.Alignment = alignment;   

            return textRangeLabel;
        }

        //Завершение всех процессов Powerpoint
        private static void KillProcessesPowerPoint()
        {
            var processes = System.Diagnostics.Process.GetProcessesByName("POWERPNT");
            foreach (var process in processes)
            {
                process.Kill();
            }
        }

        //Добавляет на слайд певую картинку (она вставится криво). Потом она удаляется
        private static PowerPoint.Shape AddKostyl(PowerPoint._Slide slide, string file)
        {
            return slide.Shapes.AddPicture(file,
                   Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                   50, 50, 50, 50);
        }
    }
}
