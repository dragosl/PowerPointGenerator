using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using PowerPointGenerator.Model;

namespace PowerPointGenerator.Helpers
{
    /// <summary>
    /// Provides ppt interaction functionalities using Aspose.Slides library
    /// </summary>
    public static class PptHelper
    {
        /// <summary>
        /// String for new line insert.
        /// </summary>
        private const string NewLineStringConstant = "\r\n";

        /// <summary>
        /// String for component indent.
        /// </summary>
        private const string ComponentExportIndentStringConstant = "\t- ";

        public static bool InsertSalesInTemplate(List<Sale> sales, string templatePath, string savedPptPath)
        {
            try
            {
                // It seems Presentation constructor does not support pptx 2007 template files...
                System.IO.FileStream file = new System.IO.FileStream(templatePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                Presentation presentation = new Presentation(templatePath);
                file.Close();
                SlideCollection slides = presentation.Slides;

                // write top 20 sales in the first textbox (textholder) of the first slide and rotate it by 90 degrees
                // this is used with 20 only because 700 are too much in a textbox (20 are also too much, but still better)
                WriteTopSales(sales, slides);                

                // write the current date in all "places" which support text
                WriteDateEverywhere(slides);

                Stream randomStreamForShapes = StreamHelper.GenerateRandomStream();

                // add a new empty slide and append a note to it
                Slide bodySlide = AddNewSlideAndNote(presentation, randomStreamForShapes);

                // add a link to a background and append the link to the notes
                AddBackgroundLinkToSlide(bodySlide);

                // add shapes to the newly added slide
                AddShapesToSlide(bodySlide, randomStreamForShapes);

                // save file - open and save with user interaction not found yet in the library...
                presentation.Save(savedPptPath, Aspose.Slides.Export.SaveFormat.Ppt);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Writes the top sales and rotates it 90 degrees.
        /// </summary>
        /// <param name="sales">The sales.</param>
        /// <param name="slides">The slides.</param>
        private static void WriteTopSales(List<Sale> sales, SlideCollection slides)
        {
            if (slides.Count > 0)
            {
                PlaceholderCollection places = slides[0].Placeholders;
                if (places.Count > 0)
                {
                    TextHolder text = places[0] as TextHolder;
                    if (text != null)
                    {
                        text.Text = string.Empty;

                        //foreach (Sale sale in sales)
                        for (int i = 0; i < 20; i++)
                        {
                            text.Text += sales[i].ToString() + NewLineStringConstant;
                        }

                        text.RotateTextBy90Degrees = true;
                        text.FitTextToShape();
                    }
                }
            }
        }

        /// <summary>
        /// Writes the current date in every control of the template which supports text.
        /// </summary>
        /// <param name="slides">The slides.</param>
        private static void WriteDateEverywhere(SlideCollection slides)
        {
            foreach (Slide slide in slides)
            {
                // Set text for text frames
                ShapeCollection shapes = slide.Shapes;
                foreach (Shape shape in shapes)
                {
                    TextFrame tf = shape.TextFrame;
                    if (tf != null)
                    {
                        // write it in the first paragraph
                        if (tf.Paragraphs.Count > 0)
                        {
                            if (tf.Paragraphs[0].Portions.Count > 0)
                            {
                                tf.Paragraphs[0].Portions[0].Text = System.DateTime.Now.ToShortDateString();
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Adds a new slide and note to it.
        /// </summary>
        /// <param name="presentation">The presentation.</param>
        private static Slide AddNewSlideAndNote(Presentation presentation, Stream randomStreamForShapes)
        {
            Slide bodySlide = presentation.AddBodySlide();
            Notes notes = bodySlide.AddNotes();
            notes.Text = "just a note";           

            // add different shapes to notes - it seems this is not possible...
            ShapeCollection notesShapes = notes.Shapes;
            //notesShapes.Add(randomStreamForShapes);
            //notesShapes.AddEllipse(0, 100, 100, 100);

            // add chart?
            //byte[] chartOleData = new byte[randomStreamForShapes.Length];
            //randomStreamForShapes.Position = 0;
            //randomStreamForShapes.Read(chartOleData, 0, chartOleData.Length);
            //notesShapes.AddOleObjectFrame(0, 200, 100, 100, "Random class name", chartOleData);
            //notesShapes.AddRectangle(0, 300, 100, 100);
            //notesShapes.AddTable(0, 400, 100, 100, 5, 5);

            return bodySlide;
        }

        /// <summary>
        /// Adds a background link to a slide.
        /// </summary>
        /// <param name="bodySlide">The slide.</param>
        private static void AddBackgroundLinkToSlide(Slide bodySlide)
        {
            Background background = bodySlide.Background;
            Link bgLink = background.AddLink();
            bgLink.Begin = 10;
            bgLink.End = 20;
            bgLink.SetExternalHyperlink("www.google.ro");
            Notes notes = bodySlide.Notes;
            notes.Text += NewLineStringConstant + "background external link: " + bgLink.ExternalHyperlink;
        }

        private static void AddShapesToSlide(Slide bodySlide, Stream randomStreamForShapes)
        {
            ShapeCollection bodySlideShapes = bodySlide.Shapes;
            //bodySlideShapes.Add(randomStreamForShapes);
            bodySlideShapes.AddEllipse(0, 100, 200, 200);
            byte[] chartOleData = new byte[randomStreamForShapes.Length];
            randomStreamForShapes.Position = 0;
            randomStreamForShapes.Read(chartOleData, 0, chartOleData.Length);
            bodySlideShapes.AddOleObjectFrame(0, 420, 400, 400, "Random class name", chartOleData);
            bodySlideShapes.AddRectangle(0, 830, 200, 200);
            bodySlideShapes.AddTable(0, 1040, 200, 200, 5, 5);
        }
    }
}
