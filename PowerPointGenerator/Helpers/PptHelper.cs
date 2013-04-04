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
        private const string NewLineString = "\r\n";

        /// <summary>
        /// String for component indent.
        /// </summary>
        private const string ComponentExportIndentString = "\t- ";

        public static bool InsertSalesInTemplate(List<Sale> sales, string templatePath, string savedPptPath)
        {
            try
            {
                // write top 20 sales in the first textbox (textholder) of the first slide and rotate it by 90 degrees
                System.IO.FileStream fis = new System.IO.FileStream(templatePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                Presentation pres = new Presentation(templatePath);
                fis.Close();
                SlideCollection slides = pres.Slides;
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
                                text.Text += sales[i].ToString() + NewLineString;
                            }

                            text.RotateTextBy90Degrees = true;
                            text.FitTextToShape();
                        }
                    }
                }

                // write the current date in all "places" which support text
                for (int i = 0; i < slides.Count; i++)
                {
                    // Set text for text frames
                    ShapeCollection shapes = slides[i].Shapes;
                    for (int j = 0; j < shapes.Count; j++)
                    {
                        TextFrame tf = shapes[j].TextFrame;
                        if (tf != null)
                        {
                            tf.Paragraphs[0].Portions[0].Text = System.DateTime.Now.ToShortDateString();
                        }
                    }
                }

                // add a new empty slide and append a note to it
                Slide bodyslide = pres.AddBodySlide();
                Notes notes = bodyslide.AddNotes();
                notes.Text = "just a note";

                Stream randomStreamForShapes = StreamHelper.GenerateRandomStream();

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

                // add a link to a background and append the link to the notes
                Background background = bodyslide.Background;
                Link bgLink = background.AddLink();
                bgLink.Begin = 10;
                bgLink.End = 20;
                bgLink.SetExternalHyperlink("www.google.ro");
                notes.Text += "background external link: " + bgLink.ExternalHyperlink;

                ShapeCollection bodyslideShapes = bodyslide.Shapes;
                //bodyslideShapes.Add(randomStreamForShapes);
                bodyslideShapes.AddEllipse(0, 100, 200, 200);
                byte[] chartOleData = new byte[randomStreamForShapes.Length];
                randomStreamForShapes.Position = 0;
                randomStreamForShapes.Read(chartOleData, 0, chartOleData.Length);
                bodyslideShapes.AddOleObjectFrame(0, 420, 400, 400, "Random class name", chartOleData);
                bodyslideShapes.AddRectangle(0, 830, 200, 200);
                bodyslideShapes.AddTable(0, 1040, 200, 200, 5, 5);

                pres.Save(savedPptPath, Aspose.Slides.Export.SaveFormat.Ppt);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
