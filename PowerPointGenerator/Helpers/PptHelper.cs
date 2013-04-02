using System;
using System.Collections.Generic;
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

        public static void InsertSalesInTemplate(List<Sale> sales, string templatePath, string savedPptPath)
        {

            //System.IO.FileStream fis = new System.IO.FileStream(templatePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            //Presentation pres = new Presentation(templatePath);
            //fis.Close();
            Presentation pres = new Presentation(templatePath);
            Slides slides = pres.Slides;
            if (slides.Count > 0)
            {
                Placeholders places = slides[0].Placeholders;
                if (places.Count > 0)
                {
                    TextHolder text = places[0] as TextHolder;
                    if (text != null)
                    {
                        
                        text.Text = string.Empty;
                        foreach (Sale sale in sales)
                        {
                            text.Text += sale.ToString() + NewLineString;
                        }
                        text.RotateTextBy90Degrees = true;
                    }
                }
            }

            for (int i = 0; i < slides.Count; i++)
            {
                // Set text for text frames
                Shapes shapes = slides[i].Shapes;
                for (int j = 0; j < shapes.Count; j++)
                {
                    TextFrame tf = shapes[j].TextFrame;
                    if (tf != null)
                    {
                            tf.Paragraphs[0].Portions[0].Text = System.DateTime.Now.ToShortDateString();
                    }
                }
            }

            pres.Save(savedPptPath, Aspose.Slides.Export.SaveFormat.Ppt);
        }
    }
}
