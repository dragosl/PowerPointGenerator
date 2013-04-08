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

        /// <summary>
        /// Inserts the sales in the ppt template and other ppt operations.
        /// </summary>
        /// <param name="sales">The sales.</param>
        /// <param name="templatePath">The template path.</param>
        /// <param name="savedPptPath">The path where the ppt is saved.</param>
        /// <returns></returns>
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

                // add a new empty slide and append a note and shapes to it
                Slide bodySlide = AddBodySlide(presentation);

                // adds double slide - no difference from the last added one, as for what is known
                // modify the slide's header,footer,position and other properties; add comments,transitions and tags; save as SVG
                Slide doubleSlide = AddDoubleBodySlide(presentation);

                // add other types of slides
                Slide emptySlide = presentation.AddEmptySlide();
                Slide headerSlide = presentation.AddHeaderSlide();
                Slide titleSlide = presentation.AddTitleSlide();

                // change the ppt file properties
                ModifyPresentationProperties(presentation);                

                presentation.CloneSlide(doubleSlide, presentation.Slides.Count);
                CommentAuthorCollection commentAuthors = presentation.CommentAuthors;
                commentAuthors.AddAuthor("author2");
                commentAuthors[0].ColorIndex = 0;

                presentation.DeleteUnusedMasters();
                presentation.EncryptDocumentProperties = true;
                presentation.FirstSlideNumber = 0;

                // add new font - constructor seems to be ambiguous
                FontCollection fonts = presentation.Fonts;
                FontEntity font = fonts[0];
                fonts.Add(font);

                presentation.GetSlideByPosition(0);
                presentation.GetSlideById(258);
                MainMaster master = presentation.MainMaster;
                //master.ChangeMaster(presentation.GetSlideByPosition(1));
                ExtraColorSchemeCollection colors = master.ExtraColorSchemes;
                Shape shape = master.FindShape("shape1");
                master.FollowMasterBackground = true;
                master.FollowMasterObjects = true;
                master.FollowMasterScheme = true;
                SlideLayout layout = master.Layout;
                master.Name = "master";

                NamedSlideShowCollection namedSlides = presentation.NamedSlideShows;
                namedSlides.Add("slideshow");
                presentation.Password = "asd";
                PictureBulletCollection bullets = presentation.PictureBullets;

                // System.Drawing.* not found in .net 4.5
                //bullets.Add(new PictureBullet(presentation, @"Templates\DownArrow.png"));
                PictureCollection pictures = presentation.Pictures;

                // System.Drawing.* not found in .net 4.5
                //pictures.Add(new Picture(presentation, @"Templates\DownArrow.png"));

                // System.Drawing.* not found in .net 4.5
                //presentation.Print("printer name");
                presentation.RemoveVBAMacros();
                presentation.RemoveWriteProtection();

                // other save options
                presentation.Save(@"D:/demopptPDF.pdf", Aspose.Slides.Export.SaveFormat.Pdf);
                presentation.SetWriteProtection("asd");
                // ... other presentation properties                

                // save file - open and save with user interaction not found yet in the library...
                presentation.Save(savedPptPath, Aspose.Slides.Export.SaveFormat.Ppt);
                return true;
            }
            catch
            {
                return false;
            }
        }

        #region Private methods

        #region First level Private methods

        /// <summary>
        /// Add a new body slide to a presentation and modify some of its properties.
        /// </summary>
        /// <param name="presentation">The presentation.</param>
        /// <returns>The added slide</returns>
        private static Slide AddBodySlide(Presentation presentation)
        {
            Stream randomStreamForShapes = StreamHelper.GenerateRandomStream();

            // add a new empty slide and append a note to it
            Slide bodySlide = AddNewSlideAndNote(presentation, randomStreamForShapes);

            // add a link to a background and append the link to the notes
            AddBackgroundLinkToSlide(bodySlide);

            // add shapes to the newly added slide
            AddShapesToSlide(bodySlide, randomStreamForShapes);

            return bodySlide;
        }

        /// <summary>
        /// Add a new double body slide to a presentation and modify some of its properties..
        /// </summary>
        /// <param name="presentation">The presentation.</param>
        /// /// <returns>The added slide</returns>
        private static Slide AddDoubleBodySlide(Presentation presentation)
        {
            // adds double slide - no difference from the last added one, as for what is known
            Slide doubleSlide = presentation.AddDoubleBodySlide();
            Notes notes = doubleSlide.AddNotes();

            // modify the newly added slide's header and footer and show in notes
            ModifySlideHeaderFooter(doubleSlide, notes);

            // add slide comments, modify other poperties and show in notes
            AddSlideComments(doubleSlide, notes);

            // set the slide position in the presentation - position must be valid
            // save slide in svg format? - opens with IE...
            doubleSlide.SlidePosition = 1;
            doubleSlide.SaveToSVG(@"D:/DoubleSlide.svg");

            // adds slide transitions and tags and show in notes
            AddSlideTransitionAndTags(doubleSlide, notes);

            return doubleSlide;
        }

        /// <summary>
        /// Modifies the presentation properties.
        /// </summary>
        /// <param name="presentation">The presentation.</param>
        private static void ModifyPresentationProperties(Presentation presentation)
        {
            presentation.DocumentProperties.Author = "fav";
            presentation.DocumentProperties.Category = "categ";
            presentation.DocumentProperties.Comments = "comm1";
            presentation.DocumentProperties.Company = "comp1";
            presentation.DocumentProperties.CreatedTime = new DateTime(2001, 1, 1);
            presentation.DocumentProperties.TotalEditingTime = new TimeSpan(0, 0, 100);
            presentation.DocumentProperties.HyperlinkBase = "www.bing.com";
            presentation.DocumentProperties.Keywords = "key1";
            presentation.DocumentProperties.LastSavedBy = "admin";
            presentation.DocumentProperties.LastPrinted = new DateTime(2002, 2, 2);
            presentation.DocumentProperties.LastSavedTime = new DateTime(2003, 3, 3);
            presentation.DocumentProperties.Manager = "maneger1";
            presentation.DocumentProperties.NameOfApplication = "app1";
            presentation.DocumentProperties.RevisionNumber = 123;
            presentation.DocumentProperties.Subject = "subj1";
            presentation.DocumentProperties.Template = "new ppt";
            presentation.DocumentProperties.Title = "Title1";
        }

        #endregion First level Private methods

        #region Second level Private methods

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
            Ellipse ellipse = bodySlideShapes.AddEllipse(0, 100, 200, 200);
            var textframe =  ellipse.AddTextFrame("shape text");
            (ellipse as Shape).AlternativeText = "shape1";
            byte[] chartOleData = new byte[randomStreamForShapes.Length];
            randomStreamForShapes.Position = 0;
            randomStreamForShapes.Read(chartOleData, 0, chartOleData.Length);
            bodySlideShapes.AddOleObjectFrame(0, 420, 400, 400, "Random class name", chartOleData);
            bodySlideShapes.AddRectangle(0, 830, 200, 200);
            bodySlideShapes.AddTable(0, 1040, 200, 200, 5, 5);
        }

        /// <summary>
        /// Modifies the slide header and footer.
        /// </summary>
        /// <param name="doubleSlide">The double slide.</param>
        /// <param name="notes">The notes.</param>
        private static void ModifySlideHeaderFooter(Slide doubleSlide, Notes notes)
        {
            // System.Drawing.Color is not found? in .net 4.5
            //notes.Text = doubleSlide.GetSchemeColor(0).ToString();
            //notes.Text += doubleSlide.GetThumbnail().ToString();

            doubleSlide.HeaderFooter.FooterText = "a";
            doubleSlide.HeaderFooter.HeaderText = "b";
            doubleSlide.HeaderFooter.DateTimeText = "vwsbuei";
            doubleSlide.HeaderFooter.PageNumberVisible = true;
            doubleSlide.HeaderFooter.ShowOnTitleSlide = true;
            notes.Text += "HeaderFooter " + doubleSlide.HeaderFooter.ToString();
            notes.Text += NewLineStringConstant + "FooterText " + doubleSlide.HeaderFooter.FooterText;
            notes.Text += NewLineStringConstant + "HeaderText " + doubleSlide.HeaderFooter.HeaderText;
            notes.Text += NewLineStringConstant + "DateTimeText " + doubleSlide.HeaderFooter.DateTimeText;
            notes.Text += NewLineStringConstant + "IsMasterSlide " + doubleSlide.IsMasterSlide.ToString();
            notes.Text += NewLineStringConstant + "Layout " + doubleSlide.Layout.ToString();

            // master id can be set but this is not done for avoiding crashes/errors.
            //doubleSlide.MasterId = 1;
        }

        /// <summary>
        /// Adds the slide comments and modify other properties.
        /// </summary>
        /// <param name="doubleSlide">The double slide.</param>
        /// <param name="notes">The notes.</param>
        private static void AddSlideComments(Slide doubleSlide, Notes notes)
        {
            doubleSlide.Name = "sadf";
            notes.Text += NewLineStringConstant + "Name " + doubleSlide.Name.ToString();
            doubleSlide.Notes.Text += NewLineStringConstant + "notes accessed using property";
            notes.Text += NewLineStringConstant + "ParentPresentation " + doubleSlide.ParentPresentation.ToString();

            // System.Drawing.Color is not found? in .net 4.5
            //doubleSlide.SetSchemeColor(0, System.Drawing.Color.FromArgb(255, 0, 0));

            // comment adding is not ok because System.Drawing.Point is not found
            CommentCollection comments = doubleSlide.SlideComments;
            //comments.AddComment(null, "JFK", "comment added", DateTime.Now, new System.Drawing.Point());

            notes.Text += NewLineStringConstant + "Comments " + comments.ToString();
            notes.Text += NewLineStringConstant + "SlideId " + doubleSlide.SlideId;
        }

        /// <summary>
        /// Adds slide transition and tags and set slide properties related to this.
        /// </summary>
        /// <param name="doubleSlide">The double slide.</param>
        /// <param name="notes">The notes.</param>
        private static void AddSlideTransitionAndTags(Slide doubleSlide, Notes notes)
        {
            SlideShowTransition slideTransition = doubleSlide.SlideShowTransition;
            slideTransition.EntryEffect = SlideTransitionEffect.BlindsHorizontal;
            slideTransition.AdvanceOnClick = true;
            slideTransition.AdvanceOnTime = false;
            slideTransition.AdvanceTime = 10;
            slideTransition.Hidden = false;
            slideTransition.LoopSoundUntilNext = true;
            notes.Text += NewLineStringConstant + "Sound " + slideTransition.Sound.ToString();
            slideTransition.Speed = SlideTransitionSpeed.Medium;

            TagCollection tags = doubleSlide.Tags;
            notes.Text += NewLineStringConstant + "tags " + tags;
            tags.Add("first tag", "the value of the first tag");
            notes.Text += NewLineStringConstant + "tags after add " + tags;
            notes.Text += NewLineStringConstant + "first tag " + tags["first tag"];
        }

        #endregion Second level Private methods

        #endregion Private methods
    }
}
