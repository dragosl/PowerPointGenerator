using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
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
        #pragma warning disable 169
        private const string ComponentExportIndentStringConstant = "\t- ";
        #pragma warning restore 169

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
                FileStream file = new FileStream(templatePath, FileMode.Open, FileAccess.Read);
                Presentation presentation = new Presentation(templatePath);
                file.Close();
                SlideCollection slides = presentation.Slides;

                // write top 20 sales in the first textbox (textholder) of the first slide and rotate it by 90 degrees
                // this is used with 20 only because 700 are too much in a textbox (20 are also too much, but still better)
                WriteTopSales(sales, slides);

                // write the current date in all "places" which support text
                WriteDateEverywhere(slides);

                // add a new empty slide and append a note and shapes to it
                AddBodySlide(presentation);

                // adds double slide - no difference from the last added one, as for what is known
                // modify the slide's header,footer,position and other properties; add comments,transitions and tags; save as SVG
                AddDoubleBodySlide(presentation);

                // add other types of slides
                presentation.AddEmptySlide();
                presentation.AddHeaderSlide();
                presentation.AddTitleSlide();

                // change the ppt file properties
                ModifyPresentationProperties(presentation);

                // other save options - most used (pptx not supported yet, as for what they say)
                // also supports Pps, Xps, Ppsx, Tiff, Odp, Pptm, Ppsm, Potx, Potm, PdfNotes, Html, TiffNotes but not exemplified here
                presentation.Save(@"D:/demopptPDF.pdf", Aspose.Slides.Export.SaveFormat.Pdf);
                //presentation.Save(@"D:/demopptx.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
                presentation.SetWriteProtection("asd");

                // this is like save but with less options - no other difference as for what is known so far
                presentation.Write(savedPptPath);

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
        private static void AddBodySlide(Presentation presentation)
        {
            Stream randomStreamForShapes = StreamHelper.GenerateRandomStream();

            // add a new empty slide and append a note to it
            Slide bodySlide = AddNewSlideAndNote(presentation, randomStreamForShapes);

            // add a link to a background and append the link to the notes
            AddBackgroundLinkToSlide(bodySlide);

            // add shapes to the newly added slide
            AddShapesToSlide(bodySlide, randomStreamForShapes);
        }

        /// <summary>
        /// Add a new double body slide to a presentation and modify some of its properties..
        /// </summary>
        /// <param name="presentation">The presentation.</param>
        /// /// <returns>The added slide</returns>
        private static void AddDoubleBodySlide(Presentation presentation)
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
        }

        /// <summary>
        /// Modifies the presentation properties.
        /// </summary>
        /// <param name="presentation">The presentation.</param>
        private static void ModifyPresentationProperties(Presentation presentation)
        {
            ModifyPresentationDocumentProperties(presentation);
            ModifyPresentationCommentAuthors(presentation);        

            presentation.DeleteUnusedMasters();
            presentation.EncryptDocumentProperties = true;
            presentation.FirstSlideNumber = 0;

            ModifyPresentationFonts(presentation);
            ModifyPresentationMaster(presentation);

            NamedSlideShowCollection namedSlides = presentation.NamedSlideShows;
            namedSlides.Add("slideshow");
            presentation.Password = "asd";

            ModifyPresentationPictures(presentation);
            
            // here a valid printer name should be set in order to work
            //presentation.Print("printer name");
            presentation.RemoveVBAMacros();
            presentation.RemoveWriteProtection();

            ModifyPresentationSettings(presentation);

            presentation.SlideSize = new Size(1024, 768);
            presentation.SlideSizeType = SlideSizeType.A4Paper;
            presentation.SlideViewType = SlideViewType.SlideShowFullScreen;
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
                        Paragraph firstParagraph = text.Paragraphs[0];
                        firstParagraph.Portions.Add(new Portion("Top 20 sales"));

                        PortionCollection portions = ((TextHolder) places[1]).Paragraphs[0].Portions;

                        //foreach (Sale sale in sales)
                        for (int i = 0; i < 20; i++)
                        {
                            portions.Add(new Portion(sales[i] + NewLineStringConstant));
                            portions[i].FontColor = Color.Green;
                            portions[i].FontHeight = 5;                            
                        }

                        portions[portions.Count - 1].FontHeight = 5;
                        portions[portions.Count - 1].FontColor = Color.Green;

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
                                tf.Paragraphs[0].Portions.Add(new Portion(DateTime.Now.ToShortDateString()));
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
        #pragma warning disable 1573
        // ReSharper disable UnusedParameter.Local
        private static Slide AddNewSlideAndNote(Presentation presentation, Stream randomStreamForShapes)
        // ReSharper restore UnusedParameter.Local
        #pragma warning restore 1573
        {
            Slide bodySlide = presentation.AddBodySlide();
            Notes notes = bodySlide.AddNotes();
            notes.Text = "just a note";           

            // add different shapes to notes - it seems this is not possible...
            #pragma warning disable 168
            ShapeCollection notesShapes = notes.Shapes;
            #pragma warning restore 168
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

        /// <summary>
        /// Adds shapes to a slide.
        /// </summary>
        /// <param name="bodySlide">The slide.</param>
        /// <param name="randomStreamForShapes">A random stream for shapes.</param>
        private static void AddShapesToSlide(Slide bodySlide, Stream randomStreamForShapes)
        {
            ShapeCollection bodySlideShapes = bodySlide.Shapes;
            //bodySlideShapes.Add(randomStreamForShapes);
            Ellipse ellipse = bodySlideShapes.AddEllipse(5500, 2500, 200, 200);
            ModifyEllipseAnimationSettings(ellipse);
            ModifyEllipseFillFormat(ellipse);
            ModifyEllipseLineFormat(ellipse);
            ModifyEllipseShadowFormat(ellipse);
            ModifyEllipseSingleLineProperties(ellipse);

            // other options which repeat from ellipse were ommited
            byte[] chartOleData = new byte[randomStreamForShapes.Length];
            randomStreamForShapes.Position = 0;
            randomStreamForShapes.Read(chartOleData, 0, chartOleData.Length);
            OleObjectFrame frame = bodySlideShapes.AddOleObjectFrame(0, 720, 400, 400, "Random class name", chartOleData);
            ModifyOleObjectFrameProperties(frame);

            bodySlideShapes.AddRectangle(0, 1330, 200, 200);

            Table table = bodySlideShapes.AddTable(0, 1740, 200, 200, 5, 5);
            ModifyTableProperties(table);
        }

        /// <summary>
        /// Modifies the slide header and footer.
        /// </summary>
        /// <param name="doubleSlide">The double slide.</param>
        /// <param name="notes">The notes.</param>
        private static void ModifySlideHeaderFooter(Slide doubleSlide, Notes notes)
        {
            notes.Text = "GetSchemeColor " + doubleSlide.GetSchemeColor(0).ToString();
            notes.Text += NewLineStringConstant + "GetThumbnail " + doubleSlide.GetThumbnail();

            doubleSlide.HeaderFooter.FooterText = "a";
            doubleSlide.HeaderFooter.HeaderText = "b";
            doubleSlide.HeaderFooter.DateTimeText = "vwsbuei";
            doubleSlide.HeaderFooter.PageNumberVisible = true;
            doubleSlide.HeaderFooter.ShowOnTitleSlide = true;
            notes.Text += NewLineStringConstant + "HeaderFooter " + doubleSlide.HeaderFooter;
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
            notes.Text += NewLineStringConstant + "Name " + doubleSlide.Name;
            doubleSlide.Notes.Text += NewLineStringConstant + "notes accessed using property";
            notes.Text += NewLineStringConstant + "ParentPresentation " + doubleSlide.ParentPresentation;

            doubleSlide.SetSchemeColor(0, Color.FromArgb(1, Color.LightGray));
            CommentCollection comments = doubleSlide.SlideComments;

            // obtaining a CommentAuthor object, used in the add method, is unknown
            //comments.AddComment(null, "JFK", "comment added", DateTime.Now, new System.Drawing.Point());

            notes.Text += NewLineStringConstant + "Comments " + comments;
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

        /// <summary>
        /// Modifies the presentation document properties.
        /// </summary>
        /// <param name="presentation">The presentation.</param>
        private static void ModifyPresentationDocumentProperties(Presentation presentation)
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

        /// <summary>
        /// Modifies the presentation's comment authors.
        /// </summary>
        /// <param name="presentation">The presentation.</param>
        private static void ModifyPresentationCommentAuthors(Presentation presentation)
        {
            presentation.CloneSlide(presentation.GetSlideByPosition(2), presentation.Slides.Count);
            CommentAuthorCollection commentAuthors = presentation.CommentAuthors;
            commentAuthors.AddAuthor("author2");
            commentAuthors[0].ColorIndex = 0;
        }

        /// <summary>
        /// Modifies the presentation's fonts.
        /// </summary>
        /// <param name="presentation">The presentation.</param>
        private static void ModifyPresentationFonts(Presentation presentation)
        {
            // add new font - constructor seems to be ambiguous
            FontCollection fonts = presentation.Fonts;
            FontEntity font = fonts[0];
            fonts.Add(font);
        }

        /// <summary>
        /// Modifies the presentation master.
        /// </summary>
        /// <param name="presentation">The presentation.</param>
        private static void ModifyPresentationMaster(Presentation presentation)
        {
            presentation.GetSlideByPosition(0);
            presentation.GetSlideById(258);
            MainMaster master = presentation.MainMaster;
            //master.ChangeMaster(presentation.GetSlideByPosition(1));
            #pragma warning disable 168
            ExtraColorSchemeCollection colors = master.ExtraColorSchemes;
            Shape shape = master.FindShape("shape1");
            master.FollowMasterBackground = true;
            master.FollowMasterObjects = true;
            master.FollowMasterScheme = true;
            SlideLayout layout = master.Layout;
            master.Name = "master";
            #pragma warning restore 168
        }

        /// <summary>
        /// Modifies the presentation's pictures.
        /// </summary>
        /// <param name="presentation">The presentation.</param>
        private static void ModifyPresentationPictures(Presentation presentation)
        {
            PictureBulletCollection bullets = presentation.PictureBullets;
            bullets.Add(new PictureBullet(presentation, @"Templates\DownArrow.png"));
            PictureCollection pictures = presentation.Pictures;
            pictures.Add(new Picture(presentation, @"Templates\DownArrow.png"));
        }

        /// <summary>
        /// Modifies the presentation's settings.
        /// </summary>
        /// <param name="presentation">The presentation.</param>
        private static void ModifyPresentationSettings(Presentation presentation)
        {
            SlideShowSettings settings = presentation.SlideShowSettings;
            settings.EndingSlide = 2;
            settings.LoopUntilStopped = true;
            settings.ManualAdvance = true;
            settings.RangeType = SlideShowRangeType.ShowAll;
            settings.ShowScrollbar = true;
            settings.ShowType = SlideShowType.ShowTypeKiosk;
            settings.ShowWithAnimation = true;
            settings.ShowWithNarration = true;
            settings.StartingSlide = 0;
        }

        /// <summary>
        /// Modifies the ellipse's animation settings.
        /// </summary>
        /// <param name="ellipse">The ellipse.</param>
        private static void ModifyEllipseAnimationSettings(Ellipse ellipse)
        {
            AnimationSettings animationSettings = ellipse.AnimationSettings;
            animationSettings.AdvanceMode = new ShapeAdvanceMode();
            animationSettings.AdvanceTime = 10;
            animationSettings.AfterEffect = ShapeAfterEffect.Dim;
            animationSettings.AnimateBackground = true;
            animationSettings.AnimateTextInReverse = true;
            animationSettings.AnimationOrder = 0;
            animationSettings.AnimationSlideCount = 0;
            animationSettings.EntryEffect = ShapeEntryEffect.StretchRight;
            animationSettings.TextLevelEffect = TextLevelEffect.AnimateByAllLevels;
            animationSettings.TextUnitEffect = TextUnitEffect.AnimateByCharacter;
        }

        /// <summary>
        /// Modifies the ellipse's fill format.
        /// </summary>
        /// <param name="ellipse">The ellipse.</param>
        private static void ModifyEllipseFillFormat(Ellipse ellipse)
        {
            ellipse.ClearLink();
            FillFormat fillFormat = ellipse.FillFormat;
            fillFormat.BackColor = Color.LightCyan;
            fillFormat.ForeColor = Color.DarkBlue;
            fillFormat.GradientColorType = GradientColorType.TwoColors;
            fillFormat.GradientDegree = 10;
            fillFormat.GradientFillAngle = 30;
            fillFormat.GradientFillFocus = 10;
            fillFormat.GradientPreset = GradientPreset.CalmWater;

            // color blend configuration not known and default raises exception
            //fillFormat.GradientStops = new System.Drawing.Drawing2D.ColorBlend();

            fillFormat.GradientStyle = GradientStyle.FromCorner1;
            fillFormat.PatternStyle = PatternStyle.DarkDownwardDiagonal;
            fillFormat.RotateWithShape = true;
            fillFormat.Type = FillType.Gradient;
        }

        /// <summary>
        /// Modifies the ellipse's line format.
        /// </summary>
        /// <param name="ellipse">The ellipse.</param>
        private static void ModifyEllipseLineFormat(Ellipse ellipse)
        {
            LineFormat format = ellipse.LineFormat;
            format.BeginArrowheadLength = LineArrowheadLength.Medium;
            format.BeginArrowheadStyle = LineArrowheadStyle.Diamond;
            format.BeginArrowheadWidth = LineArrowheadWidth.Medium;
            format.DashStyle = LineDashStyle.Dash;
            format.EndArrowheadLength = LineArrowheadLength.Medium;
            format.EndArrowheadStyle = LineArrowheadStyle.Triangle;
            format.EndArrowheadWidth = LineArrowheadWidth.Medium;
            format.ForeColor = Color.AliceBlue;
            format.JoinStyle = LineJoinStyle.JoinRound;
            format.RoundEndCap = true;
            format.ShowLines = true;
            format.Style = LineStyle.ThinThin;
            format.Width = 300;
        }

        /// <summary>
        /// Modifies the ellipse's shadow format.
        /// </summary>
        /// <param name="ellipse">The ellipse.</param>
        private static void ModifyEllipseShadowFormat(Ellipse ellipse)
        {
            ShadowFormat shadow = ellipse.ShadowFormat;
            shadow.LightColor = Color.Chartreuse;
            shadow.LightColorIndex = 0;
            shadow.PerspectiveXNumerator = 1;
            shadow.PerspectiveYNumerator = 1;
            shadow.SecondShadowOffsetX = 1;
            shadow.SecondShadowOffsetY = 1;
            shadow.ShadowColor = Color.Cornsilk;
            shadow.ShadowColorIndex = 1;
            shadow.ShadowOffsetX = 1;
            shadow.ShadowOffsetY = 1;
            shadow.ShadowOriginX = 1;
            shadow.ShadowOriginY = 1;
            shadow.ShadowStyle = ShadowStyle.Style14;
            shadow.ShadowTransformM11 = float.Epsilon;
            shadow.ShadowTransformM12 = float.Epsilon;
            shadow.ShadowTransformM21 = float.Epsilon;
            shadow.ShadowTransformM22 = float.Epsilon;
            shadow.Type = ShadowType.Double;
            shadow.Visible = true;
        }

        /// <summary>
        /// Modifies the ellipse's properties which can be set by a single line.
        /// </summary>
        /// <param name="ellipse">The ellipse.</param>
        private static void ModifyEllipseSingleLineProperties(Ellipse ellipse)
        {
            #pragma warning disable 168
            var textframe = ellipse.AddTextFrame("shape text");
            ellipse.AlternativeText = "shape1";
            ellipse.FlipHorizontal = true;
            ellipse.FlipVertical = true;
            ellipse.Height = 300;
            ellipse.Hidden = false;
            ellipse.Protection = new ShapeProtection();
            ellipse.Rotation = 30;
            System.Drawing.Rectangle rectangle = ellipse.ShapeRectangle;
            ellipse.Width = 300;
            ellipse.ZOrder(ZOrderCmd.BringForward);
            #pragma warning restore 168
        }

        /// <summary>
        /// Modifies the OLE object frame's properties.
        /// </summary>
        /// <param name="frame">The frame.</param>
        private static void ModifyOleObjectFrameProperties(OleObjectFrame frame)
        {
            frame.Brightness = 10;
            frame.ColorType = PictureColorType.Grayscale;
            frame.Contrast = 10;
            frame.CropBottom = 20;
            frame.CropLeft = 20;
            frame.CropRight = 20;
            frame.CropTop = 20;
            frame.FlipHorizontal = true;
            frame.FlipVertical = true;
            frame.FollowColorScheme = OleFollowColorScheme.TextAndBackground;
            frame.IsObjectIcon = true;
            frame.PictureFileName = @"Templates\DownArrow.png";
            frame.TransparentColor = Color.DarkCyan;
        }

        /// <summary>
        /// Modifies the table's properties.
        /// </summary>
        /// <param name="table">The table.</param>
        private static void ModifyTableProperties(Table table)
        {
            table.AddColumn();
            table.AddRow();
            table.AlternativeText = "table alt text";
            table.DeleteColumn(table.ColumnsNumber - 1);
            table.DeleteRow(table.RowsNumber - 1);
            Cell cell = table.GetCell(1, 1);
            TextFrame text = cell.TextFrame;
            text.Text = "cell";
            #pragma warning disable 168
            CellBorder border = cell.BorderBottom;
            #pragma warning disable 219
            Point point = cell.BottomRightCell;
            #pragma warning restore 219
            // ReSharper disable RedundantAssignment
            point = cell.TopLeftCell;
            // ReSharper restore RedundantAssignment
            table.MergeCells(table.GetCell(1, 1), table.GetCell(1, 2));
            table.SetBorders(5, Color.DarkGoldenrod);
            table.SetColumnWidth(1, 20);
            table.SetRowHeight(3, 100);
            #pragma warning restore 168
        }

        #endregion Second level Private methods

        #endregion Private methods
    }
}
