using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;
using Rg = Rhino.Geometry;

using PP = ShapeCrawler;
using XF = DocumentFormat.OpenXml;
using XP = DocumentFormat.OpenXml.Presentation;
using XD = DocumentFormat.OpenXml.Drawing;

using System.IO;

namespace PptPlus
{
    public class Content : PpBase
    {

        #region members

        public enum ContentTypes { None, Text, List, Table, Image, Shape, Line, ChartBar, ChartBarStack, ChartPie, ChartScatter, Placeholder };

        public enum ShapeLines { Line, LineInverse, StraightConnector1, BentConnector2, BentConnector3, BentConnector4, BentConnector5, CurvedConnector2, CurvedConnector3, CurvedConnector4, CurvedConnector5 };
        public enum ShapeRectangles { RoundedRectangle, SingleCornerRoundedRectangle, TopCornersRoundedRectangle, DiagonalCornersRoundedRectangle, SnipRoundRectangle, Snip1Rectangle, Snip2SameRectangle, Snip2DiagonalRectangle, Plaque, Ellipse };
        public enum ShapePolygons { Triangle, RightTriangle, Rectangle, Diamond, Parallelogram, Trapezoid, NonIsoscelesTrapezoid, Pentagon, Hexagon, Heptagon, Octagon, Decagon, Dodecagon };
        public enum ShapeStars { Star4, Star5, Star6, Star7, Star8, Star10, Star12, Star16, Star24, Star32 };
        public enum ShapeArrows { RightArrow, LeftArrow, UpArrow, DownArrow, StripedRightArrow, NotchedRightArrow, BentUpArrow, LeftRightArrow, UpDownArrow, LeftUpArrow, LeftRightUpArrow, QuadArrow, LeftArrowCallout, RightArrowCallout, UpArrowCallout, DownArrowCallout, LeftRightArrowCallout, UpDownArrowCallout, QuadArrowCallout, BentArrow, UTurnArrow, CircularArrow, LeftCircularArrow, LeftRightCircularArrow, CurvedRightArrow, CurvedLeftArrow, CurvedUpArrow, CurvedDownArrow, SwooshArrow };
        public enum ShapeEquations { MathPlus, MathMinus, MathMultiply, MathDivide, MathEqual, MathNotEqual };
        public enum ShapeFlowchart { FlowChartProcess, FlowChartDecision, FlowChartInputOutput, FlowChartPredefinedProcess, FlowChartInternalStorage, FlowChartDocument, FlowChartMultidocument, FlowChartTerminator, FlowChartPreparation, FlowChartManualInput, FlowChartManualOperation, FlowChartConnector, FlowChartPunchedCard, FlowChartPunchedTape, FlowChartSummingJunction, FlowChartOr, FlowChartCollate, FlowChartSort, FlowChartExtract, FlowChartMerge, FlowChartOfflineStorage, FlowChartOnlineStorage, FlowChartMagneticTape, FlowChartMagneticDisk, FlowChartMagneticDrum, FlowChartDisplay, FlowChartDelay, FlowChartAlternateProcess, FlowChartOffpageConnector };
        public enum ShapeBanners { Ribbon, Ribbon2, EllipseRibbon, EllipseRibbon2, LeftRightRibbon, VerticalScroll, HorizontalScroll, Wave, DoubleWave, Plus };
        public enum ShapeCallouts { Callout1, Callout2, Callout3, AccentCallout1, AccentCallout2, AccentCallout3, BorderCallout1, BorderCallout2, BorderCallout3, AccentBorderCallout1, AccentBorderCallout2, AccentBorderCallout3, WedgeRectangleCallout, WedgeRoundRectangleCallout, WedgeEllipseCallout, CloudCallout, Cloud };
        public enum ShapeActions { ActionButtonBlank, ActionButtonHome, ActionButtonHelp, ActionButtonInformation, ActionButtonForwardNext, ActionButtonBackPrevious, ActionButtonEnd, ActionButtonBeginning, ActionButtonReturn, ActionButtonDocument, ActionButtonSound, ActionButtonMovie };
        public enum ShapeBasics { Teardrop, HomePlate, Chevron, PieWedge, Pie, BlockArc, Donut, NoSmoking, Cube, Can, LightningBolt, Heart, Sun, Moon, SmileyFace, IrregularSeal1, IrregularSeal2, FoldedCorner, Bevel, Frame, HalfFrame, Corner, DiagonalStripe, Chord, Arc, LeftBracket, RightBracket, LeftBrace, RightBrace, BracketPair, BracePair, Gear6, Gear9, Funnel, CornerTabs, SquareTabs, PlaqueTabs, ChartX, ChartStar, ChartPlus };

        public enum TableStyles { Blank, Grid, Dark1, Dark2, Medium1, Medium2, Medium3, Light1, Light2, Light3, Light4, Themed };

        protected ContentTypes ContentType = ContentTypes.None;
        protected PP.ITableStyle tableStyle = PP.Tables.CommonTableStyles.NoStyleTableGrid;
        protected PP.Geometry geometry = PP.Geometry.Rectangle;

        protected Rg.Rectangle3d boundary = new Rg.Rectangle3d(Rg.Plane.WorldXY, 100, 100);

        protected Paragraph text = new Paragraph();
        protected Sd.Bitmap image = new Sd.Bitmap(10, 10);

        string seriesName = string.Empty;
        Dictionary<string, double> categories = new Dictionary<string, double>();
        List<string> seriesNames = new List<string>();
        Dictionary<string, List<double>> categoryList = new Dictionary<string, List<double>>();
        List<List<Paragraph>> values = new List<List<Paragraph>>();

        Rg.Line line = new Rg.Line();

        protected int index = -1;
        protected PP.PlaceholderType placeholder = PP.PlaceholderType.Content;

        #endregion

        #region constructors

        public Content() : base()
        {

        }

        public Content(Content content) : base(content)
        {
            this.ContentType = content.ContentType;
            this.Boundary = content.Boundary.Duplicate();

            this.geometry = content.geometry;
            this.line = new Rg.Line(content.line.From, content.line.To);

            this.tableStyle = content.tableStyle;

            this.text = new Paragraph(content.text);
            this.image = new Sd.Bitmap(content.image);

            this.seriesName = content.seriesName;
            this.categories = content.categories.Duplicate();
            this.seriesNames = content.seriesNames.Duplicate();
            this.categoryList = content.categoryList.Duplicate();
            this.values = content.values.Duplicate();

            this.index = content.index;
            this.placeholder = content.placeholder;
        }

        public Content(PP.IShape shape, int index, double height)
        {
            this.index = index;
            this.ContentType = ContentTypes.Placeholder;
            this.placeholder = shape.PlaceholderType.GetValueOrDefault();

            Rg.Plane plane = Rg.Plane.WorldXY;
            plane.Origin = new Rg.Point3d((double)shape.X, height - (double)shape.Y - (double)shape.Height, 0);
            this.boundary = new Rg.Rectangle3d(plane, (double)shape.Width, (double)shape.Height);
            this.Graphic = shape.GetGraphic();
            this.Font = shape.GetFont();
        }

        #region -Text
        public static Content CreateTextContent(string text, Content content = null)
        {
            return Content.CreateTextContent(new Paragraph(text), content);
        }
        public static Content CreateTextContent(Fragment fragment, Content content = null)
        {
            return Content.CreateTextContent(new Paragraph(fragment), content);
        }
        public static Content CreateTextContent(Paragraph paragraph, Content content = null)
        {
            if (content == null) content = new Content();
            content.ContentType = ContentTypes.Text;

            content.text = new Paragraph(paragraph);

            content.graphic = new Graphic(Graphics.Empty);
            return content;
        }
        #endregion

        #region -List
        public static Content CreateListContent(List<Paragraph> paragraphs, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.List;

            content.values.Add(new List<Paragraph>());
            foreach (Paragraph paragraph in paragraphs) content.values[0].Add(new Paragraph(paragraph));

            content.graphic = new Graphic(Graphics.Empty);
            return content;
        }
        #endregion

        #region -Tables
        public static Content CreateTableContent(List<List<Paragraph>> paragraphs, TableStyles style, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.Table;

            content.tableStyle = style.ToStyle();

            content.values = paragraphs.Duplicate();

            content.graphic = new Graphic(Graphics.Empty);
            return content;
        }

        #endregion

        #region -Images
        public static Content CreateImageContent(Sd.Bitmap image, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.Image;

            content.image = new Sd.Bitmap(image);

            content.graphic = new Graphic(Graphics.Empty);
            return content;
        }
        #endregion

        #region -Lines
        public static Content CreateLineContent(Rg.Line line, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.Line;

            content.line = new Rg.Line(line.From, line.To);
            Rg.BoundingBox bbox = content.line.BoundingBox;
            content.boundary = new Rg.Rectangle3d(Rg.Plane.WorldXY, bbox.Corner(true, true, true), bbox.Corner(false, false, true));

            content.graphic = Graphics.Outline;
            return content;
        }
        #endregion

        #region -Shapes
        public static Content CreateShapeContent(ShapeLines shape, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.Shape;

            content.geometry = shape.ToGeometry();
            content.graphic = Graphics.Solid;
            return content;
        }
        public static Content CreateShapeContent(ShapeRectangles shape, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.Shape;

            content.geometry = shape.ToGeometry();
            content.graphic = Graphics.Solid;
            return content;
        }
        public static Content CreateShapeContent(ShapePolygons shape, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.Shape;

            content.geometry = shape.ToGeometry();
            content.graphic = Graphics.Solid;
            return content;
        }
        public static Content CreateShapeContent(ShapeStars shape, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.Shape;

            content.geometry = shape.ToGeometry();
            content.graphic = Graphics.Solid;
            return content;
        }
        public static Content CreateShapeContent(ShapeArrows shape, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.Shape;

            content.geometry = shape.ToGeometry();
            content.graphic = Graphics.Solid;
            return content;
        }
        public static Content CreateShapeContent(ShapeEquations shape, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.Shape;

            content.geometry = shape.ToGeometry();
            content.graphic = Graphics.Solid;
            return content;
        }
        public static Content CreateShapeContent(ShapeFlowchart shape, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.Shape;

            content.geometry = shape.ToGeometry();
            content.graphic = Graphics.Solid;
            return content;
        }
        public static Content CreateShapeContent(ShapeBanners shape, Rg.Rectangle3d boundary, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.Shape;

            content.geometry = shape.ToGeometry();
            content.graphic = Graphics.Solid;
            content.Boundary = boundary;
            return content;
        }
        public static Content CreateShapeContent(ShapeCallouts shape, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.Shape;

            content.geometry = shape.ToGeometry();
            content.graphic = Graphics.Solid;
            return content;
        }
        public static Content CreateShapeContent(ShapeActions shape, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.Shape;

            content.geometry = shape.ToGeometry();
            content.graphic = Graphics.Solid;
            return content;
        }
        public static Content CreateShapeContent(ShapeBasics shape, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.Shape;

            content.geometry = shape.ToGeometry();
            content.graphic = Graphics.Solid;
            return content;
        }
        #endregion

        #region -Charts
        public static Content CreateChartBarContent(List<double> values, List<string> labels, string seriesName, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.ChartBar;

            if (labels.Count == 0) labels.Add("Item 0");

            for (int i = labels.Count; i < values.Count; i++) labels.Add("Item " + i);
            for (int i = 0; i < values.Count; i++) content.categories.Add(labels[i], values[i]);

            content.seriesName = seriesName;

            return content;
        }
        public static Content CreateChartPieContent(List<double> values, List<string> labels, string seriesName, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.ChartPie;

            if (labels.Count == 0) labels.Add("Item 0");

            for (int i = labels.Count; i < values.Count; i++) labels.Add("Item " + i);
            for (int i = 0; i < values.Count; i++) content.categories.Add(labels[i], values[i]);

            content.seriesName = seriesName;

            return content;
        }
        public static Content CreateChartStackedBarContent(List<List<double>> values, List<string> labels, List<string> seriesNames, Content content = null)
        {
            if (content == null) content = new Content();
            content = new Content(content);
            content.ContentType = ContentTypes.ChartBarStack;

            if (labels.Count == 0) labels.Add("Item 0");
            if (seriesNames.Count == 0) seriesNames.Add("Series 0");

            for (int i = labels.Count; i < values.Count; i++) labels.Add("Item " + i);
            for (int i = seriesNames.Count; i < values.Count; i++) seriesNames.Add("Series " + i);
            for (int i = 0; i < values.Count; i++) content.categoryList.Add(labels[i], values[i]);

            content.seriesNames = seriesNames.Duplicate();

            return content;
        }
        #endregion

        #endregion

        #region properties

        public virtual Rg.Rectangle3d Boundary
        {
            get { return this.boundary; }
            set { this.boundary = new Rg.Rectangle3d(value.Plane, value.X, value.Y); }
        }

        public virtual List<string> CategoryLabels
        {
            get { return this.categories.Keys.ToList(); }
        }

        public virtual List<double> CategoryValues
        {
            get { return this.categories.Values.ToList(); }
        }

        public virtual List<string> CategoryListLabels
        {
            get { return this.categoryList.Keys.ToList(); }
        }

        public virtual List<List<double>> CategoryListValues
        {
            get { return this.categoryList.ToListOfLists(); }
        }

        public virtual string SeriesName
        {
            get { return this.seriesName; }
        }

        public virtual Paragraph Text
        {
            get { return this.text; }
            set { this.text = new Paragraph(value); }
        }

        public virtual List<List<Paragraph>> Values
        {
            get { return this.values.Duplicate(); }
        }

        public virtual Rg.Line Line
        {
            get { return new Rg.Line(this.line.From, this.line.To); }
        }

        public virtual Sd.Bitmap Image
        {
            get
            {
                if (this.image != null) return new Sd.Bitmap(this.image);
                return null;
            }
        }

        public virtual PP.Geometry Geometry
        {
            get { return this.geometry; }
        }

        #endregion

        #region methods

        public void Render(PP.ISlide iSlide, PpSlide slide)
        {

            int height = (int)slide.Layout.Page.Boundary.GetHeight();
            switch (this.ContentType)
            {
                case ContentTypes.Text:
                    this.RenderTextBox(iSlide);
                    break;
                case ContentTypes.List:
                    this.RenderListBox(iSlide);
                    break;
                case ContentTypes.Table:
                    this.RenderTable(iSlide);
                    break;
                case ContentTypes.Image:
                    this.RenderImage(iSlide);
                    break;
                case ContentTypes.Shape:
                    this.RenderShape(iSlide);
                    break;
                case ContentTypes.Line:
                    this.RenderLine(iSlide, height);
                    break;
                case ContentTypes.ChartBar:
                    this.RenderBarChart(iSlide);
                    break;
                case ContentTypes.ChartPie:
                    this.RenderPieChart(iSlide);
                    break;
                case ContentTypes.ChartBarStack:
                    this.RenderBarStackChart(iSlide);
                    break;
            }
            if (iSlide.Shapes.Count > 0 && this.index < 0)
            {
                PP.IShape shape = iSlide.Shapes.Last();
                shape.SetSize(this.Boundary, slide.Layout.Page.Boundary);

                //Rg.Vector3d.VectorAngle(this.boundary.Plane.YAxis, Rg.Vector3d.YAxis,Rg.Plane.WorldXY) / Math.PI * 180.0;
            }
        }

        public void RenderTextBox(PP.ISlide slide)
        {
            if (this.index >= 0)
            {
                PP.IShape shape = slide.Shapes[this.index];
                this.AddText(shape.TextBox.Paragraphs.Last(), this.text);
            }
            else
            {
                PP.IShape shape = this.AddNewShape(slide);
                shape.TextBox.Paragraphs.Last();
                this.AddText(shape.TextBox.Paragraphs.Last(), this.text);
                shape.TextBox.VerticalAlignment = this.text.Font.VerticalAlignment.ToPpt();

                shape.SetGraphics(this.graphic);
            }
        }

        public void RenderListBox(PP.ISlide slide)
        {
            PP.IShape shape = this.AddNewShape(slide);

            foreach (Paragraph paragraph in this.values[0])
            {
                shape.TextBox.Paragraphs.Add();
                PP.IParagraph list = shape.TextBox.Paragraphs.Last();
                PP.Bullet bullet = list.Bullet;

                switch (paragraph.BulletPoint)
                {
                    case Paragraph.BulletPoints.None:
                        break;
                    case Paragraph.BulletPoints.Number:
                        bullet.Type = PP.BulletType.Numbered;
                        break;
                    case Paragraph.BulletPoints.Character:
                        bullet.Type = PP.BulletType.Character;
                        bullet.Character = paragraph.BulletCharacter;
                        break;
                    default:
                        bullet.Type = PP.BulletType.Character;
                        bullet.Character = paragraph.BulletPoint.GetUnicode();
                        break;
                }

                bullet.Size = 100;
                list.Portions.AddText("\t ");
                this.AddText(list, paragraph);
            }
            shape.SetGraphics(this.graphic);
        }

        protected PP.IParagraph AddText(PP.IParagraph iParagraph, Paragraph paragraph)
        {

            Font font = this.Font;
            if (paragraph.HasFont) font = paragraph.Font;
            if (font.HorizontalAlignment != Font.HorizontalAlignments.Default) iParagraph.HorizontalAlignment = font.HorizontalAlignment.ToPpt();
            iParagraph.IndentLevel = paragraph.IndentationLevel;
            foreach (Fragment fragment in paragraph.Fragments)
            {
                iParagraph.Portions.AddText(fragment.Text);
                PP.IParagraphPortion portion = iParagraph.Portions.Last();
                this.RenderText(portion, fragment);
                if (fragment.HasLink) portion.Link.AddFile(fragment.Hyperlink);
            }

            return iParagraph;
        }

        protected void RenderText(PP.IParagraphPortion portion, Fragment fragment)
        {

            Font font = this.Font;
            if (fragment.HasFont) font = fragment.Font;
            if (font.HasFamily) portion.Font.LatinName = font.Family;
            if (font.HasSize) portion.Font.Size = ((decimal)font.Size);
            if (font.HasColor) portion.Font.Color.Set(font.Color.ToHex());
            //if (font.HasHighlight) portion.TextHighlightColor = PP.Color.Black;

            portion.Font.IsBold = font.Bold;
            portion.Font.IsItalic = font.Italic;
            if (font.Underlined) portion.Font.Underline = XD.TextUnderlineValues.Single;

        }

        public void RenderImage(PP.ISlide slide)
        {
            MemoryStream memoryStream = new MemoryStream();
            this.image.Save(memoryStream, Sd.Imaging.ImageFormat.Png);
            memoryStream.Position = 0;
            PP.IShape shape;
            if (this.index >= 0)
            {
                shape = slide.Shapes[this.index];
                shape.Fill.SetPicture(memoryStream);
            }
            else
            {
                slide.Shapes.AddPicture(memoryStream);
                shape = slide.Shapes.Last();
            }
            memoryStream.Dispose();

            shape.SetGraphics(this.graphic);
        }

        public void RenderShape(PP.ISlide slide)
        {
            PP.IShape shape;
            if (this.index >= 0)
            {
                shape = slide.Shapes[this.index];
            }
            else
            {
                shape = this.AddNewShape(slide);
                shape.SetGraphics(this.graphic);
            }
            shape.GeometryType = this.geometry;
            this.AddText(shape.TextBox.Paragraphs[0], this.text);
        }

        public void RenderLine(PP.ISlide slide, int height)
        {
            slide.Shapes.AddLine((int)this.line.FromX, (int)(height - this.line.FromY), (int)this.line.ToX, (int)(height - this.line.ToY));

            PP.IShape shape = slide.Shapes.Last();
            shape.SetGraphics(this.graphic);
        }

        public void RenderBarChart(PP.ISlide slide)
        {
            slide.Shapes.AddBarChart(0, 0, 100, 100, categories, seriesName);
            PP.IShape shape = slide.Shapes.Last();
            //shape.SetGraphics(this.graphic);
        }

        public void RenderBarStackChart(PP.ISlide slide)
        {
            slide.Shapes.AddStackedColumnChart(0, 0, 100, 100, categoryList.ToIListDictionary(), seriesNames.ToIList());
        }

        public void RenderStackedBarChart(PP.ISlide slide)
        {
            slide.Shapes.AddBarChart(0, 0, 100, 100, categories, seriesName);
        }

        public void RenderPieChart(PP.ISlide slide)
        {
            slide.Shapes.AddPieChart(0, 0, 100, 100, categories, seriesName);
        }

        public void RenderTable(PP.ISlide slide)
        {
            slide.Shapes.AddTable(0, 0, values.Count, values[0].Count, this.tableStyle);
            PP.ITable table = (PP.ITable)slide.Shapes.Last();
            decimal w = (decimal)(Math.Abs(this.boundary.Width) / values.Count);

            for (int i = 0; i < values.Count; i++)
            {
                table.Columns[i].Width = w;
                for (int j = 0; j < values[i].Count; j++)
                {
                    PP.ITableCell cell = table[j, i];
                    this.AddText(cell.TextBox.Paragraphs[0], values[i][j]);
                    cell.SetGraphics(values[i][j].Graphic);
                    cell.TextBox.VerticalAlignment = values[i][j].Font.VerticalAlignment.ToPpt();
                }
            }

        }

        public PP.IShape AddNewShape(PP.ISlide slide)
        {
            slide.Shapes.AddShape(0, 0, 100, 100);
            return slide.Shapes[slide.Shapes.Count - 1];
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            switch (this.ContentType)
            {
                default:
                    if (this.text.Text.Length < 16) return "PPT | " + this.ContentType + " (" + this.index + "){" + this.text.Text + "}";
                    return "PPT | " + this.ContentType + "(" + this.index + "){" + this.text.Text.Substring(0, 15) + "...}";
                case ContentTypes.List:
                    if (this.values.Count > 0) return "PPT | List(" + this.index + "){" + this.values[0].Count + "i}";
                    return "PPT | List(empty)";
                case ContentTypes.Image:
                    return "PPT | Image(" + this.index + "){" + this.image.Width + "w " + this.image.Height + "h}";
                case ContentTypes.Table:
                    if (this.values.Count > 0) return "PPT | Table(" + this.index + "){" + this.values.Count + "c " + this.values[0].Count + "r}";
                    return "PPT | Table(empty)";
                case ContentTypes.Line:
                    return "PPT | Line(" + this.index + "){" + this.line.ToString() + "}";
                case ContentTypes.Shape:
                    return "PPT | Shape(" + this.index + "){" + this.geometry.ToString() + "}";
                case ContentTypes.ChartBar:
                case ContentTypes.ChartBarStack:
                case ContentTypes.ChartPie:
                case ContentTypes.ChartScatter:
                    return "PPT | Chart(" + this.index + "){" + this.ContentType + "}";
                case ContentTypes.Placeholder:
                    return "PPT | Placeholder(" + this.index + "){" + this.placeholder.ToString() + "}";
            }
        }

        #endregion

    }
}
