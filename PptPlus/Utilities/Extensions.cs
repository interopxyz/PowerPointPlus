using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;
using Rg = Rhino.Geometry;

using PP = ShapeCrawler;
using XD = DocumentFormat.OpenXml.Drawing;
using Grasshopper.Kernel.Types;
using Grasshopper.Kernel.Data;

namespace PptPlus
{
    public static class Extensions
    {

        #region geometry

        public static double GetHeight(this Rg.Rectangle3d input)
        {
            Rg.Point3d pt0 = input.Corner(0);
            Rg.Point3d pt1 = input.Corner(2);

            return (int)Math.Abs(pt1.Y - pt0.Y);
        }

        public static void SetSize(this PP.IShape input, Rg.Rectangle3d contentBounds, Rg.Rectangle3d slideBounds)
        {
            int height = (int)Math.Abs(slideBounds.Height);

            Rg.Point3d pt0 = contentBounds.Corner(0);
            Rg.Point3d pt1 = contentBounds.Corner(2);

            int h = (int)Math.Abs(pt1.Y - pt0.Y);

            input.X = (int)pt0.X;
            input.Y = height-h - (int)pt0.Y;

            input.Width = (int)Math.Abs(pt1.X - pt0.X);
            input.Height= h;
        }

        public static Font GetFont(this PP.IParagraphPortion input)
        {
            Font output = new Font();
            if(input.Font.LatinName!=null) output.Name = input.Font.LatinName;
            output.Size = (double)input.Font.Size;
            output.Bold = input.Font.IsBold;
            output.Italic = input.Font.IsItalic;
            output.Underlined = input.Font.Underline != XD.TextUnderlineValues.None;
            output.Color = input.Font.Color.ToColor();
            output.Highlight = input.TextHighlightColor.ToColor();

            return output;
        }

        public static void SetGraphics(this PP.IShape input, Graphic graphic, bool renderFill = true)
        {
            if(renderFill) {
            if (graphic.NoFill)
            {
                input.Fill.SetNoFill();
            }
            else
            {
                input.Fill.SetColor(graphic.Fill.ToHex());
            }
            }

            if (graphic.NoStroke)
            {
                input.Outline.SetNoOutline();
            }
            else
            {
                input.Outline.SetHexColor(graphic.Stroke.ToHex());
                input.Outline.Weight = (decimal)(graphic.Weight);
            }
        }

        public static void SetGraphics(this PP.ITable input, Graphic graphic)
        {

            if (graphic.NoFill)
            {
                input.Fill.SetNoFill();
            }
            else
            {
                //input.Fill.SetColor(graphic.Fill.ToHex());
            }

            if (graphic.NoStroke)
            {
                input.Outline.SetNoOutline();
            }
            else
            {
                input.Outline.SetHexColor(graphic.Stroke.ToHex());
                input.Outline.Weight = (decimal)(graphic.Weight);
            }

        }

        public static void SetGraphics(this PP.ITableCell input, Graphic graphic)
        {

            if (graphic.NoFill)
            {
                input.Fill.SetNoFill();
            }
            else
            {
                input.Fill.SetColor(graphic.Fill.ToHex());
            }

            input.UpdateBorders(graphic);

        }

        public static void UpdateBorders(this PP.ITableCell input, Graphic graphic)
        {
            if (graphic.HasStroke)
            {
                if (graphic.LeftBorder)
                {
                    input.LeftBorder.Color = graphic.Stroke.ToHex();
                    input.LeftBorder.Width = (decimal)graphic.Weight;
                }
                if (graphic.RightBorder)
                {
                    input.RightBorder.Color = graphic.Stroke.ToHex();
                    input.RightBorder.Width = (decimal)graphic.Weight;
                }
                if (graphic.TopBorder)
                {
                    input.TopBorder.Color = graphic.Stroke.ToHex();
                    input.TopBorder.Width = (decimal)graphic.Weight;
                }
                if (graphic.BottomBorder)
                {
                    input.BottomBorder.Color = graphic.Stroke.ToHex();
                    input.BottomBorder.Width = (decimal)graphic.Weight;
                }
            }
        }

        public static string ToHex(this Sd.Color color)
        {
            return $"{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        public static PP.Color ToPpt(this Sd.Color color)
        {
            return PP.Color.FromHex(color.ToHex());
        }

        public static Sd.Color ToColor(this PP.IFontColor input)
        {
            return Sd.ColorTranslator.FromHtml(input.Hex);
        }

        public static Sd.Color ToColor(this PP.Color input)
        {
            return Sd.ColorTranslator.FromHtml(input.Hex);
        }

        #endregion

        #region cloning

        public static Rg.Rectangle3d Duplicate(this Rg.Rectangle3d input)
        {
            return new Rg.Rectangle3d(input.Plane, input.X, input.Y);
        }

        public static List<string> Duplicate(this List<string> input)
        {
            List<string> output = new List<string>();
            foreach (string txt in input) output.Add(txt);
            return output;
        }

        public static Dictionary<string, string> Duplicate(this Dictionary<string, string> input)
        {
            Dictionary<string, string> output = new Dictionary<string, string>();
            foreach (KeyValuePair<string, string> item in input) output.Add(item.Key, item.Value);
            return output;
        }

        public static Dictionary<string,double> Duplicate(this Dictionary<string, double> input)
        {
            Dictionary<string, double> output = new Dictionary<string, double>();
            foreach(KeyValuePair<string,double> item in input) output.Add(item.Key, item.Value);
            return output;
        }

        public static Dictionary<string, List<double>> Duplicate(this Dictionary<string, List<double>> input)
        {
            Dictionary<string, List<double>> output = new Dictionary<string, List<double>>();
            foreach (KeyValuePair<string, List<double>> item in input)
            {
                output.Add(item.Key, new List<double>());
                foreach (double value in item.Value) output[item.Key].Add(value);
            }
            return output;
        }

        public static Dictionary<string, IList<double>> ToIListDictionary(this Dictionary<string, List<double>> input)
        {
            Dictionary<string, IList<double>> output = new Dictionary<string, IList<double>>();
            foreach (KeyValuePair<string, List<double>> item in input)
            {
                IList<double> values = new List<double>();
                foreach (double value in item.Value) values.Add(value);
                output.Add(item.Key, values);
            }
            return output;
        }

        public static List<List<double>> ToListOfLists(this Dictionary<string, List<double>> input)
        {
            List<List<double>> output = new List<List<double>>();
            foreach (KeyValuePair<string, List<double>> item in input)
            {
                List<double> values = new List<double>();
                foreach (double value in item.Value) values.Add(value);
                output.Add(values);
            }
            return output;
        }

        public static Dictionary<double, double> Duplicate(this Dictionary<double, double> input)
        {
            Dictionary<double, double> output = new Dictionary<double, double>();
            foreach (KeyValuePair<double, double> item in input) output.Add(item.Key, item.Value);
            return output;
        }

        public static IList<double> ToIList(this List<double> input)
        {
            IList<double> output = new List<double>();
            foreach (double value in input) output.Add(value);
            return output;
        }

        public static IList<string> ToIList(this List<string> input)
        {
            IList<string> output = new List<string>();
            foreach (string value in input) output.Add(value);
            return output;
        }

        public static List<List<string>> Duplicate(this List<List<string>> input)
        {
            List<List<string>> output = new List<List<string>>();
            foreach (List<string> values in input)
            {
                List<string> vals = new List<string>();
                foreach (string value in values)
                {
                    vals.Add(value);
                }
                output.Add(vals);
            }
            return output;
        }

        public static List<List<Paragraph>> Duplicate(this List<List<Paragraph>> input)
        {
            List<List<Paragraph>> output = new List<List<Paragraph>>();
            foreach (List<Paragraph> values in input)
            {
                List<Paragraph> vals = new List<Paragraph>();
                foreach (Paragraph value in values)
                {
                    vals.Add(new Paragraph(value));
                }
                output.Add(vals);
            }
            return output;
        }

        public static List<Content> Duplicate(this List<Content> input)
        {
            List<Content> output = new List<Content>();
            foreach (Content item in input) output.Add(new Content(item));
            return output;
        }

        public static List<PpSlide> Duplicate(this List<PpSlide> input)
        {
            List<PpSlide> output = new List<PpSlide>();
            foreach (PpSlide item in input) output.Add(new PpSlide(item));
            return output;
        }

        public static List<Fragment> Duplicate(this List<Fragment> input)
        {
            List<Fragment> output = new List<Fragment>();
            foreach (Fragment item in input) output.Add(new Fragment(item));
            return output;
        }

        public static Dictionary<string, Sd.Color> Duplicate(this Dictionary<string, Sd.Color> input)
        {
            Dictionary<string, Sd.Color> output = new Dictionary<string, Sd.Color>();
            foreach (KeyValuePair<string, Sd.Color> item in input) output.Add(item.Key, item.Value);
            return output;
        }

        #endregion

        #region casting

        public static bool TryGetContent(this IGH_Goo input, out Content content)
        {
            if (input.CastTo<Content>(out Content content1))
            {
                content = new Content(content1);
                return true;
            }
            else if (input.CastTo<Paragraph>(out Paragraph paragraph))
            {
                content = Content.CreateTextContent(paragraph);
                return true;
            }
            else if (input.CastTo<Fragment>(out Fragment fragment1))
            {
                content = Content.CreateTextContent(fragment1);
                return true;
            }
            else if (input.CastTo<string>(out string text))
            {
                content = Content.CreateTextContent(text);
                return true;
            }
            else if (input.CastTo<double>(out double number))
            {
                content = Content.CreateTextContent(number.ToString());
                return true;
            }
            else if (input.CastTo<int>(out int integer))
            {
                content = Content.CreateTextContent(integer.ToString());
                return true;
            }
            content = null;
            return false;
        }

        public static bool TryCastToParagraph(this IGH_Goo input, out Paragraph paragraph)
        {
            if (input.CastTo<Paragraph>(out Paragraph paragraph1))
            {
                paragraph = new Paragraph(paragraph1);
                return true;
            }
            paragraph = null;
            return false;

        }

        public static bool TryGetParagraph(this IGH_Goo input, out Paragraph paragraph)
        {
            if (input.CastTo<Paragraph>(out Paragraph paragraph1))
            {
                paragraph = new Paragraph(paragraph1);
                return true;
            }
            else if (input.CastTo<Fragment>(out Fragment fragment1))
            {
                paragraph = new Paragraph(fragment1);
                return true;
            }
            else if (input.CastTo<string>(out string text))
            {
                paragraph = new Paragraph(text);
                return true;
            }
            else if (input.CastTo<double>(out double number))
            {
                paragraph = new Paragraph(number.ToString());
                return true;
            }
            else if (input.CastTo<int>(out int integer))
            {
                paragraph = new Paragraph(integer.ToString());
                return true;
            }
            paragraph = null;
            return false;
        }

        public static bool TryGetFragment(this IGH_Goo input, out Fragment fragment)
        {
            if (input.CastTo<Fragment>(out Fragment fragment1))
            {
                fragment = new Fragment(fragment1);
                return true;
            }
            else if (input.CastTo<string>(out string text))
            {
                fragment = new Fragment(text);
                return true;
            }
            else if (input.CastTo<double>(out double number))
            {
                fragment = new Fragment(number.ToString());
                return true;
            }
            else if (input.CastTo<int>(out int integer))
            {
                fragment = new Fragment(integer.ToString());
                return true;
            }
            fragment = null;
            return false;
        }

        #endregion

        #region units

        //Inches
        public static double PointToInch(this double input)
        {
            return input / 72.0;
        }

        public static double InchToPoint(this double input)
        {
            return input * 72.0;
        }

        //Millimeters
        public static double PointToMillimeter(this double input)
        {
            return input * 25.4 / 72.0;
        }

        public static double MillimeterToPoint(this double input)
        {
            return input * 72 / 25.4;
        }

        //Centimeter
        public static double PointToCentimeter(this double input)
        {
            return input * 2.54 / 72.0;
        }

        public static double CentimeterToPoint(this double input)
        {
            return input * 72.0 / 2.54;
        }

        //Pixels
        public static double PointToPixels(this double input, int dpi = 96)
        {
            return input / 72.0 * dpi;
        }

        public static double PixelsToPoint(this double input, int dpi = 96)
        {
            return input * 72.0 / dpi;
        }

        #endregion

        #region conversion

        public static PP.TextHorizontalAlignment ToPpt(this Font.HorizontalAlignments input)
        {
            switch (input)
            {
                default:
                    return PP.TextHorizontalAlignment.Left;
                case Font.HorizontalAlignments.Center:
                    return PP.TextHorizontalAlignment.Center;
                case Font.HorizontalAlignments.Right:
                    return PP.TextHorizontalAlignment.Right;
            }
        }

        public static Font.HorizontalAlignments ToFont(this PP.TextHorizontalAlignment input)
        {
            switch (input)
            {
                default:
                    return Font.HorizontalAlignments.Left;
                case PP.TextHorizontalAlignment.Center:
                    return Font.HorizontalAlignments.Center;
                case PP.TextHorizontalAlignment.Right:
                    return Font.HorizontalAlignments.Right;
            }
        }

        public static PP.TextVerticalAlignment ToPpt(this Font.VerticalAlignments input)
        {
            switch (input)
            {
                default:
                    return PP.TextVerticalAlignment.Bottom;
                case Font.VerticalAlignments.Middle:
                    return PP.TextVerticalAlignment.Middle;
                case Font.VerticalAlignments.Top:
                    return PP.TextVerticalAlignment.Top;
            }
        }

        public static Font.VerticalAlignments ToFont(this PP.TextVerticalAlignment input)
        {
            switch (input)
            {
                default:
                    return Font.VerticalAlignments.Default;
                case PP.TextVerticalAlignment.Bottom:
                    return Font.VerticalAlignments.Bottom;
                case PP.TextVerticalAlignment.Middle:
                    return Font.VerticalAlignments.Middle;
                case PP.TextVerticalAlignment.Top:
                    return Font.VerticalAlignments.Top;
            }
        }

        public static string GetUnicode(this Paragraph.BulletPoints input)
        {
            switch (input)
            {
                default:
                    return "•";
                case Paragraph.BulletPoints.HollowBullet://
                    return "○";
                case Paragraph.BulletPoints.Bar://
                    return "⁃";
                case Paragraph.BulletPoints.Box://
                    return "■";
                case Paragraph.BulletPoints.HollowBox:
                    return "❒";
                case Paragraph.BulletPoints.Arrow:
                    return "➣";
                case Paragraph.BulletPoints.Star:
                    return "❖";
                case Paragraph.BulletPoints.Check:
                    return "✓";
            }
        }

        public static GH_Structure<GH_ObjectWrapper> ToDataTree(this List<List<Paragraph>> input, GH_Path path)
        {
            GH_Structure<GH_ObjectWrapper> ghData = new GH_Structure<GH_ObjectWrapper>();
            for (int i = 0; i < input.Count; i++)
            {
                for (int j = 0; j < input[i].Count; j++)
                {
                    ghData.Append(new GH_ObjectWrapper(input[i][j]), path.AppendElement(i));
                }
            }
            return ghData;
        }

        public static GH_Structure<GH_ObjectWrapper> ToDataTree(this List<List<Content>> input, GH_Path path)
        {
            GH_Structure<GH_ObjectWrapper> ghData = new GH_Structure<GH_ObjectWrapper>();
            for (int i = 0; i < input.Count; i++)
            {
                for (int j = 0; j < input[i].Count; j++)
                {
                    ghData.Append(new GH_ObjectWrapper(input[i][j]), path.AppendElement(i));
                }
            }
            return ghData;
        }

        public static GH_Structure<GH_Number> ToDataTree(this List<List<double>> input, GH_Path path)
        {
            GH_Structure<GH_Number> ghData = new GH_Structure<GH_Number>();
            for (int i = 0; i < input.Count; i++)
            {
                for (int j = 0; j < input[i].Count; j++)
                {
                    ghData.Append(new GH_Number(input[i][j]), path.AppendElement(i));
                }
            }
            return ghData;
        }

        #endregion

        #region extract

        public static Graphic GetGraphic(this PP.IShape input)
        {
            Graphic output = new Graphic();

            if(input.Fill.Type != PP.FillType.NoFill)
            {
                output.Fill = Sd.ColorTranslator.FromHtml(input.Fill.Color);
            }
            if(input.Outline.Weight != 0)
            {
                output.Stroke = Sd.ColorTranslator.FromHtml(input.Outline.HexColor);
                output.Weight = (double)input.Width;
            }

            return output;
        }

        public static Font GetFont(this PP.IShape input)
        {
            Font output = new Font();
            bool isCustom = true;
            if(input.PlaceholderType.HasValue)
            {
                switch (input.PlaceholderType.Value)
                {
                    default:
                        output = Fonts.Normal;
                        break;
                    case PP.PlaceholderType.Title:
                        output = Fonts.Title;
                        output.HorizontalAlignment = Font.HorizontalAlignments.Center;
                        isCustom = false;
                        break;
                    case PP.PlaceholderType.SubTitle:
                        output = Fonts.Subtitle;
                        output.HorizontalAlignment = Font.HorizontalAlignments.Center;
                        isCustom = false;
                        break;
                };
            }

            output.VerticalAlignment = input.TextBox.VerticalAlignment.ToFont();
            output.Wrapped = input.TextBox.TextWrapped;

            if (isCustom)
            {
                if (input.TextBox.Paragraphs.Count > 0)
                {
                    if (input.TextBox.Paragraphs[0].Spacing.LineSpacingPoints.HasValue) output.LineSpacing = (double)input.TextBox.Paragraphs[0].Spacing.LineSpacingPoints.Value;
                    output.HorizontalAlignment = input.TextBox.Paragraphs[0].HorizontalAlignment.ToFont();

                    if (input.TextBox.Paragraphs[0].Portions.Count > 0)
                    {
                        PP.ITextPortionFont portionFont = input.TextBox.Paragraphs[0].Portions[0].Font;
                        if (portionFont.EastAsianName != "")
                        {
                            output.Family = portionFont.LatinName;
                            output.Size = (double)input.TextBox.Paragraphs[0].Portions[0].Font.Size;

                            output.Bold = input.TextBox.Paragraphs[0].Portions[0].Font.IsBold;
                            output.Italic = input.TextBox.Paragraphs[0].Portions[0].Font.IsItalic;
                            output.Underlined = input.TextBox.Paragraphs[0].Portions[0].Font.Underline != XD.TextUnderlineValues.None;

                            output.Color = input.TextBox.Paragraphs[0].Portions[0].Font.Color.ToColor();
                            output.Highlight = input.TextBox.Paragraphs[0].Portions[0].TextHighlightColor.ToColor();
                        }
                    }
                }
            }

            return output;
        }

        #endregion

        #region shapes

        public static PP.ITableStyle ToStyle(this Content.TableStyles input)
        {
            switch (input)
            {
                default:
                    return PP.Tables.CommonTableStyles.NoStyleTableGrid;
                case Content.TableStyles.Blank:
                    return PP.Tables.CommonTableStyles.NoStyleNoGrid;
                case Content.TableStyles.Dark1:
                    return PP.Tables.CommonTableStyles.DarkStyle1;
                case Content.TableStyles.Dark2:
                    return PP.Tables.CommonTableStyles.DarkStyle2;
                case Content.TableStyles.Medium1:
                    return PP.Tables.CommonTableStyles.MediumStyle1;
                case Content.TableStyles.Medium2:
                    return PP.Tables.CommonTableStyles.MediumStyle2;
                case Content.TableStyles.Medium3:
                    return PP.Tables.CommonTableStyles.MediumStyle3;
                case Content.TableStyles.Light1:
                    return PP.Tables.CommonTableStyles.LightStyle1;
                case Content.TableStyles.Light2:
                    return PP.Tables.CommonTableStyles.LightStyle2;
                case Content.TableStyles.Light3:
                    return PP.Tables.CommonTableStyles.LightStyle3;
                case Content.TableStyles.Themed:
                    return PP.Tables.CommonTableStyles.ThemedStyle1Accent1;
            }
        }

        public static PP.Geometry ToGeometry(this Content.ShapeLines input)
        {
            switch (input)
            {
                default:
                    return PP.Geometry.Line;
                case Content.ShapeLines.LineInverse:
                    return PP.Geometry.LineInverse;
                case Content.ShapeLines.StraightConnector1:
                    return PP.Geometry.StraightConnector1;
                case Content.ShapeLines.BentConnector2:
                    return PP.Geometry.BentConnector2;
                case Content.ShapeLines.BentConnector3:
                    return PP.Geometry.BentConnector3;
                case Content.ShapeLines.BentConnector4:
                    return PP.Geometry.BentConnector4;
                case Content.ShapeLines.BentConnector5:
                    return PP.Geometry.BentConnector5;
                case Content.ShapeLines.CurvedConnector2:
                    return PP.Geometry.CurvedConnector2;
                case Content.ShapeLines.CurvedConnector3:
                    return PP.Geometry.CurvedConnector3;
                case Content.ShapeLines.CurvedConnector4:
                    return PP.Geometry.CurvedConnector4;
                case Content.ShapeLines.CurvedConnector5:
                    return PP.Geometry.CurvedConnector5;
            }
        }

        public static PP.Geometry ToGeometry(this Content.ShapeRectangles input)
        {
            switch (input)
            {
                default:
                    return PP.Geometry.RoundedRectangle;
                case Content.ShapeRectangles.SingleCornerRoundedRectangle:
                    return PP.Geometry.SingleCornerRoundedRectangle;
                case Content.ShapeRectangles.TopCornersRoundedRectangle:
                    return PP.Geometry.TopCornersRoundedRectangle;
                case Content.ShapeRectangles.DiagonalCornersRoundedRectangle:
                    return PP.Geometry.DiagonalCornersRoundedRectangle;
                case Content.ShapeRectangles.SnipRoundRectangle:
                    return PP.Geometry.SnipRoundRectangle;
                case Content.ShapeRectangles.Snip1Rectangle:
                    return PP.Geometry.Snip1Rectangle;
                case Content.ShapeRectangles.Snip2SameRectangle:
                    return PP.Geometry.Snip2SameRectangle;
                case Content.ShapeRectangles.Snip2DiagonalRectangle:
                    return PP.Geometry.Snip2DiagonalRectangle;
                case Content.ShapeRectangles.Plaque:
                    return PP.Geometry.Plaque;
                case Content.ShapeRectangles.Ellipse:
                    return PP.Geometry.Ellipse;
            }
        }

        public static PP.Geometry ToGeometry(this Content.ShapePolygons input)
        {
            switch (input)
            {
                default:
                    return PP.Geometry.Triangle;
                case Content.ShapePolygons.RightTriangle:
                    return PP.Geometry.RightTriangle;
                case Content.ShapePolygons.Rectangle:
                    return PP.Geometry.Rectangle;
                case Content.ShapePolygons.Diamond:
                    return PP.Geometry.Diamond;
                case Content.ShapePolygons.Parallelogram:
                    return PP.Geometry.Parallelogram;
                case Content.ShapePolygons.Trapezoid:
                    return PP.Geometry.Trapezoid;
                case Content.ShapePolygons.NonIsoscelesTrapezoid:
                    return PP.Geometry.NonIsoscelesTrapezoid;
                case Content.ShapePolygons.Pentagon:
                    return PP.Geometry.Pentagon;
                case Content.ShapePolygons.Hexagon:
                    return PP.Geometry.Hexagon;
                case Content.ShapePolygons.Heptagon:
                    return PP.Geometry.Heptagon;
                case Content.ShapePolygons.Octagon:
                    return PP.Geometry.Octagon;
                case Content.ShapePolygons.Decagon:
                    return PP.Geometry.Decagon;
                case Content.ShapePolygons.Dodecagon:
                    return PP.Geometry.Dodecagon;
            }
        }

        public static PP.Geometry ToGeometry(this Content.ShapeStars input)
        {
            switch (input)
            {
                default:
                    return PP.Geometry.Star4;
                case Content.ShapeStars.Star5:
                    return PP.Geometry.Star5;
                case Content.ShapeStars.Star6:
                    return PP.Geometry.Star6;
                case Content.ShapeStars.Star7:
                    return PP.Geometry.Star7;
                case Content.ShapeStars.Star8:
                    return PP.Geometry.Star8;
                case Content.ShapeStars.Star10:
                    return PP.Geometry.Star10;
                case Content.ShapeStars.Star12:
                    return PP.Geometry.Star12;
                case Content.ShapeStars.Star16:
                    return PP.Geometry.Star16;
                case Content.ShapeStars.Star24:
                    return PP.Geometry.Star24;
                case Content.ShapeStars.Star32:
                    return PP.Geometry.Star32;
            }
        }

        public static PP.Geometry ToGeometry(this Content.ShapeArrows input)
        {
            switch (input)
            {
                default:
                    return PP.Geometry.RightArrow;
                case Content.ShapeArrows.LeftArrow:
                    return PP.Geometry.LeftArrow;
                case Content.ShapeArrows.UpArrow:
                    return PP.Geometry.UpArrow;
                case Content.ShapeArrows.DownArrow:
                    return PP.Geometry.DownArrow;
                case Content.ShapeArrows.StripedRightArrow:
                    return PP.Geometry.StripedRightArrow;
                case Content.ShapeArrows.NotchedRightArrow:
                    return PP.Geometry.NotchedRightArrow;
                case Content.ShapeArrows.BentUpArrow:
                    return PP.Geometry.BentUpArrow;
                case Content.ShapeArrows.LeftRightArrow:
                    return PP.Geometry.LeftRightArrow;
                case Content.ShapeArrows.UpDownArrow:
                    return PP.Geometry.UpDownArrow;
                case Content.ShapeArrows.LeftUpArrow:
                    return PP.Geometry.LeftUpArrow;
                case Content.ShapeArrows.LeftRightUpArrow:
                    return PP.Geometry.LeftRightUpArrow;
                case Content.ShapeArrows.QuadArrow:
                    return PP.Geometry.QuadArrow;
                case Content.ShapeArrows.LeftArrowCallout:
                    return PP.Geometry.LeftArrowCallout;
                case Content.ShapeArrows.RightArrowCallout:
                    return PP.Geometry.RightArrowCallout;
                case Content.ShapeArrows.UpArrowCallout:
                    return PP.Geometry.UpArrowCallout;
                case Content.ShapeArrows.DownArrowCallout:
                    return PP.Geometry.DownArrowCallout;
                case Content.ShapeArrows.LeftRightArrowCallout:
                    return PP.Geometry.LeftRightArrowCallout;
                case Content.ShapeArrows.UpDownArrowCallout:
                    return PP.Geometry.UpDownArrowCallout;
                case Content.ShapeArrows.QuadArrowCallout:
                    return PP.Geometry.QuadArrowCallout;
                case Content.ShapeArrows.BentArrow:
                    return PP.Geometry.BentArrow;
                case Content.ShapeArrows.UTurnArrow:
                    return PP.Geometry.UTurnArrow;
                case Content.ShapeArrows.CircularArrow:
                    return PP.Geometry.CircularArrow;
                case Content.ShapeArrows.LeftCircularArrow:
                    return PP.Geometry.LeftCircularArrow;
                case Content.ShapeArrows.LeftRightCircularArrow:
                    return PP.Geometry.LeftRightCircularArrow;
                case Content.ShapeArrows.CurvedRightArrow:
                    return PP.Geometry.CurvedRightArrow;
                case Content.ShapeArrows.CurvedLeftArrow:
                    return PP.Geometry.CurvedLeftArrow;
                case Content.ShapeArrows.CurvedUpArrow:
                    return PP.Geometry.CurvedUpArrow;
                case Content.ShapeArrows.CurvedDownArrow:
                    return PP.Geometry.CurvedDownArrow;
                case Content.ShapeArrows.SwooshArrow:
                    return PP.Geometry.SwooshArrow;
            }
        }

        public static PP.Geometry ToGeometry(this Content.ShapeEquations input)
        {
            switch (input)
            {
                default:
                    return PP.Geometry.MathPlus;
                case Content.ShapeEquations.MathMinus:
                    return PP.Geometry.MathMinus;
                case Content.ShapeEquations.MathMultiply:
                    return PP.Geometry.MathMultiply;
                case Content.ShapeEquations.MathDivide:
                    return PP.Geometry.MathDivide;
                case Content.ShapeEquations.MathEqual:
                    return PP.Geometry.MathEqual;
                case Content.ShapeEquations.MathNotEqual:
                    return PP.Geometry.MathNotEqual;
            }
        }

        public static PP.Geometry ToGeometry(this Content.ShapeFlowchart input)
        {
            switch (input)
            {
                default:
                    return PP.Geometry.FlowChartProcess;
                case Content.ShapeFlowchart.FlowChartDecision:
                    return PP.Geometry.FlowChartDecision;
                case Content.ShapeFlowchart.FlowChartInputOutput:
                    return PP.Geometry.FlowChartInputOutput;
                case Content.ShapeFlowchart.FlowChartPredefinedProcess:
                    return PP.Geometry.FlowChartPredefinedProcess;
                case Content.ShapeFlowchart.FlowChartInternalStorage:
                    return PP.Geometry.FlowChartInternalStorage;
                case Content.ShapeFlowchart.FlowChartDocument:
                    return PP.Geometry.FlowChartDocument;
                case Content.ShapeFlowchart.FlowChartMultidocument:
                    return PP.Geometry.FlowChartMultidocument;
                case Content.ShapeFlowchart.FlowChartTerminator:
                    return PP.Geometry.FlowChartTerminator;
                case Content.ShapeFlowchart.FlowChartPreparation:
                    return PP.Geometry.FlowChartPreparation;
                case Content.ShapeFlowchart.FlowChartManualInput:
                    return PP.Geometry.FlowChartManualInput;
                case Content.ShapeFlowchart.FlowChartManualOperation:
                    return PP.Geometry.FlowChartManualOperation;
                case Content.ShapeFlowchart.FlowChartConnector:
                    return PP.Geometry.FlowChartConnector;
                case Content.ShapeFlowchart.FlowChartPunchedCard:
                    return PP.Geometry.FlowChartPunchedCard;
                case Content.ShapeFlowchart.FlowChartPunchedTape:
                    return PP.Geometry.FlowChartPunchedTape;
                case Content.ShapeFlowchart.FlowChartSummingJunction:
                    return PP.Geometry.FlowChartSummingJunction;
                case Content.ShapeFlowchart.FlowChartOr:
                    return PP.Geometry.FlowChartOr;
                case Content.ShapeFlowchart.FlowChartCollate:
                    return PP.Geometry.FlowChartCollate;
                case Content.ShapeFlowchart.FlowChartSort:
                    return PP.Geometry.FlowChartSort;
                case Content.ShapeFlowchart.FlowChartExtract:
                    return PP.Geometry.FlowChartExtract;
                case Content.ShapeFlowchart.FlowChartMerge:
                    return PP.Geometry.FlowChartMerge;
                case Content.ShapeFlowchart.FlowChartOfflineStorage:
                    return PP.Geometry.FlowChartOfflineStorage;
                case Content.ShapeFlowchart.FlowChartOnlineStorage:
                    return PP.Geometry.FlowChartOnlineStorage;
                case Content.ShapeFlowchart.FlowChartMagneticTape:
                    return PP.Geometry.FlowChartMagneticTape;
                case Content.ShapeFlowchart.FlowChartMagneticDisk:
                    return PP.Geometry.FlowChartMagneticDisk;
                case Content.ShapeFlowchart.FlowChartMagneticDrum:
                    return PP.Geometry.FlowChartMagneticDrum;
                case Content.ShapeFlowchart.FlowChartDisplay:
                    return PP.Geometry.FlowChartDisplay;
                case Content.ShapeFlowchart.FlowChartDelay:
                    return PP.Geometry.FlowChartDelay;
                case Content.ShapeFlowchart.FlowChartAlternateProcess:
                    return PP.Geometry.FlowChartAlternateProcess;
                case Content.ShapeFlowchart.FlowChartOffpageConnector:
                    return PP.Geometry.FlowChartOffpageConnector;
            }
        }

        public static PP.Geometry ToGeometry(this Content.ShapeBanners input)
        {
            switch (input)
            {
                default:
                    return PP.Geometry.Ribbon;
                case Content.ShapeBanners.Ribbon2:
                    return PP.Geometry.Ribbon2;
                case Content.ShapeBanners.EllipseRibbon:
                    return PP.Geometry.EllipseRibbon;
                case Content.ShapeBanners.EllipseRibbon2:
                    return PP.Geometry.EllipseRibbon2;
                case Content.ShapeBanners.LeftRightRibbon:
                    return PP.Geometry.LeftRightRibbon;
                case Content.ShapeBanners.VerticalScroll:
                    return PP.Geometry.VerticalScroll;
                case Content.ShapeBanners.HorizontalScroll:
                    return PP.Geometry.HorizontalScroll;
                case Content.ShapeBanners.Wave:
                    return PP.Geometry.Wave;
                case Content.ShapeBanners.DoubleWave:
                    return PP.Geometry.DoubleWave;
                case Content.ShapeBanners.Plus:
                    return PP.Geometry.Plus;
            }
        }

        public static PP.Geometry ToGeometry(this Content.ShapeCallouts input)
        {
            switch (input)
            {
                default:
                    return PP.Geometry.Callout1;
                case Content.ShapeCallouts.Callout2:
                    return PP.Geometry.Callout2;
                case Content.ShapeCallouts.Callout3:
                    return PP.Geometry.Callout3;
                case Content.ShapeCallouts.AccentCallout1:
                    return PP.Geometry.AccentCallout1;
                case Content.ShapeCallouts.AccentCallout2:
                    return PP.Geometry.AccentCallout2;
                case Content.ShapeCallouts.AccentCallout3:
                    return PP.Geometry.AccentCallout3;
                case Content.ShapeCallouts.BorderCallout1:
                    return PP.Geometry.BorderCallout1;
                case Content.ShapeCallouts.BorderCallout2:
                    return PP.Geometry.BorderCallout2;
                case Content.ShapeCallouts.BorderCallout3:
                    return PP.Geometry.BorderCallout3;
                case Content.ShapeCallouts.AccentBorderCallout1:
                    return PP.Geometry.AccentBorderCallout1;
                case Content.ShapeCallouts.AccentBorderCallout2:
                    return PP.Geometry.AccentBorderCallout2;
                case Content.ShapeCallouts.AccentBorderCallout3:
                    return PP.Geometry.AccentBorderCallout3;
                case Content.ShapeCallouts.WedgeRectangleCallout:
                    return PP.Geometry.WedgeRectangleCallout;
                case Content.ShapeCallouts.WedgeRoundRectangleCallout:
                    return PP.Geometry.WedgeRoundRectangleCallout;
                case Content.ShapeCallouts.WedgeEllipseCallout:
                    return PP.Geometry.WedgeEllipseCallout;
                case Content.ShapeCallouts.CloudCallout:
                    return PP.Geometry.CloudCallout;
                case Content.ShapeCallouts.Cloud:
                    return PP.Geometry.Cloud;
            }
        }

        public static PP.Geometry ToGeometry(this Content.ShapeActions input)
        {
            switch (input)
            {
                default:
                    return PP.Geometry.ActionButtonBlank;
                case Content.ShapeActions.ActionButtonHome:
                    return PP.Geometry.ActionButtonHome;
                case Content.ShapeActions.ActionButtonHelp:
                    return PP.Geometry.ActionButtonHelp;
                case Content.ShapeActions.ActionButtonInformation:
                    return PP.Geometry.ActionButtonInformation;
                case Content.ShapeActions.ActionButtonForwardNext:
                    return PP.Geometry.ActionButtonForwardNext;
                case Content.ShapeActions.ActionButtonBackPrevious:
                    return PP.Geometry.ActionButtonBackPrevious;
                case Content.ShapeActions.ActionButtonEnd:
                    return PP.Geometry.ActionButtonEnd;
                case Content.ShapeActions.ActionButtonBeginning:
                    return PP.Geometry.ActionButtonBeginning;
                case Content.ShapeActions.ActionButtonReturn:
                    return PP.Geometry.ActionButtonReturn;
                case Content.ShapeActions.ActionButtonDocument:
                    return PP.Geometry.ActionButtonDocument;
                case Content.ShapeActions.ActionButtonSound:
                    return PP.Geometry.ActionButtonSound;
                case Content.ShapeActions.ActionButtonMovie:
                    return PP.Geometry.ActionButtonMovie;
            }
        }

        public static PP.Geometry ToGeometry(this Content.ShapeBasics input)
        {
            switch (input)
            {
                default:
                    return PP.Geometry.Teardrop;
                case Content.ShapeBasics.HomePlate:
                    return PP.Geometry.HomePlate;
                case Content.ShapeBasics.Chevron:
                    return PP.Geometry.Chevron;
                case Content.ShapeBasics.PieWedge:
                    return PP.Geometry.PieWedge;
                case Content.ShapeBasics.Pie:
                    return PP.Geometry.Pie;
                case Content.ShapeBasics.BlockArc:
                    return PP.Geometry.BlockArc;
                case Content.ShapeBasics.Donut:
                    return PP.Geometry.Donut;
                case Content.ShapeBasics.NoSmoking:
                    return PP.Geometry.NoSmoking;
                case Content.ShapeBasics.Cube:
                    return PP.Geometry.Cube;
                case Content.ShapeBasics.Can:
                    return PP.Geometry.Can;
                case Content.ShapeBasics.LightningBolt:
                    return PP.Geometry.LightningBolt;
                case Content.ShapeBasics.Heart:
                    return PP.Geometry.Heart;
                case Content.ShapeBasics.Sun:
                    return PP.Geometry.Sun;
                case Content.ShapeBasics.Moon:
                    return PP.Geometry.Moon;
                case Content.ShapeBasics.SmileyFace:
                    return PP.Geometry.SmileyFace;
                case Content.ShapeBasics.IrregularSeal1:
                    return PP.Geometry.IrregularSeal1;
                case Content.ShapeBasics.IrregularSeal2:
                    return PP.Geometry.IrregularSeal2;
                case Content.ShapeBasics.FoldedCorner:
                    return PP.Geometry.FoldedCorner;
                case Content.ShapeBasics.Bevel:
                    return PP.Geometry.Bevel;
                case Content.ShapeBasics.Frame:
                    return PP.Geometry.Frame;
                case Content.ShapeBasics.HalfFrame:
                    return PP.Geometry.HalfFrame;
                case Content.ShapeBasics.Corner:
                    return PP.Geometry.Corner;
                case Content.ShapeBasics.DiagonalStripe:
                    return PP.Geometry.DiagonalStripe;
                case Content.ShapeBasics.Chord:
                    return PP.Geometry.Chord;
                case Content.ShapeBasics.Arc:
                    return PP.Geometry.Arc;
                case Content.ShapeBasics.LeftBracket:
                    return PP.Geometry.LeftBracket;
                case Content.ShapeBasics.RightBracket:
                    return PP.Geometry.RightBracket;
                case Content.ShapeBasics.LeftBrace:
                    return PP.Geometry.LeftBrace;
                case Content.ShapeBasics.RightBrace:
                    return PP.Geometry.RightBrace;
                case Content.ShapeBasics.BracketPair:
                    return PP.Geometry.BracketPair;
                case Content.ShapeBasics.BracePair:
                    return PP.Geometry.BracePair;
                case Content.ShapeBasics.Gear6:
                    return PP.Geometry.Gear6;
                case Content.ShapeBasics.Gear9:
                    return PP.Geometry.Gear9;
                case Content.ShapeBasics.Funnel:
                    return PP.Geometry.Funnel;
                case Content.ShapeBasics.CornerTabs:
                    return PP.Geometry.CornerTabs;
                case Content.ShapeBasics.SquareTabs:
                    return PP.Geometry.SquareTabs;
                case Content.ShapeBasics.PlaqueTabs:
                    return PP.Geometry.PlaqueTabs;
                case Content.ShapeBasics.ChartX:
                    return PP.Geometry.ChartX;
                case Content.ShapeBasics.ChartStar:
                    return PP.Geometry.ChartStar;
                case Content.ShapeBasics.ChartPlus:
                    return PP.Geometry.ChartPlus;
            }
        }

        #endregion

    }
}
