using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;
using Rg = Rhino.Geometry;

namespace PptPlus
{
    public class Constants
    {

        #region naming

        public static string UniqueName
        {
            get { return ShortName + "_" + DateTime.UtcNow.ToString("yyyy-dd-M_HH-mm-ss"); }
        }

        public static string LongName
        {
            get { return ShortName + " v" + Major + "." + Minor; }
        }

        public static string ShortName
        {
            get { return "PowerPointPlus"; }
        }

        private static string Minor
        {
            get { return typeof(Constants).Assembly.GetName().Version.Minor.ToString(); }
        }
        private static string Major
        {
            get { return typeof(Constants).Assembly.GetName().Version.Major.ToString(); }
        }

        public static string SubPresentation
        {
            get { return "Presentation"; }
        }

        public static string SubSlide
        {
            get { return "Slide"; }
        }

        public static string SubContent
        {
            get { return "Content"; }
        }

        public static string SubModify
        {
            get { return "Modify"; }
        }

        #endregion

        #region inputs / outputs

        public static Descriptor Presentation
        {
            get { return new Descriptor("Ppt Presentation", "Prs", "A PowerPoint Presentation", "PowerPoint Presentations", "A PowerPoint Presentation Object", "PowerPoint Presentation Objects"); }
        }

        public static Descriptor Slide
        {
            get { return new Descriptor("Ppt Slide", "Sld", "A PowerPoint Slide Object", "PowerPoint Slide Objects", "A PowerPoint Slide Object", "PowerPoint Slide Objects"); }
        }

        public static Descriptor Content
        {
            get { return new Descriptor("Ppt Content", "Con", "A Slide Content Object, Slide Content Object", "Slide Content Objects, Slide Content Objects", "Slide Content Object", "Slide Content Objects"); }
        }

        public static Descriptor Paragraph
        {
            get { return new Descriptor("Ppt Paragraph", "Par", "A Text Paragraph Object, Text Fragment Object, or String", "Text Paragraph Objects, Text Fragment Objects, or Strings", "A Text Object", "Text Objects"); }
        }

        public static Descriptor Fragment
        {
            get { return new Descriptor("Ppt Text", "Txt", "A Text Fragment Object, or String", "Text Fragment Objects, or Strings", "A Text Fragment Object", "Text Fragment Objects"); }
        }

        public static Descriptor Boundary
        {
            get { return new Descriptor("Boundary Rectangle", "B", "The bounding rectangle of the shape", "The bounding rectangles of the shapes", "The bounding rectangle of the shape", "The bounding rectangles of the shapes"); }
        }

        public static Descriptor Activate
        {
            get { return new Descriptor("Activate", "_A", "If true, the component will be activated", "If true, the component will be activated", "The active status of the component", "The active statuses of the component"); }
        }

        #endregion

        #region geometry

        public static Rg.Rectangle3d DefaultBoundary(double width = 100, double height = 100)
        {
            return new Rg.Rectangle3d(Rg.Plane.WorldXY, width, height);
        }

        #endregion

        #region color

        public static Sd.Color StartColor
        {
            get { return Sd.Color.FromArgb(99, 190, 123); }
        }

        public static Sd.Color MidColor
        {
            get { return Sd.Color.FromArgb(255, 235, 132); }
        }

        public static Sd.Color EndColor
        {
            get { return Sd.Color.FromArgb(248, 105, 107); }
        }

        #endregion
    }

    public class Descriptor
    {
        private string name = string.Empty;
        private string nickname = string.Empty;
        private string input = string.Empty;
        private string inputs = string.Empty;
        private string output = string.Empty;
        private string outputs = string.Empty;

        public Descriptor(string name, string nickname, string input, string inputs, string output, string outputs)
        {
            this.name = name;
            this.nickname = nickname;
            this.input = input;
            this.inputs = inputs;
            this.output = output;
            this.outputs = outputs;
        }

        public virtual string Name
        {
            get { return name; }
        }

        public virtual string NickName
        {
            get { return nickname; }
        }

        public virtual string Input
        {
            get { return input; }
        }

        public virtual string Output
        {
            get { return output; }
        }

        public virtual string Inputs
        {
            get { return inputs; }
        }

        public virtual string Outputs
        {
            get { return outputs; }
        }
    }
}
