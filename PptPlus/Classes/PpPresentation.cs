using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;
using Rg = Rhino.Geometry;

using PP = ShapeCrawler;
using System.IO;

namespace PptPlus
{
    public class PpPresentation:PpBase
    {

        #region members

        public enum Units { Inches, Centimeters, Millimeters, Points, Pixels };

        protected PP.IPresentation PresenationObject = new PP.Presentation();
        protected List<PpSlide> slides = new List<PpSlide>();

        protected Page page = new Page();

        protected Dictionary<string, Sd.Color> themeColors = new Dictionary<string, Sd.Color>();
        protected Dictionary<string, string> themeFonts = new Dictionary<string, string>();

        #endregion

        #region constructors

        public PpPresentation():base()
        {

        }

        public PpPresentation(PpPresentation presentation) : base(presentation)
        {
            this.slides = presentation.slides.Duplicate();
            this.themeColors = presentation.themeColors.Duplicate();
            this.themeFonts = presentation.themeFonts.Duplicate();
            this.Page = new Page(presentation.Page);
        }

        public PpPresentation(Page page)
        {
            this.Page = page;
        }

        #endregion

        #region properties

        public virtual Page Page
        {
            get { return this.page; }
            set { this.page = new Page(value); }
        }

        public virtual Sd.Color Accent1
        {
            set
            {
                if (!this.themeColors.ContainsKey("Accent1")) this.themeColors.Add("Accent1", value);
                this.themeColors["Accent1"] = value;
            }
        }
        public virtual Sd.Color Accent2
        {
            set
            {
                if (!this.themeColors.ContainsKey("Accent2")) this.themeColors.Add("Accent2", value);
                this.themeColors["Accent2"] = value;
            }
        }
        public virtual Sd.Color Accent3
        {
            set
            {
                if (!this.themeColors.ContainsKey("Accent3")) this.themeColors.Add("Accent3", value);
                this.themeColors["Accent3"] = value;
            }
        }
        public virtual Sd.Color Accent4
        {
            set
            {
                if (!this.themeColors.ContainsKey("Accent4")) this.themeColors.Add("Accent4", value);
                this.themeColors["Accent4"] = value;
            }
        }
        public virtual Sd.Color Accent5
        {
            set
            {
                if (!this.themeColors.ContainsKey("Accent5")) this.themeColors.Add("Accent5", value);
                this.themeColors["Accent5"] = value;
            }
        }
        public virtual Sd.Color Accent6
        {
            set
            {
                if (!this.themeColors.ContainsKey("Accent6")) this.themeColors.Add("Accent6", value);
                this.themeColors["Accent6"] = value;
            }
        }
        public virtual Sd.Color Light1
        {
            set
            {
                if (!this.themeColors.ContainsKey("Light1")) this.themeColors.Add("Light1", value);
                this.themeColors["Light1"] = value;
            }
        }
        public virtual Sd.Color Light2
        {
            set
            {
                if (!this.themeColors.ContainsKey("Light2")) this.themeColors.Add("Light2", value);
                this.themeColors["Light2"] = value;
            }
        }
        public virtual Sd.Color Dark1
        {
            set
            {
                if (!this.themeColors.ContainsKey("Dark1")) this.themeColors.Add("Dark1", value);
                this.themeColors["Dark1"] = value;
            }
        }
        public virtual Sd.Color Dark2
        {
            set
            {
                if (!this.themeColors.ContainsKey("Dark2")) this.themeColors.Add("Dark2", value);
                this.themeColors["Dark2"] = value;
            }
        }
        public virtual Sd.Color Hyperlink
        {
            set
            {
                if (!this.themeColors.ContainsKey("Hyperlink")) this.themeColors.Add("Hyperlink", value);
                this.themeColors["Hyperlink"] = value;
            }
        }
        public virtual Sd.Color FollowedHyperlink
        {
            set
            {
                if (!this.themeColors.ContainsKey("FollowedHyperlink")) this.themeColors.Add("FollowedHyperlink", value);
                this.themeColors["FollowedHyperlink"] = value;
            }
        }

        public virtual string HeaderFont
        {
            set
            {
                if (!this.themeFonts.ContainsKey("Header")) this.themeFonts.Add("Header", value);
                this.themeFonts["Header"] = value;
            }
        }
        public virtual string BodyFont
        {
            set
            {
                if (!this.themeFonts.ContainsKey("Body")) this.themeFonts.Add("Body", value);
                this.themeFonts["Body"] = value;
            }
        }

        #endregion

        #region methods

        public void AddSlide(PpSlide slide)
        {
            this.slides.Add(new PpSlide(slide));
        }

        public List<PpSlide> GetSlides()
        {
            return this.slides.Duplicate();
        }

        public void ClearSlides()
        {
            this.slides.Clear();
        }

        public void Render()
        {
            this.PreparePresentation();
            this.PresenationObject.Slides[0].Remove();

            foreach (PpSlide slide in this.slides)
            {
                slide.Render(this.PresenationObject);
            }

        }

        private void PreparePresentation()
        {
            this.PresenationObject = new PP.Presentation();
            this.UpdateThemeColors();

            if (this.page.Orientation == Page.Orientations.Portrait)
            {
                this.PresenationObject.SlideWidth = this.page.Width;
                this.PresenationObject.SlideHeight = this.page.Height;
            }
            else
            {
                this.PresenationObject.SlideWidth = this.page.Height;
                this.PresenationObject.SlideHeight = this.page.Width;
            }
        }

        protected void UpdateThemeColors()
        {
            for (int i = 0; i < this.PresenationObject.SlideMasters.Count(); i++)
            {
                if (this.themeColors.ContainsKey("Accent1")) this.PresenationObject.SlideMasters[i].Theme.ColorScheme.Accent1 = this.themeColors["Accent1"].ToHex();
                if (this.themeColors.ContainsKey("Accent2")) this.PresenationObject.SlideMasters[i].Theme.ColorScheme.Accent2 = this.themeColors["Accent2"].ToHex();
                if (this.themeColors.ContainsKey("Accent3")) this.PresenationObject.SlideMasters[i].Theme.ColorScheme.Accent3 = this.themeColors["Accent3"].ToHex();
                if (this.themeColors.ContainsKey("Accent4")) this.PresenationObject.SlideMasters[i].Theme.ColorScheme.Accent4 = this.themeColors["Accent4"].ToHex();
                if (this.themeColors.ContainsKey("Accent5")) this.PresenationObject.SlideMasters[i].Theme.ColorScheme.Accent5 = this.themeColors["Accent5"].ToHex();
                if (this.themeColors.ContainsKey("Accent6")) this.PresenationObject.SlideMasters[i].Theme.ColorScheme.Accent6 = this.themeColors["Accent6"].ToHex();

                if (this.themeColors.ContainsKey("Light1")) this.PresenationObject.SlideMasters[i].Theme.ColorScheme.Light1 = this.themeColors["Light1"].ToHex();
                if (this.themeColors.ContainsKey("Light2")) this.PresenationObject.SlideMasters[i].Theme.ColorScheme.Light2 = this.themeColors["Light2"].ToHex();

                if (this.themeColors.ContainsKey("Dark1")) this.PresenationObject.SlideMasters[i].Theme.ColorScheme.Dark1 = this.themeColors["Dark1"].ToHex();
                if (this.themeColors.ContainsKey("Dark2")) this.PresenationObject.SlideMasters[i].Theme.ColorScheme.Dark2 = this.themeColors["Dark2"].ToHex();

                if (this.themeColors.ContainsKey("Hyperlink")) this.PresenationObject.SlideMasters[i].Theme.ColorScheme.Hyperlink = this.themeColors["Hyperlink"].ToHex();
                if (this.themeColors.ContainsKey("FollowedHyperlink")) this.PresenationObject.SlideMasters[i].Theme.ColorScheme.FollowedHyperlink = this.themeColors["FollowedHyperlink"].ToHex();

                if (this.themeFonts.ContainsKey("Header")) this.PresenationObject.SlideMasters[i].Theme.FontScheme.HeadLatinFont = this.themeFonts["Header"];
                if (this.themeFonts.ContainsKey("Body")) this.PresenationObject.SlideMasters[i].Theme.FontScheme.BodyLatinFont = this.themeFonts["Body"];
            }
        }

        public void Save(string filepath)
        {
            this.Render();
            this.PresenationObject.Save(filepath);
        }

        public void Load(string filepath)
        {

        }

        public List<PpSlide> GetTemplateSlides()
        {
            this.PreparePresentation();
            List<PpSlide> output = new List<PpSlide>();
            foreach (PP.ISlideLayout layout in this.PresenationObject.SlideMasters[0].SlideLayouts)
            {

                PpSlide slide = new PpSlide(layout, (double)this.PresenationObject.SlideHeight);
                output.Add(slide);
            }

            return output;
        }

        public string Stream()
        {
            this.Render();
            Stream stream = new MemoryStream();
            this.PresenationObject.Save(stream);
            stream.Position = 0;
            var buffer = new byte[stream.Length];
            stream.Read(buffer, 0, (int)stream.Length);
            string output = Encoding.Default.GetString(buffer);
            stream.Dispose();
            return output;
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            return "PPT | Presentation " + "{w:" + this.Page.Width + " h:" + this.Page.Height + " s:" + this.slides.Count + "}";
        }

        #endregion

    }
}
