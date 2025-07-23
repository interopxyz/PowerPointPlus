using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;
using Rg = Rhino.Geometry;

using PP = ShapeCrawler;

namespace PptPlus
{
    public class PpSlide:PpBase
    {

        #region members

        protected List<Content> contents = new List<Content>();
        protected PpLayout layout = new PpLayout();

        protected int index = 7;
        protected string name = "Blank";

        #endregion

        #region constructors

        public PpSlide():base()
        {

        }
        public PpSlide(PP.ISlideLayout slide, double height) : base()
        {
            this.index = slide.Number;
            this.name = slide.Name;

            int i = 0;
            foreach (PP.IShape shape in slide.Shapes)
            {
                switch (shape.PlaceholderType)
                {
                    case PP.PlaceholderType.SmartArt:
                    case PP.PlaceholderType.DateAndTime:
                    case PP.PlaceholderType.SlideNumber:
                    case PP.PlaceholderType.Footer:
                        break;
                    default:
                        Content content = new Content(shape, i, height);
                        this.AddContent(content);
                        i++;
                        break;
                }
            }
        }

        public PpSlide(PpSlide slide) : base(slide)
        {
            this.Layout = new PpLayout(slide.layout);
            this.contents = slide.contents.Duplicate();
            this.index = slide.index;
            this.name = slide.name;
        }

        public PpSlide(Page page) : base()
        {
            this.layout = new PpLayout(page);
        }

        public PpSlide(PpLayout layout) : base()
        {
            this.layout = new PpLayout(layout);
        }

        #endregion

        #region properties

        public virtual PpLayout Layout
        {
            get { return this.layout; }
            set { this.layout = new PpLayout(value); }
        }

        public virtual List<Rg.Rectangle3d> Boundaries
        {
            get 
            { 
                List<Rg.Rectangle3d> output = new List<Rg.Rectangle3d>();
                foreach (Content content in this.contents) output.Add(content.Boundary);
                return output;
            }
        }

        #endregion

        #region methods

        public void AddContent(Content content)
        {
            this.contents.Add(new Content(content));
        }

        public List<Content> GetContents()
        {
            return this.contents.Duplicate();
        }

        public void ClearContents()
        {
            this.contents.Clear();
        }

        public void Render(PP.IPresentation presentation)
        {
            presentation.Slides.Add(this.index);
            PP.ISlide slide = presentation.Slides[presentation.Slides.Count - 1];
            for (int i = 0; i < slide.Shapes.Count; i++) slide.Shapes[i].TextBox.SetText("");
            foreach (Content content in this.contents) content.Render(slide,this);
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            return "PPT | Slide("+ this.index + ")"+ this.name + " " + "{c:" + this.contents.Count + "}";
        }

        #endregion

    }
}
