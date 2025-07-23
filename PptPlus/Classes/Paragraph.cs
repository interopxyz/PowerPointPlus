using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using PP = ShapeCrawler;

namespace PptPlus
{
    public class Paragraph : PpBase
    {
        #region methods

        protected List<Fragment> fragments = new List<Fragment>();
        protected double lineSpacing = 1.0;
        public enum BulletPoints { None, Bar, Bullet, HollowBullet, Box, HollowBox, Star, Arrow, Check, Character, Number };
        public BulletPoints BulletPoint = BulletPoints.None;
        protected string bulletCharacter = "-";
        protected int indentation = 1;
        protected bool hasFont = false;

        #endregion

        #region constructors

        public Paragraph() : base()
        {

        }

        public Paragraph(Paragraph paragraph) : base(paragraph)
        {
            this.fragments = paragraph.fragments.Duplicate();
            this.lineSpacing = paragraph.lineSpacing;
            this.BulletPoint = paragraph.BulletPoint;
            this.bulletCharacter = paragraph.bulletCharacter;
            this.indentation = paragraph.indentation;
            this.font = new Font(paragraph.font);
            this.graphic = new Graphic(paragraph.graphic);
            this.hasFont = paragraph.hasFont;
        }

        public Paragraph(string text) : base()
        {
            this.fragments.Add(new Fragment(text));
        }

        public Paragraph(Fragment fragment) : base()
        {
            this.font = new Font(fragment.Font);
            this.graphic = new Graphic(fragment.Graphic);
            this.fragments.Add(new Fragment(fragment));
        }

        public Paragraph(List<Fragment> fragments)
        {
            this.fragments = fragments.Duplicate();
        }

        public Paragraph(List<Paragraph> paragraphs)
        {
            foreach (Paragraph paragraph in paragraphs) this.fragments.AddRange(paragraph.Fragments.Duplicate());
        }

        #endregion

        #region properties

        public override Font Font
        {
            get { return this.font; }
            set
            {
                this.font = new Font(value);
                this.hasFont = true;
            }
        }

        public bool HasFont
        {
            get { return this.font.IsModified; }
        }

        public int IndentationLevel
        {
            get { return this.indentation; }
            set { this.indentation = value; }
        }

        public string BulletCharacter
        {
            get { return this.bulletCharacter; }
            set { this.bulletCharacter = value; }
        }

        public double LineSpacing
        {
            get { return this.lineSpacing; }
            set { this.lineSpacing = value; }
        }

        public List<Fragment> Fragments
        {
            get { return fragments; }
            set { fragments = value.Duplicate(); }
        }

        public string Text
        {
            get
            {
                StringBuilder output = new StringBuilder();
                foreach (Fragment fragment in this.fragments) output.Append(fragment.Text);
                return output.ToString();
            }
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            if (this.Text.Length < 16) return "Paragraph(" + this.Fragments.Count + "f){" + this.Text + "}";
                return "Paragraph(" + this.Fragments.Count + "f){"+ this.Text.Substring(0, 15)+"...}";
        }

        #endregion

    }
}
