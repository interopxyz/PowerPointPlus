using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PptPlus
{
    public class Fragment: PpBase
    {

        #region members

        public string Text = string.Empty;
        public string Hyperlink = string.Empty;

        #endregion

        #region constructors

        public Fragment() : base()
        {
            this.font = Fonts.Normal;
            this.graphic = Graphics.Outline;
        }

        public Fragment(Fragment fragment):base(fragment)
        {
            this.Text = fragment.Text;
            this.Hyperlink = fragment.Hyperlink;
            this.font = new Font(fragment.font);
            this.graphic = new Graphic(fragment.graphic);
        }

        public Fragment(string text):base()
        {
            this.font = Fonts.Normal;
            this.graphic = Graphics.Outline;
            this.Text = text;
        }

        public Fragment(string text, string hyperlink) : base()
        {
            this.font = Fonts.Normal;
            this.graphic = Graphics.Outline;
            this.Text = text;
            this.Hyperlink = hyperlink;
        }

        public Fragment(string text, Font font) : base()
        {
            this.font = font;
            this.graphic = Graphics.Outline;
            this.Text = text;
            this.font = new Font(font);
        }

        #endregion

        #region properties

        public override Font Font
        {
            get { return this.font; }
            set { 
                this.font = new Font(value);
            }
        }

        public bool HasFont
        {
            get { return this.font.IsModified; }
        }

        public bool HasLink
        {
            get { return this.Hyperlink != string.Empty; }
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            if (this.Text.Length < 16) return "Fragment {" + this.Text + "}";
            return "Fragment {" + this.Text.Substring(0, 15) + "...}";
        }

        #endregion

    }
}
