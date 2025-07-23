using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

namespace PptPlus
{
    public class Graphic
    {

        #region members

        protected Sd.Color fill = Sd.Color.Transparent;
        protected bool hasFill = false;

        protected Sd.Color stroke = Sd.Color.Black;
        protected double weight = 1.0;
        protected bool hasStroke = false;

        protected bool topBorder = false;
        protected bool bottomBorder = false;
        protected bool leftBorder = false;
        protected bool rightBorder = false;

        #endregion

        #region constructor

        public Graphic()
        {

        }

        public Graphic(Graphic graphic)
        {
            this.fill = graphic.fill;
            this.hasFill = graphic.hasFill;

            this.topBorder = graphic.topBorder;
            this.bottomBorder = graphic.bottomBorder;
            this.leftBorder = graphic.leftBorder;
            this.rightBorder = graphic.rightBorder;

            this.weight = graphic.weight;
            this.stroke = graphic.stroke;
            this.hasStroke = graphic.hasStroke;
        }

        public Graphic(Sd.Color stroke, double weight)
        {
            this.Stroke = stroke;
            this.Weight = weight;
            this.hasFill = false;
        }

        public Graphic(Sd.Color stroke, Sd.Color fillColor)
        {
            this.Stroke = stroke;
            this.Fill = fillColor;
        }

        public Graphic(Sd.Color color)
        {
            this.Fill = color;
            this.hasStroke = false;
        }

        #endregion

        #region properties

        public virtual Sd.Color Fill
        {
            get { return this.fill; }
            set
            {
                this.fill = value;
                this.hasFill = true;
            }
        }

        public virtual bool HasFill
        {
            get { return hasFill; }
        }

        public virtual Sd.Color Stroke
        {
            get { return stroke; }
            set
            {
                this.stroke = value;
                this.hasStroke = (value !=Sd.Color.Transparent);
            }
        }

        public virtual double Weight
        {
            get { return weight; }
            set
            {
                this.weight = value;
                this.hasStroke = (this.weight>0);
            }
        }

        public virtual bool HasStroke
        {
            get { return hasStroke; }
        }

        public virtual bool NoStroke
        {
            get
            {
                if (this.stroke.A == 0) return true;
                if (this.weight <= 0) return true;
                return false;
            }
        }

        public virtual bool NoFill
        {
            get
            {
                if (this.fill.A == 0) return true;
                return false;
            }
        }

        public virtual bool OnlyStroke
        {
            get
            {
                return (this.NoFill && !this.NoStroke);
            }
        }

        public virtual bool OnlyFill
        {
            get
            {
                return (!this.NoFill && this.NoStroke);
            }
        }

        public virtual bool TopBorder
        {
            get { return this.topBorder; }
            set { this.topBorder = value; }
        }

        public virtual bool BottomBorder
        {
            get { return this.bottomBorder; }
            set { this.bottomBorder = value; }
        }

        public virtual bool LeftBorder
        {
            get { return this.leftBorder; }
            set { this.leftBorder = value; }
        }

        public virtual bool RightBorder
        {
            get { return this.rightBorder; }
            set { this.rightBorder = value; }
        }

        public virtual bool HasBorders
        {
            get { return (this.leftBorder & this.rightBorder & this.topBorder & this.bottomBorder); }
            set
            {
                this.leftBorder = value;
                this.rightBorder = value;
                this.topBorder = value;
                this.bottomBorder = value;
            }
        }

        #endregion

        #region methods



        #endregion

    }

    public static class Graphics
    {
        public static Graphic Outline { get { return new Graphic(Sd.Color.Black, 1); } }
        public static Graphic Solid { get { return new Graphic(Sd.Color.Black); } }
        public static Graphic Empty { get { return new Graphic(Sd.Color.Transparent, Sd.Color.Transparent); } }

    }
}
