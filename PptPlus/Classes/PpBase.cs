using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PptPlus
{
    public abstract class PpBase
    {

        #region members

        protected Graphic graphic = new Graphic();
        protected Font font = new Font();

        #endregion

        #region constructors

        protected PpBase()
        {

        }

        protected PpBase(PpBase ppBase)
        {

            this.Graphic = ppBase.graphic;
            this.Font = ppBase.font;

        }

        #endregion

        #region properties

        public virtual Font Font
        {
            get { return this.font; }
            set { this.font = new Font(value); }
        }

        public virtual Graphic Graphic
        {
            get { return this.graphic; }
            set { this.graphic = new Graphic(value); }
        }

        #endregion

        #region methods



        #endregion

        #region overrides



        #endregion

    }
}
