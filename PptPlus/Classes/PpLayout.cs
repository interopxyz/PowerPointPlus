using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PptPlus
{
    public class PpLayout: PpBase
    {

        #region members

        protected Page page = Page.R4x3();
        List<Content> contents = new List<Content>();
        protected int index = 0;

        #endregion

        #region constructors

        public PpLayout():base()
        {

        }

        public PpLayout(PpLayout layout):base(layout)
        {
            this.Page = layout.page;
            this.contents = layout.contents.Duplicate();
            this.index = layout.index;
        }

        public PpLayout(Page page) : base()
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

        #endregion

        #region methods



        #endregion

        #region overrides

        public override string ToString()
        {
            return base.ToString();
        }

        #endregion

    }
}
