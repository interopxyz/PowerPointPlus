using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace PptPlus.Components.Presentation
{
    public class GH_PP_Prs_Theme_Font : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Prs_Theme_Font class.
        /// </summary>
        public GH_PP_Prs_Theme_Font()
          : base("PowerPoint Theme Font", "Ppt Theme Fnt",
              "Set the theme fonts of a PowerPoint Presentation",
              Constants.ShortName, Constants.SubPresentation)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Presentation.Name, Constants.Presentation.NickName, Constants.Presentation.Output, GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddTextParameter("Body Font", "B", "Optional Presentation Theme Body Font", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddTextParameter("Heading Font", "H", "Optional Presentation Theme Header Font", GH_ParamAccess.item);
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Presentation.Name, Constants.Presentation.NickName, Constants.Presentation.Output, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            PpPresentation presentation = new PpPresentation();
            IGH_Goo gooA = null;
            if (DA.GetData(0, ref gooA))
            {
                if (!gooA.CastTo<PpPresentation>(out presentation))
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Prs input must be a PowerPoint Presentation Object");
                    return;
                }
                presentation = new PpPresentation(presentation);
            }

            string font = string.Empty;
            if (DA.GetData(1, ref font)) presentation.BodyFont = font;
            if (DA.GetData(2, ref font)) presentation.HeaderFont = font;

            DA.SetData(0, presentation);
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return Properties.Resources.Pt_Prs_Theme_Font;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("57A2DD13-13C9-489E-A911-31473DF59B21"); }
        }
    }
}