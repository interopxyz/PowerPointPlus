using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

using Sd = System.Drawing;

namespace PptPlus.Components.Presentation
{
    public class GH_PP_Prs_Theme_Accent : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Prs_ThemeColors class.
        /// </summary>
        public GH_PP_Prs_Theme_Accent()
          : base("PowerPoint Theme Primary Colors", "Ppt Theme Clr 1",
              "Set the Primary Theme Colors of a PowerPoint Presentation",
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
            pManager.AddColourParameter("Accent1", "A1", "Optional Accent1 Color", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddColourParameter("Accent2", "A2", "Optional Accent2 Color", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddColourParameter("Accent3", "A3", "Optional Accent3 Color", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddColourParameter("Accent4", "A4", "Optional Accent4 Color", GH_ParamAccess.item);
            pManager[4].Optional = true;
            pManager.AddColourParameter("Accent5", "A5", "Optional Accent5 Color", GH_ParamAccess.item);
            pManager[5].Optional = true;
            pManager.AddColourParameter("Accent6", "A6", "Optional Accent6 Color", GH_ParamAccess.item);
            pManager[6].Optional = true;
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

            Sd.Color color = Sd.Color.Black;
            if (DA.GetData(1, ref color))presentation.Accent1 = color;
            if (DA.GetData(2, ref color)) presentation.Accent2 = color;
            if (DA.GetData(3, ref color)) presentation.Accent3 = color;
            if (DA.GetData(4, ref color)) presentation.Accent4 = color;
            if (DA.GetData(5, ref color)) presentation.Accent5 = color;
            if (DA.GetData(6, ref color)) presentation.Accent6 = color;

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
                return Properties.Resources.Pt_Prs_Theme_Accent_Colors;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("0A9A56D3-FCE6-4367-A18D-C07DFC4CC008"); }
        }
    }
}