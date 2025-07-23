using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace PptPlus.Components
{
    public class GH_PP_Prs_Templates : GH_PP_preview
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Prs_Templates class.
        /// </summary>
        public GH_PP_Prs_Templates()
          : base("PowerPoint Presentation Templates", "Ppt Template",
              "Get default PowerPoint Presentation Templates",
              Constants.ShortName, Constants.SubPresentation)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.senary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Presentation.Name, Constants.Presentation.NickName, Constants.Presentation.Output, GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Slide.Name, Constants.Slide.NickName, Constants.Slide.Outputs, GH_ParamAccess.list);
            pManager.AddRectangleParameter("Slide Boundary", "B", "The Page Boundary. (Units are represented in Points)", GH_ParamAccess.item);

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

            DA.SetDataList(0, presentation.GetTemplateSlides());
            DA.SetData(1, presentation.Page.Boundary);
            presentation.ClearSlides();

            this.AddRectangles(presentation.Page.Boundary);
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
                return Properties.Resources.Pt_Prs_Template;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("2A198862-A554-4D75-A670-5D85A5C31265"); }
        }
    }
}