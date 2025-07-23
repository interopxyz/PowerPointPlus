using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace PptPlus.Components.Presentation
{
    public class GH_PP_Prs_Slides : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Prs_AddSlides class.
        /// </summary>
        public GH_PP_Prs_Slides()
          : base("PowerPoint Add Slides", "Ppt Save",
              "Add Slides to a PowerPoint Presentation",
              Constants.ShortName, Constants.SubPresentation)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.quarternary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Presentation.Name, Constants.Presentation.NickName, Constants.Presentation.Input, GH_ParamAccess.item);
            pManager[0].Optional = true;
            pManager.AddGenericParameter(Constants.Slide.Name, Constants.Slide.NickName, Constants.Slide.Inputs, GH_ParamAccess.list);
            pManager[1].Optional = true;
            pManager.AddBooleanParameter("Clear", "X", "Optionally clear the Slides from the Presentation", GH_ParamAccess.item, false);
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Presentation.Name, Constants.Presentation.NickName, Constants.Presentation.Output, GH_ParamAccess.item);
            pManager.AddGenericParameter(Constants.Slide.Name, Constants.Slide.NickName, Constants.Slide.Outputs, GH_ParamAccess.list);
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
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Sld input must be a PowerPoint Slide Object");
                    return;
                }
                presentation = new PpPresentation(presentation);
            }

            bool clear = false;
            DA.GetData(2, ref clear);
            if (clear) presentation.ClearSlides();

            List<IGH_Goo> goos = new List<IGH_Goo>();
            if (!DA.GetDataList(1, goos)) return;
            foreach (IGH_Goo goo in goos) if (goo.CastTo<PpSlide>(out PpSlide slide)) presentation.AddSlide(slide);

            DA.SetData(0, presentation);
            DA.SetDataList(1, presentation.GetSlides());
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
                return Properties.Resources.Pt_Prs_AddSlides;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("48af60c5-16c0-45b6-9eb6-c276959b78f0"); }
        }
    }
}