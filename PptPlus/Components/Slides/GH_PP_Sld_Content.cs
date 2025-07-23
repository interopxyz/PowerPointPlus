using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace PptPlus.Components.Slides
{
    public class GH_PP_Sld_Content : GH_PP_preview
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Sld_AddContent class.
        /// </summary>
        public GH_PP_Sld_Content()
          : base("PowerPoint Slide Contents", "Ppt Slide Content",
              "Add Contents to a PowerPoint Slide",
              Constants.ShortName, Constants.SubSlide)
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Slide.Name, Constants.Slide.NickName, Constants.Slide.Input, GH_ParamAccess.item);
            pManager[0].Optional = true;
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Inputs, GH_ParamAccess.list);
            pManager[1].Optional = true;
            pManager.AddBooleanParameter("Clear", "X", "Optionally clear the Contents of the Slide", GH_ParamAccess.item, false);
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Slide.Name, Constants.Slide.NickName, Constants.Slide.Output, GH_ParamAccess.item);
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Outputs, GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            PpSlide slide = new PpSlide();
            IGH_Goo gooA = null;
            if (DA.GetData(0, ref gooA))
            {
                if (!gooA.CastTo<PpSlide>(out slide))
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Sld input must be a PowerPoint Slide Object");
                    return;
                }
                slide = new PpSlide(slide);
            }

            bool clear = false;
            DA.GetData(2, ref clear);
            if (clear) slide.ClearContents();

            List<IGH_Goo> goos = new List<IGH_Goo>();
            DA.GetDataList(1, goos);
            foreach (IGH_Goo goo in goos) if (goo.TryGetContent(out Content content)) slide.AddContent(content);

            DA.SetData(0, slide);
            DA.SetDataList(1, slide.GetContents());
            this.AddRectangles(slide.Boundaries);
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
                return Properties.Resources.Ppt_Slide_AddContents;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("c27ced0c-fd03-4343-bfd1-94becfe7e9c2"); }
        }
    }
}