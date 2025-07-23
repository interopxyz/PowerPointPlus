using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace PptPlus.Components.Presentation
{
    public class GH_PP_Pag_IsoA : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Pag_Paper class.
        /// </summary>
        public GH_PP_Pag_IsoA()
          : base("PowerPoint Iso A Presentation", "Ppt IsoA",
              "Create a PowerPoint Presentation from a standard Iso A Paper Size",
              Constants.ShortName, Constants.SubPresentation)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.primary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddPlaneParameter("Plane", "P", "The origin plane for the slide." + Environment.NewLine + "(Note: This affects the relative location of any content to the slide)", GH_ParamAccess.item, Plane.WorldXY);
            pManager[0].Optional = true;
            pManager.AddIntegerParameter("Size", "S", "A standard Iso A paper size", GH_ParamAccess.item, 0);
            pManager[1].Optional = true;
            pManager.AddIntegerParameter("Orientation", "O", "Set the portrait or landscape orientation of the page", GH_ParamAccess.item, 1);
            pManager[2].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[1];
            foreach (Page.SizesIsoA value in Enum.GetValues(typeof(Page.SizesIsoA)))
            {
                paramA.AddNamedValue((int)value + " | " + value.ToString(), (int)value);
            }

            Param_Integer paramB = (Param_Integer)pManager[2];
            paramB.AddNamedValue("Portrait", 0);
            paramB.AddNamedValue("Landscape", 1);

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Presentation.Name, Constants.Presentation.NickName, Constants.Presentation.Output, GH_ParamAccess.item);
            pManager.AddGenericParameter(Constants.Slide.Name, Constants.Slide.NickName, Constants.Slide.Output, GH_ParamAccess.item);
            pManager.AddRectangleParameter("Boundary", "B", "The Page Boundary. (Units are represented in Points)", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            int type = 0;
            DA.GetData(1, ref type);

            Page page = Page.Preset((Page.SizesIsoA)type);

            int orientation = 0;
            DA.GetData(2, ref orientation);
            if (orientation != 0) page.Orientation = Page.Orientations.Landscape;

            Plane plane = Plane.WorldXY;
            if (DA.GetData(0, ref plane)) page.Plane = plane;

            PpSlide slide = new PpSlide(page);
            PpPresentation presentation = new PpPresentation(page);

            DA.SetData(0, presentation);
            DA.SetData(1, slide);
            DA.SetData(2, page.Boundary);
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
                return Properties.Resources.Ppt_Pag_Size_IsoA;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("cad90a9d-a9e4-48c4-92cc-6cb16dfcbd32"); }
        }
    }
}