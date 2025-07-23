using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace PptPlus.Components.Presentation
{
    public class GH_PP_Pag_Construct : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Pag_Construct class.
        /// </summary>
        public GH_PP_Pag_Construct()
          : base("PowerPoint Construct Presentation", "WD Doc",
              "Construct a PowerPoint Presentation from Width and Height",
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
            pManager.AddIntegerParameter("Units", "U", "The Units for the Presentation Width and Height" + Environment.NewLine + "(Note: Pixels are at 96dpi)", GH_ParamAccess.item, 1);
            pManager[1].Optional = true;
            pManager.AddNumberParameter("Width", "W", "The Width of the Presentation Slides in the specified Units.", GH_ParamAccess.item, 8.5);
            pManager[2].Optional = true;
            pManager.AddNumberParameter("Height", "H", "The Height of the Presentation Slides in the specified Units", GH_ParamAccess.item, 11);
            pManager[3].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[1];
            foreach (PpPresentation.Units value in Enum.GetValues(typeof(PpPresentation.Units)))
            {
                paramA.AddNamedValue((int)value + " | " + value.ToString(), (int)value);
            }

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
            Page page = new Page();

            int units = 0;
            DA.GetData(1, ref units);

            double width = 8.5;
            DA.GetData(2, ref width);

            double height = 11;
            DA.GetData(3, ref height);

            switch ((PpPresentation.Units)units)
            {
                default://Points
                    page.Width = (int)width;
                    page.Height = (int)height;
                    break;
                case PpPresentation.Units.Inches:
                    page.Width = (int)width.InchToPoint();
                    page.Height = (int)height.InchToPoint();
                    break;
                case PpPresentation.Units.Centimeters:
                    page.Width = (int)width.CentimeterToPoint();
                    page.Height = (int)height.CentimeterToPoint();
                    break;
                case PpPresentation.Units.Millimeters:
                    page.Width = (int)width.MillimeterToPoint();
                    page.Height = (int)height.MillimeterToPoint();
                    break;
                case PpPresentation.Units.Pixels:
                    page.Width = (int)width.PixelsToPoint();
                    page.Height = (int)height.PixelsToPoint();
                    break;
            }

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
                return Properties.Resources.Ppt_Pag_Size_Custom;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("b4f73be0-133b-40d8-928c-788d34db0581"); }
        }
    }
}