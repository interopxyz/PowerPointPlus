using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;

namespace PptPlus.Components.Contents
{
    public class GH_PP_Con_Create_Line : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Con_Create_Line class.
        /// </summary>
        public GH_PP_Con_Create_Line()
          : base("Powerpoint Line Content", "Ppt Line",
              "Construct a PowerPoint Line Content Object",
              Constants.ShortName, Constants.SubContent)
        {
        }


        public override GH_Exposure Exposure => GH_Exposure.tertiary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, "_" + Constants.Content.NickName, "Optional " + Constants.Content.Input, GH_ParamAccess.item);
            pManager[0].Optional = true;
            pManager.AddLineParameter("Line", "L", "A line object", GH_ParamAccess.item);
            pManager[1].Optional = true;

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Output, GH_ParamAccess.item);
            pManager.AddLineParameter("Line", "L", "A line object", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            Content content = null;
            if (DA.GetData(0, ref content)) content = new Content(content);

            Line line = new Line();
            bool hasLine = DA.GetData(1, ref line);

            if(hasLine) content = Content.CreateLineContent(line, content);

            if (content != null)
            {
                DA.SetData(0, content);
                DA.SetData(1, content.Line);
            }
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
                return Properties.Resources.Ppt_Con_Shape_LineA;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("F86C9B7E-1393-477F-9F95-708AD69051AD"); }
        }
    }
}