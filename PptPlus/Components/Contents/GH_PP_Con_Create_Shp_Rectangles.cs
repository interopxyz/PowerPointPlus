using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

using PP = ShapeCrawler;

namespace PptPlus.Components.Contents
{
    public class GH_PP_Con_Create_Shp_Rectangles : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Con_Create_Shape_Rectangles class.
        /// </summary>
        public GH_PP_Con_Create_Shp_Rectangles()
          : base("Powerpoint Rectangle Shape Content", "Ppt Rect Shp",
              "Construct a PowerPoint Rect Shape Content Object",
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
            pManager.AddIntegerParameter("Shape", "S", "The Shape Geometry", GH_ParamAccess.item, 4);
            pManager[1].Optional = true;
            pManager.AddRectangleParameter(Constants.Boundary.Name, Constants.Boundary.NickName, Constants.Boundary.Input, GH_ParamAccess.item);
            pManager[2].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[1];
            foreach (Content.ShapeRectangles value in Enum.GetValues(typeof(Content.ShapeRectangles)))
            {
                paramA.AddNamedValue((int)value + " | " + value.ToString(), (int)value);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Output, GH_ParamAccess.item);
            pManager.AddTextParameter("Name", "N", "The name of the shape", GH_ParamAccess.item);
            pManager.AddRectangleParameter(Constants.Boundary.Name, Constants.Boundary.NickName, Constants.Boundary.Input, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            Content content = null;
            if (DA.GetData(0, ref content)) content = new Content(content);

            int shape = 4;
            bool hasShape = DA.GetData(1, ref shape);

            Content.ShapeRectangles shp = (Content.ShapeRectangles)shape;
            if(hasShape) content = Content.CreateShapeContent(shp, content);

            Rectangle3d boundary = new Rectangle3d();
            if (DA.GetData(2, ref boundary)) content.Boundary = boundary;

            if (content!=null) { 
            DA.SetData(0, content);
                DA.SetData(1, content.Geometry.ToString());
                DA.SetData(2, content.Boundary);
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
                return Properties.Resources.Ppt_Con_Shape_Rect;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("CF573F0D-84DD-4417-88B6-8D92A102C178"); }
        }
    }
}