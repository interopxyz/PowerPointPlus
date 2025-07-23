using System;
using System.Collections.Generic;
using System.Linq;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace PptPlus.Components.Contents
{
    public class GH_PP_Con_Create_Cht_Bar : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Create_Chart_Bar class.
        /// </summary>
        public GH_PP_Con_Create_Cht_Bar()
          : base("Powerpoint Bar Chart Contents", "Ppt Bar Chart",
              "Construct a Powerpoint Bar Chart Content Object",
              Constants.ShortName, Constants.SubContent)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.quarternary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, "_" + Constants.Content.NickName, "Optional " + Constants.Content.Input, GH_ParamAccess.item);
            pManager[0].Optional = true;
            pManager.AddNumberParameter("Numeric Values", "V", "The numeric values to plot", GH_ParamAccess.list);
            pManager[1].Optional = true;
            pManager.AddGenericParameter("Text Labels", "L", "Corresponding labels to numeric values", GH_ParamAccess.list);
            pManager[2].Optional = true;
            pManager.AddRectangleParameter(Constants.Boundary.Name, Constants.Boundary.NickName, Constants.Boundary.Input, GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddTextParameter("Series Name", "S", "Optional Series Name", GH_ParamAccess.item, "Unnamed");
            pManager[4].Optional = true;

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Output, GH_ParamAccess.item);
            pManager.AddNumberParameter("Numeric Values", "V", "The numeric values to plot", GH_ParamAccess.list);
            pManager.AddGenericParameter("Text Labels", "L", "Corresponding labels to numeric values", GH_ParamAccess.list);
            pManager.AddRectangleParameter(Constants.Boundary.Name, Constants.Boundary.NickName, Constants.Boundary.Input, GH_ParamAccess.item);
            pManager.AddTextParameter("Series Name", "S", "Optional Series Name", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            Content content = null;
            if (DA.GetData(0, ref content)) content = new Content(content);

            List<double> values = new List<double>();
            bool hasValues = DA.GetDataList(1, values);

            List<string> labels = new List<string>();
            DA.GetDataList(2, labels);

            if (labels.Count != labels.Distinct().Count())
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Labels must be unqiue");
                return;
            }

            string name = "Unnamed";
            DA.GetData(4, ref name);

            if (hasValues) content = Content.CreateChartBarContent(values, labels, name, content);

            Rectangle3d boundary = new Rectangle3d();
            if (DA.GetData(3, ref boundary)) content.Boundary = boundary;

            if (content != null)
            {
                DA.SetData(0, content);
                DA.SetData(3, content.Boundary);
                DA.SetData(4, content.SeriesName);
                if (content.CategoryLabels.Count > 0)
                {
                    DA.SetDataList(1, content.CategoryValues);
                    DA.SetDataList(2, content.CategoryLabels);
                }
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
                return Properties.Resources.Ppt_Con_Chart_Bar;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("0ABA8EFB-416B-4C79-BF25-570D4D853687"); }
        }
    }
}