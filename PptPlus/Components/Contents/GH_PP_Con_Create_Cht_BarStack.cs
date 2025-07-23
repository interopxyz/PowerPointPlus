using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PptPlus.Components.Contents
{
    public class GH_PP_Con_Create_Cht_BarStack : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Con_Create_Chart_BarStack class.
        /// </summary>
        public GH_PP_Con_Create_Cht_BarStack()
          : base("Powerpoint Stacked Bar Chart Contents", "Ppt Stack Bar Chart",
              "Construct a Powerpoint Stacked Bar Chart Content Object",
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
            pManager.AddNumberParameter("Numeric Values", "V", "The numeric values to plot", GH_ParamAccess.tree);
            pManager[1].Optional = true;
            pManager.AddGenericParameter("Text Labels", "L", "Corresponding labels to numeric values", GH_ParamAccess.list);
            pManager[2].Optional = true;
            pManager.AddRectangleParameter(Constants.Boundary.Name, Constants.Boundary.NickName, Constants.Boundary.Input, GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddTextParameter("Series Name", "S", "Optional Series Name", GH_ParamAccess.list);
            pManager[4].Optional = true;

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Output, GH_ParamAccess.item);
            pManager.AddNumberParameter("Numeric Values", "V", "The numeric values to plot", GH_ParamAccess.tree);
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

            List<List<double>> dataSet = new List<List<double>>();
            bool hasValues = DA.GetDataTree(1, out GH_Structure<GH_Number> gooSet);

            foreach (List<GH_Number> goos in gooSet.Branches)
            {
                List<double> values = new List<double>();
                foreach (GH_Number goo in goos)
                {
                        values.Add(goo.Value);
                }
                dataSet.Add(values);
            }

            List<string> labels = new List<string>();
            DA.GetDataList(2, labels);

            if (labels.Count != labels.Distinct().Count())
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Labels must be unqiue");
                return;
            }

            List<string> series = new List<string>();
            DA.GetDataList(4, series);

            if (hasValues) content = Content.CreateChartStackedBarContent(dataSet, labels, series, content);

            Rectangle3d boundary = new Rectangle3d();
            if (DA.GetData(3, ref boundary)) content.Boundary = boundary;

            if (content != null)
            {
                DA.SetData(0, content);
                DA.SetData(3, content.Boundary);
                DA.SetData(4, content.SeriesName);
                if (content.CategoryLabels.Count > 0)
                {
                    GH_Path path = new GH_Path();
                    if (this.Params.Input[0].VolatileData.PathCount > 1) path = this.Params.Input[0].VolatileData.get_Path(this.RunCount - 1);
                    path = path.AppendElement(this.RunCount - 1);

                    GH_Structure<GH_Number> data = content.CategoryListValues.ToDataTree(path);
                    DA.SetDataTree(1, data);
                    DA.SetDataList(2, content.CategoryListLabels);
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
                return Properties.Resources.Ppt_Con_Chart_Bar_Stacked;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("F27A1684-9D18-42B3-B352-04DD4EDFDAE4"); }
        }
    }
}