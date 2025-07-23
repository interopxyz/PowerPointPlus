using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace PptPlus.Components.Contents
{
    public class GH_PP_Con_Create_Txt_Table : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Con_Create_Table class.
        /// </summary>
        public GH_PP_Con_Create_Txt_Table()
          : base("Powerpoint Table Contents", "Ppt Tbl",
              "Construct a Powerpoint Table Content Object",
              Constants.ShortName, Constants.SubContent)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.primary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, "_" + Constants.Content.NickName, "Optional " + Constants.Content.Input, GH_ParamAccess.item);
            pManager[0].Optional = true;
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.tree);
            pManager[1].Optional = true;
            pManager.AddRectangleParameter(Constants.Boundary.Name, Constants.Boundary.NickName, Constants.Boundary.Input, GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddIntegerParameter("Style", "S", "Optional table style", GH_ParamAccess.item, 0);
            pManager[3].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[3];
            foreach (Content.TableStyles value in Enum.GetValues(typeof(Content.TableStyles)))
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
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.tree);
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

            List<List<Paragraph>> dataSet = new List<List<Paragraph>>();
            DA.GetDataTree(1, out GH_Structure<IGH_Goo> gooSet);

            foreach (List<IGH_Goo> goos in gooSet.Branches)
            {
                List<Paragraph> paragraphs = new List<Paragraph>();
                foreach (IGH_Goo goo in goos)
                {
                    if (goo.TryGetParagraph(out Paragraph paragraph))
                    {
                        paragraphs.Add(paragraph);
                    }
                }
                dataSet.Add(paragraphs);
            }

            int style = 0;
            DA.GetData(3, ref style);

            content = Content.CreateTableContent(dataSet,(Content.TableStyles)style, content);

            Rectangle3d boundary = new Rectangle3d();
            if (DA.GetData(2, ref boundary)) content.Boundary = boundary;

            GH_Path path = new GH_Path();
            if (this.Params.Input[0].VolatileData.PathCount > 1) path = this.Params.Input[0].VolatileData.get_Path(this.RunCount - 1);
            path = path.AppendElement(this.RunCount - 1);

            if (content != null)
            {
                GH_Structure<GH_ObjectWrapper> data = content.Values.ToDataTree(path);
                DA.SetData(0, content);
                DA.SetDataTree(1, data);
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
                return Properties.Resources.Ppt_Con_Table;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("E7719323-FAB8-4CB2-96D0-77127CF8F4CC"); }
        }
    }
}