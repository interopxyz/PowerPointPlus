using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace PptPlus.Components.Contents
{
    public class GH_PP_Con_Create_Txt_List : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Create_List class.
        /// </summary>
        public GH_PP_Con_Create_Txt_List()
          : base("Powerpoint List Contents", "Ppt List",
              "Construct a Powerpoint Text List Content Object",
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
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.list);
            pManager[1].Optional = true;
            pManager.AddRectangleParameter(Constants.Boundary.Name, Constants.Boundary.NickName, Constants.Boundary.Input, GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddIntegerParameter("Bullet Type", "T", "List Bullet Type", GH_ParamAccess.item, 1);
            pManager[3].Optional = true;
            pManager.AddTextParameter("Bullet Character", "C", "The bullet character if bullet type is set to Character", GH_ParamAccess.item, "-");
            pManager[4].Optional = true;
            pManager.AddIntegerParameter("Bullet Indentation", "I", "List Bullet Indentation", GH_ParamAccess.item, 1);
            pManager[5].Optional = true;
            pManager.AddIntegerParameter("Formatting", "F", "Optional predefined text formatting", GH_ParamAccess.item);
            pManager[6].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[3];
            foreach (Paragraph.BulletPoints value in Enum.GetValues(typeof(Paragraph.BulletPoints)))
            {
                paramA.AddNamedValue((int)value + " | " + value.ToString(), (int)value);
            }

            Param_Integer paramB = (Param_Integer)pManager[6];
            foreach (Font.Presets value in Enum.GetValues(typeof(Font.Presets)))
            {
                paramB.AddNamedValue((int)value + " | " + value.ToString(), (int)value);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Output, GH_ParamAccess.item);
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.list);
            pManager.AddRectangleParameter(Constants.Boundary.Name, Constants.Boundary.NickName, Constants.Boundary.Input, GH_ParamAccess.item);
            pManager.AddIntegerParameter("Bullet Type", "T", "List Bullet Type", GH_ParamAccess.item);
            pManager.AddTextParameter("Bullet Character", "C", "The bullet character if bullet type is set to Character", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Bullet Indentation", "I", "List Bullet Indentation", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            Content content = null;
            if (DA.GetData(0, ref content)) content = new Content(content);

            List<IGH_Goo> goos = new List<IGH_Goo>();
            DA.GetDataList(1, goos);

            int type = 0;
            DA.GetData(3, ref type);

            string character = "";
            DA.GetData(4, ref character);

            int indentation = 0;
            DA.GetData(5, ref indentation);

            int preset = 0;
            bool hasPreset = DA.GetData(6, ref preset);

            List<Paragraph> paragraphs = new List<Paragraph>();
            foreach (IGH_Goo goo in goos)
            {
                if (goo.TryGetParagraph(out Paragraph paragraph))
                {
                    if(hasPreset)foreach (Fragment fragment in paragraph.Fragments) fragment.Font = Fonts.GetPreset((Font.Presets)preset);
                    paragraph.BulletPoint = (Paragraph.BulletPoints)type;
                    paragraph.BulletCharacter = character;
                    paragraph.IndentationLevel = indentation;
                    paragraphs.Add(paragraph);
                }
            }

            content = Content.CreateListContent(paragraphs, content);

            Rectangle3d boundary = new Rectangle3d();
            if (DA.GetData(2, ref boundary)) content.Boundary = boundary;

            if (content != null)
            {
                DA.SetData(0, content);
                DA.SetData(2, content.Boundary);

                if (content.Values.Count > 0)
                {
                    DA.SetDataList(1, content.Values[0]);
                    if (content.Values[0].Count > 0)
                    {
                        DA.SetData(3, content.Values[0][0].BulletPoint);
                        DA.SetData(4, content.Values[0][0].BulletCharacter);
                        DA.SetData(5, content.Values[0][0].IndentationLevel);
                    }
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
                return Properties.Resources.Ppt_Con_List;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("B203F89B-3B71-46DF-A69C-72088747B2E4"); }
        }
    }
}