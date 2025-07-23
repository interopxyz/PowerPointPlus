using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace PptPlus.Components
{
    public class GH_PP_Con_Create_Txt_Text : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Con_Add_Text class.
        /// </summary>
        public GH_PP_Con_Create_Txt_Text()
          : base("Powerpoint Text Contents", "Ppt Txt",
              "Construct a Powerpoint Text Content Object",
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
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddRectangleParameter(Constants.Boundary.Name, Constants.Boundary.NickName, Constants.Boundary.Input, GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddIntegerParameter("Formatting", "F", "Optional predefined text formatting", GH_ParamAccess.item, 0);
            pManager[3].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[3];
            foreach (Font.Presets value in Enum.GetValues(typeof(Font.Presets)))
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
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.item);
            pManager.AddRectangleParameter(Constants.Boundary.Name, Constants.Boundary.NickName, Constants.Boundary.Input, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            Content content = null;
            if(DA.GetData(0, ref content))content = new Content(content);

            IGH_Goo gooA = null;
            bool hasText = DA.GetData(1, ref gooA);

            Paragraph paragraph = new Paragraph();
            if (hasText)
            {
                if (!gooA.TryGetParagraph(out paragraph))
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Txt input must be a Text Fragment Object or a string");
                    return;
                }
            }
                    content = Content.CreateTextContent(paragraph, content);

            int preset = 0;
            if (DA.GetData(3, ref preset))
            {
                if (preset > 0)
                {
                    paragraph.Font = paragraph.Font.ApplyPreset((Font.Presets)preset);
                    foreach (Fragment fragment in content.Text.Fragments) fragment.Font = fragment.Font.ApplyPreset((Font.Presets)preset);
                }
            }

            Rectangle3d boundary = new Rectangle3d();
            if (DA.GetData(2, ref boundary)) content.Boundary = boundary;

            if (content != null) { 
            DA.SetData(0, content);
            DA.SetData(1, content.Text);
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
                return Properties.Resources.Ppt_Con_Text;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("62097e0c-372d-437f-9e6b-e34056581d1b"); }
        }
    }
}