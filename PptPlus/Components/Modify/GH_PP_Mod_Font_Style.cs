using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace PptPlus.Components.Modify
{
    public class GH_PP_Mod_Font_Style : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Mod_Font_Style class.
        /// </summary>
        public GH_PP_Mod_Font_Style()
          : base("PowerPoint Font Style", "Ppt Font Style",
              "Modify Powerpoint Font Style if applicable",
              Constants.ShortName, Constants.SubModify)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.tertiary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddBooleanParameter("Bold", "B", "Optionally set the Bold status of the text", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddBooleanParameter("Italic", "I", "Optionally set the Italic status of the text", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddBooleanParameter("Underlined", "U", "Optionally set the Underlined status of the text", GH_ParamAccess.item);
            pManager[3].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Output, GH_ParamAccess.item);
            pManager.AddBooleanParameter("Bold", "B", "Get the Bold status of the text", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Italic", "I", "Get the Italic status of the text", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Underlined", "U", "Get the Underlined status of the text", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            Paragraph paragraph = null;
            bool isParagraph = false;
            Fragment fragment = null;

            if (gooA.TryCastToParagraph(out paragraph))
            {
                isParagraph = true;
            }
            else if (!gooA.TryGetFragment(out fragment))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Txt input must be a Text Paragraph Object, Text Fragment Object or a string");
                return;
            }

            bool bold = false;
            if (DA.GetData(1, ref bold))
            {
                if (isParagraph)
                {
                    paragraph.Font.Bold = bold;
                    foreach (Fragment f in paragraph.Fragments) f.Font.Bold = bold;
                }
                else
                {
                    fragment.Font.Bold = bold;
                }
            }

            bool italic = false;
            if (DA.GetData(2, ref italic))

            {
                if (isParagraph)
                {
                    paragraph.Font.Italic = italic;
                    foreach (Fragment f in paragraph.Fragments) f.Font.Italic = italic;
                }
                else
                {
                    fragment.Font.Italic = italic;
                }
            }

            bool underlined = false;
            if (DA.GetData(3, ref underlined))
            {
                if (isParagraph)
                {
                    paragraph.Font.Underlined = underlined;
                    foreach (Fragment f in paragraph.Fragments) f.Font.Underlined = underlined;
                }
                else
                {
                    fragment.Font.Underlined = underlined;
                }
            }

            if (isParagraph)
            {
                DA.SetData(0, paragraph);
                DA.SetData(1, paragraph.Font.Bold);
                DA.SetData(2, paragraph.Font.Italic);
                DA.SetData(3, paragraph.Font.Underlined);
            }
            else
            {
                DA.SetData(0, fragment);
                DA.SetData(1, fragment.Font.Bold);
                DA.SetData(2, fragment.Font.Italic);
                DA.SetData(3, fragment.Font.Underlined);
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
                return Properties.Resources.Ppt_Mod_Font_Style;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("E608104B-7647-4C2C-9A8D-319670D2037A"); }
        }
    }
}