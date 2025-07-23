using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

using Sd = System.Drawing;

namespace PptPlus.Components
{
    public class GH_PP_Mod_Font : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the MyComponent1 class.
        /// </summary>
        public GH_PP_Mod_Font()
          : base("PowerPoint Font", "Ppt Font",
              "Modify Powerpoint Font if applicable",
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
            pManager.AddTextParameter("Family Name", "F", "Optionally set Font Family name", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddNumberParameter("Size", "S", "Optionaly set Text Size", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddColourParameter("Color", "C", "Optionaly set Text Color", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddColourParameter("Highlight", "H", "Optionaly set Text Highlight Color", GH_ParamAccess.item);
            pManager[4].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.item);
            pManager.AddTextParameter("Family Name", "F", "Optionally set Font Family name", GH_ParamAccess.item);
            pManager.AddNumberParameter("Size", "S", "Optionaly set Text Size", GH_ParamAccess.item);
            pManager.AddColourParameter("Color", "C", "Optionaly set Text Color", GH_ParamAccess.item);
            pManager.AddColourParameter("Highlight", "H", "Optionaly set Text Highlight Color", GH_ParamAccess.item);
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

            string family = "Arial";
            if (DA.GetData(1, ref family))
            {
                if (isParagraph)
                {
                    paragraph.Font.Family = family;
                    foreach (Fragment f in paragraph.Fragments) f.Font.Family = family;
                }
                else
                {
                    fragment.Font.Family = family;
                }
            }

            double size = 8.0;
            if (DA.GetData(2, ref size))
            {
                if (isParagraph)
                {
                    paragraph.Font.Size = size;
                    foreach (Fragment f in paragraph.Fragments) f.Font.Size = size;
                }
                else
                {
                    fragment.Font.Size = size;
                }
            }

            Sd.Color color = Sd.Color.Black;
            if (DA.GetData(3, ref color))
            {
                if (isParagraph)
                {
                    paragraph.Font.Color = color;
                    foreach (Fragment f in paragraph.Fragments) f.Font.Color = color;
                }
                else
                {
                    fragment.Font.Color = color;
                }
            }

            Sd.Color highlight = Sd.Color.Empty;
            if (DA.GetData(4, ref highlight))
            {
                if (isParagraph)
                {
                    paragraph.Font.Highlight = highlight;
                    foreach (Fragment f in paragraph.Fragments) f.Font.Highlight = highlight;
                }
                else
                {
                    fragment.Font.Highlight = highlight;
                }
            }

            if (isParagraph)
            {
                DA.SetData(0, paragraph);
                DA.SetData(1, paragraph.Font.Family);
                DA.SetData(2, paragraph.Font.Size);
                DA.SetData(3, paragraph.Font.Color);
                DA.SetData(4, paragraph.Font.Highlight);
            }
            else
            {
                DA.SetData(0, fragment);
                DA.SetData(1, fragment.Font.Family);
                DA.SetData(2, fragment.Font.Size);
                DA.SetData(3, fragment.Font.Color);
                DA.SetData(4, fragment.Font.Highlight);
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
                return Properties.Resources.Ppt_Mod_Font;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("5f8c46cc-7607-458a-9ab9-7b653c0d75ba"); }
        }
    }
}