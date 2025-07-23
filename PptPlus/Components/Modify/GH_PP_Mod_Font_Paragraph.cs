using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace PptPlus.Components.Modify
{
    public class GH_PP_Mod_Font_Paragraph : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Mod_Font_Paragraph class.
        /// </summary>
        public GH_PP_Mod_Font_Paragraph()
          : base("PowerPoint Paragraph", "Ppt Para",
              "Modify Powerpoint Font Paragraph if applicable",
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
            pManager.AddNumberParameter("Line Spacing", "L", "The line spacing of a paragraph", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddIntegerParameter("Horizontal Alignment", "H", "Optionally override the horizontal alignment", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddIntegerParameter("Vertical Alignment", "V", "Optionally override the vertical alignment", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddBooleanParameter("Wrap", "W", "Optionally override the text Wrapping", GH_ParamAccess.item);
            pManager[4].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[2];
            foreach (Font.HorizontalAlignments value in Enum.GetValues(typeof(Font.HorizontalAlignments)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }

            Param_Integer paramB = (Param_Integer)pManager[3];
            foreach (Font.VerticalAlignments value in Enum.GetValues(typeof(Font.VerticalAlignments)))
            {
                paramB.AddNamedValue(value.ToString(), (int)value);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.item);
            pManager.AddNumberParameter("Line Spacing", "L", "The line spacing of a paragraph", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Horizontal Alignment", "H", "Get the horizontal alignment", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Vertical Alignment", "V", "Get the vertical alignment", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Wrap", "W", "Optionally override the text Wrapping", GH_ParamAccess.item);
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

            double spacing = 1.0;
            if (DA.GetData(1, ref spacing)) if (isParagraph) paragraph.LineSpacing = spacing;

            int horizontal = 0;
            if (DA.GetData(2, ref horizontal))
            {
                if (isParagraph)
                {
                    paragraph.Font.HorizontalAlignment = (Font.HorizontalAlignments)horizontal;
                    foreach (Fragment f in paragraph.Fragments) f.Font.HorizontalAlignment = (Font.HorizontalAlignments)horizontal;
                }
                else
                {
                    fragment.Font.HorizontalAlignment = (Font.HorizontalAlignments)horizontal;
                }
            }

            int vertical = 0;
            if (DA.GetData(3, ref vertical))
            {
                if (isParagraph)
                {
                    paragraph.Font.VerticalAlignment = (Font.VerticalAlignments)vertical;
                    foreach (Fragment f in paragraph.Fragments) f.Font.VerticalAlignment = (Font.VerticalAlignments)vertical;
                }
                else
                {
                    fragment.Font.VerticalAlignment = (Font.VerticalAlignments)vertical;
                }
            }

            bool wrap = true;
            if (DA.GetData(4, ref wrap))
            {
                if (isParagraph)
                {
                    paragraph.Font.Wrapped = wrap;
                    foreach (Fragment f in paragraph.Fragments) f.Font.Wrapped = wrap;
                }
                else
                {
                    fragment.Font.Wrapped = wrap;
                }
            }

            if (isParagraph)
            {
                DA.SetData(0, paragraph);
                DA.SetData(1, paragraph.Font.LineSpacing);
                DA.SetData(2, paragraph.Font.HorizontalAlignment);
                DA.SetData(3, paragraph.Font.Wrapped);
            }
            else
            {
                DA.SetData(0, fragment);
                DA.SetData(1, fragment.Font.LineSpacing);
                DA.SetData(2, fragment.Font.HorizontalAlignment);
                DA.SetData(3, fragment.Font.Wrapped);
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
                return Properties.Resources.Ppt_Mod_Font_Paragraph;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("44694AEC-75B7-4E8E-9D83-554060D29CB4"); }
        }
    }
}