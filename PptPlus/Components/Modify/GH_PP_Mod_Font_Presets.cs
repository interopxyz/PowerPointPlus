using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace PptPlus.Components.Modify
{
    public class GH_PP_Mod_Font_Presets : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Mod_Font_Presets class.
        /// </summary>
        public GH_PP_Mod_Font_Presets()
          : base("PowerPoint Font Presets", "Ppt Font Presets",
              "Apply a Preset Font to a Paragraph, or a Fragment if applicable",
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
            pManager.AddIntegerParameter("Font Preset", "F", "A preset font configuration", GH_ParamAccess.item);
            pManager[1].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[1];
            foreach (Font.Presets value in Enum.GetValues(typeof(Font.Presets)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Output, GH_ParamAccess.item);
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

            int type = 0;
            if (DA.GetData(1, ref type))
            {
                if (isParagraph)
                {
                    paragraph.Font = Fonts.GetPreset((Font.Presets)type);
                    foreach (Fragment f in paragraph.Fragments) f.Font = Fonts.GetPreset((Font.Presets)type);
                    DA.SetData(0, paragraph);
                }
                else
                {
                    fragment.Font = Fonts.GetPreset((Font.Presets)type);
                    DA.SetData(0, fragment);
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
                return Properties.Resources.Ppt_Mod_Font_Presets_01;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("6B3EA799-44D5-4D9A-A4A4-27BF8EEC178B"); }
        }
    }
}