using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace PptPlus.Components.Modify
{
    public class GH_PP_Mod_Borders : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Mod_Borders class.
        /// </summary>
        public GH_PP_Mod_Borders()
          : base("Powerpoint Content Borders", "Ppt Borders",
              "Modify Powerpoint Content Borders if applicable",
              Constants.ShortName, Constants.SubModify)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Input, GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddBooleanParameter("Top", "T", "Toggle visbility of Top Border", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddBooleanParameter("Bottom", "B", "Toggle visbility of Bottom Border", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddBooleanParameter("Left", "L", "Toggle visbility of Left Border", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddBooleanParameter("Right", "R", "Toggle visbility of Right Border", GH_ParamAccess.item);
            pManager[4].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Input, GH_ParamAccess.item);
            pManager.AddBooleanParameter("Top", "T", "Toggle visbility of Top Border", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Bottom", "B", "Toggle visbility of Bottom Border", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Left", "L", "Toggle visbility of Left Border", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Right", "R", "Toggle visbility of Right Border", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            bool top = false;
            bool bottom = false;
            bool left = false;
            bool right = false;

            if (gooA.TryGetFragment(out Fragment fragment))
            {
                if (DA.GetData(1, ref top)) fragment.Graphic.TopBorder = top;
                if (DA.GetData(2, ref bottom)) fragment.Graphic.BottomBorder = bottom;
                if (DA.GetData(3, ref left)) fragment.Graphic.LeftBorder = left;
                if (DA.GetData(4, ref right)) fragment.Graphic.RightBorder = right;

                DA.SetData(0, fragment);
                DA.SetData(1, fragment.Graphic.TopBorder);
                DA.SetData(2, fragment.Graphic.BottomBorder);
                DA.SetData(3, fragment.Graphic.LeftBorder);
                DA.SetData(4, fragment.Graphic.RightBorder);
            }
            else if (gooA.TryGetParagraph(out Paragraph paragraph))
            {
                if (DA.GetData(1, ref top)) foreach (Fragment fragment1 in paragraph.Fragments) fragment1.Graphic.TopBorder = top;
                if (DA.GetData(2, ref bottom)) foreach (Fragment fragment1 in paragraph.Fragments) fragment1.Graphic.BottomBorder = bottom;
                if (DA.GetData(3, ref left)) foreach (Fragment fragment1 in paragraph.Fragments) fragment1.Graphic.LeftBorder = left;
                if (DA.GetData(4, ref right)) foreach (Fragment fragment1 in paragraph.Fragments) fragment1.Graphic.RightBorder = right;

                DA.SetData(0, paragraph);
                DA.SetData(1, paragraph.Fragments[0].Graphic.TopBorder);
                DA.SetData(2, paragraph.Fragments[0].Graphic.BottomBorder);
                DA.SetData(3, paragraph.Fragments[0].Graphic.LeftBorder);
                DA.SetData(3, paragraph.Fragments[0].Graphic.RightBorder);
            }
            else if (gooA.TryGetContent(out Content content))
            {
                if (DA.GetData(1, ref top)) content.Graphic.TopBorder = top;
                if (DA.GetData(2, ref bottom)) content.Graphic.BottomBorder = bottom;
                if (DA.GetData(3, ref left)) content.Graphic.LeftBorder = left;
                if (DA.GetData(4, ref right)) content.Graphic.RightBorder = right;

                DA.SetData(0, content);
                DA.SetData(1, content.Graphic.TopBorder);
                DA.SetData(2, content.Graphic.BottomBorder);
                DA.SetData(3, content.Graphic.LeftBorder);
                DA.SetData(4, content.Graphic.RightBorder);
            }
            else if (gooA.CastTo<PpSlide>(out PpSlide slide))
            {
                if (DA.GetData(1, ref top)) slide.Graphic.TopBorder = top;
                if (DA.GetData(2, ref bottom)) slide.Graphic.BottomBorder = bottom;
                if (DA.GetData(3, ref left)) slide.Graphic.LeftBorder = left;
                if (DA.GetData(4, ref right)) slide.Graphic.RightBorder = right;

                DA.SetData(0, slide);
                DA.SetData(1, slide.Graphic.TopBorder);
                DA.SetData(2, slide.Graphic.BottomBorder);
                DA.SetData(3, slide.Graphic.LeftBorder);
                DA.SetData(4, slide.Graphic.RightBorder);
            }
            else
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Con input must be a Content Object, Paragraph Object, Text Fragment Object or a string");
                return;
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
                return Properties.Resources.Ppt_Mod_Borders;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("301EBFFB-8E2A-41FF-AA7D-D21F9CFD9BF8"); }
        }
    }
}