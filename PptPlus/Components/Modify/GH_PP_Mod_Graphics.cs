using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

using Sd = System.Drawing;

namespace PptPlus.Components.Modify
{
    public class GH_PP_Mod_Graphics : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Mod_Graphics class.
        /// </summary>
        public GH_PP_Mod_Graphics()
          : base("PowerPoint Graphics", "Ppt Graphics",
              "Modify Powerpoint Graphics when applicable",
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
            pManager.AddColourParameter("Fill Color", "F", "The background or fill color when applicable.", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddColourParameter("Stroke Color", "S", "The line stroke color when applicable.", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddNumberParameter("Stroke Weight", "W", "The line weight when applicable.", GH_ParamAccess.item);
            pManager[3].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Output, GH_ParamAccess.item);
            pManager.AddColourParameter("Fill Color", "F", "The background or fill color.", GH_ParamAccess.item);
            pManager.AddColourParameter("Stroke Color", "S", "The line stroke color.", GH_ParamAccess.item);
            pManager.AddNumberParameter("Stroke Weight", "W", "The line weight.", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            Sd.Color fill = Sd.Color.Empty;
            bool hasFill = DA.GetData(1, ref fill);

            Sd.Color stroke = Sd.Color.Empty;
            bool hasStroke = DA.GetData(2, ref stroke);

            double weight = 1.0;
            bool hasWeight = DA.GetData(3, ref weight);

            if (gooA.TryGetFragment(out Fragment fragment))
            {
                if (hasFill) fragment.Graphic.Fill = fill;
                if (hasStroke) fragment.Graphic.Stroke = stroke;
                if (hasWeight) fragment.Graphic.Weight = weight;

                DA.SetData(0, fragment);
                DA.SetData(1, fragment.Graphic.Fill);
                DA.SetData(2, fragment.Graphic.Stroke);
                DA.SetData(3, fragment.Graphic.Weight);
            }
            else if (gooA.TryGetParagraph(out Paragraph paragraph))
            {

                if (hasFill)
                {
                    paragraph.Graphic.Fill = fill;
                    foreach (Fragment fragment1 in paragraph.Fragments) fragment1.Graphic.Fill = fill;
                }
                if (hasStroke)
                {
                    paragraph.Graphic.Stroke = stroke;
                    foreach (Fragment fragment1 in paragraph.Fragments) fragment1.Graphic.Stroke = stroke;
                }
                if (hasWeight)
                {
                    paragraph.Graphic.Weight = weight;
                    foreach (Fragment fragment1 in paragraph.Fragments) fragment1.Graphic.Weight = weight;
                }

                DA.SetData(0, paragraph);
                DA.SetData(1, paragraph.Graphic.Fill);
                DA.SetData(2, paragraph.Graphic.Stroke);
                DA.SetData(3, paragraph.Graphic.Weight);
            }
            else if (gooA.TryGetContent(out Content content))
            {
                if (hasFill) content.Graphic.Fill = fill;
                if (hasStroke) content.Graphic.Stroke = stroke;
                if (hasWeight) content.Graphic.Weight = weight;

                DA.SetData(0, content);
                DA.SetData(1, content.Graphic.Fill);
                DA.SetData(2, content.Graphic.Stroke);
                DA.SetData(3, content.Graphic.Weight);
            }
            else if (gooA.CastTo<PpPresentation>(out PpPresentation presentation))
            {
                if (hasFill) presentation.Graphic.Fill = fill;
                if (hasStroke) presentation.Graphic.Stroke = stroke;
                if (hasWeight) presentation.Graphic.Weight = weight;

                DA.SetData(0, presentation);
                DA.SetData(1, presentation.Graphic.Fill);
                DA.SetData(2, presentation.Graphic.Stroke);
                DA.SetData(3, presentation.Graphic.Weight);
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
                return Properties.Resources.Ppt_Mod_Graphic;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("3D82FF62-72F8-47D8-A9D6-9752CCD2998F"); }
        }
    }
}