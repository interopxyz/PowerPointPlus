﻿using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace PptPlus.Components.Modify
{
    public class GH_PP_Mod_CombineText : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_PP_Mod_CombineText class.
        /// </summary>
        public GH_PP_Mod_CombineText()
          : base("Combine PowerPoint Text", "Ppt CombineTxt",
              "Combine PowerPoint Text Paragraphs, Text Fragments, or Strings into a Text Paragraph Object",
              Constants.ShortName, Constants.SubModify)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.quinary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Fragment.Name, Constants.Fragment.NickName, Constants.Fragment.Inputs, GH_ParamAccess.list);
            pManager[0].Optional = false;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Paragraph.Name, Constants.Paragraph.NickName, Constants.Paragraph.Input, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            List<IGH_Goo> goos = new List<IGH_Goo>();
            if (!DA.GetDataList(0, goos)) return;
            List<Paragraph> paragraphs = new List<Paragraph>();
            foreach (IGH_Goo goo in goos) if (goo.TryGetParagraph(out Paragraph para)) paragraphs.Add(para);
            Paragraph paragraph = new Paragraph(paragraphs);

            DA.SetData(0, paragraph);
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
                return Properties.Resources.Ppt_Mod_Text_Combine;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("B733AE88-4053-429D-9150-E97E2CBCE4C2"); }
        }
    }
}