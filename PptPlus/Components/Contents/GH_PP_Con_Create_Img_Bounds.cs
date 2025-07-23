using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

using Sd = System.Drawing;

namespace PptPlus.Components
{
    public class GH_PP_Con_Create_Img_Bounds : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_WD_Con_Add_Image class.
        /// </summary>
        public GH_PP_Con_Create_Img_Bounds()
          : base("Powerpoint Image Content", "Ppt Img",
              "Construct a PowerPoint Image Content Object",
              Constants.ShortName, Constants.SubContent)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, "_" + Constants.Content.NickName, "Optional " + Constants.Content.Input, GH_ParamAccess.item);
            pManager[0].Optional = true;
            pManager.AddGenericParameter("Image", "I", "A System Drawing Bitmap or full FilePath to an image file", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddRectangleParameter(Constants.Boundary.Name, Constants.Boundary.NickName, Constants.Boundary.Input, GH_ParamAccess.item);
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Content.Name, Constants.Content.NickName, Constants.Content.Output, GH_ParamAccess.item);
            pManager.AddGenericParameter("Image", "I", "A System.Drawing.Bitmap image", GH_ParamAccess.item);
            pManager.AddRectangleParameter(Constants.Boundary.Name, Constants.Boundary.NickName, Constants.Boundary.Input, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            Content content = null;
            bool hasContent = DA.GetData(0, ref content);
            if(hasContent) content = new Content(content);

            string warning = "I input must be a System Drawing Bitmap or a full file path to an image file";
            IGH_Goo gooA = null;
            bool hasImage =DA.GetData(1, ref gooA);

            Sd.Bitmap image = null;
            bool isValid = false;
            if (hasImage)
            {
                if (gooA.CastTo<Sd.Bitmap>(out image)) isValid = true;

                if (!isValid)
                {
                    if (gooA.CastTo<string>(out string filepath))
                    {
                        if (System.IO.File.Exists(filepath))
                        {
                            image = new Sd.Bitmap(filepath);
                            isValid = true;
                        }
                        else
                        {
                            this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, warning);
                            return;
                        }
                    }
                    else
                    {
                        this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, warning);
                        return;
                    }
                }
            }

            if (isValid) content = Content.CreateImageContent(image, content);
            
                Rectangle3d boundary = new Rectangle3d();
                if (DA.GetData(2, ref boundary)) content.Boundary = boundary;

            if (content != null)
            {
                DA.SetData(0, content);
                DA.SetData(1, content.Image);
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
                return Properties.Resources.Ppt_Con_Image;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("dff706c3-bb28-4c5e-8bd6-34004f6d8850"); }
        }
    }
}