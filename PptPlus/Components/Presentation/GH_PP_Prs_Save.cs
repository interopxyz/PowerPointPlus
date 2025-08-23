using System;
using System.Collections.Generic;
using System.IO;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace PptPlus.Components
{
    public class GH_PP_Prs_Save : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the PT_GH_Presentation class.
        /// </summary>
        public GH_PP_Prs_Save()
          : base("PowerPoint Save Presentation", "Ppt Save",
              "Save the PowerPoint Presentation to a file",
              Constants.ShortName, Constants.SubPresentation)
        {
        }

        public override GH_Exposure Exposure => GH_Exposure.septenary;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Presentation.Name, Constants.Presentation.NickName, Constants.Presentation.Output, GH_ParamAccess.item);
            pManager[0].Optional = false;

            pManager.AddTextParameter("Directory", "D", "The directory or folder where the file will be saved", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddTextParameter("File Name", "N", "The Document name", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddBooleanParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item);
            pManager[3].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Filepath", "P", "The full filepath to the document", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool activate = false;
            if (DA.GetData(3, ref activate))
            {
                if (activate)
                {
                    PpPresentation presentation = new PpPresentation();
                    IGH_Goo gooA = null;
                    if (DA.GetData(0, ref gooA))
                    {
                        if (!gooA.CastTo<PpPresentation>(out presentation))
                        {
                            this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Prs input must be a PowerPoint Presentation Object");
                            return;
                        }
                        presentation = new PpPresentation(presentation);
                    }

                    string directory = "C:\\Users\\Public\\Documents\\";
                    if (DA.GetData(1, ref directory))
                    {
                        if (!Directory.Exists(directory))
                        {
                            this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "The specified directory does not exist");
                            return;
                        }
                    }
                    else
                    {
                        if (this.OnPingDocument().FilePath != null)
                        {
                            directory = Path.GetDirectoryName(this.OnPingDocument().FilePath) + "\\";
                        }
                    }
                    directory = Path.GetDirectoryName(directory);

                    string name = Constants.UniqueName;
                    DA.GetData(2, ref name);
                    name = Path.GetFileNameWithoutExtension(name);

                    string path = Path.Combine(directory, name + ".pptx");

                    presentation.Save(path);

                    DA.SetData(0, path);
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
                return Properties.Resources.Pt_Prs_Save;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("1F2B3F1C-AD85-467D-AB45-835F73A2DC84"); }
        }
    }
}