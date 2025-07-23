using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Components;
using Grasshopper.Kernel.Types;
using Rhino;
using Rhino.DocObjects;
using Rhino.Geometry;

namespace PptPlus.Components
{
    public abstract class GH_PP_preview : GH_Component
    {
        private List<NurbsCurve> boundaries = new List<NurbsCurve>();
        private BoundingBox _displayBox = new BoundingBox();
        public override bool IsBakeCapable => (boundaries.Count) > 0;
        /// <summary>
        /// Initializes a new instance of the GH_PP_preview class.
        /// </summary>
        public GH_PP_preview()
          : base("GH_PP_preview", "Nickname",
              "Description",
              "Category", "Subcategory")
        {
        }
        public GH_PP_preview(string Name, string NickName, string Description, string Category, string Subcategory) : base(Name, NickName, Description, Category, Subcategory)
        {
            //this.DisplayExpired += GH_Pdf__Base_DisplayExpired;
        }

        protected override void BeforeSolveInstance()
        {
            boundaries = new List<NurbsCurve>();
            _displayBox = BoundingBox.Unset;
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
        }

        protected void AddRectangles(Rectangle3d rectangle)
        {
            if (!this.Hidden & !this.Locked)
            {
                this.boundaries.Add(rectangle.ToNurbsCurve());
                this._displayBox.Union(rectangle.BoundingBox);
            }
        }

        protected void AddRectangles(List<Rectangle3d> rectangles)
        {
            if (!this.Hidden & !this.Locked)
            {
                foreach (Rectangle3d rectangle in rectangles)
                {
                    this.boundaries.Add(rectangle.ToNurbsCurve());
                    this._displayBox.Union(rectangle.BoundingBox);
                }
            }
        }
        public override BoundingBox ClippingBox
        {
            get
            {
                return _displayBox;
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
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("ADE82A7C-EBE9-4283-9914-DBC0579130F7"); }
        }

        public override void DrawViewportMeshes(IGH_PreviewArgs args)
        {
            if (Hidden) return;
            if (Locked) return;

            Rhino.Display.DisplayMaterial mat = new Rhino.Display.DisplayMaterial();
            if (Attributes.Selected)
            {
                mat = args.ShadeMaterial_Selected;
            }
            else
            {
                mat = args.ShadeMaterial;
            }

            foreach (NurbsCurve curve in this.boundaries)
            {
                args.Display.DrawCurve(curve, mat.Diffuse);

            }

            // Set Display Override
            base.DrawViewportWires(args);
        }

        public override void DrawViewportWires(IGH_PreviewArgs args)
        {
            if (Hidden) return;
            if (Locked) return;

            Rhino.Display.DisplayMaterial mat = new Rhino.Display.DisplayMaterial();
            if (Attributes.Selected)
            {
                mat = args.ShadeMaterial_Selected;
            }
            else
            {
                mat = args.ShadeMaterial;
            }

            foreach (NurbsCurve curve in this.boundaries)
            {
                args.Display.DrawCurve(curve, mat.Diffuse);

            }

            // Set Display Override
            base.DrawViewportWires(args);
        }

        public override void BakeGeometry(RhinoDoc doc, ObjectAttributes att, List<Guid> obj_ids)
        {
            if (this.boundaries.Count > 0 && this.IsBakeCapable)
            {
                foreach (NurbsCurve rectangle in this.boundaries)
                {
                    GH_CustomPreviewItem item = new GH_CustomPreviewItem();
                    item.Geometry = (IGH_PreviewData)GH_Convert.ToGeometricGoo(rectangle);
                    GH_Material shader = new GH_Material(System.Drawing.Color.Black);
                    item.Material = shader;
                    item.Shader = shader.Value;
                    item.Colour = shader.Value.Diffuse;
                    Guid guid = item.PushToRhinoDocument(doc, att);
                    if (guid != Guid.Empty)
                    {
                        obj_ids.Add(guid);
                    }
                }
            }
        }
    }
}