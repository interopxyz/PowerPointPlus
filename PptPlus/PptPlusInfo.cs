using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace PptPlus
{
    public class PptPlusInfo : GH_AssemblyInfo
    {
        public override string Name => "PowerPointPlus";

        //Return a 24x24 pixel bitmap to represent this GHA library.
        public override Bitmap Icon => Properties.Resources.PptPlus_Logo_24;

        //Return a short string describing the purpose of this GHA library.
        public override string Description => "A Grasshopper 3d plugin for Powerpoint File Read & Write";

        public override Guid Id => new Guid("5395B4B9-D17F-4504-B212-52CD879BB8CB");

        //Return a string identifying you or your company.
        public override string AuthorName => "David Mans";

        //Return a string representing your preferred contact details.
        public override string AuthorContact => "interopxyz@gmail.com";

        public override string AssemblyVersion => "1.0.2.0";
    }

    public class PptPlusCategoryIcon : GH_AssemblyPriority
    {
        public object Properties { get; private set; }

        public override GH_LoadingInstruction PriorityLoad()
        {
            Instances.ComponentServer.AddCategoryIcon(Constants.ShortName, PptPlus.Properties.Resources.PptPlus_Logo_16);
            Instances.ComponentServer.AddCategorySymbolName(Constants.ShortName, 'W');
            return GH_LoadingInstruction.Proceed;
        }
    }

}