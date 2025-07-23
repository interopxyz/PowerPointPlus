using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Rg = Rhino.Geometry;

namespace PptPlus
{
    public class Page
    {

        #region members

        public enum SizesIsoA { A0, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10 };
        public enum SizesIsoB { B0, B1, B2, B3, B4, B5, B6, B7, B8, B9, B10 };
        public enum SizesUS { Letter, Legal, Statement, Tabloid, Executive };
        public enum SizesANSI { ANSIA, ANSIB, ANSIC, ANSID, ANSIE };
        public enum SizesRatio { R1x1, R5x4, R4x3, R3x2, R8x5, R16x9, R2x1, R7x3, Widescreen };

        public enum Orientations { Portrait, Landscape };

        Rg.Plane plane = Rg.Plane.WorldXY;

        string name = "Widescreen";
        //Units in Points
        double width = 959;
        double height = 540;

        public Orientations Orientation = Orientations.Portrait;

        #endregion

        #region constructors

        public Page()
        {

        }

        public Page(Page page)
        {
            this.plane = new Rg.Plane(page.plane);
            this.name = page.name;

            this.width = page.width;
            this.height = page.height;

            this.Orientation = page.Orientation;
        }

        public Page(string name, double width, double height)
        {
            this.name = name;
            this.width = width;
            this.height = height;
        }

        //North American
        public static Page Tabloid() { return new Page("Tabloid", 792, 1224); }
        public static Page Legal() { return new Page("Legal", 612, 1008); }
        public static Page Letter() { return new Page("Letter", 612, 792); }
        public static Page Statement() { return new Page("Statement", 396, 612); }
        public static Page Executive() { return new Page("Executive", 306, 756); }

        //ANSIC
        public static Page ANSIA() { return new Page("ANSIA", 612, 792); }
        public static Page ANSIB() { return new Page("ANSIB", 792, 1224); }
        public static Page ANSIC() { return new Page("ANSIC", 1224, 1584); }
        public static Page ANSID() { return new Page("ANSID", 1584, 2448); }
        public static Page ANSIE() { return new Page("ANSIE", 2448, 3168); }

        //ISO A Series
        public static Page A0() { return new Page("A0", 2383, 3370); }
        public static Page A1() { return new Page("A1", 1683, 2383); }
        public static Page A2() { return new Page("A2", 1190, 1683); }
        public static Page A3() { return new Page("A3", 841, 1190); }
        public static Page A4() { return new Page("A4", 595, 841); }
        public static Page A5() { return new Page("A5", 419, 595); }
        public static Page A6() { return new Page("A6", 297, 419); }
        public static Page A7() { return new Page("A7", 209, 297); }
        public static Page A8() { return new Page("A8", 147, 209); }
        public static Page A9() { return new Page("A9", 104, 147); }
        public static Page A10() { return new Page("A10", 73, 104); }

        //ISO B Series
        public static Page B0() { return new Page("B0", 2834, 4008); }
        public static Page B1() { return new Page("B1", 2004, 2834); }
        public static Page B2() { return new Page("B2", 1417, 2004); }
        public static Page B3() { return new Page("B3", 1000, 1417); }
        public static Page B4() { return new Page("B4", 708, 1000); }
        public static Page B5() { return new Page("B5", 498, 708); }
        public static Page B6() { return new Page("B6", 354, 498); }
        public static Page B7() { return new Page("B7", 249, 354); }
        public static Page B8() { return new Page("B8", 175, 249); }
        public static Page B9() { return new Page("B9", 124, 175); }
        public static Page B10() { return new Page("B10", 87, 124); }

        //RATIO
        public static Page R1x1() { return new Page("1:1", 720, 720); }
        public static Page R5x4() { return new Page("5:4", 720, 576); }
        public static Page R4x3() { return new Page("4:3", 720, 540); }
        public static Page R3x2() { return new Page("3:2", 720, 480); }
        public static Page R8x5() { return new Page("8:5", 720, 450); }
        public static Page R16x9() { return new Page("16:9", 720, 405); }
        public static Page R2x1() { return new Page("2:1", 720, 360); }
        public static Page R7x3() { return new Page("7:3", 720, 308); }
        public static Page Widescreen() { return new Page("Widescreen", 959, 540); }

        #endregion

        #region properties

        public virtual string Name
        {
            get { return this.name; }
            set { this.name = value; }
        }

        public virtual Rg.Plane Plane
        {
            get { return this.plane; }
            set { this.plane = new Rg.Plane(value); }
        }

        public virtual int Width
        {
            get { return (int)this.width; }
            set 
            { 
                this.width = value;
            }
        }

        public virtual int Height
        {
            get { return (int)this.height; }
            set 
            { 
                this.height = value;
            }
        }

        public virtual double Ratio
        {
            get { return this.width/this.height; }
        }

        public virtual Rg.Rectangle3d Boundary
        {
            get {
                if(this.Orientation == Orientations.Landscape) return new Rg.Rectangle3d(plane, height, width);
                return new Rg.Rectangle3d(plane, width, height); 
            }
        }


        #endregion

        #region methods

        public static Page Preset(Page.SizesIsoA type)
        {
            switch (type)
            {
                default:
                    return Page.A4();
                case SizesIsoA.A0:
                    return Page.A0();
                case SizesIsoA.A1:
                    return Page.A1();
                case SizesIsoA.A2:
                    return Page.A2();
                case SizesIsoA.A3:
                    return Page.A3();
                case SizesIsoA.A5:
                    return Page.A5();
                case SizesIsoA.A6:
                    return Page.A6();
                case SizesIsoA.A7:
                    return Page.A7();
                case SizesIsoA.A8:
                    return Page.A8();
                case SizesIsoA.A9:
                    return Page.A9();
                case SizesIsoA.A10:
                    return Page.A10();
            }
        }

        public static Page Preset(Page.SizesIsoB type)
        {
            switch (type)
            {
                default:
                    return Page.B4();
                case SizesIsoB.B0:
                    return Page.B0();
                case SizesIsoB.B1:
                    return Page.B1();
                case SizesIsoB.B2:
                    return Page.B2();
                case SizesIsoB.B3:
                    return Page.B3();
                case SizesIsoB.B5:
                    return Page.B5();
                case SizesIsoB.B6:
                    return Page.B6();
                case SizesIsoB.B7:
                    return Page.B7();
                case SizesIsoB.B8:
                    return Page.B8();
                case SizesIsoB.B9:
                    return Page.B9();
                case SizesIsoB.B10:
                    return Page.B10();
            }
        }

        public static Page Preset(Page.SizesUS type)
        {
            switch (type)
            {
                default:
                    return Page.Letter();
                case SizesUS.Executive:
                    return Page.Executive();
                case SizesUS.Legal:
                    return Page.Legal();
                case SizesUS.Statement:
                    return Page.Statement();
                case SizesUS.Tabloid:
                    return Page.Tabloid();
            }
        }

        public static Page Preset(Page.SizesANSI type)
        {
            switch (type)
            {
                default:
                    return Page.ANSIC();
                case SizesANSI.ANSIA:
                    return Page.ANSIA();
                case SizesANSI.ANSIB:
                    return Page.ANSIB();
                case SizesANSI.ANSID:
                    return Page.ANSID();
                case SizesANSI.ANSIE:
                    return Page.ANSIE();
            }
        }

        public static Page Preset(Page.SizesRatio type)
        {
            switch (type)
            {
                default:
                    return Page.R4x3();
                case SizesRatio.R16x9:
                    return Page.R16x9();
                case SizesRatio.R1x1:
                    return Page.R1x1();
                case SizesRatio.R2x1:
                    return Page.R2x1();
                case SizesRatio.R3x2:
                    return Page.R3x2();
                case SizesRatio.R5x4:
                    return Page.R5x4();
                case SizesRatio.R7x3:
                    return Page.R7x3();
                case SizesRatio.R8x5:
                    return Page.R8x5();
                case SizesRatio.Widescreen:
                    return Page.Widescreen();
            }
        }

        #endregion

    }
}
