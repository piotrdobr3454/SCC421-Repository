using System;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace MigrationFormApp
{
    public class SlidePPT
    {
        public string Name { get; set; }
        public float Ratio { get; set; }
        public static PPT.Slide createSlidePPT(PPT.Slides slides, String name, int slideindex, PPT.CustomLayout cl)
        {
            PPT.Slide slide = slides.AddSlide(slideindex, cl);
            slide.Name = name;
            return slide;
        }
        public static float getRatio(float val)
        {
            return (float)1*val;
        }
    }
}
