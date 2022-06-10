using System;
using Newtonsoft.Json;
using Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;
using System.IO;

namespace MigrationFormApp
{
    static class ProgramConvertJSOn
    {

        [STAThread]
        public static bool CreateJSON(string pathToPPT, string exportPath)
        {
            var app = new PPT.Application()
            {
                //  Visible = MsoTriState.msoFalse,
                AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
            };

            var pres = app.Presentations;

            var file = pres.Open(pathToPPT, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

            MPPT myPPT = new MPPT();

            foreach (PPT.Slide sld in file.Slides)
            {
                MSlide mySlide = MSlide.Create(sld);

                foreach (PPT.Shape shp in sld.Shapes)
                {
                    MShape myShape = MShape.Create(shp);
                    myShape.GetShapeParameters(myShape);
                    myShape.GetDirectionParameters(myShape.Children);
                    mySlide.Shapes.Add(myShape);

                    string d = myShape.DeviceName;
                    if (d != "" && d != null)
                    {
                        DXQrow myRow = DXQrow.createDXQrow(myShape.DXQrow, d);
                        if (!mySlide.DXQconveyor.Elements.ContainsKey(d))
                        {
                            mySlide.DXQconveyor.Elements.Add(d, myRow);
                        }
                    }
                    SaveImage(shp, pathToPPT.Substring(0, pathToPPT.LastIndexOf("\\") + 1));
                }
                myPPT.Slides.Add(mySlide);
            }

            //int charLocation = file.Name.IndexOf(".", StringComparison.Ordinal);

            string currDateTime = DateTime.Now.ToString("yyyyMMdd_HHmm");

            using (StreamWriter savefile = File.CreateText(exportPath + @".json"))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(savefile, myPPT);
            }

            return true;
        }
        public static void SaveImage(PPT.Shape shp, string createText)
        {
            if (shp.Type.Equals(MsoShapeType.msoGroup))
            {
                for (int x = 1; x <= shp.GroupItems.Count; x++)
                {
                    if (shp.GroupItems[x].Type.Equals(MsoShapeType.msoPicture))
                    {
                        try
                        {
                            shp.GroupItems[x].Export(createText + @"Pictures\" + shp.GroupItems[x].Name + ".png", Microsoft.Office.Interop.PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {

                        }
                    }
                }
            }
            else if (shp.Type.Equals(MsoShapeType.msoPicture))
            {
                try
                {
                    shp.Export(createText + @"Pictures\" + shp.Name + ".png", Microsoft.Office.Interop.PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                }
            }
        }
    }

}
