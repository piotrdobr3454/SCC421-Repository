using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using static MigrationFormApp.Engine;
using System.Collections.ObjectModel;
using Microsoft.Office.Interop.Excel;
using Aspose.Slides.Export;

namespace MigrationFormApp
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            System.Windows.Forms.Application.Run(new Form1());
            
        }
        public static bool SaveData(string jsonString)
        {
            string jsonString2 = jsonString.Replace(@"\", "/");
            string jsonString3 = File.ReadAllText(jsonString2);
           
            ProjectProperties items = JsonConvert.DeserializeObject<ProjectProperties>(jsonString3);
            string _path = jsonString.Substring(0, jsonString.IndexOf(".")) + ".ec";
            Engine.GetData(items, _path);
            return true;
        }
    }

    class Engine
    {
        static string pathToExcel = "C:\\ConveyorConverter\\Migration\\Migration\\fb4.xlsx";
        public static void setPathToExcel(string excelPath)
        {
            pathToExcel = excelPath;
        }
        public static void GetData(ProjectProperties items, string path)
        {
            string sep = ";";
            string tab = "\t";

            List<myProperties> propertiesList = new List<myProperties>();
            List<Shape> shapes = new List<Shape>();
            Dictionary<string, string> myDict = new Dictionary<string, string>();
            List<string> dXQconveyors = new List<string>();
            Dictionary<string, bool> elementsList = new Dictionary<string, bool>();
            var shapeList = new ArrayList();
            int index = 0;
            int ab = 0;

            List<string> mParams = new List<string>();
            List<myProperties> propertiesList2 = new List<myProperties>();
            List<Shape> shapes2 = new List<Shape>();
            Dictionary<string, string> myDict2 = new Dictionary<string, string>();
            Dictionary<string, string> tagsLis = new Dictionary<string, string>();
            List<string> tagsLisKey = new List<string>();
            List<string> tagsLisVal = new List<string>();
            Dictionary<string, string> ChildrenDict = new Dictionary<string, string>();
            List<string> ChildrenList = new List<string>();
            List<string> ChildrenListOutput = new List<string>();

            List<string> ShapesTags = new List<String>();
            var shapeList2 = new ArrayList();
            int index2 = 0;
            int ab2 = 0;
            //int index3 = 0;
            string currTime = DateTime.Now.ToString("yyyMMdd_HHmm");
            string topicName = "319101";
            string path2 = path.Substring(0, path.IndexOf(".")) + "_2.ec"; 


            List<string> safetySegments = new List<string>();
            string topicSeg = "";

            CreateList();

            for (int i = 1; i <= items.mySlide.Length; i++)
            {
                int k = i - 1;
                foreach (var item2 in items.mySlide[k].DXQconveyor.MultiDevices)
                {
                    dXQconveyors.Add(item2);

                }

                foreach (var item in items.mySlide[k].Shapes)
                {
                    string slideName = items.mySlide[k].Name;
                    string topic = item.Topic;
                    string deviceName = item.DeviceName;
                    string dataStructure = item.DataStructure;
                    string fileName = item.FileName;
                    string statusWindow = item.StatusWindow;
                    string directions = "";
                    string multiDevice = "GR1";
                    string parameters = item.Parameters;
                    string name = item.Name;

                    int pTop = (int)item.Params.Top;
                    int pLeft = (int)item.Params.Left;
                    int pHeight = (int)item.Params.Height;
                    int pWidth = (int)item.Params.Width;
                    pTop = ReScale(pTop);
                    pHeight = ReScale(pHeight);
                    pWidth = ReScale(pWidth);
                    string position = pLeft.ToString() + sep + pTop.ToString() + sep + pHeight.ToString() + sep + pWidth.ToString();
                    string settings = "";

                    Collection<MShape> children = item.Children;

                    Dictionary<string, string> tags = item.Tags;
                    string txtTmp = item.Text;
                    string ChildrenTmp = "";
                    string TagsTmp = "";

                    string myDictKey = slideName + "_" + deviceName;

                    if (deviceName != null)
                    {
                        propertiesList.Add(new myProperties()
                        {
                            myKey = myDictKey,
                            //myOutput = slideName + "\t" + topic + "\t" + deviceName + "\t" + dataStructure + "\t" + "\t" + statusWindow + "\t" + "\t" + "\t" + position + "\t" + parameters
                            myOutput = slideName + tab + topic + tab + deviceName + tab + dataStructure + tab + fileName + tab + statusWindow + tab + directions + tab + multiDevice + tab + position + tab + parameters
                        });
                        if (deviceName.Length < 7)
                        {
                            deviceName = deviceName + "XXXXX";
                        }
                        shapes.Add(new Shape(slideName, topic, deviceName, dataStructure, fileName, statusWindow, directions, multiDevice, position, parameters, name, settings
                            , children
                           , tags
                           , ChildrenTmp
                           , TagsTmp
                           , txtTmp
                            ));
                        index = index + 1;
                    }



                }

                shapes.Sort((x, y) => (x.Slidename + x.DeviceName.Substring(0, 1) + x.DeviceName.Substring(x.DeviceName.Length - 3) + x.DeviceName.Substring(x.DeviceName.Length - 5, 2) + x.Position).CompareTo(y.Slidename + y.DeviceName.Substring(0, 1) + y.DeviceName.Substring(y.DeviceName.Length - 3) + y.DeviceName.Substring(y.DeviceName.Length - 5, 2) + y.Position));
                shapes = OnlyUnique(shapes);
                shapes = GroupData2(shapes, dXQconveyors, pathToExcel);
                myDict.Add("A" + ab, "\t" + "\t" + "\t" + "\t" + "\t" + "\t" + "\t" + "\t" + "\t");
                ab = ab + 1;
                foreach (var shape in shapes)
                {
                    myDict.Add("A" + ab, shape.Slidename + tab + shape.Topic + tab + shape.DeviceName + tab + shape.DataStructure + tab + shape.FileName + tab + shape.StatusWindow + tab + shape.Directions + tab + shape.Multidevice + tab + shape.Position + tab + shape.Parameters);
                    ab = ab + 1;
                }

                shapes.Clear();
                //elementsList.Clear();

                if (propertiesList.Count > 0)
                {
                    propertiesList.Add(new myProperties()
                    {
                        myKey = "A" + i,
                        //myOutput = "\t" + "\t" + "\t" + "\t" + "\t" + "\t" + "\t" + "\t" + "\t"
                        myOutput = tab + tab + tab + tab + tab + tab + tab + tab + tab
                    });
                }




                foreach (var item4 in items.mySlide[k].Shapes)
                {


                    string slideName = items.mySlide[k].Name;
                    string deviceName2 = item4.DeviceName;
                    string multiDevice2 = "GR1";
                    string directions2 = "";
                    string name2 = item4.Name;

                    string parameters2 = "";
                    Collection<MShape> children2 = item4.Children;
                    Dictionary<string, string> tags2 = item4.Tags;


                    string children2Tmp = JsonConvert.SerializeObject(children2);

                    children2Tmp = children2Tmp.ToString();

                    string tsgs2Tmp = JsonConvert.SerializeObject(tags2);
                    tsgs2Tmp = tsgs2Tmp.ToString();


                    string topic2 = item4.Topic;
                    string dataStructure2 = item4.DataStructure;
                    string fileName2 = item4.FileName;
                    string statusWindow2 = item4.StatusWindow;

                    int pTop2 = (int)item4.Params.Top;
                    int pLeft2 = (int)item4.Params.Left;
                    int pHeight2 = (int)item4.Params.Height;
                    int pWidth2 = (int)item4.Params.Width;

                    string position2 = pLeft2.ToString() + sep + pTop2.ToString() + sep + pHeight2.ToString() + sep + pWidth2.ToString();

                    string type = item4.Params.Type;
                    string autoShapeType = item4.Params.AutoShapeType;
                    string horizontalFlip = item4.Params.HorizontalFlip;
                    string verticalFlip = item4.Params.VerticalFlip;
                    string rotation = item4.Params.Rotation;
                    string zOrderPosition = item4.Params.ZOrderPosition;
                    string foreColor = item4.Params.Fill.ForeColor;
                    string backColor = item4.Params.Fill.BackColor;
                    string type2 = item4.Params.Fill.Type;
                    string transparency = item4.Params.Fill.Transparency;
                    string foreColor2 = item4.Params.Line.ForeColor;
                    string dashStyle = item4.Params.Line.DashStyle;
                    string style = item4.Params.Line.Style;
                    string beginArrowheadLength = item4.Params.Line.BeginArrowheadLength;
                    string beginArrowheadStyle = item4.Params.Line.BeginArrowheadStyle;
                    string beginArrowheadWidth = item4.Params.Line.BeginArrowheadWidth;
                    string endArrowheadLength = item4.Params.Line.EndArrowheadLength;
                    string endArrowheadStyle = item4.Params.Line.EndArrowheadStyle;
                    string endArrowheadWidth = item4.Params.Line.EndArrowheadWidth;
                    string transparency2 = item4.Params.Line.Transparency;
                    string weight = item4.Params.Line.Weight;

                    string ChildrenTmp2 = children2Tmp;
                    string TagsTmp2 = tsgs2Tmp;
                    string txtTmp2 = item4.Text;

                    string myDictKey2 = slideName + "_" + name2;

                    string settings2 = type.ToString() + sep + autoShapeType.ToString() + sep + horizontalFlip.ToString() + sep + verticalFlip.ToString() + sep + rotation.ToString() + sep + zOrderPosition.ToString() + sep + foreColor.ToString() + sep + backColor.ToString() + sep + type2.ToString()
                                         + sep + transparency.ToString() + sep + foreColor2.ToString() + sep + dashStyle.ToString() + sep + style.ToString() + sep + beginArrowheadLength.ToString() + sep + beginArrowheadStyle.ToString() + sep + beginArrowheadWidth.ToString() + sep + endArrowheadLength.ToString()
                                         + sep + endArrowheadStyle.ToString() + sep + endArrowheadWidth.ToString() + sep + transparency2.ToString() + sep + weight.ToString();
                    ;


                    if (ChildrenTmp2 == "[]")
                    {

                    }
                    else if (deviceName2 == null)
                    {


                        multiDevice2 = name2;
                        string[] strSep = new string[] { "Params" };
                        ChildrenList = ChildrenTmp2.Split(strSep, StringSplitOptions.None).ToList();



                        for (int m = 1; m <= ChildrenList.Count - 1; m++)

                        {

                            string[] strSep2 = new string[] { "{:" };
                            char[] charsep = new char[] { '\\', '{', '}', '\"', '[', ']' };
                            string ListOutput = "";
                            string ListOutputParamBeforeFill = "";
                            string ListOutputBeforeLine = "";
                            string ListOutputAfterLine = "";
                            string ListOutputBeforeTopic = "";
                            string ListOutputAfterTopic = "";
                            string ListOutputBeforeTags = "";
                            string ListOutputAfterTags = "";
                            List<string> ChildrenListTmp = ChildrenList[m].Split(charsep, StringSplitOptions.RemoveEmptyEntries).ToList();

                            for (int ll = 1; ll <= ChildrenListTmp.Count - 1; ll++)
                            {

                                ListOutput = ListOutput + ChildrenListTmp[ll]
                                        ;
                            }


                            String input = ListOutput;
                            if (input.EndsWith(","))
                            {
                                input = input.Remove(input.Length - 1, 1);
                            }

                            int index99 = input.IndexOf("Fill");
                            if (index99 >= 0)
                            {
                                ListOutputParamBeforeFill = input.Substring(0, index99);
                                string ListOutputAfterFill = input.Substring(index99);
                                int index98 = ListOutputAfterFill.IndexOf("Line");
                                if (index98 >= 0)
                                {
                                    ListOutputBeforeLine = ListOutputAfterFill.Substring(0, index98);
                                    ListOutputAfterLine = ListOutputAfterFill.Substring(index98);
                                }
                                int index97 = ListOutputAfterLine.IndexOf("Topic");
                                if (index97 >= 0)
                                {
                                    ListOutputBeforeTopic = ListOutputAfterLine.Substring(0, index97);
                                    ListOutputAfterTopic = ListOutputAfterLine.Substring(index97);
                                }
                                int index96 = ListOutputAfterTopic.IndexOf("Tags");
                                if (index96 >= 0)
                                {
                                    ListOutputBeforeTags = ListOutputAfterTopic.Substring(0, index96);
                                    ListOutputAfterTags = ListOutputAfterTopic.Substring(index96);
                                }

                            }



                            var dictionaryMain = ListOutputParamBeforeFill
                                .Split(',')
                                .Select(part => part.Split(':'))
                                .Where(part => part.Length == 2)
                                .ToDictionary(sp => sp[0], sp => sp[1]);

                            var dictionaryFill = ListOutputBeforeLine

                                .Split(',')
                                .Select(part => part.Split(':'))
                                .Where(part => part.Length == 2)
                                .ToDictionary(sp => sp[0], sp => sp[1]);

                            var dictionaryLine = ListOutputBeforeTopic
                                .Split(',')
                                .Select(part => part.Split(':'))
                                .Where(part => part.Length == 2)
                                .ToDictionary(sp => sp[0], sp => sp[1]);

                            var dictionaryTopic = ListOutputBeforeTags
                                .Split(',')
                                .Select(part => part.Split(':'))
                                .Where(part => part.Length == 2)
                                .ToDictionary(sp => sp[0], sp => sp[1]);



                            string PosTmp = dictionaryMain["Top"] + sep + dictionaryMain["Left"] + dictionaryMain["Height"] + sep + dictionaryMain["Width"];

                            string outputForChildren = slideName + tab + dictionaryTopic["Name"] + tab + multiDevice2 + tab + PosTmp
                                ;
                            string AftertmpPos = "";
                            var kk = 0;
                            foreach (var item3 in dictionaryMain)
                            {
                                if (kk >= 4)
                                {

                                    AftertmpPos = AftertmpPos + item3.Value + sep;
                                    kk++;
                                }
                                else
                                {

                                    kk++;

                                }
                            }



                            string FillTmp = "";

                            foreach (var item in dictionaryFill)
                            {

                                FillTmp = FillTmp + item.Value + sep;

                            }

                            string LineTmp = "";

                            foreach (var item2 in dictionaryLine)
                            {

                                LineTmp = LineTmp + item2.Value + sep;

                            }
                            LineTmp = LineTmp.Remove(LineTmp.Length - 1, 1);

                            string tmpChild = "";
                            if (dictionaryTopic["Children"] == "")
                            {
                                tmpChild = "[]";
                            }
                            else
                            {
                                tmpChild = dictionaryTopic["Children"];
                            }

                            ListOutputAfterTags = ListOutputAfterTags.Replace("Tags:", "");
                            if (ListOutputAfterTags == "")
                            {
                                ListOutputAfterTags = "{}";
                            }
                            else
                            {
                                ListOutputAfterTags = "{\"" + ListOutputAfterTags + "\"}";
                            }
                            string txtTmpChil = "";
                            txtTmpChil = dictionaryTopic["Text"];

                            outputForChildren = outputForChildren + tab + tab + AftertmpPos + FillTmp + LineTmp + tab + ListOutputAfterTags + tab + txtTmpChil;
                            ChildrenListOutput.Add(outputForChildren);


                        }

                    }




                    if (deviceName2 == null)
                    {



                        propertiesList2.Add(new myProperties()
                        {

                            myKey2 = myDictKey2,
                            myOutput2 = slideName + tab + name2 + tab + multiDevice2 + tab + position2 + tab + settings2
                                        + ChildrenTmp2
                                        + TagsTmp2
                        });



                        shapes2.Add(new Shape(slideName, topic2, deviceName2, dataStructure2, fileName2, statusWindow2, directions2, multiDevice2, position2, parameters2, name2, settings2
                            , children2
                            , tags2
                            , ChildrenTmp2
                            , TagsTmp2
                            , txtTmp2
                                    ));
                        ShapesTags.Add(slideName + sep + topic2 + sep + deviceName2 + sep + dataStructure2 + sep + fileName2 + sep + statusWindow2 + sep + directions2 + sep + multiDevice2 + sep + position2 + sep + parameters2 + sep + name2 + sep + settings2 + sep
                                   + ChildrenTmp2
                                   + TagsTmp2
                            );

                        index2 = index2 + 1;
                    }

                }



                if (shapes2 != null)
                {
                    shapes2.Sort((xx, yy) => (xx.Slidename + xx.Name.Substring(0, 1) + xx.Position).CompareTo(yy.Slidename + yy.Name.Substring(0, 1) + yy.Position));
                    shapes2 = OnlyNoDevices(shapes2);
                    myDict2.Add("A" + ab2, "\t" + "\t" + "\t" + "\t" + "\t" + "\t");
                    ab2 = ab2 + 1;
                    foreach (var shape2 in shapes2)
                    {
                        myDict2.Add("A" + ab2, shape2.Slidename + tab + shape2.Name + tab + shape2.Multidevice + tab + shape2.Position + tab + shape2.Parameters + tab + shape2.Settings + tab
                            + shape2.TagsTmp + shape2.Text
                            );
                        ab2 = ab2 + 1;
                    }

                    shapes2.Clear();

                    foreach (var item99 in ChildrenListOutput)
                    {
                        myDict2.Add("A" + ab2, item99);
                        ab2 = ab2 + 1;
                    }

                    ChildrenListOutput.Clear();

                    if (propertiesList2.Count > 0)
                    {
                        propertiesList2.Add(new myProperties()
                        {
                            myKey2 = "A" + i,

                            myOutput2 = tab + tab + tab + tab + tab + tab + tab
                        });
                    }
                }

            }

            string _txtHeader = "Slide" + "\t" + "Topic" + "\t" + "Device Name" + "\t" + "Data Structure" + "\t" + "File Name"
                + "\t" + "Status Window" + "\t" + "Directions" + "\t" + "Multidevice" + "\t" + "Positions" + "\t" + "Parameters";
            if (!File.Exists(path))
            {
                using (StreamWriter swCreate = File.CreateText(path))
                {
                    swCreate.WriteLine(_txtHeader);
                    swCreate.Close();
                };
            }

            if (!File.Exists(getPathToTxt()))
            {
                using (StreamWriter swCreate = File.CreateText(getPathToTxt()))
                {
                    swCreate.WriteLine(_txtHeader);
                    swCreate.Close();
                };
            }


            foreach (var item in myDict)
            {
                using (StreamWriter swAppend = File.AppendText(path))
                {
                    swAppend.WriteLine(item.Value);
                    swAppend.Close();
                }

                using (StreamWriter swAppend = File.AppendText(getPathToTxt()))
                {
                    swAppend.WriteLine(item.Value);
                    swAppend.Close();
                }
            }

            string _txtHeader2 = "Slide" + "\t" + "Name" + "\t" + "Multidevice" + "\t" + "Positions" + "\t" + "Parameters" + "\t" + "Settings" + "\t" + "Tags" + "\t" + "Text";
            if (!File.Exists(path2))
            {
                using (StreamWriter swCreate = File.CreateText(path2))
                {
                    swCreate.WriteLine(_txtHeader2);
                    swCreate.Close();
                };
            }

            if (!File.Exists(getPathToTxt()))
            {
                using (StreamWriter swCreate = File.CreateText(getPathToTxt()))
                {
                    swCreate.WriteLine(_txtHeader2);
                    swCreate.Close();
                };
            }


            foreach (var item in myDict2)
            {
                using (StreamWriter swAppend = File.AppendText(path2))
                {
                    swAppend.WriteLine(item.Value);
                    swAppend.Close();
                }

                using (StreamWriter swAppend = File.AppendText(getPathToTxt()))
                {
                    swAppend.WriteLine(item.Value);
                    swAppend.Close();
                }
            }
        }

        public static void CreateSafetySegments(List<string> safetySegments, string path, string topic)
        {
            if (!File.Exists(path))
            {
                using (StreamWriter swCreate = File.CreateText(path))
                {
                    swCreate.Close();
                };
            }

            using (StreamWriter swAppend = File.AppendText(path))
            {
                swAppend.WriteLine("[SWVersion=TIA]");
                swAppend.WriteLine("[PLC=" + topic + "]");
                swAppend.WriteLine("[EcoScreenDB=True]");
            }


            foreach (var item in safetySegments)
            {
                using (StreamWriter swAppend = File.AppendText(path))
                {
                    swAppend.WriteLine(item);
                }
            }
        }
        public static List<Shape> OnlyNoDevices(List<Shape> shapes2)
        {

            List<Shape> shapeList2 = new List<Shape>();
            List<string> shapeListDeviceName2 = new List<string>();

            foreach (var shape2 in shapes2)
            {
                if (!shapeListDeviceName2.Contains(shape2.Slidename + shape2.Name))
                {
                    shapeListDeviceName2.Add(shape2.Slidename + shape2.Name);
                    shapeList2.Add(shape2);
                }
            }
            return shapeList2;

        }
        public static List<Shape> OnlyUnique(List<Shape> shapes)
        {
            List<Shape> shapeList = new List<Shape>();
            List<string> shapeListDeviceName = new List<string>();

            foreach (var shape in shapes)
            {
                if (!shapeListDeviceName.Contains(shape.Slidename + shape.DeviceName))
                {
                    shapeListDeviceName.Add(shape.Slidename + shape.DeviceName);
                    shapeList.Add(shape);
                }
            }
            return shapeList;
        }
        public static List<Shape> OrganizeGroupStatus(List<Shape> shapes)
        {
            foreach (var shape in shapes)
            {
                if (shape.Slidename.Contains("Group_Status"))
                {
                    shape.Slidename = "GroupStatus";
                    if (shape.DataStructure.Contains("DIM_CT"))
                    {
                        shape.DataStructure = "ud_Gp_DIMCT";
                    }
                    if (shape.FileName.Contains("DIM_CT"))
                    {
                        shape.FileName = "GP_DIMDIMCT";
                    }
                }
            }
            return shapes;
        }

        public static List<Shape> GroupData2(List<Shape> shapes, List<string> dXQconveyors, string pathToExcel)
        {
            ReplaceDS(shapes, pathToExcel);
            HandleMultiDevices(shapes, dXQconveyors);
            SetDirections(shapes);
            OrganizeGroupStatus(shapes);
            HandleSlideName(shapes);
            HandleRFID(shapes);
            shapes = HandleRoDip(shapes);
            HandleTT_LT_elem(shapes);
            HandleSafetyDeviceName(shapes);

            return shapes;
        }

        public static List<Shape> HandleMultiDevices(List<Shape> shapes, List<string> dXQconveyors)
        {
            int x = 0;
            foreach (var shape in shapes)
            {

                // Handling multidevice elements
                if (!shape.DataStructure.Contains("Safety") && !shape.DeviceName.Contains("MES") && dXQconveyors.Contains(shape.DeviceName.Substring(shape.DeviceName.Length - 3)))
                {
                    bool exist = false;
                    int saveIndex2 = 0;
                    for (int i = 0; i < x; i++)
                    {
                        if (shapes[i].DeviceName.Substring(shapes[i].DeviceName.Length - 3).Equals(shape.DeviceName.Substring(shape.DeviceName.Length - 3)))
                        {
                            exist = true;
                            saveIndex2 = i;
                            break;
                        }
                    }
                    if (exist && x > 0)
                    {
                        shape.Multidevice = "";
                        shape.Position = "";
                        shape.Parameters = "";
                        string num = shapes[saveIndex2].Multidevice.Substring(2);
                        int num1 = int.Parse(num);
                        num1 = num1 + 1;
                        shapes[saveIndex2].Multidevice = shapes[saveIndex2].Multidevice.Substring(0, 2) + num1.ToString();
                    }
                }
                else if (shape.DataStructure.Contains("SafetyLightBarrier"))
                {
                    bool exist2 = false;
                    int saveIndex3 = 0;
                    if (x > 0)
                    {

                        if (shapes[x - 1].DataStructure.Contains("SafetyLightBarrier") && !shapes[x - 1].Position.Equals(""))
                        {
                            exist2 = true;
                            saveIndex3 = x - 1;
                        }
                    }
                    if (exist2 && x > 0)
                    {
                        shape.Multidevice = "";
                        shape.Position = "";
                        shape.Parameters = "";
                        string num = shapes[saveIndex3].Multidevice.Substring(2);
                        int num1 = int.Parse(num);
                        num1 = num1 + 1;
                        shapes[saveIndex3].Multidevice = shapes[saveIndex3].Multidevice.Substring(0, 2) + num1.ToString();
                    }
                }
                x = x + 1;
            }
            return shapes;
        }
        public static List<Shape> ReplaceDS(List<Shape> shapes, string pathToExcel)
        {
            Application excel = new Application();
            Workbook wb = excel.Workbooks.Open(pathToExcel);
            Worksheet excelSheet = wb.ActiveSheet;

            //Read the first cell
            foreach (var shape in shapes)
            {
                for (int i = 1; i < excelSheet.UsedRange.Rows.Count + 1; i++)
                {
                    if (shape.StatusWindow.Length > 4)
                    {
                        if (shape.StatusWindow.Contains(excelSheet.Cells[i, 3].Value.ToString()))
                        {
                            shape.FileName = excelSheet.Cells[i, 1].Value.ToString();
                            shape.DataStructure = "ud_" + shape.FileName;
                        }
                    }
                    else if (shape.DataStructure.Contains("Safety"))
                    {
                        if (shape.DataStructure.Contains(excelSheet.Cells[i, 3].Value.ToString()))
                        {
                            shape.FileName = excelSheet.Cells[i, 1].Value.ToString();
                            shape.DataStructure = "ud_" + shape.FileName;
                        }
                    }

                }
            }
            wb.Close();
            excel.Quit();
            return shapes;
        }

        public static List<Shape> HandleRoDip(List<Shape> shapes)
        {
            List<Shape> shapes1 = new List<Shape>();
            foreach (var shape in shapes)
            {
                if (shape.FileName != null && !shape.FileName.Contains("RoDip"))
                {
                    shapes1.Add(shape);
                }
            }
            return shapes1;
        }

        public static List<Shape> HandleRFID(List<Shape> shapes)
        {
            foreach (var shape in shapes)
            {
                if (shape.DeviceName.Contains("IS_VISU_DB"))
                {
                    shape.DeviceName = shape.Topic + "_" + shape.Slidename.Substring(0, 4) + "_IS010";
                    shape.FileName = "RFID";
                    shape.DataStructure = "ud_RFID";
                    shape.StatusWindow = "FU556";
                    shape.Parameters = "TT;LL";
                }
            }
            return shapes;
        }

        public static List<Shape> HandleSlideName(List<Shape> shapes)
        {
            foreach (var shape in shapes)
            {
                if (shape.Slidename != null && shape.Slidename.Contains("_") && shape.Slidename.Contains("C"))
                {
                    shape.Slidename = shape.Slidename.Replace("_", "");
                    shape.Slidename = shape.Slidename.Substring(shape.Slidename.IndexOf("C"));
                    if (shape.Slidename.Contains("Part1"))
                    {
                        shape.Slidename = shape.Slidename.Replace("Part1", "A");
                    }
                    else if (shape.Slidename.Contains("Part2"))
                    {
                        shape.Slidename = shape.Slidename.Replace("Part2", "B");
                    }
                }

            }
            return shapes;
        }

        public static int ReScale(int value)
        {
            double val = (value / 1.3);
            return (int)val;
        }

        public static List<Shape> HandleTT_LT_elem(List<Shape> shapes)
        {
            foreach (var shape in shapes)
            {
                if (shape.FileName.Contains("TT2P2D2S"))
                {
                    shape.FileName = "Mtr_TTTT2P2D2SMovS";
                }
                else if (shape.FileName.Contains("EL16P2SEn"))
                {
                    shape.FileName = "Mtr_ELEL16P2SEnc";
                    shape.DataStructure = "ud_Mtr_EL16P2SEn";
                }
                else if (shape.FileName.Contains("TC16P2S"))
                {
                    shape.FileName = "Mtr_TCTC16P2SEnc";
                }
                else if (shape.DataStructure.Contains("Enc"))
                {
                    shape.DataStructure = shape.DataStructure.Substring(0, shape.DataStructure.Length);
                }
                else if (shape.FileName.Contains("LT2P2"))
                {
                    shape.FileName = "Mtr_LT" + shape.FileName.Substring(shape.FileName.IndexOf("_") + 1);
                }
            }
            return shapes;
        }

        public static List<Shape> HandleSafetyDeviceName(List<Shape> shapes)
        {
            foreach (var shape in shapes)
            {
                if (shape.DeviceName.Contains("DB_ESTOP"))
                {
                    shape.DeviceName = shape.Topic + "_DXQ.ES." + shape.DeviceName.Substring(shape.DeviceName.IndexOf(".") + 1);
                }
            }
            return shapes;
        }

        public static List<Shape> SetDirections(List<Shape> shapes)
        {
            List<string> files = CreateList();
            List<string> tempList = new List<string>();
            string tempElem = "";
            bool added = false;
            bool containsDirection = false;
            foreach (var shape in shapes)
            {
                foreach (string file in files)
                {
                    if (shape.FileName != null && shape.FileName.Contains("Safety"))
                    {
                        string sub = file;
                        sub = sub.Replace(" ", "");
                        if (sub.Contains(shape.FileName))
                        {
                            if (shape.FileName.Length + 4 != sub.Length)
                            {
                                string direction = sub.Substring(shape.FileName.Length);
                                if (shape.FileName.Contains("SafetyLightBarrier"))
                                {
                                    if (shape.Parameters.Equals("TT;LL"))
                                    {
                                        shape.Parameters = "BB;RR";
                                    }
                                    if (!tempList.Contains("H") || !tempList.Contains("V"))
                                    {
                                        tempList.Add("H");
                                        tempList.Add("V");
                                    }
                                }
                                else
                                {
                                    direction = direction.Substring(0, direction.IndexOf("."));
                                    if (!tempList.Contains(direction))
                                    {
                                        tempList.Add(direction);
                                    }
                                }
                            }

                        }
                    }
                    else if (shape.FileName != null && shape.FileName.Contains("ContrDesk"))
                    {
                        shape.FileName = "GP_CDKCDK";
                    }

                    else if (shape.FileName != null)
                    {
                        string sub = file.Replace("  ", " ");
                        sub = sub.Substring(sub.IndexOf(" ") + 1);
                        string tempo = sub.Replace(" ", "");
                        //  Console.WriteLine(shape.FileName.Substring(shape.FileName.IndexOf("_")));
                        int occurences = sub.Count(c => c == ' ');

                        int wordCount = 0, index = 0;

                        // skip whitespace until first word
                        while (index < sub.Length && char.IsWhiteSpace(sub[index]))
                            index++;

                        while (index < sub.Length)
                        {
                            // check if current char is part of a word
                            while (index < sub.Length && !char.IsWhiteSpace(sub[index]))
                                index++;

                            wordCount++;

                            // skip whitespace until next word
                            while (index < sub.Length && char.IsWhiteSpace(sub[index]))
                                index++;
                        }

                        if (wordCount > 2)
                        {
                            string part = sub.Replace("  ", " ");
                            string part2 = part.Substring(0, part.IndexOf(" "));
                            part = part.Substring(part2.Length + 1);
                            part = part.Substring(0, part.IndexOf(" "));
                            part2 = part2 + part;
                            if (part2.Equals(shape.FileName.Substring(shape.FileName.IndexOf("_") + 1)))
                            {
                                string direction = sub.Replace("  ", " ");
                                direction = direction.Substring((direction.IndexOf(" ") + 1 + part.Length + 1));

                                direction = direction.Substring(0, direction.Length - 4);
                                direction = direction.Replace("  ", "");
                                direction = direction.Replace(" ", "");

                                if (!tempList.Contains(direction))
                                {
                                    tempList.Add(direction);
                                }

                                if (!added)
                                {
                                    tempElem = shape.FileName.Substring(0, shape.FileName.IndexOf("_") + 1) + file.Substring(file.IndexOf("_") + 1, 2) + shape.FileName.Substring(shape.FileName.IndexOf("_") + 1);
                                    added = true;
                                }
                            }

                        }

                        else if (sub.Contains(" ") && sub.Substring(0, sub.IndexOf(" ")).Equals(shape.FileName.Substring(shape.FileName.IndexOf("_") + 1)))
                        {
                            string direction = sub.Substring(sub.IndexOf(" ") + 1);

                            direction = direction.Substring(0, direction.Length - 4);
                            direction = direction.Replace("  ", "");
                            direction = direction.Replace(" ", "");
                            // Console.WriteLine(direction);
                            if (!tempList.Contains(direction))
                            {
                                containsDirection = true;
                                tempList.Add(direction);
                            }

                            if (!added)
                            {
                                tempElem = shape.FileName.Substring(0, shape.FileName.IndexOf("_") + 1) + file.Substring(file.IndexOf("_") + 1, 2) + shape.FileName.Substring(shape.FileName.IndexOf("_") + 1);
                                added = true;
                            }
                        }
                    }
                }
                foreach (string elem in tempList)
                {
                    if (elem.Contains(" "))
                    {
                        elem.Replace(" ", "");
                    }
                    if (elem == tempList.Last())
                    {
                        shape.Directions = shape.Directions + elem;
                    }
                    else if (tempList.Count > 0)
                    {
                        shape.Directions = shape.Directions + elem + ";";
                    }

                }
                if (shape.FileName != null && !shape.FileName.Contains("Safety") && !shape.FileName.Contains("CDK") && containsDirection)
                {
                    shape.FileName = tempElem;
                }
                if (shape.FileName != null && shape.FileName.Contains("RU2"))
                {
                    shape.FileName = shape.FileName.Replace("Mtr_Ro", "Rolldoor");
                    shape.Parameters = "TT;RR;0;0";
                }
                tempList.Clear();
                added = false;
                containsDirection = false;
            }

            return shapes;
        }

        public static List<string> CreateList()
        {
            List<string> files = new List<string>();
            List<string> theList = new List<string>();
            String path = @"C:/Users/plraddop/Downloads/Library/ConveyorDBS3";
            String sdira = @"C:/Users/plraddop/Downloads/Library/ConveyorDBS3/List.txt";


            if (System.IO.Directory.Exists(path))
            {
                //Get files for the base path
                string[] baseDirectoryFiles = System.IO.Directory.GetFiles(path);
                files.AddRange(baseDirectoryFiles);

                // Get files in subdirectories (first level) of base path
                foreach (string directory in System.IO.Directory.GetDirectories(path))
                {
                    string[] directoryFiles = System.IO.Directory.GetFiles(directory);
                    files.AddRange(directoryFiles);
                }

                string[] filesArray = files.ToArray();
                File.WriteAllText(sdira, String.Empty);
                foreach (var file in files)
                {
                    using (StreamWriter swAppend = File.AppendText(sdira))
                    {
                        if (file.Contains(".png"))
                        {
                            swAppend.WriteLine(file.Substring(file.IndexOf("\\") + 1));
                            swAppend.Close();
                            theList.Add(file.Substring(file.IndexOf("\\") + 1));
                        }

                    }
                }

            }
            else
            {
                Console.WriteLine("NIE MA");
            }

            return theList;
        }

        public static String getPathToTxt()
        {
            return @"C:/ConveyorConverter/BMW_BRASIL/BMW_Paint_Shop/Conveyor/319101/319101- Visu/" + "generatedFile.txt";
        }
    }

    public class myProperties
    {
        public string myKey { get; set; }
        public string myOutput { get; set; }
        public string myKey2 { get; set; }
        public string myOutput2 { get; set; }
    }

    public class ProjectProperties
    {
        [JsonProperty("Slides")]
        public Slide[] mySlide { get; set; }
    }

    public class Slide
    {
        public List<Shape> Shapes { get; set; }
        public string Name { get; set; }
        public DXQconveyor DXQconveyor { get; set; }
        public TagListOutput TagListOutput { get; set; }
    }

    public class Shape
    {
        public Params Params { get; set; }
        public string Topic { get; set; }
        public string DeviceName { get; set; }
        public string DataStructure { get; set; }
        public string FileName { get; set; }
        public string Multidevice { get; set; }
        public string Parameters { get; set; }
        public string StatusWindow { get; set; }
        public string Directions { get; set; }
        public string Name { get; set; }
        public string Position { get; set; }
        public string Slidename { get; set; }
        public string Settings { get; set; }

        public Dictionary<string, string> Tags { set; get; }

        public Collection<MShape> Children { get; set; }

        public string ChildrenTmp { get; set; }
        public string TagsTmp { get; set; }

        public string Text { get; set; }

        public Shape(string Slidename, string Topic, string DeviceName, string DataStructure, string FileName, string StatusWindow, string Directions, string Multidevice, string Position, string Parameters, string Name, string Settings
            , Collection<MShape> Children
            , Dictionary<string, string> Tags
            , string ChildrenTmp
            , string TagsTmp
            , string Text
            ) 
        {
            this.Slidename = Slidename;
            this.Topic = Topic;
            this.DeviceName = DeviceName;
            this.DataStructure = DataStructure;
            this.FileName = FileName;
            this.StatusWindow = StatusWindow;
            this.Directions = Directions;
            this.Multidevice = Multidevice;
            this.Position = Position;
            this.Parameters = Parameters;
            this.Name = Name;

            this.Settings = Settings;
            this.Tags = Tags;
            this.Children = Children;
            this.ChildrenTmp = ChildrenTmp;
            this.TagsTmp = TagsTmp;
            this.Text = Text;
        }
    }

    public class ShapeComparer : IComparer
    {
        public int Compare(object x, object y)
        {
            return (new CaseInsensitiveComparer()).Compare(((Shape)x).Slidename + ((Shape)x).DeviceName.Substring(((Shape)x).DeviceName.Length - 3), (((Shape)y).Slidename + ((Shape)y).DeviceName.Substring(((Shape)y).DeviceName.Length - 3)));
        }
    }
    public class Params
    {
        public double Top { get; set; }
        public double Left { get; set; }
        public double Height { get; set; }
        public double Width { get; set; }

        public string Type { get; set; }
        public string AutoShapeType { get; set; }
        public string HorizontalFlip { get; set; }

        public string VerticalFlip { get; set; }
        public string Rotation { get; set; }
        public string ZOrderPosition { get; set; }

        public MFill Fill { get; set; }
        public MLine Line { get; set; }


        public class MFill
        {

            public string ForeColor { get; set; }
            public string BackColor { get; set; }
            public string Type { get; set; }
            public string Transparency { get; set; }
        }

        public class MLine
        {

            public string ForeColor { get; set; }
            public string DashStyle { get; set; }
            public string Style { get; set; }
            public string BeginArrowheadLength { get; set; }
            public string BeginArrowheadStyle { get; set; }
            public string BeginArrowheadWidth { get; set; }
            public string EndArrowheadLength { get; set; }
            public string EndArrowheadStyle { get; set; }
            public string EndArrowheadWidth { get; set; }
            public string Transparency { get; set; }
            public string Weight { get; set; }
        }

    }

}
