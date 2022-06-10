using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace MigrationFormApp
{
    public class MObject
    {
        public string Name
        {
            get;
            set;
        }
        public Dictionary<string, string> Tags
        {
            get;
            set;
        }

        public MObject()
        {
            Tags = new Dictionary<string, string>();
        }
        public void CreateTags(PPT.Tags t)
        {
            for (int i = 1; i <= t.Count; i++)
            {
                string n = t.Name(i);
                string v = t.Value(i);
                if (!n.Contains("SHADOW"))
                {
                    Tags.Add(n, v);
                }
            }
        }
    }

    public class MPPT
    {
        public Collection<MSlide> Slides
        {
            get;
            set;
        }

        public MPPT()
        {
            Slides = new Collection<MSlide>();
        }
    }
    public class MSlide : MObject
    {
        public Collection<MShape> Shapes
        {
            get;
            protected set;
        }
        public DXQconveyor DXQconveyor
        {
            get;
            set;
        }

        public TagListOutput TagListOutput

        {
            get;
            set;
        }

        private MSlide()
        {
            Shapes = new Collection<MShape>();
            DXQconveyor = new DXQconveyor();

            //kod od LS
            TagListOutput = new TagListOutput();

            //koniec kodu od LS
        }

        static public MSlide Create(PPT.Slide sld)
        {
            MSlide _slide = new MSlide();
            _slide.Name = getTitle(sld, sld.Name);
            _slide.CreateTags(sld.Tags);
            _slide.getMultiDevices(sld.Tags);

            //kod od LS
            _slide.getTagList(sld.Tags);
            //koniec kodu od LS
            return _slide;
        }
        //kod od LS
        public void getTagList(PPT.Tags t)
        {

            for (int i = 1; i <= t.Count; i++)
            {
                string n = t.Name(i);
                string v = t.Value(i);

                TagListOutput.TagListTemp.Add(n, v);
            }

        }
        //koniec kodu od LS
        static public string getTitle(PPT.Slide sld, string defaultValue)
        {
            if (sld.Shapes.HasTitle == MsoTriState.msoTrue)
            {
                if (sld.Shapes.Title.TextFrame.HasText == MsoTriState.msoTrue)
                {
                    return "_" + sld.Shapes.Title.TextFrame.TextRange.Text;
                }
            }

            return defaultValue;
        }

        public void getMultiDevices(PPT.Tags t)
        {
            for (int i = 1; i <= t.Count; i++)
            {
                string n = t.Name(i);
                string v = t.Value(i);

                if (n.Contains("GLOBAL"))
                {
                    MatchCollection matches = Regex.Matches(v, @"iVisibility[A-Za-z0-9]+?((?=[ ])|(?=,)|(?=:))");

                    foreach (Match match in matches)
                    {
                        string mKey = match.ToString().Replace("iVisibility", "");
                        string mVal = Regex.Replace(mKey, @"[A-Za-z]+?(?=[0-9])", "");
                        DXQconveyor.MultiDevices.Add(mVal);
                    }
                }
            }
        }
    }
    public class MShape : MObject
    {
        public MParams Params
        {
            get;
            set;
        }
        public string Topic
        {
            get;
            set;
        }
        public string DeviceName
        {
            get;
            set;
        }
        public string DataStructure
        {
            get;
            set;
        }
        public string FileName
        {
            get;
            set;
        }
        public string Parameters
        {
            get;
            set;
        }
        public string StatusWindow
        {
            get;
            set;
        }
        public PPT.TextFrame Text
        {
            get;
            set;
        }
        public Collection<MShape> Children
        {
            get;
            protected set;
        }

        public string DXQrow
        {
            get;
            set;
        }

        public MShape()
        {
            Children = new Collection<MShape>();

        }


        static public MShape Create(PPT.Shape shp)
        {
            MShape _shape = new MShape();
            _shape.Name = shp.Name;
            _shape.Params = MParams.Create(shp);
            _shape.CreateTags(shp.Tags);
            _shape.CreateLinks(shp);
            // _shape.SetValues(shp);

            if (shp.Type == MsoShapeType.msoGroup)
            {
                foreach (PPT.Shape gShp in shp.GroupItems)
                {
                    MShape _gShape = Create(gShp);
                    _shape.Children.Add(_gShape);
                }
            }

            if (shp.HasTextFrame == MsoTriState.msoCTrue)
                _shape.Text = shp.TextFrame;


            _shape.GetDirectionParameters(_shape.Children);

            return _shape;
        }

        public void SetValues(PPT.Shape shp)
        {
            //MShape _shape = new MShape();
            Name = shp.Name;
            Params = MParams.Create(shp);
            CreateTags(shp.Tags);
            CreateLinks(shp);

            if (shp.Type == MsoShapeType.msoGroup)
            {

                foreach (PPT.Shape gShp in shp.GroupItems)
                {
                    MShape _gShape = Create(gShp);
                    Children.Add(_gShape);
                }

            }

            if (shp.HasTextFrame == MsoTriState.msoCTrue)
                Text = shp.TextFrame;

            GetDirectionParameters(Children);
        }

        public void CreateLinks(PPT.Shape s)
        {
            if (s.Type != MsoShapeType.msoGroup && s.Type != MsoShapeType.msoOLEControlObject)
            {
                foreach (PPT.ActionSetting act in s.ActionSettings)
                {
                    if (act.Hyperlink.Address != null)
                    {
                        string n = "ESCR_MODULE_0_HYPERLINK01";
                        string v = $"&ocrlfPLACEHOLDER_Adress&ocrlf{act.Hyperlink.Address}&ocrlfPLACEHOLDER_Subadress&ocrlf";
                        Tags.Add(n, v);
                    }
                }
            }
        }

        public void GetShapeParameters(MShape shape)
        {
            foreach (KeyValuePair<string, string> tag in shape.Tags)
            {
                if (DeviceName != "null")
                {
                    bool anim = tag.Key.Contains("COLORBLINK");
                    bool swap = tag.Key.Contains("COLORSWAP"); //added
                    bool condRB = shape.Name.Contains("ArrowLeft") && tag.Value.Contains("_Rev");
                    bool condFX = shape.Name.Contains("ArrowOpen") && tag.Value.Contains("_Cl");
                    bool condEL = shape.Name.Contains("ArrowDown") && tag.Value.Contains("_Lower");
                    bool condTT = shape.Name.Contains("ArrowTurnCCW") && tag.Value.Contains("_Fwd");
                    bool condTC = shape.Name.Contains("ArrowTurnCW") && tag.Value.Contains("_Fwd");
                    bool condRG = tag.Value.Contains("_Pos") && tag.Value.Contains("_RG");

                    if (anim)
                    {

                        int rot = (int)shape.Params.Rotation;

                        if (condFX || condRB || condRG)
                        {
                            switch (rot)
                            {
                                case 0:
                                    Parameters = "BB;RR;0;0";
                                    break;
                                case 180:
                                    Parameters = "TT;RR;0;0";
                                    break;
                                case 90:
                                    Parameters = "RR;BB;0;0";
                                    break;
                                default:
                                    Parameters = "LL;BB;0;0";
                                    break;
                            }
                        }
                        else if (condEL)
                        {
                            switch (rot)
                            {
                                case 0:
                                    Parameters = "LL;BB;1;2";
                                    break;
                                case 180:
                                    Parameters = "RR;RR;1;2";
                                    break;
                                case 90:
                                    Parameters = "TT;RR;1;2";
                                    break;
                                default:
                                    Parameters = "BB;BB;1;2";
                                    break;
                            }
                        }
                        else if (condTT)
                        {
                            switch (rot)
                            {
                                case 0:
                                    Parameters = "RR;LL;1;2;0;Left;Mid;0";
                                    break;
                                case 180:
                                    Parameters = "LL;LL;1;2;90;Left;Mid;0";
                                    break;
                                case 90:
                                    Parameters = "BB;LL;1;2;90;Left;Mid;0";
                                    break;
                                default:
                                    Parameters = "TT;LL;1;2;90;Left;Mid;0";
                                    break;
                            }
                        }
                        else if (condTC)
                        {
                            switch (rot)
                            {
                                case 0:
                                    Parameters = "RR;LL;1;2;0;Right;Mid;0";
                                    break;
                                case 180:
                                    Parameters = "LL;LL;1;2;90;Rigt;Mid;0";
                                    break;
                                case 90:
                                    Parameters = "BB;LL;1;2;90;Right;Mid;0";
                                    break;
                                default:
                                    Parameters = "TT;LL;1;2;90;Right;Mid;0";
                                    break;
                            }
                        }
                    }

                    if (tag.Key.Contains("DIAGNOSISWINDOW"))
                    {
                        string workV = tag.Value.Replace("\r", "").Replace("\n", "");
                        //Data Structure
                        string mStr = "PLACEHOLDER_DIAGNOSISCONTROL";
                        int searchSt = workV.IndexOf(mStr);
                        int st = workV.IndexOf((char)39, searchSt) + 1;
                        int len = workV.IndexOf((char)39, st) - st;
                        string fu = workV.Substring(st, len);
                        StatusWindow = fu;

                        //Status Window
                        if (StatusWindow.Contains("."))
                            DataStructure = StatusWindow.Substring(0, StatusWindow.IndexOf("."));
                        else
                            DataStructure = StatusWindow;

                        //// Replacing data for the PPT2016
                        //if(StatusWindow.Contains("1101"))
                        //{
                        //    DataStructure = "Mtr_CRB2D2S2S";
                        //}

                        //else if(StatusWindow.Contains("122"))
                        //{
                        //    DataStructure = "Mtr_TT2P2D2SM";
                        //}
                        //else if(StatusWindow.Contains("133"))
                        //{
                        //    DataStructure = "Mtr_LT2P2D1SIL";
                        //}

                        //if(StatusWindow.Contains("FU010"))
                        //{
                        //    DataStructure = "Gp_Panel_CT";
                        //    FileName = "GP_MainMain_CT";
                        //}
                        //else if(StatusWindow.Contains("FU011"))
                        //{
                        //    DataStructure = "Gp_DIM_CT";
                        //    FileName = "GP_DIMDIM_CT";
                        //}
                        //else if(StatusWindow.Contains("FU209"))
                        //{
                        //    DataStructure = "Gp_Ilock16";
                        //    FileName = "GP_ILIlock16";
                        //}
                        //else if(StatusWindow.Contains("FU045"))
                        //{
                        //    DataStructure = "Gp_ContrDesk";
                        //    FileName = "GP_CDKCDK";
                        //}
                        //else if(StatusWindow.Contains("FU101"))
                        //{
                        //    DataStructure = "Mtr_RB2D2S2S";
                        //    FileName = "Mtr_RBRB2D2S2S";
                        //}
                        //else if(StatusWindow.Contains("FU172"))
                        //{
                        //    DataStructure = "Mtr_RB1D3SChln";
                        //    FileName = "Mtr_RBRB1D3SChainIn";
                        //}
                        //else if(StatusWindow.Contains("FU171"))
                        //{
                        //    DataStructure = "Mtr_SC2D1S";
                        //    FileName = "Mtr_RBSC2D1S";
                        //}
                        //else if(StatusWindow.Contains("FU173"))
                        //{
                        //    DataStructure = "Mtr_RB1D3SChOut";
                        //    FileName = "Mtr_RBR1D3SChainOut";
                        //}
                        //else if(StatusWindow.Contains("FU148"))
                        //{
                        //    DataStructure = "Mtr_EL7P2SM";
                        //    FileName = "Mtr_ELEL7P2SMovS";
                        //}
                        //else if(StatusWindow.Contains("FU128"))
                        //{
                        //    DataStructure = "Mtr_FX2P2D1S";
                        //    FileName = "Mtr_FXFX2P2D1S";
                        //}
                        //else if(StatusWindow.Contains("FU101"))
                        //{
                        //    DataStructure = "Mtr_RB2D2S2S";
                        //    FileName = "Mtr_RBRB2D2S2S";
                        //}



                        //Topic
                        mStr = "PLACEHOLDER_PLC";
                        searchSt = workV.IndexOf(mStr);
                        st = workV.IndexOf((char)39, searchSt) + 1;
                        len = workV.IndexOf((char)39, st) - st;
                        string plc = workV.Substring(st, len);
                        Topic = plc;

                        //DeviceName                    
                        mStr = "PLACEHOLDER_ID";
                        searchSt = workV.IndexOf(mStr);
                        st = workV.IndexOf((char)39, searchSt) + 2;
                        len = workV.IndexOf((char)39, st) - st;
                        string did = workV.Substring(st, len);
                        DeviceName = did;
                        //  DeviceName = Topic + "_" + DeviceName;

                        if (DeviceName.Contains("_CDK"))
                        {
                            Parameters = $"{DeviceName};0";
                        }

                    }

                    //safety
                    if (tag.Value.Contains("ESTOP") && anim)
                    {
                        string workV = tag.Value.Replace("\r", "").Replace("\n", "");
                        //Data Structure
                        string mStr = "PLACEHOLDER_CONDITION";
                        int searchSt = workV.IndexOf(mStr);
                        int st = workV.IndexOf((char)39, searchSt) + 1;
                        int len = workV.IndexOf((char)39, st) - st;
                        string fu = workV.Substring(st, len);
                        StatusWindow = "";

                        if (fu.Contains("E01"))
                        {
                            DataStructure = "Safety:LB";
                        }
                        else if (fu.Contains("S01F"))
                        {
                            DataStructure = "Safety:AD";
                        }
                        else
                        {
                            DataStructure = "Safety:EE";
                        }


                        //Topic
                        mStr = "PLACEHOLDER_CONDITION";
                        searchSt = workV.IndexOf(mStr);
                        st = workV.IndexOf((char)39, searchSt) + 1;
                        len = workV.IndexOf((char)46, st) - st;
                        string plc = workV.Substring(st, len);
                        Topic = plc;

                        //DeviceName                    
                        mStr = "PLACEHOLDER_CONDITION";
                        searchSt = workV.IndexOf(mStr);
                        st = workV.IndexOf((char)46, searchSt) + 2;
                        len = workV.IndexOf((char)39, st) - st;
                        string did = workV.Substring(st, len);
                        DeviceName = did;

                        switch ((int)shape.Params.Rotation)
                        {
                            case 0:
                                Parameters = "TT;LL";
                                break;
                            case 180:
                                Parameters = "TT:LL";
                                break;
                            case 90:
                                Parameters = "LL;TT";
                                break;
                            default:
                                Parameters = "LL;TT";
                                break;
                        }

                    }


                    else if (tag.Value.Contains("IS_VISU") && anim)
                    {
                        string workV = tag.Value.Replace("\r", "").Replace("\n", "");
                        //Data Structure
                        string mStr = "PLACEHOLDER_CONDITION";
                        int searchSt = workV.IndexOf(mStr);
                        int st = workV.IndexOf((char)39, searchSt) + 1;
                        int len = workV.IndexOf((char)39, st) - st;
                        string fu = workV.Substring(st, len);
                        StatusWindow = "";

                        ////Status Window
                        //if (StatusWindow.Contains("."))
                        //    DataStructure = StatusWindow.Substring(0, StatusWindow.IndexOf("."));
                        //else
                        //    DataStructure = StatusWindow;

                        DataStructure = "Fault";

                        //Topic
                        mStr = "PLACEHOLDER_CONDITION";
                        searchSt = workV.IndexOf(mStr);
                        st = workV.IndexOf((char)39, searchSt) + 1;
                        len = workV.IndexOf((char)46, st) - st;
                        string plc = workV.Substring(st, len);
                        Topic = plc;

                        //DeviceName                    
                        mStr = "PLACEHOLDER_CONDITION";
                        searchSt = workV.IndexOf(mStr);
                        st = workV.IndexOf((char)46, searchSt) + 2;
                        len = workV.IndexOf((char)39, st) - st;
                        string did = workV.Substring(st, len);
                        DeviceName = did;
                    }

                    else if (tag.Value.Contains("ESTOP") && swap) // safety segments
                    {
                        string workV = tag.Value.Replace("\r", "").Replace("\n", "");
                        //Data Structure
                        string mStr = "PLACEHOLDER_CONDITION";
                        int searchSt = workV.IndexOf(mStr);
                        int st = workV.IndexOf((char)39, searchSt) + 1;
                        int len = workV.IndexOf((char)39, st) - st;
                        string fu = workV.Substring(st, len);
                        StatusWindow = "";

                        ////Status Window
                        //if (StatusWindow.Contains("."))
                        //    DataStructure = StatusWindow.Substring(0, StatusWindow.IndexOf("."));
                        //else
                        //    DataStructure = StatusWindow;

                        DataStructure = "SafetySeg";

                        //Topic
                        mStr = "PLACEHOLDER_CONDITION";
                        searchSt = workV.IndexOf(mStr);
                        st = workV.IndexOf((char)39, searchSt) + 1;
                        len = workV.IndexOf((char)46, st) - st;
                        string plc = workV.Substring(st, len);
                        Topic = plc;

                        //DeviceName                    
                        mStr = "PLACEHOLDER_CONDITION";
                        searchSt = workV.IndexOf(mStr);
                        st = workV.IndexOf((char)46, searchSt) + 2;
                        len = workV.IndexOf((char)39, st) - st;
                        string did = workV.Substring(st, len);
                        DeviceName = did;

                        switch ((int)shape.Params.Rotation)
                        {
                            case 0:
                                Parameters = "TT;LL";
                                break;
                            case 180:
                                Parameters = "TT:LL";
                                break;
                            case 90:
                                Parameters = "LL;TT";
                                break;
                            default:
                                Parameters = "LL;TT";
                                break;
                        }
                    }

                    DXQrow = $"{Topic}{(char)9}{DeviceName}{(char)9}{DataStructure}{(char)9}{DataStructure}{(char)9}{StatusWindow}{(char)9}{(char)9}GR1{(char)9}{shape.Params.Top};{shape.Params.Left};{shape.Params.Width};{shape.Params.Height}{(char)9}{Parameters}";

                }
                //kod od LS
                else if (DeviceName == "null")
                {

                    //   Tags.Add(tag.Key, tag.Value);
                    //    Console.WriteLine(Tags);

                    // string TagsCollection = "";

                    //dd(tag.Key + tag.Value);



                    Parameters = Parameters + ";" + tag.Key.ToString() + tag.Value.ToString() + ";";
                    Console.WriteLine(Parameters);



                }
            }
        }

        public void GetDirectionParameters(Collection<MShape> children)
        {
            foreach (MShape child in children)
            {
                GetShapeParameters(child);
            }
        }


    }

    public class MChild : MShape
    {

    }

    public class MParams
    {
        public int Top
        {
            get;
            set;
        }
        public int Left
        {
            get;
            set;
        }
        public int Height
        {
            get;
            set;
        }
        public int Width
        {
            get;
            set;
        }

        public MsoShapeType Type
        {
            get;
            set;
        }
        public MsoAutoShapeType AutoShapeType
        {
            get;
            set;
        }

        public MsoTriState HorizontalFlip
        {
            get;
            set;
        }
        public MsoTriState VerticalFlip
        {
            get;
            set;
        }
        public float Rotation
        {
            get;
            set;
        }
        public int ZOrderPosition
        {
            get;
            set;
        }

        public MFill Fill
        {
            get;
            set;
        }
        public MLine Line
        {
            get;
            set;
        }

        static public MParams Create(PPT.Shape shp)
        {
            MParams _Params = new MParams();
            _Params.Top = (int)shp.Top;
            _Params.Left = (int)shp.Left;
            _Params.Height = (int)shp.Height;
            _Params.Width = (int)shp.Width;
            _Params.Type = shp.Type;
            _Params.AutoShapeType = shp.AutoShapeType;
            _Params.HorizontalFlip = shp.HorizontalFlip;
            _Params.VerticalFlip = shp.VerticalFlip;
            _Params.Rotation = shp.Rotation;
            _Params.ZOrderPosition = shp.ZOrderPosition;
            _Params.Fill = MFill.Create(shp);
            _Params.Line = MLine.Create(shp);
            return _Params;
        }

    }
    public class MFill
    {
        public int ForeColor
        {
            get;
            set;
        }
        public int BackColor
        {
            get;
            set;
        }
        public MsoFillType Type
        {
            get;
            set;
        }
        public float Transparency
        {
            get;
            set;
        }
        static public MFill Create(PPT.Shape shp)
        {
            MFill _Fill = new MFill();
            if (shp.Type != MsoShapeType.msoOLEControlObject)
            {
                _Fill.ForeColor = shp.Fill.ForeColor.RGB;
                _Fill.BackColor = shp.Fill.BackColor.RGB;
                _Fill.Type = shp.Fill.Type;
                _Fill.Transparency = shp.Fill.Transparency;
            }

            return _Fill;
        }
    }

    public class MLine
    {
        public int ForeColor
        {
            get;
            set;
        }

        public MsoLineDashStyle DashStyle
        {
            get;
            set;
        }
        public MsoLineStyle Style
        {
            get;
            set;
        }
        public MsoArrowheadLength BeginArrowheadLength
        {
            get;
            set;
        }
        public MsoArrowheadStyle BeginArrowheadStyle
        {
            get;
            set;
        }
        public MsoArrowheadWidth BeginArrowheadWidth
        {
            get;
            set;
        }
        public MsoArrowheadLength EndArrowheadLength
        {
            get;
            set;
        }
        public MsoArrowheadStyle EndArrowheadStyle
        {
            get;
            set;
        }
        public MsoArrowheadWidth EndArrowheadWidth
        {
            get;
            set;
        }

        public float Transparency
        {
            get;
            set;
        }
        public float Weight
        {
            get;
            set;
        }

        static public MLine Create(PPT.Shape shp)
        {
            MLine _Line = new MLine();
            if (shp.Type != MsoShapeType.msoOLEControlObject)
            {
                _Line.ForeColor = shp.Line.ForeColor.RGB;
                _Line.DashStyle = shp.Line.DashStyle;
                _Line.Style = shp.Line.Style;
                _Line.BeginArrowheadLength = shp.Line.BeginArrowheadLength;
                _Line.BeginArrowheadStyle = shp.Line.BeginArrowheadStyle;
                _Line.BeginArrowheadWidth = shp.Line.BeginArrowheadWidth;
                _Line.EndArrowheadLength = shp.Line.EndArrowheadLength;
                _Line.EndArrowheadStyle = shp.Line.EndArrowheadStyle;
                _Line.EndArrowheadWidth = shp.Line.EndArrowheadWidth;
            }

            return _Line;
        }

    }

    public class MTags : Dictionary<string, string>
    {
        public Dictionary<string, string> Tags
        {
            get;
            protected set;
        }

        public MTags()
        {
            // MTags _Tags = new MTags();

            Tags = new Dictionary<string, string>();
        }

    }

    public class DXQconveyor
    {
        public Dictionary<string, DXQrow> Elements
        {
            get;
            set;
        }



        public Collection<string> MultiDevices
        {
            get;
            set;
        }



        public DXQconveyor()
        {
            Elements = new Dictionary<string, DXQrow>();
            MultiDevices = new Collection<string>();
        }
        public void createMultiElements()
        {
            Dictionary<string, string> myDevices = new Dictionary<string, string>();
            foreach (KeyValuePair<string, DXQrow> elem in Elements)
            {
                foreach (string mdev in MultiDevices)
                {
                    if (elem.Key.Contains(mdev))
                    {
                        myDevices.Add(elem.Key, elem.Value.RawData);
                    }
                    var sortedDevices = myDevices.OrderByDescending(x => x.Key);
                }
            }
        }
    }



    public class DXQrow
    {
        public string RawData { get; set; }
        public string CD { get; set; }
        public string Group { get; set; }
        public string Device { get; set; }
        public Collection<string> Multi { get; set; }
        static public DXQrow createDXQrow(string row, string devicename)
        {

            DXQrow _row = new DXQrow();
            _row.RawData = row;
            Match c = Regex.Match(devicename, @"^.+?(?=_)");
            Match g = Regex.Match(devicename, @"_.+?_");
            Match d = Regex.Match(devicename, @"_.{0,6}?$");
            _row.CD = c.Value; //.Replace("_","");
            _row.Group = g.Value; //.Replace("_", "");
            _row.Device = d.Value; //.Replace("_", "");
            _row.Multi = new Collection<string>();
            return _row;
        }

    }
    //kod od LS
    public class TagListOutput

    {
        public Dictionary<string, string> TagListTemp { get; set; }

        public TagListOutput()

        {

            TagListTemp = new Dictionary<string, string>();


        }

    }
    //koniec kodu od LS
}