using Microsoft.Office.Core;
using System;
using PPT = Microsoft.Office.Interop.PowerPoint;

public class ShapePPT
{
    public string Topic { get; set; }
    public string DeviceName { get; set; }
    public string DataStructure { get; set; }
    public string Parameters { get; set; }
    public string StatusWindow { get; set; }
    public PPT.TextFrame Text { get; set; }
 //   public Collection<MShape> Children { get; protected set; }

    // Params
    public int Top { get; set; }
    public int Left { get; set; }
    public int Height { get; set; }
    public int Width { get; set; }
    public MsoShapeType Type { get; set; }
    public MsoAutoShapeType AutoShapeType { get; set; }
    public MsoTriState HorizontalFlip { get; set; }
    public MsoTriState VerticalFlip { get; set; }
    public float Rotation { get; set; }
    public int ZOrderPosition { get; set; }

    // Fill
    public int ForeColor { get; set; }
    public int BackColor { get; set; }
    public MsoFillType FillType { get; set; }
    public float Transparency { get; set; }

    // Line
    public int LineForeColor { get; set; }
    public MsoLineDashStyle DashStyle { get; set; }
    public MsoLineStyle Style { get; set; }
    public MsoArrowheadLength BeginArrowheadLength { get; set; }
    public MsoArrowheadStyle BeginArrowheadStyle { get; set; }
    public MsoArrowheadWidth BeginArrowheadWidth { get; set; }
    public MsoArrowheadLength EndArrowheadLength { get; set; }
    public MsoArrowheadStyle EndArrowheadStyle { get; set; }
    public MsoArrowheadWidth EndArrowheadWidth { get; set; }
    public float LineTransparency { get; set; }
    public float Weight { get; set; }

    static public PPT.Shape createShapePPT(PPT.Slide slide, String name, float top, float left, float height, float width, MsoAutoShapeType type
        )
    {
        PPT.Shape shape = slide.Shapes.AddShape(
            (MsoAutoShapeType)3, left, top, width, height);
        shape.Name = name;
        return shape;
    }
    

}
