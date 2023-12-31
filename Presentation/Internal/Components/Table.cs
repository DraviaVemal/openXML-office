using OpenXMLOffice.Global;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation

{
    public class Table
    {
        private int X = 0;
        private int Y = 0;
        private int Width = 8128000;
        private int Height = 741680;
        private P.GraphicFrame? GraphicFrame;

        public void UpdatePosition(int X, int Y)
        {
            this.X = X;
            this.Y = Y;
            if (GraphicFrame != null)
            {
                GraphicFrame.Transform = new P.Transform
                {
                    Offset = new A.Offset { X = X, Y = Y },
                    Extents = new A.Extents { Cx = Width, Cy = Height }
                };
            }
        }

        public void UpdateSize(int Width, int Height)
        {
            this.Width = Width;
            this.Height = Height;
            if (GraphicFrame != null)
            {
                GraphicFrame.Transform = new P.Transform
                {
                    Offset = new A.Offset { X = X, Y = Y },
                    Extents = new A.Extents { Cx = Width, Cy = Height }
                };
            }
        }

        private A.Table CreateTable(TableRow[] TableRows, TableSetting TableSetting)
        {
            A.Table Table = new()
            {
                TableProperties = new A.TableProperties()
                {
                    FirstRow = true,
                    BandRow = true
                },
                TableGrid = CreateTableGrid(TableSetting)
            };
            // Add Table Data Row
            foreach (TableRow row in TableRows)
            {
                Table.Append(CreateTableRow(row));
            }
            return Table;
        }

        private A.TableRow CreateTableRow(TableRow Row)
        {
            A.TableRow TableRow = new()
            {
                Height = Row.Height
            };
            foreach (TableCell cell in Row.TableCells)
            {
                TableRow.Append(CreateTableCell(cell));
            }
            return TableRow;
        }

        private A.TableGrid CreateTableGrid(TableSetting TableSetting)
        {
            A.TableGrid TableGrid = new();
            foreach (TableColumnSetting Column in TableSetting.TableColumnSettings)
            {
                TableGrid.Append(new A.GridColumn() { Width = Column.Width });
            }
            return TableGrid;
        }

        private A.TableCell CreateTableCell(TableCell Cell)
        {
            A.TableCell TableCell = new();
            A.Paragraph Paragraph = new();
            if (Cell.Value == null)
            {
                Paragraph.Append(new A.EndParagraphRunProperties() { Language = "en-IN" });
            }
            else
            {
                Paragraph.Append(new TextBox().CreateTextRun(new TextBoxSetting()
                {
                    Text = Cell.Value,
                    TextBackground = Cell.TextBackground,
                    TextColor = Cell.TextColor
                }));
            }
            TableCell.Append(new A.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                Paragraph
            ));
            A.TableCellProperties TableCellProperties = new();
            TableCellProperties.Append(new A.LeftBorderLineProperties(
                Cell.LeftBorder ? new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.RightBorderLineProperties(
                Cell.RightBorder ? new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.TopBorderLineProperties(
                Cell.TopBorder ? new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.BottomBorderLineProperties(
                Cell.BottomBorder ? new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.TopLeftToBottomRightBorderLineProperties(
                Cell.BottomBorder ? new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.BottomLeftToTopRightBorderLineProperties(
                Cell.BottomBorder ? new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex() { Val = Cell.CellBackground }));
            TableCell.Append(TableCellProperties);
            return TableCell;
        }

        public P.GraphicFrame CreateTableGraphicFrame(TableRow[] TableRows, TableSetting TableSetting)
        {
            A.GraphicData GraphicData = new(CreateTable(TableRows, TableSetting))
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
            };
            GraphicFrame = new()
            {
                NonVisualGraphicFrameProperties = new P.NonVisualGraphicFrameProperties(
               new P.NonVisualDrawingProperties()
               {
                   Id = 1,
                   Name = "Table 1"
               },
               new P.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoGrouping = true }),
               new P.ApplicationNonVisualDrawingProperties()),
                Graphic = new A.Graphic()
                {
                    GraphicData = GraphicData
                },
                Transform = new P.Transform()
                {
                    Offset = new A.Offset()
                    {
                        X = X,
                        Y = Y
                    },
                    Extents = new A.Extents()
                    {
                        Cx = Width,
                        Cy = Height
                    }
                }
            };
            return GraphicFrame;
        }
    }
}