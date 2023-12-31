using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using OpenXMLOffice.Global;

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

        private A.TableRow CreateTableRow(TableRow row)
        {
            A.TableRow TableRow = new()
            {
                Height = row.Height
            };
            foreach (TableCell cell in row.TableCells)
            {
                TableRow.Append(CreateTableCell(cell));
            }
            return TableRow;
        }

        private A.TableCell CreateTableCell(TableCell cell)
        {
            A.TableCell TableCell = new();
            A.Paragraph Paragraph = new();
            if (cell.Value == null)
            {
                Paragraph.Append(new A.EndParagraphRunProperties() { Language = "en-IN" });
            }
            else
            {
                Paragraph.Append(new A.Run(new A.RunProperties() { Language = "en-IN", Dirty = false }, new A.Text(cell.Value)));
            }
            TableCell.Append(new A.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                Paragraph
            ));
            TableCell.Append(new A.TableCellProperties());
            return TableCell;
        }

        private A.TableGrid CreateTableGrid(TableSetting TableSetting)
        {
            A.TableGrid TableGrid = new();
            foreach (TableColumnSetting Column in TableSetting.TableColumnSettings)
            {
                TableGrid.Append(new A.GridColumn() { Width = 4064000 });
            }
            return TableGrid;
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