using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace PurchasingProcedures
{
    public class DataGridViewHelper
    {
        public DataGridViewHelper(DataGridView gridview)
        {
            gridview.CellPainting += new DataGridViewCellPaintingEventHandler(gridview_CellPainting);
        }
        int top = 0;
        int left = 0;
        int height = 0;
        int width1 = 0;
        public void gridview_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            #region 重绘datagridview表头
            DataGridView dgv = (DataGridView)(sender);
            if (e.RowIndex != -1) return;
            foreach (TopHeader item in Headers)
            {
                if (e.ColumnIndex >= item.Index && e.ColumnIndex < item.Index + item.Span)
                {
                    if (e.ColumnIndex == item.Index)
                    {
                        top = e.CellBounds.Top;
                        left = e.CellBounds.Left;
                        height = e.CellBounds.Height;
                    }
                    int width = 0;
                    for (int i = item.Index; i < item.Span + item.Index; i++)
                    {
                        width += dgv.Columns[i].Width;
                    }
                    Rectangle rect = new Rectangle(left, top, width, e.CellBounds.Height);
                    using (Brush backColorBrush = new SolidBrush(e.CellStyle.BackColor))
                    {
                        //抹去原来的cell背景
                        e.Graphics.FillRectangle(backColorBrush, rect);
                    }
                    using (Pen gridLinePen = new Pen(dgv.GridColor))
                    {
                        e.Graphics.DrawLine(gridLinePen, left, top, left + width, top);
                        e.Graphics.DrawLine(gridLinePen, left, top + height / 2, left + width, top + height / 2);
                        width1 = 0;
                        e.Graphics.DrawLine(gridLinePen, left, top, left, top + height);
                        for (int i = item.Index; i < item.Span + item.Index; i++)
                        {
                            width1 += dgv.Columns[i].Width;
                            e.Graphics.DrawLine(gridLinePen, left + width1, top + height / 2, left + width1, top + height);
                        }
                        SizeF sf = e.Graphics.MeasureString(item.Text, e.CellStyle.Font);
                        float lstr = (width - sf.Width) / 2;
                        float rstr = (height / 2 - sf.Height) / 2;
                        //画出文本框
                        if (item.Text != "")
                        {
                            e.Graphics.DrawString(item.Text, e.CellStyle.Font,
                                                       new SolidBrush(e.CellStyle.ForeColor),
                                                         left + lstr,
                                                         top + rstr,
                                                         StringFormat.GenericDefault);
                        }
                        width = 0;
                        width1 = 0;
                        for (int i = item.Index; i < item.Span + item.Index; i++)
                        {
                            string columnValue = dgv.Columns[i].HeaderText;
                            width1 = dgv.Columns[i].Width;
                            sf = e.Graphics.MeasureString(columnValue, e.CellStyle.Font);
                            lstr = (width1 - sf.Width) / 2;
                            rstr = (height / 2 - sf.Height) / 2;
                            if (columnValue != "")
                            {
                                e.Graphics.DrawString(columnValue, e.CellStyle.Font,
                                                           new SolidBrush(e.CellStyle.ForeColor),
                                                             left + width + lstr,
                                                             top + height / 2 + rstr,
                                                             StringFormat.GenericDefault);
                            }
                            width += dgv.Columns[i].Width;
                        }
                    }
                    e.Handled = true;
                }
            }
            #endregion
        }
        private List<TopHeader> _headers = new List<TopHeader>();
        public List<TopHeader> Headers
        {
            get { return _headers; }
        }

        public struct TopHeader
        {
            public TopHeader(int index, int span, string text)
            {
                this.Index = index;
                this.Span = span;
                this.Text = text;
            }
            public int Index;
            public int Span;
            public string Text;
        }

    }
}
