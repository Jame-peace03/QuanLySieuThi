using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.SqlClient;
using app = Microsoft.Office.Interop.Excel.Application;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
namespace QuanLySieuThi
{
    class ExportExecl
    {
        public static string duongdanex = Application.StartupPath;
        /*public static void exportecxel(DataGridView g, string duongdan, string tenfile)
        {
            duongdan = duongdanex + @"\taikhoan\";
            app obj = new app();
            obj.Application.Workbooks.Add(Type.Missing);
            obj.Columns.ColumnWidth = 25;
            for (int i = 1; i < g.Columns.Count + 1; i++)
            {
                obj.Cells[1, i] = g.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < g.Rows.Count; i++)
                for (int j = 0; j < g.Columns.Count; j++)
                {
                    if (g.Rows[i].Cells[j].Value != null)
                    {
                        obj.Cells[i + 2, j + 1] = g.Rows[i].Cells[j].Value;
                    }
                }
            obj.Range["A1", "M100"].Font.Name = "Times New Roman";
            obj.Range["A1", "M100"].HorizontalAlignment = 3;
            obj.ActiveWorkbook.SaveCopyAs(duongdan + tenfile + ".xlsx");
            obj.ActiveWorkbook.Saved = true;
            //obj.Quit();
        }
        */

        public static void exportecxelchitietdonhang(DataGridView g, string duongdan, string tenfile)
        {
            duongdan = duongdanex + @"\ChiTietDonHang\";
            app obj = new app();
            obj.Application.Workbooks.Add(Type.Missing);
            obj.Columns.ColumnWidth = 25;
            for (int i = 1; i < g.Columns.Count + 1; i++)
            {
                obj.Cells[1, i] = g.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < g.Rows.Count; i++)
                for (int j = 0; j < g.Columns.Count; j++)
                {
                    if (g.Rows[i].Cells[j].Value != null)
                    {
                        obj.Cells[i + 2, j + 1] = g.Rows[i].Cells[j].Value;
                    }
                }
            obj.Range["A1", "M100"].Font.Name = "Times New Roman";
            obj.Range["A1", "M100"].HorizontalAlignment = 3;
            obj.ActiveWorkbook.SaveCopyAs(duongdan + tenfile + ".xlsx");
            obj.ActiveWorkbook.Saved = true;
            //obj.Quit();
        }
        public static void ExportToPDF(DataGridView g, string duongdan, string tenfile)
        {
            // Đường dẫn đến thư mục lưu trữ file PDF
            duongdan = duongdanex + @"\ChiTietDonHang-PDF\";

            // Tạo một tài liệu PDF mới
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(duongdan + tenfile + ".pdf", FileMode.Create));

            // Mở tài liệu để bắt đầu viết
            doc.Open();

            // Tạo một phần tử tiêu đề cho tài liệu PDF
            Paragraph title = new Paragraph("Thông tin don hàng", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 18f, iTextSharp.text.Font.BOLD));
            title.Alignment = Element.ALIGN_CENTER;
            doc.Add(title);

            // Thêm một đoạn trống (space)
            Paragraph space = new Paragraph(" ");
            doc.Add(space);

            PdfPTable pdfTable = new PdfPTable(g.Columns.Count);
            pdfTable.DefaultCell.Padding = 3;
            pdfTable.WidthPercentage = 100;
            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;

            // Thêm tiêu đề từng cột vào tệp PDF
            for (int i = 0; i < g.Columns.Count; i++)
            {
                PdfPCell cell = new PdfPCell(new Phrase(g.Columns[i].HeaderText, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12f, iTextSharp.text.Font.BOLD)));
                pdfTable.AddCell(cell);
            }

            // Thêm dữ liệu từ DataGridView vào tệp PDF
            for (int i = 0; i < g.Rows.Count; i++)
            {
                for (int j = 0; j < g.Columns.Count; j++)
                {
                    if (g.Rows[i].Cells[j].Value != null)
                    {
                        pdfTable.AddCell(g.Rows[i].Cells[j].Value.ToString());
                    }
                    else
                    {
                        pdfTable.AddCell("");
                    }
                }
            }

            // Thêm bảng vào tài liệu PDF
            doc.Add(pdfTable);

            // Đóng tài liệu
            doc.Close();
            writer.Close();
        }

        /* public static void export_phieu(DataGridView g, string duongdan, string tenfile, string solg)
         {
             duongdan = duongdanex + @"\ThongTinPhieu\";
             app obj = new app();
             obj.Application.Workbooks.Add(Type.Missing);
             obj.Columns.ColumnWidth = 25;


             for (int i = 1; i < g.Columns.Count + 1; i++)
             {
                 obj.Cells[1, i] = g.Columns[i - 1].HeaderText;
             }
             for (int i = 0; i < g.Rows.Count; i++)
                 for (int j = 0; j < g.Columns.Count; j++)
                 {
                     if (g.Rows[i].Cells[j].Value != null)
                     {
                         obj.Cells[i + 2, j + 1] = g.Rows[i].Cells[j].Value;
                     }
                 }
             obj.Cells[g.Rows.Count + 2, g.Columns.Count - 1] = "Số lượng : ";
             obj.Cells[g.Rows.Count + 2, g.Columns.Count] = " " + solg + "";

             obj.Range["A1", "M100"].Font.Name = "Times New Roman";
             obj.Range["A1", "M100"].HorizontalAlignment = 3;
             obj.ActiveWorkbook.SaveCopyAs(duongdan + tenfile + ".xlsx");
             obj.ActiveWorkbook.Saved = true;
             //obj.Quit();
         }
        */
        public static void nhapnhieu(DataGridView g, string duongdan, string tenfile, string s, string tile)
        {
            duongdan = duongdanex + @"\NhapNhieu\";
            app obj = new app();
            obj.Application.Workbooks.Add(Type.Missing);
            obj.Columns.ColumnWidth = 25;

 
            for (int i = 1; i < g.Columns.Count + 1; i++)
            {
                obj.Cells[1, i] = g.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < g.Rows.Count; i++)
                for (int j = 0; j < g.Columns.Count; j++)
                {
                    if (g.Rows[i].Cells[j].Value != null)
                    {
                        obj.Cells[i + 2, j + 1] = g.Rows[i].Cells[j].Value;
                    }
                }
            //obj.Cells[g.Rows.Count + 2, g.Columns.Count - 1] = "Chiếu khấu : ";
            //obj.Cells[g.Rows.Count + 2, g.Columns.Count] = " " + chietkhau + " %";
            obj.Cells[g.Rows.Count + 3, g.Columns.Count - 1] = "Tổng Tiền : ";
            obj.Cells[g.Rows.Count + 3, g.Columns.Count] = " " + s;
            obj.Range["A1", "M100"].Font.Name = "Times New Roman";
            obj.Range["A1", "M100"].HorizontalAlignment = 3;
            obj.ActiveWorkbook.SaveCopyAs(duongdan + tenfile + ".xlsx");
            obj.ActiveWorkbook.Saved = true;
            //obj.Quit();
        }
    }
}
