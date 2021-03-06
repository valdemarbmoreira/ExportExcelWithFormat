﻿using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ProjectDownloadExcel
{
    public partial class Main : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
                            ScriptManager.ScriptResourceMapping.AddDefinition("jquery",
                    new ScriptResourceDefinition
                    {
                        Path = "~/scripts/jquery-1.4.1.min.js",
                        DebugPath = "~/scripts/jquery-1.4.1.js",
                        CdnPath = "http://ajax.microsoft.com/ajax/jQuery/jquery-1.4.1.min.js",
                        CdnDebugPath = "http://ajax.microsoft.com/ajax/jQuery/jquery-1.4.1.js"
                    });

                            panel.CssClass = "formatPanelFather";
                        Panel p = new Panel();
                        p.Width= 100;
                        //p.Height = 100;
            
                        p.BackColor = Color.Azure;

                        p.CssClass = "formatPanel";
                        panel.Controls.Add(p);

                        Panel p2 = new Panel();
                        p2.Width = 100;
                        p2.Height = 100;
                        p2.CssClass = "formatPanel";
                        p2.BackColor = Color.Green;

            
                        panel.Controls.Add(p2);


                            CheckBox[] cbl = new CheckBox[20];
                            for (int i = 0; i < 10; i++)
                            {
                                
                                cbl[i] = new CheckBox();
                                cbl[i].Text = "teste " + i;
                                cbl[i].CssClass = "checkBoxFormat";
                                p.Controls.Add(cbl[i]);
                                // Response.Write("\n" + words[i]);
                            }

        }

        protected void btnGerar_Click(object sender, EventArgs e)
        {
            if (!Page.IsValid)
            {
                return;
            }


            //limpar mensagem de validacao:


            var data = new[]{ 
                               new{ Name="Ram", Email="ram@techbrij.com safaadfafafadsfas sdfdsf sfdasfs asdfsfa", Phone="111-222-3333" },
                               new{ Name="ShyamÇÇÇÇ", Email="shyam@techbrij.com", Phone="159-222-1596" },
                               new{ Name="Mohan", Email="mohan@techbrij.com", Phone="456-222-4569" },
                               new{ Name="Sohan", Email="sohan@techbrij.com", Phone="789-456-3333" },
                               new{ Name="Karan", Email="karan@techbrij.com", Phone="111-222-1234" },
                               new{ Name="Brij", Email="brij@techbrij.com", Phone="111-222-3333" }                       
                      };

            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add("Sony");
            workSheet.Cells[1, 1].LoadFromCollection(data, true);
            workSheet.Cells[1, 1, 1, 3].Style.Font.Bold = true;
            workSheet.Cells[1, 1, 1, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                var fill = workSheet.Cells[1, 1, 1, 3].Style.Fill;
                    fill.PatternType = ExcelFillStyle.Solid;
                    fill.BackgroundColor.SetColor(Color.Gray);

                    var border = workSheet.Cells[1, 1, 1, 3].Style.Border;
                    border.Bottom.Style = ExcelBorderStyle.Thin;
                    border.Top.Style = ExcelBorderStyle.Thin;
                    border.Left.Style = ExcelBorderStyle.Thin;
                    border.Right.Style = ExcelBorderStyle.Thin;

                    //workSheet.Cells.AutoFitColumns();
                    workSheet.Column(2).Width = 50;
                    workSheet.Column(2).Style.WrapText = true;
                    workSheet.Column(2).Style.VerticalAlignment = ExcelVerticalAlignment.Center;


            //workSheet.Cells[1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            //workSheet.Cells[1,1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Green);
            using (var memoryStream = new MemoryStream())
            {
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=Contact.xlsx");
                excel.SaveAs(memoryStream);
                memoryStream.WriteTo(Response.OutputStream);
                Response.Flush();
                Response.End();
            }
        }
    }
}