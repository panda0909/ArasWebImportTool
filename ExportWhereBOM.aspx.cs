using Aras.IOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ArasWebImportTool
{
    public partial class ExportWhereBOM : System.Web.UI.Page
    {

        Innovator inn;
        string user_name = "";
        string dbname = "";
        string Col_STK_ID = "主件";
        string Col_Mat_ID = "元件";
        string Col_Mat_No = "序號";
        string Col_Mat_Qty = "組成用量";
        string Col_Position = "插件位置";
        string Col_Alternative = "Alternative";
        string Col_Shrink_Rate = "損耗率";
        string Col_Remark = "備註";
        string Col_RelationAlternative = "Related Alt(STK_ID,Unit,QTY,SRate)";
        protected void Page_Load(object sender, EventArgs e)
        {
            user_name = Request.QueryString["user_name"];
            dbname = Request.QueryString["dbname"];
            
            lblLog.Text = "User Name = " + user_name + ",DataBase = " + dbname + "</br>";
        }

        protected void btnExport_Click(object sender, EventArgs e)
        {
            string saveDir = @"\Files\";
            string appDir = Request.PhysicalApplicationPath;
            string savePath = appDir + saveDir;
            string mode = "zt";
            if (FileUpload_Excel.HasFile == false)
            {
                lblLog.Text = "請上傳檔案";
                return;
            }
            string saveResultPath = savePath + FileUpload_Excel.FileName;
            FileUpload_Excel.SaveAs(saveResultPath);

            if (File.Exists(saveResultPath) == false)
            {
                lblLog.Text = "檔案上傳失敗<br>";
                return;
            }

            Login();
            if (inn == null)
            {
                lblLog.Text += "登入失敗<br>";
                return;
            }

            DataTable dtBOM = null;
            string txtPath = saveResultPath;
            string sheetname = "BOM_Import";
            dtBOM = ExcelLib.ReadExcelToDataTable(txtPath, "BOM_Import");
            if (dtBOM == null)
            {
                sheetname = "BOM_Import_ENG";
                dtBOM = ExcelLib.ReadExcelToDataTable(txtPath, "BOM_Import_ENG");
                mode = "en";
            }
            if (dtBOM == null)
            {
                lblLog.Text += "無法找到分頁 BOM_Import或 BOM_Import_ENG。Can't find the sheet name.<br>";
                return;
            }

            if (mode == "en")
            {
                Col_STK_ID = "STK_ID";
                Col_Mat_ID = "Mat_ID";
                Col_Mat_No = "Mat_No";
                Col_Mat_Qty = "Mat_Qty";
                Col_Position = "Position";
                Col_Alternative = "Alternative";
                Col_Shrink_Rate = "Shrink_Rate";
                Col_Remark = "Remark";
            }
            dtBOM.Columns.Add(Col_RelationAlternative);
            foreach(DataRow row in dtBOM.Rows)
            {
                string STK_ID = row[Col_STK_ID].ToString();
                string Mat_ID = row[Col_Mat_ID].ToString();
                string Mat_No = row[Col_Mat_No].ToString();
                string Mat_Qty = row[Col_Mat_Qty].ToString();
                string Position = row[Col_Position].ToString();
                string Alternative = row[Col_Alternative].ToString();
                string Shrink_Rate = row[Col_Shrink_Rate].ToString();
                string Remark = row[Col_Remark].ToString();

                string aml = @"<AML>
                              <Item action='get' type='Part BOM'>
                                <source_id>
                                  <Item type='Part' action='get'>
                                    <item_number>{0}</item_number>
                                    <is_current>1</is_current>
                                  </Item>
                                </source_id>
                                <related_id>
                                <Item type='Part' action='get'>
                                  <item_number>{1}</item_number>
                                </Item>
                                </related_id>
                              </Item>
                            </AML>";
                aml = string.Format(aml, STK_ID, Mat_ID);
                Item itmBOM = inn.applyAML(aml);
                if (itmBOM.getItemCount() > 0)
                {
                    row[Col_Mat_No] = itmBOM.getItemByIndex(0).getProperty("sort_order", "");
                    row[Col_Mat_Qty] = itmBOM.getItemByIndex(0).getProperty("quantity", "");
                    row[Col_Position] = itmBOM.getItemByIndex(0).getProperty("reference_designator", "");
                    row[Col_Alternative] = "";
                    row[Col_Remark] = itmBOM.getItemByIndex(0).getProperty("cn_bom_note", "");
                    row[Col_Shrink_Rate] = itmBOM.getItemByIndex(0).getProperty("cn_attrition_rate", "");

                    string subAML = @"<AML>
                                      <Item action='get' type='BOM Substitute' select='cn_sort_order,cn_substitute_unit,cn_substitute_quantity,cn_substitute_shrinkrate,related_id(item_number)'>
                                        <source_id>
                                          <Item type='Part BOM' action='get'>
                                            <source_id>
                                              <Item type='Part' action='get'>
                                                <item_number>{0}</item_number>
                                                <is_current>1</is_current>
                                              </Item>
                                            </source_id>
                                            <related_id>
                                            <Item type='Part' action='get'>
                                              <item_number>{1}</item_number>
                                            </Item>
                                            </related_id>
                                          </Item>
                                        </source_id>
                                      </Item>
                                    </AML>";
                    subAML = string.Format(subAML, STK_ID, Mat_ID);
                    Item subAltItms = inn.applyAML(subAML);

                    if (subAltItms.isError()==false)
                    {
                        for(int i=0;i< subAltItms.getItemCount(); i++)
                        {
                            Item itmSubAlt = subAltItms.getItemByIndex(i);
                            string altStr = itmSubAlt.getRelatedItem().getProperty("item_number","")+"_";
                            altStr += itmSubAlt.getProperty("cn_substitute_unit", "") + "_";
                            altStr += itmSubAlt.getProperty("cn_substitute_quantity", "") + "_";
                            altStr += itmSubAlt.getProperty("cn_substitute_shrinkrate", "") + "_";

                            int rAlt = int.Parse(itmSubAlt.getProperty("cn_sort_order", "1"));
                            rAlt += 1;
                            altStr += "R"+rAlt.ToString();
                            if(row[Col_RelationAlternative].ToString().Trim() == "")
                            {
                                row[Col_RelationAlternative] = altStr;
                            }
                            else
                            {
                                row[Col_RelationAlternative] = row[Col_Shrink_Rate].ToString() + "\r\n" + altStr;
                            }
                            
                        }
                    }
                }
            }

            string result_path = ExcelLib.SaveExcelFromDataTable(saveResultPath, dtBOM, sheetname);
            if (xDownload(result_path, "BOM_Import.xlsx"))
            {

            }
        }

        protected void btnDownloadTemplate_Click(object sender, EventArgs e)
        {
            string docupath = Request.PhysicalApplicationPath;
            if (xDownload(docupath + "Files\\BOM_Template2.xlsx", "BOM_Template.xlsx"))
            {

            }
            else
            {

            }
        }

        protected void btnDownloadTemplate_EN_Click(object sender, EventArgs e)
        {
            string docupath = Request.PhysicalApplicationPath;
            if (xDownload(docupath + "Files\\BOM_Template_ENG.xlsx", "BOM_Template_ENG.xlsx"))
            {

            }
            else
            {

            }
        }

        public bool xDownload(string xFile, string out_file)
        //xFile 路徑+檔案, 設定另存的檔名
        {
            if (File.Exists(xFile))
            {
                try
                {
                    FileInfo xpath_file = new FileInfo(xFile);  //要 using System.IO;
                                                                // 將傳入的檔名以 FileInfo 來進行解析（只以字串無法做）
                    System.Web.HttpContext.Current.Response.Clear(); //清除buffer
                    System.Web.HttpContext.Current.Response.ClearHeaders(); //清除 buffer 表頭
                    System.Web.HttpContext.Current.Response.Buffer = false;
                    System.Web.HttpContext.Current.Response.ContentType = "application/octet-stream";
                    // 檔案類型還有下列幾種"application/pdf"、"application/vnd.ms-excel"、"text/xml"、"text/HTML"、"image/JPEG"、"image/GIF"
                    System.Web.HttpContext.Current.Response.AppendHeader("Content-Disposition", "attachment;filename=" + System.Web.HttpUtility.UrlEncode(out_file, System.Text.Encoding.UTF8));
                    // 考慮 utf-8 檔名問題，以 out_file 設定另存的檔名
                    System.Web.HttpContext.Current.Response.AppendHeader("Content-Length", xpath_file.Length.ToString()); //表頭加入檔案大小
                    System.Web.HttpContext.Current.Response.WriteFile(xpath_file.FullName);

                    // 將檔案輸出
                    System.Web.HttpContext.Current.Response.Flush();
                    // 強制 Flush buffer 內容
                    System.Web.HttpContext.Current.Response.End();
                    return true;

                }
                catch (Exception)
                { return false; }

            }
            else
                return false;
        }

        private void Login()
        {
            HttpServerConnection cnx = IomFactory.CreateHttpServerConnection("http://localhost/plm", dbname, "admin", "innovator");
            Item login_result = cnx.Login();
            if (!login_result.isError())
            {
                inn = IomFactory.CreateInnovator(cnx);

            }
        }
    }
}