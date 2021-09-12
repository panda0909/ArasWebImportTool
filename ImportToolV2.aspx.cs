using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aras.IOM;

namespace ArasWebImportTool
{
    public partial class ImportToolV2 : System.Web.UI.Page
    {
        Innovator inn;
        string user_name = "";
        string dbname = "";
        
        //中文欄位
        string STK_ID = "主件";
        string Mat_ID = "元件";
        string Mat_No = "序號";
        string Mat_Qty = "組成用量";
        string Position = "插件位置";
        string Alternative = "Alternative";
        string Shrink_Rate = "損耗率";
        string Remark = "備註";
        protected void Page_Load(object sender, EventArgs e)
        {
            user_name = Request.QueryString["user_name"];
            dbname = Request.QueryString["dbname"];
            
            lblLog.Text = "User Name = " + user_name + ",DataBase = " + dbname + "</br>";
            
        }

        protected void btnImport_Click(object sender, EventArgs e)
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

            //登入Aras
            Login();
            if (inn == null)
            {
                lblLog.Text += "登入失敗<br>";
                return;
            }

            string write_mode = DropDownList_Model.SelectedValue; //all完整刪除重建,diff差異匯入
            string txtPath = saveResultPath;

            DataTable dtBOM = null;
            
            dtBOM = ExcelLib.ReadExcelToDataTable(txtPath, "BOM_Import");
            if(dtBOM == null)
            {
                dtBOM = ExcelLib.ReadExcelToDataTable(txtPath, "BOM_Import_ENG");
                mode = "en";
            }
            if(dtBOM == null)
            {
                lblLog.Text += "無法找到分頁 BOM_Import或 BOM_Import_ENG。Can't find the sheet name.<br>";
                return;
            }
            dtBOM.Columns.Add("Result");

            //英文版欄位
            if(mode == "en")
            {
                STK_ID = "STK_ID";
                Mat_ID = "Mat_ID";
                Mat_No = "Mat_No";
                Mat_Qty = "Mat_Qty";
                Position = "Position";
                Alternative = "Alternative";
                Shrink_Rate = "Shrink_Rate";
                Remark = "Remark";
            }
            var RootParts = from t in dtBOM.AsEnumerable()
                            group t by new { root_part = t.Field<string>(STK_ID) } into m
                            select new
                            {
                                rootPart = m.Key.root_part
                            };
            foreach (var root in RootParts)
            {
                List<Item> rootBOMsAras = GetItemBOMByRoot(root.rootPart);//取得系統BOM表
                if (write_mode == "all")
                {
                    //完整刪除BOM
                    foreach(Item arasBOMItem in rootBOMsAras)
                    {
                        DeltePartBOM(arasBOMItem.getID());
                    }
                }
                DataRow[] rootBOMsFindExcel = dtBOM.Select(STK_ID+" = '" + root.rootPart + "'");//ExcelBOM
                if (write_mode == "diff2")
                {
                    //先用Aras查每一個主料，如果Excel沒有，就刪除
                    foreach (Item arasBOMItem in rootBOMsAras)
                    {
                        string arasSubBOM_ItemNumber = arasBOMItem.getRelatedItem().getProperty("item_number", "");
                        if (arasSubBOM_ItemNumber != "")
                        {
                            var findRow = rootBOMsFindExcel.Where(r => r[Mat_ID].ToString().Trim() == arasSubBOM_ItemNumber 
                            && (r[Alternative].ToString().Trim()=="" || r[Alternative].ToString().Trim() == "R1")).FirstOrDefault();
                            if(findRow == null)
                            {
                                DeltePartBOM(arasBOMItem.getID());
                            }
                        }
                    }
                }
                foreach (DataRow row in rootBOMsFindExcel)
                {
                    string rootPart = row[STK_ID].ToString().Trim();
                    string subPart = row[Mat_ID].ToString().Trim();
                    string sort_order = row[Mat_No].ToString().Trim();
                    string quantity = row[Mat_Qty].ToString().Trim();
                    string reference_designator = row[Position].ToString().Trim();
                    string alternative = row[Alternative].ToString().Trim();
                    string attrition_rate = row[Shrink_Rate].ToString().Trim();
                    string note = row[Remark].ToString().Trim();

                    Item currentAffectItem = inn.newItem("Part", "get");
                    currentAffectItem.setProperty("item_number", rootPart);
                    currentAffectItem.setProperty("is_current", "1");
                    currentAffectItem = currentAffectItem.apply();
                    if (currentAffectItem.isError())
                    {
                        row["Result"] = rootPart+"父階料號不存在";
                        continue;
                    }
                    else
                    {
                        if(currentAffectItem.getProperty("state","") != "Preliminary")
                        {
                            row["Result"] = rootPart + "父階料號沒有可編輯權限，請至執行變更在操作";
                            continue;
                        }
                    }

                    //主料變更
                    if (alternative == "" || alternative == "R1")
                    {
                        string chkAml = CheckBOMItemAML(rootPart, subPart);

                        Item chkItm = inn.applyAML(chkAml);
                        if (chkItm.isError())
                        {
                            //子階料號不存在，新增
                            string chkOrder = CheckBOMItemAMLBySortOrder(rootPart, sort_order);
                            Item chkItmOrder = inn.applyAML(chkOrder);
                            if (chkItmOrder.isError() == false)
                            {
                                //把舊序號刪除
                                if (chkItmOrder.getItemByIndex(0).getRelatedItem().getProperty("item_number", "") != subPart)
                                {
                                    DeltePartBOM(chkItmOrder.getItemByIndex(0).getID());
                                    string importAml_Add = BOMAddAML(row);
                                    Item resultItem = inn.applyAML(importAml_Add);
                                    if (resultItem.isError())
                                    {
                                        row["Result"] = "Error:" + resultItem.getErrorString();
                                    }
                                    else
                                    {
                                        row["Result"] = "success";
                                    }
                                }
                                else
                                {
                                    string importAml_Merge = BOMMergeAML(row, chkItmOrder.getID());
                                    Item resultItem = inn.applyAML(importAml_Merge);
                                    if (resultItem.isError())
                                    {
                                        row["Result"] = "Error:" + resultItem.getErrorString();
                                    }
                                    else
                                    {
                                        row["Result"] = "success";
                                    }
                                }
                            }
                            else
                            {
                                string importAml_Add = BOMAddAML(row);
                                Item resultItem = inn.applyAML(importAml_Add);
                                if (resultItem.isError())
                                {
                                    row["Result"] = "Error:" + resultItem.getErrorString();
                                }
                                else
                                {
                                    row["Result"] = "success";
                                }
                            }
                        }
                        else
                        {
                            //子階料號存在，修改
                            string chkOrder = CheckBOMItemAMLBySortOrder(rootPart, sort_order);
                            Item chkItmOrder = inn.applyAML(chkOrder);
                            if (chkItmOrder.isError() == false)
                            {
                                //把舊序號刪除
                                if(chkItmOrder.getItemByIndex(0).getRelatedItem().getProperty("item_number","")!=subPart)
                                    DeltePartBOM(chkItmOrder.getItemByIndex(0).getID());
                            }
                            string importAml_Merge = BOMMergeAML(row, chkItm.getID());
                            Item resultItem = inn.applyAML(importAml_Merge);
                            if (resultItem.isError())
                            {
                                row["Result"] = "Error:" + resultItem.getErrorString();
                            }
                            else
                            {
                                row["Result"] = "success";
                            }
                        }
                    }
                    else if (alternative == "D") //表示刪除
                    {
                        string subsititue = row[Mat_ID].ToString().Trim();  //替代料號
                        string find_number = "";
                        rootPart = row[STK_ID].ToString().Trim();
                        sort_order = row[Mat_No].ToString().Trim();
                        DataRow[] subPartFind = dtBOM.Select(STK_ID + " = '" + rootPart + "' and " + Mat_No + " = '" + sort_order + "' and Alternative='R1'");
                        if (subPartFind.Count() > 0)
                        {
                            subPart = subPartFind[0][Mat_ID].ToString().Trim(); //主件
                            
                            string chkAml = CheckAlternativeAML(rootPart, subPart, subsititue, sort_order);
                            Item chkItm = inn.applyAML(chkAml);

                            if (chkItm.isError()==false)
                            {
                                string chkOrder = CheckAlternativeAMLByAlterOrder(rootPart, subPart, subsititue);
                                Item chkItmOrder = inn.applyAML(chkOrder);
                                if (chkItmOrder.isError() == false)
                                {
                                    Item delResult = DelteBOMSubsitute(chkItmOrder.getItemByIndex(0).getID());
                                    if (delResult.isError() == false)
                                    {
                                        row["Result"] = "success";

                                        //更新是否有替代件欄位
                                        int is_alter = CheckPartBOM_IsAlterNative(rootPart, subPart, sort_order);
                                        string chkBOMAML = CheckBOMItemAML(rootPart, subPart);
                                        Item chkBOM = inn.applyAML(chkBOMAML);
                                        if (chkBOM.isError() == false)
                                        {
                                            string sql = "update innovator.Part_BOM set cn_isalternate = '" + is_alter + "' where id = '" + chkBOM.getID() + "'";
                                            Item qry = inn.applySQL(sql);
                                            inn.applyAML(ApplyEditPartBOM(chkBOM.getID()));
                                        }
                                        
                                    }
                                    else
                                    {
                                        row["Result"] = "Error:" + delResult.getErrorString();
                                    }
                                }
                            }
                        }
                        else
                        {
                            //純主件刪除
                            string chkOrder = CheckBOMItemAML(rootPart, subPart);
                            Item chkItmOrder = inn.applyAML(chkOrder);
                            if (chkItmOrder.isError() == false)
                            {
                                if (chkItmOrder.getItemByIndex(0).getRelatedItem().getProperty("item_number", "") == subPart)
                                {
                                    Item delResult = DeltePartBOM(chkItmOrder.getItemByIndex(0).getID());
                                    if (delResult.isError() == false)
                                    {
                                        row["Result"] = "success";
                                    }
                                    else
                                    {
                                        row["Result"] = "Error:" + delResult.getErrorString();
                                    }
                                    
                                }
                            }
                        }
                    }
                    else
                    {
                        //替代料變更
                        rootPart = row[STK_ID].ToString().Trim();
                        subPart = "";
                        string subsititue = row[Mat_ID].ToString().Trim();  //替代料號
                        sort_order = row[Mat_No].ToString().Trim();
                        string find_number = "";
                        quantity = row[Mat_Qty].ToString().Trim();
                        attrition_rate = row[Shrink_Rate].ToString().Trim();
                        note = row[Remark].ToString().Trim();

                        //是否有R1主料且序號相同
                        DataRow[] subPartFind = dtBOM.Select(STK_ID+" = '" + rootPart + "' and "+Mat_No+" = '" + sort_order + "' and Alternative='R1'");
                        if (subPartFind.Count() > 0)
                        {
                            //R1料號
                            subPart = subPartFind[0][Mat_ID].ToString().Trim();
                        }
                        if (alternative != "")
                        {
                            string alternativeNumber = alternative.Substring(1, alternative.Length - 1);
                            int aNumber = int.Parse(alternativeNumber) - 1;
                            find_number = aNumber.ToString();
                        }

                        string chkAml = CheckAlternativeAML(rootPart, subPart, subsititue, sort_order);
                        Item chkItm = inn.applyAML(chkAml);
                        if (chkItm.isError())
                        {
                            string chkOrder = CheckAlternativeAMLByAlterOrder(rootPart, subPart, subsititue, sort_order, find_number);
                            Item chkItmOrder = inn.applyAML(chkOrder);
                            if (chkItmOrder.isError() == false)
                            {
                                //把舊序號刪除
                                DelteBOMSubsitute(chkItmOrder.getItemByIndex(0).getID());
                                string importAml_AddR = AlternativeBOMAddAML(rootPart, subPart, subsititue, quantity, sort_order, find_number, attrition_rate, note);
                                Item resultItem = inn.applyAML(importAml_AddR);
                                if (resultItem.isError())
                                {
                                    row["Result"] = "Error:" + resultItem.getErrorString();
                                }
                                else
                                {
                                    row["Result"] = "success";
                                }
                            }
                            else
                            {
                                string importAml_AddR = AlternativeBOMAddAML(rootPart, subPart, subsititue, quantity, sort_order, find_number, attrition_rate, note);
                                Item resultItem = inn.applyAML(importAml_AddR);
                                if (resultItem.isError())
                                {
                                    row["Result"] = "Error:" + resultItem.getErrorString();
                                }
                                else
                                {
                                    row["Result"] = "success";
                                    //修改紀錄
                                    Item thisBom = inn.applyAML(CheckBOMItemAML(rootPart, subPart));
                                    if (thisBom.isError() == false)
                                    {
                                        inn.applyAML(ApplyEditPartBOM(thisBom.getID()));
                                    }
                                }
                            }

                        }
                        else
                        {
                            string chkOrder = CheckAlternativeAMLByAlterOrder(rootPart, subPart, subsititue, sort_order, find_number);
                            Item chkItmOrder = inn.applyAML(chkOrder);
                            if (chkItmOrder.isError() == false)
                            {
                                //把舊序號刪除
                                if (chkItmOrder.getItemByIndex(0).getRelatedItem().getProperty("item_number", "") != subsititue)
                                    DelteBOMSubsitute(chkItmOrder.getItemByIndex(0).getID());
                            }
                            string importAml_MergeR = AlternativeBOMMergeAML(rootPart, subPart, subsititue, quantity, sort_order, find_number, attrition_rate, note, chkItm.getID());
                            Item resultItem = inn.applyAML(importAml_MergeR);
                            if (resultItem.isError())
                            {
                                row["Result"] = "Error:" + resultItem.getErrorString();
                            }
                            else
                            {
                                row["Result"] = "success";
                                //修改紀錄
                                Item thisBom = inn.applyAML(CheckBOMItemAML(rootPart, subPart));
                                if (thisBom.isError() == false)
                                {
                                    inn.applyAML(ApplyEditPartBOM(thisBom.getID()));
                                }
                            }
                        }
                    }
                }
            }
            gvBOM.DataSource = dtBOM;
            gvBOM.DataBind();
        }
        public List<Item> GetItemBOMByRoot(string item_number)
        {
            List<Item> result = new List<Item>();
            string aml = @"<AML>
                                  <Item action='get' type='Part BOM'>
                                    <source_id>
                                        <Item type='Part' action='get'>
                                        <item_number>@source_id</item_number>
                                        <is_current>1</is_current>
                                        </Item>
                                    </source_id>
                                    </Item>
                                </AML>";
            aml = aml.Replace("@source_id", item_number);
            Item itm = inn.applyAML(aml);
            if (itm.isError() == false)
            {
                for (int i = 0; i < itm.getItemCount(); i++)
                {
                    Item itmPB = itm.getItemByIndex(i);
                    result.Add(itmPB);
                }
            }
            return result;
        }
        public Item DeltePartBOM(string id)
        {
            string aml = @"<AML><Item type='Part BOM' action='delete' id='{0}'></Item></AML>";
            aml = string.Format(aml, id);
            
            return inn.applyAML(aml);
        }
        public Item DelteBOMSubsitute(string id)
        {
            //string sql = @"delete from innovator.bom_substitute where id = '{0}'";
            //sql = string.Format(sql, id);
            string aml = @"<AML><Item type='BOM Substitute' action='delete' id='{0}'></Item></AML>";
            aml = string.Format(aml, id);
            Item itm = inn.applyAML(aml);
            return itm;
        }
        public string CheckBOMItemAML(string rootPart, string subPart)
        {
            string chkAml = @"<AML>
                                  <Item action='get' type='Part BOM'>
                                    <source_id>
                                        <Item type='Part' action='get'>
                                        <item_number>@source_id</item_number>
                                        <is_current>1</is_current>
                                        </Item>
                                    </source_id>
                                    <related_id>
                                    <Item type='Part' action='get'>
                                        <item_number>@related_id</item_number>
                                    </Item>
                                    </related_id>
                                    </Item>
                                </AML>";
            chkAml = chkAml.Replace("@source_id", rootPart);
            chkAml = chkAml.Replace("@related_id", subPart);
            return chkAml;
        }
        public string CheckBOMItemAMLBySortOrder(string rootPart, string sort_order)
        {
            string chkAml = @"<AML>
                                  <Item action='get' type='Part BOM'>
                                    <source_id>
                                        <Item type='Part' action='get'>
                                        <item_number>@source_id</item_number>
                                        <is_current>1</is_current>
                                        </Item>
                                    </source_id>
                                    <sort_order>@sort_order</sort_order>
                                    </Item>
                                </AML>";
            chkAml = chkAml.Replace("@source_id", rootPart);
            chkAml = chkAml.Replace("@sort_order", sort_order);
            return chkAml;
        }
        public string CheckBOMItemAMLBySortOrder(string rootPart)
        {
            string chkAml = @"<AML>
                                  <Item action='get' type='Part BOM'>
                                    <source_id>
                                        <Item type='Part' action='get'>
                                        <item_number>@source_id</item_number>
                                        <is_current>1</is_current>
                                        </Item>
                                    </source_id>
                                    </Item>
                                </AML>";
            chkAml = chkAml.Replace("@source_id", rootPart);
            return chkAml;
        }
        public int CheckPartBOM_IsAlterNative(string rootPart,string subPart,string sort_order)
        {
            string chkAml = @"<AML>
                                    <Item action='get' type='BOM Substitute'>
                                    <source_id>
                                        <Item type='Part BOM' action='get'>
                                        <sort_order>@sort_order</sort_order>
                                        <source_id>
                                            <Item type='Part' action='get'>
                                            <item_number>@rootPart</item_number>
                                            <is_current>1</is_current>
                                            </Item>
                                        </source_id>
                                        <related_id>
                                            <Item type='Part' action='get'>
                                            <item_number>@subPart</item_number>
                                            </Item>
                                        </related_id>
                                        </Item>
                                    </source_id>
                                    </Item>
                                </AML>";
            chkAml = chkAml.Replace("@rootPart", rootPart);
            chkAml = chkAml.Replace("@subPart", subPart);
            chkAml = chkAml.Replace("@sort_order", sort_order);
            Item itm = inn.applyAML(chkAml);
            if (itm.isError() == false)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }
        public string CheckAlternativeAML(string rootPart, string subPart, string subsititue, string sort_order)
        {
            string chkAml = @"<AML>
                                    <Item action='get' type='BOM Substitute'>
                                    <source_id>
                                        <Item type='Part BOM' action='get'>
                                        <sort_order>@sort_order</sort_order>
                                        <source_id>
                                            <Item type='Part' action='get'>
                                            <item_number>@rootPart</item_number>
                                            <is_current>1</is_current>
                                            </Item>
                                        </source_id>
                                        <related_id>
                                            <Item type='Part' action='get'>
                                            <item_number>@subPart</item_number>
                                            </Item>
                                        </related_id>
                                        </Item>
                                    </source_id>
                                    <related_id>
                                        <Item type='Part' action='get'>
                                        <item_number>@subsititue</item_number>
                                        </Item>
                                    </related_id>
                                    </Item>
                                </AML>";

            chkAml = chkAml.Replace("@rootPart", rootPart);
            chkAml = chkAml.Replace("@subPart", subPart);
            chkAml = chkAml.Replace("@subsititue", subsititue);
            chkAml = chkAml.Replace("@sort_order", sort_order);
            return chkAml;
        }
        public string CheckAlternativeAML(string rootPart, string subPart, string subsititue)
        {
            string chkAml = @"<AML>
                                    <Item action='get' type='BOM Substitute'>
                                    <source_id>
                                        <Item type='Part BOM' action='get'>
                                        <source_id>
                                            <Item type='Part' action='get'>
                                            <item_number>@rootPart</item_number>
                                            <is_current>1</is_current>
                                            </Item>
                                        </source_id>
                                        <related_id>
                                            <Item type='Part' action='get'>
                                            <item_number>@subPart</item_number>
                                            </Item>
                                        </related_id>
                                        </Item>
                                    </source_id>
                                    <related_id>
                                        <Item type='Part' action='get'>
                                        <item_number>@subsititue</item_number>
                                        </Item>
                                    </related_id>
                                    </Item>
                                </AML>";

            chkAml = chkAml.Replace("@rootPart", rootPart);
            chkAml = chkAml.Replace("@subPart", subPart);
            chkAml = chkAml.Replace("@subsititue", subsititue);
            return chkAml;
        }
        public string CheckAlternativeAMLByAlterOrder(string rootPart, string subPart, string subsititue, string sort_order, string alter_order)
        {
            string chkAml = @"<AML>
                                    <Item action='get' type='BOM Substitute'>
                                    <source_id>
                                        <Item type='Part BOM' action='get'>
                                        <sort_order>@sort_order</sort_order>
                                        <source_id>
                                            <Item type='Part' action='get'>
                                            <item_number>@rootPart</item_number>
                                            <is_current>1</is_current>
                                            </Item>
                                        </source_id>
                                        <related_id>
                                            <Item type='Part' action='get'>
                                            <item_number>@subPart</item_number>
                                            </Item>
                                        </related_id>
                                        </Item>
                                    </source_id>
                                    <cn_sort_order>@cn_sort_order</cn_sort_order>
                                    </Item>
                                </AML>";

            chkAml = chkAml.Replace("@rootPart", rootPart);
            chkAml = chkAml.Replace("@subPart", subPart);
            chkAml = chkAml.Replace("@subsititue", subsititue);
            chkAml = chkAml.Replace("@sort_order", sort_order);
            chkAml = chkAml.Replace("@cn_sort_order", alter_order);
            return chkAml;
        }
        public string CheckAlternativeAMLByAlterOrder(string rootPart, string subPart, string subsititue)
        {
            string chkAml = @"<AML>
                                    <Item action='get' type='BOM Substitute'>
                                    <source_id>
                                        <Item type='Part BOM' action='get'>
                                        <source_id>
                                            <Item type='Part' action='get'>
                                            <item_number>@rootPart</item_number>
                                            <is_current>1</is_current>
                                            </Item>
                                        </source_id>
                                        <related_id>
                                            <Item type='Part' action='get'>
                                            <item_number>@subPart</item_number>
                                            </Item>
                                        </related_id>
                                        </Item>
                                    </source_id>
                                    </Item>
                                </AML>";

            chkAml = chkAml.Replace("@rootPart", rootPart);
            chkAml = chkAml.Replace("@subPart", subPart);
            chkAml = chkAml.Replace("@subsititue", subsititue);
            return chkAml;
        }
        public string BOMAddAML(DataRow row)
        {
            string rootPart = row[STK_ID].ToString().Trim();
            string subPart = row[Mat_ID].ToString().Trim();
            string sort_order = row[Mat_No].ToString().Trim();
            string quantity = row[Mat_Qty].ToString().Trim();
            string reference_designator = row[Position].ToString().Trim();
            string alternative = row[Alternative].ToString().Trim();
            string attrition_rate = row[Shrink_Rate].ToString().Trim();
            string note = row[Remark].ToString().Trim();
            string importAml_Add = @"<AML>
                                  <Item action='add' type='Part BOM'>
                                    <source_id>
                                      <Item type='Part' action='get'>
                                        <item_number>@source_id</item_number>
                                        <is_current>1</is_current>
                                      </Item>
                                    </source_id>
                                    <related_id>
                                      <Item type='Part' action='get'>
                                        <item_number>@related_id</item_number>
                                      </Item>
                                    </related_id>
                                    <quantity>@quantity</quantity>
                                    <sort_order>@sort_order</sort_order>
                                    <ch_order>@ch_order</ch_order>
                                    <cn_bom_note>@cn_bom_note</cn_bom_note>
                                    <cn_attrition_rate>@cn_attrition_rate</cn_attrition_rate>
                                    <reference_designator>@reference_designator</reference_designator>
                                  </Item>
                                </AML>";
            importAml_Add = importAml_Add.Replace("@source_id", rootPart);
            importAml_Add = importAml_Add.Replace("@related_id", subPart);
            importAml_Add = importAml_Add.Replace("@quantity", quantity);
            importAml_Add = importAml_Add.Replace("@sort_order", sort_order);
            importAml_Add = importAml_Add.Replace("@cn_bom_note", note.Replace("<", "&lt;").Replace(">", "&gt;"));
            importAml_Add = importAml_Add.Replace("@reference_designator", reference_designator.Replace("<", "&lt;").Replace(">", "&gt;"));
            importAml_Add = importAml_Add.Replace("@cn_attrition_rate", attrition_rate);
            importAml_Add = importAml_Add.Replace("@ch_order", GetMaxChOrder(rootPart));
            return importAml_Add;
        }
        public string BOMMergeAML(DataRow row, string bom_id)
        {
            string rootPart = row[STK_ID].ToString().Trim();
            string subPart = row[Mat_ID].ToString().Trim();
            string sort_order = row[Mat_No].ToString().Trim();
            string quantity = row[Mat_Qty].ToString().Trim();
            string reference_designator = row[Position].ToString().Trim();
            string alternative = row[Alternative].ToString().Trim();
            string attrition_rate = row[Shrink_Rate].ToString().Trim();
            string note = row[Remark].ToString().Trim();
            string importAml_Add = @"<AML>
                                  <Item action='edit' type='Part BOM' id='@id'>
                                    <source_id>
                                      <Item type='Part' action='get'>
                                        <item_number>@source_id</item_number>
                                        <is_current>1</is_current>
                                      </Item>
                                    </source_id>
                                    <related_id>
                                      <Item type='Part' action='get'>
                                        <item_number>@related_id</item_number>
                                      </Item>
                                    </related_id>
                                    <quantity>@quantity</quantity>
                                    <sort_order>@sort_order</sort_order>
                                    <ch_order>@ch_order</ch_order>
                                    <cn_bom_note>@cn_bom_note</cn_bom_note>
                                    <cn_attrition_rate>@cn_attrition_rate</cn_attrition_rate>
                                    <reference_designator>@reference_designator</reference_designator>
                                  </Item>
                                </AML>";
            importAml_Add = importAml_Add.Replace("@id", bom_id);
            importAml_Add = importAml_Add.Replace("@source_id", rootPart);
            importAml_Add = importAml_Add.Replace("@related_id", subPart);
            importAml_Add = importAml_Add.Replace("@quantity", quantity);
            importAml_Add = importAml_Add.Replace("@sort_order", sort_order);
            importAml_Add = importAml_Add.Replace("@cn_bom_note", note.Replace("<", "&lt;").Replace(">", "&gt;"));
            importAml_Add = importAml_Add.Replace("@reference_designator", reference_designator.Replace("<", "&lt;").Replace(">", "&gt;"));
            importAml_Add = importAml_Add.Replace("@cn_attrition_rate", attrition_rate);
            importAml_Add = importAml_Add.Replace("@ch_order", GetMaxChOrder(rootPart));
            return importAml_Add;
        }
        public string AlternativeBOMAddAML(string rootPart, string subPart, string subsititue, string quantity, string sort_order, string find_number, string attrition_rate, string cn_note)
        {
            string importAml_AddR = @"<AML>
                                    <Item action='add' type='BOM Substitute'>
                                    <source_id>
                                        <Item type='Part BOM' action='get'>
                                        <sort_order>@sort_order</sort_order>
                                        <source_id>
                                            <Item type='Part' action='get'>
                                            <item_number>@rootPart</item_number>
                                            <is_current>1</is_current>
                                            </Item>
                                        </source_id>
                                        <related_id>
                                            <Item type='Part' action='get'>
                                            <item_number>@subPart</item_number>
                                            </Item>
                                        </related_id>
                                        </Item>
                                    </source_id>
                                    <related_id>
                                        <Item type='Part' action='get'>
                                        <item_number>@subsititue</item_number>
                                        </Item>
                                    </related_id>
                                    <cn_sort_order>@cn_sort_order</cn_sort_order>
                                    <cn_substitute_quantity>@cn_substitute_quantity</cn_substitute_quantity>
                                    <cn_substitute_shrinkrate>@cn_attrition_rate</cn_substitute_shrinkrate>
                                    </Item>
                                </AML>";
            importAml_AddR = importAml_AddR.Replace("@rootPart", rootPart);
            importAml_AddR = importAml_AddR.Replace("@subPart", subPart);
            importAml_AddR = importAml_AddR.Replace("@subsititue", subsititue);
            importAml_AddR = importAml_AddR.Replace("@cn_substitute_quantity", quantity);
            importAml_AddR = importAml_AddR.Replace("@sort_order", sort_order);
            importAml_AddR = importAml_AddR.Replace("@cn_sort_order", find_number);
            importAml_AddR = importAml_AddR.Replace("@cn_attrition_rate", attrition_rate);
            importAml_AddR = importAml_AddR.Replace("@cn_note", cn_note);
            return importAml_AddR;
        }
        public string AlternativeBOMMergeAML(string rootPart, string subPart, string subsititue, string quantity, string sort_order, string find_number, string attrition_rate, string cn_note, string alternative_id)
        {
            string importAml_AddR = @"<AML>
                                    <Item action='edit' type='BOM Substitute' id='@alternative_id'>
                                    <source_id>
                                        <Item type='Part BOM' action='get'>
                                        <sort_order>@sort_order</sort_order>
                                        <source_id>
                                            <Item type='Part' action='get'>
                                            <item_number>@rootPart</item_number>
                                            <is_current>1</is_current>
                                            </Item>
                                        </source_id>
                                        <related_id>
                                            <Item type='Part' action='get'>
                                            <item_number>@subPart</item_number>
                                            </Item>
                                        </related_id>
                                        </Item>
                                    </source_id>
                                    <related_id>
                                        <Item type='Part' action='get'>
                                        <item_number>@subsititue</item_number>
                                        </Item>
                                    </related_id>
                                    <cn_sort_order>@cn_sort_order</cn_sort_order>
                                    <cn_substitute_quantity>@cn_substitute_quantity</cn_substitute_quantity>
                                    <cn_substitute_shrinkrate>@cn_attrition_rate</cn_substitute_shrinkrate>
                                    </Item>
                                </AML>";
            importAml_AddR = importAml_AddR.Replace("@alternative_id", alternative_id);
            importAml_AddR = importAml_AddR.Replace("@rootPart", rootPart);
            importAml_AddR = importAml_AddR.Replace("@subPart", subPart);
            importAml_AddR = importAml_AddR.Replace("@subsititue", subsititue);
            importAml_AddR = importAml_AddR.Replace("@cn_substitute_quantity", quantity);
            importAml_AddR = importAml_AddR.Replace("@sort_order", sort_order);
            importAml_AddR = importAml_AddR.Replace("@cn_sort_order", find_number);
            importAml_AddR = importAml_AddR.Replace("@cn_attrition_rate", attrition_rate);
            importAml_AddR = importAml_AddR.Replace("@cn_note", cn_note);
            return importAml_AddR;
        }
        private string GetMaxChOrder(string source_part)
        {
            int max_order = 1;

            string chkAml = @"<AML>
                                  <Item action='get' type='Part BOM'>
                                    <source_id>
                                        <Item type='Part' action='get'>
                                        <item_number>@source_id</item_number>
                                        <is_current>1</is_current>
                                        </Item>
                                    </source_id>
                                    </Item>
                                </AML>";
            chkAml = chkAml.Replace("@source_id", source_part);
            Item boms = inn.applyAML(chkAml);

            if (!boms.isError())
            {
                for (int i = 0; i < boms.getItemCount(); i++)
                {
                    Item bom = boms.getItemByIndex(i);
                    string ch_order = bom.getProperty("ch_order", "0");
                    int ch_order_int = int.Parse(ch_order);
                    if (max_order < ch_order_int)
                    {
                        max_order = ch_order_int;
                    }
                }
                max_order = max_order + 1;
            }
            return max_order.ToString();
        }
        private void Log(string msg)
        {
            Item log = inn.newItem("JPC_Method_Log", "add");
            log.setProperty("jpc_run_method", "WebImportTool");
            log.setProperty("jpc_method_event", "Web");
            log.setProperty("jpc_log", msg);
            log = log.apply();
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
        private string ApplyEditPartBOM(string id)
        {
            string aml = @"<AML>
                          <Item action='edit' type='Part BOM' id='{0}'>
                          </Item>
                        </AML>";
            aml = string.Format(aml, id);
            return aml;
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
    }
}