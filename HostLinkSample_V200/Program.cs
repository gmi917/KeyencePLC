using System;
using System.Collections;
using KHL = KvHostLink.KvHostLink;
using KHST = KvStruct.KvStruct;
using System.Data.SqlClient;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using System.Windows.Forms;
using System.Xml.Linq;
using HostLinkSample_V200;

namespace cs_dll_sample
{
    class Program
    {
        static void Main(string[] args)
        {
            int err = 0;
            int sock = 0;
            err = KHL.KHLInit();
            string connectionString = "Server=192.168.1.9,1433;Database=TEST1;User Id=pwc;Password=PWC@admin;";
            SqlConnection connection = null;
            ArrayList ipqc1ResultList = new ArrayList();
            ArrayList ipqc2ResultList = new ArrayList();
            ArrayList ipqc3ResultList = new ArrayList();
            ArrayList ipqcItemList = new ArrayList();
            float[] rdEM3000float = new float[2];
            float[] rdEM4000float = new float[2];
            float[] rdEM5000float = new float[2];
            string itemNumber = "";
            string maxFormNumber = "";
            string productLineId = "";
            string processCode = "";
            string qcId ="";
            string employeeID = "";//員工工號
            string predictQty = "";//預計產量
            string firstItemDate = "";//首件日期
            string ZF0 = "";//1→現場人員 or 2→品檢人員
            string ZF4 = ""; //1→OP1 or 2→OP2
            string ZF8 = "";//量測數據完成信號(1→三個工件都完成)
            string ZF10 = "";//(預留)
            string ZF16ipqc1 = "";//工件一量測數據信號(1→OK;2→NG)
            string ZF18ipqc2 = "";//工件二量測數據信號(1→OK;2→NG)
            string ZF20ipqc3 = "";//工件三量測數據信號(1→OK;2→NG)
            string mo = "";//製令單別
            string mn = "";//製令單號
            string defective = "不合格";//判定結果
            JYT012 jyt012 = new JYT012();
            IpqcItem primumQCData = new IpqcItem();
            IpqcItem item = new IpqcItem();
            if (err != 0)
            {
                Console.WriteLine(err);
            }

            err = KHL.KHLConnect("10.1.9.106", 8500, 3000, KvHostLink.KHLSockType.SOCK_TCP, ref sock);
            if (err != 0)
            {
                Console.WriteLine(err);
            }
            Console.WriteLine("Connection Established");

            // Interface Buffer
            byte[] readBuf = new byte[2048];
            byte[] writeBuf = new byte[2048];


            Console.WriteLine("--- operate bool start ---");
           
            //讀EM3000工件一量測數據暫存器
            Console.WriteLine("讀EM3000工件一量測數據暫存器(EM3000~EM3098)");

            for (uint i = 0; i < 100; i += 2)
            {
                err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_EM, 3000 + i, 2, readBuf);
                KHST.ByteToFloat(ref rdEM3000float, readBuf, 2, 0);

                if (rdEM3000float[0] == 0)
                {
                    break;
                }
                //將rdEM3000float[0]的值存到ArrayList
                ipqc1ResultList.Add(rdEM3000float[0]);
            }
            Console.WriteLine("Result ipqc1ResultList List:");
            foreach (float value in ipqc1ResultList)
            {
                Console.WriteLine(value);
            }

            //讀EM4000工件二量測數據暫存器
            Console.WriteLine("讀EM4000工件一量測數據暫存器(EM4000~EM4098)");
            for (uint i = 0; i < 100; i += 2)
            {
                err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_EM, 4000 + i, 2, readBuf);
                KHST.ByteToFloat(ref rdEM4000float, readBuf, 2, 0);

                if (rdEM4000float[0] == 0)
                {
                    break;
                }
                //將rdEM4000float[0]的值存到ArrayList
                ipqc2ResultList.Add(rdEM4000float[0]);
            }
            Console.WriteLine("Result ipqc2ResultList List:");
            foreach (float value in ipqc2ResultList)
            {
                Console.WriteLine(value);
            }

            //讀EM5000工件三量測數據暫存器
            Console.WriteLine("讀EM5000工件一量測數據暫存器(EM5000~EM5098)");
            for (uint i = 0; i < 100; i += 2)
            {
                err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_EM, 5000 + i, 2, readBuf);
                KHST.ByteToFloat(ref rdEM5000float, readBuf, 2, 0);

                if (rdEM5000float[0] == 0)
                {
                    break;
                }
                //將rdEM5000float[0]的值存到ArrayList
                ipqc3ResultList.Add(rdEM5000float[0]);
            }
            Console.WriteLine("Result ipqc3ResultList List:");
            foreach (float value in ipqc3ResultList)
            {
                Console.WriteLine(value);
            }

            //讀現場人員 or 品檢人員(ZF0)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 0, 1, readBuf);
            Console.WriteLine("讀現場人員 or 品檢人員(ZF0)");
            byte[] rdZF0Str = new byte[1];
            KHST.ByteToString(ref rdZF0Str, readBuf, 1, 0);
            ZF0= rdZF0Str[0].ToString();
            Console.WriteLine("\tZF0:{0}", ZF0);
            //for (int i = 0; i < 1; i++) Console.WriteLine("\tZF0:{0}", rdZF0Str[i]);

            //OP1 or OP2(ZF4)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 4, 1, readBuf);
            Console.WriteLine("讀OP1 or OP2(ZF4)");
            byte[] rdZF4Str = new byte[1];
            KHST.ByteToString(ref rdZF4Str, readBuf, 1, 0);
            ZF4 = rdZF4Str[0].ToString();
            Console.WriteLine("\tZF4:{0}", ZF4);
            //for (int i = 0; i < 1; i++) Console.WriteLine("\tZF4:{0}", rdZF4Str[i]);

            //量測數據完成信號(ZF8)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 8, 1, readBuf);
            Console.WriteLine("讀量測數據完成信號(ZF8)");
            byte[] rdZF8Str = new byte[1];
            KHST.ByteToString(ref rdZF8Str, readBuf, 1, 0);
            ZF8 = rdZF8Str[0].ToString();
            Console.WriteLine("\tZF8:{0}", ZF8);
            //for (int i = 0; i < 1; i++) Console.WriteLine("\tZF8:{0}", rdZF8Str[i]);

            //工件一量測數據OK or NG信號(ZF16)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 16, 1, readBuf);
            Console.WriteLine("讀工件一量測數據OK or NG信號(ZF16)");
            byte[] rdZF16Str = new byte[1];
            KHST.ByteToString(ref rdZF16Str, readBuf, 1, 0);
            ZF16ipqc1 = rdZF16Str[0].ToString();
            Console.WriteLine("\tZF16:{0}", ZF16ipqc1);
            //for (int i = 0; i < 1; i++) Console.WriteLine("\tZF16:{0}", rdZF16Str[i]);

            //工件二量測數據OK or NG信號(ZF18)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 18, 1, readBuf);
            Console.WriteLine("讀工件二量測數據OK or NG信號(ZF18)");
            byte[] rdZF18Str = new byte[1];
            KHST.ByteToString(ref rdZF18Str, readBuf, 1, 0);
            ZF18ipqc2 = rdZF18Str[0].ToString();
            Console.WriteLine("\tZF18:{0}", ZF18ipqc2);
            //for (int i = 0; i < 1; i++) Console.WriteLine("\tZF18:{0}", rdZF18Str[i]);

            //工件三量測數據OK or NG信號(ZF20)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 20, 1, readBuf);
            Console.WriteLine("讀工件三量測數據OK or NG信號(ZF20)");
            byte[] rdZF20Str = new byte[1];
            KHST.ByteToString(ref rdZF20Str, readBuf, 1, 0);
            ZF20ipqc3 = rdZF20Str[0].ToString();
            Console.WriteLine("\tZF20:{0}", ZF20ipqc3);
            //for (int i = 0; i < 1; i++) Console.WriteLine("\tZF20:{0}", rdZF20Str[0]);


            //讀製令單別(ZF30~ZF31)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 30, 2, readBuf);
            if (err != 0)
            {
                Console.WriteLine(err);
            }
            Console.WriteLine("讀製令單別(ZF30~ZF31)");
            byte[] rdZF30Str = new byte[4];
            KHST.ByteToString(ref rdZF30Str, readBuf, 1, 0, 4);
            //只有ZF和EM暫存器編號是用十進制累加,其餘暫存器編號則為16進制累加
            //存英文字串到暫存器需要使用ASCII16BIT格式,但是讀出來是二進制的值,還需要轉碼成文字才能顯示該字串
            mo = System.Text.Encoding.GetEncoding(65001).GetString(rdZF30Str);
            Console.WriteLine(System.Text.Encoding.GetEncoding(65001).GetString(rdZF30Str));
            //讀製令單號(ZF32~ZF37)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 32, 6, readBuf);
            if (err != 0)
            {
                Console.WriteLine(err);
            }
            Console.WriteLine("讀製令單號(ZF32~ZF37)");
            byte[] rdZF32Str = new byte[12];
            KHST.ByteToString(ref rdZF32Str, readBuf, 1, 0, 12);
            //只有ZF和EM暫存器編號是用十進制累加,其餘暫存器編號則為16進制累加
            //存英文字串到暫存器需要使用ASCII16BIT格式,但是讀出來是二進制的值,還需要轉碼成文字才能顯示該字串
            mn = System.Text.Encoding.GetEncoding(65001).GetString(rdZF32Str);
            if (mn.StartsWith("-"))
            {
                mn = mn.Substring(1);
            }
            Console.WriteLine(mn);

            //讀員工工號(ZF40~ZF44)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 40, 5, readBuf);
            if (err != 0)
            {
                Console.WriteLine(err);
            }
            Console.WriteLine("讀員工工號(ZF40)");
            byte[] rdZF40Str = new byte[10];
            KHST.ByteToString(ref rdZF40Str, readBuf, 1, 0, 10);
            //轉碼成文字
            employeeID = System.Text.Encoding.GetEncoding(65001).GetString(rdZF40Str);
            Console.WriteLine(employeeID);

            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result;
            try
            {
                connection = new SqlConnection(connectionString);

                connection.Open();
                Console.WriteLine("Connection opened successfully.");
                String DBName = "TEST1";//測試資料庫
                //public static String DBName="JOYTECH";//正式資料庫
                mo = "1";
                mn = "1";
                ZF0 = "1";
                if (ZF0!=null&& !ZF0.Equals("")&&ZF0.Equals("1"))//現場人員
                {
                    string processCodeQuery = "select TA004 from " + DBName + ".dbo.SFCTA where TA001='" + mo + "' and TA002 ='" + mn + "'";
                    using (SqlCommand processCodeQuerycommand = new SqlCommand(processCodeQuery, connection))
                    {
                        using (SqlDataReader processCodeQueryreader = processCodeQuerycommand.ExecuteReader())
                        {
                            if (processCodeQueryreader.Read())
                            {
                                processCode = processCodeQueryreader.GetString(processCodeQueryreader.GetOrdinal("TA004"));

                                string itemNumberQuery = "select TA006 from " + DBName + ".dbo.MOCTA"
                                     + " where TA001='" + mo + "' and TA002='" + mn + "'";
                                using (SqlCommand itemNumberQuerycommand = new SqlCommand(itemNumberQuery, connection))
                                {
                                    using (SqlDataReader itemNumberQueryreader = itemNumberQuerycommand.ExecuteReader())
                                    {
                                        if (itemNumberQueryreader.Read())
                                        {
                                            itemNumber = itemNumberQueryreader.GetString(itemNumberQueryreader.GetOrdinal("TA006"));
                                        }
                                        else
                                        {
                                            result = MessageBox.Show("找不到產品品號的資料，寫入資料失敗", "警告", buttons, MessageBoxIcon.Warning);
                                            Console.WriteLine("找不到產品品號的資料");
                                            return;
                                        }
                                    }
                                }
                                string maxFormNumberQuery = "SELECT  JYT012a002 FROM " + DBName + ".dbo.JYT012 where JYT012a002=(select  max(JYT012a002) from " + DBName + ".dbo.JYT012"
                                    + " where JYT012a005='" + itemNumber + "')";
                                using (SqlCommand maxFormNumberQuerycommand = new SqlCommand(maxFormNumberQuery, connection))
                                {
                                    //SqlDataReader reader = maxFormNumberQuerycommand.ExecuteReader();
                                    using (SqlDataReader maxFormNumberQueryreader = maxFormNumberQuerycommand.ExecuteReader())
                                    {
                                        if (maxFormNumberQueryreader.Read())
                                        {
                                            maxFormNumber = maxFormNumberQueryreader.GetString(maxFormNumberQueryreader.GetOrdinal("JYT012a002"));
                                        }
                                        else
                                        {
                                            result = MessageBox.Show("找不到產品品號的表單號碼最大值資料，寫入資料失敗", "警告", buttons, MessageBoxIcon.Warning);
                                            Console.WriteLine("找不到產品品號的表單號碼最大值資料");
                                            return;
                                        }
                                    }
                                }

                                //去DSS找基本資料
                                string getQCData = "select JYT012a002,JYT012a003,JYT012a004,JYT012a005,JYT012a006,UDF01,UDF02 from " + DBName + ".dbo.JYT012 where JYT012a005='" + itemNumber + "'"
                                    + " and JYT012a002='" + maxFormNumber + "'";
                                using (SqlCommand getQCDatacommand = new SqlCommand(getQCData, connection))
                                {
                                    //SqlDataReader reader = getQCDatacommand.ExecuteReader();
                                    using (SqlDataReader getQCDatareader = getQCDatacommand.ExecuteReader())
                                    {
                                        if (getQCDatareader.Read())
                                        {
                                            jyt012.manufactureOrder = mo;
                                            jyt012.manufactureNo = mn;
                                            jyt012.processCode = processCode;
                                            jyt012.productionLine = productLineId;
                                            jyt012.JYT012a002 = getQCDatareader.GetString(getQCDatareader.GetOrdinal("JYT012a002"));
                                            jyt012.JYT012a003 = getQCDatareader.GetString(getQCDatareader.GetOrdinal("JYT012a003"));
                                            jyt012.UDF02 = getQCDatareader.GetString(getQCDatareader.GetOrdinal("UDF02"));
                                            jyt012.JYT012a004 = getQCDatareader.GetString(getQCDatareader.GetOrdinal("JYT012a004"));
                                            jyt012.JYT012a005 = getQCDatareader.GetString(getQCDatareader.GetOrdinal("JYT012a005"));
                                            jyt012.JYT012a006 = getQCDatareader.GetString(getQCDatareader.GetOrdinal("JYT012a006"));
                                            jyt012.UDF01 = getQCDatareader.GetString(getQCDatareader.GetOrdinal("UDF01"));
                                        }
                                        else
                                        {
                                            result = MessageBox.Show("在DSS找不到基本資料，寫入資料失敗", "警告", buttons, MessageBoxIcon.Warning);
                                            Console.WriteLine("在DSS找不到基本資料");
                                            return;
                                        }
                                    }
                                }

                                //找首件/自主檢查項目
                                String getIpqc = "select JYT012b003,JYT012b005,JYT012b003,JYT012b006,JYT012b007,JYT012b008,JYT012b009,JYT012b010,UDF03 from " + DBName + ".dbo.JYT012 where JYT012a005='" + itemNumber + "'"
                                     + " and JYT012a002='" + maxFormNumber + "' order by JYT012b003 asc";
                                using (SqlCommand getIpqccommand = new SqlCommand(getIpqc, connection))
                                {
                                    using (SqlDataReader getIpqcreader = getIpqccommand.ExecuteReader())
                                    {
                                        while (getIpqcreader.Read())
                                        {
                                            item.JYT012b003 = getIpqcreader["JYT012b003"].ToString();
                                            item.JYT012b005 = getIpqcreader["JYT012b005"].ToString();
                                            item.JYT012b006 = getIpqcreader["JYT012b006"].ToString();
                                            item.JYT012b007 = getIpqcreader["JYT012b007"].ToString();
                                            item.JYT012b008 = getIpqcreader["JYT012b008"].ToString();
                                            item.JYT012b009 = getIpqcreader["JYT012b009"].ToString();
                                            item.JYT012b010 = getIpqcreader["JYT012b010"].ToString();
                                            item.UDF03 = getIpqcreader["UDF03"].ToString();

                                            // 將 IpqcItem 物件添加到 ArrayList 中
                                            ipqcItemList.Add(item);
                                        }
                                    }
                                }

                                string productLineIdQuery = "SELECT TA021 from " + DBName + ".dbo.MOCTA where TA001 ='" + mo + "' and TA002 ='" + mn + "'";
                                using (SqlCommand productLineIdQuerycommand = new SqlCommand(productLineIdQuery, connection))
                                {
                                    using (SqlDataReader productLineIdQueryreader = productLineIdQuerycommand.ExecuteReader())
                                    {
                                        if (productLineIdQueryreader.Read())
                                        {
                                            productLineId = productLineIdQueryreader.GetString(productLineIdQueryreader.GetOrdinal("TA021"));
                                        }
                                        else
                                        {
                                            result = MessageBox.Show("找不到生產線別的資料，寫入資料失敗", "警告", buttons, MessageBoxIcon.Warning);
                                            Console.WriteLine("找不到生產線別的資料");
                                            return;
                                        }
                                    }
                                }

                                //insert檢驗說明書基本資料並回傳qc_id
                                string insSpecData = "INSERT INTO " + DBName + ".dbo.QCDataCollection (itemName,specification,manufactureOrder,manufactureNo,partNumber,imageNumber,processCode,version," +
                                     "formNumber,customerCode,firstItemDate,firstItemStaff,predictQty,machineNumber,productionLine,imageFileName,ipqcCREATOR)" +
                                     ") VALUES (@itemName, @specification, @manufactureOrder,@manufactureNo,@partNumber,@imageNumbe,@processCode,@version," +
                                     "@formNumber,@customerCode,@firstItemDate,@firstItemStaff,@predictQty,@machineNumber,@productionLine,@imageFileName,@ipqcCREATOR)" +
                                     "; SELECT IDENT_CURRENT('" + DBName + ".dbo.QCDataCollection') AS id;";
                                // 資料參數
                                SqlParameter[] QCDataCollectionparameters = {
                                new SqlParameter("@itemName", jyt012.JYT012a003),
                                new SqlParameter("@specification", jyt012.UDF02),
                                new SqlParameter("@manufactureOrder", jyt012.manufactureOrder),
                                new SqlParameter("@manufactureNo", jyt012.manufactureNo),
                                new SqlParameter("@partNumber", jyt012.JYT012a005),
                                new SqlParameter("@imageNumber", jyt012.JYT012a004),
                                new SqlParameter("@processCode", jyt012.processCode),
                                new SqlParameter("@version", jyt012.JYT012a006),
                                new SqlParameter("@formNumber", jyt012.JYT012a002),
                                new SqlParameter("@customerCode", jyt012.customerCode),
                                new SqlParameter("@firstItemDate", jyt012.firstItemDate),
                                new SqlParameter("@firstItemStaff", jyt012.firstItemStaff),
                                new SqlParameter("@predictQty", jyt012.predictQty),
                                new SqlParameter("@machineNumber", jyt012.machineNumber),
                                new SqlParameter("@productionLine", jyt012.productionLine),
                                new SqlParameter("@imageFileName", jyt012.UDF01),
                                new SqlParameter("@ipqcCREATOR", employeeID)
                            };

                                using (SqlCommand insSpecDatacommand = new SqlCommand(insSpecData, connection))
                                {
                                    // 設定參數
                                    insSpecDatacommand.Parameters.AddRange(QCDataCollectionparameters);

                                    // 執行插入和取得ID的查詢
                                    using (SqlDataReader insSpecDatareader = insSpecDatacommand.ExecuteReader())
                                    {
                                        if (insSpecDatareader.Read() && insSpecDatareader["id"] != DBNull.Value)
                                        {
                                            // 取得自增ID
                                            qcId = insSpecDatareader["id"].ToString();
                                            Console.WriteLine("Inserted row ID: " + qcId);
                                        }
                                        else
                                        {
                                            result = MessageBox.Show("無回傳qc_id資料，寫入資料失敗", "警告", buttons, MessageBoxIcon.Warning);
                                            Console.WriteLine("無回傳qc_id資料");
                                            return;
                                        }
                                    }
                                }

                                for (int i = 0; i <= ipqcItemList.Count; i++)
                                {
                                    IpqcItem itemList = (IpqcItem)ipqcItemList[i];
                                    //找首件/自主檢查項目
                                    String getIpqcData = "select JYT012b003,JYT012b005,JYT012b006,JYT012b007,JYT012b008,JYT012b009,JYT012b010,UDF03 from " + DBName + ".dbo.JYT012 where JYT012a005='" + itemNumber + "'"
                                    + " and JYT012a002='" + maxFormNumber + "' and JYT012b003='" + itemList.JYT012b003 + "'";
                                    using (SqlCommand getIpqcDatacommand = new SqlCommand(getIpqcData, connection))
                                    {
                                        //SqlDataReader reader = command.ExecuteReader();
                                        using (SqlDataReader getIpqcDatareader = getIpqcDatacommand.ExecuteReader())
                                        {
                                            if (getIpqcDatareader.Read())
                                            {
                                                primumQCData.JYT012b003 = getIpqcDatareader.GetString(getIpqcDatareader.GetOrdinal("JYT012b003"));
                                                primumQCData.JYT012b005 = getIpqcDatareader.GetString(getIpqcDatareader.GetOrdinal("JYT012b005"));
                                                primumQCData.JYT012b006 = getIpqcDatareader.GetString(getIpqcDatareader.GetOrdinal("JYT012b006"));
                                                primumQCData.JYT012b007 = getIpqcDatareader.GetString(getIpqcDatareader.GetOrdinal("JYT012b007"));
                                                primumQCData.JYT012b008 = getIpqcDatareader.GetString(getIpqcDatareader.GetOrdinal("JYT012b008"));
                                                primumQCData.JYT012b009 = getIpqcDatareader.GetString(getIpqcDatareader.GetOrdinal("JYT012b009"));
                                                primumQCData.JYT012b010 = getIpqcDatareader.GetString(getIpqcDatareader.GetOrdinal("JYT012b010"));
                                                primumQCData.UDF03 = getIpqcDatareader.GetString(getIpqcDatareader.GetOrdinal("UDF03"));
                                            }
                                            else
                                            {
                                                result = MessageBox.Show("找不到首件/自主檢查項目的資料，寫入資料失敗", "警告", buttons, MessageBoxIcon.Warning);
                                                Console.WriteLine("找不到首件/自主檢查項目的資料");
                                                return;
                                            }
                                        }

                                        //insert首件/自主檢查記錄實測狀況(ipqc1)
                                        String insIpqcData = "INSERT INTO " + DBName + ".dbo.QCDataCollectionContent(qc_id,itemSN,testItem,testUnit,standardValue,upperLimit,lowerLimit," +
                                                 "testTool,flag,ipqc1,ipqc2,ipqc3,ipqcCREATOR)" +
                                                 " VALUES (@qc_id,@itemSN,@testItem,@testUnit,@standardValue,@upperLimit,@lowerLimit,@testTool,@flag," +
                                                 "@ipqc1,@ipqc2,@ipqc3,@ipqcCREATOR)";
                                        // 資料參數
                                        SqlParameter[] QCDataCollectionContentparameters = {
                                            new SqlParameter("@qc_id", "value1"),
                                            new SqlParameter("@itemSN", "value2"),
                                            new SqlParameter("@testItem", "value2"),
                                            new SqlParameter("@testUnit", "value2"),
                                            new SqlParameter("@standardValue", "value2"),
                                            new SqlParameter("@upperLimit", "value2"),
                                            new SqlParameter("@lowerLimit", "value2"),
                                            new SqlParameter("@testTool", "value2"),
                                            new SqlParameter("@flag", "value2"),
                                            new SqlParameter("@ipqc1", "value2"),
                                            new SqlParameter("@ipqc2", "value2"),
                                            new SqlParameter("@ipqc3", "value2"),
                                            new SqlParameter("@ipqcCREATOR", "value2")
                                        };
                                        using (SqlCommand insIpqcDatacommand = new SqlCommand(insIpqcData, connection))
                                        {
                                            // 設定參數
                                            insIpqcDatacommand.Parameters.AddRange(QCDataCollectionContentparameters);
                                            // 執行 INSERT 操作
                                            int insIpqcRowsAffected = insIpqcDatacommand.ExecuteNonQuery();
                                            if (insIpqcRowsAffected > 0)
                                            {
                                                //寫上傳完成信號(PC→PLC)寫1筆資料(ZF14)
                                                int[] wzf14 = new int[1] { 0 };
                                                KHST.IntToByte(ref writeBuf, wzf14, 1, 0);
                                                err = KHL.KHLWriteDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 14, 1, writeBuf);
                                                result = MessageBox.Show("現場人員上傳資料完成", "結果", buttons);
                                                Console.WriteLine("現場人員寫入資料成功");
                                                return;
                                            }
                                            else
                                            {
                                                result = MessageBox.Show("寫入資料操作失敗", "警告", buttons, MessageBoxIcon.Warning);
                                                Console.WriteLine("寫入資料操作失敗。");
                                                return;
                                            }
                                        }
                                        //insert首件/自主檢查記錄實測狀況(ipqc2)

                                        //insert首件/自主檢查記錄實測狀況(ipqc3)

                                    }
                                }                                      
                            }
                            else
                            {
                                result = MessageBox.Show("查無製令單別單號，寫入資料失敗", "警告", buttons, MessageBoxIcon.Warning);
                                Console.WriteLine("查無製令單別單號，寫入資料失敗");
                                return;
                            }
                        }                        
                    }                                  
                }
                else if(ZF0 != null && !ZF0.Equals("") && ZF0.Equals("2"))//品檢人員
                {
                    //品檢人員找尋qc_id
                    string getQc_idQuery = "SELECT  qc_id FROM " + DBName + ".dbo.QCDataCollection where manufactureOrder='" + mo + "'"
                         + " and manufactureNo='" + mo + "'";
                    using (SqlCommand getQc_idQuerycommand = new SqlCommand(getQc_idQuery, connection))
                    {
                        using (SqlDataReader getQc_idQueryreader = getQc_idQuerycommand.ExecuteReader())
                        {
                            if (getQc_idQueryreader.Read())
                            {
                                qcId = getQc_idQueryreader["qc_id"].ToString();

                                //寫品檢成品檢驗記錄
                                String getQCDataCollectionData = "select qc_id,itemName,specification,manufactureOrder,manufactureNo,partNumber,imageNumber,version,formNumber,customerCode,firstItemDate,firstItemStaff,predictQty,machineNumber,imageFileName from " + DBName + ".dbo.QCDataCollection where "
                                     + "qc_id='" + qcId + "'";
                                using (SqlCommand getQCDataCollectionDatacommand = new SqlCommand(getQCDataCollectionData, connection))
                                {
                                    using (SqlDataReader getQCDataCollectionDatareader = getQCDataCollectionDatacommand.ExecuteReader())
                                    {
                                        if (getQCDataCollectionDatareader.Read())
                                        {
                                            jyt012.JYT012a003 = getQCDataCollectionDatareader.GetString(getQCDataCollectionDatareader.GetOrdinal("itemName"));
                                            jyt012.UDF02 = getQCDataCollectionDatareader.GetString(getQCDataCollectionDatareader.GetOrdinal("specification"));
                                            jyt012.JYT012a005 = getQCDataCollectionDatareader.GetString(getQCDataCollectionDatareader.GetOrdinal("partNumber"));
                                            jyt012.manufactureOrder = mo;
                                            jyt012.manufactureNo = mn;
                                            jyt012.customerCode = getQCDataCollectionDatareader.GetString(getQCDataCollectionDatareader.GetOrdinal("customerCode"));
                                            jyt012.JYT012a002 = getQCDataCollectionDatareader.GetString(getQCDataCollectionDatareader.GetOrdinal("formNumber"));
                                            jyt012.machineNumber = getQCDataCollectionDatareader.GetString(getQCDataCollectionDatareader.GetOrdinal("machineNumber"));
                                            jyt012.JYT012a006 = getQCDataCollectionDatareader.GetString(getQCDataCollectionDatareader.GetOrdinal("version"));
                                        }
                                        else
                                        {
                                            result = MessageBox.Show("查無檢驗記錄資料，寫入資料失敗", "警告", buttons, MessageBoxIcon.Warning);
                                            Console.WriteLine("查無檢驗記錄資料，寫入資料失敗");
                                            return;
                                        }
                                        //品檢人員找尋未檢驗的首件/自主檢查記錄詳細資料
                                        String getQCTestData = "select qc_id,itemSN,testItem,testUnit,standardValue,upperLimit,lowerLimit,testTool,flag,ipqc1,ipqc2,ipqc3 from " + DBName + ".dbo.QCDataCollectionContent where "
                                             + "qc_id='" + qcId + "'";
                                        using (SqlCommand getQCTestDatacommand = new SqlCommand(getQCTestData, connection))
                                        {
                                            using (SqlDataReader getQCTestDatareader = getQCTestDatacommand.ExecuteReader())
                                            {
                                                while (getQCTestDatareader.Read())
                                                {
                                                    item.JYT012b003 = getQCTestDatareader["itemSN"].ToString();
                                                    item.JYT012b005 = getQCTestDatareader["testItem"].ToString();
                                                    item.JYT012b006 = getQCTestDatareader["testUnit"].ToString();
                                                    item.JYT012b007 = getQCTestDatareader["standardValue"].ToString();
                                                    item.JYT012b008 = getQCTestDatareader["upperLimit"].ToString();
                                                    item.JYT012b009 = getQCTestDatareader["lowerLimit"].ToString();
                                                    ipqcItemList.Add(item);
                                                }
                                            }
                                        }
                                        //寫品檢成品檢驗記錄
                                        for (int i = 0; i <= ipqcItemList.Count; i++)
                                        {
                                            IpqcItem itemList = (IpqcItem)ipqcItemList[i];
                                            String updPqcData = "update " + DBName + ".dbo.QCDataCollectionContent set pqc1=@pqc1,pqc2=@pqc2,pqc3=@pqc3"
                                                 + " where qc_id='" + qcId + "' and itemSN='" + itemList.JYT012b003 + "'";
                                            using (SqlCommand updPqcDatacommand = new SqlCommand(updPqcData, connection))
                                            {
                                                updPqcDatacommand.Parameters.AddWithValue("@pqc1", "NewValue1");
                                                updPqcDatacommand.Parameters.AddWithValue("@pqc2", "NewValue2");
                                                updPqcDatacommand.Parameters.AddWithValue("@pqc3", "Condition");
                                                int updateRowsAffected = updPqcDatacommand.ExecuteNonQuery();
                                                if (updateRowsAffected > 0)
                                                {
                                                    Console.WriteLine("資料更新成功。");
                                                    if(ZF16ipqc1.Equals("1")&& ZF16ipqc1.Equals("1")&& ZF16ipqc1.Equals("1"))
                                                    {
                                                        defective = "合格";
                                                    }
                                                    else
                                                    {
                                                        defective = "不合格";
                                                    }
                                                    //寫判定結果
                                                    String insPqcResult = "insert into " + DBName + ".dbo.QCDataCollectionResult(qc_id,defective,determination,supervisor,inspectors,handleResult,pqcCREATOR)"
                                                         + " values(@qc_id,@defective,@determination,@supervisor,@inspectors,@handleResult,@pqcCREATOR)";
                                                    // 資料參數
                                                    SqlParameter[] QCDataCollectionResultparameters = {
                                                        new SqlParameter("@qc_id", qcId),
                                                        new SqlParameter("@itemSN", defective),
                                                        new SqlParameter("@testItem", "value2"),
                                                        new SqlParameter("@testUnit", "value2"),
                                                        new SqlParameter("@standardValue", "value2"),
                                                        new SqlParameter("@upperLimit", "value2"),
                                                        new SqlParameter("@lowerLimit", "value2"),
                                                        new SqlParameter("@testTool", "value2"),
                                                        new SqlParameter("@flag", "value2"),
                                                        new SqlParameter("@ipqc1", "value2"),
                                                        new SqlParameter("@ipqc2", "value2"),
                                                        new SqlParameter("@ipqc3", "value2"),
                                                        new SqlParameter("@ipqcCREATOR", "value2")
                                                    };
                                                    using (SqlCommand insPqcResultcommand = new SqlCommand(insPqcResult, connection))
                                                    {
                                                        // 設定參數
                                                        insPqcResultcommand.Parameters.AddRange(QCDataCollectionResultparameters);
                                                        // 執行 INSERT 操作
                                                        int insRowsAffected = insPqcResultcommand.ExecuteNonQuery();
                                                        if (insRowsAffected > 0)
                                                        {
                                                            //寫上傳完成信號(PC→PLC)寫1筆資料(ZF14)
                                                            int[] wzf14 = new int[1] { 0 };
                                                            KHST.IntToByte(ref writeBuf, wzf14, 1, 0);
                                                            err = KHL.KHLWriteDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 14, 1, writeBuf);
                                                            result = MessageBox.Show("品檢人員上傳資料完成", "結果", buttons);
                                                            Console.WriteLine("品檢人員寫入資料成功");
                                                            return;
                                                        }
                                                        else
                                                        {
                                                            result = MessageBox.Show("寫入資料操作失敗", "警告", buttons, MessageBoxIcon.Warning);
                                                            Console.WriteLine("寫入資料操作失敗。");
                                                            return;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    result = MessageBox.Show("沒有資料被更新，寫入資料失敗", "警告", buttons, MessageBoxIcon.Warning);
                                                    Console.WriteLine("沒有資料被更新。");
                                                    return;
                                                }

                                            }

                                        }

                                    }
                                }
                            }
                            else
                            {
                                result = MessageBox.Show("製令單別單號找不到現場人員檢查過的資料，寫入資料失敗", "警告", buttons, MessageBoxIcon.Warning);
                                Console.WriteLine("製令單別單號找不到現場人員檢查過的資料");
                                return;
                            }
                        }
                    }
                }
                else
                {
                    result = MessageBox.Show("現場人員和品檢人員回傳值有誤，操作失敗", "警告", buttons, MessageBoxIcon.Warning);
                    Console.WriteLine("現場人員和品檢人員回傳值有誤");
                    return;
                }
            }
            catch (Exception ex)
            {
                result = MessageBox.Show("系統有誤，操作失敗", "警告", buttons, MessageBoxIcon.Warning);
                Console.WriteLine("An error occurred: " + ex.Message);
                return;
            }
            finally
            {
                err = KHL.KHLDisconnect(sock);
                if (err != 0)
                {
                    Console.WriteLine(err);
                }

                if (connection != null)
                {
                    connection.Close();
                }
            }
        }
    }
}
