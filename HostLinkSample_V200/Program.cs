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
            string ZF10 = "";
            string ZF16ipqc1 = "";//工件一量測數據信號(1→OK;2→NG)
            string ZF18ipqc2 = "";//工件二量測數據信號(1→OK;2→NG)
            string ZF20ipqc3 = "";//工件三量測數據信號(1→OK;2→NG)
            string mo = "";//製令單別
            string mn = "";//製令單號
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


            try
            {
                connection = new SqlConnection(connectionString);

                connection.Open();
                Console.WriteLine("Connection opened successfully.");
                String DBName = "TEST1";//測試資料庫
                //public static String DBName="JOYTECH";//正式資料庫

                if (ZF0!=null&& !ZF0.Equals("")&&ZF0.Equals("1"))//現場人員
                {
                    string query = "SELECT  * FROM " + DBName + ".dbo.QCDataCollection;";
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        if (reader.Read())
                        {
                            string name = reader.GetString(reader.GetOrdinal("specification"));
                            Console.WriteLine($"Name: {name}");
                        }
                        reader.Close();

                    }


                    string itemNumberQuery = "select TA006 from " + DBName + ".dbo.MOCTA"
                        + " where TA001='" + mo + "' and TA002='" + mn + "'";
                    using (SqlCommand command = new SqlCommand(itemNumberQuery, connection))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        if (reader.Read())
                        {
                            itemNumber = reader.GetString(reader.GetOrdinal("TA006"));
                        }
                        reader.Close();
                    }
                    string maxFormNumberQuery = "SELECT  JYT012a002 FROM " + DBName + ".dbo.JYT012 where JYT012a002=(select  max(JYT012a002) from " + DBName + ".dbo.JYT012"
                        + " where JYT012a005='" + itemNumber + "')";
                    using (SqlCommand command = new SqlCommand(maxFormNumberQuery, connection))
                    {
                        SqlDataReader reader = command.ExecuteReader();

                        if (reader.Read())
                        {
                            maxFormNumber = reader.GetString(reader.GetOrdinal("JYT012a002"));
                        }

                        reader.Close();
                    }

                    //去DSS找基本資料
                    string getQCData = "select JYT012a002,JYT012a003,JYT012a004,JYT012a005,JYT012a006,UDF01,UDF02 from " + DBName + ".dbo.JYT012 where JYT012a005='" + itemNumber + "'"
                        + " and JYT012a002='" + maxFormNumber + "'";
                    using (SqlCommand command = new SqlCommand(getQCData, connection))
                    {
                        SqlDataReader reader = command.ExecuteReader();

                        if (reader.Read())
                        {
                            jyt012.manufactureOrder = mo;
                            jyt012.manufactureNo = mn;
                            jyt012.processCode = processCode;
                            jyt012.productionLine = productLineId;
                            jyt012.JYT012a002 = reader.GetString(reader.GetOrdinal("JYT012a002"));
                            jyt012.JYT012a003 = reader.GetString(reader.GetOrdinal("JYT012a003"));
                            jyt012.UDF02 = reader.GetString(reader.GetOrdinal("UDF02"));
                            jyt012.JYT012a004 = reader.GetString(reader.GetOrdinal("JYT012a004"));
                            jyt012.JYT012a005 = reader.GetString(reader.GetOrdinal("JYT012a005"));
                            jyt012.JYT012a006 = reader.GetString(reader.GetOrdinal("JYT012a006"));
                            jyt012.UDF01 = reader.GetString(reader.GetOrdinal("UDF01"));
                        }

                        reader.Close();
                    }

                    //找首件/自主檢查項目
                    String getIpqc = "select JYT012b003,JYT012b005,JYT012b003,JYT012b006,JYT012b007,JYT012b008,JYT012b009,JYT012b010,UDF03 from " + DBName + ".dbo.JYT012 where JYT012a005='" + itemNumber + "'"
                         + " and JYT012a002='" + maxFormNumber + "' order by JYT012b003 asc";
                    using (SqlCommand command = new SqlCommand(getIpqc, connection))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            item.JYT012b003 = reader["JYT012b003"].ToString();
                            item.JYT012b005 = reader["JYT012b005"].ToString();
                            item.JYT012b006 = reader["JYT012b006"].ToString();
                            item.JYT012b007 = reader["JYT012b007"].ToString();
                            item.JYT012b008 = reader["JYT012b008"].ToString();
                            item.JYT012b009 = reader["JYT012b009"].ToString();
                            item.JYT012b010 = reader["JYT012b010"].ToString();
                            item.UDF03 = reader["UDF03"].ToString();

                            // 將 IpqcItem 物件添加到 ArrayList 中
                            ipqcItemList.Add(item);
                        }
                        reader.Close();
                    }


                    string productLineIdQuery = "SELECT TA021 from " + DBName + ".dbo.MOCTA where TA001 ='" + mo + "' and TA002 ='" + mn + "'";
                    using (SqlCommand command = new SqlCommand(productLineIdQuery, connection))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        if (reader.Read())
                        {
                            productLineId = reader.GetString(reader.GetOrdinal("TA021"));
                        }
                        reader.Close();
                    }
                    string processCodeQuery = "select TA004 from " + DBName + ".dbo.SFCTA where TA001='" + mo + "' and TA002 ='" + mn + "'";
                    using (SqlCommand command = new SqlCommand(processCodeQuery, connection))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        if (reader.Read())
                        {
                            processCode = reader.GetString(reader.GetOrdinal("TA004"));
                        }
                        reader.Close();
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

                    using (SqlCommand command = new SqlCommand(insSpecData, connection))
                    {
                        // 設定參數
                        command.Parameters.AddRange(QCDataCollectionparameters);

                        // 執行插入和取得ID的查詢
                        SqlDataReader reader = command.ExecuteReader();
                        if (reader.Read() && reader["id"] != DBNull.Value)
                        {
                            // 取得自增ID
                            qcId = reader["id"].ToString();
                            Console.WriteLine("Inserted row ID: " + qcId);
                        }
                        reader.Close();
                    }

                    for (int i = 0; i <= ipqcItemList.Count; i++)
                    {
                        IpqcItem itemList = (IpqcItem)ipqcItemList[i];
                        //找首件/自主檢查項目
                        String getIpqcData = "select JYT012b003,JYT012b005,JYT012b006,JYT012b007,JYT012b008,JYT012b009,JYT012b010,UDF03 from " + DBName + ".dbo.JYT012 where JYT012a005='" + itemNumber + "'"
                        + " and JYT012a002='" + maxFormNumber + "' and JYT012b003='" + itemList.JYT012b003 + "'";
                        using (SqlCommand command = new SqlCommand(getIpqcData, connection))
                        {
                            SqlDataReader reader = command.ExecuteReader();
                            if (reader.Read())
                            {
                                primumQCData.JYT012b003 = reader.GetString(reader.GetOrdinal("JYT012b003"));
                                primumQCData.JYT012b005 = reader.GetString(reader.GetOrdinal("JYT012b005"));
                                primumQCData.JYT012b006 = reader.GetString(reader.GetOrdinal("JYT012b006"));
                                primumQCData.JYT012b007 = reader.GetString(reader.GetOrdinal("JYT012b007"));
                                primumQCData.JYT012b008 = reader.GetString(reader.GetOrdinal("JYT012b008"));
                                primumQCData.JYT012b009 = reader.GetString(reader.GetOrdinal("JYT012b009"));
                                primumQCData.JYT012b010 = reader.GetString(reader.GetOrdinal("JYT012b010"));
                                primumQCData.UDF03 = reader.GetString(reader.GetOrdinal("UDF03"));
                            }
                            reader.Close();
                            //insert首件/自主檢查記錄實測狀況
                            String insertSQL = "INSERT INTO " + DBName + ".dbo.QCDataCollectionContent(qc_id,itemSN,testItem,testUnit,standardValue,upperLimit,lowerLimit," +
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
                            command.Parameters.AddRange(QCDataCollectionContentparameters);
                            command.ExecuteNonQuery();
                            Console.WriteLine("Insert operation completed successfully.");
                        }
                    }
                    MessageBoxButtons buttons = MessageBoxButtons.OK;
                    DialogResult result;

                    // Displays the MessageBox.
                    result = MessageBox.Show("上傳資料完成", "結果", buttons);

                    //寫上傳完成信號(PC→PLC)寫1筆資料(ZF14)
                    int[] wzf14 = new int[1] { 0 };
                    KHST.IntToByte(ref writeBuf, wzf14, 1, 0);
                    err = KHL.KHLWriteDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 14, 1, writeBuf);
                }
                else if(ZF0 != null && !ZF0.Equals("") && ZF0.Equals("2"))//品檢人員
                {
                    //品檢人員找尋qc_id

                    //寫品檢成品檢驗記錄

                    //寫判定結果


                }
                else
                {
                    Console.WriteLine("現場人員和品檢人員回傳值有誤");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
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
