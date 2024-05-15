using System;
using System.Collections;
using KHL = KvHostLink.KvHostLink;
using KHST = KvStruct.KvStruct;
using System.Data.SqlClient;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using System.Windows.Forms;
using System.Xml.Linq;
using HostLinkSample_V200;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Header;
using System.Threading;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
namespace cs_dll_sample
{
    class Program
    {
        static async Task Main(string[] args)
        {
            int err = 0;
            int sock = 0;  
            //測試DB
            //string connectionString = "Server=192.168.1.9,1433;Database=TEST1;User Id=pwc;Password=PWC@admin;";
            //正式DB
            string connectionString = "Server=192.168.1.9,1433;Database=JOYTECH;User Id=pwc;Password=PWC@admin;";
            //Log DB
            string LogconnectionString = "Server=192.168.1.9,1433;Database=LOG;User Id=pwc;Password=PWC@admin;";
            //string DBName = "TEST1";//測試資料庫
            string DBName="JOYTECH";//正式資料庫
            SqlConnection connection = null;
            SqlConnection Logconnection = null;
            ArrayList ipqc1ResultList = new ArrayList();
            ArrayList ipqc2ResultList = new ArrayList();
            ArrayList ipqc3ResultList = new ArrayList();
            ArrayList ipqcItemList = new ArrayList();
            float[] rdEM3000float = new float[2];//工件一量測數據暫存器
            float[] rdEM4000float = new float[2];//工件二量測數據暫存器
            float[] rdEM5000float = new float[2];//工件三量測數據暫存器
            string itemNumber = "";//產品品號
            string maxFormNumber = ""; //產品品號的表單號碼最大值
            string productLineId = "";//生產線別
            string processCode = "";//製程代號
            string qcId ="";//主鍵
            string employeeID = "";//員工工號
            //string predictQty = "";//預計產量
            string ZF50YearStr = "";//年
            string ZF60MonthStr = "";//月
            string ZF70DayStr = "";//日
            string firstItemDate = "";//首件日期(格式:年-月-日)
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
            int insIpqc1RowsAffected = 0;
            int updIpqc2RowsAffected = 0;
            int updIpqc3RowsAffected = 0;
            int updPqc1RowsAffected = 0;
            int updPqc2RowsAffected = 0;
            int updPqc3RowsAffected = 0;
            string JYT012a008 = "";//2080→OP1;2090→OP2
            JYT012 jyt012 = new JYT012();
            IpqcItem primumQCData = new IpqcItem();
            //IpqcItem item = new IpqcItem();
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result;

            Console.WriteLine("PLC連線建立中...");
            err = KHL.KHLInit();
            // 创建一个 CancellationTokenSource 用于取消任务
            var cts = new CancellationTokenSource();

            // 设置超时时间为3000毫秒
            cts.CancelAfter(3);

            // 创建一个Task并传入取消令牌
            var task = Task.Run(() =>
            {
                try
                {                    
                    err = KHL.KHLConnect("10.1.9.106", 8500, 3000, KvHostLink.KHLSockType.SOCK_TCP, ref sock);
                    if (err != 0)
                    {
                        result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                        Console.WriteLine("PLC連線發生異常");
                        Console.WriteLine(err);
                        KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                        return;
                    }
                }
                catch (Exception ex)
                {
                    result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                    Console.WriteLine("PLC連線發生異常：" + ex.Message);
                    KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗" + ex.Message, "錯誤", "");
                    return;
                }
            }, cts.Token);

            try
            {
                // 等待任务完成
                await task;
            }
            catch (Exception ex)
            {
                result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                Console.WriteLine("PLC連線發生異常");
                KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗" + ex.Message, "錯誤", "");
                return;            
            }          
           
            // Interface Buffer
            byte[] readBuf = new byte[2048];
            byte[] writeBuf = new byte[2048];
 
            //讀EM3000工件一量測數據暫存器
            Console.WriteLine("讀EM3000工件一量測數據暫存器(EM3000~EM3098)");

            for (uint i = 0; i < 100; i += 2)
            {
                //有部分暫存器存一筆資料需要放2個暫存器,所以迴圈i一次要加2,wordNum參數為2
                err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_EM, 3000 + i, 2, readBuf);
                if (err != 0)
                {
                    result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                    Console.WriteLine("PLC連線發生異常");
                    Console.WriteLine(err);
                    KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗" , "錯誤", err.ToString());
                    return;
                }
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
                if (err != 0)
                {
                    result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                    Console.WriteLine("PLC連線發生異常");
                    Console.WriteLine(err);
                    KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                    return;
                }
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
                if (err != 0)
                {
                    result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                    Console.WriteLine("PLC連線發生異常");
                    Console.WriteLine(err);
                    KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                    return;
                }
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
            if (err != 0)
            {
                result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                Console.WriteLine("PLC連線發生異常");
                Console.WriteLine(err);
                KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                return;
            }
            Console.WriteLine("讀現場人員 or 品檢人員(ZF0)");
            int[] rdZF0Str = new int[2];
            KHST.ByteToInt(ref rdZF0Str, readBuf, 2, 0);
            ZF0= rdZF0Str[0].ToString();
            Console.WriteLine("\tZF0:{0}", ZF0);
            //for (int i = 0; i < 1; i++) Console.WriteLine("\tZF0:{0}", rdZF0Str[0]);

            //OP1 or OP2(ZF4)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 4, 1, readBuf);
            if (err != 0)
            {
                result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                Console.WriteLine("PLC連線發生異常");
                Console.WriteLine(err);
                KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                return;
            }
            Console.WriteLine("讀OP1 or OP2(ZF4)");
            int[] rdZF4Str = new int[2];
            KHST.ByteToInt(ref rdZF4Str, readBuf, 2, 0);
            ZF4 = rdZF4Str[0].ToString();
            Console.WriteLine("\tZF4:{0}", ZF4);
            if(!ZF4.Equals("") && ZF4.Equals("1")) //表OP1
            {
                JYT012a008 = "2080";
            }
            else //表OP2
            {
                JYT012a008 = "2090";
            }
            //for (int i = 0; i < 1; i++) Console.WriteLine("\tZF4:{0}", rdZF4Str[i]);

            //量測數據完成信號(ZF8)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 8, 1, readBuf);
            if (err != 0)
            {
                result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                Console.WriteLine("PLC連線發生異常");
                Console.WriteLine(err);
                KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                return;
            }
            Console.WriteLine("讀量測數據完成信號(ZF8)");
            int[] rdZF8Str = new int[2];
            KHST.ByteToInt(ref rdZF8Str, readBuf, 2, 0);
            ZF8 = rdZF8Str[0].ToString();
            Console.WriteLine("\tZF8:{0}", ZF8);
            //for (int i = 0; i < 1; i++) Console.WriteLine("\tZF8:{0}", rdZF8Str[i]);

            //工件一量測數據OK or NG信號(ZF16)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 16, 1, readBuf);
            if (err != 0)
            {
                result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                Console.WriteLine("PLC連線發生異常");
                Console.WriteLine(err);
                KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                return;
            }
            Console.WriteLine("讀工件一量測數據OK or NG信號(ZF16)");
            int[] rdZF16Str = new int[2];
            KHST.ByteToInt(ref rdZF16Str, readBuf, 2, 0);
            ZF16ipqc1 = rdZF16Str[0].ToString();
            Console.WriteLine("\tZF16:{0}", ZF16ipqc1);
            //for (int i = 0; i < 2; i++) Console.WriteLine("\tZF16:{0}", rdZF16Str[1]);

            //工件二量測數據OK or NG信號(ZF18)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 18, 1, readBuf);
            if (err != 0)
            {
                result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                Console.WriteLine("PLC連線發生異常");
                Console.WriteLine(err);
                KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                return;
            }
            Console.WriteLine("讀工件二量測數據OK or NG信號(ZF18)");
            int[] rdZF18Str = new int[2];
            KHST.ByteToInt(ref rdZF18Str, readBuf, 2, 0);
            ZF18ipqc2 = rdZF18Str[0].ToString();
            Console.WriteLine("\tZF18:{0}", ZF18ipqc2);
            //for (int i = 0; i < 1; i++) Console.WriteLine("\tZF18:{0}", rdZF18Str[i]);

            //工件三量測數據OK or NG信號(ZF20)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 20, 1, readBuf);
            if (err != 0)
            {
                result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                Console.WriteLine("PLC連線發生異常");
                Console.WriteLine(err);
                KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                return;
            }
            Console.WriteLine("讀工件三量測數據OK or NG信號(ZF20)");
            int[] rdZF20Str = new int[2];
            KHST.ByteToInt(ref rdZF20Str, readBuf, 2, 0);
            ZF20ipqc3 = rdZF20Str[0].ToString();
            Console.WriteLine("\tZF20:{0}", ZF20ipqc3);
            //for (int i = 0; i < 1; i++) Console.WriteLine("\tZF20:{0}", rdZF20Str[0]);

            //首件日期-年(ZF50)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 50, 2, readBuf);
            if (err != 0)
            {
                result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                Console.WriteLine("PLC連線發生異常");
                Console.WriteLine(err);
                KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                return;
            }
            int[] rdZF50Str = new int[4];
            KHST.ByteToInt(ref rdZF50Str, readBuf, 4, 0);
            ZF50YearStr = rdZF50Str[0].ToString();
            Console.WriteLine("\tZF50:{0}", ZF50YearStr);
            //for (int i = 0; i < 2; i++) Console.WriteLine("\tZF50:{0}", rdZF50Str[0]);

            //首件日期-月(ZF60)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 60, 2, readBuf);
            if (err != 0)
            {
                result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                Console.WriteLine("PLC連線發生異常");
                Console.WriteLine(err);
                KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                return;
            }
            int[] rdZF60Str = new int[4];
            KHST.ByteToInt(ref rdZF60Str, readBuf, 4, 0);
            ZF60MonthStr = rdZF60Str[0].ToString();
            Console.WriteLine("\tZF60:{0}", ZF60MonthStr);
            //for (int i = 0; i < 2; i++) Console.WriteLine("\tZF60:{0}", rdZF60Str[0]);

            //首件日期-日(ZF70)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 70, 2, readBuf);
            if (err != 0)
            {
                result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                Console.WriteLine("PLC連線發生異常");
                Console.WriteLine(err);
                KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                return;
            }
            int[] rdZF70Str = new int[4];
            KHST.ByteToInt(ref rdZF70Str, readBuf, 4, 0);
            ZF70DayStr = rdZF70Str[0].ToString();
            Console.WriteLine("\tZF70:{0}", ZF70DayStr);
            //for (int i = 0; i < 2; i++) Console.WriteLine("\tZF70:{0}", rdZF70Str[0]);

            //格式化為 YYYY-MM-DD 的格式
            DateTime date = DateTime.Parse(ZF50YearStr + "-" + ZF60MonthStr + "-" + ZF70DayStr);
            firstItemDate = date.ToString("yyyy-MM-dd");
            Console.WriteLine("\tfirstItemDate:{0}", firstItemDate);

            //預計產量(ZF80)
            //err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 80, 4, readBuf);
            //if (err != 0)
            //{
            //    result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
            //    Console.WriteLine("PLC連線發生異常");
            //    Console.WriteLine(err);
            //    KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
            //    return;
            //}
            //Console.WriteLine("讀預計產量(ZF80)");
            //int[] rdZF80Str = new int[8];
            //KHST.ByteToInt(ref rdZF80Str, readBuf, 4, 0);
            //predictQty = rdZF80Str[0].ToString();
            //Console.WriteLine("\tZF80:{0}", predictQty);

            //讀製令單別(ZF30~ZF31)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 30, 2, readBuf);
            if (err != 0)
            {
                result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                Console.WriteLine("PLC連線發生異常");
                Console.WriteLine(err);
                KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                return;                
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
                result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                Console.WriteLine("PLC連線發生異常");
                Console.WriteLine(err);
                KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                return;
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
                result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                Console.WriteLine("PLC連線發生異常");
                Console.WriteLine(err);
                KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                return;
            }
            Console.WriteLine("讀員工工號(ZF40)");
            byte[] rdZF40Str = new byte[10];
            KHST.ByteToString(ref rdZF40Str, readBuf, 1, 0, 10);
            //轉碼成文字
            employeeID = System.Text.Encoding.GetEncoding(65001).GetString(rdZF40Str);
            Console.WriteLine(employeeID);

            //ZF18ipqc2 = "1";
            //ZF20ipqc3 = "2";
            //ZF0 = "2";

            try
            {
                connection = new SqlConnection(connectionString);

                connection.Open();
                Console.WriteLine("資料庫連線...");                
                
                if (ZF0!=null&& !ZF0.Equals("")&&ZF0.Equals("1"))//1→現場人員
                {
                    if(!ZF16ipqc1.Equals("") && ZF16ipqc1.Equals("1") && !ZF18ipqc2.Equals("") && ZF18ipqc2.Equals("1")
                        && !ZF20ipqc3.Equals("") && ZF20ipqc3.Equals("1"))
                    {
                        //取製程代號
                        string processCodeQuery = "select TA004 from " + DBName + ".dbo.SFCTA where TA001='" + mo + "' and TA002 ='" + mn + "'";
                        using (SqlCommand processCodeQuerycommand = new SqlCommand(processCodeQuery, connection))
                        {
                            using (SqlDataReader processCodeQueryreader = processCodeQuerycommand.ExecuteReader())
                            {
                                if (processCodeQueryreader.Read())
                                {
                                    processCode = processCodeQueryreader.GetString(processCodeQueryreader.GetOrdinal("TA004"));
                                    processCodeQueryreader.Close();
                                    //找產品品號
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
                                                result = MessageBox.Show("找不到產品品號的資料，操作失敗", "警告", buttons, MessageBoxIcon.Warning);
                                                Console.WriteLine("找不到產品品號的資料");
                                                KvHostLinkLog(employeeID, mo, mn, itemNumber, "找不到產品品號的資料，操作失敗", "警告", "");
                                                return;
                                            }
                                            itemNumberQueryreader.Close();
                                        }
                                    }
                                    //找該產品品號的表單號碼最大值
                                    string maxFormNumberQuery = "SELECT  JYT012a002 FROM " + DBName + ".dbo.JYT012 where JYT012a002=(select  max(JYT012a002) from " + DBName + ".dbo.JYT012"
                                        + " where JYT012a008='"+ JYT012a008+"' and JYT012a005='" + itemNumber + "')";
                                    using (SqlCommand maxFormNumberQuerycommand = new SqlCommand(maxFormNumberQuery, connection))
                                    {
                                        using (SqlDataReader maxFormNumberQueryreader = maxFormNumberQuerycommand.ExecuteReader())
                                        {
                                            if (maxFormNumberQueryreader.Read())
                                            {
                                                maxFormNumber = maxFormNumberQueryreader.GetString(maxFormNumberQueryreader.GetOrdinal("JYT012a002"));
                                            }
                                            else
                                            {
                                                result = MessageBox.Show("找不到產品品號的表單號碼最大值資料，操作失敗", "警告", buttons, MessageBoxIcon.Warning);
                                                Console.WriteLine("找不到產品品號的表單號碼最大值資料");
                                                KvHostLinkLog(employeeID, mo, mn, itemNumber, "找不到產品品號的表單號碼最大值資料，操作失敗", "警告", "");
                                                return;
                                            }
                                            maxFormNumberQueryreader.Close();
                                        }
                                    }

                                    //去DSS找基本資料
                                    string getQCData = "select JYT012a002,JYT012a003,JYT012a004,JYT012a005,JYT012a006,UDF01,UDF02 from " + DBName + ".dbo.JYT012 where JYT012a005='" + itemNumber + "'"
                                        + " and JYT012a002='" + maxFormNumber + "'";
                                    using (SqlCommand getQCDatacommand = new SqlCommand(getQCData, connection))
                                    {
                                        using (SqlDataReader getQCDatareader = getQCDatacommand.ExecuteReader())
                                        {
                                            if (getQCDatareader.Read())
                                            {
                                                jyt012.manufactureOrder = mo;
                                                jyt012.manufactureNo = mn;
                                                jyt012.processCode = processCode;
                                                //jyt012.productionLine = productLineId;
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
                                                result = MessageBox.Show("在DSS找不到基本資料，操作失敗", "警告", buttons, MessageBoxIcon.Warning);
                                                Console.WriteLine("在DSS找不到基本資料");
                                                KvHostLinkLog(employeeID, mo, mn, itemNumber, "在DSS找不到基本資料，操作失敗", "警告", "");
                                                return;
                                            }
                                            getQCDatareader.Close();
                                        }
                                    }

                                    //找首件/自主檢查項目
                                    String getIpqc = "select JYT012b003,JYT012b005,JYT012b006,JYT012b007,JYT012b008,JYT012b009,JYT012b010,UDF03 from " + DBName + ".dbo.JYT012 where JYT012a005='" + itemNumber + "'"
                                         + " and JYT012a002='" + maxFormNumber + "' and UPPER(JYT012b010)='TMX' order by JYT012b003 asc";
                                    using (SqlCommand getIpqccommand = new SqlCommand(getIpqc, connection))
                                    {
                                        using (SqlDataReader getIpqcreader = getIpqccommand.ExecuteReader())
                                        {
                                            while (getIpqcreader.Read())
                                            {
                                                IpqcItem item = new IpqcItem();
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
                                            getIpqcreader.Close();
                                        }
                                    }
                                    if (ipqcItemList.Count > 0)
                                    {
                                        if (!ipqc1ResultList.Count.Equals(ipqcItemList.Count))
                                        {
                                            result = MessageBox.Show("DSS資料與影像量測儀兩者檢測數據筆數不一致，操作失敗", "警告", buttons, MessageBoxIcon.Warning);
                                            Console.WriteLine("DSS資料與影像量測儀兩者檢測數據筆數不一致");
                                            KvHostLinkLog(employeeID, mo, mn, itemNumber, "DSS資料與影像量測儀兩者檢測數據筆數不一致，操作失敗", "警告", "");
                                            return;
                                        }
                                    }
                                    else
                                    {
                                        result = MessageBox.Show("在DSS系統找不到資料，操作失敗", "警告", buttons, MessageBoxIcon.Warning);
                                        Console.WriteLine("在DSS系統找不到資料");
                                        KvHostLinkLog(employeeID, mo, mn, itemNumber, "在DSS系統找不到資料，操作失敗", "警告", "");
                                        return;
                                    }

                                    //找生產線別
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
                                                result = MessageBox.Show("找不到生產線別的資料，操作失敗", "警告", buttons, MessageBoxIcon.Warning);
                                                Console.WriteLine("找不到生產線別的資料");
                                                KvHostLinkLog(employeeID, mo, mn, itemNumber, "找不到生產線別的資料，操作失敗", "警告", "");
                                                return;
                                            }
                                            productLineIdQueryreader.Close();
                                        }
                                    }

                                    //insert檢驗說明書基本資料並回傳qc_id
                                    string insSpecData = "INSERT INTO " + DBName + ".dbo.QCDataCollection (itemName,specification,manufactureOrder,manufactureNo,partNumber,imageNumber,processCode,version," +
                                         "formNumber,firstItemDate,firstItemStaff,machineNumber,productionLine,imageFileName,ipqcCREATOR)" +
                                         " VALUES (@itemName, @specification, @manufactureOrder,@manufactureNo,@partNumber,@imageNumber,@processCode,@version," +
                                         "@formNumber,@firstItemDate,@firstItemStaff,@machineNumber,@productionLine,@imageFileName,@ipqcCREATOR)" +
                                         "; SELECT SCOPE_IDENTITY();";
                                    // 資料參數
                                    SqlParameter[] QCDataCollectionparameters = {
                                        new SqlParameter("@itemName", jyt012.JYT012a003.ToString()),
                                        new SqlParameter("@specification", jyt012.UDF02.ToString()),
                                        new SqlParameter("@manufactureOrder", mo),
                                        new SqlParameter("@manufactureNo", mn),
                                        new SqlParameter("@partNumber", jyt012.JYT012a005),
                                        new SqlParameter("@imageNumber", jyt012.JYT012a004),
                                        new SqlParameter("@processCode", processCode),
                                        new SqlParameter("@version", jyt012.JYT012a006),
                                        new SqlParameter("@formNumber", jyt012.JYT012a002),
                                        new SqlParameter("@firstItemDate", firstItemDate),
                                        new SqlParameter("@firstItemStaff", employeeID),                                    
                                        new SqlParameter("@machineNumber", ""),
                                        new SqlParameter("@productionLine", productLineId),
                                        new SqlParameter("@imageFileName", jyt012.UDF01),
                                        new SqlParameter("@ipqcCREATOR", employeeID)
                                    };

                                    using (SqlCommand insSpecDatacommand = new SqlCommand(insSpecData, connection))
                                    {
                                        // 設定參數
                                        insSpecDatacommand.Parameters.AddRange(QCDataCollectionparameters);

                                        // 執行插入和取得ID的查詢
                                        var insertedId = insSpecDatacommand.ExecuteScalar();
                                        qcId = insertedId.ToString();
                                        Console.WriteLine("qcId: " + qcId);                                       
                                    }

                                    for (int i = 0; i < ipqcItemList.Count; i++)
                                    {
                                        IpqcItem itemList = (IpqcItem)ipqcItemList[i];
                                        //找首件/自主檢查項目
                                        String getIpqcData = "select JYT012b003,JYT012b005,JYT012b006,JYT012b007,JYT012b008,JYT012b009,JYT012b010,UDF03 from " + DBName + ".dbo.JYT012 where JYT012a005='" + itemNumber + "'"
                                        + " and JYT012a002='" + maxFormNumber + "' and JYT012b003='" + itemList.JYT012b003 + "' and  UPPER(JYT012b010)='TMX'";
                                        using (SqlCommand getIpqcDatacommand = new SqlCommand(getIpqcData, connection))
                                        {
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
                                                    KvHostLinkLog(employeeID, mo, mn, itemNumber, "找不到首件/自主檢查項目的資料，寫入資料失敗", "警告", "");
                                                    return;
                                                }
                                                getIpqcDatareader.Close();
                                            }

                                            //insert首件/自主檢查記錄實測狀況(ipqc1)
                                            String insIpqc1Data = "INSERT INTO " + DBName + ".dbo.QCDataCollectionContent(qc_id,itemSN,testItem,testUnit,standardValue,upperLimit,lowerLimit," +
                                                     "testTool,flag,ipqc1,ipqcCREATOR)" +
                                                     " VALUES (@qc_id,@itemSN,@testItem,@testUnit,@standardValue,@upperLimit,@lowerLimit,@testTool,@flag," +
                                                     "@ipqc1,@ipqcCREATOR)";
                                            // 資料參數
                                            SqlParameter[] QCDataCollectionContentparameters = {
                                                new SqlParameter("@qc_id", qcId),
                                                new SqlParameter("@itemSN", primumQCData.JYT012b003),
                                                new SqlParameter("@testItem", primumQCData.JYT012b005),
                                                new SqlParameter("@testUnit", primumQCData.JYT012b006),
                                                new SqlParameter("@standardValue", primumQCData.JYT012b007),
                                                new SqlParameter("@upperLimit", primumQCData.JYT012b008),
                                                new SqlParameter("@lowerLimit", primumQCData.JYT012b009),
                                                new SqlParameter("@testTool", primumQCData.JYT012b010),
                                                new SqlParameter("@flag", primumQCData.UDF03),
                                                new SqlParameter("@ipqc1", ipqc1ResultList[i]),
                                                new SqlParameter("@ipqcCREATOR", "")
                                            };
                                            using (SqlCommand insIpqc1Datacommand = new SqlCommand(insIpqc1Data, connection))
                                            {
                                                // 設定參數
                                                insIpqc1Datacommand.Parameters.AddRange(QCDataCollectionContentparameters);
                                                // 執行 INSERT 操作
                                                insIpqc1RowsAffected = insIpqc1Datacommand.ExecuteNonQuery();
                                            }
                                        }
                                    }

                                    //update首件/自主檢查記錄實測狀況(ipqc2)
                                    for (int i = 0; i < ipqcItemList.Count; i++)
                                    {
                                        IpqcItem itemList = (IpqcItem)ipqcItemList[i];
                                        String updIpqc2Data = "update " + DBName + ".dbo.QCDataCollectionContent set ipqc2=@ipqc2"
                                             + " where qc_id='" + qcId + "' and itemSN='" + itemList.JYT012b003 + "'";
                                        using (SqlCommand updIpqc2Datacommand = new SqlCommand(updIpqc2Data, connection))
                                        {
                                            updIpqc2Datacommand.Parameters.AddWithValue("@ipqc2", ipqc2ResultList[i]);
                                            int updIpqc2Rows = updIpqc2Datacommand.ExecuteNonQuery();
                                            updIpqc2RowsAffected = updIpqc2RowsAffected + updIpqc2Rows;
                                        }
                                    }

                                    //update首件/自主檢查記錄實測狀況(ipqc3)
                                    for (int i = 0; i < ipqcItemList.Count; i++)
                                    {
                                        IpqcItem itemList = (IpqcItem)ipqcItemList[i];
                                        String updIpqc3Data = "update " + DBName + ".dbo.QCDataCollectionContent set ipqc3=@ipqc3"
                                             + " where qc_id='" + qcId + "' and itemSN='" + itemList.JYT012b003 + "'";
                                        using (SqlCommand updIpqc3Datacommand = new SqlCommand(updIpqc3Data, connection))
                                        {
                                            updIpqc3Datacommand.Parameters.AddWithValue("@ipqc3", ipqc3ResultList[i]);
                                            int updIpqc3Rows = updIpqc3Datacommand.ExecuteNonQuery();
                                            updIpqc3RowsAffected = updIpqc3RowsAffected + updIpqc3Rows;
                                        }
                                    }

                                    //寫上傳完成信號值"1"到ZF14暫存器(PC→PLC)
                                    if (insIpqc1RowsAffected > 0 && updIpqc2RowsAffected > 0 && updIpqc3RowsAffected > 0)
                                    {
                                        int[] wzf14 = new int[1] { 1 };
                                        KHST.IntToByte(ref writeBuf, wzf14, 1, 0);
                                        err = KHL.KHLWriteDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 14, 1, writeBuf);
                                        if (err != 0)
                                        {
                                            result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                                            Console.WriteLine("PLC連線發生異常");
                                            Console.WriteLine(err);
                                            KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                                            return;
                                        }
                                        result = MessageBox.Show("現場人員上傳資料完成", "結果", buttons);
                                        Console.WriteLine("現場人員上傳資料完成");
                                        KvHostLinkLog(employeeID, mo, mn, itemNumber, "現場人員上傳資料完成", "結果", "");
                                        return;
                                    }
                                    else
                                    {
                                        result = MessageBox.Show("現場人員上傳資料失敗", "警告", buttons, MessageBoxIcon.Warning);
                                        Console.WriteLine("現場人員上傳資料失敗");
                                        KvHostLinkLog(employeeID, mo, mn, itemNumber, "現場人員上傳資料失敗", "警告", "");
                                        return;
                                    }
                                }
                                else
                                {
                                    result = MessageBox.Show("查無製令單別單號，操作失敗", "警告", buttons, MessageBoxIcon.Warning);
                                    Console.WriteLine("查無製令單別單號");
                                    KvHostLinkLog(employeeID, mo, mn, itemNumber, "查無製令單別單號，操作失敗", "警告", "");
                                    return;
                                }
                            }
                        }
                    }
                    else
                    {
                        result = MessageBox.Show("工件量測數據不是全部都OK，操作失敗", "警告", buttons, MessageBoxIcon.Warning);
                        Console.WriteLine("工件量測數據不是全部都OK");
                        KvHostLinkLog(employeeID, mo, mn, itemNumber, "工件量測數據不是全部都OK，操作失敗", "警告", "");
                        return;
                    }                                                     
                }
                else if(ZF0 != null && !ZF0.Equals("") && ZF0.Equals("2"))//品檢人員
                {
                    //品檢人員找尋qc_id
                    string getQc_idQuery = "SELECT  qc_id FROM " + DBName + ".dbo.QCDataCollection where pqcStatus='0' and manufactureOrder='" + mo + "'"
                         + " and manufactureNo='" + mn + "'";
                    using (SqlCommand getQc_idQuerycommand = new SqlCommand(getQc_idQuery, connection))
                    {
                        using (SqlDataReader getQc_idQueryreader = getQc_idQuerycommand.ExecuteReader())
                        {
                            if (getQc_idQueryreader.Read())
                            {
                                qcId = getQc_idQueryreader["qc_id"].ToString();
                                getQc_idQueryreader.Close();
                                //寫品檢成品檢驗記錄
                                String getQCDataCollectionData = "select qc_id,itemName,specification,manufactureOrder,manufactureNo,partNumber,imageNumber,version,formNumber,firstItemDate,firstItemStaff,machineNumber,imageFileName from " + DBName + ".dbo.QCDataCollection where "
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
                                            jyt012.JYT012a002 = getQCDataCollectionDatareader.GetString(getQCDataCollectionDatareader.GetOrdinal("formNumber"));
                                            jyt012.machineNumber = getQCDataCollectionDatareader.GetString(getQCDataCollectionDatareader.GetOrdinal("machineNumber"));
                                            jyt012.JYT012a006 = getQCDataCollectionDatareader.GetString(getQCDataCollectionDatareader.GetOrdinal("version"));
                                        }
                                        else
                                        {
                                            result = MessageBox.Show("查無檢驗記錄資料，操作失敗", "警告", buttons, MessageBoxIcon.Warning);
                                            Console.WriteLine("查無檢驗記錄資料，操作失敗");
                                            KvHostLinkLog(employeeID, mo, mn, itemNumber, "查無檢驗記錄資料，操作失敗", "警告", "");
                                            return;
                                        }
                                        getQCDataCollectionDatareader.Close();
                                        //品檢人員找尋未檢驗的首件/自主檢查記錄詳細資料
                                        String getQCTestData = "select qc_id,itemSN,testItem,testUnit,standardValue,upperLimit,lowerLimit,testTool,flag,ipqc1,ipqc2,ipqc3 from " + DBName + ".dbo.QCDataCollectionContent where "
                                             + "qc_id='" + qcId + "' order by itemSN asc";
                                        using (SqlCommand getQCTestDatacommand = new SqlCommand(getQCTestData, connection))
                                        {
                                            using (SqlDataReader getQCTestDatareader = getQCTestDatacommand.ExecuteReader())
                                            {
                                                while (getQCTestDatareader.Read())
                                                {
                                                    IpqcItem item = new IpqcItem();
                                                    item.JYT012b003 = getQCTestDatareader["itemSN"].ToString();
                                                    item.JYT012b005 = getQCTestDatareader["testItem"].ToString();
                                                    item.JYT012b006 = getQCTestDatareader["testUnit"].ToString();
                                                    item.JYT012b007 = getQCTestDatareader["standardValue"].ToString();
                                                    item.JYT012b008 = getQCTestDatareader["upperLimit"].ToString();
                                                    item.JYT012b009 = getQCTestDatareader["lowerLimit"].ToString();
                                                    ipqcItemList.Add(item);
                                                }
                                                getQCTestDatareader.Close();
                                            }
                                        }

                                        //寫品檢成品檢驗記錄(pqc1)
                                        for (int i = 0; i < ipqcItemList.Count; i++)
                                        {
                                            IpqcItem itemList = (IpqcItem)ipqcItemList[i];
                                            String updPqc1Data = "update " + DBName + ".dbo.QCDataCollectionContent set pqc1=@pqc1"
                                                 + " where qc_id='" + qcId + "' and itemSN='" + itemList.JYT012b003 + "'";
                                            using (SqlCommand updPqc1Datacommand = new SqlCommand(updPqc1Data, connection))
                                            {
                                                updPqc1Datacommand.Parameters.AddWithValue("@pqc1", ipqc1ResultList[i]);
                                                int updateRowsAffected = updPqc1Datacommand.ExecuteNonQuery();
                                                updPqc1RowsAffected = updPqc1RowsAffected + updateRowsAffected;
                                            }
                                        }

                                        //寫品檢成品檢驗記錄(pqc2)
                                        for (int i = 0; i < ipqcItemList.Count; i++)
                                        {
                                            IpqcItem itemList = (IpqcItem)ipqcItemList[i];
                                            String updPqc2Data = "update " + DBName + ".dbo.QCDataCollectionContent set pqc2=@pqc2"
                                                 + " where qc_id='" + qcId + "' and itemSN='" + itemList.JYT012b003 + "'";
                                            using (SqlCommand updPqc2Datacommand = new SqlCommand(updPqc2Data, connection))
                                            {
                                                updPqc2Datacommand.Parameters.AddWithValue("@pqc2", ipqc2ResultList[i]);
                                                int updateRowsAffected = updPqc2Datacommand.ExecuteNonQuery();
                                                updPqc2RowsAffected = updPqc2RowsAffected + updateRowsAffected;
                                            }
                                        }

                                        //寫品檢成品檢驗記錄(pqc3)
                                        for (int i = 0; i < ipqcItemList.Count; i++)
                                        {
                                            IpqcItem itemList = (IpqcItem)ipqcItemList[i];
                                            String updPqc3Data = "update " + DBName + ".dbo.QCDataCollectionContent set pqc3=@pqc3"
                                                 + " where qc_id='" + qcId + "' and itemSN='" + itemList.JYT012b003 + "'";
                                            using (SqlCommand updPqc3Datacommand = new SqlCommand(updPqc3Data, connection))
                                            {
                                                updPqc3Datacommand.Parameters.AddWithValue("@pqc3", ipqc3ResultList[i]);
                                                int updateRowsAffected = updPqc3Datacommand.ExecuteNonQuery();
                                                updPqc3RowsAffected = updPqc3RowsAffected + updateRowsAffected;
                                            }
                                        }

                                        if (ZF16ipqc1.Equals("1") && ZF18ipqc2.Equals("1") && ZF20ipqc3.Equals("1"))
                                        {
                                            defective = "合格";
                                        }
                                        else
                                        {
                                            defective = "不合格";
                                        }
                                        //寫判定結果
                                        if (updPqc1RowsAffected > 0 && updPqc2RowsAffected > 0 && updPqc3RowsAffected > 0)
                                        {
                                            String insPqcResult = "insert into " + DBName + ".dbo.QCDataCollectionResult(qc_id,defective,determination,supervisor,inspectors,handleResult,pqcCREATOR)"
                                                 + " values(@qc_id,@defective,@determination,@supervisor,@inspectors,@handleResult,@pqcCREATOR)";
                                            // 資料參數
                                            SqlParameter[] QCDataCollectionResultparameters = {
                                                new SqlParameter("@qc_id", qcId),
                                                new SqlParameter("@defective", ""),
                                                new SqlParameter("@determination", defective),
                                                new SqlParameter("@supervisor", ""),
                                                new SqlParameter("@inspectors", employeeID),
                                                new SqlParameter("@handleResult", ""),
                                                new SqlParameter("@pqcCREATOR", employeeID)
                                            };
                                            using (SqlCommand insPqcResultcommand = new SqlCommand(insPqcResult, connection))
                                            {
                                                // 設定參數
                                                insPqcResultcommand.Parameters.AddRange(QCDataCollectionResultparameters);
                                                // 執行 INSERT 操作
                                                int insRowsAffected = insPqcResultcommand.ExecuteNonQuery();
                                                if (insRowsAffected > 0)
                                                {
                                                    //回寫品檢有檢驗flag
                                                    String updpqcStatusflag = "update " + DBName + ".dbo.QCDataCollection set pqcStatus=@pqcStatus where qc_id='" + qcId + "'";
                                                    using (SqlCommand updpqcStatusflagcommand = new SqlCommand(updpqcStatusflag, connection))
                                                    {
                                                        updpqcStatusflagcommand.Parameters.AddWithValue("@pqcStatus", "1");
                                                        int updateRowsAffected = updpqcStatusflagcommand.ExecuteNonQuery();
                                                        Console.WriteLine("更新品檢人員檢驗flag資料筆數=" + updateRowsAffected + "筆");
                                                    }
                                                    //寫上傳完成信號值"1"到ZF14暫存器(PC→PLC)
                                                    int[] wzf14 = new int[1] { 1 };
                                                    KHST.IntToByte(ref writeBuf, wzf14, 1, 0);
                                                    err = KHL.KHLWriteDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 14, 1, writeBuf);
                                                    if (err != 0)
                                                    {
                                                        result = MessageBox.Show("PLC連線發生異常，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                                                        Console.WriteLine("PLC連線發生異常，操作失敗");
                                                        Console.WriteLine(err);
                                                        KvHostLinkLog(employeeID, mo, mn, itemNumber, "PLC連線發生異常，操作失敗", "錯誤", err.ToString());
                                                        return;
                                                    }
                                                    result = MessageBox.Show("品檢人員上傳資料完成", "結果", buttons);
                                                    Console.WriteLine("品檢人員寫入資料成功");
                                                    KvHostLinkLog(employeeID, mo, mn, itemNumber, "品檢人員寫入資料成功", "結果", "");
                                                    //確認品檢人員需檢驗次數
                                                    string getQcCountQuery = "SELECT  count(*) as count FROM " + DBName + ".dbo.QCDataCollection where pqcStatus='0' and manufactureOrder='" + mo + "'"
                                                       + " and manufactureNo='" + mn + "'";
                                                    using (SqlCommand getQcCountQuerycommand = new SqlCommand(getQcCountQuery, connection))
                                                    {
                                                        using (SqlDataReader getQcCountQueryreader = getQcCountQuerycommand.ExecuteReader())
                                                        {
                                                            if (getQcCountQueryreader.Read())
                                                            {
                                                                qcId = getQcCountQueryreader["count"].ToString();
                                                                result = MessageBox.Show("品檢總共還需要檢測" + qcId + "筆資料", "結果", buttons);
                                                                Console.WriteLine("品檢總共還需要檢測" + qcId + "筆資料");
                                                                getQcCountQueryreader.Close();
                                                            }
                                                        }
                                                    }
                                                    return;
                                                }
                                                else
                                                {
                                                    result = MessageBox.Show("品檢人員上傳資料失敗", "警告", buttons, MessageBoxIcon.Warning);
                                                    Console.WriteLine("品檢人員上傳資料失敗。");
                                                    KvHostLinkLog(employeeID, mo, mn, itemNumber, "品檢人員上傳資料失敗", "警告", "");
                                                    return;
                                                }
                                            }
                                        }                                                    
                                    }
                                }
                            }
                            else
                            {
                                result = MessageBox.Show("製令單別單號找不到需要檢測的資料", "結果", buttons);
                                Console.WriteLine("製令單別單號找不到需要檢測的資料");
                                KvHostLinkLog(employeeID, mo, mn, itemNumber, "製令單別單號找不到需要檢測的資料", "結果", "");
                                return;
                            }
                        }
                    }                                                                          
                }
                else
                {
                    result = MessageBox.Show("現場人員和品檢人員回傳值有誤，操作失敗", "警告", buttons, MessageBoxIcon.Warning);
                    Console.WriteLine("現場人員和品檢人員回傳值有誤");
                    KvHostLinkLog(employeeID, mo, mn, itemNumber, "現場人員和品檢人員回傳值有誤，操作失敗", "警告", "");
                    return;
                }                
            }
            catch (Exception ex)
            {
                result = MessageBox.Show("系統有誤，操作失敗", "錯誤", buttons, MessageBoxIcon.Error);
                Console.WriteLine("An error occurred: " + ex.Message);
                KvHostLinkLog(employeeID, mo, mn, itemNumber, "系統有誤，操作失敗" + ex.Message, "錯誤", "");
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

            void KvHostLinkLog(string employeeID, string manufactureOrder, string manufactureNo,
                string itemNumber, string Msg, string Status, string KvHostLinkEErrorCode)
            {
                Logconnection = new SqlConnection(LogconnectionString);

                Logconnection.Open();

                string insertQuery = "INSERT INTO KvHostLinkLog (employeeID,manufactureOrder,manufactureNo,itemNumber, Msg," +
                     " Status,KvHostLinkEErrorCode) VALUES (@employeeID, @manufactureOrder, @manufactureNo,@itemNumber," +
                     "@Msg,@Status,@KvHostLinkEErrorCode)";
                using (SqlCommand insertQuerycommand = new SqlCommand(insertQuery, Logconnection))
                {
                    insertQuerycommand.Parameters.AddWithValue("@employeeID", employeeID);
                    insertQuerycommand.Parameters.AddWithValue("@manufactureOrder", manufactureOrder);
                    insertQuerycommand.Parameters.AddWithValue("@manufactureNo", manufactureNo);
                    insertQuerycommand.Parameters.AddWithValue("@itemNumber", itemNumber);
                    insertQuerycommand.Parameters.AddWithValue("@Msg", Msg);
                    insertQuerycommand.Parameters.AddWithValue("@Status", Status);
                    insertQuerycommand.Parameters.AddWithValue("@KvHostLinkEErrorCode", KvHostLinkEErrorCode);
                    insertQuerycommand.ExecuteNonQuery();
                }
                if (Logconnection != null)
                {
                    Logconnection.Close();
                }
            }         
        }
    }
}
