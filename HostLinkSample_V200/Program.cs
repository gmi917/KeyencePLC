using System;
using System.Collections;
using KHL = KvHostLink.KvHostLink;
using KHST = KvStruct.KvStruct;
using System.Data.SqlClient;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using System.Windows.Forms;

namespace cs_dll_sample
{
    class Program
    {
        static void Main(string[] args)
        {
            int err = 0;
            int sock = 0;
            err = KHL.KHLInit();
            string connectionString = "Data Source=123.456.789.0;Initial Catalog=myDataBase;User ID=myUsername;Password=myPassword;";
            SqlConnection connection = null;
            ArrayList ipqc1ResultList = new ArrayList();
            ArrayList ipqc2ResultList = new ArrayList();
            ArrayList ipqc3ResultList = new ArrayList();
            float[] rdEM3000float = new float[2];
            float[] rdEM4000float = new float[2];
            float[] rdEM5000float = new float[2];
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

            //讀現場人員 or 品檢人員(ZF0)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 0, 1, readBuf);
            Console.WriteLine("讀現場人員 or 品檢人員(ZF0)");
            int[] rdZF0Str = new int[1];
            KHST.ByteToInt(ref rdZF0Str, readBuf, 1, 0);
            for (int i = 0; i < 1; i++) Console.WriteLine("\tZF0:{0}", rdZF0Str[i]);

            //OP1 or OP2(ZF4)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 4, 1, readBuf);
            Console.WriteLine("讀OP1 or OP2(ZF4)");
            int[] rdZF4Str = new int[1];
            KHST.ByteToInt(ref rdZF4Str, readBuf, 1, 0);
            for (int i = 0; i < 1; i++) Console.WriteLine("\tZF4:{0}", rdZF4Str[i]);

            //量測數據完成信號(ZF8)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 8, 1, readBuf);
            Console.WriteLine("讀量測數據完成信號(ZF8)");
            int[] rdZF8Str = new int[1];
            KHST.ByteToInt(ref rdZF8Str, readBuf, 1, 0);
            for (int i = 0; i < 1; i++) Console.WriteLine("\tZF8:{0}", rdZF8Str[i]);

            //工件一量測數據OK or NG信號(ZF16)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 16, 1, readBuf);
            Console.WriteLine("讀工件一量測數據OK or NG信號(ZF16)");
            int[] rdZF16Str = new int[1];
            KHST.ByteToInt(ref rdZF16Str, readBuf, 1, 0);
            for (int i = 0; i < 1; i++) Console.WriteLine("\tZF16:{0}", rdZF16Str[i]);

            //工件二量測數據OK or NG信號(ZF18)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 18, 1, readBuf);
            Console.WriteLine("讀工件二量測數據OK or NG信號(ZF18)");
            int[] rdZF18Str = new int[1];
            KHST.ByteToInt(ref rdZF18Str, readBuf, 1, 0);
            for (int i = 0; i < 1; i++) Console.WriteLine("\tZF18:{0}", rdZF18Str[i]);

            //工件三量測數據OK or NG信號(ZF20)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 20, 1, readBuf);
            Console.WriteLine("讀工件三量測數據OK or NG信號(ZF20)");
            int[] rdZF20Str = new int[1];
            KHST.ByteToInt(ref rdZF20Str, readBuf, 1, 0);
            for (int i = 0; i < 1; i++) Console.WriteLine("\tZF20:{0}", rdZF20Str[i]);


            //讀工單號碼(ZF30~ZF37)
            err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 30, 8, readBuf);
            if (err != 0)
            {
                Console.WriteLine(err);
            }
            Console.WriteLine("讀工單號碼(ZF30~ZF37)");
            byte[] rdZF30Str = new byte[20];
            KHST.ByteToString(ref rdZF30Str, readBuf, 1, 0, 20);
            //只有ZF和EM暫存器編號是用十進制累加,其餘暫存器編號則為16進制累加
            //存英文字串到暫存器需要使用ASCII16BIT格式,但是讀出來是二進制的值,還需要轉碼成文字才能顯示該字串
            Console.WriteLine(System.Text.Encoding.GetEncoding(65001).GetString(rdZF30Str));

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
            Console.WriteLine(System.Text.Encoding.GetEncoding(65001).GetString(rdZF40Str));


            Console.WriteLine("Result ipqc1ResultList List:");
            foreach (float value in ipqc1ResultList)
            {
                Console.WriteLine(value);
            }

            Console.WriteLine("Result ipqc2ResultList List:");
            foreach (float value in ipqc2ResultList)
            {
                Console.WriteLine(value);
            }

            Console.WriteLine("Result ipqc3ResultList List:");
            foreach (float value in ipqc3ResultList)
            {
                Console.WriteLine(value);
            }

            try
            {
                //connection = new SqlConnection(connectionString);

                //connection.Open();
                //foreach (float value in ipqc1ResultList)
                //{
                //    string query = $"INSERT INTO tableName (FloatValue) VALUES (@Value)";
                //    using (SqlCommand command = new SqlCommand(query, connection))
                //    {
                //        command.Parameters.AddWithValue("@Value", value);
                //        command.ExecuteNonQuery();
                //    }
                //}

                MessageBoxButtons buttons = MessageBoxButtons.OK;
                DialogResult result;

                // Displays the MessageBox.
                result = MessageBox.Show("上傳資料完成", "結果", buttons);

                //寫上傳完成信號(PC→PLC)寫1筆資料(ZF14)
                int[] wzf14 = new int[1] { 0 };
                KHST.IntToByte(ref writeBuf, wzf14, 1, 0);
                err = KHL.KHLWriteDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 14, 1, writeBuf);

                //讀上傳完成信號(PC→PLC)讀1筆資料(ZF14)
                //err = KHL.KHLReadDevicesAsWords(sock, KvHostLink.KHLDevType.DEV_ZF, 14, 1, readBuf);
                //Console.WriteLine("讀1筆上傳完成信號(ZF14)");
                //byte[] rdZF14Str = new byte[1];
                //KHST.ByteToString(ref rdZF14Str, readBuf, 1, 0);
                //for (int i = 0; i < 1; i++) Console.WriteLine("\tZF14:{0}", rdZF14Str[i]);
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
