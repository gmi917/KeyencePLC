using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HostLinkSample_V200
{
    internal class JYT012
    {
        public String manufactureOrder {  get; set; }//製令單別
        public String manufactureNo {  get; set; }//製令單號
        public String processCode {  get; set; }//製程代號
        public String productionLine {  get; set; }//生產線別
        public String JYT012a002 {  get; set; }//表單單號
        public String JYT012a003 {  get; set; }//品名
        public String UDF02 {  get; set; }//規格
        public String JYT012a004 {  get; set; }//圖號
        public String JYT012a005 {  get; set; }//料號
        public String JYT012b006 {  get; set; }//檢驗單位
        public String customerCode {  get; set; }//客戶代號
        public String firstItemDate {  get; set; }//首件日期
        public String firstItemStaff {  get; set; }//首件人員
        public String predictQty {  get; set; }//預計產量
        public String machineNumber {  get; set; }//機台編號
        public String JYT012a006 {  get; set; }//版次資訊
        public String UDF01 {  get; set; }//圖檔檔名
        public String fullImageName {  get; set; }//URL+圖檔檔名
        public String qc_id {  get; set; }
    }
}
