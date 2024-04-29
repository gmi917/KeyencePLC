using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HostLinkSample_V200
{
    internal class IpqcItem
    {
        public String JYT012b003 { get; set; }//項次
        public String JYT012b005 {  get; set; }//檢驗項目
        public String JYT012b006 {  get; set; }//檢驗單位
        public String JYT012b007 {  get; set; }//檢驗標準值
        public String JYT012b008 {  get; set; }//上限值
        public String JYT012b009 {  get; set; }//下限值
        public String JYT012b010 {  get; set; }//檢驗量具
        public String UDF03 {  get; set; }//檢驗類型
        public String min {  get; set; }
        public String max { get; set; }
        public String ipqc1 { get; set; }//首件/自主檢查記錄實測狀況1
        public String ipqc2 {  get; set; }//首件/自主檢查記錄實測狀況2
        public String ipqc3 {  get; set; }//首件/自主檢查記錄實測狀況3
        public String pqc1 {  get; set; }//品檢成品檢驗記錄實測狀況1
        public String pqc2 {  get; set; }//品檢成品檢驗記錄實測狀況2
        public String pqc3 {  get; set; }//品檢成品檢驗記錄實測狀況3

    }
}
