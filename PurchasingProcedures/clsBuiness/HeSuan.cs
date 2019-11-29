using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace clsBuiness
{
    public class HeSuan
    {
        public string LOT{get;set;}
        public string 订单数量 { get; set; }
        public string Name { get; set; }
        public string 实际出口数量 { get; set; }
        public string 色号颜色 { get; set; }
        public string 单价 { get; set; }
        public string 预计单耗 { get; set; }
        public string 预计成本 { get; set; }
        public string 预计用量 { get; set; }
        public string 库存 { get; set; }
        public string 订量 { get; set; }
        public string 实际到货量 { get; set; }
        public string 实际到货金额 { get; set; }
        public string 剩余数量 { get; set; }
        public string 平均单耗 { get; set; }
        public string 结算成本 { get; set; }
        public string 小计 { get; set; }
    }
    public class softTime_info
    {
        public string _id { get; set; }//玩法种类

        public string starttime { get; set; }//玩法种类
        public string name { get; set; }//玩法种类
        public string endtime { get; set; }//玩法种类
        public string soft_name { get; set; }//玩法种类
        public string denglushijian { get; set; }//玩法种类


        public string password { get; set; }//玩法种类
        public string pid { get; set; }//玩法种类
        public string mark1 { get; set; }//玩法种类
        public string mark2 { get; set; }//玩法种类
        public string mark3 { get; set; }//玩法种类
        public string mark4 { get; set; }//玩法种类
        public string mark5 { get; set; }//玩法种类
    }
}
