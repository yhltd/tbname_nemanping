using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using clsBuiness;

namespace logic
{
    public class clsAllnewLogic
    {
        public bool Login(string name , string pwd) 
        {
            using (nemanpingEntities3 can = new nemanpingEntities3())
            {
                var user = can.UserTable.First(u => u.Name.Equals(name) && u.Pwd.Equals(pwd));
                return true;
            }
            
        }
        public List<Sehao> selectSehao() 
        {
            using (nemanpingEntities3 can = new nemanpingEntities3())
            {
                List<Sehao> sehao = new List<Sehao>();
                var select = from s in can.Sehao select new { s.Name, s.SeHao1 };
                
                
                foreach (var item in select)
                {
                    Sehao sh = new Sehao();
                    sh.Name = item.Name;
                    sh.SeHao1 = item.SeHao1;
                    sehao.Add(sh);
                }
                return sehao;
            }
        }
   
    }
}
