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
            using (nemanpingEntities1 can = new nemanpingEntities1())
            {
                var user = can.UserTable.First(u => u.Name.Equals(name) && u.Pwd.Equals(pwd));
                return true;
            }
            
        }
   
    }
}
