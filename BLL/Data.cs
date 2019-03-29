using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    public class Data
    {
        private DAL.Data _objDAL = new DAL.Data();

        public DataTable getDataTemp()
        {
            return _objDAL.getDataTemp();
        }

        public DataTable getDataTemp2under()
        {
            return _objDAL.getDataTemp2under();
        }

        public DataTable CountOn()
        {
            return _objDAL.CountOn();
        }

        public DataTable CountUnder()
        {
            return _objDAL.CountUnder();
        }   

        public DataTable Max()
        {
            return _objDAL.Max();
        } 

        public DataTable NotPrice()
               {
                 return  _objDAL.NotPrice();
               } 

        public DataTable NotPrice2under()
             {
                 return _objDAL.NotPrice2under();
             } 

        public DataTable CheckFilename(string filename)
           {
                   return _objDAL.CheckFilename(filename);
               } 

        public void InsertFilename(string filename)
        {
            _objDAL.InsertFilename(filename);
        } 

        public DataTable BindFilename()
         {
             return _objDAL.BindFilename();
         }  

        public void ExecuteNonQuery(string sql)
            {
                _objDAL.ExecuteNonQuery(sql);
            } 

        public DataTable ExecuteDatatable(string sql)
        {
            return _objDAL.ExecuteDatatable(sql);
        }
    }
}
