using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL
{
    public class Data
    {

        public DataTable getDataTemp()
        {
            try
            {
                DataSet ds = new DataSet();
                String sql;
                sql = "Select * From Temp";
                Class.clsDB db = new Class.clsDB();
                ds = db.ExecuteDataSet(sql);
                db.Close();
                return ds.Tables[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }

        public DataTable getDataTemp2under()
        {
            try
            {
                DataSet ds = new DataSet();
                String sql;
                sql = "Select * From Temp2under";
                Class.clsDB db = new Class.clsDB();
                ds = db.ExecuteDataSet(sql);
                db.Close();
                return ds.Tables[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }

        public DataTable CountOn()
        {
            try
            {
                DataSet ds = new DataSet();
                String sql;
                sql = "select count(*) as c from temp where price > 0 ";
                Class.clsDB db = new Class.clsDB();
                ds = db.ExecuteDataSet(sql);
                db.Close();
                return ds.Tables[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }

        public DataTable CountUnder()
        {
            try
            {
                DataSet ds = new DataSet();
                String sql;
                sql = "select count(*) as c from temp2under where price > 0";
                Class.clsDB db = new Class.clsDB();
                ds = db.ExecuteDataSet(sql);
                db.Close();
                return ds.Tables[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }

        public DataTable NotPrice()
        {
            try
            {
                DataSet ds = new DataSet();
                String sql;
                sql = "select no,price  from temp where price = 0";
                Class.clsDB db = new Class.clsDB();
                ds = db.ExecuteDataSet(sql);
                db.Close();
                return ds.Tables[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }

        public DataTable NotPrice2under()
        {
            try
            {
                DataSet ds = new DataSet();
                String sql;
                sql = "select no,price  from temp2under where price = 0";
                Class.clsDB db = new Class.clsDB();
                ds = db.ExecuteDataSet(sql);
                db.Close();
                return ds.Tables[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }

        public DataTable Max()
        {
            try
            {
                DataSet ds = new DataSet();
                String sql;
                sql = "select avg(price) from temp";
                Class.clsDB db = new Class.clsDB();
                ds = db.ExecuteDataSet(sql);
                db.Close();
                return ds.Tables[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }

        public DataTable CheckFilename(string filename)
        {
            try
            {
                DataSet ds = new DataSet();
                String sql;
                sql = "select * from tb_file_name where file_name= '" + filename + "'";
                Class.clsDB db = new Class.clsDB();
                ds = db.ExecuteDataSet(sql);
                db.Close();
                return ds.Tables[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }

        public void InsertFilename(string filename)
        {
            Class.clsDB db = new Class.clsDB();
            string sql = "insert into tb_file_name(file_name) values('" + filename + "')";

            db.ExecuteNonQuery(sql);
            db.Close();
        }

        public DataTable BindFilename()
        {
            try
            {
                DataSet ds = new DataSet();
                String sql;
                sql = "select * from tb_file_name ";
                Class.clsDB db = new Class.clsDB();
                ds = db.ExecuteDataSet(sql);
                db.Close();
                return ds.Tables[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }

        public void ExecuteNonQuery(string sql)
        {
            Class.clsDB db = new Class.clsDB();
            db.ExecuteNonQuery(sql);
            db.Close();
        }

        public DataTable ExecuteDatatable(string sql)
        {
            try
            {
                DataSet ds = new DataSet();
            
                Class.clsDB db = new Class.clsDB();
                ds = db.ExecuteDataSet(sql);
                db.Close();
                return ds.Tables[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }
    }
}
