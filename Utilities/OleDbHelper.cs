using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Threading.Tasks;

namespace LogictecaTest.Utilities
{
    public class OleDbHelper
    {
        protected OleDbConnection conn = new OleDbConnection();
        protected OleDbCommand comm = new OleDbCommand();

        string dbFile = "";
        public OleDbHelper(string file)
        {
            dbFile = file;
            //var app = new Microsoft.Office.Interop.Excel.Application();
            //var wb = app.Workbooks.Add();
            //wb.SaveAs(file);
            //wb.Close();
        }

        public void openConnection()
        {
            if (conn.State == ConnectionState.Closed)
            {
                conn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbFile + ";Persist Security Info=False;";
                comm.Connection = conn;
                try
                {
                    conn.Open();
                }
                catch (Exception e)
                { throw new Exception(e.Message); }

            }

        }
        public void closeConnection()
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
                conn.Dispose();
                comm.Dispose();
            }
        }
        /// <summary>
        /// Execute SQL statement
        /// </summary>
        /// <param name="sqlstr"></param>
        public int ExecuteNonQuery(string sqlstr)
        {
            int i = 0;
            try
            {
                openConnection();
                comm.CommandType = CommandType.Text;
                comm.CommandText = sqlstr;
                i = comm.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {

                closeConnection();

            }
            return i;
        }
 
 
        /// <summary>
                 /// Execute SQL statement
        /// </summary>
        /// <param name="sqlstr"></param>
        public object executeScalarSql(string sqlstr)
        {
            object o;
            try
            {
                openConnection();
                comm.CommandType = CommandType.Text;
                comm.CommandText = sqlstr;
                o = comm.ExecuteScalar();
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {

                closeConnection();

            }

            return o;
        }


        public int bulkInsert(DataTable dtExcel,string sheetName)
        {
            int counter = 0;
            string qryFieldValue = "", qryFieldValueTemp="";
            openConnection();
            for (int r = 0; r < dtExcel.Rows.Count; r++)
            {
                qryFieldValue = "";
                for (int j = 0; j < dtExcel.Columns.Count; j++)
                {
                    qryFieldValueTemp = dtExcel.Rows[r][j].ToString();
                    qryFieldValue = qryFieldValue + (qryFieldValue.Trim() != "" ? ", '" : "'") + qryFieldValueTemp.Replace("'", "''") + "'";
                }

                try
                {
                    comm.CommandType = CommandType.Text;
                    comm.CommandText = "Insert into [" + sheetName + "$] Values (" + qryFieldValue + ")";
                    counter += comm.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    throw new Exception(e.Message);
                }
                finally
                {


                }

            }
            closeConnection();
            return counter;
        }

        /// <summary>
        /// Returns the OLEDBDATAREADER object of the specified SQL statement, please pay attention to the object when you are using.
        /// </summary>
        /// <param name="sqlstr"></param>
        /// <returns></returns>
        public OleDbDataReader dataReader(string sqlstr)
        {
            OleDbDataReader dr = null;
            try
            {
                openConnection();
                comm.CommandText = sqlstr;
                comm.CommandType = CommandType.Text;

                dr = comm.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch
            {
                try
                {
                    dr.Close();
                    closeConnection();
                }
                catch { }
            }
            return dr;
        }
        /// <summary>
        /// Return the OLEDBDATAREADER object of the specified SQL statement, please pay attention to shutting down when using
        /// </summary>
        /// <param name="sqlstr"></param>
        /// <param name="dr"></param>
        public void dataReader(string sqlstr, ref OleDbDataReader dr)
        {
            try
            {
                openConnection();
                comm.CommandText = sqlstr;
                comm.CommandType = CommandType.Text;
                dr = comm.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch
            {
                try
                {
                    if (dr != null && !dr.IsClosed)
                        dr.Close();
                }
                catch
                {
                }
                finally
                {
                    closeConnection();
                }
            }
        }
        /// <summary>
                 // / Return to the DataSet of the specified SQL statement
        /// </summary>
        /// <param name="sqlstr"></param>
        /// <returns></returns>
        public DataSet dataSet(string sqlstr)
        {
            DataSet ds = new DataSet();
            OleDbDataAdapter da = new OleDbDataAdapter();
            try
            {
                openConnection();
                comm.CommandType = CommandType.Text;
                comm.CommandText = sqlstr;
                da.SelectCommand = comm;
                da.Fill(ds);

            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                closeConnection();
            }
            return ds;
        }
        /// <summary>
                 // / Return to the DataSet of the specified SQL statement
        /// </summary>
        /// <param name="sqlstr"></param>
        /// <param name="ds"></param>
        public void dataSet(string sqlstr, ref DataSet ds)
        {
            OleDbDataAdapter da = new OleDbDataAdapter();
            try
            {
                openConnection();
                comm.CommandType = CommandType.Text;
                comm.CommandText = sqlstr;
                da.SelectCommand = comm;
                da.Fill(ds);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                closeConnection();
            }
        }
        /// <summary>
                 /// Return to the DataTable specified by the SQL statement
        /// </summary>
        /// <param name="sqlstr"></param>
        /// <returns></returns>
        public DataTable dataTable(string sqlstr)
        {
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter();
            try
            {
                openConnection();
                comm.CommandType = CommandType.Text;
                comm.CommandText = sqlstr;
                da.SelectCommand = comm;
                da.Fill(dt);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                closeConnection();
            }
            return dt;
        }
        /// <summary>
                 /// Return to the DataTable specified by the SQL statement
        /// </summary>
        /// <param name="sqlstr"></param>
        /// <param name="dt"></param>
        public void dataTable(string sqlstr, ref DataTable dt)
        {
            OleDbDataAdapter da = new OleDbDataAdapter();
            try
            {
                openConnection();
                comm.CommandType = CommandType.Text;
                comm.CommandText = sqlstr;
                da.SelectCommand = comm;
                da.Fill(dt);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                closeConnection();
            }
        }
        /// <summary>
        /// Returns the DataView specified by the SQL statement
        /// </summary>
        /// <param name="sqlstr"></param>
        /// <returns></returns>
        public DataView dataView(string sqlstr)
        {
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataView dv = new DataView();
            DataSet ds = new DataSet();
            try
            {
                openConnection();
                comm.CommandType = CommandType.Text;
                comm.CommandText = sqlstr;
                da.SelectCommand = comm;
                da.Fill(ds);
                dv = ds.Tables[0].DefaultView;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                closeConnection();
            }
            return dv;
        }
    }
}

