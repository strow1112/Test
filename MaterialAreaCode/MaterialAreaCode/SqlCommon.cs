using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.Net.Sockets;

namespace MaterialAreaCode
{
    class SqlCommon
    {
        private SqlDataAdapter sDataAdapter;
        public SqlConnection sConn;
        private SqlTransaction sTr;

        public SqlCommon()
        {
            string myIP = "";
            IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());

            foreach (IPAddress ip in host.AddressList)
            {
                if(ip.AddressFamily == AddressFamily.InterNetwork)
                {
                    myIP = ip.ToString();
                }
            }

            if (myIP.Substring(0, 2) == "10")
            {
                this.sConn = new SqlConnection("Server=10.224.189.17,1432;database=ILSHINMES;uid=sa;pwd=ilshin");
            }

            else if (myIP.Substring(0, 2) == "19")
            {
                this.sConn = new SqlConnection("Server=192.168.0.110,1432;database=ILSHINMES;uid=sa;pwd=ilshin");
            }

        }

        public DataTable getTable(string strSQL)
        {
            DataTable dt = new DataTable();
            try
            {
                //Console.WriteLine("1111=" + strSQL);
                this.sDataAdapter = new SqlDataAdapter(strSQL, this.sConn);
                //this.oDataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;
                this.sDataAdapter.Fill(dt);
                this.sConn.Close();
                this.sConn.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                this.sConn.Close();
                this.sConn.Dispose();
            }
            return dt;
        }

        public object getSimpleScalar(string strSQL)
        {
            object obj;

            try
            {
                SqlCommand cmd = new SqlCommand(strSQL, this.sConn);

                this.sConn.Open();
                obj = cmd.ExecuteScalar();
                this.sConn.Close();
                this.sConn.Dispose();
                if (obj == null)
                {
                    obj = "";
                }
                return obj;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                this.sConn.Close();
                this.sConn.Dispose();
                return null;
            }
        }

        public int execNonQuery(string strSQL)
        {
            int affectRows = 0;
            bool isConnOpened = false;

            if (this.sConn.State == ConnectionState.Open)
                isConnOpened = true;
            else
            {
                this.sConn.Open();
            }

            try
            {
                this.sTr = this.sConn.BeginTransaction();
                SqlCommand cmd = new SqlCommand(strSQL, this.sConn, this.sTr);
                cmd.CommandType = CommandType.Text;
                affectRows = cmd.ExecuteNonQuery();
                //Console.WriteLine(">>>>>>>>>>>>>>>>> isConnOpened =" + isConnOpened);
                if (!isConnOpened)
                {
                    this.sTr.Commit();
                    this.sConn.Close();
                }
                return affectRows;
            }
            catch (SqlException oex)
            {
                this.sTr.Rollback();
                this.sConn.Close();
                throw oex;
            }
            catch (Exception ex)
            {
                this.sTr.Rollback();
                this.sConn.Close();
                throw ex;
            }
        }
    }
}
