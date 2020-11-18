using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;

namespace OPR
{
    public class cSQL
    {
        OleDbConnection cn;
        string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\OPR.mdb";


        public void cSQL_init(string ConnectionString)
        {
            this.ConnectionString = ConnectionString;
        }


        public void Connect()
        {
            try
            {
                cn = new OleDbConnection();
                cn.ConnectionString = ConnectionString;
                cn.Open();
            }

            catch (Exception exp)
            {
                System.Windows.Forms.MessageBox.Show("Can't connect to Server" + exp.Message);
                return;
            }
        }

        public void Disconnect()
        {
            try
            {
                cn.Close();
            }
            catch (Exception exp)
            {
                System.Windows.Forms.MessageBox.Show("Can't connect to Server" + exp.Message);
                return;
            }
        }

        public DataTable Query(string SQLstring)
        {
            OleDbCommand cm = null;
            try
            {
                cm = new OleDbCommand(SQLstring, cn);
                OleDbDataAdapter da = new OleDbDataAdapter();

                da.SelectCommand = cm;
                DataTable _table = new DataTable();
                da.Fill(_table);
                da.Dispose();
                cm.Dispose();
                return _table;
            }
            catch (Exception exp)
            {
                System.Windows.Forms.MessageBox.Show(exp.Message);
                cm.Dispose();
                return null;
            }

        }

        public void QueryBox(string SQLstring, ComboBox _box)
        {
            OleDbCommand cm = null;
            try
            {
                cm = new OleDbCommand(SQLstring, cn);
                OleDbDataAdapter da = new OleDbDataAdapter();
                DataSet ds = new DataSet();

                da.SelectCommand = cm;
                da.Fill(ds);
                _box.DataSource = ds.Tables[0];

                da.Dispose();
                cm.Dispose();
            }
            catch (Exception exp)
            {
                System.Windows.Forms.MessageBox.Show(exp.Message);
                cm.Dispose();
            }

        }

        public int SetCommand(string SQLstring)
        {
            OleDbCommand cm = null;
            int res = -1;
            try
            {
                cm = new OleDbCommand(SQLstring, cn);
                res = Int32.Parse(cm.ExecuteScalar().ToString());
                cm.Dispose();
            }
            catch (Exception exp)
            {
                System.Windows.Forms.MessageBox.Show("Base Drive.\n" + exp.Message);
                cm.Dispose();
                return res;
            }
            return res;

        }

        public string SetCommandStr(string SQLstring)
        {
            OleDbCommand cm = null;
            object res = null;
            string nul = "";
            try
            {
                cm = new OleDbCommand(SQLstring, cn);
                res = cm.ExecuteScalar();
                cm.Dispose();
            }
            catch (Exception exp)
            {
                System.Windows.Forms.MessageBox.Show("Base Drive.\n" + exp.Message);
                cm.Dispose();
                return res.ToString();
            }
            if (res != null) return res.ToString();
            else return nul;
        }


        public void SetCommandUpIn(string SQLstring)
        {
            OleDbCommand cm = null;
            try
            {
                cm = new OleDbCommand(SQLstring, cn);
                cm.ExecuteNonQuery();
                cm.Dispose();
               // MessageBox.Show("Выполнено!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception exp)
            {
                System.Windows.Forms.MessageBox.Show("Base Drive.\n" + exp.Message);
                cm.Dispose();
            }

        }
    }
}
