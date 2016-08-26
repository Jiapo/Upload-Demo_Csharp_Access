using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using AccessLib;

namespace Access
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string strConnection = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=" + Application.StartupPath + "\\demo.mdb";
        OleDbConnection myCon;
        OleDbCommand myCommand;
        string sql;
        string DemoTable = "table1";      //表名
        string num;
        DataSet ds;
        DataTable dt;
        OleDbDataAdapter da;
        private void Process(string strCom)
        {
            myCon = new OleDbConnection(strConnection);
            myCon.Open();
            switch (strCom)
            {
                case "insert":
                    sql = "insert into " + DemoTable + "(Num,Name1) values ('123','王')";
                    break;
                case "select":
                    //sql = "select * from " + DemoTable + " where id=50";
                    sql = "select top 1 * from " + DemoTable;    //from后面要有空格
                    ds = new DataSet();
                    da = new OleDbDataAdapter(sql, myCon);
                    da.Fill(ds,DemoTable);
                    dt = ds.Tables[0];
                    num = dt.Rows[0]["id"].ToString();
                    textBox2.Text = dt.Rows[0]["Num"].ToString();
                    textBox3.Text = dt.Rows[0]["Name1"].ToString();
                    break;
                case "delete":
                    sql = "delete from " + DemoTable + " where id = "+num;
                    break;
                case "update":
                    sql = "update " + DemoTable + " set Num='" + "456" + "' where id=49";
                    break;

            }
            myCommand = new OleDbCommand(sql, myCon);
            myCommand.ExecuteNonQuery();
            myCon.Close();
        }
        int k = 0;
        private void button1_Click(object sender, EventArgs e)
        {
            Process("insert");
            textBox1.Text = k.ToString();
            k++;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process("select");
            Process("delete");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Process("update");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Process("select");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'demoDataSet.table1' table. You can move, or remove it, as needed.

        }

        private void button5_Click(object sender, EventArgs e)
        {
            AccessClass Acc = new AccessClass();
            string info = Acc.AccessProce("insert", strConnection, DemoTable);
            textBox2.Text = info;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            AccessClass Acc = new AccessClass();
            string info = Acc.AccessProce("select", strConnection, DemoTable);
            textBox2.Text = info;
        }
    }

}
