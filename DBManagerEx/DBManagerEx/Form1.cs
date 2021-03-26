using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DBManagerEx
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlCom = new SqlCommand();
        private void mnuNew_Click(object sender, EventArgs e)   // grid 초기화 -> Table 명칭 초기화 -> DB 초기화
        {
            dataGrid.Rows.Clear();                              // 초기화 순서  Row -> Column
            dataGrid.Columns.Clear();

            sbDBName.Text = "DB File Name";
            sbTables.Text = "Table List";
            sbTables.DropDownItems.Clear();
            sbMessage.Text = "Initialized.";
            sqlConn.Close();
        }

        private void mnuMigration_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            StreamReader sr = new StreamReader(openFileDialog1.FileName);
            string buf = sr.ReadLine();     // 첫번째 Line에는 각 Column의 HeaderText
            string[] sArr = buf.Split(','); // ','로 구분
            for (int i = 0; i < sArr.Length; i++)
            {
                dataGrid.Columns.Add(sArr[i], sArr[i]);
            }
            while(true)     //무한루프를 통해서 데이터 불러옴
            {
                buf = sr.ReadLine();
                if (buf == null) break;
                sArr = buf.Split(',');      // string array
                dataGrid.Rows.Add(sArr);    // Rows.Add의 4번째 오버로드 : obj argument에는 str 사용 가능 (반대로는 x)
            }
            sr.Close();
        }

        private void mnuFileCSV_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;
            StreamWriter sw = new StreamWriter(saveFileDialog1.FileName,false,Encoding.Default);   //1.경로 2. append 모드 3.Encoding
            string buf = "";
            for(int i=0; i<dataGrid.ColumnCount; i++)
            {
                buf += dataGrid.Columns[i].HeaderText;
                if (i < dataGrid.ColumnCount - 1) buf += ",";
            }
            sw.Write(buf+"\r\n");

            for(int k=0;k<dataGrid.RowCount; k++)
            {
                buf = "";
                for (int i = 0; i < dataGrid.ColumnCount; i++)
                {
                    buf += dataGrid.Rows[k].Cells[i].Value;
                    if (i < dataGrid.ColumnCount - 1) buf += ",";
                }
                sw.Write(buf + "\r\n");
            }
            sw.Close();
        }
        string sC   on = @"Data Source = (LocalDB)\MSSQLLocalDB;AttachDbFilename=;Integrated Security = True; Connect Timeout = 30";
        private void mnuDBOpen_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            try 
            {
                string[] sArr = sCon.Split(';');
                sCon = $"{sArr[0]};{sArr[1]}{openFileDialog1.FileName};{sArr[2]};{sArr[3]}";
                sqlConn.ConnectionString = sCon;
                sqlConn.Open();
                sqlCom.Connection = sqlConn;
                sbDBName.Text = openFileDialog1.SafeFileName;
                sbDBName.BackColor = Color.Gold;

                DataTable dt = sqlConn.GetSchema("Tables");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sbTables.DropDownItems.Add(dt.Rows[i].ItemArray[2].ToString());
                }
                sbMessage.Text = "Success";
                sbMessage.BackColor = Color.Gray;
            }
            catch(Exception e1)     // 현재 이벤트 Argument가 e를 가지고 있기 때문.
            {
                MessageBox.Show(e1.Message, "Error!");
                sbMessage.Text = "Error";
                sbMessage.BackColor = Color.IndianRed;
            }
        }

        private void sbTables_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            sbTables.Text = e.ClickedItem.Text;
            Runsql($"select * from {sbTables.Text}");
        }

        private void ClearGrid()        // 현재의 Grid를 Clear시켜주는 함수
        {
            dataGrid.Rows.Clear();
            dataGrid.Columns.Clear();
        }

        public string GetToken(int i, string src, char del)     //Sql문에서 가장 첫번째 단어를 얻을 수 있다.
        {
            string[] sArr = src.Split(del);
            return sArr[i];
        }

        private void Runsql(string Sql)                         // Runsql : string으로 sql문을 받는 함수.
        {   // select id, fCode from Facility
            try 
            {
                string ss = GetToken(0, Sql.Trim().ToLower(), ' ');  // ss: start string
                sqlCom.CommandText = Sql;
                if (ss == "select")
                {
                    ClearGrid();
                    SqlDataReader sdr = sqlCom.ExecuteReader();
                    for (int i = 0; i < sdr.FieldCount; i++)
                    {
                        dataGrid.Columns.Add(sdr.GetName(i), sdr.GetName(i));
                    }

                    for (int k = 0; sdr.Read(); k++)
                    {
                        object[] oArr = new object[sdr.FieldCount];
                        sdr.GetValues(oArr);
                        dataGrid.Rows.Add(oArr);
                    }
                }
                else
                {
                    sqlCom.ExecuteNonQuery();
                }
            }

            catch(Exception e1)
            {
                MessageBox.Show(e1.Message);
            }
            
        }

        private void dataGrid_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            dataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].ToolTipText = ".";
        }

        private void mnuUpdate_Click(object sender, EventArgs e)
        {
            for(int i=0;i<dataGrid.RowCount;i++)
            {
                for(int k=0;k<dataGrid.ColumnCount;k++)
                {
                    if(dataGrid.Rows[i].Cells[k].ToolTipText == ".")
                    {//"update {TableName} set{currentCellHeader} = {currentCellValue} where {idHeader} = {idValue}"
                      
                        string tn = sbTables.Text;                      //TableName
                        string ht = dataGrid.Columns[k].HeaderText;     //currentCellHeader
                        object cv = dataGrid.Rows[i].Cells[k].Value;    //currentCellValue
                        string it = dataGrid.Columns[0].HeaderText;     //idHeader
                        object id = dataGrid.Rows[i].Cells[0].Value;    //idValue

                        string sql = $"update {tn} set{ht}=N'{cv}' where {it}={id}";
                        Runsql(sql);
                    }
                }
            }
        }
    }
}
