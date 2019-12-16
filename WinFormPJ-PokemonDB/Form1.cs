using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;

namespace WinFormPJ_PokemonDB
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        MySqlConnection conn;
        DataSet dataSet;
        MySqlDataAdapter adapter, adapter1, adapter2, adapter3, adapter4, adapter5, adapter6;
        int IndexOfSelectedRow;
        string NameOfSelectedTab = "전국도감";
        string[] join_pokeNum = {
            "SELECT 전국번호, 포켓몬이름, 타입1, 타입2 FROM 포켓몬정보 INNER JOIN 전국도감 using(전국번호);",
            "SELECT 관동번호, 전국번호, 포켓몬이름, 타입1, 타입2 FROM 관동도감 INNER JOIN 전국도감 using(전국번호) INNER JOIN 포켓몬정보 using(전국번호);",
            "SELECT 성도번호, 전국번호, 포켓몬이름, 타입1, 타입2 FROM 성도도감 INNER JOIN 전국도감 using(전국번호) INNER JOIN 포켓몬정보 using(전국번호);",
            "SELECT 호연번호, 전국번호, 포켓몬이름, 타입1, 타입2 FROM 호연도감 INNER JOIN 전국도감 using(전국번호) INNER JOIN 포켓몬정보 using(전국번호);",
            "SELECT 신오번호, 전국번호, 포켓몬이름, 타입1, 타입2 FROM 신오도감 INNER JOIN 전국도감 using(전국번호) INNER JOIN 포켓몬정보 using(전국번호);",
            "SELECT 하나번호, 전국번호, 포켓몬이름, 타입1, 타입2 FROM 하나도감 INNER JOIN 전국도감 using(전국번호) INNER JOIN 포켓몬정보 using(전국번호);",
            "SELECT 칼로스번호, 전국번호, 포켓몬이름, 타입1, 타입2 FROM 칼로스도감 INNER JOIN 전국도감 using(전국번호) INNER JOIN 포켓몬정보 using(전국번호);"
        };
        string[] Tables = { "전국", "관동", "성도", "호연", "신오", "하나", "칼로스" };

        private void Form1_Load(object sender, EventArgs e)
        {
            string connStr = "server=poke-gsm.mysql.database.azure.com;port=3306;database=poke;uid=woskaangel@poke-gsm;pwd=k06120393@";
            conn = new MySqlConnection(connStr);
            dataSet = new DataSet();

            adapter = new MySqlDataAdapter(join_pokeNum[0], conn);
            adapter.Fill(dataSet, "전국도감");

            adapter1 = new MySqlDataAdapter(join_pokeNum[1], conn);
            adapter1.Fill(dataSet, "관동도감");

            adapter2 = new MySqlDataAdapter(join_pokeNum[2], conn);
            adapter2.Fill(dataSet, "성도도감");

            adapter3 = new MySqlDataAdapter(join_pokeNum[3], conn);
            adapter3.Fill(dataSet, "호연도감");

            adapter4 = new MySqlDataAdapter(join_pokeNum[4], conn);
            adapter4.Fill(dataSet, "신오도감");

            adapter5 = new MySqlDataAdapter(join_pokeNum[5], conn);
            adapter5.Fill(dataSet, "하나도감");

            adapter6 = new MySqlDataAdapter(join_pokeNum[6], conn);
            adapter6.Fill(dataSet, "칼로스도감");

            PokemonDataView.DataSource = dataSet.Tables["전국도감"];
            Setting("전국도감");
        }

        //포켓몬 탭 인덱스가 바뀔 경우(도감의 종류 변경)
        private void PokemonTabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            NameOfSelectedTab = PokemonTabControl.SelectedTab.Text + "도감";
            switch (NameOfSelectedTab)
            {
                case "전국도감":
                    dataSet.Tables["전국도감"].Clear();
                    adapter.Fill(dataSet, "전국도감");
                    PokemonDataView.DataSource = dataSet.Tables[NameOfSelectedTab];
                    break;
                case "관동도감":
                    dataSet.Tables["관동도감"].Clear();
                    adapter1.Fill(dataSet, "관동도감");
                    PokemonDataView.DataSource = dataSet.Tables[NameOfSelectedTab];
                    break;
                case "성도도감":
                    dataSet.Tables["성도도감"].Clear();
                    adapter2.Fill(dataSet, "성도도감");
                    PokemonDataView.DataSource = dataSet.Tables[NameOfSelectedTab];
                    break;
                case "호연도감":
                    dataSet.Tables["호연도감"].Clear();
                    adapter3.Fill(dataSet, "호연도감");
                    PokemonDataView.DataSource = dataSet.Tables[NameOfSelectedTab];
                    break;
                case "신오도감":
                    dataSet.Tables["신오도감"].Clear();
                    adapter4.Fill(dataSet, "신오도감");
                    PokemonDataView.DataSource = dataSet.Tables[NameOfSelectedTab];
                    break;
                case "하나도감":
                    dataSet.Tables["하나도감"].Clear();
                    adapter5.Fill(dataSet, "하나도감");
                    PokemonDataView.DataSource = dataSet.Tables[NameOfSelectedTab];
                    break;
                case "칼로스도감":
                    dataSet.Tables["칼로스도감"].Clear();
                    adapter6.Fill(dataSet, "칼로스도감");
                    PokemonDataView.DataSource = dataSet.Tables[NameOfSelectedTab];
                    break;
            }
            /*for (int i = 0; i < Tables.Length; i++)
            {
                if (NameOfSelectedTab == Tables[i].ToString())
                    PokemonDataView.DataSource = dataSet.Tables[NameOfSelectedTab];
            }*/
        }

        private void Setting(string currentTab)
        {
            try
            {
                if (currentTab == "전국도감")
                {
                    // 모든 타입의 종류는 같으므로 타입은 타입 1에서 불러와서 타입 2의 검색을 허용
                    // 타입2 같은 경우는 2번째 타입이 없는 포켓몬도 있어 null 값이 있기에 제대로 갱신이 되지 않음
                    string queryStr = "SELECT distinct 타입1 FROM 포켓몬정보 INNER JOIN 전국도감 using(전국번호);";
                    MySqlCommand commands = new MySqlCommand(queryStr, conn);

                    conn.Open();
                    MySqlDataReader reader = commands.ExecuteReader();
                    while (reader.Read())
                    {
                        PokemonType1ComboBox.Items.Add(reader.GetString("타입1"));
                        PokemonType2ComboBox.Items.Add(reader.GetString("타입1"));
                    }
                    reader.Close();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        private void ComboBox_ReSetting(string currentTab)
        {
            switch (currentTab)
            {
                case "전국도감":
                    PokemonType1ComboBox.Items.Clear();
                    PokemonType2ComboBox.Items.Clear();
                    Setting(currentTab);
                    break;
                case "관동도감":
                    PokemonType1ComboBox.Items.Clear();
                    PokemonType2ComboBox.Items.Clear();
                    Setting(currentTab);
                    break;
                case "성도도감":
                    PokemonType1ComboBox.Items.Clear();
                    PokemonType2ComboBox.Items.Clear();
                    Setting(currentTab);
                    break;
                case "호연도감":
                    PokemonType1ComboBox.Items.Clear();
                    PokemonType2ComboBox.Items.Clear();
                    Setting(currentTab);
                    break;
                case "신오도감":
                    PokemonType1ComboBox.Items.Clear();
                    PokemonType2ComboBox.Items.Clear();
                    Setting(currentTab);
                    break;
                case "하나도감":
                    PokemonType1ComboBox.Items.Clear();
                    PokemonType2ComboBox.Items.Clear();
                    Setting(currentTab);
                    break;
                case "칼로스도감":
                    PokemonType1ComboBox.Items.Clear();
                    PokemonType2ComboBox.Items.Clear();
                    Setting(currentTab);
                    break;
            }
        }

        private void PokemonDataView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            IndexOfSelectedRow = e.RowIndex;
            DataGridViewRow row = PokemonDataView.Rows[IndexOfSelectedRow];
            NameOfSelectedTab = PokemonTabControl.SelectedTab.Text + "도감";

            Form2 modalForm;
            if (NameOfSelectedTab == "전국도감")
            {
                modalForm = new Form2(
                    IndexOfSelectedRow,
                    NameOfSelectedTab,
                    row.Cells[0].Value.ToString(),
                    row.Cells[1].Value.ToString(),
                    row.Cells[2].Value.ToString(),
                    row.Cells[3].Value.ToString()
                )
                {
                    Owner = this
                };
            }
            else
            {
                modalForm = new Form2(
                    IndexOfSelectedRow,
                    NameOfSelectedTab,
                    row.Cells[0].Value.ToString(),
                    row.Cells[1].Value.ToString(),
                    row.Cells[2].Value.ToString(),
                    row.Cells[3].Value.ToString(),
                    row.Cells[4].Value.ToString()
                )
                {
                    Owner = this
                };
            }
            // 새로운 폼에 선택된 row의 정보를 담아서 생성

            modalForm.ShowDialog();
            modalForm.Dispose();
        }

        private void dictWorldSearchBtn_Click(object sender, EventArgs e)
        {
            string queryString;

            string[] options = new string[4];
            string options_PokemonNumber;
            if (PokemonNumBox_min.Text != "" && PokemonNumBox_max.Text != "")
            {
                options_PokemonNumber = "전국번호>=@PokemonNum_min and 전국번호<=@PokemonNum_max";
            }
            else if (PokemonNumBox_min.Text != "" || PokemonNumBox_max.Text != "")
            {
                if (PokemonNumBox_min.Text != "")
                    options_PokemonNumber = "전국번호>=@PokemonNum_min";
                else options_PokemonNumber = "전국번호<=@PokemonNum_max";
            }
            else
            {
                options_PokemonNumber = null;
            }
            options[0] = options_PokemonNumber;
            options[1] = (PokemonNameBox.Text != "") ? "포켓몬이름=@포켓몬이름" : null;
            options[2] = (PokemonType1ComboBox.Text != "") ? "타입1=@타입1" : null;
            options[3] = (PokemonType2ComboBox.Text != "") ? "타입2=@타입2" : null;

            if (options[0] != null || options[1] != null || options[2] != null || options[3] != null)
            {
                queryString = $"SELECT * FROM 포켓몬정보 INNER JOIN 전국도감 using(전국번호) WHERE ";
                bool firstOption = true;
                for (int i = 0; i < options.Length; i++)
                {
                    if (options[i] != null)
                    {
                        if (firstOption)
                        {
                            queryString += options[i];
                            firstOption = false;
                        }
                        else queryString += " and " + options[i];
                    }
                }
            }
            else queryString = "SELECT * FROM 포켓몬정보 INNER JOIN 전국도감 using(전국번호)";

            adapter.SelectCommand = new MySqlCommand(queryString, conn);
            adapter.SelectCommand.Parameters.AddWithValue("@PokemonNum_min", PokemonNumBox_min.Text);
            adapter.SelectCommand.Parameters.AddWithValue("@PokemonNum_max", PokemonNumBox_max.Text);
            adapter.SelectCommand.Parameters.AddWithValue("@포켓몬이름", PokemonNameBox.Text);
            adapter.SelectCommand.Parameters.AddWithValue("@타입1", PokemonType1ComboBox.Text);
            adapter.SelectCommand.Parameters.AddWithValue("@타입2", PokemonType2ComboBox.Text);

            try
            {
                conn.Open();
                dataSet.Tables["전국도감"].Clear();
                if (adapter.Fill(dataSet, "전국도감") > 0)
                    PokemonDataView.DataSource = dataSet.Tables["전국도감"];
                else MessageBox.Show("찾는 데이터가 엄서용!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        private void dictWorldOthersBtn_Click(object sender, EventArgs e)
        {
            Form2 Form = new Form2("전국도감");
            Form.Owner = this;
            Form.ShowDialog();
            Form.Dispose();
        }

        private void dictWorldTBClearBtn_Click(object sender, EventArgs e)
        {
            PokemonNameBox.Clear();
            PokemonNumBox_min.Clear();
            PokemonNumBox_max.Clear();
            PokemonType1ComboBox.Text = "";
            PokemonType2ComboBox.Text = "";
        }

        public void InsertRow(string[] PokemonData, string NameOfSelectedTab)
        {
            string queryString;
            queryString = "INSERT INTO 포켓몬정보 VALUES(@전국번호, @포켓몬이름, @타입1, @타입2)";
            adapter.InsertCommand = new MySqlCommand(queryString, conn);
            adapter.InsertCommand.Parameters.Add("@전국번호", MySqlDbType.Int32);
            adapter.InsertCommand.Parameters.Add("@포켓몬이름", MySqlDbType.VarChar);
            adapter.InsertCommand.Parameters.Add("@타입1", MySqlDbType.VarChar);
            adapter.InsertCommand.Parameters.Add("@타입2", MySqlDbType.VarChar);

            #region Parameters를 이용한 데이터 삽입 처리
            adapter.InsertCommand.Parameters["@전국번호"].Value = Convert.ToInt32(PokemonData[0]);
            adapter.InsertCommand.Parameters["@포켓몬이름"].Value = PokemonData[1];
            adapter.InsertCommand.Parameters["@타입1"].Value = PokemonData[2];
            adapter.InsertCommand.Parameters["@타입2"].Value = PokemonData[3];
            #endregion

            #region 기타 지방 테이블 삽입 처리
            // 데이터가 너무 많아 트래픽 처리가 힘들 것 같아 구현하지 않음
            #endregion

            try
            {
                conn.Open();
                adapter.InsertCommand.ExecuteNonQuery();
                dataSet.Tables[NameOfSelectedTab].Clear();  // 선택한 테이블의 이전 데이터 지우기
                adapter.Fill(dataSet, NameOfSelectedTab);
                PokemonDataView.DataSource = dataSet.Tables[NameOfSelectedTab];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
            ComboBox_ReSetting(NameOfSelectedTab);
        }

        public void UpdateRow(string[] PokemonData, string NameOfSelectedTab)
        {
            string queryString;
            queryString = "UPDATE 포켓몬정보 SET 포켓몬이름=@포켓몬이름, 타입1=@타입1, 타입2=@타입2 WHERE 전국번호=@전국번호";
            adapter.UpdateCommand = new MySqlCommand(queryString, conn);
            adapter.UpdateCommand.Parameters.AddWithValue("@전국번호", PokemonData[0]);
            adapter.UpdateCommand.Parameters.AddWithValue("@포켓몬이름", PokemonData[1]);
            adapter.UpdateCommand.Parameters.AddWithValue("@타입1", PokemonData[2]);
            adapter.UpdateCommand.Parameters.AddWithValue("@타입2", PokemonData[3]);

            try
            {
                conn.Open();
                adapter.UpdateCommand.ExecuteNonQuery();
                dataSet.Tables[NameOfSelectedTab].Clear();  // 선택한 테이블의 이전 데이터 지우기
                adapter.Fill(dataSet, NameOfSelectedTab);
                PokemonDataView.DataSource = dataSet.Tables[NameOfSelectedTab];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
            ComboBox_ReSetting(NameOfSelectedTab);
        }

        public void DeleteRow(string PrimaryKey, string NameOfSelectedTab)
        {
            string queryString;
            queryString = "DELETE FROM 포켓몬정보 WHERE 전국번호=@전국번호";
            adapter.DeleteCommand = new MySqlCommand(queryString, conn);
            adapter.DeleteCommand.Parameters.AddWithValue("@전국번호", PrimaryKey);

            try
            {
                conn.Open();
                adapter.DeleteCommand.ExecuteNonQuery();
                dataSet.Tables[NameOfSelectedTab].Clear();  // 선택한 테이블의 이전 데이터 지우기
                adapter.Fill(dataSet, NameOfSelectedTab);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
            ComboBox_ReSetting(NameOfSelectedTab);
        }

        // 파일로 저장하기(txt, excel)
        private void btnSave_Click(object sender, EventArgs e)
        {
            // dataGridView에 데이터가 존재하는지 체크
            if (PokemonDataView.RowCount == 0)
            {
                MessageBox.Show("저장할 데이가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 라디오 버튼 상태에 따라 txt 또는 excel 파일로 저장 
            if (rbText.Checked)
            {
                saveFileDialog1.Filter = "텍스트 파일(*.txt)|*.txt";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    SaveTextFile(saveFileDialog1.FileName);
                    MessageBox.Show("메모장에 저장하였습니다!");
                }
            }
            else if (rbExcel.Checked)
            {
                saveFileDialog1.Filter = "엑셀 파일(*.xlsx)|*.xlsx";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    SaveExcelFile(saveFileDialog1.FileName);
                    MessageBox.Show("엑셀에 저장하였습니다!");
                }
            }
        }

        // 엑셀 파일 저장
        private void SaveExcelFile(string filePath)
        {
            // 1. 엑셀 사용에 필요한 객체들 생성
            Excel.Application eApp;     // 엑셀 프로그램
            Excel.Workbook eWorkbook;   // 엑셀 시트를 여러개 포함하는 단위
            Excel.Worksheet eWorkSheet; // 엑셀 워크시트

            eApp = new Excel.Application();
            eWorkbook = eApp.Workbooks.Add();   // eApp에 워크북 추가
            eWorkSheet = eWorkbook.Sheets[1];   // 엑셀 워크시트는 index가 1부터 시작한다.

            // 2. 엑셀에 저장할 데이터를 2차원 배열 형태로 준비
            string[,] dataArr;
            int colCount = dataSet.Tables[NameOfSelectedTab].Columns.Count + 1;
            int rowCount = dataSet.Tables[NameOfSelectedTab].Rows.Count + 1;
            dataArr = new string[rowCount, colCount];

            // 2-1 Column 이름 저장
            for (int i = 0; i < dataSet.Tables[NameOfSelectedTab].Columns.Count; i++)
            {
                dataArr[0, i] = dataSet.Tables[NameOfSelectedTab].Columns[i].ColumnName;  // 첫 행에 컬럼이름 저장
            }

            // 2-2 행 데이터 저장
            for (int i = 0; i < dataSet.Tables[NameOfSelectedTab].Rows.Count; i++)
            {
                for (int j = 0; j < dataSet.Tables[NameOfSelectedTab].Columns.Count; j++)
                {
                    dataArr[i + 1, j] = dataSet.Tables[NameOfSelectedTab].Rows[i].ItemArray[j].ToString();
                }
            }

            // 3. 준비된 데이터를 엑셀파일에 저장
            string endCell = $"E{rowCount}";        // 데이터가 저장이 끝나는 셀의 주소
            eWorkSheet.get_Range("A1:" + endCell).Value = dataArr;  // 데이터가 저장될 셀의 범위에 2차원 배열 저장

            eWorkbook.SaveAs(filePath, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Excel.XlSaveAsAccessMode.xlShared, false, false, Type.Missing, Type.Missing,
                Type.Missing);
            eWorkbook.Close(false, Type.Missing, Type.Missing);
            eApp.Quit();
        }

        // 텍스트 파일 저장
        void SaveTextFile(string filePath)
        {
            // SaveFileDialog에서 지정한 파일경로에 Strem 생성 -> 저장 
            using (StreamWriter sw = new StreamWriter(filePath))
            {
                // Column 이름들 저장
                foreach (DataColumn col in dataSet.Tables[NameOfSelectedTab].Columns)
                {
                    sw.Write($"{col.ColumnName}\t");
                }
                sw.WriteLine();

                // DataSet의 데이터 행 저장
                foreach (DataRow row in dataSet.Tables[NameOfSelectedTab].Rows)
                {
                    string rowString = "";
                    foreach (var item in row.ItemArray)
                    {
                        rowString += $"{item.ToString()}\t";
                    }
                    sw.WriteLine(rowString);
                }
            }
        }
    }
}
