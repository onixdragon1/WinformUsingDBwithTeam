using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinFormPJ_PokemonDB
{
    public partial class Form2 : Form
    {
        #region 전국도감 변수
        private string dictWorldPokemonNum, dictWorldPokemonName;
        private string dictWorldPokemonType1, dictWorldPokemonType2;
        #endregion

        #region 지방도감 변수
        private string dictRegionPokemonName, dictRegionPokemonNum;
        private string dictRegionPokemonType1, dictRegionPokemonType2;
        private string dictRegionNum;
        #endregion

        private string NameOfSelectedTab;
        private int IndexOfSelectedRow;

        public Form2()
        {
            InitializeComponent();
        }

        public Form2(string NameOfSelectedTab)
        {
            InitializeComponent();
            this.NameOfSelectedTab = NameOfSelectedTab;
        }

        public Form2(int IndexOfSelectedRow, string NameOfSelectedTab, string value1, string value2, string value3, string value4)
        {
            InitializeComponent();
            this.NameOfSelectedTab = NameOfSelectedTab;
            this.IndexOfSelectedRow = IndexOfSelectedRow;
            this.dictWorldPokemonNum = value1;
            this.dictWorldPokemonName = value2;
            this.dictWorldPokemonType1 = value3;
            this.dictWorldPokemonType2 = value4;
        }

        public Form2(int IndexOfSelectedRow, string NameOfSelectedTab, string value1, string value2, string value3, string value4, string value5)
        {
            InitializeComponent();
            this.NameOfSelectedTab = NameOfSelectedTab;
            this.IndexOfSelectedRow = IndexOfSelectedRow;
            this.dictWorldPokemonNum = value1;
            this.dictWorldPokemonName = value2;
            this.dictWorldPokemonType1 = value3;
            this.dictWorldPokemonType2 = value4;
            this.dictRegionNum = value5;
        }

        Form1 mainForm;
        private void Form2_Load(object sender, EventArgs e)
        {
            OtherQueryLabel.Text += NameOfSelectedTab + " Table";
            value1Label.Text = "전국번호";
            value1TextBox.Text = dictWorldPokemonNum;
            value2Label.Text = "포켓몬이름";
            value2TextBox.Text = dictWorldPokemonName;
            value3Label.Text = "타입1";
            value3TextBox.Text = dictWorldPokemonType1;
            value4Label.Text = "타입2";
            value4TextBox.Text = dictWorldPokemonType2;
            value5Label.Hide();
            value5TextBox.Hide();

            if (Owner != null)
            {
                mainForm = Owner as Form1;
            }
        }

        public string[] IfTextBoxIsBlank(string[] data, string currentTab)
        {
            if (data[0] == "") data[0] = dictWorldPokemonNum;
            if (data[1] == "") data[1] = dictWorldPokemonName;
            if (data[2] == "") data[2] = dictWorldPokemonType1;
            if (data[3] == "") data[3] = dictWorldPokemonType2;
            return data;
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            string[] rowDatas = {
                    value1TextBox.Text,
                    value2TextBox.Text,
                    value3TextBox.Text,
                    value4TextBox.Text
                };
            mainForm.InsertRow(rowDatas, NameOfSelectedTab);
            this.Close();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            string[] rowDatas = {
                    value1TextBox.Text,
                    value2TextBox.Text,
                    value3TextBox.Text,
                    value4TextBox.Text
                };
            mainForm.UpdateRow(IfTextBoxIsBlank(rowDatas, NameOfSelectedTab), NameOfSelectedTab);
            this.Close();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            mainForm.DeleteRow(dictWorldPokemonNum, NameOfSelectedTab);
            this.Close();
        }

        private void btnTextBoxClear_Click(object sender, EventArgs e)
        {
            value1TextBox.Clear();
            value2TextBox.Clear();
            value3TextBox.Clear();
            value4TextBox.Clear();
        }
    }
}
