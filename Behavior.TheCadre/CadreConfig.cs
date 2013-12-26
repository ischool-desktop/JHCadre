using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data.Configuration;
using System.Xml;
using FISCA.DSAUtil;

namespace Behavior.TheCadre
{
    public partial class CadreConfig : BaseForm
    {
         private string ConfigName = "BehaviorCadreConfig";
         private string SchoolCadre = "學校幹部";
         private string ClassCadre = "班級幹部";
         private string AssnCadre = "社團幹部";

        public CadreConfig()
        {
            InitializeComponent();
        }

        private void SchoolCadre_Load(object sender, EventArgs e)
        {
            SetupForm();
        }

        /// <summary>
        /// 初始畫介面
        /// </summary>
        private void SetupForm()
        {
            ConfigData cd = K12.Data.School.Configuration[ConfigName];

            string SchoolConfig = cd[SchoolCadre];

            if (!string.IsNullOrEmpty(SchoolConfig))
            {
                XmlElement xml = DSXmlHelper.LoadXml(SchoolConfig);
                foreach (XmlElement XmlEach in xml.SelectNodes("Item"))
                {
                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(dataGridViewX1);
                    row.Cells[0].Value = XmlEach.GetAttribute("CadreName");
                    dataGridViewX1.Rows.Add(row);
                }
            }

            string ClassConfig = cd[ClassCadre];

            if (!string.IsNullOrEmpty(ClassConfig))
            {
                XmlElement xml = DSXmlHelper.LoadXml(ClassConfig);
                foreach (XmlElement XmlEach in xml.SelectNodes("Item"))
                {
                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(dataGridViewX2);
                    row.Cells[0].Value = XmlEach.GetAttribute("CadreName");
                    dataGridViewX2.Rows.Add(row);
                }
            }

            string AssnConfig = cd[AssnCadre];

            if (!string.IsNullOrEmpty(AssnConfig))
            {
                XmlElement xml = DSXmlHelper.LoadXml(AssnConfig);
                foreach (XmlElement XmlEach in xml.SelectNodes("Item"))
                {
                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(dataGridViewX3);
                    row.Cells[0].Value = XmlEach.GetAttribute("CadreName");
                    dataGridViewX3.Rows.Add(row);
                }
            }
        }

        /// <summary>
        /// 儲存設定檔
        /// </summary>
        private void btnSave_Click(object sender, EventArgs e)
        {
            ConfigData cd = K12.Data.School.Configuration[ConfigName];

            DSXmlHelper helper1 = new DSXmlHelper();
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.IsNewRow)
                    continue;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    string cellValue = "" + cell.Value;
                    if (!string.IsNullOrEmpty(cellValue)) //如果是空的就跳過
                    {
                        helper1.AddElement("Item");
                        helper1.SetAttribute("Item", "CadreName", cellValue);
                    }
                }
            }
            cd[SchoolCadre] = helper1.BaseElement.OuterXml;

            DSXmlHelper helper2 = new DSXmlHelper();
            foreach (DataGridViewRow row in dataGridViewX2.Rows)
            {
                if (row.IsNewRow)
                    continue;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    string cellValue = "" + cell.Value;
                    if (!string.IsNullOrEmpty(cellValue)) //如果是空的就跳過
                    {
                        helper2.AddElement("Item");
                        helper2.SetAttribute("Item", "CadreName", cellValue);
                    }
                }
            }
            cd[ClassCadre] = helper2.BaseElement.OuterXml;

            DSXmlHelper helper3 = new DSXmlHelper();
            foreach (DataGridViewRow row in dataGridViewX3.Rows)
            {
                if (row.IsNewRow)
                    continue;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    string cellValue = "" + cell.Value;
                    if (!string.IsNullOrEmpty(cellValue)) //如果是空的就跳過
                    {
                        helper3.AddElement("Item");
                        helper3.SetAttribute("Item", "CadreName", cellValue);
                    }
                }
            }
            cd[AssnCadre] = helper3.BaseElement.OuterXml;

            cd.Save();

            MsgBox.Show("儲存成功");
            this.Close();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
