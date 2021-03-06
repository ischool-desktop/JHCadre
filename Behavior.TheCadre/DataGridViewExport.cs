﻿using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Cells;
using System.Windows.Forms;
using FISCA.Presentation.Controls;

namespace Behavior.TheCadre
{
    internal class DataGridViewExport
    {
        Workbook _workbook;
        Worksheet _worksheet;
        List<int> _colIndexes;

        public DataGridViewExport(DataGridView dgv)
        {
            _workbook = new Workbook();
            _workbook.Worksheets.Clear();
            _worksheet = _workbook.Worksheets[_workbook.Worksheets.Add()];
            _worksheet.Name = "Sheet1";
            _colIndexes = new List<int>();

            int sheetRowIndex = 0;
            int sheetColIndex = 0;
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (col.Visible == false) continue;
                _colIndexes.Add(col.Index);
                _worksheet.Cells[sheetRowIndex, sheetColIndex++].PutValue(col.HeaderText);
            }

            foreach (DataGridViewRow row in dgv.Rows)
            {
                sheetRowIndex++;
                sheetColIndex = 0;
                foreach (int colIndex in _colIndexes)
                {
                     if (dgv.Columns[row.Cells[colIndex].ColumnIndex].HeaderText == "參與比序")
                     {
                          if ("" + row.Cells[colIndex].Value == "True")
                          {
                               _worksheet.Cells[sheetRowIndex, sheetColIndex++].PutValue("是");
                          }
                     }
                     else
                     {
                          _worksheet.Cells[sheetRowIndex, sheetColIndex++].PutValue("" + row.Cells[colIndex].Value);
                     }
                }
            }

            _worksheet.AutoFitColumns();
        }

        public void Save(string path)
        {
            try
            {
                _workbook.Save(path, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                MsgBox.Show("匯出失敗：" + ex.Message);
            }
        }
    }
}
