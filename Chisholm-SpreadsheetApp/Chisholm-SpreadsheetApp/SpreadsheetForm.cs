/*
 * Programmer: Jesse Chisholm 11278684
 * Program: Excell Program Part 4 (HW10)
 * Date Started: 9/26/14
 * Date Completed: 11/21/14
 * Hours Worked: 50+
 * Collaboration: Light collab with Ben Tatham, Tim Vierow, and Zach Allen, and often referenced MSDN.
 * 
 * Program Description: This program is a simple unfinished version of Excell using Winforms and C#.
 * 
 * File Description: This file contains the UI implementation of a spreadsheet using the Workbook 
 *                   class with UI elements.
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ExpressionEngine;
using SpreadsheetEngine;
using SpreadsheetEngine.Undos;

namespace Chisholm_SpreadsheetApp
{
    public partial class SpreadsheetForm : Form
    {
        public Workbook book = new Workbook(50, 26);

        public SpreadsheetForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Subscribe to events.
            dataGridView1.CellBeginEdit += dataGridView1_CellBeginEdit;
            dataGridView1.CellEndEdit += dataGridView1_CellEndEdit;
            book.WorkbookSheetChanged += OnWorkbookSheetChanged;

            // Clear our grid.
            dataGridView1.Columns.Clear();

            // Create each column and name them with capital letters.
            for (char c = 'A'; c <= 'Z'; c++)
            {
                string name = Convert.ToString(c);
                dataGridView1.Columns.Add(name, name);
            }

            // Add 50 rows to grid.
            dataGridView1.Rows.Add(50);

            // For each row in the grid, set the row header value to the string version of the row number.
            int rowNumber = 1;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.HeaderCell.Value = Convert.ToString(rowNumber++);
            }

            // Resize row widths to fit row headers.
            dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            // Get the row and column indices of the dataGridView cell that started being edited.
            int row = e.RowIndex;
            int col = e.ColumnIndex;

            // Get the corresponding cell from our sheet.
            Cell currCell = book.ActiveSheet.GetCell(row, col);

            // Set the dataGridView cell's value to the sheet cell's Text property.
            dataGridView1.Rows[row].Cells[col].Value = currCell.Text;
        }

        void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // Get the row and column indices of the dataGridView cell that is done being edited.
            int row = e.RowIndex;
            int col = e.ColumnIndex;

            // Get the corresponding cell from our sheet.
            Cell ssCell = book.ActiveSheet.GetCell(row, col);

            // New text from the datagrid.
            string newText;

            // Now, the text from the dataGridView cell can be written to our cell's Text property.
            // If the dataGridView cell's value is null, just write the empty string.
            try
            {
                newText = dataGridView1.Rows[row].Cells[col].Value.ToString();
            }
            catch (NullReferenceException)
            {
                newText = "";
            }

            // Our undos for each action.
            IUndoRedoCmd[] undos = new IUndoRedoCmd[1];

            // The only undo action will be a restore text command with the new text and the cell name.
            undos[0] = new RestoreTextCmd(ssCell.Text, ssCell.Name);

            // Set the text in the spreadsheet cell.
            ssCell.Text = newText;

            // Add the undos to our undo/redo system.
            book.UndoRedo.AddUndos(new UndoRedoCollection(undos, "cell text change"));

            // Now, we can write the new sheet cell value to the dataGridView cell.
            dataGridView1.Rows[row].Cells[col].Value = ssCell.Value;

            UpdateToolStripMenu();
        }

        private void chooseBackgroundColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Our undos for each action.
            List<IUndoRedoCmd> undos = new List<IUndoRedoCmd>();

            // If they've chosen a color...
            ColorDialog colorDialog = new ColorDialog();
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                //...Get the chosen color as an int.
                int chosenColor = colorDialog.Color.ToArgb();

                // For each cell selected...
                foreach (DataGridViewCell dgCell in dataGridView1.SelectedCells)
                {
                    //...Get the corresponding spreadsheet cell.
                    Cell ssCell = book.ActiveSheet.GetCell(dgCell.RowIndex, dgCell.ColumnIndex);

                    // Add this color change to the undos.
                    undos.Add(new RestoreBackColorCmd(ssCell.BackColor, ssCell.Name));

                    // Set the color in the spreadsheet cell.
                    ssCell.BackColor = chosenColor;
                }

                // Add the undos to our undo/redo system.
                book.UndoRedo.AddUndos(new UndoRedoCollection(undos, "cell background color change"));

                UpdateToolStripMenu();
            }
        }
        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            book.UndoRedo.Undo(book.ActiveSheet);
            UpdateToolStripMenu();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            book.UndoRedo.Redo(book.ActiveSheet);
            UpdateToolStripMenu();
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XML files (*.xml)|*.xml";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Stream stream = new FileStream(saveFileDialog.FileName, FileMode.Create, FileAccess.Write);
                book.Save(stream);
                stream.Dispose();
            }
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XML files (*.xml)|*.xml";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                Stream stream = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read);
                book.Load(stream);
                stream.Dispose();
            }

            UpdateToolStripMenu();
        }

        /// <summary>
        /// Updates the undo/redo tool strip items based on the status of the workbook's UndoRedoSystem.
        /// </summary>
        private void UpdateToolStripMenu()
        {
            ToolStripMenuItem group = menuStrip1.Items[1] as ToolStripMenuItem;

            foreach (ToolStripItem item in group.DropDownItems)
            {
                if (item.Text.Substring(0, 4) == "Undo")
                {
                    item.Enabled = book.UndoRedo.CanUndo;
                    item.Text = "Undo " + book.UndoRedo.UndoDescription;
                }
                else if (item.Text.Substring(0, 4) == "Redo")
                {
                    item.Enabled = book.UndoRedo.CanRedo;
                    item.Text = "Redo " + book.UndoRedo.RedoDescription;
                }
            }
        }

        /// <summary>
        /// Handles any Cell's PropertyChanged event in the spreadsheet.
        /// </summary>
        /// <param name="sender">The Cell object which originally fired the PropertyChanged event.</param>
        /// <param name="e">The event arguments from the Spreadsheet.</param>
        private void OnWorkbookSheetChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Value")
            {
                Cell currCell = sender as Cell;

                if (currCell != null)
                {
                    int row = currCell.RowIndex;
                    int col = currCell.ColumnIndex;
                    string value = currCell.Value;

                    dataGridView1.Rows[row].Cells[col].Value = value;
                }
            }
            else if (e.PropertyName == "BackColor")
            {
                Cell currCell = sender as Cell;

                if (currCell != null)
                {
                    int row = currCell.RowIndex;
                    int col = currCell.ColumnIndex;
                    int color = currCell.BackColor;

                    dataGridView1.Rows[row].Cells[col].Style.BackColor = Color.FromArgb(color);
                }
            }
        }
    }
}