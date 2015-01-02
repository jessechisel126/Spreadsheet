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
 * File Description: This file contains the Cell class, used in the Spreadsheet class.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;

namespace SpreadsheetEngine
{
    // Our abstract base Cell class, which is able to use property changed events.
    public abstract class Cell : INotifyPropertyChanged
    {
        #region Fields

        private readonly int _row = 0;
        private readonly int _col = 0;
        private readonly string _name = "";

        protected string _text = "";
        protected string _value = "";
        protected int _backColor = -1;

        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        #endregion

        #region Properties

        // Row location of this cell.
        public int RowIndex
        {
            get { return _row; }
        }

        // Column location of this cell.
        public int ColumnIndex
        {
            get { return _col; }
        }

        // Name of this cell.
        public string Name
        {
            get { return _name; }
        }

        // Text in this cell.
        public string Text
        {
            get { return _text; }

            set
            {
                // If our text changed, set it and fire property changed event.
                if (value != _text)
                {
                    _text = value;
                    PropertyChanged(this, new PropertyChangedEventArgs("Text"));
                }
            }
        }

        // Value of this cell.
        public string Value
        {
            get { return _value; }
        }

        // Background color of this cell.
        public int BackColor
        {
            get { return _backColor; }

            set
            {
                // If our color changed, set it and fire property changed event.
                if (value != _backColor)
                {
                    _backColor = value;
                    PropertyChanged(this, new PropertyChangedEventArgs("BackColor"));
                }
            }
        }

        // Whether this cell contains all default values.
        public bool HasDefaults
        {
            get
            {
                if (BackColor == -1 && string.IsNullOrEmpty(Text))
                {
                    return true;
                }

                return false;
            }
        }

        #endregion

        #region Constructors

        public Cell()
        {
        }

        // Row, column and name can only be set through this constructor.
        public Cell(int row, int col)
        {
            _row = row;
            _col = col;
            _name += Convert.ToChar('A' + col);
            _name += (row + 1).ToString();
        }

        #endregion

        #region Methods

        /// <summary>
        /// Clears the cell.
        /// </summary>
        public void Clear()
        {
            Text = "";
            BackColor = -1;
        }

        #endregion
    }
}