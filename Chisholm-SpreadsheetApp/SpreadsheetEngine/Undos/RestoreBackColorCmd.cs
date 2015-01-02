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
 * File Description: This file contains the RestoreBackColorCmd class.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine.Undos
{
    // A restore back color command.
    public class RestoreBackColorCmd : IUndoRedoCmd
    {
        private int _cellColor;
        private string _cellName;

        public RestoreBackColorCmd(int cellColor, string cellName)
        {
            _cellColor = cellColor;
            _cellName = cellName;
        }

        public IUndoRedoCmd Execute(Spreadsheet sheet)
        {
            Cell cell = sheet.GetCell(_cellName);
            int oldColor = cell.BackColor;
            cell.BackColor = _cellColor;
            return new RestoreBackColorCmd(oldColor, _cellName);
        }
    }
}
