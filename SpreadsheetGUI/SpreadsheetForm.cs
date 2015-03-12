using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using SS;
using Tweetinvi;

namespace SpreadsheetGUI
{
    public partial class SpreadsheetForm : Form
    {
        Spreadsheet m_spreadsheet;
        //"A1" name of selection
        string m_selectedCell;
        //path to use when saving
        string m_saveFilePath;
        //used so that setCellContents isn't set every time you select a cell
        bool m_contentChanged;

        const string kFileFilter = "Spreadsheet files (*.sprd)|*.sprd|All files (*.*)|*.*";

        //twitter integration variables
        bool m_twitterEnabled;
        string m_twitterAlias;

        public SpreadsheetForm(string path)
        {
            m_spreadsheet = new Spreadsheet(path, s => Regex.IsMatch(s, @"^[A-Z][1-9][0-9]?$"), s => s.ToUpper(), "ps6");
            m_selectedCell = getCellName(0, 0);
            m_saveFilePath = path;

            InitializeComponent();
            this.Text = getFileName(path) + "- Spreadsheet";
            spreadsheetPanel1.SelectionChanged += displaySelection;
            this.FormClosing += saveOnExitPrompt;
            spreadsheetPanel1.SetSelection(0, 0);
            foreach (string cell in m_spreadsheet.GetNamesOfAllNonemptyCells())
            {
                setGUICellValue(cell);
            }

            cellContentsTextBox.Text = getCellContents(m_selectedCell);

            cellNameTextBox.Text = m_selectedCell;
            cellContentsTextBox.Text = getCellContents(m_selectedCell);
            cellValueTextBox.Text = m_spreadsheet.GetCellValue(m_selectedCell).ToString();
        }

        public SpreadsheetForm()
        {
            m_spreadsheet = new Spreadsheet(s => Regex.IsMatch(s, @"^[A-Z][1-9][0-9]?$"), s => s.ToUpper(), "ps6");
            m_selectedCell = getCellName(0, 0);
            m_saveFilePath = "";

            InitializeComponent();
            this.Text = "untitled - Spreadsheet";
            spreadsheetPanel1.SelectionChanged += displaySelection;
            this.FormClosing += saveOnExitPrompt;
            spreadsheetPanel1.SetSelection(0, 0);

            cellNameTextBox.Text = m_selectedCell;
            cellContentsTextBox.Text = getCellContents(m_selectedCell);
            cellValueTextBox.Text = m_spreadsheet.GetCellValue(m_selectedCell).ToString();
        }

        private void setTwitterCredentials()
        {
             //initialize Tweetinvi credentials
            //TwitterCredentials.ApplicationCredentials = TwitterCredentials.CreateCredentials(
            TwitterCredentials.SetCredentials(
                "2891661433-2SaAWATqnHBjgY0Em4ilWCE62Vmg6DxYwhgwGcq",
                "5w7qKeW1dmYQrcKjMy0iHKuyAqPnqxrZOMTCOeYpaYQgX",
                "7VUtsTP60Xswl6ao68JYb9zkN",
                "RrXHlCkfjhvae07P8jkoBGSoqhDr1lZwBhOr5UX7UnVaj1xNTP");
        }

        private string getFileName(string path)
        {
            if (path.Length == 0)
                return "untitled";
            int nameIndex = path.LastIndexOf('\\') + 1;
            return path.Substring(nameIndex, path.LastIndexOf('.') - nameIndex);
        }

        private void saveOnExitPrompt(Object sender, FormClosingEventArgs e)
        {
            if (m_spreadsheet.Changed)
            {
                DialogResult result = MessageBox.Show("Save changes to " + getFileName(m_saveFilePath) + "?",
                                "Save changes?", MessageBoxButtons.YesNoCancel);
                switch (result)
                {
                    case DialogResult.Yes:
                        saveToolStripMenuItem_Click(null, null);
                        break;
                    case DialogResult.No:
                        break;
                    case DialogResult.Cancel:
                        e.Cancel = true;
                        break;
                }
            }
        }

        private string getCellName(int col, int row)
        {
            return (char)(col + 65) + "" + (row + 1);
        }

        private void getCellPosition(string name, out int col, out int row)
        {
            col = name[0] - 65;
            row = Int32.Parse(name.Substring(1)) - 1;
        }

        private void displaySelection(SpreadsheetPanel ss)
        {
            setCellContents(m_selectedCell);

            int row, col;
            ss.GetSelection(out col, out row);
            m_selectedCell = getCellName(col, row);

            cellNameTextBox.Text = m_selectedCell;
            cellContentsTextBox.Text = getCellContents(m_selectedCell);
            cellValueTextBox.Text = m_spreadsheet.GetCellValue(m_selectedCell).ToString();

            cellContentsTextBox.Focus();
        }

        private string getCellContents(string name)
        {
            object contents = m_spreadsheet.GetCellContents(m_selectedCell);
            if (contents.GetType() == typeof(SpreadsheetUtilities.Formula))
                return "=" + contents.ToString();
            else
                return contents.ToString();
        }

        private void setCellContents(string name)
        {
            if (!m_contentChanged)
                return;
            bool error = false;
            errorText.Text = "";
            ISet<string> affectedCells;
            try
            {
                affectedCells = m_spreadsheet.SetContentsOfCell(name, cellContentsTextBox.Text);
                foreach (string cell in affectedCells)
                {
                    setGUICellValue(cell);
                }
            }
            catch(CircularException)
            {
                errorText.Text = "Circular Dependency";
                errorText.ToolTipText = "Circular Dependency";
                error = true;

            }
            catch(Exception e)
            {
                errorText.Text = e.Message;
                errorText.ToolTipText = e.Message;
                error = true;
            }

            if (m_twitterEnabled && !error && cellContentsTextBox.Text.Length > 0)
            {
                string truncatedContents = cellContentsTextBox.Text;
                string message = m_twitterAlias + " set contents of " + name + " to ";
                //twitter messages are limited to 140 chars, if the message is too long, truncate the contents
                //and add ... to the end
                if (truncatedContents.Length + message.Length + 1 > 140)
                    truncatedContents = truncatedContents.Substring(0, 140 - (message.Length + 1 + 3)) + "...";
                
                Task tweetTask = Task.Factory.StartNew(()=>{Tweet.PublishTweet(message + truncatedContents + "!");});
            }
        }

        private void setGUICellValue(string cell)
        {
            int col, row;
            getCellPosition(cell, out col, out row);
            object cellValue = m_spreadsheet.GetCellValue(cell);
            if (cellValue.GetType() == typeof(SpreadsheetUtilities.FormulaError))
                spreadsheetPanel1.SetValue(col, row, "FormulaError");
            else
                spreadsheetPanel1.SetValue(col, row, cellValue.ToString());

            if (cell == m_selectedCell)
            {
                cellValueTextBox.Text = m_spreadsheet.GetCellValue(cell).ToString();
            }
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SpreadsheetApplicationContext.getAppContext().RunForm(new SpreadsheetForm());
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void cellContentsTextBox_Leave(object sender, EventArgs e)
        {
            setCellContents(m_selectedCell);
        }

        private void cellContentsTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                setCellContents(m_selectedCell);
                e.SuppressKeyPress = true;
            }
            else if(e.Control)
            {
                int col, row;
                getCellPosition(m_selectedCell, out col, out row);

                switch(e.KeyCode)
                {
                    case Keys.Right:
                        col += 1;
                        break;
                    case Keys.Left:
                        col -= 1;
                        break;
                    case Keys.Up:
                        row -= 1;
                        break;
                    case Keys.Down:
                        row += 1;
                        break;
                    default:
                        break;
                }

                if(spreadsheetPanel1.SetSelection(col,row))
                {
                    displaySelection(spreadsheetPanel1);
                }
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = kFileFilter;
            openFileDialog.RestoreDirectory = true;

            if(openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    SpreadsheetApplicationContext.getAppContext().RunForm(new SpreadsheetForm(openFileDialog.FileName));
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. " + ex.Message);
                }
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (m_saveFilePath.Length == 0)
            {
                saveAsToolStripMenuItem_Click(sender, e);
            }
            else
            {
                try
                {
                    m_spreadsheet.Save(m_saveFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not save file to disk. " + ex.Message);
                }
            }
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = kFileFilter;
            saveFileDialog.RestoreDirectory = true;

            if(saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                m_saveFilePath = saveFileDialog.FileName;
                try
                {
                    m_spreadsheet.Save(m_saveFilePath);
                    this.Text = getFileName(m_saveFilePath) + " - Spreadsheet";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not save file to disk. " + ex.Message);
                }
            }
        }

        private void viewHelpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SpreadsheetApplicationContext.getAppContext().RunForm(new HelpForm());
        }

        private void twitterButton_ButtonClick(object sender, EventArgs e)
        {
            if(!m_twitterEnabled)
            {
                TwitterEnableDialog dialog = new TwitterEnableDialog();
                if(dialog.ShowDialog(this) ==  DialogResult.OK)
                {
                    m_twitterEnabled = true;
                    m_twitterAlias = dialog.userAlias;
                    twitterButton.Text = "Alias: " + m_twitterAlias;
                    setTwitterCredentials();
                }
            }
            else
            {
                DialogResult result = MessageBox.Show("Disable Twitter Integration?",
                                "Disable Twitter Integration?", MessageBoxButtons.OKCancel);
                if(result == DialogResult.OK)
                {
                    m_twitterEnabled = false;
                    twitterButton.Text = "Enable Twitter Integration";
                }
            }
        }

        private void cellContentsTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            m_contentChanged = true;
        }
    }
}
