//Spreadsheet.cs
//Created by Elliot Hatch on October 2014

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Xml;
using SpreadsheetUtilities;

namespace SS
{
    public class Spreadsheet : AbstractSpreadsheet
    {
        /// <summary>
        /// Cells are stored in a Dictionary with the cell name as the key, and objects of the Cell class as values.
        /// Dependencies are tracked when a cell is added or removed by using the DependencyGraph class.
        /// </summary>
        private Dictionary<string, Cell> m_cells;
        private DependencyGraph m_dependencyGraph;

        /// <summary>
        /// True if this spreadsheet has been modified since it was created or saved                  
        /// (whichever happened most recently); false otherwise.
        /// </summary>
        public override bool Changed { get; protected set; }

        /// <summary>
        /// Create a spreadsheet with isValid delegate that always returns true, an identity normalizer function, and the default version ("default")
        /// </summary>
        public Spreadsheet()
            :this(s => true, s => s, "default")
        {
        }

        /// <summary>
        /// Create a spreadsheet with the provided isValid and normalize delegates and the given version number.
        /// </summary>
        /// <param name="isValid"></param>
        /// <param name="normalize"></param>
        /// <param name="version"></param>
        public Spreadsheet(Func<string, bool> isValid, Func<string, string> normalize, string version)
            :base(isValid, normalize, version)
        {
            Changed = false;
            m_cells = new Dictionary<string, Cell>();
            m_dependencyGraph = new DependencyGraph();
        }

        /// <summary>
        /// Opens a saved spreadsheet located at path, only if the cell names conform to the isValid delegate and the spreadsheet has the same version as the one given in this constructor.
        /// Also normalizes all cell names using the given normalie delegate
        /// </summary>
        /// <param name="path">Spreadsheet file path</param>
        /// <param name="isValid"></param>
        /// <param name="normalize"></param>
        /// <param name="version"></param>
        public Spreadsheet(string path, Func<string, bool> isValid, Func<string, string> normalize, string version)
            : this(isValid, normalize, version)
        {
            XmlReader xmlReader = null;
            try
            {
                xmlReader = XmlReader.Create(path);
                xmlReader.ReadToFollowing("spreadsheet");
                //get version number
                string fileVersion = xmlReader.GetAttribute("version");
                if(fileVersion != version)
                    throw new SpreadsheetReadWriteException("Wrong version: Tried to open " + path + "as version " + version + "but the file was version " + fileVersion);
                
                //read all cells and add them to spreadsheet
                while (xmlReader.ReadToFollowing("cell"))
                {
                    xmlReader.ReadToFollowing("name");
                    xmlReader.Read();
                    string name = xmlReader.Value;
                    xmlReader.ReadToFollowing("contents");
                    xmlReader.Read();
                    string contents = xmlReader.Value;
                    SetContentsOfCell(name, contents);
                }
            }
            catch(Exception e)
            {
                throw new SpreadsheetReadWriteException("Failed reading " + path + ": " + e);
            }
            finally
            {
                if (xmlReader != null)
                    xmlReader.Close();
            }

            Changed = false;
        }

        /// <summary>
        /// Returns the version information of the spreadsheet saved in the named file.
        /// If there are any problems opening, reading, or closing the file, the method
        /// should throw a SpreadsheetReadWriteException with an explanatory message.
        /// </summary>
        public override string GetSavedVersion(String filename)
        {
            string version = "";
            XmlReader xmlReader = null;
            try
            {
                xmlReader = XmlReader.Create(filename);
                xmlReader.ReadToFollowing("spreadsheet");
                version = xmlReader.GetAttribute("version");
            }
            catch(Exception e)
            {
                throw new SpreadsheetReadWriteException("Failed reading " + filename + ": " + e);
            }
            finally
            {
                if (xmlReader != null)
                    xmlReader.Close();
            }
            return version;
        }

        /// <summary>
        /// Writes the contents of this spreadsheet to the named file using an XML format.
        /// The XML elements should be structured as follows:
        /// 
        /// <spreadsheet version="version information goes here">
        /// 
        /// <cell>
        /// <name>
        /// cell name goes here
        /// </name>
        /// <contents>
        /// cell contents goes here
        /// </contents>    
        /// </cell>
        /// 
        /// </spreadsheet>
        /// 
        /// There should be one cell element for each non-empty cell in the spreadsheet.  
        /// If the cell contains a string, it should be written as the contents.  
        /// If the cell contains a double d, d.ToString() should be written as the contents.  
        /// If the cell contains a Formula f, f.ToString() with "=" prepended should be written as the contents.
        /// 
        /// If there are any problems opening, writing, or closing the file, the method should throw a
        /// SpreadsheetReadWriteException with an explanatory message.
        /// </summary>
        public override void Save(String filename)
        {
            XmlWriter xmlWriter = null;
            try
            {
                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Indent = true;
                xmlWriter = XmlWriter.Create(filename, settings);
                xmlWriter.WriteStartElement("spreadsheet");
                xmlWriter.WriteAttributeString("version", Version);

                //write each cell
                foreach (KeyValuePair<string, Cell> cellKeyValue in m_cells)
                {
                    xmlWriter.WriteStartElement("cell");
                    xmlWriter.WriteElementString("name", cellKeyValue.Key);
                    xmlWriter.WriteElementString("contents", cellKeyValue.Value.ToString());
                    xmlWriter.WriteEndElement();
                }
                xmlWriter.WriteEndElement();
            }
            catch (Exception e)
            {
                throw new SpreadsheetReadWriteException("Failed writing " + filename + ": " + e);
            }
            finally
            {
                if (xmlWriter != null)
                    xmlWriter.Close();
            }
            Changed = false;
        }

        /// <summary>
        /// If name is null or invalid, throws an InvalidNameException.
        /// 
        /// Otherwise, returns the value (as opposed to the contents) of the named cell.  The return
        /// value should be either a string, a double, or a SpreadsheetUtilities.FormulaError.
        /// </summary>
        public override object GetCellValue(String name)
        {
            string normalizedName = Normalize(name);
            checkCellName(normalizedName);

            Cell cell;
            if (m_cells.TryGetValue(normalizedName, out cell))
            {
                return cell.getValue(variableLookup);
            }
            return "";
        }

        /// <summary>
        /// If content is null, throws an ArgumentNullException.
        /// 
        /// Otherwise, if name is null or invalid, throws an InvalidNameException.
        /// 
        /// Otherwise, if content parses as a double, the contents of the named
        /// cell becomes that double.
        /// 
        /// Otherwise, if content begins with the character '=', an attempt is made
        /// to parse the remainder of content into a Formula f using the Formula
        /// constructor.  There are then three possibilities:
        /// 
        ///   (1) If the remainder of content cannot be parsed into a Formula, a 
        ///       SpreadsheetUtilities.FormulaFormatException is thrown.
        ///       
        ///   (2) Otherwise, if changing the contents of the named cell to be f
        ///       would cause a circular dependency, a CircularException is thrown.
        ///       
        ///   (3) Otherwise, the contents of the named cell becomes f.
        /// 
        /// Otherwise, the contents of the named cell becomes content.
        /// 
        /// If an exception is not thrown, the method returns a set consisting of
        /// name plus the names of all other cells whose value depends, directly
        /// or indirectly, on the named cell.
        /// 
        /// For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        /// set {A1, B1, C1} is returned.
        /// </summary>
        public override ISet<String> SetContentsOfCell(String name, String content)
        {
            if(content == null)
                throw new ArgumentNullException();

            string normalizedName = Normalize(name);

            checkCellName(normalizedName);

            ISet<string> cellsToRecalculate = null;

            double doubleValue;
            if(Double.TryParse(content, out doubleValue))
            {
                //add as double
                cellsToRecalculate = SetCellContents(normalizedName, doubleValue);
            }
            else if(content.Length > 0 && content.First() == '=')
            {
                //add as formula
                cellsToRecalculate = SetCellContents(normalizedName, new Formula(content.Substring(1), Normalize, IsValid));
            }
            else
            {
                //add as string
                cellsToRecalculate = SetCellContents(normalizedName, content);
            }

            //set each dependent cell to need a value recalculation
            foreach(string cellName in cellsToRecalculate)
            {
                Cell cell;
                if (m_cells.TryGetValue(cellName, out cell))
                {
                    cell.setNeedsRecalulation();
                }
            }

            Changed = true;
            return cellsToRecalculate;
        }

        /// <summary>
        /// Enumerates the names of all the non-empty cells in the spreadsheet.
        /// </summary>
        public override IEnumerable<String> GetNamesOfAllNonemptyCells()
        {
            return m_cells.Keys;
        }

        /// <summary>
        /// If name is null or invalid, throws an InvalidNameException.
        /// 
        /// Otherwise, returns the contents (as opposed to the value) of the named cell.  The return
        /// value should be either a string, a double, or a Formula.
        public override object GetCellContents(String name)
        {
            string normalizedName = Normalize(name);
            checkCellName(normalizedName);

            Cell cell;
            if (m_cells.TryGetValue(normalizedName, out cell))
            {
                return cell.asObject();
            }

            //cell not found - must be an empty string
            return "";
        }

        /// <summary>
        /// If name is null or invalid, throws an InvalidNameException.
        /// 
        /// Otherwise, the contents of the named cell becomes number.  The method returns a
        /// set consisting of name plus the names of all other cells whose value depends, 
        /// directly or indirectly, on the named cell.
        /// 
        /// For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        /// set {A1, B1, C1} is returned.
        /// </summary>
        protected override ISet<String> SetCellContents(String name, double number)
        {
            removeCellDependencies(name);
            m_cells[name] = new Cell(number, Cell.CellType.Number);

            return new HashSet<string>(GetCellsToRecalculate(name));
        }

        /// <summary>
        /// If text is null, throws an ArgumentNullException.
        /// 
        /// Otherwise, if name is null or invalid, throws an InvalidNameException.
        /// 
        /// Otherwise, the contents of the named cell becomes text.  The method returns a
        /// set consisting of name plus the names of all other cells whose value depends, 
        /// directly or indirectly, on the named cell.
        /// 
        /// For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        /// set {A1, B1, C1} is returned.
        /// </summary>
        protected override ISet<String> SetCellContents(String name, String text)
        {
            removeCellDependencies(name);
            //set contents of cell - if passed an empty string we can remove the cell
            if (text != "")
                m_cells[name] = new Cell(text, Cell.CellType.String);
            else
                m_cells.Remove(name);

            return new HashSet<string>(GetCellsToRecalculate(name));
        }

        /// <summary>
        /// If the formula parameter is null, throws an ArgumentNullException.
        /// 
        /// Otherwise, if name is null or invalid, throws an InvalidNameException.
        /// 
        /// Otherwise, if changing the contents of the named cell to be the formula would cause a 
        /// circular dependency, throws a CircularException.  (No change is made to the spreadsheet.)
        /// 
        /// Otherwise, the contents of the named cell becomes formula.  The method returns a
        /// Set consisting of name plus the names of all other cells whose value depends,
        /// directly or indirectly, on the named cell.
        /// 
        /// For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        /// set {A1, B1, C1} is returned.
        /// </summary>
        protected override ISet<String> SetCellContents(String name, Formula formula)
        {
            //save off old cell value
            Cell oldCell = null;
            m_cells.TryGetValue(name, out oldCell);

            removeCellDependencies(name);
            m_cells[name] = new Cell(formula, Cell.CellType.Formula);
            IEnumerable<string> variables = formula.GetVariables();
            foreach(string v in variables)
            {
                m_dependencyGraph.AddDependency(v, name);
            }

            try
            {
                HashSet<string> cellsToRecalculate = new HashSet<string>(GetCellsToRecalculate(name));
                return cellsToRecalculate;
            }
            catch(CircularException e)
            {
                //if there was a circular exception, restore old cell
                if (oldCell == null)
                    m_cells.Remove(name);
                else
                {
                    switch(oldCell.cellType)
                    {
                        case Cell.CellType.Formula:
                            Formula f = oldCell.asFormula();
                            SetCellContents(name, f);
                            break;
                        case Cell.CellType.Number:
                            double d = oldCell.asDouble();
                            SetCellContents(name, d);
                            break;
                        case Cell.CellType.String:
                             string s = oldCell.asString();
                            SetCellContents(name, s);
                            break;
                    }
                }
                throw e;
            }
        }

        /// <summary>
        /// If name is null, throws an ArgumentNullException.
        /// 
        /// Otherwise, if name isn't a valid cell name, throws an InvalidNameException.
        /// 
        /// Otherwise, returns an enumeration, without duplicates, of the names of all cells whose
        /// values depend directly on the value of the named cell.  In other words, returns
        /// an enumeration, without duplicates, of the names of all cells that contain
        /// formulas containing name.
        /// 
        /// For example, suppose that
        /// A1 contains 3
        /// B1 contains the formula A1 * A1
        /// C1 contains the formula B1 + A1
        /// D1 contains the formula B1 - C1
        /// The direct dependents of A1 are B1 and C1
        /// </summary>
        protected override IEnumerable<String> GetDirectDependents(String name)
        {
            return m_dependencyGraph.GetDependents(name);
        }

        /// <summary>
        /// if name is null, throws ArgumentNullException
        /// if name is not valid throws InvalidNameException
        /// A string is a cell name if and only if it consists of one or more letters,
        /// followed by one or more digits AND it satisfies the predicate IsValid.
        /// </summary>
        /// <param name="name"></param>
        protected void checkCellName(String name)
        {
            if(name == null)
                throw new InvalidNameException();

            if(!Regex.IsMatch(name, @"^[a-zA-Z]+[0-9]+$"))
                throw new InvalidNameException();

            if(!IsValid(name))
                throw new InvalidNameException();
        }

        /// <summary>
        /// Removes all dependees of a cell. This method should be called anytime a cell is replaced or removed.
        /// </summary>
        /// <param name="name">Name of the cell whose dependees should be removed</param>
        protected void removeCellDependencies(String name)
        {
            Cell cell;
            if (m_cells.TryGetValue(name, out cell))
            {
                if (cell.cellType == Cell.CellType.Formula)
                {
                    m_dependencyGraph.ReplaceDependees(name, Enumerable.Empty<string>());
                }
            }
        }
        /// <summary>
        /// Delegate that gets the double value of a cell. If it can't, throws ArgumentException()
        /// Recursively calls cell.GetValue(variableLookup) until the formula is evaluated.
        /// </summary>
        /// <param name="variableName"></param>
        /// <returns></returns>
        protected double variableLookup(string variableName)
        {
            Cell cell;
            if (m_cells.TryGetValue(variableName, out cell))
            {
                //only allow lookup if the cell is a number or formula, and is not a FormulaError
                if (cell.cellType == Cell.CellType.Number || cell.cellType == Cell.CellType.Formula)
                {
                    object value = cell.getValue(variableLookup);
                    if (!(value is FormulaError))
                        return (double)value;
                    else
                        throw new ArgumentException("Cell " + variableName + " has a Formula Error:" + ((FormulaError)value).Reason);
                }
                else
                    throw new ArgumentException("Cell " + variableName + " does not contain a number or formula.");
            }
            throw new ArgumentException("Cell " + variableName + " is empty.");
        }
    }
}

