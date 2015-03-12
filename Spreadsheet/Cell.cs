//Cell.cs
//Created by Elliot Hatch on October 2014
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetUtilities;

namespace SS
{
    /// <summary>
    /// Internal Cell class, used by Spreadsheet
    /// A cell represents one cell in a spreadsheet. It has a contents (string, number, or formula) and a value (string, number, or evaluated formula (double))
    /// Cell also stores its CellType and provides helper casting functions.
    /// </summary>
    class Cell
    {
        protected object m_contents;
        protected CellType m_type;
        protected object m_value;
        protected bool m_needsRecalculation;

        public Cell(object contents, CellType type)
        {
            m_contents = contents;
            m_type = type;
            m_value = null;
            m_needsRecalculation = true;
            if (type == CellType.Number)
            {
                m_value = contents;
                m_needsRecalculation = false;
            }


        }
        public CellType cellType
        {
            get { return m_type; }
        }
        public object asObject()
        {
            return m_contents;
        }
        public string asString()
        {
            return m_contents as string;
        }
        public double asDouble()
        {
            return Convert.ToDouble(m_contents);
        }
        public Formula asFormula()
        {
            return m_contents as Formula;
        }
        /// <summary>
        /// Returns the value of a cell. If the cell is a formula, calls asFormula().Evaluate(lookup) to calculate the value, then returns it
        /// If the cell doesn't need to be reclaculated the cell returns its stored value
        /// </summary>
        /// <param name="lookup"></param>
        /// <returns></returns>
        public object getValue(Func<string, double> lookup)
        {
            if (m_type == CellType.String)
                return m_contents;

            if (m_needsRecalculation && m_type == CellType.Formula)
            {
                m_value = asFormula().Evaluate(lookup);
                m_needsRecalculation = false;
            }

            return m_value;
        }
        public void setNeedsRecalulation()
        {
            m_needsRecalculation = true;
        }

        public override string ToString()
        {
            switch (m_type)
            {
                case CellType.Formula:
                    return "=" + m_contents.ToString();
                case CellType.Number:
                    return m_contents.ToString();
                case CellType.String:
                    return m_contents.ToString();
                default:
                    return m_contents.ToString();
            }
        }

        public enum CellType
        {
            String,
            Number,
            Formula
        }
    }
}
