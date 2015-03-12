Spreadsheet - created by Elliot Hatch - October 2014

Included Libraries:
	Formula - Debug: Updated October 9 2014
	Dependency Graph - Debug: Updated October 9 2014

10/15/2014
Updated to PS5 specs
Added "value" member to Cell class
	Formulas are initially uncalculated until GetCellValue is called, then each dependee of the cell has GetValue recursively called to get the final value.
	The calculated value is stored in the value field until the cell has its NeedsRecalculation flag reset. The flag is reset in SetContentsOfCell using the dependent cells that will be returned.
Added saving and loading using XmlReader and XmlWriter


SpreadsheetGUI

11/04/2014
Added Windows Form GUI
Added twitter integration
	using Tweetinvi library - https://tweetinvi.codeplex.com/
