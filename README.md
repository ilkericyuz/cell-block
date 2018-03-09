# CellBlock
An Microsoft Excel Class Module which lets you define a cell block as a table and reach values in it

 Cell Block Class defines a cell block and lets you reach values in it
 A Cell Block is a group of cells which can be defined as a table.

 Referencing from an initial cell, class finds the edges of the cell block and defines:
 topRow:          Range object which refers to top row of the cell block
 bottomRow:       Range object which refers to bottom row of the cell block
 leftColumn:      Range object which refers to left column of the cell block
 rightColumn:     Range object which refers to right column of the cell block
 titleRow:        Range object which refers to entire top row of the cell block
 titleColumn:     Range object which refers to entire left column of the cell block
 titleRowCells:   Range object which refers to cell block part of the title row
 titleColumnCells:Range object which refers to cell block part of the title column
 top:             Row number of the top row
 left:            Column number of the left column
 width:           Width of the cell block
 height:          Height of the cell block
 size:            Total number of cells in the cell block
 activeRowCells:  Dictionary object which let you reach any value in the selected row
                  using title of the row as key
    

 activeColumnCells:  Just like activeRowCells, but refers to cell values according to their
                     title columns
                     
 Usage:
      ' Select any cell in a group of cells or a table and
      
      dim cb as new CellBlock
      cb.InitializeProperties()
      
      ' Now you can reach your table and any value in it
      ' cb.activeRowCells([title])
      Range(cb.firstCell, cb.lastCell).Select
      
Example:
    
      ' Date        A   B   C   D
      ' 20170101    1   2   3   ronaldinho
      ' 20170102   [4]  5   6   messi
      ' 20170103    7   8   9   c. ronaldo

      ' As 4 is selected, activeRowCells("D") gives you messi, activeRowCells("C") gives you 6
