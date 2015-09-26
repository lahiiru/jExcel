package jExcel;

import javax.swing.table.*;

class SheetModel extends AbstractTableModel {

    final private Sheet guiTable;
    private int guiRow;
    private int guiColumn;

    protected Cell[][] cells;

    SheetModel(Cell[][] cells, Sheet table) {
        guiTable = table;
        guiRow = cells.length;
        guiColumn = cells[0].length;
        this.cells = cells;
    }

    public String getFormularBarText(int r, int c) {
        String s = "Click on a cell";
        if ((r + 1) * (c + 1) > 0) {
            if (cells[r][c].getFormula().length() < 1) {
                s = cells[r][c].getValue();
            } else {
                s = "=" + cells[r][c].getFormula();
            }
            guiTable.getBook().getContainer().getFormulaBar().setEditable(true);
        }else{
        guiTable.getBook().getContainer().getFormulaBar().setEditable(false);
        }
        return s;
    }

    public int getRowCount() {
        return guiRow;
    }

    public int getColumnCount() {
        return guiColumn;
    }

    public boolean isCellEditable(int row, int col) {
        return true;
    }

    public Object getValueAt(int row, int column) {
        //if(cells[row][column].rawvalued==false){cells[row][column].Compute(cells[row][column].formula, cells[row][column].sheet.book);}
        String r = cells[row][column].getValue();
        cells[row][column].setEditing(false);
        return r;

    }

    public void setValueAt(Object value, int row, int column) {
        String input = (String) value;
        if (input.equals("=")) {
            return;
        }
        guiTable.getBook().setSaved(false);
        if (value != null) {
            cells[row][column].Compute(input, guiTable.getBook());
        } else {
            cells[row][column].setValueObject("");
            cells[row][column].setFormula("");
        }
        guiTable.repaint();
    }

}
