package jExcel;

import java.util.ArrayList;

/**
 *
 * @author adm.trine@gmail.com
 */
class Cell {

    private Object value;
    private String formula;
    private String rformula;
    int state;
    private ArrayList<Cell> listeners;
    private ArrayList<Cell> references;
    private int row;
    private int column;
    private boolean editing;
    private Cell[][] cells;
    private Sheet sheet;
    private boolean rawvalued;

    Cell(Sheet sheet, int r, int c) {
        row = r;
        column = c;
        value = "";
        formula = "";
        rformula = "";
        listeners = new ArrayList<>();
        editing = false;
        rawvalued = false;
        this.sheet = sheet;
    }

    Cell(Sheet sheet, int r, int c, String value, String formula) {
        this(sheet, r, c);
        this.value = value;
        this.formula = formula;
    }

    public String getValue() {

        if (getFormula().length() < 1 || !isEditing()) {
            String strV = getValueObject().toString();
            if (strV.matches("[-+]?([0-9]+\\.[0])")) {
                strV = strV.split("\\.")[0];                //removes unwanted zeros after decimal pont.
            }
            if (strV.equals(formula.toUpperCase())) {
                strV = formula;                             
            }
            return strV;
        } else {
            setEditing(false);
            return "=" + getFormula();
        }

    }

    public void Compute(String input, Book book) {
        setCells(book.getSheets().get(0).getSheetModel().cells);
        setReferences(new ArrayList<Cell>());
        if (input.length() < 1) {
            this.setFormula("");
            this.setValueObject("");
        } else if (input.charAt(0) != '=') {
            this.setFormula("");
            this.setValueObject(input);
        } else {
            input = input.split("=")[1];
            //make formula
            this.setFormula(input);
            String tf = this.getFormula().replaceAll("\\[\\-", "\\[\\~");
            for (String x : tf.split("[\\+\\*\\-\\)\\(/,:\\!]")) {
                x = x.replaceAll("\\[\\~", "\\[\\-");
                if (x.matches("R(\\[[0-9\\-]+\\])*C(\\[[0-9\\-]+\\])*")) {
                    int r, c;
                    String tcol = x.replaceAll("R(\\[[0-9\\-]+\\])*", "");
                    String trow = x.replaceAll("C(\\[[0-9\\-]+\\])*", "");
                    if (tcol.matches("C")) {
                        c = 0;
                    } else {
                        c = Integer.parseInt(tcol.replaceAll("(C\\[)|(\\])", ""));
                    }
                    if (trow.matches("R")) {
                        r = 0;
                    } else {
                        r = Integer.parseInt(trow.replaceAll("(R\\[)|(\\])", ""));
                    }
                    c += getColumn();
                    r += getRow();
                    String s = (c / 26 < 1 ? "" : Character.valueOf((char) (c / 26 + 64)).toString()) + Character.valueOf((char) (c % 26 + 65)).toString() + (r + 1);
                    setFormula(getFormula().replace(x, s));
                    input = getFormula();
                }
            }

            //capitalizer
            for (String x : input.split("[\\+\\*\\-\\)\\(/,:\\!]")) {
                if (x.matches("[a-z]{1,2}[1-9][0-9]*") | x.matches("[A-Za-z0-9]+")) {
                    input = input.replaceAll(x, x.toUpperCase());
                }
            }

            //make relative formula
            this.setRformula(input);
            for (String x : getRformula().split("[\\+\\*\\-\\)\\(/,:\\!]")) {
                if (x.matches("[A-Z]{1,2}[1-9][0-9]*")) {
                    String col = x.replaceAll("[0-9]", "");
                    if (col.length() == 1) {
                        col = "@".concat(col);
                    }
                    int tmp = col.charAt(1) - 64;
                    int c = 26 * (col.charAt(0) - 64) + tmp;
                    int r = Integer.parseInt(x.replaceAll("[A-Z]", ""));
                    String s = "";
                    if (r - 1 - getRow() == 0) {
                        s += "R";
                    } else {
                        s += "R[" + (r - 1 - getRow()) + "]";
                    }
                    if (c - 1 - getColumn() == 0) {
                        s += "C";
                    } else {
                        s += "C[" + (c - 1 - getColumn()) + "]";
                    }
                    setRformula(getRformula().replaceFirst(x, s));
                }
            }
            //: simplifier
            if (input.contains(":")) {
                input = Cell.colonRemove(input);
            }
            //replacer
            for (String x : input.split("[\\+\\*\\-\\)\\(/,]")) {
                String tmpx = x;
                String orinput = input;
                if (x.matches("[A-Z]{1,2}[1-9][0-9]*") | x.matches("[A-Za-z0-9]+![A-Z]{1,2}[1-9][0-9]*")) {
                    if (x.matches("[A-Za-z0-9]+![A-Z]{1,2}[1-9][0-9]*")) {
                        int i = book.getSheetCollection().indexOfTab(x.split("!")[0]);
                        setCells(book.getSheets().get(i).getSheetModel().cells);
                        tmpx = x.split("!")[1];
                    } else {
                        setCells(getSheet().getSheetModel().cells);
                    }
                    String col = tmpx.replaceAll("[0-9]", "");
                    if (col.length() == 1) {
                        col = "@".concat(col);
                    }
                    int tmp = col.charAt(1) - 64;
                    int c = 26 * (col.charAt(0) - 64) + tmp;
                    int r = Integer.parseInt(tmpx.replaceAll("[A-Z]", ""));
                    if(r>getCells().length || c>getCells()[0].length){
                        setValueObject("#ERROR");
                        return;
                    }
                    if (getCells()[r - 1][c - 1].equals(this)) {
                        this.setValueObject("0");
                        System.out.println("Same cell cannot be reffered");
                        return;
                    }
                    if (!cells[r - 1][c - 1].listeners.contains(this)) {
                        getCells()[r - 1][c - 1].getListeners().add(this);
                    }
                    String s = getCells()[r - 1][c - 1].getValue();
                    if (s.length() < 1) {
                        s = "0";
                    }
                    input = input.replaceFirst(x, s);
                    this.getReferences().add(getCells()[r - 1][c - 1]);
                }
            }
            if (input.matches("([A-Z]+|[a-z]+)\\([0-9][0-9,-\\.]+[\\)]")) {
                input = resultValue.getRawExpression(input);
            }
            try {
                resultValue result = new resultValue(input);
                this.setValueObject(result.getResultStr());
            } catch (Exception e) {
                this.setValueObject("#ERROR");
            }
        }
            //listner update

        for (Cell x : new ArrayList<>(getListeners())) {

            if (x.getReferences().contains(this)) {
                x.Compute("=" + x.getFormula(), book);
            } else {
                this.getListeners().remove(x);
            }
        }
        this.setRawvalued(false);
    }

    public static String colonRemove(String input) {
        for (String x : input.split("[\\+\\*\\-\\)\\(/]")) {
            if (!x.contains(":")) {
                continue;
            }
            String l = x.split(":")[0];
            String r = x.split(":")[1];
            String l0 = l.replaceAll("[0-9]*", "");
            int l1 = Integer.parseInt(l.replaceAll("[A-Z]*", ""));
            String r0 = r.replaceAll("[0-9]*", "");
            int r1 = Integer.parseInt(r.replaceAll("[A-Z]*", ""));
            if (l0.equals(r0)) {
                if ((l1 - r1) * (l1 - r1) == 1) {
                    input = input.replaceAll(x, l + "," + r);
                } else {
                    if (l1 > r1) {
                        int t = l1;
                        l1 = r1;
                        r1 = t;
                    }
                    String midString = ",";
                    for (Integer i = l1 + 1; i < r1; i++) {
                        midString += l0 + i.toString() + ",";
                    }
                    input = input.replaceAll(x, l + midString + r);
                }
            } else if (l1 == r1) {
                Character cl0 = l0.charAt(0);
                Character cr0 = r0.charAt(0);
                if (Math.abs(cl0 - cr0) == 1) {
                    input = input.replaceAll(x, l + "," + r);
                } else {
                    if (cl0 > cr0) {
                        Character t = cl0;
                        cl0 = cr0;
                        cr0 = t;
                    }
                    String midString = ",";
                    for (Character i = ++cl0; i < cr0; i++) {
                        midString += i.toString() + l1 + ",";
                    }
                    input = input.replaceAll(x, l + midString + r);
                }
            }
        }

        return input;
    }

    /**
     * @return the formula
     */
    public String getFormula() {
        return formula;
    }

    /**
     * @param formula the formula to set
     */
    public void setFormula(String formula) {
        this.formula = formula;
    }

    /**
     * @return the rformula
     */
    public String getRformula() {
        return rformula;
    }

    /**
     * @param rformula the rformula to set
     */
    public void setRformula(String rformula) {
        this.rformula = rformula;
    }

    /**
     * @return the listeners
     */
    public ArrayList<Cell> getListeners() {
        return listeners;
    }

    /**
     * @param listeners the listeners to set
     */
    public void setListeners(ArrayList<Cell> listeners) {
        this.listeners = listeners;
    }

    /**
     * @return the references
     */
    public ArrayList<Cell> getReferences() {
        return references;
    }

    /**
     * @param references the references to set
     */
    public void setReferences(ArrayList<Cell> references) {
        this.references = references;
    }

    /**
     * @return the row
     */
    public int getRow() {
        return row;
    }

    /**
     * @param row the row to set
     */
    public void setRow(int row) {
        this.row = row;
    }

    /**
     * @return the column
     */
    public int getColumn() {
        return column;
    }

    /**
     * @param column the column to set
     */
    public void setColumn(int column) {
        this.column = column;
    }

    /**
     * @return the editing
     */
    public boolean isEditing() {
        return editing;
    }

    /**
     * @param editing the editing to set
     */
    public void setEditing(boolean editing) {
        this.editing = editing;
    }

    /**
     * @return the cells
     */
    public Cell[][] getCells() {
        return cells;
    }

    /**
     * @param cells the cells to set
     */
    public void setCells(Cell[][] cells) {
        this.cells = cells;
    }

    /**
     * @return the sheet
     */
    public Sheet getSheet() {
        return sheet;
    }

    /**
     * @param sheet the sheet to set
     */
    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    /**
     * @return the rawvalued
     */
    public boolean isRawvalued() {
        return rawvalued;
    }

    /**
     * @param rawvalued the rawvalued to set
     */
    public void setRawvalued(boolean rawvalued) {
        this.rawvalued = rawvalued;
    }

    /**
     * @return the value
     */
    public Object getValueObject() {
        return value;
    }

    /**
     * @param value the value to set
     */
    public void setValueObject(Object value) {
        this.value = value;
    }
}
