package jExcel;

import javax.swing.*;
import java.awt.*;
import java.util.ArrayList;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

/**
 *
 * @author adm.trine@gmail.com
 */
public class Book extends JPanel {

    private int sheetCount;
    private ArrayList<Sheet> sheets;
    private boolean saved;
    private String savedPath;
    private JTabbedPane sheetCollection;

    private Window container;

    public void addNewSheet() {
        String name = "SHEET";
        for (int i = 1; i < 1000; i++) {
            ArrayList<String> titles = new ArrayList<>();
            for (int j = 0; j < getSheetCollection().getTabCount(); j++) {
                titles.add(getSheetCollection().getTitleAt(j));
            }
            if (titles.contains(name + i)) {
                continue;
            }
            name += i;
            break;
        }
        Sheet s = new Sheet(this, 40, 40, name);
        getSheets().add(s);
        setSheetCount(getSheetCount() + 1);
        getSheetCollection().add(getSheets().get(getSheetCount() - 1).getSheetName(), getSheets().get(getSheetCount() - 1).getSheet());
    }

    public void closeSheet(int i) {
        if (i < 0) {
            i = getSheetCollection().getSelectedIndex();
        }
        if (i < getSheetCount()) {
            if (i >= getSheetCollection().getTabCount()) {
                getSheetCollection().repaint();
            } else {
                if (getSheetCollection().getTabCount() == 1) {
                    getSheetCollection().removeAll();
                    getContainer().getFormulaBar().setVisible(false);
                    getContainer().getBarCaption().setVisible(false);
                    getContainer().menuBar.getMenu(0).getItem(2).setEnabled(false);
                    getContainer().menuBar.getMenu(0).getItem(3).setEnabled(false);
                    getContainer().menuBar.getMenu(0).getItem(4).setEnabled(false);
                    getContainer().menuBar.getMenu(0).getItem(6).setEnabled(false);
                    ((JButton) getContainer().Ribbon.getComponents()[2]).setEnabled(false);
                    ((JButton) getContainer().Ribbon.getComponents()[3]).setEnabled(false);
                } else {
                    getSheetCollection().removeTabAt(i);
                }
                getSheets().remove(i);
            }
            //container.pack();
        }
        setSheetCount(getSheetCollection().getTabCount());
    }

    public Sheet getActiveSheet() {
        if (getSheetCollection().getTabCount() == 0) {
            return null;
        }
        return getSheets().get(getSheetCollection().getSelectedIndex());
    }

    public Book(Window window) {
        super();
        saved = true;
        this.container = window;
        this.savedPath = "";
        sheets = new ArrayList<>();
        sheetCollection = new JTabbedPane();
        sheetCollection.setTabPlacement(JTabbedPane.BOTTOM);

        ChangeListener changeListener = new ChangeListener() {
            public void stateChanged(ChangeEvent changeEvent) {
                if (getSheetCollection().getTabCount() > 0) {

                    getContainer().getFormulaBar().setVisible(true);
                    getContainer().getBarCaption().setVisible(true);
                    getContainer().menuBar.getMenu(0).getItem(2).setEnabled(true);
                    getContainer().menuBar.getMenu(0).getItem(3).setEnabled(true);
                    getContainer().menuBar.getMenu(0).getItem(4).setEnabled(true);
                    getContainer().menuBar.getMenu(0).getItem(6).setEnabled(true);
                    ((JButton) getContainer().Ribbon.getComponents()[2]).setEnabled(true);
                    ((JButton) getContainer().Ribbon.getComponents()[3]).setEnabled(true);
                    try {
                        if (getSheets().toArray().length > getSheetCollection().getSelectedIndex()) {
                            int r, c;
                            r = getSheets().get(getSheetCollection().getSelectedIndex()).getSelectedRow();
                            c = getSheets().get(getSheetCollection().getSelectedIndex()).getSelectedColumn();
                            getContainer().getFormulaBar().setText(getSheets().get(getSheetCollection().getSelectedIndex()).getSheetModel().getFormularBarText(r, c));
                        }
                    } catch (Exception e) {
                    }

                } else {
                    getContainer().getFormulaBar().setVisible(false);
                    getContainer().getBarCaption().setVisible(false);
                    getContainer().menuBar.getMenu(0).getItem(2).setEnabled(false);
                    getContainer().menuBar.getMenu(0).getItem(3).setEnabled(false);
                    getContainer().menuBar.getMenu(0).getItem(4).setEnabled(false);
                    getContainer().menuBar.getMenu(0).getItem(6).setEnabled(false);
                    ((JButton) getContainer().Ribbon.getComponents()[2]).setEnabled(false);
                    ((JButton) getContainer().Ribbon.getComponents()[3]).setEnabled(false);
                }
            }
        };
        sheetCollection.addChangeListener(changeListener);
        //this.setPreferredSize(new Dimension(500,200));
        this.setLayout(new BoxLayout(this, BoxLayout.PAGE_AXIS));

        this.add(sheetCollection, BorderLayout.CENTER);

    }

    /**
     * @return the sheetCount
     */
    public int getSheetCount() {
        return sheetCount;
    }

    /**
     * @param sheetCount the sheetCount to set
     */
    public void setSheetCount(int sheetCount) {
        this.sheetCount = sheetCount;
    }

    /**
     * @return the sheets
     */
    public ArrayList<Sheet> getSheets() {
        return sheets;
    }

    /**
     * @param sheets the sheets to set
     */
    public void setSheets(ArrayList<Sheet> sheets) {
        this.sheets = sheets;
    }

    /**
     * @return the saved
     */
    public boolean isSaved() {
        return saved;
    }

    /**
     * @param saved the saved to set
     */
    public void setSaved(boolean saved) {
        this.saved = saved;
    }

    /**
     * @return the sheetCollection
     */
    public JTabbedPane getSheetCollection() {
        return sheetCollection;
    }

    /**
     * @param sheetCollection the sheetCollection to set
     */
    public void setSheetCollection(JTabbedPane sheetCollection) {
        this.sheetCollection = sheetCollection;
    }

    /**
     * @return the container
     */
    public Window getContainer() {
        return container;
    }

    /**
     * @param container the container to set
     */
    public void setContainer(Window container) {
        this.container = container;
    }

    /**
     * @return the savedPath
     */
    public String getSavedPath() {
        return savedPath;
    }

    /**
     * @param savedPath the savedPath to set
     */
    public void setSavedPath(String savedPath) {
        this.savedPath = savedPath;
    }
}
